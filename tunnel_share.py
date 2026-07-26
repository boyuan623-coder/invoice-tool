"""
本机外网分享：通过 Cloudflare 临时隧道生成公网链接。
不需要购买云服务器；本机关闭或停止程序后链接失效。

依赖：首次运行自动下载 cloudflared（Windows/macOS/Linux）。
"""

from __future__ import annotations

import os
import platform
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import urllib.request
from typing import Optional

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOOLS_DIR = os.path.join(SCRIPT_DIR, "tools")
PUBLIC_URL_RE = re.compile(r"https://[a-zA-Z0-9\-]+\.trycloudflare\.com")


def get_lan_ips() -> list[str]:
    """获取本机局域网 IPv4 地址（排除回环）。"""
    ips = []
    try:
        hostname = socket.gethostname()
        for info in socket.getaddrinfo(hostname, None, socket.AF_INET):
            ip = info[4][0]
            if ip and not ip.startswith("127.") and ip not in ips:
                ips.append(ip)
    except OSError:
        pass

    # 更可靠：连一个外部地址探测本机出口网卡 IP（不真正发包）
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        if ip and not ip.startswith("127.") and ip not in ips:
            ips.insert(0, ip)
    except OSError:
        pass
    return ips


def _cloudflared_download_url() -> Optional[str]:
    system = platform.system().lower()
    machine = platform.machine().lower()
    base = "https://github.com/cloudflare/cloudflared/releases/latest/download"
    if system == "windows":
        return f"{base}/cloudflared-windows-amd64.exe"
    if system == "darwin":
        if "arm" in machine or "aarch64" in machine:
            return f"{base}/cloudflared-darwin-arm64.tgz"
        return f"{base}/cloudflared-darwin-amd64.tgz"
    if system == "linux":
        if "arm" in machine or "aarch64" in machine:
            return f"{base}/cloudflared-linux-arm64"
        return f"{base}/cloudflared-linux-amd64"
    return None


def _local_cloudflared_path() -> str:
    name = "cloudflared.exe" if platform.system().lower() == "windows" else "cloudflared"
    return os.path.join(TOOLS_DIR, name)


def resolve_cloudflared() -> Optional[str]:
    """优先使用 PATH / 本地 tools 目录中的 cloudflared。"""
    found = shutil.which("cloudflared")
    if found:
        return found
    local = _local_cloudflared_path()
    if os.path.isfile(local):
        return local
    return None


def _download_urls() -> list[str]:
    """优先镜像（国内直连 GitHub 常卡住），最后才试官方源。"""
    primary = _cloudflared_download_url()
    if not primary:
        return []
    return [
        "https://ghfast.top/" + primary,
        "https://ghproxy.net/" + primary,
        "https://mirror.ghproxy.com/" + primary,
        primary,
    ]


def _download_file_curl(url: str, dest: str, timeout: int = 120) -> bool:
    """优先用 curl，超时更快、国内成功率更高。"""
    curl = shutil.which("curl") or shutil.which("curl.exe")
    if not curl:
        return False
    cmd = [
        curl, "-L", "--fail",
        "--connect-timeout", "15",
        "--max-time", str(timeout),
        "-A", "invoice-tool/share",
        "-o", dest,
        url,
    ]
    creationflags = 0
    if platform.system().lower() == "windows":
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    proc = subprocess.run(
        cmd, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
        creationflags=creationflags,
    )
    if proc.returncode != 0:
        raise RuntimeError((proc.stderr or proc.stdout or "curl failed")[-300:])
    return True


def _download_file_urllib(url: str, dest: str, timeout: int = 45):
    req = urllib.request.Request(url, headers={"User-Agent": "invoice-tool/share"})
    with urllib.request.urlopen(req, timeout=timeout) as resp, open(dest, "wb") as out:
        shutil.copyfileobj(resp, out)


def _download_file(url: str, dest: str, timeout: int = 120):
    try:
        if _download_file_curl(url, dest, timeout=timeout):
            return
    except Exception:
        # curl 失败再回退 urllib
        pass
    _download_file_urllib(url, dest, timeout=min(timeout, 60))


def ensure_cloudflared(progress_print=print) -> str:
    """若本地没有 cloudflared，则自动下载到 tools/。"""
    existing = resolve_cloudflared()
    if existing:
        return existing

    urls = _download_urls()
    if not urls:
        raise RuntimeError(f"当前系统暂不支持自动下载 cloudflared: {platform.system()}")

    os.makedirs(TOOLS_DIR, exist_ok=True)
    dest = _local_cloudflared_path()
    progress_print("  [分享] 首次使用，正在下载 Cloudflare 隧道工具（约 50MB）...")
    progress_print("  [分享] 若长时间无进度，请关闭窗口后按使用说明手动下载")

    tmp = dest + ".download"
    last_error = None
    for url in urls:
        progress_print(f"  [分享] 尝试: {url}")
        try:
            _download_file(url, tmp)
            is_tgz = ".tgz" in url.split("?")[0]
            if is_tgz:
                import tarfile
                with tarfile.open(tmp, "r:gz") as tar:
                    member = next(
                        (m for m in tar.getmembers() if os.path.basename(m.name) == "cloudflared"),
                        None,
                    )
                    if not member:
                        raise RuntimeError("压缩包中未找到 cloudflared")
                    member.name = os.path.basename(dest)
                    tar.extract(member, TOOLS_DIR)
                try:
                    os.unlink(tmp)
                except OSError:
                    pass
            else:
                if os.path.getsize(tmp) < 1_000_000:
                    raise RuntimeError("下载文件过小，可能不是有效安装包")
                os.replace(tmp, dest)
            if platform.system().lower() != "windows":
                os.chmod(dest, 0o755)
            last_error = None
            break
        except Exception as e:
            last_error = e
            progress_print(f"  [分享] 该地址失败，尝试下一个... ({e})")
            try:
                if os.path.exists(tmp):
                    os.unlink(tmp)
            except OSError:
                pass

    if last_error or not os.path.isfile(dest):
        manual = _cloudflared_download_url() or ""
        raise RuntimeError(
            "cloudflared 自动下载失败。\n"
            f"  请手动下载: {manual}\n"
            f"  或镜像: https://ghfast.top/{manual}\n"
            f"  保存为: {dest}\n"
            f"  保存后重新双击 invoice_tool_share.bat\n"
            f"  原因: {last_error}"
        )

    progress_print(f"  [分享] 已保存到: {dest}")
    return dest


class ShareTunnel:
    """Cloudflare 临时隧道，进程随主程序退出而结束。"""

    def __init__(self, local_port: int = 5000):
        self.local_port = local_port
        self.public_url: Optional[str] = None
        self._proc: Optional[subprocess.Popen] = None
        self._error: Optional[str] = None

    def start(self, timeout: float = 45.0) -> str:
        binary = ensure_cloudflared()
        local_url = f"http://127.0.0.1:{self.local_port}"
        creationflags = 0
        if platform.system().lower() == "windows":
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        self._proc = subprocess.Popen(
            [binary, "tunnel", "--url", local_url],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=creationflags,
        )

        deadline = time.time() + timeout
        assert self._proc.stdout is not None
        while time.time() < deadline:
            if self._proc.poll() is not None:
                leftover = self._proc.stdout.read() or ""
                self._error = leftover[-500:] or "cloudflared 异常退出"
                raise RuntimeError(f"隧道启动失败: {self._error}")

            line = self._proc.stdout.readline()
            if not line:
                time.sleep(0.1)
                continue
            m = PUBLIC_URL_RE.search(line)
            if m:
                self.public_url = m.group(0)
                # 继续在后台读掉输出，避免管道阻塞
                threading.Thread(target=self._drain, daemon=True).start()
                return self.public_url

        self.stop()
        raise RuntimeError("等待公网链接超时，请检查网络后重试")

    def _drain(self):
        try:
            if self._proc and self._proc.stdout:
                for _ in self._proc.stdout:
                    pass
        except Exception:
            pass

    def stop(self):
        if self._proc and self._proc.poll() is None:
            try:
                self._proc.terminate()
                self._proc.wait(timeout=5)
            except Exception:
                try:
                    self._proc.kill()
                except Exception:
                    pass
        self._proc = None


def start_share_tunnel(port: int = 5000, progress_print=print) -> ShareTunnel:
    """启动隧道并打印访问信息。"""
    progress_print("")
    progress_print("  [分享] 正在创建临时公网链接（Cloudflare Tunnel）...")
    progress_print("  [分享] 不需要云服务器；关闭本窗口后链接立即失效")
    tunnel = ShareTunnel(local_port=port)
    url = tunnel.start()
    progress_print("")
    progress_print("=" * 56)
    progress_print("  ✅ 外网分享已开启（把下面链接发给别人即可）")
    progress_print(f"  {url}")
    progress_print("=" * 56)
    progress_print("  注意：")
    progress_print("  · 链接含随机地址，请勿随意公开到互联网")
    progress_print("  · 对方上传的文件会保存在你这台电脑上")
    progress_print("  · 你的电脑需保持开机且本程序保持运行")
    progress_print("")
    return tunnel


def print_lan_hints(port: int = 5000, progress_print=print):
    ips = get_lan_ips()
    if not ips:
        progress_print(f"  局域网: http://<你的电脑IP>:{port}")
        return
    progress_print("  局域网（同一 WiFi 下可直接访问）：")
    for ip in ips[:3]:
        progress_print(f"    http://{ip}:{port}")


if __name__ == "__main__":
    # 单独测试隧道（需本机 5000 已有服务，或仅测下载）
    if len(sys.argv) > 1 and sys.argv[1] == "download":
        print(ensure_cloudflared())
    else:
        t = start_share_tunnel()
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            t.stop()
