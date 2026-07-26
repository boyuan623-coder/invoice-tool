"""
固定域名分享：Cloudflare Named Tunnel
默认主机名: invoice.bingshanforprivate.asia
"""

from __future__ import annotations

import os
import platform
import re
import subprocess
import sys
import threading
import time
from typing import Optional

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOOLS_DIR = os.path.join(SCRIPT_DIR, "tools")
DEFAULT_HOSTNAME = "invoice.bingshanforprivate.asia"
DEFAULT_TUNNEL_NAME = "invoice-tool"
CONFIG_PATH = os.path.join(TOOLS_DIR, "tunnel-config.yml")


def _cloudflared_bin() -> str:
    from tunnel_share import ensure_cloudflared
    return ensure_cloudflared()


def _user_cloudflared_dir() -> str:
    return os.path.join(os.path.expanduser("~"), ".cloudflared")


def config_ready() -> bool:
    return os.path.isfile(CONFIG_PATH)


def read_hostname_from_config() -> str:
    if not os.path.isfile(CONFIG_PATH):
        return DEFAULT_HOSTNAME
    try:
        text = open(CONFIG_PATH, encoding="utf-8").read()
        m = re.search(r"hostname:\s*([^\s#]+)", text)
        if m:
            return m.group(1).strip()
    except OSError:
        pass
    return DEFAULT_HOSTNAME


def write_tunnel_config(tunnel_id: str, hostname: str, port: int = 5000) -> str:
    cred = os.path.join(_user_cloudflared_dir(), f"{tunnel_id}.json")
    # Windows 路径在 yaml 里用正斜杠更稳
    cred_yaml = cred.replace("\\", "/")
    content = (
        f"tunnel: {tunnel_id}\n"
        f"credentials-file: {cred_yaml}\n"
        f"\n"
        f"ingress:\n"
        f"  - hostname: {hostname}\n"
        f"    service: http://127.0.0.1:{port}\n"
        f"  - service: http_status:404\n"
    )
    os.makedirs(TOOLS_DIR, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)
    return CONFIG_PATH


def _run(cmd: list[str], check: bool = True) -> subprocess.CompletedProcess:
    creationflags = 0
    if platform.system().lower() == "windows":
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    # login / create 需要可见窗口交互时不要隐藏
    return subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=creationflags,
    )


def _run_visible(cmd: list[str]) -> int:
    """需要浏览器登录时用可见窗口。"""
    print("  [固定分享] 执行:", " ".join(cmd))
    return subprocess.call(cmd)


def list_tunnels() -> str:
    bin_path = _cloudflared_bin()
    r = _run([bin_path, "tunnel", "list"])
    return (r.stdout or "") + (r.stderr or "")


def find_tunnel_id(name: str = DEFAULT_TUNNEL_NAME) -> Optional[str]:
    text = list_tunnels()
    # 典型行: UUID  NAME  CREATED
    for line in text.splitlines():
        if name in line:
            m = re.search(
                r"([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})",
                line,
                re.I,
            )
            if m:
                return m.group(1)
    return None


def setup_fixed_tunnel(
    hostname: str = DEFAULT_HOSTNAME,
    tunnel_name: str = DEFAULT_TUNNEL_NAME,
    port: int = 5000,
) -> str:
    """
    交互式初始化（需已在 Cloudflare 添加域名并改好 NS）。
    返回固定访问 URL。
    """
    bin_path = _cloudflared_bin()
    print("=" * 60)
    print("  固定域名隧道初始化")
    print(f"  域名: {hostname}")
    print("=" * 60)
    print()
    print("  使用前请确认：")
    print("  1. 域名已添加到 Cloudflare")
    print("  2. 域名商处 NS 已改为 Cloudflare 提供的两行 nameserver")
    print("  3. 等待 DNS 生效后（可能几分钟到几小时）再继续")
    print()
    input("  已完成后按回车继续登录 Cloudflare ... ")

    # 登录（会打开浏览器）
    code = subprocess.call([bin_path, "tunnel", "login"])
    if code != 0:
        raise RuntimeError("cloudflared tunnel login 失败")

    tunnel_id = find_tunnel_id(tunnel_name)
    if not tunnel_id:
        print(f"  [固定分享] 创建隧道: {tunnel_name}")
        r = subprocess.run(
            [bin_path, "tunnel", "create", tunnel_name],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        out = (r.stdout or "") + (r.stderr or "")
        print(out)
        if r.returncode != 0:
            raise RuntimeError("创建隧道失败")
        tunnel_id = find_tunnel_id(tunnel_name)
        if not tunnel_id:
            m = re.search(
                r"([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})",
                out,
                re.I,
            )
            tunnel_id = m.group(1) if m else None
    else:
        print(f"  [固定分享] 已存在隧道 {tunnel_name}: {tunnel_id}")

    if not tunnel_id:
        raise RuntimeError("无法获取 tunnel id，请手动执行: cloudflared tunnel list")

    print(f"  [固定分享] 绑定 DNS: {hostname} -> {tunnel_name}")
    r = subprocess.run(
        [bin_path, "tunnel", "route", "dns", tunnel_name, hostname],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    print((r.stdout or "") + (r.stderr or ""))
    # DNS 已存在时也可能非 0，继续写配置

    cfg = write_tunnel_config(tunnel_id, hostname, port=port)
    print(f"  [固定分享] 配置已写入: {cfg}")
    url = f"https://{hostname}"
    print()
    print("=" * 60)
    print(f"  固定网址: {url}")
    print("  以后请双击 invoice_tool_fixed.bat 启动")
    print("=" * 60)
    return url


class FixedShareTunnel:
    def __init__(self, port: int = 5000):
        self.port = port
        self.public_url: Optional[str] = None
        self._proc: Optional[subprocess.Popen] = None

    def start(self) -> str:
        if not config_ready():
            raise RuntimeError(
                "尚未配置固定隧道。请先运行: python tunnel_fixed.py setup\n"
                "或双击 setup_fixed_tunnel.bat"
            )
        # 确保 config 里的端口与当前一致
        hostname = read_hostname_from_config()
        # 若端口变了，轻量改写 service 行
        try:
            text = open(CONFIG_PATH, encoding="utf-8").read()
            text2 = re.sub(
                r"service:\s*http://127\.0\.0\.1:\d+",
                f"service: http://127.0.0.1:{self.port}",
                text,
            )
            if text2 != text:
                open(CONFIG_PATH, "w", encoding="utf-8", newline="\n").write(text2)
        except OSError:
            pass

        bin_path = _cloudflared_bin()
        creationflags = 0
        if platform.system().lower() == "windows":
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        self._proc = subprocess.Popen(
            [bin_path, "tunnel", "--config", CONFIG_PATH, "run"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=creationflags,
        )
        self.public_url = f"https://{hostname}"
        threading.Thread(target=self._drain, daemon=True).start()

        # 稍等确认进程未立刻退出
        time.sleep(1.5)
        if self._proc.poll() is not None:
            leftover = ""
            try:
                leftover = self._proc.stdout.read() if self._proc.stdout else ""
            except Exception:
                pass
            raise RuntimeError(f"固定隧道启动失败:\n{leftover[-800:]}")
        return self.public_url

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


def start_fixed_share_tunnel(port: int = 5000, progress_print=print) -> FixedShareTunnel:
    progress_print("")
    progress_print("  [固定分享] 正在启动 Named Tunnel ...")
    tunnel = FixedShareTunnel(port=port)
    url = tunnel.start()
    progress_print("")
    progress_print("=" * 56)
    progress_print("  固定网址（可收藏，长期不变）:")
    progress_print(f"  {url}")
    progress_print("=" * 56)
    progress_print("  注意: 关闭本窗口后，外网将无法访问")
    progress_print("")
    return tunnel


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "setup":
        host = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_HOSTNAME
        setup_fixed_tunnel(hostname=host)
    elif len(sys.argv) > 1 and sys.argv[1] == "status":
        print("config:", CONFIG_PATH, "ready=", config_ready())
        print("hostname:", read_hostname_from_config())
        print(list_tunnels())
    else:
        print("用法:")
        print("  python tunnel_fixed.py setup")
        print("  python tunnel_fixed.py setup invoice.bingshanforprivate.asia")
        print("  python tunnel_fixed.py status")
