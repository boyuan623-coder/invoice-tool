#!/usr/bin/env bash
# 腾讯云 / Linux 一键安装依赖并以前台或后台启动
# 用法:
#   cd /path/to/invoice-tool
#   bash deploy/start_cloud.sh           # 安装 + 前台运行
#   bash deploy/start_cloud.sh --daemon  # 安装 + nohup 后台
#   bash deploy/start_cloud.sh --install-only

set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

DAEMON=0
INSTALL_ONLY=0
for arg in "$@"; do
  case "$arg" in
    --daemon) DAEMON=1 ;;
    --install-only) INSTALL_ONLY=1 ;;
  esac
done

if [[ -f "$ROOT/deploy/invoice-tool.env" ]]; then
  set -a
  # shellcheck disable=SC1091
  source "$ROOT/deploy/invoice-tool.env"
  set +a
elif [[ -f "$ROOT/.env" ]]; then
  set -a
  # shellcheck disable=SC1091
  source "$ROOT/.env"
  set +a
fi

export INVOICE_TOOL_CLOUD="${INVOICE_TOOL_CLOUD:-1}"
export INVOICE_TOOL_LOW_MEM="${INVOICE_TOOL_LOW_MEM:-1}"
export INVOICE_TOOL_NO_BROWSER="${INVOICE_TOOL_NO_BROWSER:-1}"
export INVOICE_TOOL_DEVICE="${INVOICE_TOOL_DEVICE:-cpu}"
export INVOICE_TOOL_HOST="${INVOICE_TOOL_HOST:-0.0.0.0}"
export INVOICE_TOOL_PORT="${INVOICE_TOOL_PORT:-5000}"

echo "==> 项目目录: $ROOT"
echo "==> 端口: $INVOICE_TOOL_PORT  低内存: $INVOICE_TOOL_LOW_MEM"

MEM_MB="$(awk '/MemTotal/ {print int($2/1024)}' /proc/meminfo 2>/dev/null || echo 0)"
if [[ "$MEM_MB" -gt 0 && "$MEM_MB" -le 2500 ]]; then
  echo "==> 检测到约 ${MEM_MB}MB 内存（2G 档）"
  if ! swapon --show 2>/dev/null | grep -q .; then
    echo "[建议] 先执行: sudo bash deploy/setup_swap.sh"
    echo "       否则图片 OCR 很容易被系统 OOM 杀掉"
  fi
fi


if command -v apt-get >/dev/null 2>&1; then
  sudo apt-get update -y
  sudo DEBIAN_FRONTEND=noninteractive apt-get install -y \
    python3 python3-pip python3-venv libgl1 libglib2.0-0
fi

if [[ ! -d "$ROOT/.venv" ]]; then
  python3 -m venv "$ROOT/.venv"
fi
# shellcheck disable=SC1091
source "$ROOT/.venv/bin/activate"
pip install -U pip
pip install -r "$ROOT/requirements.txt" -i https://pypi.tuna.tsinghua.edu.cn/simple

mkdir -p "$ROOT/uploads" "$ROOT/excel"

if [[ "$INSTALL_ONLY" -eq 1 ]]; then
  echo "==> 依赖安装完成（--install-only）"
  exit 0
fi

if [[ -z "${INVOICE_TOOL_ACCESS_TOKEN:-}" ]]; then
  echo "[警告] 未设置 INVOICE_TOOL_ACCESS_TOKEN，公网任何人都能用。"
  echo "       可复制 deploy/env.example 为 deploy/invoice-tool.env 后填写口令。"
fi

LOW_MEM_FLAG=()
if [[ "${INVOICE_TOOL_LOW_MEM}" == "1" || "${INVOICE_TOOL_LOW_MEM}" == "true" ]]; then
  LOW_MEM_FLAG=(--low-mem)
fi

echo "==> 启动: python app.py --cloud ${LOW_MEM_FLAG[*]:-} --host $INVOICE_TOOL_HOST --port $INVOICE_TOOL_PORT"
if [[ "$DAEMON" -eq 1 ]]; then
  nohup python "$ROOT/app.py" --cloud "${LOW_MEM_FLAG[@]}" \
    --host "$INVOICE_TOOL_HOST" \
    --port "$INVOICE_TOOL_PORT" \
    > "$ROOT/app.log" 2>&1 &
  echo $! > "$ROOT/app.pid"
  echo "==> 已后台启动 PID=$(cat "$ROOT/app.pid")，日志: $ROOT/app.log"
  echo "==> 访问: http://<公网IP>:$INVOICE_TOOL_PORT"
else
  exec python "$ROOT/app.py" --cloud "${LOW_MEM_FLAG[@]}" \
    --host "$INVOICE_TOOL_HOST" \
    --port "$INVOICE_TOOL_PORT"
fi
