#!/usr/bin/env bash
# 2G 小机强烈建议先加 2G Swap，再启动服务
# 用法: sudo bash deploy/setup_swap.sh

set -euo pipefail

SWAPFILE="${SWAPFILE:-/swapfile}"
SIZE_GB="${SIZE_GB:-2}"

if swapon --show | grep -q .; then
  echo "==> 已存在 Swap:"
  swapon --show
  free -h
  exit 0
fi

echo "==> 创建 ${SIZE_GB}G Swap: $SWAPFILE"
fallocate -l "${SIZE_GB}G" "$SWAPFILE" 2>/dev/null || dd if=/dev/zero of="$SWAPFILE" bs=1M count=$((SIZE_GB * 1024))
chmod 600 "$SWAPFILE"
mkswap "$SWAPFILE"
swapon "$SWAPFILE"

if ! grep -q "$SWAPFILE" /etc/fstab 2>/dev/null; then
  echo "$SWAPFILE none swap sw 0 0" >> /etc/fstab
fi

echo "==> 完成"
free -h
