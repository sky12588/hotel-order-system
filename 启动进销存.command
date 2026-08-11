#!/bin/bash
# 酒店订单管理系统启动脚本

cd "$(dirname "$0")"
if [ -f .env ]; then
  set -a
  source .env
  set +a
fi
export PORT="${PORT:-5011}"
echo "正在启动酒店订单管理系统..."
python3 server.py
