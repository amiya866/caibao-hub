#!/bin/bash
# 财报跟踪平台可移植版 · Linux/macOS 一键启动
set -e
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt --quiet 2>/dev/null || true
python3 run.py "$@"
