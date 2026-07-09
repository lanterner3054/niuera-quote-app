#!/usr/bin/env bash
# ── NIUERA 报价工作台 启动脚本 (Mac/Linux) ──
cd "$(dirname "$0")"
( sleep 2; (open http://127.0.0.1:8000 || xdg-open http://127.0.0.1:8000) >/dev/null 2>&1 ) &
python3 -m uvicorn app:app --host 127.0.0.1 --port 8000
