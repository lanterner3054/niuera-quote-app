#!/usr/bin/env bash
# 在服务器上执行的安装脚本(由本机部署流程通过 ssh 触发)
# 前提:代码包 quote-app.tar.gz 与数据包 quote-data.tar.gz 已 scp 到 /home/ubuntu/
set -e
cd /home/ubuntu

echo "== 解包 =="
mkdir -p quote-app
tar -xzf quote-app.tar.gz -C quote-app
tar -xzf quote-data.tar.gz -C quote-app   # data/ output/ legacy/

cd quote-app

echo "== 修正旧档路径(容器视角 /app)=="
python3 tools/fix_legacy_paths.py --base /app || \
  docker run --rm -v "$PWD":/app -w /app python:3.12-slim python tools/fix_legacy_paths.py --base /app

echo "== 构建并启动 =="
docker compose up -d --build

echo "== 装备份 cron =="
chmod +x backup.sh
( crontab -l 2>/dev/null | grep -v quote-app/backup.sh ; \
  echo "30 2 * * * /home/ubuntu/quote-app/backup.sh >> /home/ubuntu/backups/quote-app/backup.log 2>&1" ) | crontab -

echo "== 自检 =="
sleep 3
docker compose ps
curl -sk https://127.0.0.1:9443/api/bootstrap | head -c 200; echo
echo "DONE"
