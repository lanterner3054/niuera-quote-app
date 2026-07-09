#!/usr/bin/env bash
# 每日备份 quote-app 的 data/ 与 output/,保留最近 30 份
# crontab: 30 2 * * * /home/ubuntu/quote-app/backup.sh >> /home/ubuntu/backups/quote-app/backup.log 2>&1
set -e
APP_DIR="/home/ubuntu/quote-app"
BAK_DIR="/home/ubuntu/backups/quote-app"
mkdir -p "$BAK_DIR"
TS=$(date +%Y%m%d_%H%M%S)
tar -czf "$BAK_DIR/quote-app_$TS.tar.gz" -C "$APP_DIR" data output
ls -1t "$BAK_DIR"/quote-app_*.tar.gz | tail -n +31 | xargs -r rm -f
echo "$(date '+%F %T') backup ok: quote-app_$TS.tar.gz"
