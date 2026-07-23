#!/bin/bash
# RFC3161 可信时间戳:每天给留档台账的链头打一个第三方时间戳。
# 效果:第三方 TSA 证明"此链头哈希在某时刻已存在",配合哈希链 =
# 台账在该时刻之前的所有条目都不可能是事后伪造的。
# 服务器 cron: 0 3 * * * /home/ubuntu/quote-app/tools/tsa_stamp.sh
# 查看某个戳: openssl ts -reply -in xxx.tsr -text
set -u
ARCH="${1:-/home/ubuntu/quote-app/data/archive}"
LEDGER="$ARCH/ledger.jsonl"
TSA_DIR="$ARCH/tsa"
mkdir -p "$TSA_DIR"
[ -s "$LEDGER" ] || exit 0   # 还没有任何留档

read -r chain seq <<EOF
$(tail -n 1 "$LEDGER" | python3 -c 'import sys,json;e=json.load(sys.stdin);print(e["chain"],e["seq"])')
EOF
[ -n "$chain" ] || exit 1

# 链头没变(当天没有新单)就不重复打戳
last=$(cat "$TSA_DIR/.last_chain" 2>/dev/null || true)
[ "$chain" = "$last" ] && exit 0

ts=$(date +%Y%m%d-%H%M%S)
base="$TSA_DIR/${ts}_seq${seq}"
printf '%s' "$chain" > "${base}.head"
tsq=$(mktemp)
openssl ts -query -data "${base}.head" -sha256 -cert -out "$tsq" || { rm -f "${base}.head" "$tsq"; exit 1; }

for url in https://freetsa.org/tsr http://timestamp.digicert.com; do
    if curl -sf -H 'Content-Type: application/timestamp-query' \
            --data-binary @"$tsq" "$url" -o "${base}.tsr" --max-time 30 \
       && openssl ts -reply -in "${base}.tsr" -text 2>/dev/null | grep -q 'Time stamp:'; then
        echo "$chain" > "$TSA_DIR/.last_chain"
        echo "$(date '+%F %T') seq=$seq tsa=$url OK" >> "$TSA_DIR/tsa.log"
        rm -f "$tsq"
        exit 0
    fi
    rm -f "${base}.tsr"
done
rm -f "${base}.head" "$tsq"
echo "$(date '+%F %T') seq=$seq FAILED(所有 TSA 均不可达,明天重试)" >> "$TSA_DIR/tsa.log"
exit 1
