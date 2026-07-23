# -*- coding: utf-8 -*-
"""校验留档台账: python tools/verify_ledger.py [data 目录,默认 ./data]

逐条重算哈希链与留档文件 SHA-256;任何事后改动/删除都会在这里现形。
服务器上跑: docker exec quote-app python tools/verify_ledger.py /app/data
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from engine.archive import verify_ledger  # noqa: E402


def main():
    data_dir = sys.argv[1] if len(sys.argv) > 1 else "data"
    ok, problems, count = verify_ledger(data_dir)
    print(f"台账条目: {count}")
    if ok:
        print("✔ 哈希链完整,所有留档文件与登记哈希一致")
        return 0
    print(f"✘ 发现 {len(problems)} 处问题:")
    for p in problems:
        print("  -", p)
    return 1


if __name__ == "__main__":
    sys.exit(main())
