# -*- coding: utf-8 -*-
"""
config.json 结构升级(V4 方案 §2.3)。幂等:可重复执行,已是新结构时只补缺失键。
执行前自动备份为 config.json.bak_时间戳。

用法:  python tools/upgrade_config.py
"""
import os, sys, json, shutil, datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE, "data", "config.json")

# PI 默认条款(依据样例 pi__20260605_XNY-AD260045.pdf)
PI_TERM_OVERRIDES = {
    "price_term": "FOB SHANGHAI",
    "delivery_time": "Within 10 working days after receipt of the payment.",
    "package": "Standard Export Packing.",
    "validity": "30 days.",
}

# 占位模板:升级后请到「设置」页把真实银行信息补齐
DEFAULT_BANK_ACCOUNT = {
    "id": "bank_main",
    "label": "默认账户(请到设置页修改)",
    "ac_bank": "YOUR BANK NAME, BRANCH",
    "address": "YOUR BANK ADDRESS",
    "beneficiary": "YOUR COMPANY NAME",
    "account_no": "000000000000",
    "swift": "XXXXXXXX",
    "currency": "USD",
    "is_default": True,
}

DEFAULT_LEGACY_DIRS = []


def upgrade(config):
    changed = []
    defaults = config.get("defaults", {})
    if "quotation" not in defaults:
        # 旧扁平结构 -> 嵌套。字段名保持与现有代码一致(delivery_time 等)。
        flat = dict(defaults)
        pi_defaults = dict(flat)
        pi_defaults.update(PI_TERM_OVERRIDES)
        config["defaults"] = {"quotation": flat, "pi": pi_defaults}
        changed.append("defaults -> defaults.quotation / defaults.pi")
    if "sales_list" not in config:
        from_name = config["defaults"]["quotation"].get("from_name", "")
        config["sales_list"] = [from_name] if from_name else []
        changed.append("sales_list")
    if "bank_accounts" not in config:
        config["bank_accounts"] = [dict(DEFAULT_BANK_ACCOUNT)]
        changed.append("bank_accounts")
    if "stamp" not in config:
        config["stamp"] = {"path": "stamp.png", "enabled_default": True}
        changed.append("stamp")
    if "legacy_dirs" not in config:
        config["legacy_dirs"] = [d for d in DEFAULT_LEGACY_DIRS if os.path.isdir(d)]
        changed.append("legacy_dirs")
    return changed


def main():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = json.load(f)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = CONFIG_PATH + ".bak_" + ts
    shutil.copy2(CONFIG_PATH, bak)
    changed = upgrade(config)
    if not changed:
        os.remove(bak)
        print("config 已是新结构,无需升级。")
        return
    tmp = CONFIG_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=1)
    os.replace(tmp, CONFIG_PATH)
    print("已升级:", "; ".join(changed))
    print("备份:", os.path.basename(bak))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    main()
