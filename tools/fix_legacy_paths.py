# -*- coding: utf-8 -*-
"""
云端迁移路径修正:把存档里指向 Windows 本地的 legacy_file
(D:\\报价\\PI\\...)改写为服务器路径(/app/legacy/PI/...),
并同步更新 config.json 的 legacy_dirs,最后重建索引。

幂等,可重复执行。在服务器容器内或宿主机(python3)均可运行:
    python3 tools/fix_legacy_paths.py [--base /app]
"""
import os, sys, json, glob, argparse

MAPPING = [
    ("D:\\报价\\PI\\PI Excel\\", "{base}/legacy/PI/PI Excel/"),
    ("D:\\报价\\PI\\", "{base}/legacy/PI/"),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default="/app", help="应用根目录(容器内为 /app)")
    args = ap.parse_args()
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    sys.path.insert(0, root)
    from engine.docstore import DocStore

    changed = 0
    for p in glob.glob(os.path.join(root, "data", "documents", "*.json")):
        if os.path.basename(p) == "index.json":
            continue
        with open(p, "r", encoding="utf-8") as f:
            doc = json.load(f)
        lf = doc.get("legacy_file")
        if not lf:
            continue
        new = lf
        for old_prefix, new_prefix in MAPPING:
            if lf.startswith(old_prefix):
                new = new_prefix.format(base=args.base) + lf[len(old_prefix):]
                break
        if new != lf:
            doc["legacy_file"] = new.replace("\\", "/")
            tmp = p + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(doc, f, ensure_ascii=False, indent=1)
            os.replace(tmp, p)
            changed += 1

    cfg_path = os.path.join(root, "data", "config.json")
    with open(cfg_path, "r", encoding="utf-8") as f:
        config = json.load(f)
    new_dirs = [f"{args.base}/legacy/PI", f"{args.base}/legacy/PI/PI Excel"]
    if config.get("legacy_dirs") != new_dirs:
        config["legacy_dirs"] = new_dirs
        tmp = cfg_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=1)
        os.replace(tmp, cfg_path)

    ds = DocStore(os.path.join(root, "data"))
    idx = ds.rebuild_index()
    print(f"legacy_file 改写 {changed} 条; legacy_dirs -> {new_dirs}; 索引重建 {len(idx['entries'])} 条")


if __name__ == "__main__":
    main()
