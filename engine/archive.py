# -*- coding: utf-8 -*-
"""生成留档 + 防篡改台账。

每次生成单据(xlsx/zip/pdf)时把文件复制进 data/archive/YYYY/MM/,
文件名带生成时刻时间戳: {doc_type}_{no}_R{rev}_{YYYYMMDD-HHMMSS}.{ext};
主文件(xlsx/zip)同时落一份单据数据快照 json。

台账 data/archive/ledger.jsonl 采用哈希链:每条记录含上一条的链哈希,
任何一条被事后修改/删除都会使后续整条链校验失败(校验: tools/verify_ledger.py)。
"""
import datetime
import hashlib
import json
import os
import re
import shutil
import threading

_LOCK = threading.Lock()
GENESIS = "GENESIS"


def _safe(s):
    return re.sub(r"[^A-Za-z0-9_.\-]", "_", str(s or "").strip()) or "NA"


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def chain_hash(prev, entry_without_chain):
    canon = json.dumps(entry_without_chain, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256((prev + "|" + canon).encode("utf-8")).hexdigest()


def _iter_ledger(ledger_path):
    if not os.path.exists(ledger_path):
        return
    with open(ledger_path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                yield json.loads(line)


def _ledger_tail(ledger_path):
    seq, prev = 0, GENESIS
    for e in _iter_ledger(ledger_path) or ():
        seq = e.get("seq", seq + 1)
        prev = e.get("chain", prev)
    return seq, prev


def archive_document(data_dir, doc, file_path, with_snapshot=True):
    """留档一个生成产物。幂等:同 no/rev 且内容哈希相同的产物只记一次。
    留档失败不应阻断业务,调用方可裸调;返回台账条目 dict 或 None(跳过)。"""
    with _LOCK:
        arch_root = os.path.join(data_dir, "archive")
        ledger = os.path.join(arch_root, "ledger.jsonl")
        digest = sha256_file(file_path)
        no = str(doc.get("no") or "NA")
        rev = int(doc.get("revision") or 0)
        doc_type = str(doc.get("doc_type") or "doc")

        for e in _iter_ledger(ledger) or ():
            if e.get("no") == no and e.get("revision") == rev and e.get("sha256") == digest:
                return None  # 同内容已留档

        now = datetime.datetime.now()
        ts = now.strftime("%Y%m%d-%H%M%S")
        sub = os.path.join(arch_root, now.strftime("%Y"), now.strftime("%m"))
        os.makedirs(sub, exist_ok=True)
        ext = (os.path.splitext(file_path)[1].lstrip(".") or "bin").lower()
        base = f"{doc_type}_{_safe(no)}_R{rev}_{ts}"
        dest = os.path.join(sub, f"{base}.{ext}")
        i = 1
        while os.path.exists(dest):
            dest = os.path.join(sub, f"{base}-{i}.{ext}")
            i += 1
        shutil.copy2(file_path, dest)

        snap_rel = snap_digest = None
        if with_snapshot:
            snap_path = os.path.splitext(dest)[0] + ".json"
            with open(snap_path, "w", encoding="utf-8") as f:
                json.dump(doc, f, ensure_ascii=False, sort_keys=True, indent=1)
            snap_rel = os.path.relpath(snap_path, arch_root).replace("\\", "/")
            snap_digest = sha256_file(snap_path)

        seq, prev = _ledger_tail(ledger)
        entry = {
            "seq": seq + 1,
            "ts": now.strftime("%Y-%m-%d %H:%M:%S"),
            "doc_type": doc_type, "no": no, "revision": rev,
            "file": os.path.relpath(dest, arch_root).replace("\\", "/"),
            "sha256": digest,
            "snapshot": snap_rel, "snapshot_sha256": snap_digest,
            "prev": prev,
        }
        entry["chain"] = chain_hash(prev, entry)
        with open(ledger, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False, sort_keys=True) + "\n")
        return entry


def verify_ledger(data_dir):
    """全量校验:哈希链连续性 + 每个留档文件/快照的内容哈希。
    返回 (ok, problems:list[str], count)。"""
    arch_root = os.path.join(data_dir, "archive")
    ledger = os.path.join(arch_root, "ledger.jsonl")
    problems = []
    prev, count = GENESIS, 0
    expected_seq = 0
    for e in _iter_ledger(ledger) or ():
        count += 1
        expected_seq += 1
        tag = f"seq={e.get('seq')} no={e.get('no')} file={e.get('file')}"
        if e.get("seq") != expected_seq:
            problems.append(f"{tag}: 序号断档(期望 {expected_seq}),可能有条目被删除")
            expected_seq = e.get("seq", expected_seq)
        if e.get("prev") != prev:
            problems.append(f"{tag}: 链头不衔接,前序条目被改动或删除")
        body = {k: v for k, v in e.items() if k != "chain"}
        if chain_hash(e.get("prev", ""), body) != e.get("chain"):
            problems.append(f"{tag}: 本条内容与链哈希不符,条目被篡改")
        for key, hkey in (("file", "sha256"), ("snapshot", "snapshot_sha256")):
            rel = e.get(key)
            if not rel:
                continue
            full = os.path.join(arch_root, rel)
            if not os.path.exists(full):
                problems.append(f"{tag}: 留档文件缺失 {rel}")
            elif sha256_file(full) != e.get(hkey):
                problems.append(f"{tag}: 留档文件内容被改动 {rel}")
        prev = e.get("chain", prev)
    return (not problems), problems, count
