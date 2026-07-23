# -*- coding: utf-8 -*-
import json
import os

import pytest

from engine.archive import archive_document, verify_ledger, sha256_file


@pytest.fixture
def data_dir(tmp_path):
    return str(tmp_path)


def _make_doc(no="XNY-AD0001", rev=0):
    return {"no": no, "doc_type": "quotation", "revision": rev,
            "customer": "ACME", "items": [{"item": "A", "qty": 1, "unit_price": 2}]}


def _make_file(tmp_path, name="q.xlsx", content=b"fake-xlsx-bytes"):
    p = tmp_path / name
    p.write_bytes(content)
    return str(p)


def test_archive_creates_timestamped_copy_and_snapshot(data_dir, tmp_path):
    src = _make_file(tmp_path)
    entry = archive_document(data_dir, _make_doc(), src)
    assert entry is not None
    arch = os.path.join(data_dir, "archive")
    full = os.path.join(arch, entry["file"])
    assert os.path.exists(full)
    assert os.path.basename(full).startswith("quotation_XNY-AD0001_R0_")
    assert entry["sha256"] == sha256_file(src)
    snap = os.path.join(arch, entry["snapshot"])
    assert json.load(open(snap, encoding="utf-8"))["customer"] == "ACME"


def test_same_content_archived_once(data_dir, tmp_path):
    src = _make_file(tmp_path)
    assert archive_document(data_dir, _make_doc(), src) is not None
    assert archive_document(data_dir, _make_doc(), src) is None  # 幂等
    ok, problems, count = verify_ledger(data_dir)
    assert ok and count == 1


def test_chain_links_and_verify_ok(data_dir, tmp_path):
    e1 = archive_document(data_dir, _make_doc(), _make_file(tmp_path, "a.xlsx", b"v1"))
    e2 = archive_document(data_dir, _make_doc(rev=1), _make_file(tmp_path, "b.xlsx", b"v2"))
    assert e2["prev"] == e1["chain"]
    ok, problems, count = verify_ledger(data_dir)
    assert ok and count == 2 and problems == []


def test_tampered_file_detected(data_dir, tmp_path):
    entry = archive_document(data_dir, _make_doc(), _make_file(tmp_path))
    full = os.path.join(data_dir, "archive", entry["file"])
    with open(full, "ab") as f:
        f.write(b"tamper")
    ok, problems, _ = verify_ledger(data_dir)
    assert not ok
    assert any("内容被改动" in p for p in problems)


def test_tampered_ledger_entry_detected(data_dir, tmp_path):
    archive_document(data_dir, _make_doc(), _make_file(tmp_path, "a.xlsx", b"v1"))
    archive_document(data_dir, _make_doc(rev=1), _make_file(tmp_path, "b.xlsx", b"v2"))
    ledger = os.path.join(data_dir, "archive", "ledger.jsonl")
    lines = open(ledger, encoding="utf-8").read().splitlines()
    e = json.loads(lines[0])
    e["no"] = "FORGED-999"  # 事后篡改台账内容
    lines[0] = json.dumps(e, ensure_ascii=False, sort_keys=True)
    open(ledger, "w", encoding="utf-8").write("\n".join(lines) + "\n")
    ok, problems, _ = verify_ledger(data_dir)
    assert not ok
    assert any("篡改" in p for p in problems)


def test_deleted_ledger_entry_detected(data_dir, tmp_path):
    archive_document(data_dir, _make_doc(), _make_file(tmp_path, "a.xlsx", b"v1"))
    archive_document(data_dir, _make_doc(rev=1), _make_file(tmp_path, "b.xlsx", b"v2"))
    ledger = os.path.join(data_dir, "archive", "ledger.jsonl")
    lines = open(ledger, encoding="utf-8").read().splitlines()
    open(ledger, "w", encoding="utf-8").write(lines[1] + "\n")  # 删掉第一条
    ok, problems, _ = verify_ledger(data_dir)
    assert not ok
