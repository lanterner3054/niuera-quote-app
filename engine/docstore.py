"""
NIUERA document store
---------------------
Archive layer for generated documents (quotation / pi).

Layout (under <data>/documents/):
    {no}__{doc_type}__r{revision}.json   full payload snapshot, immutable
    index.json                           list-page fields only, rebuildable

Concurrency: a single cross-process file lock (data/.docstore.lock) guards
number claiming + document save + index write. Lock file holds "pid ts";
stale locks (>60s) are removed automatically. Pure stdlib, Windows-safe.

All JSON writes are atomic (temp file + os.replace).
"""
import os, json, time, re, datetime

INDEX_VERSION = 1
SCHEMA_VERSION = 1
LOCK_STALE_SECONDS = 60
LOCK_RETRY_INTERVAL = 0.1
LOCK_RETRIES = 50


class DocStoreError(Exception):
    pass


class DuplicateKeyError(DocStoreError):
    pass


class DocStore:
    def __init__(self, data_dir):
        self.data_dir = data_dir
        self.docs_dir = os.path.join(data_dir, "documents")
        self.index_path = os.path.join(self.docs_dir, "index.json")
        self.lock_path = os.path.join(data_dir, ".docstore.lock")
        os.makedirs(self.docs_dir, exist_ok=True)

    # ---------- lock ----------
    def lock(self):
        return _FileLock(self.lock_path)

    # ---------- paths / keys ----------
    @staticmethod
    def doc_filename(no, doc_type, revision):
        safe_no = re.sub(r"[^A-Za-z0-9_.\-]", "_", str(no))
        return f"{safe_no}__{doc_type}__r{int(revision)}.json"

    def doc_path(self, no, doc_type, revision):
        return os.path.join(self.docs_dir, self.doc_filename(no, doc_type, revision))

    def exists(self, no, doc_type, revision):
        return os.path.exists(self.doc_path(no, doc_type, revision))

    # ---------- io ----------
    def load_index(self):
        if not os.path.exists(self.index_path):
            return {"version": INDEX_VERSION, "entries": []}
        try:
            with open(self.index_path, "r", encoding="utf-8") as f:
                idx = json.load(f)
            if not isinstance(idx.get("entries"), list):
                raise ValueError("bad index shape")
            return idx
        except Exception:
            # index corrupt -> rebuild from document files
            return self.rebuild_index()

    def get_document(self, no, doc_type, revision):
        p = self.doc_path(no, doc_type, revision)
        if not os.path.exists(p):
            return None
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)

    def save_document(self, doc, replace=False):
        """Persist one document + its index entry. Caller must hold the lock
        for write paths that also claim numbers. Raises DuplicateKeyError if
        the (no, doc_type, revision) key already exists and replace=False."""
        no, doc_type = doc["no"], doc["doc_type"]
        revision = int(doc.get("revision", 0))
        doc.setdefault("schema_version", SCHEMA_VERSION)
        doc.setdefault("created_at", _now_iso())
        p = self.doc_path(no, doc_type, revision)
        if os.path.exists(p) and not replace:
            raise DuplicateKeyError(f"{no} {doc_type} r{revision} already archived")
        _atomic_write_json(p, doc)
        idx = self.load_index()
        entry = self.index_entry(doc)
        entries = [e for e in idx["entries"]
                   if not (e["no"] == no and e["doc_type"] == doc_type
                           and int(e.get("revision", 0)) == revision)]
        entries.append(entry)
        idx["entries"] = entries
        _atomic_write_json(self.index_path, idx)
        return entry

    @staticmethod
    def index_entry(doc):
        src = doc.get("source") or {}
        return {
            "no": doc["no"],
            "doc_type": doc["doc_type"],
            "revision": int(doc.get("revision", 0)),
            "customer": doc.get("customer", ""),
            "from_name": doc.get("from_name", ""),
            "currency": doc.get("currency", ""),
            "date": doc.get("date", ""),
            "amount": (doc.get("totals") or {}).get("amount"),
            "source_no": src.get("no"),
            "archived_only": bool(doc.get("archived_only", False)),
            "file": doc.get("file"),
            "legacy_file": doc.get("legacy_file"),
            "created_at": doc.get("created_at", ""),
        }

    def rebuild_index(self):
        """Rebuild index.json from all per-document JSON files."""
        entries = []
        for name in os.listdir(self.docs_dir):
            if name == "index.json" or not name.endswith(".json"):
                continue
            try:
                with open(os.path.join(self.docs_dir, name), "r", encoding="utf-8") as f:
                    entries.append(self.index_entry(json.load(f)))
            except Exception:
                continue
        idx = {"version": INDEX_VERSION, "entries": entries}
        _atomic_write_json(self.index_path, idx)
        return idx

    # ---------- query ----------
    def query(self, doc_type=None, from_name=None, q=None,
              date_from=None, date_to=None, page=1, page_size=20):
        entries = self.load_index()["entries"]
        if doc_type:
            entries = [e for e in entries if e.get("doc_type") == doc_type]
        if from_name:
            entries = [e for e in entries if e.get("from_name") == from_name]
        if q:
            ql = str(q).lower()
            entries = [e for e in entries
                       if ql in str(e.get("no", "")).lower()
                       or ql in str(e.get("customer", "")).lower()]
        if date_from:
            df = _compact_date(date_from)
            entries = [e for e in entries if _compact_date(e.get("date")) >= df]
        if date_to:
            dt = _compact_date(date_to)
            entries = [e for e in entries if _compact_date(e.get("date")) <= dt]
        entries.sort(key=lambda e: (_compact_date(e.get("date")), str(e.get("no", "")),
                                    str(e.get("created_at", ""))), reverse=True)
        total = len(entries)
        page = max(1, int(page or 1))
        start = (page - 1) * page_size
        return {"total": total, "page": page, "page_size": page_size,
                "entries": entries[start:start + page_size]}

    def from_names(self):
        seen, out = set(), []
        for e in self.load_index()["entries"]:
            n = e.get("from_name") or ""
            if n and n not in seen:
                seen.add(n)
                out.append(n)
        return sorted(out)

    def max_seq(self, prefix):
        """Largest numeric sequence used by any indexed document number."""
        pat = re.compile(rf"^{re.escape(prefix)}(\d+)")
        mx = 0
        for e in self.load_index()["entries"]:
            m = pat.match(str(e.get("no", "")))
            if m:
                try:
                    mx = max(mx, int(m.group(1)))
                except ValueError:
                    pass
        return mx


class _FileLock:
    """Exclusive-create file lock with stale-lock cleanup (Windows-safe)."""

    def __init__(self, path):
        self.path = path
        self.fd = None

    def __enter__(self):
        for _ in range(LOCK_RETRIES):
            try:
                self.fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(self.fd, f"{os.getpid()} {time.time()}".encode())
                return self
            except FileExistsError:
                self._clear_if_stale()
                time.sleep(LOCK_RETRY_INTERVAL)
            except PermissionError:
                # Windows: file in delete-pending state while another holder
                # is releasing -- just retry.
                time.sleep(LOCK_RETRY_INTERVAL)
        raise DocStoreError("系统繁忙(锁等待超时),请稍后重试")

    def __exit__(self, *exc):
        if self.fd is not None:
            try:
                os.close(self.fd)
            except OSError:
                pass
            self.fd = None
        try:
            os.remove(self.path)
        except OSError:
            pass

    def _clear_if_stale(self):
        # Staleness is judged by mtime via os.stat only. Never open() the lock
        # file here: on Windows a concurrent read handle blocks the owner's
        # os.remove (no FILE_SHARE_DELETE), which would leak the lock forever.
        try:
            if time.time() - os.path.getmtime(self.path) > LOCK_STALE_SECONDS:
                os.remove(self.path)
        except OSError:
            pass


def _atomic_write_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def _now_iso():
    return datetime.datetime.now().astimezone().isoformat(timespec="seconds")


def _compact_date(s):
    """'2026/06/05' | '2026-06-05' | '20260605' -> '20260605' for comparison."""
    digits = re.sub(r"\D", "", str(s or ""))
    return digits[:8].ljust(8, "0")
