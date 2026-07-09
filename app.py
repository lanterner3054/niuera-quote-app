"""
NIUERA Quote Workbench - local web backend (FastAPI).

Run:  python -m uvicorn app:app --host 127.0.0.1 --port 8000
Then open http://127.0.0.1:8000 in your browser.
(The start script does this and opens the browser for you.)

V4: dual document types (quotation / pi), archive layer + history,
convert-to-pi / revise flows. Write path is single: POST /api/generate.
"""
import os, sys, json, shutil, datetime, subprocess, re
from fastapi import FastAPI, Body, UploadFile, File, Form, Query
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

def _base_dir():
    # When packaged by PyInstaller (--onefile), bundled files live in sys._MEIPASS.
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))

def _exe_dir():
    # Directory of the running exe (for external, editable data folders).
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

BASE = _base_dir()
# Prefer an external editable data/ next to the exe; fall back to the bundled one.
_ext_data = os.path.join(_exe_dir(), "data")
DATA = _ext_data if os.path.isdir(_ext_data) else os.path.join(BASE, "data")
STATIC = os.path.join(BASE, "static")
OUTPUT = os.path.join(_exe_dir(), "output")
os.makedirs(OUTPUT, exist_ok=True)

from engine.quote_engine import generate_quote
from engine.pi_engine import generate_pi, compute_total
from engine.amount_words import amount_to_words
from engine.docstore import DocStore, DuplicateKeyError, DocStoreError

app = FastAPI(title="NIUERA Quote Workbench")
DOCSTORE = DocStore(DATA)

TERM_KEYS = ("origin", "payment", "price_term", "delivery_time", "package", "validity")
BANK_REQUIRED = ("ac_bank", "address", "beneficiary", "account_no", "swift")


def _load(name):
    with open(os.path.join(DATA, name), "r", encoding="utf-8") as f:
        return json.load(f)


def _save(name, obj):
    path = os.path.join(DATA, name)
    if os.path.exists(path):
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        bdir = os.path.join(DATA, "backups")
        os.makedirs(bdir, exist_ok=True)
        shutil.copy2(path, os.path.join(bdir, f"{name}.{ts}.bak"))
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


def get_defaults(config, doc_type="quotation"):
    """Read per-doc-type defaults; tolerates the legacy flat structure."""
    d = config.get("defaults", {})
    if "quotation" in d or "pi" in d:
        base = dict(d.get("quotation", {}))
        if doc_type == "pi":
            base.update(d.get("pi", {}))
        return base
    return dict(d)  # legacy flat config


def _int_or(v, default):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def _seq_width(v):
    s = str(v or "")
    return len(s) if s.isdigit() and s.startswith("0") else 0


def _format_seq(seq, width):
    return f"{seq:0{width}d}" if width else str(seq)


def _max_used_pi_seq(prefix):
    """Largest sequence seen anywhere: output/ filenames + document index."""
    max_seq = DOCSTORE.max_seq(prefix)
    if prefix and os.path.isdir(OUTPUT):
        pat = re.compile(rf"{re.escape(prefix)}(\d+)")
        for name in os.listdir(OUTPUT):
            for m in pat.finditer(name):
                max_seq = max(max_seq, _int_or(m.group(1), 0))
    return max_seq


def _peek_next_no(config):
    """Compute the next number WITHOUT persisting. Commit via _commit_seq."""
    pi = config.get("pi", {})
    prefix = pi.get("prefix", "XNY-AD")
    configured = _int_or(pi.get("next_seq"), 1)
    seq = max(configured, _max_used_pi_seq(prefix) + 1)
    return f"{prefix}{_format_seq(seq, _seq_width(pi.get('next_seq')))}", seq


def _commit_seq(config, seq):
    config.setdefault("pi", {})["next_seq"] = seq + 1
    _save("config.json", config)


@app.get("/", response_class=HTMLResponse)
def index():
    with open(os.path.join(STATIC, "index.html"), "r", encoding="utf-8") as f:
        return f.read()


@app.get("/api/bootstrap")
def bootstrap():
    config = _load("config.json")
    products = _load("products.json")
    return {
        "company": config.get("company", {}),
        # keep flat "defaults" = quotation defaults for backward compatibility
        "defaults": get_defaults(config, "quotation"),
        "defaults_pi": get_defaults(config, "pi"),
        "sales_list": config.get("sales_list", []),
        "bank_accounts": config.get("bank_accounts", []),
        "stamp": config.get("stamp", {}),
        "from_names": DOCSTORE.from_names(),
        "next_pi": _peek_next_no(config)[0],
        "today": datetime.date.today().strftime("%Y/%m/%d"),
        "products": products,
    }


def _resolve_items(mode, items, products):
    pmap = {p["id"]: p for p in products}
    out = []
    for it in items:
        base = pmap.get(it.get("id"), {})
        row = {
            "item": it.get("item", base.get("item", "")),
            "desc": it.get("desc", base.get("desc", "")),
            "unit": it.get("unit", base.get("unit", "PCS")),
            "image": it.get("image", base.get("image", "")),
        }
        if mode == "tiered":
            row.update({
                "band1": it.get("band1", base.get("band1")),
                "price1": _num(it.get("price1", base.get("price1"))),
                "band2": it.get("band2", base.get("band2")),
                "price2": _num(it.get("price2", base.get("price2"))),
            })
        else:
            row.update({
                "qty": _num(it.get("qty")),
                "unit_price": _num(it.get("unit_price", base.get("price1"))),
            })
        out.append(row)
    return out


def _num(v):
    if v in (None, "", "None"):
        return None
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return None


def _safe(s):
    return re.sub(r"[^A-Za-z0-9_.\-]", "_", str(s or "").strip())


def _price_error(mode, items):
    for idx, it in enumerate(items, 1):
        label = it.get("item") or f"第 {idx} 项"
        if not str(it.get("item") or "").strip():
            return f"第 {idx} 项产品名称为空,请先填写产品名称。"
        if mode == "tiered":
            if it.get("price1") in (None, ""):
                return f"第 {idx} 项 {label} 的价格档①为空,请先填写价格。"
            if it.get("band2") not in (None, "") and it.get("price2") in (None, ""):
                return f"第 {idx} 项 {label} 的价格档②为空,请先填写价格或清空档②。"
        elif it.get("unit_price") in (None, ""):
            return f"第 {idx} 项 {label} 的单价为空,请先填写价格。"
    return None


def _source_triple(payload):
    s = payload.get("source")
    if isinstance(s, dict) and s.get("no"):
        return {"no": s.get("no"), "doc_type": s.get("doc_type", "quotation"),
                "revision": _int_or(s.get("revision"), 0)}
    return None


def _doc_filename(doc_type, date, no, revision):
    date_compact = str(date).replace("/", "").replace("-", "")
    rev_sfx = f"-R{revision}" if revision else ""
    return f"{doc_type}_{date_compact}_{_safe(no) or 'NA'}{rev_sfx}.xlsx"


@app.post("/api/generate")
def generate(payload: dict = Body(...)):
    doc_type = payload.get("doc_type", "quotation")
    try:
        if doc_type == "pi":
            return _generate_pi(payload)
        return _generate_quotation(payload)
    except DuplicateKeyError:
        hint = "PI" if doc_type == "pi" else "报价"
        return JSONResponse({"error": f"该编号的{hint}已存在。如需修改请在历史页用「复制为新版」。"},
                            status_code=409)
    except DocStoreError as e:
        return JSONResponse({"error": str(e)}, status_code=503)


def _generate_quotation(payload):
    config = _load("config.json")
    products = _load("products.json")
    defaults = get_defaults(config, "quotation")
    mode = payload.get("mode", "standard")
    if not payload.get("customer"):
        return JSONResponse({"error": "请填写客户公司 (Customer Company)"}, status_code=400)
    items_in = payload.get("items", [])
    if not items_in:
        return JSONResponse({"error": "请至少添加一个产品"}, status_code=400)

    resolved_items = _resolve_items(mode, items_in, products)
    err = _price_error(mode, resolved_items)
    if err:
        return JSONResponse({"error": err}, status_code=400)

    revision = max(0, _int_or(payload.get("revision"), 0))
    quote = {
        "mode": mode,
        "customer": payload.get("customer", ""),
        "contact": payload.get("contact", ""),
        "pi_no": payload.get("pi_no", ""),
        "date": payload.get("date", datetime.date.today().strftime("%Y/%m/%d")),
        "from_name": payload.get("from_name", defaults.get("from_name", "")),
        "currency": payload.get("currency", defaults.get("currency", "USD")),
        "items": resolved_items,
        "terms": payload.get("terms", defaults),
    }

    total = None
    if mode == "standard":
        total = 0.0
        for it in resolved_items:
            if it.get("qty") not in (None, "") and it.get("unit_price") not in (None, ""):
                total += float(it["qty"]) * float(it["unit_price"])
        total = round(total, 2)

    with DOCSTORE.lock():
        config = _load("config.json")
        claimed_seq = None
        if payload.get("advance_pi"):
            quote["pi_no"], claimed_seq = _peek_next_no(config)
        if DOCSTORE.exists(quote["pi_no"], "quotation", revision):
            raise DuplicateKeyError(quote["pi_no"])
        fname = _doc_filename("quotation", quote["date"], quote["pi_no"], revision)
        out_path = os.path.join(OUTPUT, fname)
        generate_quote(config, quote, out_path, base_dir=DATA)
        doc = {
            "no": quote["pi_no"], "doc_type": "quotation", "revision": revision,
            "source": _source_triple(payload),
            "date": quote["date"], "from_name": quote["from_name"],
            "currency": quote["currency"],
            "customer": quote["customer"], "contact": quote["contact"],
            "mode": mode, "items": resolved_items, "terms": quote["terms"],
            "totals": {"amount": total, "amount_words": None},
            "file": os.path.join("output", fname).replace("\\", "/"),
            "legacy_file": None, "archived_only": False,
        }
        DOCSTORE.save_document(doc)
        if claimed_seq is not None:
            _commit_seq(config, claimed_seq)

    return _file_or_pdf_response(payload, out_path, fname, quote["pi_no"])


def _pi_validation_error(items, total, bank, customer):
    """Plan §7 pre-generation checks for PI."""
    if not customer:
        return "请填写客户抬头 (Customer Company)"
    if not items:
        return "PI 不能为空单,请至少添加一行明细"
    for idx, it in enumerate(items, 1):
        if it.get("line_type") == "fee":
            if not str(it.get("item") or "").strip():
                return f"费用行第 {idx} 行名称为空"
            if it.get("amount") is None or float(it["amount"]) < 0:
                return f"费用行第 {idx} 行金额不完整"
        else:
            if not str(it.get("item") or "").strip():
                return f"第 {idx} 行产品名称为空"
            if not it.get("qty") or float(it["qty"]) <= 0:
                return f"第 {idx} 行数量缺失或为 0"
            if not it.get("unit_price") or float(it["unit_price"]) <= 0:
                return f"第 {idx} 行单价缺失或为 0"
    if total <= 0:
        return "合计为 0,不能开 PI"
    if not bank:
        return "未找到银行账户,请到设置页添加"
    for k in BANK_REQUIRED:
        if not str(bank.get(k) or "").strip():
            return "银行信息不完整,请到设置页补全(户行/地址/收款人/账号/SWIFT)"
    return None


def _generate_pi(payload):
    config = _load("config.json")
    defaults = get_defaults(config, "pi")

    items = []
    for raw in payload.get("items", []):
        it = {
            "line_type": "fee" if raw.get("line_type") == "fee" else "product",
            "item": str(raw.get("item") or "").strip(),
            "desc": raw.get("desc", ""),
        }
        if it["line_type"] == "fee":
            it["amount"] = _num(raw.get("amount"))
        else:
            it["unit"] = raw.get("unit", "PCS")
            it["qty"] = _num(raw.get("qty"))
            it["unit_price"] = _num(raw.get("unit_price"))
        items.append(it)

    total = compute_total(items)
    accounts = config.get("bank_accounts", [])
    acc_id = payload.get("bank_account_id")
    bank = next((a for a in accounts if a.get("id") == acc_id), None)
    if bank is None:
        bank = next((a for a in accounts if a.get("is_default")), accounts[0] if accounts else None)

    customer = str(payload.get("customer") or "").strip()
    err = _pi_validation_error(items, total, bank, customer)
    if err:
        return JSONResponse({"error": err}, status_code=400)

    terms = {k: (payload.get("terms") or {}).get(k, defaults.get(k, "")) for k in TERM_KEYS}
    revision = max(0, _int_or(payload.get("revision"), 0))
    doc_date = payload.get("date", datetime.date.today().strftime("%Y/%m/%d"))

    with DOCSTORE.lock():
        config = _load("config.json")
        claimed_seq = None
        no = str(payload.get("pi_no") or payload.get("no") or "").strip()
        if payload.get("advance_pi") or not no:
            no, claimed_seq = _peek_next_no(config)
        if DOCSTORE.exists(no, "pi", revision):
            raise DuplicateKeyError(no)
        fname = _doc_filename("pi", doc_date, no, revision)
        out_path = os.path.join(OUTPUT, fname)
        doc = {
            "no": no, "doc_type": "pi", "revision": revision,
            "source": _source_triple(payload),
            "date": doc_date,
            "from_name": payload.get("from_name", defaults.get("from_name", "")),
            "currency": payload.get("currency", defaults.get("currency", "USD")),
            "customer": customer, "contact": payload.get("contact", ""),
            "items": items, "terms": terms,
            "bank_account_id": bank.get("id"),
            "bank_account": bank,
            "stamp": bool(payload.get("stamp",
                          config.get("stamp", {}).get("enabled_default", True))),
            "totals": {"amount": total, "amount_words": amount_to_words(total)},
            "file": os.path.join("output", fname).replace("\\", "/"),
            "legacy_file": None, "archived_only": False,
        }
        generate_pi(config, doc, out_path, base_dir=DATA)
        DOCSTORE.save_document(doc)
        if claimed_seq is not None:
            _commit_seq(config, claimed_seq)

    return _file_or_pdf_response(payload, out_path, fname, no)


def _file_or_pdf_response(payload, out_path, fname, no):
    fmt = payload.get("fmt", "xlsx")
    if fmt == "pdf":
        pdf_path = _to_pdf(out_path)
        if pdf_path and os.path.exists(pdf_path):
            out_path, fname = pdf_path, os.path.basename(pdf_path)
        else:
            return JSONResponse({"error": "PDF 转换需要本机安装 LibreOffice;已生成 xlsx,可手动另存为 PDF。",
                                 "fallback_xlsx": fname,
                                 "pi_no": no,
                                 "next_pi": _peek_next_no(_load("config.json"))[0]},
                                status_code=200)
    media = ("application/pdf" if fname.endswith(".pdf")
             else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return FileResponse(out_path, filename=fname, media_type=media)


def _to_pdf(xlsx_path):
    for exe in ("libreoffice", "soffice"):
        if shutil.which(exe):
            try:
                subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir",
                                os.path.dirname(xlsx_path), xlsx_path],
                               check=True, timeout=120,
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                return xlsx_path[:-5] + ".pdf"
            except Exception:
                return None
    return None


# ---------------- history / archive APIs ----------------

@app.get("/api/documents")
def list_documents(type: str = Query(None), q: str = Query(None),
                   from_name: str = Query(None, alias="from"),
                   date_from: str = Query(None), date_to: str = Query(None),
                   page: int = Query(1)):
    res = DOCSTORE.query(doc_type=type or None, from_name=from_name or None,
                         q=q or None, date_from=date_from or None,
                         date_to=date_to or None, page=page)
    config = _load("config.json")
    opts = set(config.get("sales_list", [])) | set(DOCSTORE.from_names())
    res["from_options"] = sorted(o for o in opts if o)
    return res


def _latest_revision(no, doc_type):
    revs = [int(e.get("revision", 0)) for e in DOCSTORE.load_index()["entries"]
            if e["no"] == no and e["doc_type"] == doc_type]
    return max(revs) if revs else None


@app.get("/api/documents/{no}/{doc_type}/{rev}")
def get_document(no: str, doc_type: str, rev: int):
    doc = DOCSTORE.get_document(no, doc_type, rev)
    if doc is None:
        return JSONResponse({"error": "单据不存在"}, status_code=404)
    return doc


@app.get("/api/documents/{no}/{doc_type}/{rev}/download")
def download_document(no: str, doc_type: str, rev: int):
    doc = DOCSTORE.get_document(no, doc_type, rev)
    if doc is None:
        return JSONResponse({"error": "单据不存在"}, status_code=404)
    for key in ("file", "legacy_file"):
        p = doc.get(key)
        if not p:
            continue
        full = p if os.path.isabs(p) else os.path.join(_exe_dir(), p)
        if os.path.exists(full):
            media = ("application/pdf" if full.lower().endswith(".pdf")
                     else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            return FileResponse(full, filename=os.path.basename(full), media_type=media)
    return JSONResponse({"error": "文件已被移动或删除,无法下载(存档数据仍可查看)"}, status_code=404)


@app.post("/api/documents/convert-to-pi")
def convert_to_pi(payload: dict = Body(...)):
    """Return a PI prefill from a quotation. Never claims a number,
    never persists anything (plan §4)."""
    no = str(payload.get("no") or "").strip()
    rev = payload.get("revision")
    if rev is None:
        rev = _latest_revision(no, "quotation")
    if rev is None:
        return JSONResponse({"error": "找不到该报价"}, status_code=404)
    doc = DOCSTORE.get_document(no, "quotation", int(rev))
    if doc is None:
        return JSONResponse({"error": "找不到该报价"}, status_code=404)
    if doc.get("archived_only") or not doc.get("items"):
        return JSONResponse({"error": "旧单没有明细数据,无法一键转 PI,请手工新建。"}, status_code=400)

    config = _load("config.json")
    defaults = get_defaults(config, "pi")
    tiered = doc.get("mode") == "tiered"
    items = []
    for it in doc["items"]:
        items.append({
            "line_type": "product",
            "item": it.get("item", ""), "desc": it.get("desc", ""),
            "unit": it.get("unit", "PCS"),
            # tiered source: no committed qty; price defaults to band-1 price
            "qty": None if tiered else it.get("qty"),
            "unit_price": it.get("price1") if tiered else it.get("unit_price"),
        })
    prefill = {
        "doc_type": "pi",
        "customer": doc.get("customer", ""), "contact": doc.get("contact", ""),
        "from_name": doc.get("from_name", defaults.get("from_name", "")),
        "currency": doc.get("currency", defaults.get("currency", "USD")),
        "items": items,
        "terms": {k: defaults.get(k, "") for k in TERM_KEYS},
        "source": {"no": no, "doc_type": "quotation", "revision": int(rev)},
        "tiered_source": tiered,
        "source_no": no,
        "pi_exists": DOCSTORE.exists(no, "pi", 0),
    }
    return prefill


@app.post("/api/documents/revise")
def revise_document(payload: dict = Body(...)):
    """Return a rev+1 prefill of an existing document. Not persisted."""
    no = str(payload.get("no") or "").strip()
    doc_type = payload.get("doc_type", "quotation")
    rev = payload.get("revision")
    if rev is None:
        rev = _latest_revision(no, doc_type)
    if rev is None:
        return JSONResponse({"error": "找不到该单据"}, status_code=404)
    doc = DOCSTORE.get_document(no, doc_type, int(rev))
    if doc is None:
        return JSONResponse({"error": "找不到该单据"}, status_code=404)
    if doc.get("archived_only") or not doc.get("items"):
        return JSONResponse({"error": "旧单没有明细数据,无法复制为新版,请手工新建。"}, status_code=400)
    latest = _latest_revision(no, doc_type)
    prefill = dict(doc)
    prefill["revision"] = (latest if latest is not None else int(rev)) + 1
    prefill.pop("created_at", None)
    prefill.pop("file", None)
    prefill.pop("legacy_file", None)
    return prefill


@app.get("/api/amount_words")
def get_amount_words(value: float = Query(...)):
    try:
        return {"words": amount_to_words(value)}
    except (ValueError, TypeError):
        return JSONResponse({"error": "金额无效"}, status_code=400)


# ---------------- products / config / images ----------------

@app.post("/api/products")
def save_products(payload: dict = Body(...)):
    products = payload.get("products")
    if not isinstance(products, list):
        return JSONResponse({"error": "格式错误"}, status_code=400)
    # light validation
    seen = set()
    for p in products:
        if not p.get("select") or not p.get("item"):
            return JSONResponse({"error": f"产品缺少 选型名/对外名: {p.get('id','?')}"}, status_code=400)
        if p["select"] in seen:
            return JSONResponse({"error": f"选型名重复: {p['select']}"}, status_code=400)
        seen.add(p["select"])
    _save("products.json", products)
    return {"ok": True, "count": len(products)}


@app.post("/api/config")
def save_config(payload: dict = Body(...)):
    config = _load("config.json")
    if "defaults" in payload and isinstance(payload["defaults"], dict):
        cur = config.setdefault("defaults", {})
        incoming = payload["defaults"]
        if "quotation" in incoming or "pi" in incoming:
            for dt in ("quotation", "pi"):
                if dt in incoming:
                    cur.setdefault(dt, {}).update(incoming[dt])
        else:
            # legacy flat update -> quotation bucket (or flat if still legacy)
            (cur.setdefault("quotation", {}) if "quotation" in cur else cur).update(incoming)
    if "company" in payload:
        config["company"].update(payload["company"])
    if "pi" in payload:
        config.setdefault("pi", {}).update(payload["pi"])
    if "sales_list" in payload:
        names = [str(n).strip() for n in payload["sales_list"] if str(n).strip()]
        if not names:
            return JSONResponse({"error": "报价人列表不能为空"}, status_code=400)
        config["sales_list"] = names
    if "bank_accounts" in payload:
        accounts = payload["bank_accounts"]
        if not isinstance(accounts, list) or not accounts:
            return JSONResponse({"error": "至少保留一个银行账户"}, status_code=400)
        for i, a in enumerate(accounts, 1):
            for k in BANK_REQUIRED:
                if not str(a.get(k) or "").strip():
                    return JSONResponse({"error": f"银行账户第 {i} 条信息不完整(户行/地址/收款人/账号/SWIFT 均必填)"},
                                        status_code=400)
            a.setdefault("id", _safe(a.get("account_no")) or f"bank{i}")
            a.setdefault("currency", "USD")
            a.setdefault("label", a["ac_bank"])
        if not any(a.get("is_default") for a in accounts):
            accounts[0]["is_default"] = True
        config["bank_accounts"] = accounts
    if "stamp" in payload and isinstance(payload["stamp"], dict):
        config.setdefault("stamp", {}).update(payload["stamp"])
    _save("config.json", config)
    return {"ok": True}


@app.post("/api/upload_image")
async def upload_image(file: UploadFile = File(...), name_hint: str = Form("")):
    imgdir = os.path.join(DATA, "images")
    os.makedirs(imgdir, exist_ok=True)
    orig = file.filename or "image.png"
    ext = os.path.splitext(orig)[1].lower() or ".png"
    if ext not in (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif"):
        return JSONResponse({"error": "仅支持 jpg/png/webp/bmp/gif 图片"}, status_code=400)
    stem = _safe(name_hint) or _safe(os.path.splitext(orig)[0]) or "image"
    fname = stem + ext
    dest = os.path.join(imgdir, fname)
    n = 2
    while os.path.exists(dest) and _safe(name_hint) == "":
        fname = f"{stem}_{n}{ext}"; dest = os.path.join(imgdir, fname); n += 1
    data = await file.read()
    with open(dest, "wb") as f:
        f.write(data)
    return {"filename": fname}


@app.get("/api/img/{name}")
def get_img(name: str):
    p = os.path.join(DATA, "images", os.path.basename(name))
    if os.path.exists(p):
        return FileResponse(p)
    return JSONResponse({"error": "not found"}, status_code=404)


if os.path.isdir(STATIC):
    app.mount("/static", StaticFiles(directory=STATIC), name="static")
