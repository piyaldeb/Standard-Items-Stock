import os
import time
from datetime import datetime
from pathlib import Path

import requests
import json
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# =========================
# CONFIG — edit these only
# =========================
# Odoo
ODOO_URL   = "https://taps.odoo.com"
DB         = "masbha-tex-taps-master-2093561"
USERNAME   = "ranak@texzipperbd.com"
PASSWORD   = "2326"

MODEL      = "pending.stock.config"   # target model of the export preset
EXPORT_ID  = 670                      # ir.exports preset id (e.g., 550, 351, etc.)
DOMAIN     = []                       # Odoo domain filter; [] = all records
ALLOWED_COMPANY_IDS = [1]             # active company context (e.g., [3] for Metal)
TZ         = "Asia/Dhaka"

# Local file (optional: saved then deleted; kept with timestamp fallback if locked)
OUTFILE = "pending_stock_00_ranak.xlsx"

# Google Sheets
GOOGLE_SHEET_URL     = "https://docs.google.com/spreadsheets/d/1fnOSIWQa_mbfMHdgPatjYEIhG3kQlzPy0djHG8TOszk/edit?gid=1326846174"
SHEET_NAME           = "Zipper Raw"
SERVICE_ACCOUNT_JSON = "credentials.json"

# Paste behavior
PASTE_COLUMNS = 10  # keep first 10 columns (A:J)

# =========================
# Helpers
# =========================
def log(msg: str):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)

def col_letter(n: int) -> str:
    """1 -> A, 2 -> B, ..."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

# =========================
# HTTP session + Odoo RPC
# =========================
session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})

def call_kw(model, method, args=None, kwargs=None):
    """Call Odoo JSON-RPC endpoint /web/dataset/call_kw/{model}/{method}"""
    url = f"{ODOO_URL}/web/dataset/call_kw/{model}/{method}"
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {"model": model, "method": method, "args": args or [], "kwargs": kwargs or {}},
    }
    r = session.post(url, json=payload)
    r.raise_for_status()
    res = r.json()
    if "error" in res:
        raise RuntimeError(res["error"])
    return res.get("result")

# =========================
# Main
# =========================
def main():
    # 1) Login
    log("Logging into Odoo…")
    login = session.post(f"{ODOO_URL}/web/session/authenticate", json={
        "jsonrpc": "2.0",
        "params": {"db": DB, "login": USERNAME, "password": PASSWORD}
    })
    login.raise_for_status()
    uid = login.json().get("result", {}).get("uid")
    if not uid:
        raise RuntimeError("Login failed")
    CTX = {"lang": "en_US", "tz": TZ, "uid": uid, "allowed_company_ids": ALLOWED_COMPANY_IDS}
    log(f"✅ Logged in (uid={uid})")

    # 2) Load export preset (ir.exports)
    log(f"Loading export preset {EXPORT_ID} …")
    exports = call_kw(
        "ir.exports", "search_read",
        args=[[["id", "=", EXPORT_ID]]],
        kwargs={"fields": ["id", "name", "resource", "export_fields"], "context": CTX},
    )
    exp_rec = exports[0] if exports else None
    if not exp_rec:
        raise RuntimeError(f"Export preset with ID {EXPORT_ID} not found.")
    if exp_rec["resource"] != MODEL:
        raise RuntimeError(f"Preset {EXPORT_ID} is for model '{exp_rec['resource']}', not '{MODEL}'")
    export_line_ids = exp_rec["export_fields"]  # numeric ir.exports.line IDs

    # 3) Resolve ordered field names (server has no 'label' on ir.exports.line)
    log("Resolving preset lines (field names in preset order)…")
    lines = call_kw("ir.exports.line", "read", args=[export_line_ids],
                    kwargs={"fields": ["id", "name"], "context": CTX})
    by_id = {l["id"]: l for l in lines}
    ordered = [by_id[i] for i in export_line_ids if i in by_id]
    field_names = [l["name"] for l in ordered]  # e.g., "inventory_code", "product_type", "product_type/id"
    missing = [i for i in export_line_ids if i not in by_id]
    if missing:
        log(f"⚠️ Missing export line IDs (ignored): {missing}")
    if not field_names:
        raise RuntimeError("No export fields resolved (field_names is empty).")

    # Pretty headers via fields_get on base field (handles '/id', '/display_name', etc.)
    base_fields = sorted(set(n.split("/")[0] for n in field_names))
    fg = call_kw(MODEL, "fields_get", args=[base_fields],
                 kwargs={"attributes": ["string"], "context": CTX})

    def pretty_label(name: str) -> str:
        if "/" in name:
            base, suffix = name.split("/", 1)
            base_label = fg.get(base, {}).get("string", base)
            if suffix in ("display_name", "name"):
                return base_label
            if suffix == "id":
                return f"{base_label} (ID)"
            return f"{base_label}/{suffix}"
        return fg.get(name, {}).get("string", name)

    columns = [pretty_label(n) for n in field_names]

    # 4) Get record ids to export
    log("Searching records…")
    ids = call_kw(MODEL, "search", args=[DOMAIN], kwargs={"context": CTX})
    log(f"Found {len(ids)} records")
    if not ids:
        log("No records match the domain; nothing to export.")
        return

    # 5) Export data
    log("Exporting data via export_data…")
    export_res = call_kw(MODEL, "export_data", args=[ids, field_names], kwargs={"context": CTX})
    rows = export_res.get("datas", [])
    df = pd.DataFrame(rows, columns=columns)
    log(f"DataFrame shape: {df.shape}")

    # 6) Optional local save with Windows-safe fallback
    saved_path = None
    try:
        # Try to overwrite existing file (may fail if open in Excel)
        p = Path(OUTFILE)
        if p.exists():
            try:
                p.unlink()
            except PermissionError:
                pass
        df.to_excel(OUTFILE, index=False)
        saved_path = OUTFILE
        log(f"Saved local copy: {OUTFILE}")
    except PermissionError:
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        alt = Path(OUTFILE).with_name(f"{Path(OUTFILE).stem}_{ts}.xlsx")
        df.to_excel(alt, index=False)
        saved_path = str(alt)
        log(f"⚠️ '{OUTFILE}' is in use. Saved to '{alt}' instead.")

    # 7) Trim to first N columns for Sheet
    if PASTE_COLUMNS and df.shape[1] > PASTE_COLUMNS:
        df = df.iloc[:, :PASTE_COLUMNS]
        log(f"Trimmed to first {PASTE_COLUMNS} columns → shape: {df.shape}")

    # 8) Google Sheets upload
    log("Authorizing Google Sheets…")
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_JSON, scope)
    client = gspread.authorize(creds)

    log("Opening spreadsheet and worksheet…")
    spreadsheet = client.open_by_url(GOOGLE_SHEET_URL)
    worksheet = spreadsheet.worksheet(SHEET_NAME)

    last_col_letter = col_letter(max(df.shape[1], 1))
    log(f"Clearing range A:{last_col_letter} …")
    worksheet.batch_clear([f"A:{last_col_letter}"])

    values = [df.columns.tolist()] + df.values.tolist()
    end_row = len(values)  # header + rows
    log(f"Uploading to A1:{last_col_letter}{end_row} …")
    worksheet.update(f"A1:{last_col_letter}{end_row}", values)
    log("✅ Uploaded to Google Sheet.")

    # 9) Cleanup local file (best-effort)
    if saved_path:
        try:
            Path(saved_path).unlink()
            log(f"🗑️ Deleted local file: {saved_path}")
        except PermissionError:
            log(f"⚠️ Could not delete '{saved_path}' (still open). Close it and delete manually.")

    log("🎉 Done.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"❌ ERROR: {e}")
        raise
