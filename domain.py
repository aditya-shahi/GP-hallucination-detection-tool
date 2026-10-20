#!/usr/bin/env python3
"""
frontend.py - single-file backend + front end for the DOMAIN-WISE sheet.

Two pages:

1) Checker — for every row of the Fund .xlsx (Sheet1), check whether the
   value in each datapoint cell actually appears in the page text of the
   reference URL(s) listed next to it. Several URLs in one reference cell
   are separated by "|".

   Datapoints handled (12 sheets total):
     * Firm AUM                       -> numeric match
     * Firm AUM Date                  -> DD/MM/YYYY date match
     * Firm Geo Focus
       Firm Country Focus
       Firm Industry Focus            -> plain list, split on "," and "|"
     * Asset Class Geo Focus
       Asset Class Country Focus
       Asset Class Industry Focus
       Investment Type                -> asset-class list, "Class - Data1, Data2 | Class - Data3"
     * Asset Class AUM Date           -> asset-class date, "Class - DD/MM/YYYY | ..."
     * Asset Class AUM
       Total Capital Invested         -> asset-class numbers, "Class - 1,250 | ..."

   For every list/asset-list datapoint, rule (B): a row is a hit only if
   EVERY listed value is found somewhere across the reference URLs.

   Columns NOT checked (no sheets): Firm AUM Currency, Asset Class,
   Total Capital Invested Currency.

2) Keyword Match — input one or more domains and their keywords, e.g.

       eatonvance.com: "AUM", "assets"
       morganstanley.com, blackrock.com: "AUM", "assets"

   For each input domain we filter the URL-wise CSV's "Domain" column
   (exact match, case-insensitive), then fetch the s3status of every row
   whose Domain matches, extract url_visible_text, and match ALL keywords
   assigned to that domain. Output: one row per domain in a single sheet
   called "Domain_keyword_matches" (Domain, Keywords, URLs scanned,
   Matched with, URLs matched, Evidence, Notes).

Output destination (chosen on each page):
  - Google Sheet   : written straight into a Google Sheet you own (free)
  - Excel          : downloaded as .xlsx

----------------------------------------------------------------------
GOOGLE SHEETS SETUP (all free, no billing needed)
  1. https://console.cloud.google.com -> create a project
  2. "APIs & Services" -> Library -> enable "Google Sheets API"
  3. "IAM & Admin" -> Service Accounts -> create one -> Keys -> Add key ->
     JSON. Save the file next to this script as service_account.json
     (or point GOOGLE_SA_JSON at it).
  4. Also enable the "Google Drive API" (needed to create sheets).
  5. Leave the Google Sheet URL blank on the page: a NEW sheet named after your
     input file is created and shared with the Google email you type on the page.
     (Or paste an existing sheet's URL and share it with the service account's
     e-mail - the "client_email" inside the JSON - as Editor.)
----------------------------------------------------------------------
"""

import base64
import csv
import io
import json
import os
import queue
import re
import threading
import time
from urllib.parse import urlparse, quote, unquote

import brotli
import requests
from flask import Flask, jsonify, request, Response

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ----------------------------- CONFIG ---------------------------------
GOOGLE_SA_JSON = os.getenv("GOOGLE_SA_JSON", "service_account.json")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")

S3_LOGIN_ID = os.getenv("S3_LOGIN_ID", "")
S3_PASSWORD = os.getenv("S3_PASSWORD", "")
DECOM_BASE = "http://forage-internal.datables.ai/forage-fi2depress/decom"

# ---- Special datapoints ----
AUM_COL     = "Firm AUM"
AUM_URL_COL = "Firm AUM Reference URL"

DATE_COL     = "Firm AUM Date"
DATE_URL_COL = "Firm AUM Date Reference URL"

# ---- Plain list: split on "," and "|", rule (B) ----
PLAIN_LIST_DATAPOINTS = {
    "Firm Geo Focus":      {"value_col": "Firm Geo Focus",      "url_col": "Firm Geo Focus Reference URL"},
    "Firm Country Focus":  {"value_col": "Firm Country Focus",  "url_col": "Firm Country Focus Reference URL"},
    "Firm Industry Focus": {"value_col": "Firm Industry Focus", "url_col": "Firm Industry Focus Reference URL"},
}

# ---- Asset-class list: "Class - Data1, Data2 | Class - Data3" ----
ASSET_LIST_DATAPOINTS = {
    "Asset Class Geo Focus":      {"value_col": "Asset Class Geo Focus",      "url_col": "Asset Class Geo Focus Reference URL"},
    "Asset Class Country Focus":  {"value_col": "Asset Class Country Focus",  "url_col": "Asset Class Country Focus Reference URL"},
    "Asset Class Industry Focus": {"value_col": "Asset Class Industry Focus", "url_col": "Asset Class Industry Focus Reference URL"},
    "Investment Type":            {"value_col": "Investment Type",            "url_col": "Investment Type Reference URL"},
}

# ---- Asset-class date: "Class - DD/MM/YYYY | Class - DD/MM/YYYY, ..." ----
ASSET_DATE_DATAPOINTS = {
    "Asset Class AUM Date": {"value_col": "Asset Class AUM Date",
                             "url_col":   "Asset Class AUM Date Reference URL"},
}

# ---- Asset-class numbers: "Class - 1,250 | Class - 800, 450" ----
ASSET_AUM_DATAPOINTS = {
    "Asset Class AUM":        {"value_col": "Asset Class AUM",
                               "url_col":   "Asset Class AUM Reference URL"},
    "Total Capital Invested": {"value_col": "Total Capital Invested",
                               "url_col":   "Total Capital Invested Reference URL"},
}

# ---- Columns carried through but never checked (no sheets) ----
SKIP_COLS = ("Firm AUM Currency", "Asset Class", "Total Capital Invested Currency")

# ---- Column names in the URL-wise CSV ----
LOOKUP_URL_COL = "URL"
LOOKUP_JSON_COL = "s3status"

# ---- Behaviour ----
TARGET_KEY = "url_visible_text"
CONTEXT_WORDS = 10
MAX_MATCHES_SHOWN = 10
MATCH_SEPARATOR = " | "
MAX_BYTES = 512 * 1024 * 1024
PORT = 8767

# ---- Keyword Match page ----
STATE_FILE = os.getenv("CHECKER_STATE_FILE", "checker_state.json")
COL_MATCHED_WITH = "Matched with"
COL_URLS_MATCHED = "Urls matched"
COL_EVIDENCE = "Evidence"
MATCH_OUT_COLS = [COL_MATCHED_WITH, COL_URLS_MATCHED, COL_EVIDENCE]

# ---- Stable tab names ----
def _tab(dp):
    if dp == "Firm AUM":      return "AUM_results"
    if dp == "Firm AUM Date": return "Date_results"
    return re.sub(r"[^A-Za-z0-9_]", "_", dp) + "_results"

def _kw_tab(dp):
    if dp == "Firm AUM":      return "AUM_keyword_matches"
    if dp == "Firm AUM Date": return "Date_keyword_matches"
    return re.sub(r"[^A-Za-z0-9_]", "_", dp) + "_keyword_matches"

SHEET_ORDER = (
    ["AUM_results", "Date_results"]
    + [_tab(n) for n in PLAIN_LIST_DATAPOINTS]
    + [_tab(n) for n in ASSET_LIST_DATAPOINTS]
    + [_tab(n) for n in ASSET_DATE_DATAPOINTS]
    + [_tab(n) for n in ASSET_AUM_DATAPOINTS]
)

# ---- Keyword-match registry for all 12 datapoints ----
MATCH_TABS = {}
MATCH_TABS["Firm AUM"] = {
    "source_tab": _tab("Firm AUM"),
    "out_tab":    _kw_tab("Firm AUM"),
    "url_col":    AUM_URL_COL,
}
MATCH_TABS["Firm AUM Date"] = {
    "source_tab": _tab("Firm AUM Date"),
    "out_tab":    _kw_tab("Firm AUM Date"),
    "url_col":    DATE_URL_COL,
}
for _fam in (PLAIN_LIST_DATAPOINTS, ASSET_LIST_DATAPOINTS,
             ASSET_DATE_DATAPOINTS, ASSET_AUM_DATAPOINTS):
    for _n, _c in _fam.items():
        MATCH_TABS[_n] = {
            "source_tab": _tab(_n),
            "out_tab":    _kw_tab(_n),
            "url_col":    _c["url_col"],
        }

# ---- Google Sheets limits ----
GSHEET_MAX_CELL_CHARS = 49000
GSHEET_CHUNK_ROWS = 2000
# ----------------------------------------------------------------------

app = Flask(__name__)


# ------------------------- URL helpers --------------------------------

def split_urls(cell):
    if not cell:
        return []
    parts = [p.strip() for p in str(cell).split("|")]
    return [p for p in parts if p]


def norm_url(u):
    u = (u or "").strip()
    if not u:
        return ""
    p = urlparse(u)
    if not p.netloc:
        return u.lower()
    path = p.path.rstrip("/")
    return f"{p.netloc.lower()}{path}" + (f"?{p.query}" if p.query else "")


# ------------------------- Multi-value splitters ----------------------

def split_multi_values(cell):
    """Split on "," and "|"; dedupe case-insensitively."""
    if cell is None:
        return []
    out, seen = [], set()
    for p in re.split(r"[,|]", str(cell)):
        p = p.strip()
        if not p:
            continue
        key = p.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return out


_ASSET_PREFIX_RE = re.compile(r"^(.*?)\s+-\s+(.*)$")


def split_asset_class_values(cell):
    """
    "Equity - India, China | Private Debt - Mexico | Infrastructure - France"
      -> ["India", "China", "Mexico", "France"]   (class labels dropped)

    Split on "|" first (page groups), then drop the "Class - " prefix from
    each group, then split the remainder on ",". Dedupes case-insensitively.
    A group without " - " is kept whole as data.
    """
    if cell is None:
        return []
    out, seen = [], set()
    for group in str(cell).split("|"):
        group = group.strip()
        if not group:
            continue
        m = _ASSET_PREFIX_RE.match(group)
        data_part = m.group(2) if m else group
        for p in data_part.split(","):
            p = p.strip()
            if not p:
                continue
            key = p.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(p)
    return out


# ------------------------- generic helpers ----------------------------

def words_before(text, pos, n):
    parts = text[:pos].split()
    return " ".join(parts[-n:]) if parts else ""


def words_after(text, pos, n):
    parts = text[pos:].split()
    return " ".join(parts[:n]) if parts else ""


def _context(text, start, end):
    pre = words_before(text, start, CONTEXT_WORDS)
    post = words_after(text, end, CONTEXT_WORDS)
    return (pre + " [" + text[start:end] + "] " + post).strip()


# ------------------------- AUM number helpers -------------------------

def extract_number(aum_text):
    m = re.search(r"\d[\d,]*(?:\.\d+)?", str(aum_text or ""))
    return m.group(0) if m else None


def build_number_patterns(num):
    variants = {num, num.replace(",", "")}
    pats = []
    for v in variants:
        body = re.escape(v)
        if "." in v:
            body += r"0*"
        pats.append(re.compile(r"(?<!\d)(?<!\d,)(?<!\d\.)" + body + r"(?!\d)(?!\.\d)"))
    return pats


# ------------------------- Date helpers -------------------------------

FUND_DATE_RE = re.compile(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$")


def parse_fund_date(text):
    m = FUND_DATE_RE.match(str(text or ""))
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if not (1 <= d <= 31 and 1 <= mo <= 12):
        return None
    return (d, mo, y)


# ------------------------- Match finders ------------------------------

def find_number_matches(text, pats):
    seen = {}
    for p in pats:
        for m in p.finditer(text):
            if m.start() in seen:
                continue
            seen[m.start()] = {
                "match": m.group(0),
                "context": _context(text, m.start(), m.end()),
            }
    return [seen[k] for k in sorted(seen)]


def find_date_matches(text, d, mo, y):
    dd, mm = f"{d:02d}", f"{mo:02d}"
    d1, m1 = str(d), str(mo)
    yy4 = f"{y:04d}"
    yy2 = f"{y % 100:02d}"

    variants = [
        (f"DD/MM/YYYY ({dd}/{mm}/{yy4})", dd, mm, yy4),
        (f"DD-MM-YYYY ({dd}-{mm}-{yy4})", dd, mm, yy4),
        (f"MM/DD/YYYY ({mm}/{dd}/{yy4})", mm, dd, yy4),
        (f"DD/MM/YY ({dd}/{mm}/{yy2})", dd, mm, yy2),
        (f"MM/DD/YY ({mm}/{dd}/{yy2})", mm, dd, yy2),
        ("unpadded D/M/YYYY", d1, m1, yy4),
    ]

    seps = r"[\/.\-]"
    seen_keys = set()
    patterns, labels = [], []
    for label, a, b, yy in variants:
        key = (a, b, yy)
        if key in seen_keys:
            continue
        seen_keys.add(key)
        patterns.append(re.compile(
            r"(?<!\d)"
            r"(?:" + re.escape(a) + seps + re.escape(b) + r"|" +
            re.escape(b) + seps + re.escape(a) + r")"
            + seps + re.escape(yy) +
            r"(?!\d)"
        ))
        labels.append(label)

    found = {}
    for pat, label in zip(patterns, labels):
        for m in pat.finditer(text):
            if m.start() in found:
                continue
            found[m.start()] = {
                "match": m.group(0),
                "context": _context(text, m.start(), m.end()),
                "label": label,
            }
    return [found[k] for k in sorted(found)]


def find_text_matches(text, value):
    """Case-insensitive, word-boundary, whitespace-flexible literal search."""
    if not value:
        return []
    needle = re.escape(value.strip().lower()).replace(r"\ ", r"\s+")
    pat = re.compile(r"(?<![A-Za-z0-9])" + needle + r"(?![A-Za-z0-9])")
    return [(m.start(), m.end()) for m in pat.finditer(text.lower())]


# ------------------------- fetch / decompress -------------------------

def to_decom_url(s3link):
    if s3link.startswith("s3://"):
        rest = s3link[len("s3://"):]
        slash = rest.find("/")
        if slash != -1:
            bucket, key = rest[:slash], rest[slash + 1:]
            s3link = f"https://{bucket}.s3.amazonaws.com/{key}"
    return f"{DECOM_BASE}?s3link={quote(s3link, safe='')}"


def extract_s3link(url):
    i = url.find("s3link=")
    if i == -1:
        return None
    link = url[i + len("s3link="):].strip()
    if re.match(r"^https?%3A", link, re.I):
        link = unquote(link)
    return link or None


def _maybe_decompress(data):
    try:
        out = brotli.decompress(data)
        out[:4096].decode("utf-8")
        return out
    except Exception:
        return data


def _looks_like_text_json(b):
    try:
        s = b[:4096].decode("utf-8")
    except UnicodeDecodeError:
        return False
    return s.lstrip()[:1] in ("{", "[")


def _get(url, auth):
    r = requests.get(
        url, auth=auth, timeout=60, stream=True,
        headers={"User-Agent": "domain-checker/1.0", "Accept": "*/*"},
    )
    r.raise_for_status()
    raw = r.raw.read(MAX_BYTES + 1)
    if len(raw) > MAX_BYTES:
        raise ValueError("File is larger than 512 MB")
    return raw


def fetch_json_text(link, user, pw):
    link = (link or "").strip()
    if not link:
        raise ValueError("Empty s3status link")

    auth = (user, pw) if (user or pw) else None

    attempts = []
    if "decom?" in link or link.endswith("/decom"):
        attempts.append(link)
        inner = extract_s3link(link)
        if inner:
            attempts.append(inner)
    else:
        attempts.append(to_decom_url(link))
        attempts.append(link)

    errors = []
    for url in attempts:
        try:
            raw = _get(url, auth)
        except requests.HTTPError as e:
            errors.append(f"{url.split('?')[0]} -> HTTP {e.response.status_code}")
            continue
        except Exception as e:
            errors.append(f"{url.split('?')[0]} -> {e}")
            continue
        if _looks_like_text_json(raw):
            return raw.decode("utf-8", errors="replace")
        return _maybe_decompress(raw).decode("utf-8", errors="replace")

    raise RuntimeError(" | ".join(errors))


def get_target_text(raw_json):
    try:
        data = json.loads(raw_json)
    except json.JSONDecodeError as e:
        return None, f"JSON parse failed: {e}"
    if not isinstance(data, dict):
        return None, f"JSON root is not an object (no '{TARGET_KEY}' key)"
    if TARGET_KEY not in data:
        return None, f"Key '{TARGET_KEY}' not present in JSON"
    v = data[TARGET_KEY]
    if v is None:
        return None, f"Key '{TARGET_KEY}' is null"
    return (v if isinstance(v, str) else str(v)), ""


# ------------------------- CSV helpers --------------------------------

def parse_csv_bytes(data):
    text = data.decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    rows = list(reader)
    fields = list(reader.fieldnames or [])
    return fields, rows


def parse_xlsx_sheet(data, sheet_name="Sheet1"):
    """
    Read one sheet of an .xlsx file into (fields, rows) — same shape as
    parse_csv_bytes(). Falls back to the first sheet if `sheet_name` isn't
    present. Every cell is stringified (or "" if None) so downstream code
    sees the same types it would from a CSV.
    """
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    try:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.worksheets[0] if wb.worksheets else None
        if ws is None:
            raise ValueError("Workbook has no sheets")
        rows_iter = ws.iter_rows(values_only=True)
        try:
            header = next(rows_iter)
        except StopIteration:
            return [], []
        fields = ["" if h is None else str(h).strip() for h in header]
        # drop trailing empty header cells
        while fields and fields[-1] == "":
            fields.pop()

        def cell(v):
            if v is None:
                return ""
            if isinstance(v, bool):
                return str(v)
            if isinstance(v, float) and v.is_integer():
                return str(int(v))
            return str(v)

        rows = []
        for r in rows_iter:
            if r is None:
                continue
            if all(v is None or v == "" for v in r):
                continue
            rows.append({fields[i]: cell(r[i]) if i < len(r) else ""
                         for i in range(len(fields))})
        return fields, rows
    finally:
        wb.close()


# ------------------------- Per-URL orchestration ----------------------

def _fetch_text_for_url(ref, lookup, cache, user, pw):
    link = lookup.get(norm_url(ref))
    if not link:
        return None, "URL not in URL-wise file"
    if link not in cache:
        try:
            cache[link] = fetch_json_text(link, user, pw)
        except Exception as e:
            return None, f"fetch error: {e}"
    text, note = get_target_text(cache[link])
    if text is None:
        return None, note
    return text, ""


def _fetch_text_for_s3(s3link, cache, user, pw):
    """
    Fetch a bare s3status link (not keyed by URL) and return its
    url_visible_text. Reuses the shared cache of raw JSON so two domains
    that point at the same s3status don't refetch.
    """
    if s3link in cache:
        return get_target_text(cache[s3link])
    try:
        raw = fetch_json_text(s3link, user, pw)
    except Exception as e:
        return None, f"fetch error: {e}"
    cache[s3link] = raw
    return get_target_text(raw)


def process_aum_row(row, cfg, lookup, cache, user, pw):
    out = {
        "Searched value": "",
        "Present in content": "false",
        "Match count": 0,
        "Multiple matches": "false",
        "Match context (10 words each side)": "",
        "Per-URL Results": "",
        "Notes": "",
    }
    raw = (row.get(cfg["value_col"]) or "").strip()
    refs = split_urls(row.get(cfg["url_col"]))
    if not raw:
        out["Notes"] = f"Blank {cfg['value_col']} (skipped)"; return out
    if not refs:
        out["Notes"] = f"Blank {cfg['url_col']} (skipped)"; return out

    num = extract_number(raw)
    out["Searched value"] = num or ""
    if not num:
        out["Notes"] = "No number found"; return out
    pats = build_number_patterns(num)

    total, any_match = 0, False
    per_url, notes, contexts = [], [], []
    for ref in refs:
        text, note = _fetch_text_for_url(ref, lookup, cache, user, pw)
        if text is None:
            per_url.append(f"{ref} -> {note}"); notes.append(f"{ref} -> {note}"); continue
        ms = find_number_matches(text, pats)
        total += len(ms)
        if ms:
            any_match = True
            per_url.append(f"{ref} -> true (x{len(ms)})")
            for m in ms[:MAX_MATCHES_SHOWN]:
                contexts.append(f"{ref} :: {m['context']}")
        else:
            per_url.append(f"{ref} -> false (x0)")

    out["Present in content"] = "true" if any_match else "false"
    out["Match count"] = total
    out["Multiple matches"] = "true" if total > 1 else "false"
    out["Match context (10 words each side)"] = MATCH_SEPARATOR.join(contexts)
    out["Notes"] = MATCH_SEPARATOR.join(notes)
    out["Per-URL Results"] = MATCH_SEPARATOR.join(per_url)
    return out


def process_date_row(row, cfg, lookup, cache, user, pw):
    out = {
        "Searched date": "",
        "Parsed as": "",
        "Date present in content": "false",
        "Date match count": 0,
        "Date multiple matches": "false",
        "Matched version": "",
        "Date match context (10 words each side)": "",
        "Per-URL Results": "",
        "Date notes": "",
    }
    raw = (row.get(cfg["value_col"]) or "").strip()
    refs = split_urls(row.get(cfg["url_col"]))
    if not raw:
        out["Date notes"] = f"Blank {cfg['value_col']} (skipped)"; return out
    if not refs:
        out["Date notes"] = f"Blank {cfg['url_col']} (skipped)"; return out

    out["Searched date"] = raw
    parsed = parse_fund_date(raw)
    if not parsed:
        out["Date notes"] = f"{cfg['value_col']} is not DD/MM/YYYY"; return out
    d, mo, y = parsed
    out["Parsed as"] = f"{d:02d}/{mo:02d}/{y:04d}"

    total, any_match = 0, False
    per_url, notes, contexts, versions = [], [], [], []
    for ref in refs:
        text, note = _fetch_text_for_url(ref, lookup, cache, user, pw)
        if text is None:
            per_url.append(f"{ref} -> {note}"); notes.append(f"{ref} -> {note}"); continue
        ms = find_date_matches(text, d, mo, y)
        total += len(ms)
        if ms:
            any_match = True
            per_url.append(f"{ref} -> true (x{len(ms)})")
            for m in ms[:MAX_MATCHES_SHOWN]:
                contexts.append(f"{ref} :: {m['context']}")
            for m in ms:
                base = m.get("label", "").split(" (")[0]
                if base and base not in versions:
                    versions.append(base)
        else:
            per_url.append(f"{ref} -> false (x0)")

    out["Date present in content"] = "true" if any_match else "false"
    out["Date match count"] = total
    out["Date multiple matches"] = "true" if total > 1 else "false"
    out["Matched version"] = MATCH_SEPARATOR.join(versions)
    out["Date match context (10 words each side)"] = MATCH_SEPARATOR.join(contexts)
    out["Date notes"] = MATCH_SEPARATOR.join(notes)
    out["Per-URL Results"] = MATCH_SEPARATOR.join(per_url)
    return out


def process_list_row(row, cfg, lookup, cache, user, pw):
    out = {
        "Values searched": "",
        "Values count": 0,
        "Values found": "",
        "Values missing": "",
        "All values found": "false",
        "Total match count": 0,
        "Per-value evidence (10 words each side)": "",
        "Per-URL Results": "",
        "Notes": "",
    }
    values = split_multi_values(row.get(cfg["value_col"]))
    refs = split_urls(row.get(cfg["url_col"]))
    if not values:
        out["Notes"] = f"Blank {cfg['value_col']} (skipped)"; return out
    if not refs:
        out["Notes"] = f"Blank {cfg['url_col']} (skipped)"; return out
    out["Values searched"] = MATCH_SEPARATOR.join(values)
    out["Values count"] = len(values)

    found_map = {v: [] for v in values}
    total_matches = 0
    per_url_lines, notes_parts = [], []
    for ref in refs:
        text, note = _fetch_text_for_url(ref, lookup, cache, user, pw)
        if text is None:
            per_url_lines.append(f"{ref} -> {note}")
            notes_parts.append(f"{ref} -> {note}")
            continue
        hits_here = []
        for v in values:
            spans = find_text_matches(text, v)
            if not spans:
                continue
            hits_here.append(v)
            total_matches += len(spans)
            for s, e in spans[:MAX_MATCHES_SHOWN]:
                found_map[v].append((ref, _context(text, s, e)))
        per_url_lines.append(
            f"{ref} -> true ({', '.join(hits_here)})" if hits_here
            else f"{ref} -> false (x0)")

    found_values   = [v for v in values if found_map[v]]
    missing_values = [v for v in values if not found_map[v]]
    evidence_lines = []
    for v in found_values:
        for ref, ctx in found_map[v][:MAX_MATCHES_SHOWN]:
            evidence_lines.append(f"[{v}] {ref} :: {ctx}")

    out["Values found"]   = MATCH_SEPARATOR.join(found_values)
    out["Values missing"] = MATCH_SEPARATOR.join(missing_values)
    out["All values found"] = "true" if not missing_values else "false"
    out["Total match count"] = total_matches
    out["Per-value evidence (10 words each side)"] = MATCH_SEPARATOR.join(evidence_lines)
    out["Per-URL Results"] = MATCH_SEPARATOR.join(per_url_lines)
    out["Notes"] = MATCH_SEPARATOR.join(notes_parts)
    return out


def process_asset_list_row(row, cfg, lookup, cache, user, pw):
    out = {
        "Class + data searched": "",
        "Values searched": "",
        "Values count": 0,
        "Values found": "",
        "Values missing": "",
        "All values found": "false",
        "Total match count": 0,
        "Per-value evidence (10 words each side)": "",
        "Per-URL Results": "",
        "Notes": "",
    }
    raw = (row.get(cfg["value_col"]) or "").strip()
    refs = split_urls(row.get(cfg["url_col"]))
    if not raw:
        out["Notes"] = f"Blank {cfg['value_col']} (skipped)"; return out
    if not refs:
        out["Notes"] = f"Blank {cfg['url_col']} (skipped)"; return out

    out["Class + data searched"] = raw
    values = split_asset_class_values(raw)
    out["Values searched"] = MATCH_SEPARATOR.join(values)
    out["Values count"] = len(values)
    if not values:
        out["Notes"] = "No data values after stripping class labels"; return out

    found_map = {v: [] for v in values}
    total_matches = 0
    per_url_lines, notes_parts = [], []
    for ref in refs:
        text, note = _fetch_text_for_url(ref, lookup, cache, user, pw)
        if text is None:
            per_url_lines.append(f"{ref} -> {note}")
            notes_parts.append(f"{ref} -> {note}")
            continue
        hits_here = []
        for v in values:
            spans = find_text_matches(text, v)
            if not spans:
                continue
            hits_here.append(v)
            total_matches += len(spans)
            for s, e in spans[:MAX_MATCHES_SHOWN]:
                found_map[v].append((ref, _context(text, s, e)))
        per_url_lines.append(
            f"{ref} -> true ({', '.join(hits_here)})" if hits_here
            else f"{ref} -> false (x0)")

    found_values   = [v for v in values if found_map[v]]
    missing_values = [v for v in values if not found_map[v]]
    evidence_lines = []
    for v in found_values:
        for ref, ctx in found_map[v][:MAX_MATCHES_SHOWN]:
            evidence_lines.append(f"[{v}] {ref} :: {ctx}")

    out["Values found"]   = MATCH_SEPARATOR.join(found_values)
    out["Values missing"] = MATCH_SEPARATOR.join(missing_values)
    out["All values found"] = "true" if not missing_values else "false"
    out["Total match count"] = total_matches
    out["Per-value evidence (10 words each side)"] = MATCH_SEPARATOR.join(evidence_lines)
    out["Per-URL Results"] = MATCH_SEPARATOR.join(per_url_lines)
    out["Notes"] = MATCH_SEPARATOR.join(notes_parts)
    return out


def process_asset_date_row(row, cfg, lookup, cache, user, pw):
    out = {
        "Class + data searched": "",
        "Dates searched": "",
        "Dates count": 0,
        "Dates parsed": "",
        "Dates unparseable": "",
        "Dates found": "",
        "Dates missing": "",
        "All dates found": "false",
        "Total match count": 0,
        "Matched version": "",
        "Per-date evidence (10 words each side)": "",
        "Per-URL Results": "",
        "Notes": "",
    }
    raw = (row.get(cfg["value_col"]) or "").strip()
    refs = split_urls(row.get(cfg["url_col"]))
    if not raw:
        out["Notes"] = f"Blank {cfg['value_col']} (skipped)"; return out
    if not refs:
        out["Notes"] = f"Blank {cfg['url_col']} (skipped)"; return out

    out["Class + data searched"] = raw
    values = split_asset_class_values(raw)
    out["Dates searched"] = MATCH_SEPARATOR.join(values)
    out["Dates count"] = len(values)
    if not values:
        out["Notes"] = "No date values after stripping class labels"; return out

    parsed, unparsed = [], []
    for v in values:
        p = parse_fund_date(v)
        if p:
            parsed.append((v, p))
        else:
            unparsed.append(v)
    out["Dates parsed"]      = MATCH_SEPARATOR.join(v for v, _ in parsed)
    out["Dates unparseable"] = MATCH_SEPARATOR.join(unparsed)
    if not parsed:
        out["Notes"] = "No DD/MM/YYYY values after stripping class labels"; return out

    found_map = {v: [] for v, _ in parsed}
    versions_set = set()
    total_matches = 0
    per_url_lines, notes_parts = [], []
    for ref in refs:
        text, note = _fetch_text_for_url(ref, lookup, cache, user, pw)
        if text is None:
            per_url_lines.append(f"{ref} -> {note}")
            notes_parts.append(f"{ref} -> {note}")
            continue
        hits_here = []
        for v, (d, mo, y) in parsed:
            ms = find_date_matches(text, d, mo, y)
            if not ms:
                continue
            hits_here.append(v)
            total_matches += len(ms)
            for m in ms[:MAX_MATCHES_SHOWN]:
                found_map[v].append((ref, m["context"]))
                base = m.get("label", "").split(" (")[0]
                if base:
                    versions_set.add(base)
        per_url_lines.append(
            f"{ref} -> true ({', '.join(hits_here)})" if hits_here
            else f"{ref} -> false (x0)")

    found_dates   = [v for v, _ in parsed if found_map[v]]
    missing_dates = [v for v, _ in parsed if not found_map[v]]
    evidence_lines = []
    for v in found_dates:
        for ref, ctx in found_map[v][:MAX_MATCHES_SHOWN]:
            evidence_lines.append(f"[{v}] {ref} :: {ctx}")

    out["Dates found"]   = MATCH_SEPARATOR.join(found_dates)
    out["Dates missing"] = MATCH_SEPARATOR.join(missing_dates)
    out["All dates found"] = "true" if not missing_dates else "false"
    out["Total match count"] = total_matches
    out["Matched version"] = MATCH_SEPARATOR.join(sorted(versions_set))
    out["Per-date evidence (10 words each side)"] = MATCH_SEPARATOR.join(evidence_lines)
    out["Per-URL Results"] = MATCH_SEPARATOR.join(per_url_lines)
    out["Notes"] = MATCH_SEPARATOR.join(notes_parts)
    return out


def process_asset_aum_row(row, cfg, lookup, cache, user, pw):
    out = {
        "Class + data searched": "",
        "Numbers searched": "",
        "Numbers count": 0,
        "Numbers found": "",
        "Numbers missing": "",
        "All numbers found": "false",
        "Total match count": 0,
        "Per-number evidence (10 words each side)": "",
        "Per-URL Results": "",
        "Notes": "",
    }
    raw = (row.get(cfg["value_col"]) or "").strip()
    refs = split_urls(row.get(cfg["url_col"]))
    if not raw:
        out["Notes"] = f"Blank {cfg['value_col']} (skipped)"; return out
    if not refs:
        out["Notes"] = f"Blank {cfg['url_col']} (skipped)"; return out

    out["Class + data searched"] = raw
    values = split_asset_class_values(raw)
    nums, seen = [], set()
    for v in values:
        n = extract_number(v)
        if n and n not in seen:
            seen.add(n); nums.append(n)
    out["Numbers searched"] = MATCH_SEPARATOR.join(nums)
    out["Numbers count"] = len(nums)
    if not nums:
        out["Notes"] = "No numbers after stripping class labels"; return out

    patterns_by_num = {n: build_number_patterns(n) for n in nums}
    found_map = {n: [] for n in nums}
    total_matches = 0
    per_url_lines, notes_parts = [], []
    for ref in refs:
        text, note = _fetch_text_for_url(ref, lookup, cache, user, pw)
        if text is None:
            per_url_lines.append(f"{ref} -> {note}")
            notes_parts.append(f"{ref} -> {note}")
            continue
        hits_here = []
        for n in nums:
            ms = find_number_matches(text, patterns_by_num[n])
            if not ms:
                continue
            hits_here.append(n)
            total_matches += len(ms)
            for m in ms[:MAX_MATCHES_SHOWN]:
                found_map[n].append((ref, m["context"]))
        per_url_lines.append(
            f"{ref} -> true ({', '.join(hits_here)})" if hits_here
            else f"{ref} -> false (x0)")

    found_numbers   = [n for n in nums if found_map[n]]
    missing_numbers = [n for n in nums if not found_map[n]]
    evidence_lines = []
    for n in found_numbers:
        for ref, ctx in found_map[n][:MAX_MATCHES_SHOWN]:
            evidence_lines.append(f"[{n}] {ref} :: {ctx}")

    out["Numbers found"]   = MATCH_SEPARATOR.join(found_numbers)
    out["Numbers missing"] = MATCH_SEPARATOR.join(missing_numbers)
    out["All numbers found"] = "true" if not missing_numbers else "false"
    out["Total match count"] = total_matches
    out["Per-number evidence (10 words each side)"] = MATCH_SEPARATOR.join(evidence_lines)
    out["Per-URL Results"] = MATCH_SEPARATOR.join(per_url_lines)
    out["Notes"] = MATCH_SEPARATOR.join(notes_parts)
    return out


# ------------------------- Lookup / state -----------------------------

KW_JOIN = "|"
_STATE_LOCK = threading.Lock()


def build_lookup(lookup_bytes):
    fields, rows = parse_csv_bytes(lookup_bytes)
    missing = [c for c in (LOOKUP_URL_COL, LOOKUP_JSON_COL) if c not in fields]
    if missing:
        raise ValueError(f"URL-wise CSV is missing column(s): {missing}. Found: {fields}")
    return {
        norm_url(r.get(LOOKUP_URL_COL, "")): (r.get(LOOKUP_JSON_COL) or "").strip()
        for r in rows
        if r.get(LOOKUP_URL_COL)
    }


def _now():
    return time.strftime("%d %b %Y, %H:%M")


def load_state():
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_state(**kw):
    with _STATE_LOCK:
        st = load_state()
        st.update(kw)
        tmp = STATE_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(st, f)
        os.replace(tmp, STATE_FILE)


# ------------------------- Keyword helpers ----------------------------

def parse_keywords(text):
    text = (text or "").strip()
    if not text:
        return []
    items = re.findall(r'["“”]([^"“”]+)["“”]', text)
    if not items:
        items = text.split(",")
    out, seen = [], set()
    for k in items:
        k = " ".join(k.split())
        if k and k.lower() not in seen:
            seen.add(k.lower())
            out.append(k)
    return out


def compile_keywords(kws):
    pats = []
    for kw in kws:
        needle = r"\s+".join(re.escape(w) for w in kw.split())
        pats.append((kw, re.compile(r"(?<![A-Za-z0-9])" + needle + r"(?![A-Za-z0-9])",
                                    re.IGNORECASE)))
    return pats


# ------------------------- Domain keyword match -----------------------

def _norm_domain(s):
    """
    Normalise a domain string for EXACT match against the URL-wise CSV's
    "Domain" column. Only strips surrounding whitespace and lowercases.
    Scheme / www. / trailing slash are NOT removed.
    """
    return (s or "").strip().lower()


def parse_domain_keyword_block(text):
    """
    Parse lines like:

        eatonvance.com: "AUM", "assets"
        morganstanley.com: "fee-based", "advisory"
        morganstanley.com, blackrock.com: "AUM", "assets"

    Returns {normalised_domain: [kw, ...]}. Duplicate domains across lines
    are UNIONed (case-insensitive, order-preserving). Blank lines skipped.

    Errors:
      - a non-blank line without ":"
      - a line whose value side yields no keywords
      - a line whose key side yields no domains
      - no usable lines at all
    """
    out = {}
    for n, raw in enumerate((text or "").splitlines(), start=1):
        line = raw.strip()
        if not line:
            continue
        if ":" not in line:
            raise ValueError(f'Line {n}: expected   domain[, domain]: "kw", "kw"')
        key_part, val_part = line.split(":", 1)
        domains = [_norm_domain(d) for d in key_part.split(",")]
        domains = [d for d in domains if d]
        if not domains:
            raise ValueError(f"Line {n}: no domains before the colon")
        kws = parse_keywords(val_part)
        if not kws:
            raise ValueError(f'Line {n}: no keywords in double quotes after the colon')
        for d in domains:
            existing = out.setdefault(d, [])
            seen = {k.lower() for k in existing}
            for k in kws:
                if k.lower() not in seen:
                    existing.append(k)
                    seen.add(k.lower())
    if not out:
        raise ValueError('Enter at least one line like   domain: "keyword", "keyword"')
    return out


def _rows_by_domain(lookup_rows):
    """
    Build {normalised_domain: [(url, s3status), ...]} from the URL-wise CSV
    rows. Requires columns 'Domain', 'URL', 's3status'.
    """
    fields = set()
    for r in lookup_rows:
        fields.update(r.keys())
    for need in ("Domain", "URL", "s3status"):
        if need not in fields:
            raise ValueError(f"URL-wise CSV is missing required column '{need}'. "
                             f"Found: {sorted(fields)}")
    out = {}
    for r in lookup_rows:
        dom = _norm_domain(r.get("Domain"))
        if not dom:
            continue
        out.setdefault(dom, []).append(
            ((r.get("URL") or "").strip(), (r.get("s3status") or "").strip())
        )
    return out


def run_domain_keyword_match(domain_groups, lookup_rows, user, pw,
                             progress=None, status=None):
    """
    domain_groups: {normalised_domain: [kw, ...]}
    lookup_rows:   raw URL-wise CSV rows (list of dicts)

    For every input domain:
      - find URL-wise rows whose 'Domain' == that domain (exact, lowercased)
      - for each such row, fetch its s3status, extract url_visible_text
      - match ALL of that domain's keywords against the page text
    Emits ONE row per domain. Returns (fields, rows).
    """
    def say(m):
        if status:
            try: status(m)
            except Exception: pass
    def emit(**kw):
        if progress:
            try: progress(kw)
            except Exception: pass

    by_domain = _rows_by_domain(lookup_rows)

    fields = [
        "Domain",
        "Keywords",
        "URLs scanned",
        "Matched with",
        "URLs matched",
        "Evidence",
        "Notes",
    ]
    rows = []
    cache = {}   # s3status -> raw json text (shared across domains)
    done = 0
    total = len(domain_groups)

    for dom, kws in domain_groups.items():
        say(f"Matching '{dom}' ({len(kws)} keyword(s)) ...")
        pats = compile_keywords(kws)

        entries = by_domain.get(dom, [])
        if not entries:
            rows.append({
                "Domain": dom,
                "Keywords": KW_JOIN.join(kws),
                "URLs scanned": 0,
                "Matched with": "",
                "URLs matched": "",
                "Evidence": "",
                "Notes": "domain not in URL-wise file",
            })
            done += 1
            emit(i=done, total=total, domain=dom,
                 matched=False, urls_scanned=0, url_hits=0,
                 matched_with="", note="domain not in URL-wise file")
            continue

        hit_kws, hit_urls, evidence_lines = set(), [], []
        notes = []
        urls_scanned = 0

        for url, s3 in entries:
            if not s3:
                notes.append(f"{url or '(no URL)'} -> blank s3status")
                continue
            urls_scanned += 1
            text, note = _fetch_text_for_s3(s3, cache, user, pw)
            if text is None:
                notes.append(f"{url or s3} -> {note}")
                continue
            found_here = []
            for kw, p in pats:
                for m in p.finditer(text):
                    ctx = _context(text, m.start(), m.end())
                    evidence_lines.append(f"{url or s3} :: {ctx}")
                    found_here.append(kw)
            if found_here:
                hit_urls.append(url or s3)
                hit_kws.update(found_here)

        matched_with = KW_JOIN.join(kw for kw, _ in pats if kw in hit_kws)
        rows.append({
            "Domain": dom,
            "Keywords": KW_JOIN.join(kws),
            "URLs scanned": urls_scanned,
            "Matched with": matched_with,
            "URLs matched": KW_JOIN.join(hit_urls),
            "Evidence": KW_JOIN.join(evidence_lines),
            "Notes": "; ".join(notes)[:2000],
        })

        done += 1
        emit(i=done, total=total, domain=dom,
             matched=bool(hit_urls), urls_scanned=urls_scanned,
             url_hits=len(hit_urls), matched_with=matched_with,
             note="; ".join(notes)[:300])

    return fields, rows


# ------------------------- Core run -----------------------------------

def _run_check_inner(fund_fields, fund_rows, lookup_bytes, user, pw, progress=None):
    def emit(**kw):
        if progress:
            try: progress(kw)
            except Exception: pass

    lookup = build_lookup(lookup_bytes)
    cache = {}
    total = len(fund_rows)

    aum_rows, date_rows = [], []
    plain_rows      = {n: [] for n in PLAIN_LIST_DATAPOINTS}
    asset_list_rows = {n: [] for n in ASSET_LIST_DATAPOINTS}
    asset_date_rows = {n: [] for n in ASSET_DATE_DATAPOINTS}
    asset_aum_rows  = {n: [] for n in ASSET_AUM_DATAPOINTS}

    for i, row in enumerate(fund_rows, start=1):
        # Firm AUM
        aum_res = process_aum_row(
            row, {"value_col": AUM_COL, "url_col": AUM_URL_COL},
            lookup, cache, user, pw)
        a = dict(row); a.update(aum_res); aum_rows.append(a)

        # Firm AUM Date
        date_res = process_date_row(
            row, {"value_col": DATE_COL, "url_col": DATE_URL_COL},
            lookup, cache, user, pw)
        d = dict(row); d.update(date_res); date_rows.append(d)

        # Plain list
        plain_summary = {}
        for n, cfg in PLAIN_LIST_DATAPOINTS.items():
            res = process_list_row(row, cfg, lookup, cache, user, pw)
            r = dict(row); r.update(res); plain_rows[n].append(r)
            plain_summary[n] = res["All values found"]

        # Asset-class list
        asset_list_summary = {}
        for n, cfg in ASSET_LIST_DATAPOINTS.items():
            res = process_asset_list_row(row, cfg, lookup, cache, user, pw)
            r = dict(row); r.update(res); asset_list_rows[n].append(r)
            asset_list_summary[n] = res["All values found"]

        # Asset-class date
        asset_date_summary = {}
        for n, cfg in ASSET_DATE_DATAPOINTS.items():
            res = process_asset_date_row(row, cfg, lookup, cache, user, pw)
            r = dict(row); r.update(res); asset_date_rows[n].append(r)
            asset_date_summary[n] = res["All dates found"]

        # Asset-class numbers
        asset_aum_summary = {}
        for n, cfg in ASSET_AUM_DATAPOINTS.items():
            res = process_asset_aum_row(row, cfg, lookup, cache, user, pw)
            r = dict(row); r.update(res); asset_aum_rows[n].append(r)
            asset_aum_summary[n] = res["All numbers found"]

        emit(i=i, total=total,
             aum_flag=aum_res["Present in content"],
             date_flag=date_res["Date present in content"],
             plain=plain_summary,
             asset_list=asset_list_summary,
             asset_date=asset_date_summary,
             asset_aum=asset_aum_summary)

        print(f"[{i}/{total}] AUM {aum_res['Present in content']} "
              f"| Date {date_res['Date present in content']} "
              f"| plain {plain_summary} | asset_list {asset_list_summary} "
              f"| asset_date {asset_date_summary} | asset_aum {asset_aum_summary}")

    aum_fields = list(fund_fields) + [
        "Searched value", "Present in content", "Match count",
        "Multiple matches", "Match context (10 words each side)",
        "Per-URL Results", "Notes",
    ]
    date_fields = list(fund_fields) + [
        "Searched date", "Parsed as", "Date present in content",
        "Date match count", "Date multiple matches", "Matched version",
        "Date match context (10 words each side)", "Per-URL Results", "Date notes",
    ]
    plain_fields = list(fund_fields) + [
        "Values searched", "Values count", "Values found", "Values missing",
        "All values found", "Total match count",
        "Per-value evidence (10 words each side)", "Per-URL Results", "Notes",
    ]
    asset_list_fields = list(fund_fields) + [
        "Class + data searched",
        "Values searched", "Values count", "Values found", "Values missing",
        "All values found", "Total match count",
        "Per-value evidence (10 words each side)", "Per-URL Results", "Notes",
    ]
    asset_date_fields = list(fund_fields) + [
        "Class + data searched",
        "Dates searched", "Dates count", "Dates parsed", "Dates unparseable",
        "Dates found", "Dates missing", "All dates found", "Total match count",
        "Matched version", "Per-date evidence (10 words each side)",
        "Per-URL Results", "Notes",
    ]
    asset_aum_fields = list(fund_fields) + [
        "Class + data searched",
        "Numbers searched", "Numbers count", "Numbers found", "Numbers missing",
        "All numbers found", "Total match count",
        "Per-number evidence (10 words each side)", "Per-URL Results", "Notes",
    ]

    sheets = {
        "AUM_results":  (aum_fields,  aum_rows),
        "Date_results": (date_fields, date_rows),
    }
    for n in PLAIN_LIST_DATAPOINTS:  sheets[_tab(n)] = (plain_fields,      plain_rows[n])
    for n in ASSET_LIST_DATAPOINTS:  sheets[_tab(n)] = (asset_list_fields, asset_list_rows[n])
    for n in ASSET_DATE_DATAPOINTS:  sheets[_tab(n)] = (asset_date_fields, asset_date_rows[n])
    for n in ASSET_AUM_DATAPOINTS:   sheets[_tab(n)] = (asset_aum_fields,  asset_aum_rows[n])

    return {name: sheets[name] for name in SHEET_ORDER if name in sheets}


def run_check_from_rows(fund_fields, fund_rows, lookup_bytes, user, pw, progress=None):
    """
    Same as run_check() but takes the Fund sheet as already-parsed
    (fields, rows) instead of raw CSV bytes. Keeps a single place where
    the CSV↔XLSX distinction lives.
    """
    return _run_check_inner(fund_fields, fund_rows, lookup_bytes, user, pw, progress)


# ------------------------- Excel writer -------------------------------

_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _col_width(name):
    low = name.lower()
    if "context" in low or "per-url results" in low or "evidence" in low:
        return 60
    if low.startswith(("firm", "asset", "investment", "total", "url")):
        return 34
    return 18


def read_excel_sheets(data):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    out = {}
    for ws in wb.worksheets:
        rows_iter = ws.iter_rows(values_only=True)
        try:
            header = next(rows_iter)
        except StopIteration:
            continue
        fields = [str(h) if h is not None else "" for h in header]
        rows = []
        for r in rows_iter:
            if r is None:
                continue
            rows.append({fields[i]: ("" if v is None else v)
                         for i, v in enumerate(r) if i < len(fields)})
        out[ws.title] = (fields, rows)
    wb.close()
    return out


def build_excel(sheets):
    wb = Workbook()
    first = True
    for title, (fields, rows) in sheets.items():
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = title[:31] if len(title) > 31 else title
        ws.append(fields)
        for r in rows:
            ws.append([_ILLEGAL_XML.sub("", v) if isinstance(v, str) else v
                       for v in (r.get(f, "") for f in fields)])
        for c in range(1, len(fields) + 1):
            ws.cell(row=1, column=c).font = Font(bold=True)
        for idx, name in enumerate(fields, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = _col_width(name)
        ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ------------------------- Google Sheets writer -----------------------

def extract_sheet_id(s):
    s = (s or "").strip()
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", s)
    if m:
        return m.group(1)
    return s if re.fullmatch(r"[A-Za-z0-9_-]{20,}", s) else ""


def service_account_email():
    try:
        with open(GOOGLE_SA_JSON, encoding="utf-8") as f:
            return json.load(f).get("client_email", "")
    except Exception:
        return ""


def _gs_cell(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return v
    return str(v)[:GSHEET_MAX_CELL_CHARS]


def _gs_retry_factory(say):
    from gspread.exceptions import APIError
    def retry(fn, *a, **kw):
        for attempt in range(5):
            try:
                return fn(*a, **kw)
            except APIError as e:
                code = getattr(getattr(e, "response", None), "status_code", None)
                if code in (429, 500, 503) and attempt < 4:
                    wait = 15 * (attempt + 1)
                    say(f"Google rate limit hit, waiting {wait}s ...")
                    time.sleep(wait); continue
                raise
    return retry


def open_spreadsheet(sheet_id, status=None):
    import gspread
    from gspread.exceptions import APIError
    def say(msg):
        if status:
            try: status(msg)
            except Exception: pass
    if not os.path.exists(GOOGLE_SA_JSON):
        raise RuntimeError(f"Service-account key not found at '{GOOGLE_SA_JSON}'.")
    email = service_account_email()
    retry = _gs_retry_factory(say)
    try:
        gc = gspread.service_account(filename=GOOGLE_SA_JSON)
        sh = retry(gc.open_by_key, sheet_id)
    except APIError as e:
        code = getattr(getattr(e, "response", None), "status_code", None)
        if code in (403, 404):
            raise RuntimeError(
                f"Cannot open the Google Sheet (HTTP {code}). Share it with "
                f"{email or 'the service account'} as Editor.")
        raise
    except gspread.exceptions.SpreadsheetNotFound:
        raise RuntimeError(
            f"Google Sheet not found. Share it with {email or 'the service account'}.")
    return sh, retry


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def sheet_title_from_filename(filename):
    base = os.path.basename((filename or "").replace("\\", "/"))
    base = re.sub(r"\.(csv|xlsx|xls)$", "", base, flags=re.I)
    base = re.sub(r"[\x00-\x1f]", "", base).strip()[:100].strip()
    return base or "Domain QA results"


def create_spreadsheet(title, share_email, status=None):
    import gspread
    from gspread.exceptions import APIError
    def say(msg):
        if status:
            try: status(msg)
            except Exception: pass
    if not os.path.exists(GOOGLE_SA_JSON):
        raise RuntimeError(f"Service-account key not found at '{GOOGLE_SA_JSON}'.")
    retry = _gs_retry_factory(say)
    def code_of(e):
        return getattr(getattr(e, "response", None), "status_code", None)
    try:
        gc = gspread.service_account(filename=GOOGLE_SA_JSON)
        say(f"Creating Google Sheet '{title}' ...")
        sh = retry(gc.create, title)
    except APIError as e:
        raise RuntimeError(f"Could not create the Google Sheet (HTTP {code_of(e)}). Details: {e}")
    try:
        say(f"Sharing it with {share_email} ...")
        retry(sh.share, share_email, perm_type="user", role="writer")
    except APIError as e:
        raise RuntimeError(
            f"The sheet was created ({sh.url}) but sharing it with {share_email} failed "
            f"(HTTP {code_of(e)}): {e}")
    return sh.id


def write_gsheet(sheet_id, sheets, status=None):
    def say(msg):
        if status:
            try: status(msg)
            except Exception: pass
    sh, retry = open_spreadsheet(sheet_id, status)
    existing = {ws.title: ws for ws in retry(sh.worksheets)}
    for title, (fields, rows) in sheets.items():
        say(f"Writing '{title}' ({len(rows)} rows) to Google Sheets ...")
        data = [list(fields)] + [[_gs_cell(r.get(f, "")) for f in fields] for r in rows]
        ws = existing.get(title)
        if ws is None:
            ws = retry(sh.add_worksheet, title=title,
                       rows=max(len(data), 2), cols=max(len(fields), 1))
        else:
            retry(ws.clear)
            retry(ws.resize, rows=max(len(data), 2), cols=max(len(fields), 1))
        for start in range(0, len(data), GSHEET_CHUNK_ROWS):
            chunk = data[start:start + GSHEET_CHUNK_ROWS]
            retry(ws.update, values=chunk, range_name=f"A{start + 1}", raw=True)
        retry(ws.format, "1:1", {"textFormat": {"bold": True}})
        retry(ws.freeze, rows=1)
    default = existing.get("Sheet1")
    if default is not None and "Sheet1" not in sheets:
        try:
            if not default.get_all_values():
                sh.del_worksheet(default)
        except Exception:
            pass
    return sh.url, sh.title


# ------------------------- Routes -------------------------------------

STREAMS = {}
STREAM_LOCK = threading.Lock()
NEXT_STREAM_ID = [0]


@app.route("/")
def index():
    return Response(PAGE_HTML, mimetype="text/html; charset=utf-8")


# ---------- Checker ----------

@app.route("/api/run-stream", methods=["POST"])
def api_run_stream():
    fund = request.files.get("fund_xlsx")
    lookup = request.files.get("lookup_csv")
    if not fund or not lookup:
        return jsonify({"error": "Upload the Fund .xlsx and the URL-wise CSV."}), 400

    user = request.form.get("user") or S3_LOGIN_ID
    pw = request.form.get("pass") or S3_PASSWORD
    mode = request.form.get("mode", "gsheet")
    if mode not in ("gsheet", "excel"):
        return jsonify({"error": "Unknown output mode."}), 400

    sheet_id = ""
    new_title = ""
    share_email = ""
    if mode == "gsheet":
        sheet_id = (extract_sheet_id(request.form.get("sheet"))
                    or extract_sheet_id(GOOGLE_SHEET_ID))
        if not sheet_id:
            share_email = (request.form.get("share_email") or "").strip()
            if not EMAIL_RE.match(share_email):
                return jsonify({"error": "Enter your Google email (the new Google Sheet is "
                                         "shared with it), or paste an existing Google Sheet URL."}), 400
            new_title = sheet_title_from_filename(fund.filename)
        if not os.path.exists(GOOGLE_SA_JSON):
            return jsonify({"error": f"Service-account key not found at '{GOOGLE_SA_JSON}'."}), 400

    # ---- read + parse inputs BEFORE def worker() ----
    fund_bytes = fund.read()
    lookup_bytes = lookup.read()
    try:
        fund_fields, fund_rows = parse_xlsx_sheet(fund_bytes, sheet_name="Sheet1")
    except Exception as e:
        return jsonify({"error": f"Could not read the Fund .xlsx: {e}"}), 400
    if not fund_fields:
        return jsonify({"error": "Sheet1 in the Fund .xlsx is empty."}), 400

    with STREAM_LOCK:
        NEXT_STREAM_ID[0] += 1
        sid = str(NEXT_STREAM_ID[0])
        STREAMS[sid] = queue.Queue()

    # Every free variable the worker reads is bound as a default arg, so the
    # thread sees a concrete value even though api_run_stream has returned.
    def worker(
        sid=sid,
        fund_fields=fund_fields,
        fund_rows=fund_rows,
        lookup_bytes=lookup_bytes,
        user=user,
        pw=pw,
        mode=mode,
        sheet_id=sheet_id,
        new_title=new_title,
        share_email=share_email,
    ):
        q = STREAMS[sid]
        try:
            sheets = run_check_from_rows(fund_fields, fund_rows, lookup_bytes,
                                         user, pw,
                                         progress=lambda ev: q.put(("progress", ev)))
            if mode == "gsheet":
                target_id = sheet_id
                if not target_id:
                    target_id = create_spreadsheet(
                        new_title, share_email,
                        status=lambda m: q.put(("status", {"message": m})))
                url, title = write_gsheet(
                    target_id, sheets,
                    status=lambda m: q.put(("status", {"message": m})))
                try:
                    save_state(sheet_id=target_id, sheet_url=url, sheet_title=title,
                               saved_at=_now(), lookup=build_lookup(lookup_bytes),
                               lookup_saved_at=_now())
                except Exception:
                    pass
                q.put(("done", {"mode": "gsheet", "url": url, "title": title}))
            else:
                q.put(("status", {"message": "Building Excel file ..."}))
                out = build_excel(sheets)
                q.put(("done", {"mode": "excel",
                                "filename": "domain_results.xlsx",
                                "b64": base64.b64encode(out).decode("ascii")}))
        except Exception as e:
            q.put(("fail", {"message": str(e)}))
        finally:
            q.put(("__end__", {}))

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"stream_id": sid})


# ---------- Domain keyword match ----------

@app.route("/api/domain-match-stream", methods=["POST"])
def api_domain_match_stream():
    mode = request.form.get("mode", "gsheet")
    if mode not in ("gsheet", "excel"):
        return jsonify({"error": "Unknown output mode."}), 400

    try:
        domain_groups = parse_domain_keyword_block(
            request.form.get("domains_keywords", ""))
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    user = request.form.get("user") or S3_LOGIN_ID
    pw = request.form.get("pass") or S3_PASSWORD

    # ---- URL-wise CSV: upload or remembered ----
    upload = request.files.get("lookup_csv")
    st = load_state()
    if upload and upload.filename:
        try:
            _, lookup_rows = parse_csv_bytes(upload.read())
        except Exception as e:
            return jsonify({"error": f"Could not read URL-wise CSV: {e}"}), 400
        try:
            _rows_by_domain(lookup_rows)  # validate columns
        except ValueError as e:
            return jsonify({"error": str(e)}), 400
        save_state(lookup_rows=lookup_rows, lookup_saved_at=_now())
    else:
        lookup_rows = st.get("lookup_rows") or []
    if not lookup_rows:
        return jsonify({"error": "No URL-wise data saved yet. Upload the URL-wise CSV "
                                 "(or run the Checker once)."}), 400

    # ---- Google Sheet id / share target ----
    sheet_id = ""
    new_title = ""
    share_email = ""
    if mode == "gsheet":
        sheet_id = (extract_sheet_id(request.form.get("sheet"))
                    or st.get("sheet_id", "")
                    or extract_sheet_id(GOOGLE_SHEET_ID))
        if not sheet_id:
            share_email = (request.form.get("share_email") or "").strip()
            if not EMAIL_RE.match(share_email):
                return jsonify({"error": "Enter your Google email (the new Google Sheet is "
                                         "shared with it), or paste an existing Google Sheet URL."}), 400
            new_title = "Domain keyword matches"
        if not os.path.exists(GOOGLE_SA_JSON):
            return jsonify({"error": f"Service-account key not found at '{GOOGLE_SA_JSON}'."}), 400

    with STREAM_LOCK:
        NEXT_STREAM_ID[0] += 1
        sid = str(NEXT_STREAM_ID[0])
        STREAMS[sid] = queue.Queue()

    def worker(
        sid=sid,
        domain_groups=domain_groups,
        lookup_rows=lookup_rows,
        user=user,
        pw=pw,
        mode=mode,
        sheet_id=sheet_id,
        new_title=new_title,
        share_email=share_email,
    ):
        q = STREAMS[sid]
        try:
            fields, rows = run_domain_keyword_match(
                domain_groups, lookup_rows, user, pw,
                progress=lambda ev: q.put(("progress", ev)),
                status=lambda m: q.put(("status", {"message": m})))
            sheets = {"Domain_keyword_matches": (fields, rows)}

            if mode == "gsheet":
                target_id = sheet_id
                if not target_id:
                    target_id = create_spreadsheet(
                        new_title, share_email,
                        status=lambda m: q.put(("status", {"message": m})))
                url, title = write_gsheet(
                    target_id, sheets,
                    status=lambda m: q.put(("status", {"message": m})))
                try:
                    save_state(sheet_id=target_id, sheet_url=url,
                               sheet_title=title, saved_at=_now(),
                               lookup_rows=lookup_rows,
                               lookup_saved_at=_now())
                except Exception:
                    pass
                q.put(("done", {"mode": "gsheet", "url": url, "title": title,
                                "tab": "Domain_keyword_matches",
                                "rows": len(rows)}))
            else:
                q.put(("status", {"message": "Building Excel file ..."}))
                out = build_excel(sheets)
                q.put(("done", {"mode": "excel",
                                "filename": "domain_keyword_matches.xlsx",
                                "b64": base64.b64encode(out).decode("ascii"),
                                "tab": "Domain_keyword_matches",
                                "rows": len(rows)}))
        except Exception as e:
            q.put(("fail", {"message": str(e)}))
        finally:
            q.put(("__end__", {}))

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"stream_id": sid})


# ---------- SSE progress ----------

@app.route("/api/progress/<sid>")
def api_progress(sid):
    if sid not in STREAMS:
        return jsonify({"error": "unknown stream"}), 404
    q = STREAMS[sid]

    def gen():
        try:
            while True:
                try:
                    kind, payload = q.get(timeout=60)
                except queue.Empty:
                    yield "event: ping\ndata: {}\n\n"; continue
                if kind == "__end__":
                    yield "event: end\ndata: {}\n\n"; break
                yield f"event: {kind}\ndata: {json.dumps(payload)}\n\n"
        finally:
            with STREAM_LOCK:
                STREAMS.pop(sid, None)

    return Response(gen(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/api/state")
def api_state():
    st = load_state()
    return jsonify({
        "sheet_id": st.get("sheet_id", ""),
        "sheet_url": st.get("sheet_url", ""),
        "sheet_title": st.get("sheet_title", ""),
        "saved_at": st.get("saved_at", ""),
        # Fall back to the raw-rows length so the count shows after a
        # domain-match run too, not only after a Checker run.
        "lookup_count": len(st.get("lookup") or {}) or len(st.get("lookup_rows") or []),
        "lookup_saved_at": st.get("lookup_saved_at", ""),
        "gsheet_ready": os.path.exists(GOOGLE_SA_JSON),
        "default_sheet_set": bool(extract_sheet_id(GOOGLE_SHEET_ID)),
    })


@app.route("/health")
def health():
    return jsonify({
        "ok": True, "target_key": TARGET_KEY,
        "gsheet_ready": os.path.exists(GOOGLE_SA_JSON),
        "service_account_email": service_account_email(),
        "service_account_path": GOOGLE_SA_JSON,
        "default_sheet_set": bool(extract_sheet_id(GOOGLE_SHEET_ID)),
    })


# ------------------------- Front end ----------------------------------

PAGE_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>GP Domain-wise Sheet - Halluciantion Detecction Tool</title>
<style>
  :root{--bg:#0f1115;--panel:#171a21;--border:#2a2f3a;--text:#e6e8ec;--muted:#8b93a1;
        --accent:#4f8cff;--ok:#3fb950;--err:#f85149;--warn:#d29922}
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--text);
       font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif}
  header{padding:18px 24px;border-bottom:1px solid var(--border)}
  header h1{margin:0;font-size:16px;font-weight:600}
  header .sub{color:var(--muted);font-size:12px}
  main{padding:24px;max-width:960px;margin:0 auto}
  .panel{background:var(--panel);border:1px solid var(--border);border-radius:10px;
         padding:18px;margin-bottom:18px}
  label{display:block;font-size:12px;color:var(--muted);margin-bottom:6px;
        font-weight:600;text-transform:uppercase;letter-spacing:.04em}
  input[type=file],input[type=text],input[type=password]{
        width:100%;padding:10px;background:#0d0f13;border:1px solid var(--border);
        border-radius:6px;color:var(--text);font-size:13px}
  input[type=file]{border-style:dashed}
  input[type=file]:hover,input[type=text]:focus,input[type=password]:focus{border-color:var(--accent)}
  .hint{color:var(--muted);font-size:12px;margin-top:6px}
  code{background:#0d0f13;padding:2px 6px;border-radius:4px;font-size:12px}
  .radios{display:flex;gap:22px;flex-wrap:wrap}
  .radios label{display:flex;align-items:center;gap:8px;margin:0;font-size:13px;
                color:var(--text);text-transform:none;letter-spacing:0;font-weight:500;cursor:pointer}
  .actions{display:flex;gap:10px;align-items:center;margin-top:8px;flex-wrap:wrap}
  button{padding:10px 18px;border-radius:6px;border:1px solid var(--border);
         background:#22262f;color:var(--text);font-size:13px;font-weight:600;cursor:pointer}
  button.primary{background:var(--accent);border-color:var(--accent);color:#fff}
  button:disabled{opacity:.5;cursor:not-allowed}
  .status{margin-left:auto;font-size:13px;color:var(--muted)}
  .status.ok{color:var(--ok)} .status.err{color:var(--err)} .status.busy{color:var(--warn)}
  .bar{height:8px;background:#0d0f13;border:1px solid var(--border);border-radius:999px;
       overflow:hidden;margin-top:8px}
  .bar > div{height:100%;background:var(--accent);width:0%;transition:width .2s ease}
  .log{background:#0d0f13;border:1px solid var(--border);border-radius:6px;
       padding:12px;max-height:280px;overflow:auto;margin-top:10px;
       font:12px/1.5 ui-monospace,Menlo,Consolas,monospace;
       white-space:pre-wrap;word-break:break-word;color:var(--muted)}
  .log .ok{color:var(--ok)} .log .err{color:var(--err)} .log .warn{color:var(--warn)}
  .error{background:#2a1518;border:1px solid #5c2027;color:#ffb3b8;padding:14px;
         border-radius:6px;white-space:pre-wrap;font-size:13px;word-break:break-word}
  .result a{color:var(--accent);font-weight:600;word-break:break-all}
  table{width:100%;border-collapse:collapse;font-size:13px;margin-top:10px}
  th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--border);vertical-align:top}
  th{color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.04em}
  .counters{display:flex;gap:18px;flex-wrap:wrap;font-size:13px;color:var(--muted);margin-top:6px}
  .counters b{color:var(--text)}
  .shell{display:flex;min-height:100vh}
  .nav{width:208px;flex:none;background:var(--panel);border-right:1px solid var(--border);
       padding:12px 8px;transition:width .15s ease}
  .nav button.item{display:flex;align-items:center;gap:10px;width:100%;margin:0 0 4px;
       padding:9px 10px;text-align:left;background:transparent;border:1px solid transparent;
       color:var(--muted);border-radius:6px;font-size:13px;font-weight:600;white-space:nowrap}
  .nav button.item:hover{color:var(--text);background:#1d2129}
  .nav button.item.active{color:var(--text);background:#22262f;border-color:var(--border)}
  .nav .ico{width:18px;text-align:center;flex:none}
  .shell.collapsed .nav{width:58px}
  .shell.collapsed .nav .lbl{display:none}
  .content{flex:1;min-width:0}
  textarea.code{width:100%;min-height:200px;padding:12px;background:#0d0f13;
       border:1px solid var(--border);border-radius:6px;color:var(--text);
       font:13px/1.7 ui-monospace,Menlo,Consolas,monospace;resize:vertical;white-space:pre;
       overflow:auto}
  textarea.code:focus{outline:none;border-color:var(--accent)}
  .remember{background:#0d0f13;border:1px solid var(--border);border-radius:6px;
       padding:12px;font-size:13px;word-break:break-all}
  .remember a{color:var(--accent)}
  @media (max-width:640px){.nav{width:58px}.nav .lbl{display:none}}
  .hidden{display:none!important}
</style>
</head>
<body>
<div class="shell" id="shell">
<aside class="nav" id="nav">
  <button class="item" id="navToggle" type="button" title="Collapse / expand menu">
    <span class="ico">&#9776;</span><span class="lbl">Menu</span></button>
  <button class="item active" data-page="check" type="button">
    <span class="ico">&#10003;</span><span class="lbl">Checker</span></button>
  <button class="item" data-page="match" type="button">
    <span class="ico">&#8981;</span><span class="lbl">Keyword Match</span></button>
</aside>
<div class="content">
<header>
  <h1>Domain-wise Sheet - Halluciantion Detecction Tool</h1>
  <span class="sub">local backend &middot; output to Google Sheets or .xlsx</span>
</header>

<main id="page-check">
  <div class="panel">
    <div style="margin-bottom:16px">
      <label for="fund">Domain Excel (.xlsx)</label>
      <input type="file" id="fund" accept=".xlsx,.xls">
      <div class="hint">
        Reads sheet <code>Sheet1</code> (falls back to the first sheet if absent).
        <b>Special:</b> <code>Firm AUM</code> + <code>Firm AUM Reference URL</code>,
        <code>Firm AUM Date</code> + <code>Firm AUM Date Reference URL</code>.<br>
        <b>Plain list</b> (split on <code>,</code> and <code>|</code>):
        Firm Geo Focus, Firm Country Focus, Firm Industry Focus.<br>
        <b>Asset-class list</b> (<code>Class - Data1, Data2 | Class - Data3</code>;
        only the <i>Data</i> part is matched): Asset Class Geo Focus,
        Asset Class Country Focus, Asset Class Industry Focus, Investment Type.<br>
        <b>Asset-class date</b> (<code>Class - DD/MM/YYYY | &hellip;</code>):
        Asset Class AUM Date.<br>
        <b>Asset-class numbers</b> (<code>Class - 1,250 | &hellip;</code>):
        Asset Class AUM, Total Capital Invested.<br>
        <b>Ignored:</b> Firm AUM Currency, Asset Class,
        Total Capital Invested Currency.<br>
        Every reference URL cell may contain multiple URLs separated by <code>|</code>.
      </div>
    </div>
    <div style="margin-bottom:16px">
      <label for="lookup">URL-wise CSV</label>
      <input type="file" id="lookup" accept=".csv,text/csv">
      <div class="hint">Must contain <code>URL</code> and <code>s3status</code>.</div>
    </div>
    <div style="margin-bottom:16px">
      <label for="user">Login ID</label>
      <input type="text" id="user" autocomplete="username" spellcheck="false">
    </div>
    <div style="margin-bottom:16px">
      <label for="pass">Password</label>
      <input type="password" id="pass" autocomplete="current-password">
    </div>

    <div style="margin-bottom:16px">
      <label>Output</label>
      <div class="radios">
        <label><input type="radio" name="mode" value="gsheet" checked> Google Sheet</label>
        <label><input type="radio" name="mode" value="excel"> Download .xlsx</label>
      </div>
    </div>
    <div style="margin-bottom:16px" id="sheetWrap">
      <label for="sheet">Google Sheet URL</label>
      <input type="text" id="sheet" spellcheck="false"
             placeholder="Leave blank to create a new sheet named after your Fund file">
      <div class="hint" id="sheetHint">checking Google setup&hellip;</div>
      <div style="margin-top:12px">
        <label for="shareEmail">Your Google email <span style="color:var(--err)">*</span></label>
        <input type="text" id="shareEmail" spellcheck="false" autocomplete="email"
               inputmode="email" required aria-required="true" placeholder="you@company.com">
        <div class="hint">Required. The new sheet is shared with this address as Editor, so it
          shows up in your Google Drive. (Not needed if you paste an existing sheet URL above.)</div>
      </div>
    </div>

    <div class="actions">
      <button class="primary" id="run">Run check</button>
      <button id="dl" disabled>Download .xlsx</button>
      <span class="status" id="status">idle</span>
    </div>

    <div class="hidden" id="progressWrap" style="margin-top:16px">
      <label>Progress</label>
      <div class="bar"><div id="bar"></div></div>
      <div class="counters">
        <span>Processed: <b id="cDone">0</b> / <b id="cTotal">0</b></span>
        <span>AUM: <b id="cAum">0</b></span>
        <span>Date: <b id="cDate">0</b></span>
        <span>plain pass: <b id="cPlain">0</b></span>
        <span>asset-list pass: <b id="cAssetList">0</b></span>
        <span>asset-date pass: <b id="cAssetDate">0</b></span>
        <span>asset-AUM pass: <b id="cAssetAum">0</b></span>
      </div>
      <div class="log" id="logBox"></div>
    </div>
  </div>

  <div class="panel hidden result" id="resultPanel">
    <label>Result</label>
    <div id="resultBox"></div>
  </div>

  <div class="panel hidden" id="errorPanel">
    <label>Error</label>
    <div class="error" id="errorBox"></div>
  </div>

  <div class="panel">
    <label>Output sheets (12)</label>
    <table>
      <tr><th>Sheet</th><th>Contents</th></tr>
      <tr><td><code>AUM_results</code></td><td>Firm AUM match columns + Per-URL Results</td></tr>
      <tr><td><code>Date_results</code></td><td>Firm AUM Date match columns + Per-URL Results</td></tr>
      <tr><td><code>Firm_Geo_Focus_results</code></td><td rowspan="3">Plain list &mdash; Values searched/count/found/missing, All values found, Per-value evidence, Per-URL Results</td></tr>
      <tr><td><code>Firm_Country_Focus_results</code></td></tr>
      <tr><td><code>Firm_Industry_Focus_results</code></td></tr>
      <tr><td><code>Asset_Class_Geo_Focus_results</code></td><td rowspan="4">Asset-class list &mdash; Class + data searched, then plain-list columns on the Data values only</td></tr>
      <tr><td><code>Asset_Class_Country_Focus_results</code></td></tr>
      <tr><td><code>Asset_Class_Industry_Focus_results</code></td></tr>
      <tr><td><code>Investment_Type_results</code></td></tr>
      <tr><td><code>Asset_Class_AUM_Date_results</code></td><td>Asset-class date &mdash; Class + data searched, Dates searched/parsed/unparseable/found/missing, All dates found, Per-date evidence</td></tr>
      <tr><td><code>Asset_Class_AUM_results</code></td><td rowspan="2">Asset-class numbers &mdash; Class + data searched, Numbers searched/count/found/missing, All numbers found, Per-number evidence</td></tr>
      <tr><td><code>Total_Capital_Invested_results</code></td></tr>
    </table>
  </div>
</main>

<main id="page-match" class="hidden">
  <div class="panel">
    <label>Input</label>
    <div class="radios">
      <label><input type="radio" name="mmode" value="gsheet" checked> Google Sheet (remembered)</label>
      <label><input type="radio" name="mmode" value="excel"> Download .xlsx</label>
    </div>

    <div id="mSheetWrap" style="margin-top:14px">
      <label>Google Sheet</label>
      <div class="remember" id="mRemember">Loading&hellip;</div>
      <div style="margin:14px 0 0">
        <label for="mSheet">Use a different sheet (optional)</label>
        <input type="text" id="mSheet" spellcheck="false"
               placeholder="Leave blank to use the remembered sheet">
      </div>
      <div style="margin-top:12px">
        <label for="mShareEmail">Your Google email (only if creating a new sheet)</label>
        <input type="text" id="mShareEmail" spellcheck="false" autocomplete="email"
               inputmode="email" placeholder="you@company.com">
      </div>
    </div>

    <div style="margin-bottom:16px;margin-top:16px">
      <label for="mDomainsKeywords">Domains &amp; keywords &mdash; one or more per line</label>
      <textarea class="code" id="mDomainsKeywords" spellcheck="false">eatonvance.com: "AUM", "assets"
morganstanley.com, blackrock.com: "AUM", "assets"</textarea>
      <div class="hint">
        Format: <code>domain[, domain]: "kw", "kw"</code>.
        Several domains before the colon share the same keyword list.
        Duplicate domains across lines are unioned (dedupe case-insensitive).
        The domain must match a value in the URL-wise CSV&rsquo;s
        <code>Domain</code> column exactly (case-insensitive, surrounding
        whitespace ignored). All keywords for a domain are matched against
        every page whose <code>s3status</code> link sits under that domain.
        Keywords are matched as whole words, case-insensitive, whitespace-flexible.
      </div>
    </div>

    <div style="margin-bottom:16px">
      <label for="mUser">Login ID</label>
      <input type="text" id="mUser" autocomplete="username" spellcheck="false"
             placeholder="Leave blank to reuse the one typed on the Checker page">
    </div>
    <div style="margin-bottom:16px">
      <label for="mPass">Password</label>
      <input type="password" id="mPass" autocomplete="current-password">
    </div>
    <div style="margin-bottom:16px">
      <label for="mLookup">URL-wise CSV (optional)</label>
      <input type="file" id="mLookup" accept=".csv,text/csv">
      <div class="hint" id="mLookupHint"></div>
    </div>

    <div class="actions">
      <button class="primary" id="mRun">Match</button>
      <button id="mDl" disabled>Download .xlsx</button>
      <span class="status" id="mStatus">idle</span>
    </div>

    <div class="hidden" id="mProgressWrap" style="margin-top:16px">
      <label>Progress</label>
      <div class="bar"><div id="mBar"></div></div>
      <div class="counters">
        <span>Processed: <b id="mDone">0</b> / <b id="mTotal">0</b></span>
      </div>
      <div class="log" id="mLog"></div>
    </div>
  </div>

  <div class="panel hidden result" id="mResultPanel">
    <label>Result</label>
    <div id="mResultBox"></div>
  </div>

  <div class="panel hidden" id="mErrorPanel">
    <label>Error</label>
    <div class="error" id="mErrorBox"></div>
  </div>

  <div class="panel">
    <label>Output &mdash; one row per domain (sheet <code>Domain_keyword_matches</code>)</label>
    <table>
      <tr><th>Column</th><th>Meaning</th></tr>
      <tr><td><code>Domain</code></td><td>The input domain (as normalised)</td></tr>
      <tr><td><code>Keywords</code></td><td>All keywords assigned to this domain, joined with <code>|</code></td></tr>
      <tr><td><code>URLs scanned</code></td><td>Number of URL-wise rows whose Domain matched, with a non-blank s3status</td></tr>
      <tr><td><code>Matched with</code></td><td>Keywords actually found on at least one page, joined with <code>|</code></td></tr>
      <tr><td><code>URLs matched</code></td><td>URL values (from the URL-wise CSV) where at least one keyword was found, joined with <code>|</code></td></tr>
      <tr><td><code>Evidence</code></td><td><code>URL :: context</code> lines (10 words each side, match in <code>[]</code>), joined with <code>|</code></td></tr>
      <tr><td><code>Notes</code></td><td>Fetch errors, blank s3status, "domain not in URL-wise file", etc.</td></tr>
    </table>
  </div>
</main>
</div>
</div>

<script>
(function () {
  "use strict";
  var $ = function (id) { return document.getElementById(id); };
  function store(k, v) { try { localStorage.setItem(k, v); } catch (_) {} }
  function load(k) { try { return localStorage.getItem(k); } catch (_) { return null; } }

  var shell = $("shell");
  var pages = { check: $("page-check"), match: $("page-match") };
  var navItems = document.querySelectorAll("#nav .item[data-page]");

  if (load("navCollapsed") === "1") shell.classList.add("collapsed");
  $("navToggle").addEventListener("click", function () {
    var c = shell.classList.toggle("collapsed");
    store("navCollapsed", c ? "1" : "0");
  });

  function show(p) {
    if (!pages[p]) p = "check";
    Object.keys(pages).forEach(function (k) { pages[k].classList.toggle("hidden", k !== p); });
    Array.prototype.forEach.call(navItems, function (b) {
      b.classList.toggle("active", b.getAttribute("data-page") === p);
    });
    try { history.replaceState(null, "", "#" + p); } catch (_) {}
    if (p === "match") refreshState();
  }
  Array.prototype.forEach.call(navItems, function (b) {
    b.addEventListener("click", function () { show(b.getAttribute("data-page")); });
  });
  window.addEventListener("hashchange", function () {
    show((location.hash || "").replace("#", "") || "check");
  });

  // ---------- Checker page ----------
  var fundEl = $("fund"), lookupEl = $("lookup"), userEl = $("user"), passEl = $("pass"),
      sheetEl = $("sheet"), shareEl = $("shareEmail"), sheetWrap = $("sheetWrap"),
      sheetHint = $("sheetHint"),
      runBtn = $("run"), dlBtn = $("dl"), statusEl = $("status"),
      errorPanel = $("errorPanel"), errorBox = $("errorBox"),
      resultPanel = $("resultPanel"), resultBox = $("resultBox"),
      progressWrap = $("progressWrap"), bar = $("bar"),
      cDone = $("cDone"), cTotal = $("cTotal"), cAum = $("cAum"),
      cDate = $("cDate"), cPlain = $("cPlain"), cAssetList = $("cAssetList"),
      cAssetDate = $("cAssetDate"), cAssetAum = $("cAssetAum"),
      logBox = $("logBox");

  var resultBlob = null, finished = false, seen = 0, total = 0,
      aumCount = 0, dateCount = 0, plainCount = 0, assetListCount = 0,
      assetDateCount = 0, assetAumCount = 0, logLines = 0;

  function mode() { return document.querySelector('input[name="mode"]:checked').value; }
  function setStatus(t, c) { statusEl.textContent = t; statusEl.className = "status" + (c ? " " + c : ""); }
  function showError(m) { errorBox.textContent = m; errorPanel.classList.remove("hidden"); }
  function log(line, cls) {
    var el = document.createElement("div");
    if (cls) el.className = cls;
    el.textContent = line;
    logBox.appendChild(el);
    if (++logLines > 500) { logBox.removeChild(logBox.firstChild); logLines--; }
    logBox.scrollTop = logBox.scrollHeight;
  }
  function resetProgress() {
    seen = total = aumCount = dateCount = 0;
    plainCount = assetListCount = assetDateCount = assetAumCount = 0;
    logLines = 0;
    cDone.textContent = "0"; cTotal.textContent = "0";
    cAum.textContent = "0"; cDate.textContent = "0";
    cPlain.textContent = "0"; cAssetList.textContent = "0";
    cAssetDate.textContent = "0"; cAssetAum.textContent = "0";
    bar.style.width = "0%"; logBox.textContent = "";
    progressWrap.classList.remove("hidden");
  }
  function updateBar() {
    if (total > 0) bar.style.width = Math.min(100, (seen / total) * 100).toFixed(1) + "%";
  }
  function refreshMode() { sheetWrap.classList.toggle("hidden", mode() !== "gsheet"); }
  Array.prototype.forEach.call(document.querySelectorAll('input[name="mode"]'),
    function (r) { r.addEventListener("change", refreshMode); });
  refreshMode();
  try { shareEl.value = localStorage.getItem("shareEmail") || ""; } catch (_) {}

  fetch("/health").then(function (r) { return r.json(); }).then(function (h) {
    if (h.gsheet_ready) {
      sheetHint.textContent = h.default_sheet_set
        ? "Leave blank to use the default sheet set in frontend.py. To use another sheet, paste its URL (shared with " + (h.service_account_email || "the service account") + " as Editor)."
        : "Leave blank to create a new sheet named after your Fund file and share it with the email below. To use an existing sheet, paste its URL (shared with " + (h.service_account_email || "the service account") + " as Editor).";
    } else {
      sheetHint.textContent = "Service-account key not found at " + h.service_account_path + ". Use the .xlsx option or set GOOGLE_SA_JSON.";
    }
  }).catch(function () { sheetHint.textContent = ""; });

  function triggerDownload() {
    if (!resultBlob) return;
    var a = document.createElement("a");
    a.href = URL.createObjectURL(resultBlob);
    a.download = "domain_results.xlsx";
    document.body.appendChild(a); a.click();
    URL.revokeObjectURL(a.href); a.remove();
  }
  dlBtn.addEventListener("click", triggerDownload);

  runBtn.addEventListener("click", async function () {
    errorPanel.classList.add("hidden"); resultPanel.classList.add("hidden");
    dlBtn.disabled = true; resultBlob = null; finished = false;
    var fund = fundEl.files[0], lookup = lookupEl.files[0];
    if (!fund || !lookup) {
      showError("Select the Fund .xlsx and the URL-wise CSV first.");
      setStatus("missing files", "err"); return;
    }
    if (mode() === "gsheet" && !sheetEl.value.trim() &&
        !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(shareEl.value.trim())) {
      showError("Enter your Google email (the new Google Sheet is shared with it), or paste an existing Google Sheet URL.");
      setStatus("email needed", "err"); shareEl.focus(); return;
    }
    try { if (shareEl.value.trim()) localStorage.setItem("shareEmail", shareEl.value.trim()); } catch (_) {}

    runBtn.disabled = true; setStatus("starting\u2026", "busy"); resetProgress();

    var fd = new FormData();
    fd.append("fund_xlsx", fund);
    fd.append("lookup_csv", lookup);
    fd.append("user", userEl.value);
    fd.append("pass", passEl.value);
    fd.append("mode", mode());
    fd.append("sheet", sheetEl.value);
    fd.append("share_email", shareEl.value);

    var startRes;
    try { startRes = await fetch("/api/run-stream", { method: "POST", body: fd }); }
    catch (e) { runBtn.disabled = false; setStatus("network error", "err"); showError(String(e)); return; }
    if (!startRes.ok) {
      runBtn.disabled = false; setStatus("error " + startRes.status, "err");
      var msg = await startRes.text();
      try { msg = JSON.parse(msg).error || msg; } catch (_) {}
      showError(msg); return;
    }
    var sid = (await startRes.json()).stream_id;
    setStatus("running\u2026", "busy");

    var src = new EventSource("/api/progress/" + sid);

    src.addEventListener("progress", function (ev) {
      var d = JSON.parse(ev.data);
      seen = d.i; total = d.total;
      cDone.textContent = String(seen); cTotal.textContent = String(total);
      if (d.aum_flag === "true") aumCount++;
      if (d.date_flag === "true") dateCount++;
      var plainPass = 0, alPass = 0, adPass = 0, aaPass = 0;
      Object.keys(d.plain || {}).forEach(function (k) { if (d.plain[k] === "true") plainPass++; });
      Object.keys(d.asset_list || {}).forEach(function (k) { if (d.asset_list[k] === "true") alPass++; });
      Object.keys(d.asset_date || {}).forEach(function (k) { if (d.asset_date[k] === "true") adPass++; });
      Object.keys(d.asset_aum || {}).forEach(function (k) { if (d.asset_aum[k] === "true") aaPass++; });
      if (plainPass) plainCount++;
      if (alPass) assetListCount++;
      if (adPass) assetDateCount++;
      if (aaPass) assetAumCount++;
      cAum.textContent = String(aumCount);
      cDate.textContent = String(dateCount);
      cPlain.textContent = String(plainCount);
      cAssetList.textContent = String(assetListCount);
      cAssetDate.textContent = String(assetDateCount);
      cAssetAum.textContent = String(assetAumCount);

      var cls = (d.aum_flag === "true" || d.date_flag === "true") ? "ok" : "warn";
      log("[" + d.i + "/" + d.total + "] AUM " + d.aum_flag +
          " | Date " + d.date_flag +
          " | plain+" + plainPass + " asset+" + alPass +
          " adate+" + adPass + " aAUM+" + aaPass, cls);
      updateBar();
    });

    src.addEventListener("status", function (ev) {
      var d = JSON.parse(ev.data); setStatus(d.message, "busy"); log(d.message, "warn");
    });
    src.addEventListener("done", function (ev) {
      var d = JSON.parse(ev.data); finished = true;
      if (d.mode === "gsheet") {
        resultBox.textContent = "";
        var a = document.createElement("a");
        a.href = d.url; a.target = "_blank"; a.rel = "noopener"; a.textContent = d.url;
        resultBox.appendChild(a);
        resultPanel.classList.remove("hidden");
        setStatus("done \u2014 written to Google Sheets", "ok");
        window.dispatchEvent(new Event("sheet-saved"));
      } else {
        var bin = atob(d.b64);
        var bytes = new Uint8Array(bin.length);
        for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        resultBlob = new Blob([bytes],
          { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        dlBtn.disabled = false; setStatus("done \u2014 downloading", "ok");
        triggerDownload();
      }
    });
    src.addEventListener("fail", function (ev) {
      var d = {}; try { d = JSON.parse(ev.data || "{}"); } catch (_) {}
      finished = true; setStatus("error", "err"); showError(d.message || "Run failed");
    });
    src.addEventListener("error", function () {
      if (finished) return;
      setStatus("connection lost", "err"); showError("Lost connection to the local server.");
      src.close(); runBtn.disabled = false;
    });
    src.addEventListener("end", function () { src.close(); runBtn.disabled = false; updateBar(); });
  });

  // ---------- Keyword Match page ----------
  var remember = $("mRemember"), lookupHint = $("mLookupHint");
  function refreshState() {
    fetch("/api/state").then(function (r) { return r.json(); }).then(function (s) {
      remember.textContent = "";
      if (s.sheet_id) {
        var b = document.createElement("b");
        b.textContent = s.sheet_title || "(untitled sheet)";
        remember.appendChild(b); remember.appendChild(document.createElement("br"));
        if (/^https:\/\//.test(s.sheet_url || "")) {
          var a = document.createElement("a");
          a.href = s.sheet_url; a.target = "_blank"; a.rel = "noopener";
          a.textContent = s.sheet_url; remember.appendChild(a);
        }
        if (s.saved_at) {
          var t = document.createElement("div");
          t.className = "hint"; t.textContent = "Remembered from the last run on " + s.saved_at;
          remember.appendChild(t);
        }
      } else {
        remember.textContent = "No sheet remembered yet. Run a Google Sheet check on the Checker page first, or paste a sheet URL below.";
      }
      if (s.lookup_count) {
        lookupHint.textContent = "Using the saved URL-wise data (" + s.lookup_count + " row(s)" +
          (s.lookup_saved_at ? ", saved " + s.lookup_saved_at : "") +
          "). Upload a CSV only to replace it.";
      } else {
        lookupHint.textContent = "No URL-wise data saved yet \u2014 upload the URL-wise CSV once.";
      }
    }).catch(function () { remember.textContent = "Could not read the saved state."; });
  }
  window.addEventListener("sheet-saved", refreshState);

  function mmode() { return document.querySelector('input[name="mmode"]:checked').value; }
  function refreshMode2() {
    // The Keyword Match page no longer has an "Upload Excel" input;
    // both output modes share the same inputs above.
    $("mSheetWrap").classList.toggle("hidden", mmode() !== "gsheet");
  }
  Array.prototype.forEach.call(document.querySelectorAll('input[name="mmode"]'),
    function (r) { r.addEventListener("change", refreshMode2); });
  refreshMode2();

  var mRun = $("mRun"), mDl = $("mDl"), mStatus = $("mStatus");
  var mLog = $("mLog"), mBar = $("mBar");
  var mErrPanel = $("mErrorPanel"), mErrBox = $("mErrorBox");
  var mResPanel = $("mResultPanel"), mResBox = $("mResultBox");
  var mSeen = 0, mTotal = 0, mFinished = false, mBlob = null, mLogLines = 0;

  function mSetStatus(t, c) { mStatus.textContent = t; mStatus.className = "status" + (c ? " " + c : ""); }
  function mShowError(m) { mErrBox.textContent = m; mErrPanel.classList.remove("hidden"); }
  function mLogLine(line, cls) {
    var el = document.createElement("div");
    if (cls) el.className = cls;
    el.textContent = line; mLog.appendChild(el);
    if (++mLogLines > 500) { mLog.removeChild(mLog.firstChild); mLogLines--; }
    mLog.scrollTop = mLog.scrollHeight;
  }
  function mTriggerDownload() {
    if (!mBlob) return;
    var a = document.createElement("a");
    a.href = URL.createObjectURL(mBlob);
    a.download = mBlob.name || "domain_keyword_matches.xlsx";
    document.body.appendChild(a); a.click(); URL.revokeObjectURL(a.href); a.remove();
  }
  mDl.addEventListener("click", mTriggerDownload);

  mRun.addEventListener("click", async function () {
    mErrPanel.classList.add("hidden"); mResPanel.classList.add("hidden");
    mDl.disabled = true; mBlob = null; mFinished = false;
    mSeen = 0; mTotal = 0; mLogLines = 0;
    mLog.textContent = ""; mBar.style.width = "0%";
    $("mProgressWrap").classList.remove("hidden");
    $("mDone").textContent = "0"; $("mTotal").textContent = "0";

    mRun.disabled = true; mSetStatus("starting\u2026", "busy");

    var fd = new FormData();
    fd.append("mode", mmode());
    fd.append("domains_keywords", $("mDomainsKeywords").value);
    fd.append("sheet", $("mSheet").value);
    fd.append("share_email", $("mShareEmail").value);
    fd.append("user", $("mUser").value || $("user").value);
    fd.append("pass", $("mPass").value || $("pass").value);
    var lk = $("mLookup").files[0];
    if (lk) fd.append("lookup_csv", lk);

    var startRes;
    try { startRes = await fetch("/api/domain-match-stream", { method: "POST", body: fd }); }
    catch (e) { mRun.disabled = false; mSetStatus("network error", "err"); mShowError(String(e)); return; }
    if (!startRes.ok) {
      mRun.disabled = false; mSetStatus("error " + startRes.status, "err");
      var msg = await startRes.text();
      try { msg = JSON.parse(msg).error || msg; } catch (_) {}
      mShowError(msg); return;
    }
    var sid = (await startRes.json()).stream_id;
    mSetStatus("running\u2026", "busy");

    var src = new EventSource("/api/progress/" + sid);
    src.addEventListener("progress", function (ev) {
      var d = JSON.parse(ev.data);
      mSeen = d.i; mTotal = d.total;
      $("mDone").textContent = String(mSeen); $("mTotal").textContent = String(mTotal);
      if (mTotal > 0) mBar.style.width = Math.min(100, (mSeen / mTotal) * 100).toFixed(1) + "%";
      var line = "[" + d.i + "/" + d.total + "] " + d.domain + " \u2014 ";
      line += d.matched
        ? "matched " + (d.matched_with || "(none)") +
          " in " + d.url_hits + " of " + d.urls_scanned + " URL(s)"
        : "no match (" + d.urls_scanned + " URL(s) checked)";
      if (d.note) line += "  " + d.note;
      mLogLine(line, d.matched ? "ok" : (d.note ? "err" : "warn"));
    });
    src.addEventListener("status", function (ev) {
      var d = JSON.parse(ev.data); mSetStatus(d.message, "busy"); mLogLine(d.message, "warn");
    });
    src.addEventListener("done", function (ev) {
      var d = JSON.parse(ev.data); mFinished = true;
      mResBox.textContent = "";
      if (d.mode === "excel") {
        var bin = atob(d.b64);
        var bytes = new Uint8Array(bin.length);
        for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        mBlob = new Blob([bytes],
          { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        mBlob.name = d.filename || "domain_keyword_matches.xlsx";
        mDl.disabled = false; mSetStatus("done \u2014 downloading", "ok");
        mTriggerDownload();
      } else {
        if (/^https:\/\//.test(d.url || "")) {
          var a = document.createElement("a");
          a.href = d.url; a.target = "_blank"; a.rel = "noopener";
          a.textContent = (d.title ? d.title + " \u2014 " : "") + d.url;
          mResBox.appendChild(a);
        }
        mSetStatus("done \u2014 written to Google Sheets", "ok");
      }
      var p = document.createElement("div");
      p.className = "hint";
      p.textContent = (d.tab || "Domain_keyword_matches") + ": " + d.rows + " domain row(s)";
      mResBox.appendChild(p);
      mResPanel.classList.remove("hidden");
    });
    src.addEventListener("fail", function (ev) {
      var d = {}; try { d = JSON.parse(ev.data || "{}"); } catch (_) {}
      mFinished = true; mSetStatus("error", "err"); mShowError(d.message || "Run failed");
    });
    src.addEventListener("error", function () {
      if (mFinished) return;
      mSetStatus("connection lost", "err"); mShowError("Lost connection to the local server.");
      src.close(); mRun.disabled = false;
    });
    src.addEventListener("end", function () { src.close(); mRun.disabled = false; });
  });

  show((location.hash || "").replace("#", "") || "check");
})();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    print(f"Open http://localhost:{PORT}")
    app.run(host="127.0.0.1", port=PORT, debug=False, threaded=True)