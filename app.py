import base64
import hashlib
import hmac
import io
import json
import re
import textwrap
import threading
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from dateutil import parser as dateparser
from PIL import Image, ImageDraw, ImageFont, ImageOps
from google.oauth2.credentials import Credentials as UserCredentials
from google.oauth2.service_account import Credentials

from booking_parser import (EVENTS, PAYMENT_METHODS, normalize_contact, parse_date,
                            parse_text, split_amount)

# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------
BASE = Path(__file__).parent
try:
    _icon = Image.open(BASE / "logo_prowess.png")
except Exception:  # noqa
    _icon = "🧾"
st.set_page_config(page_title="Stall Booking", page_icon=_icon, layout="centered")

BASE = Path(__file__).parent
CRED_FILE = BASE / "credentials.json"      # downloaded from Google Cloud (see guide)
CONFIG_FILE = BASE / "config.json"         # created automatically by the sidebar
FILES_DIR = BASE / "booking_files"         # logos / receipts are always saved here too
TZ = ZoneInfo("Asia/Karachi")

# ----------------------------------------------------------------------------
# Look and feel (purple / gold theme, built for phones)
# ----------------------------------------------------------------------------
EVENT_COLORS = {"Jacranda": "#B3267F", "AMBLEM": "#E8702A", "Qawali Night": "#1FA39A"}
THEME_CSS = """
.block-container{padding-top:4.2rem;padding-bottom:3rem;max-width:760px}
.sbk-hdr{background:#fff;border-radius:18px;padding:12px 14px 10px;margin-bottom:12px;text-align:center}
.sbk-logos{display:flex;align-items:center;justify-content:center;gap:16px}
.sbk-logos img{height:84px;width:auto;max-width:48%}
.sbk-sub{color:#B3267F;font-size:14px;font-weight:600;margin-top:6px;letter-spacing:.3px}
.sbk-title{color:#3D1A5C;font-size:22px;font-weight:700}
.sbk-cards{display:grid;grid-template-columns:repeat(2,1fr);gap:10px;margin:6px 0 10px}
@media(min-width:640px){.sbk-cards{grid-template-columns:repeat(4,1fr)}}
.sbk-card{background:#fff;border-radius:12px;padding:10px 12px}
.sbk-k{font-size:12px;color:#7A4A9C}
.sbk-v{font-size:21px;font-weight:700;color:#3D1A5C;line-height:1.25}
.sbk-card.sbk-pf{background:#B3267F}
.sbk-card.sbk-pf .sbk-k{color:#FBE3F1}
.sbk-card.sbk-pf .sbk-v{color:#fff}
.sbk-evs{display:grid;grid-template-columns:1fr;gap:8px;margin-bottom:10px}
@media(min-width:640px){.sbk-evs{grid-template-columns:repeat(3,1fr)}}
.sbk-ev{background:#fff;border-radius:12px;padding:9px 12px}
.sbk-row{display:flex;align-items:center;justify-content:space-between;gap:8px}
.sbk-tag{display:inline-block;color:#fff;font-size:12px;padding:2px 10px;border-radius:12px;white-space:nowrap}
.sbk-n{font-size:13px;color:#3D1A5C;text-align:right}
.sbk-bar{height:6px;border-radius:3px;background:#EFE3F7;margin-top:8px}
.sbk-bar i{display:block;height:6px;border-radius:3px}
.sbk-dia{display:flex;justify-content:center;gap:8px;margin:10px 0 6px}
.sbk-dia i{width:9px;height:9px;transform:rotate(45deg);display:inline-block}
button[data-baseweb="tab"]{padding:10px 9px}
button[data-baseweb="tab"] p{font-size:14px}
div[data-baseweb="tab-highlight"]{background-color:#F5C230 !important}
button[data-testid="stBaseButton-primary"],button[data-testid="stBaseButton-primaryFormSubmit"],button[kind="primary"],button[kind="primaryFormSubmit"]{background-color:#F5C230 !important;border-color:#F5C230 !important;border-radius:24px;min-height:46px}
button[data-testid="stBaseButton-primary"] p,button[data-testid="stBaseButton-primaryFormSubmit"] p,button[kind="primary"] p,button[kind="primaryFormSubmit"] p{color:#3D1A5C !important;font-weight:700}
button[data-testid="stBaseButton-secondary"],button[data-testid="stBaseButton-secondaryFormSubmit"]{border-radius:24px;min-height:44px;border-color:#F5C230}
input,textarea{font-size:16px !important}
"""
st.markdown(f"<style>{THEME_CSS}</style>", unsafe_allow_html=True)


@st.cache_resource
def _logo_uri(name):
    p = BASE / name
    return ("data:image/png;base64," + base64.b64encode(p.read_bytes()).decode()) if p.exists() else None


def render_header():
    a, b = _logo_uri("logo_prowess.png"), _logo_uri("logo_expo.png")
    imgs = "".join(f'<img src="{u}" alt="{alt}"/>' for u, alt in ((a, "PROWESS"), (b, "Be Pakistani Buy Pakistani")) if u)
    top = f'<div class="sbk-logos">{imgs}</div>' if imgs else '<div class="sbk-title">Be Pakistani, Buy Pakistani</div>'
    st.markdown(f'<div class="sbk-hdr">{top}<div class="sbk-sub">Stall booking records</div></div>',
                unsafe_allow_html=True)


render_header()
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

HEADERS = ["Booking ID", "Saved At", "Name", "Brand Name", "Brand Items",
           "Booked For (All Events)", "This Event", "Paid Amount (This Event)",
           "Total Paid (Whole Booking)", "Booking Date",
           "Contact", "Logo", "Payment Receipt", "Payment Method", "Received By",
           "Other IDs of This Booking"]
LAST_COL = "P"      # 16 columns: A..P
EVENT_CODES = {"Jacranda": "JAC", "AMBLEM": "AMB", "Qawali Night": "QWL"}
COL_BRAND, COL_CONTACT = HEADERS.index("Brand Name"), HEADERS.index("Contact")   # 0-based


# ----------------------------------------------------------------------------
# Booking receipt (picture + text to send to the stall owner)
# ----------------------------------------------------------------------------
EXPO_TITLE = "Be Pakistani, Buy Pakistani 2026"   # shown at the top of the receipt - change if needed


def _font(size, bold=False):
    names = ["DejaVuSans-Bold.ttf", "arialbd.ttf"] if bold else ["DejaVuSans.ttf", "arial.ttf"]
    for n in names:
        try:
            return ImageFont.truetype(n, size)
        except OSError:
            continue
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def fmt_rs(x):
    x = float(x)
    return f"Rs {x:,.0f}" if x.is_integer() else f"Rs {x:,.2f}"


def receipt_rows(brand, booked_by, events, ids, amounts, paid, method, received_by, booking_date):
    """Each row: (label, [lines shown on the picture], single-line text for copy/paste)."""
    rows = [("Stall", [brand], brand),
            ("Stall booked by", [booked_by], booked_by),
            ("Booking ID(s)", [f"{ids[e]} ({e})" for e in events],
             ", ".join(f"{ids[e]} ({e})" for e in events))]
    if paid > 0:
        lines = [fmt_rs(paid)]
        text = fmt_rs(paid)
        if len(events) > 1:
            lines += [f"{e}: {fmt_rs(amounts[e])}" for e in events]
            text += " (" + ", ".join(f"{e} {fmt_rs(amounts[e])}" for e in events) + ")"
    else:
        lines, text = ["Rs 0 (no payment yet)"], "Rs 0 (no payment yet)"
    rows += [("Paid amount", lines, text), ("Payment method", [method], method)]
    if received_by:
        rows.append(("Received by", [received_by], received_by))
    rows.append(("Date of booking", [booking_date], booking_date))
    return rows


def receipt_text(rows):
    out = [f"Booking confirmed - {EXPO_TITLE}"]
    out += [f"{label}: {text}" for label, _lines, text in rows]
    out.append("Please keep your booking ID for any future contact.")
    return "\n".join(out)


def _logo_img(name, height):
    try:
        im = Image.open(BASE / name).convert("RGB")
        return im.resize((round(im.width * height / im.height), height), Image.LANCZOS)
    except Exception:  # noqa
        return None


def make_receipt_png(rows):
    W, M, LABEL_W, LH = 900, 50, 290, 40
    PURPLE, GOLD = (61, 26, 92), (245, 194, 48)
    f_title, f_sub = _font(40, True), _font(24)
    f_label, f_val = _font(24), _font(28, True)
    laid, body_h = [], 0
    for label, lines, _text in rows:
        wrapped = []
        for ln in lines:
            wrapped += textwrap.wrap(str(ln), width=28) or [""]
        laid.append((label, wrapped))
        body_h += len(wrapped) * LH + 22
    HEAD, FOOT = 190, 150
    H = HEAD + 30 + body_h + FOOT
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, W, HEAD], fill=PURPLE)
    d.rectangle([0, HEAD, W, HEAD + 8], fill=GOLD)
    d.text((M, 50), "Booking", font=f_title, fill="white")
    d.text((M, 100), "Confirmation", font=f_title, fill=GOLD)
    logo = _logo_img("logo_expo.png", 120)
    if logo:
        bw, bh = logo.width + 30, 150
        x0 = W - M - bw
        d.rounded_rectangle([x0, 20, x0 + bw, 20 + bh], radius=22, fill="white")
        img.paste(logo, (x0 + 15, 20 + (bh - logo.height) // 2))
    y = HEAD + 40
    for label, wrapped in laid:
        d.text((M, y + 6), label, font=f_label, fill=(122, 74, 156))
        for i, ln in enumerate(wrapped):
            d.text((M + LABEL_W, y + i * LH), ln, font=f_val, fill=PURPLE)
        y += len(wrapped) * LH + 22
        d.line([M, y - 10, W - M, y - 10], fill=(236, 226, 244), width=2)
    d.text((M, y + 24), "Please keep your booking ID", font=f_sub, fill=(90, 60, 120))
    d.text((M, y + 54), "for any future contact.", font=f_sub, fill=(90, 60, 120))
    plogo = _logo_img("logo_prowess.png", 100)
    if plogo:
        img.paste(plogo, (W - M - plogo.width, y + 10))
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def secrets_get(key, default=None):
    try:
        return st.secrets[key]
    except Exception:  # no secrets file / key missing (normal on your own computer)
        return default


CLOUD = secrets_get("gcp_service_account") is not None   # True when deployed online


def load_config():
    if CLOUD:
        return {"sheet_url": secrets_get("sheet_url", ""),
                "drive_folder_id": secrets_get("drive_folder_id", "")}
    if CONFIG_FILE.exists():
        return json.loads(CONFIG_FILE.read_text())
    return {"sheet_url": "", "drive_folder_id": ""}


def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2))


# ----------------------------------------------------------------------------
# Google connection
# ----------------------------------------------------------------------------
@st.cache_resource
def get_creds():
    info = secrets_get("gcp_service_account")
    if info:
        return Credentials.from_service_account_info(dict(info), scopes=SCOPES)
    return Credentials.from_service_account_file(str(CRED_FILE), scopes=SCOPES)


def drive_service():
    """Uses YOUR Google account (OAuth) if configured, otherwise the service account."""
    from googleapiclient.discovery import build
    o = secrets_get("drive_oauth")
    if o:
        creds = UserCredentials(
            token=None, refresh_token=o["refresh_token"], client_id=o["client_id"],
            client_secret=o["client_secret"], token_uri="https://oauth2.googleapis.com/token",
            scopes=["https://www.googleapis.com/auth/drive.file"])
    else:
        creds = get_creds()
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def open_spreadsheet(url):
    gc = gspread.authorize(get_creds())
    return gc.open_by_url(url)


def ensure_event_sheets(sh):
    """Creates the 3 event tabs (with headers) if they do not exist yet."""
    existing = {w.title: w for w in sh.worksheets()}
    sheets = {}
    for ev in EVENTS:
        if ev not in existing:
            ws = sh.add_worksheet(title=ev, rows=1000, cols=len(HEADERS))
            ws.append_row(HEADERS, value_input_option="RAW")
            ws.freeze(rows=1)
            ws.format(f"A1:{LAST_COL}1", {"textFormat": {"bold": True},
                                "backgroundColor": {"red": 0.85, "green": 0.92, "blue": 0.85}})
        else:
            ws = existing[ev]
            first = ws.row_values(1)
            if "Discount %" not in first:
                first = fix_heading_case(ws, first, HEADERS)
            # the discount column was removed: delete it from older tabs (only when every
            # other heading is exactly what we expect, so nothing unknown is ever deleted)
            if "Discount %" in first:
                without = [h for h in first if h != "Discount %"]
                if without == HEADERS[:len(without)]:
                    ws.delete_columns(first.index("Discount %") + 1)
                    first = ws.row_values(1)
            if first != HEADERS and first and first == HEADERS[:len(first)] and len(first) >= 13:
                if ws.col_count < len(HEADERS):             # older tab: add the newer columns at the end
                    ws.add_cols(len(HEADERS) - ws.col_count)
                ws.update(range_name=f"A1:{LAST_COL}1", values=[HEADERS])
            elif first != HEADERS:
                if not any(ws.get_all_values()):          # empty tab -> write headers
                    ws.append_row(HEADERS, value_input_option="RAW")
                else:
                    raise RuntimeError(
                        f"Tab '{ev}' already has data with different column headings. "
                        "Rename or delete that tab so the app can create it correctly.")
        sheets[ev] = ws
    return sheets


@st.cache_resource
def save_lock():
    """One lock shared by everyone using the app, so two people saving at the same
    moment can never be given the same booking ID."""
    return threading.Lock()


def next_booking_id(ws, event):
    """JAC-001, JAC-002 ... : one more than the highest number already in that event's tab."""
    code = EVENT_CODES[event]
    top = 0
    for v in ws.col_values(1)[1:]:
        m = re.fullmatch(rf"{code}-(\d+)", v.strip())
        if m:
            top = max(top, int(m.group(1)))
    return f"{code}-{top + 1:03d}"


@st.cache_data(ttl=15, show_spinner=False)
def read_all(_ws, title):
    return _ws.get_all_values()


SUMMARY_TITLE = "Summary"
EXPENSE_TITLE = "Expenses"


def _norm(row):
    return [" ".join(str(h).split()).casefold() for h in row]


def fix_heading_case(ws, first, canonical):
    """If the headings match except for capital letters / extra spaces (e.g. 'Taken by'),
    rewrite them exactly as the app expects. Returns the (possibly corrected) first row."""
    canon = canonical[:len(first)]
    if first and first != canon and _norm(first) == _norm(canon):
        ws.update(range_name=f"A1:{chr(64 + len(canon))}1", values=[canon])
        return canon
    return first
SCHEMA_VERSION = "2026-10-28-edit-log"   # bump when the sheet layout changes
EXPENSE_HEADERS = ["Expense ID", "Date", "What it was for", "Amount (Rs)", "Taken By", "Receipt Photo"]


def ensure_expenses_sheet(sh):
    """Creates the 'Expenses' tab if it does not exist yet (and adds the Taken By column to an older one)."""
    existing = {w.title: w for w in sh.worksheets()}
    if EXPENSE_TITLE not in existing:
        ws = sh.add_worksheet(title=EXPENSE_TITLE, rows=1000, cols=len(EXPENSE_HEADERS))
        ws.append_row(EXPENSE_HEADERS, value_input_option="RAW")
        ws.freeze(rows=1)
        try:
            ws.format("A1:F1", {"textFormat": {"bold": True},
                                "backgroundColor": {"red": 0.98, "green": 0.9, "blue": 0.8}})
        except Exception as e:  # noqa
            print(f"Expenses formatting skipped: {e}")
        return ws
    ws = existing[EXPENSE_TITLE]
    first = fix_heading_case(ws, ws.row_values(1), EXPENSE_HEADERS)
    if first in (EXPENSE_HEADERS[:4], EXPENSE_HEADERS[:5]):   # older tab: add the newer columns at the end
        if ws.col_count < len(EXPENSE_HEADERS):
            ws.add_cols(len(EXPENSE_HEADERS) - ws.col_count)
        ws.update(range_name="A1:F1", values=[EXPENSE_HEADERS])
    elif first != EXPENSE_HEADERS:
        if not any(ws.get_all_values()):
            ws.append_row(EXPENSE_HEADERS, value_input_option="RAW")
        else:
            raise RuntimeError("The 'Expenses' tab already has data with different column headings. "
                               "Rename or delete it so the app can create it correctly.")
    return ws


LOG_TITLE = "Edit Log"
LOG_HEADERS = ["Time", "Edited By", "Record Type", "ID", "Field", "Old Value", "New Value", "Reason"]


def ensure_log_sheet(sh):
    """Creates the 'Edit Log' tab: one row for every change ever made to a saved record."""
    existing = {w.title: w for w in sh.worksheets()}
    if LOG_TITLE not in existing:
        ws = sh.add_worksheet(title=LOG_TITLE, rows=2000, cols=len(LOG_HEADERS))
        ws.append_row(LOG_HEADERS, value_input_option="RAW")
        ws.freeze(rows=1)
        try:
            ws.format("A1:H1", {"textFormat": {"bold": True},
                                "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.97}})
        except Exception as e:  # noqa
            print(f"Edit Log formatting skipped: {e}")
        return ws
    ws = existing[LOG_TITLE]
    first = fix_heading_case(ws, ws.row_values(1), LOG_HEADERS)
    if first != LOG_HEADERS:
        if not any(ws.get_all_values()):
            ws.append_row(LOG_HEADERS, value_input_option="RAW")
        else:
            raise RuntimeError("The 'Edit Log' tab already has data with different column headings.")
    return ws


def next_expense_id(ws, log_ws=None):
    """EXP-001, EXP-002 ... never re-using a number, even one that was deleted (the Edit Log remembers it)."""
    top = 0
    ids = list(ws.col_values(1)[1:])
    if log_ws is not None:
        ids += log_ws.col_values(4)[1:]
    for v in ids:
        m = re.fullmatch(r"EXP-(\d+)", v.strip())
        if m:
            top = max(top, int(m.group(1)))
    return f"EXP-{top + 1:03d}"


def write_profit_block(ws):
    """Adds the 'Profit' lines to the Summary tab (rows 19-22), only if that spot is empty."""
    try:
        cur = ws.get_values("A19:B22")
        if any(any(str(c).strip() for c in r) for r in cur):
            return
        amount_col = chr(ord("A") + EXPENSE_HEADERS.index("Amount (Rs)"))
        rows = [["Profit", "Amount (Rs)"],
                ["Total collected from stalls", "=C8"],
                ["Total expenses", f"=SUM('{EXPENSE_TITLE}'!{amount_col}2:{amount_col})"],
                ["PROFIT (collected - expenses)", "=B20-B21"]]
        ws.update(range_name="A19", values=rows, value_input_option="USER_ENTERED")
        try:
            green = {"backgroundColor": {"red": 0.85, "green": 0.92, "blue": 0.85}}
            ws.format("A19:B19", {"textFormat": {"bold": True}, **green})
            ws.format("B20:B22", {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}})
            ws.format("A22:B22", {"textFormat": {"bold": True}})
        except Exception as e:  # noqa
            print(f"Profit block formatting skipped: {e}")
    except Exception as e:  # noqa
        print(f"Profit block could not be written: {e}")


def ensure_summary_sheet(sh):
    """Creates a 'Summary' tab (first tab) with live formulas over the three event tabs.
    Created once; if it already exists it is left alone so your own edits are never overwritten."""
    existing = {w.title: w for w in sh.worksheets()}
    if SUMMARY_TITLE in existing:             # already there: only add the new Profit lines if missing
        write_profit_block(existing[SUMMARY_TITLE])
        return
    col = lambda name: chr(ord("A") + HEADERS.index(name))        # column letter of a heading
    paid, contact, method = col("Paid Amount (This Event)"), col("Contact"), col("Payment Method")
    ids = col("Booking ID")
    q = lambda ev: f"'{ev}'"
    rows = [["Stall Booking Summary", "", "", "", ""],
            ["Updates automatically from the three event tabs. Do not type inside the tables.", "", "", "", ""],
            ["", "", "", "", ""],
            ["Event", "Stall bookings", "Payment collected (Rs)", "", ""]]
    for i, ev in enumerate(EVENTS):
        rows.append([ev, f"=COUNTA({q(ev)}!{ids}2:{ids})", f"=SUM({q(ev)}!{paid}2:{paid})", "", ""])
    first, last = 5, 4 + len(EVENTS)
    rows.append(["TOTAL", f"=SUM(B{first}:B{last})", f"=SUM(C{first}:C{last})", "", ""])
    stack = ";".join(f"{q(ev)}!{contact}2:{contact}" for ev in EVENTS)
    rows.append(["", "", "", "", ""])
    rows.append(["Unique stalls (a stall booked for several events counts once)",
                 f'=IFERROR(ROWS(UNIQUE(FILTER({{{stack}}},{{{stack}}}<>""))),0)', "", "", ""])
    unpaid = "+".join(f'COUNTIF({q(ev)}!{method}2:{method},"No payment yet")' for ev in EVENTS)
    rows.append(["Bookings with no payment yet", f"={unpaid}", "", "", ""])
    rows.append(["", "", "", "", ""])
    rows.append(["Payment collected by method (Rs)"] + EVENTS + ["Total"])
    start = len(rows) + 1
    for i, m in enumerate(PAYMENT_METHODS[:-1]):                   # all methods except 'No payment yet'
        r = start + i
        rows.append([m] + [f"=SUMIF({q(ev)}!{method}2:{method},$A{r},{q(ev)}!{paid}2:{paid})"
                           for ev in EVENTS] + [f"=SUM(B{r}:D{r})"])
    ws = sh.add_worksheet(title=SUMMARY_TITLE, rows=40, cols=6, index=0)
    ws.update(range_name="A1", values=rows, value_input_option="USER_ENTERED")
    try:                                   # looks only; the numbers are already in place
        bold = {"textFormat": {"bold": True}}
        ws.format("A1", {"textFormat": {"bold": True, "fontSize": 16}})
        ws.format("A4:C4", {**bold, "backgroundColor": {"red": 0.85, "green": 0.92, "blue": 0.85}})
        ws.format(f"A{last + 1}:C{last + 1}", bold)
        hdr = start - 1
        ws.format(f"A{hdr}:E{hdr}", {**bold, "backgroundColor": {"red": 0.85, "green": 0.92, "blue": 0.85}})
        ws.format("C5:C8", {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}})
        ws.format(f"B{start}:E{start + len(PAYMENT_METHODS) - 2}",
                  {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}})
        ws.columns_auto_resize(0, 4)
    except Exception as e:  # noqa
        print(f"Summary formatting skipped: {e}")
    write_profit_block(ws)


def find_duplicates(sheets, events, brand, contact):
    hits = []
    b = brand.strip().casefold()
    c = "".join(ch for ch in contact if ch.isdigit())
    for ev in events:
        for row in read_all(sheets[ev], ev)[1:]:
            if len(row) <= COL_CONTACT:
                continue
            same_brand = row[COL_BRAND].strip().casefold() == b
            same_phone = c and "".join(ch for ch in row[COL_CONTACT] if ch.isdigit()) == c
            if same_brand or same_phone:
                hits.append(f"{ev}: '{row[COL_BRAND]}' ({row[COL_CONTACT]}) - ID {row[0]}")
    return hits


# ----------------------------------------------------------------------------
# Files (logo + receipt)
# ----------------------------------------------------------------------------
def to_png(uploaded, max_side=1600):
    """Turns an uploaded photo (JPG/PNG/WEBP) into PNG bytes. Fixes phone rotation and
    shrinks very large photos (longest side 1600 px) so the PNG file stays a sensible size."""
    try:
        img = Image.open(io.BytesIO(uploaded.getvalue()))
        img = ImageOps.exif_transpose(img)
        if img.mode not in ("RGB", "RGBA", "L"):
            img = img.convert("RGBA" if ("A" in img.mode or img.mode == "P") else "RGB")
        if max(img.size) > max_side:
            img.thumbnail((max_side, max_side))
        buf = io.BytesIO()
        img.save(buf, "PNG", optimize=True)
        return buf.getvalue()
    except Exception as e:  # noqa
        raise RuntimeError(f"Could not read the picture '{getattr(uploaded, 'name', '')}' ({e}). "
                           "Please upload a JPG, PNG or WEBP photo.")


def save_uploads(booking_id, logo, receipt):
    """Uploads to Google Drive when a folder is configured (always, when deployed online).
    On your own computer it also keeps a local copy. Returns (logo_ref, receipt_ref, warning)."""
    cfg = load_config()
    folder_id = cfg.get("drive_folder_id")
    if CLOUD and not folder_id:
        raise RuntimeError("Drive folder is not configured, so files would be lost. "
                           "Add drive_folder_id to the app secrets.")
    refs, warning = {}, None
    drive = drive_service() if folder_id else None
    local_dir = None
    if not CLOUD:
        local_dir = FILES_DIR / booking_id
        local_dir.mkdir(parents=True, exist_ok=True)

    for label, f in (("logo", logo), ("receipt", receipt)):
        if f is None:
            refs[label] = "No logo" if label == "logo" else "No payment yet"
            continue
        ext = ".png"
        data = to_png(f)                       # logo and receipt are always saved as PNG
        if local_dir:
            (local_dir / f"{label}{ext}").write_bytes(data)
            refs[label] = str(local_dir / f"{label}{ext}")
        if drive:
            try:
                from googleapiclient.http import MediaIoBaseUpload
                media = MediaIoBaseUpload(io.BytesIO(data), mimetype="image/png")
                made = drive.files().create(
                    body={"name": f"{booking_id}_{label}{ext}", "parents": [folder_id]},
                    media_body=media, fields="id,webViewLink",
                    supportsAllDrives=True).execute()
                refs[label] = made["webViewLink"]
            except Exception as e:  # noqa
                if CLOUD:
                    raise RuntimeError(f"Google Drive upload failed: {str(e)[:200]}")
                warning = (f"Drive upload failed ({str(e)[:120]}). "
                           "Files are safe in the local 'booking_files' folder.")
    return refs["logo"], refs["receipt"], warning


# ----------------------------------------------------------------------------
# Form state helpers
# ----------------------------------------------------------------------------
def defaults():
    return {"f_name": "", "f_brand": "", "f_items": "", "f_events": [],
            "f_paid": 0.0,
            "f_date": datetime.now(TZ).date(), "f_contact": "",
            "f_method": None, "f_received": "",
            "parse_errors": [], "parse_warnings": [], "parsed": False}


for k, v in defaults().items():
    st.session_state.setdefault(k, v)
st.session_state.setdefault("raw_text", "")
st.session_state.setdefault("uploader_n", 0)
st.session_state.setdefault("last_saved", None)


def do_parse():
    data, errs, warns = parse_text(st.session_state.raw_text)
    st.session_state.f_name = data["name"]
    st.session_state.f_brand = data["brand"]
    st.session_state.f_items = data["items"]
    st.session_state.f_events = data["events"]
    st.session_state.f_paid = float(data["paid"] or 0)
    if data["date"]:
        st.session_state.f_date = data["date"]
    st.session_state.f_contact = data["contact"]
    st.session_state.f_method = data["method"]
    st.session_state.f_received = data["received_by"]
    st.session_state.parse_errors = errs
    st.session_state.parse_warnings = warns
    st.session_state.parsed = True
    st.session_state.last_saved = None


def reset_form():
    for k, v in defaults().items():
        st.session_state[k] = v
    st.session_state.raw_text = ""
    st.session_state.uploader_n += 1


if st.session_state.pop("do_reset_now", False):
    reset_form()

# ----------------------------------------------------------------------------
# Password gate (used when the app is online)
# ----------------------------------------------------------------------------
_pw = secrets_get("app_password")
if _pw and not st.session_state.get("authed"):
    st.subheader("🔒 Enter the password")
    entered = st.text_input("Password", type="password")
    if entered:
        if hmac.compare_digest(entered.encode(), str(_pw).encode()):
            st.session_state.authed = True
            st.rerun()
        else:
            st.error("Wrong password.")
    st.stop()

# ----------------------------------------------------------------------------
# Sidebar: connection
# ----------------------------------------------------------------------------
cfg = load_config()
with st.sidebar:
    st.header("⚙️ Connection")
    if CLOUD:
        email = dict(secrets_get("gcp_service_account")).get("client_email", "")
        st.caption("Settings are managed in the app secrets.")
    else:
        if not CRED_FILE.exists():
            st.error("credentials.json not found. Put the Google key file in the app folder "
                     "(see setup guide).")
            st.stop()
        email = json.loads(CRED_FILE.read_text()).get("client_email", "")
        st.caption("Share your Google Sheet (as **Editor**) with:")
        st.code(email, language=None)
        url = st.text_input("Google Sheet URL", value=cfg.get("sheet_url", ""))
        folder_id = st.text_input("Drive folder ID for logos/receipts (optional)",
                                  value=cfg.get("drive_folder_id", ""))
        if st.button("Save & connect"):
            save_config({"sheet_url": url.strip(), "drive_folder_id": folder_id.strip()})
            st.rerun()

if not cfg.get("sheet_url"):
    st.info("Paste your Google Sheet URL in the sidebar and press **Save & connect**.")
    st.stop()

@st.cache_resource
def connect(url, schema_version):
    sh = open_spreadsheet(url)
    sheets = ensure_event_sheets(sh)
    try:
        expenses_ws = ensure_expenses_sheet(sh)
    except Exception as e:  # noqa
        expenses_ws = None
        print(f"Expenses sheet problem: {e}")
    try:
        log_ws = ensure_log_sheet(sh)
    except Exception as e:  # noqa
        log_ws = None
        print(f"Edit Log problem: {e}")
    try:
        ensure_summary_sheet(sh)
    except Exception as e:  # noqa - the summary is a bonus; never stop bookings because of it
        print(f"Summary sheet could not be created: {e}")
    return sh, sheets, expenses_ws, log_ws


try:
    sh, sheets, expenses_ws, log_ws = connect(cfg["sheet_url"], SCHEMA_VERSION)
except Exception as e:  # noqa
    st.error(f"Could not open the Google Sheet: {e}")
    st.stop()

st.sidebar.success(f"Connected: {sh.title}")

# ----------------------------------------------------------------------------
# Main UI
# ----------------------------------------------------------------------------
dash = st.container()          # filled in at the very end, once all the helper functions exist
tab_new, tab_view, tab_search, tab_exp, tab_edit = st.tabs(
    ["➕ New booking", "📊 Records", "🔎 Search", "💸 Expenses", "✏️ Edit"])

with tab_new:
    st.subheader("Step 1 - Paste the booking details")
    st.text_area("Booking text", key="raw_text", height=230,
                 placeholder="Name:\nBrand Name:\nBrand Items:\nBooked for: Jacranda, AMBLEM, qawali night\n"
                             "Paid Amount:\nBooking date:\nContact:")
    st.button("🔍 Read details", on_click=do_parse, type="primary",
              disabled=not st.session_state.raw_text.strip())

    if st.session_state.last_saved:
        st.success(st.session_state.last_saved["msg"])
        st.caption("Receipt for the stall owner - download the picture, or copy the text below:")
        st.image(st.session_state.last_saved["png"], width=420)
        st.download_button("⬇️ Download receipt (PNG)", data=st.session_state.last_saved["png"],
                           file_name=f"receipt_{st.session_state.last_saved['file_tag']}.png",
                           mime="image/png")
        st.code(st.session_state.last_saved["copy"], language=None)

    if st.session_state.parsed:
        st.subheader("Step 2 - Check what the app understood")
        for e in st.session_state.parse_errors:
            st.error(e)
        for w in st.session_state.parse_warnings:
            st.warning(w)
        if not st.session_state.parse_errors and not st.session_state.parse_warnings:
            st.success("Everything was read cleanly. Quick check below, then add files.")

        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Name", key="f_name")
            st.text_input("Brand Name", key="f_brand")
            st.text_area("Brand Items", key="f_items", height=100)
            st.multiselect("Booked for", EVENTS, key="f_events")
        with c2:
            st.number_input("Paid Amount (PKR, whole booking)", min_value=0.0,
                            step=500.0, key="f_paid")
            st.date_input("Booking date", key="f_date", format="DD/MM/YYYY")
            st.text_input("Contact", key="f_contact")

        events = st.session_state.f_events
        paid = float(st.session_state.f_paid)

        # --- split of the paid amount across events (prevents double counting) ---
        amounts = {}
        if len(events) == 1:
            amounts[events[0]] = paid
        elif len(events) > 1:
            mode = st.radio("Paid amount is split across events…",
                            ["Equally", "Manually"], horizontal=True)
            eq = split_amount(paid, len(events))
            cols = st.columns(len(events))
            for i, ev in enumerate(events):
                with cols[i]:
                    if mode == "Equally":
                        amounts[ev] = float(eq[i])
                        st.metric(ev, f"Rs {eq[i]:,}")
                    else:
                        amounts[ev] = st.number_input(
                            f"{ev} (Rs)", min_value=0.0, step=500.0, value=float(eq[i]),
                            key=f"split_{ev}_{paid}_{len(events)}")
            total_split = round(sum(amounts.values()), 2)
            if abs(total_split - paid) > 0.01:
                st.error(f"Split adds up to Rs {total_split:,} but paid amount is Rs {paid:,}.")

        st.subheader("Step 3 - Payment details, logo and receipt")
        method = st.radio("How was the payment made?", PAYMENT_METHODS, index=None,
                          key="f_method", horizontal=True)
        if method and method != "No payment yet":
            st.text_input("Received by (team member who received or confirmed this payment)",
                          key="f_received")
        if method == "Cash":
            st.caption("For cash, the receipt photo is optional (e.g. a photo of the handwritten "
                       "receipt or the cash handover).")
        u1, u2 = st.columns(2)
        with u1:
            logo = st.file_uploader("Brand logo (optional) - saved as PNG",
                                    type=["png", "jpg", "jpeg", "webp"],
                                    key=f"logo_{st.session_state.uploader_n}")
            if logo:
                st.image(logo, width=160)
        with u2:
            receipt = st.file_uploader(
                "Cash receipt / photo (optional) - saved as PNG" if method == "Cash"
                else "No payment yet - nothing to upload" if method == "No payment yet"
                else "Payment receipt - saved as PNG", type=["png", "jpg", "jpeg", "webp"],
                                       key=f"receipt_{st.session_state.uploader_n}")
            if receipt:
                st.image(receipt, width=160)
        st.caption("Only photos are accepted (JPG, PNG, WEBP) and they are saved as PNG. "
                   "If you have a PDF receipt, take a screenshot of it and upload that.")

        # ------------------------- validation -------------------------------
        problems = []
        for label, key in (("Name", "f_name"), ("Brand Name", "f_brand"),
                           ("Brand Items", "f_items"), ("Contact", "f_contact")):
            if not str(st.session_state[key]).strip():
                problems.append(f"{label} is empty.")
        if not events:
            problems.append("Select at least one event.")
        if method is None:
            problems.append("Choose how the payment was made.")
        elif method == "No payment yet":
            if paid != 0:
                problems.append("Method is 'No payment yet' but a paid amount is entered - set it to 0 "
                                "or choose the real payment method.")
        else:
            if paid <= 0:
                problems.append("Paid amount is 0 - enter the amount, or choose 'No payment yet'.")
            if not st.session_state.f_received.strip():
                problems.append("Enter who received the payment.")
            if method != "Cash" and receipt is None:
                problems.append("Payment receipt not uploaded.")
        if len(events) > 1 and abs(sum(amounts.values()) - paid) > 0.01:
            problems.append("Split amounts do not add up to the paid amount.")

        if CLOUD and not cfg.get("drive_folder_id"):
            problems.append("Drive folder is not configured in the app secrets (files would be lost).")

        dupes = []
        if events and st.session_state.f_brand.strip():
            dupes = find_duplicates(sheets, events, st.session_state.f_brand, st.session_state.f_contact)
        allow_dupe = True
        if dupes:
            st.warning("Possible duplicate booking found:\n\n" + "\n".join(f"- {d}" for d in dupes))
            allow_dupe = st.checkbox("This is NOT a duplicate - save anyway")

        st.subheader("Step 4 - Save")
        if problems:
            st.info("Fix these first:\n\n" + "\n".join(f"- {p}" for p in problems))
        if st.button("✅ Save booking to Google Sheets", type="primary",
                     disabled=bool(problems) or not allow_dupe):
            now = datetime.now(TZ).strftime("%d-%b-%Y %H:%M")
            brand_clean = st.session_state.f_brand.strip()
            with st.spinner("Saving..."):
                with save_lock():          # only one booking is saved at a time
                    try:
                        ids = {ev: next_booking_id(sheets[ev], ev) for ev in events}
                    except Exception as e:  # noqa
                        st.error(f"Nothing was saved. Could not read the sheet: {e}")
                        st.stop()
                    tag = "_".join(ids.values())
                    try:
                        logo_ref, receipt_ref, drive_warn = save_uploads(tag, logo, receipt)
                        if receipt is None and method == "Cash":
                            receipt_ref = "Cash - no receipt photo"
                    except Exception as e:  # noqa
                        st.error(f"Nothing was saved. {e}")
                        st.stop()
                    saved_to = []
                    try:
                        for ev in events:
                            others = ", ".join(f"{ids[e]} ({e})" for e in events if e != ev)
                            row = [ids[ev], now, st.session_state.f_name.strip(),
                                   brand_clean, st.session_state.f_items.strip(),
                                   ", ".join(events), ev, amounts[ev], paid,
                                   st.session_state.f_date.strftime("%d-%b-%Y"),
                                   st.session_state.f_contact.strip(),
                                   logo_ref, receipt_ref, method,
                                   st.session_state.f_received.strip() if method != "No payment yet" else "",
                                   others]
                            sheets[ev].append_row(row, value_input_option="RAW")
                            saved_to.append(ev)
                    except Exception as e:  # noqa
                        st.error(f"Saved to {', '.join(saved_to) or 'nothing'} but then failed on the next "
                                 f"sheet: {e}. IDs used: {', '.join(ids.values())}. Add the missing row "
                                 "manually or delete the partial row and save again.")
                        st.stop()
            ids_text = ", ".join(f"{ids[e]} ({e})" for e in events)
            msg = f"Saved {brand_clean}. Booking ID(s): {ids_text}"
            rows_for_receipt = receipt_rows(
                brand_clean, st.session_state.f_name.strip(), events, ids, amounts, paid, method,
                st.session_state.f_received.strip() if method != "No payment yet" else "",
                st.session_state.f_date.strftime("%d-%b-%Y"))
            st.cache_data.clear()
            st.session_state.last_saved = {"msg": msg, "copy": receipt_text(rows_for_receipt),
                                           "png": make_receipt_png(rows_for_receipt),
                                           "file_tag": "_".join(ids.values())}
            st.session_state.do_reset_now = True
            st.rerun()

with tab_view:
    st.subheader("Saved bookings")
    if st.button("🔄 Refresh"):
        st.cache_data.clear()
        st.rerun()
    ev_tabs = st.tabs(EVENTS)
    for ev, t in zip(EVENTS, ev_tabs):
        with t:
            vals = read_all(sheets[ev], ev)
            if len(vals) <= 1:
                st.caption("No bookings yet.")
                continue
            df = pd.DataFrame([r + [""] * (len(vals[0]) - len(r)) for r in vals[1:]], columns=vals[0])
            paid_col = pd.to_numeric(df["Paid Amount (This Event)"], errors="coerce").fillna(0)
            m1, m2 = st.columns(2)
            m1.metric("Stalls booked", len(df))
            m2.metric("Total collected (Rs)", f"{paid_col.sum():,.0f}")
            st.dataframe(df, hide_index=True)


# ----------------------------------------------------------------------------
# Search tab
# ----------------------------------------------------------------------------
def digits_of(text):
    return re.sub(r"\D", "", text or "")


def load_all_rows():
    rows = []
    for ev in EVENTS:
        vals = read_all(sheets[ev], ev)
        if len(vals) <= 1:
            continue
        head = vals[0]
        for r in vals[1:]:
            r = r + [""] * (len(head) - len(r))
            d = dict(zip(head, r))
            d["_event"] = ev
            rows.append(d)
    return rows


def stall_key(d):
    return digits_of(d.get("Contact", "")) or d.get("Brand Name", "").strip().casefold()


def row_matches(d, q):
    ql = q.strip().casefold()
    short_num = q.strip().isdigit() and len(q.strip()) <= 4      # e.g. "14"
    if not short_num and ql in d.get("Booking ID", "").casefold():
        return True
    if ql in d.get("Brand Name", "").casefold() or ql in d.get("Name", "").casefold():
        return True
    qd = digits_of(q)
    if len(qd) >= 4 and qd in digits_of(d.get("Contact", "")):
        return True
    if short_num:                                           # "14" finds JAC-014, AMB-014 ...
        m = re.search(r"-(\d+)$", d.get("Booking ID", ""))
        if m and int(m.group(1)) == int(q.strip()):
            return True
    return False


with tab_search:
    st.subheader("Find a stall")
    q = st.text_input("Search by booking ID, brand name, person name or phone number",
                      key="search_q", placeholder="e.g. JAC-014, zara, 0300-1234567")
    if st.button("🔄 Refresh data", key="refresh_search"):
        st.cache_data.clear()
        st.rerun()
    if not q.strip():
        st.caption("Type part of a brand or person name, a phone number, or an ID such as JAC-014 "
                   "(or just 14). All three events are searched together.")
    else:
        all_rows = load_all_rows()
        groups = {}
        for d in all_rows:
            groups.setdefault(stall_key(d), []).append(d)
        hits = [g for g in groups.values() if any(row_matches(d, q) for d in g)]
        if not hits:
            st.info("No booking found. Check the spelling, or try just part of the name or the last digits "
                    "of the phone number.")
        else:
            st.caption(f"{len(hits)} stall(s) found" + (" - showing the first 20" if len(hits) > 20 else ""))
        for g in hits[:20]:
            first = g[0]
            st.markdown(f"### {first.get('Brand Name', '')}")
            st.caption(f"{first.get('Name', '')}  ·  {first.get('Contact', '')}  ·  "
                       f"Items: {first.get('Brand Items', '')}")
            cols = {"Booking ID": "Booking ID", "_event": "Event",
                    "Paid Amount (This Event)": "Paid (Rs)",
                    "Booking Date": "Booking date", "Payment Method": "Payment method",
                    "Received By": "Received by"}
            df = pd.DataFrame([{new: d.get(old, "") for old, new in cols.items()} for d in g])
            st.dataframe(df, hide_index=True)
            def _num(x):
                try:
                    return float(str(x).replace(",", ""))
                except ValueError:
                    return 0.0
            total = sum(_num(d.get("Paid Amount (This Event)", "")) for d in g)
            st.metric("Total paid (all events)", f"Rs {total:,.0f}")
            for label in ("Logo", "Payment Receipt"):
                seen = []
                for d in g:
                    ref = d.get(label, "").strip()
                    if ref and ref not in seen:
                        seen.append(ref)
                for ref in seen:
                    if ref.startswith("http"):
                        st.link_button(f"Open {label.lower()}", ref)
                    else:
                        st.caption(f"{label}: {ref}")
            st.divider()


# ----------------------------------------------------------------------------
# Expenses tab
# ----------------------------------------------------------------------------
def _amount(x):
    try:
        return float(str(x).replace(",", ""))
    except ValueError:
        return 0.0


def total_collected():
    total = 0.0
    for ev in EVENTS:
        vals = read_all(sheets[ev], ev)
        if len(vals) > 1 and "Paid Amount (This Event)" in vals[0]:
            i = vals[0].index("Paid Amount (This Event)")
            total += sum(_amount(r[i]) for r in vals[1:] if len(r) > i)
    return total


def rs(x):
    return f"Rs {x:,.0f}" if float(x).is_integer() else f"Rs {x:,.2f}"


with tab_exp:
    st.subheader("Add an expense")
    if expenses_ws is None:
        st.error("The 'Expenses' tab could not be prepared in your Google Sheet. If a tab called "
                 "'Expenses' already exists with other headings, rename it and restart the app.")
    else:
        _vals = read_all(expenses_ws, EXPENSE_TITLE)
        known_names = sorted({" ".join(r[4].split()) for r in _vals[1:] if len(r) > 4 and r[4].strip()},
                             key=str.casefold)
        with st.form("expense_form", clear_on_submit=True):
            e_date = st.date_input("Date", value=datetime.now(TZ).date(), format="DD/MM/YYYY")
            e_what = st.text_input("What was it for?", placeholder="e.g. Stage banner printing")
            e_amount = st.number_input("Amount (Rs)", min_value=0.0, step=100.0)
            e_taken = st.text_input("Money taken by (name of the person who took the money)")
            if known_names:
                st.caption("Names used so far: " + ", ".join(known_names) +
                           ". Type the name the same way each time so the totals add up.")
            e_receipt = st.file_uploader("Receipt photo as proof (optional) - saved as PNG",
                                         type=["png", "jpg", "jpeg", "webp"], key="exp_receipt")
            e_dupe_ok = st.checkbox("Save even if an identical expense is already recorded")
            e_submit = st.form_submit_button("✅ Save expense", type="primary")

        if e_submit:
            what = " ".join(e_what.split())
            taken_by = " ".join(e_taken.split())
            st.session_state.exp_msg = None
            amount = int(e_amount) if float(e_amount).is_integer() else round(float(e_amount), 2)
            date_txt = e_date.strftime("%d-%b-%Y")
            if not what:
                st.error("Please write what the expense was for. Nothing was saved.")
            elif amount <= 0:
                st.error("The amount must be more than 0. Nothing was saved.")
            elif not taken_by:
                st.error("Please write who took the money. Nothing was saved.")
            else:
                try:
                    with save_lock():
                        existing_rows = expenses_ws.get_all_values()[1:]
                        same = [r for r in existing_rows if len(r) >= 4 and r[1] == date_txt
                                and r[2].strip().casefold() == what.casefold()
                                and _amount(r[3]) == float(amount)]
                        if same and not e_dupe_ok:
                            st.warning(f"An identical expense is already recorded ({same[0][0]}). "
                                       "Nothing was saved. Tick the box if it is really a second one.")
                        else:
                            exp_id = next_expense_id(expenses_ws, log_ws)
                            receipt_ref, photo_warn = "No receipt photo", None
                            if e_receipt is not None:
                                _, receipt_ref, photo_warn = save_uploads(exp_id, None, e_receipt)
                            expenses_ws.append_row([exp_id, date_txt, what, amount, taken_by, receipt_ref],
                                                   value_input_option="RAW")
                            if photo_warn:
                                st.warning(photo_warn)
                            st.cache_data.clear()
                            st.session_state.exp_msg = f"Saved {exp_id}: {what} - {rs(amount)} - taken by {taken_by} ({date_txt})" + ("" if e_receipt is None else " - receipt photo saved")
                except Exception as e:  # noqa
                    st.error(f"Nothing was saved: {e}")
        if st.session_state.get("exp_msg") and not e_submit:
            st.success(st.session_state.exp_msg)
        elif e_submit and st.session_state.get("exp_msg"):
            st.success(st.session_state.exp_msg)

        st.divider()
        if st.button("🔄 Refresh", key="refresh_exp"):
            st.cache_data.clear()
            st.rerun()
        vals = read_all(expenses_ws, EXPENSE_TITLE)
        data = [r + [""] * (len(EXPENSE_HEADERS) - len(r)) for r in vals[1:]]
        if data:
            by_person = {}
            for r in data:
                name = " ".join(r[4].split()) or "(not recorded)"
                by_person.setdefault(name.casefold(), [name, 0.0])[1] += _amount(r[3])
            st.markdown("**Money taken, by person**")
            st.dataframe(pd.DataFrame([{"Taken by": n, "Total taken": rs(a)}
                                       for n, a in sorted(by_person.values(), key=lambda x: -x[1])]),
                         hide_index=True)
            st.markdown("**All expenses (newest first)**")
            df = pd.DataFrame(data, columns=EXPENSE_HEADERS).iloc[::-1]
            df["Receipt Photo"] = [x if x.startswith("http") else None for x in df["Receipt Photo"]]
            st.dataframe(df, hide_index=True,
                         column_config={"Receipt Photo": st.column_config.LinkColumn(
                             "Receipt photo", display_text="Open")})
            missing = sum(1 for r in data if not r[5].strip() or r[5].strip() == "No receipt photo")
            st.caption(f"{missing} of {len(data)} expenses have no receipt photo.")
        else:
            st.caption("No expenses recorded yet.")


# ----------------------------------------------------------------------------
# Editing saved records (bookings and expenses)
# ----------------------------------------------------------------------------
ID_RE = re.compile(r"[A-Z]{3}-\d+")
NO_REF = ("", "No payment yet", "Cash - no receipt photo", "No receipt photo", "No logo")
COL = {h: i for i, h in enumerate(HEADERS)}                     # heading -> column position
SHARED_FIELDS = ["Name", "Brand Name", "Brand Items", "Booking Date", "Contact", "Logo",
                 "Payment Receipt", "Payment Method", "Received By"]


def has_ref(ref):
    return str(ref).strip() not in NO_REF


def clean_num(x):
    x = float(x)
    return int(x) if x.is_integer() else round(x, 2)


def now_text():
    return datetime.now(TZ).strftime("%d-%b-%Y %H:%M:%S")


def fresh_row(ws, rid, width):
    """Reads the sheet again right now and returns (row number, row values) for an ID."""
    for i, r in enumerate(ws.get_all_values(), start=1):
        if r and r[0] == rid:
            return i, ([str(c) for c in r] + [""] * width)[:width]
    return None, None


def booking_groups():
    """A booking = all the event rows linked to each other (JAC-001 + QWL-001 ...)."""
    rows = load_all_rows()
    by_id = {r.get("Booking ID", ""): r for r in rows if r.get("Booking ID")}
    groups, seen = [], set()
    for r in rows:
        rid = r.get("Booking ID", "")
        if not rid or rid in seen:
            continue
        linked = [rid] + [i for i in ID_RE.findall(r.get("Other IDs of This Booking", ""))
                          if i in by_id and i != rid]
        seen.update(linked)
        groups.append([by_id[i] for i in linked])

    def newest(g):
        try:
            return dateparser.parse(g[0].get("Saved At", ""), dayfirst=True)
        except Exception:  # noqa
            return datetime.min
    return sorted(groups, key=newest, reverse=True)


def group_label(g):
    return f"{', '.join(d['Booking ID'] for d in g)} - {g[0].get('Brand Name', '')} ({g[0].get('Name', '')})"


def render_booking_editor(g):
    ids = [d["Booking ID"] for d in g]
    snap = {d["Booking ID"]: [str(d.get(h, "")) for h in HEADERS] for d in g}
    sig = hashlib.md5(json.dumps(snap, sort_keys=True).encode()).hexdigest()[:8]
    first = snap[ids[0]]
    st.caption("Events in this booking: " + ", ".join(f"{d['Booking ID']} ({d['_event']})" for d in g))
    for label, col in (("logo", "Logo"), ("receipt", "Payment Receipt")):
        ref = first[COL[col]]
        if ref.startswith("http"):
            st.link_button(f"Open current {label}", ref)
        else:
            st.caption(f"Current {label}: {ref or '-'}")
    try:
        cur_date = parse_date(first[COL["Booking Date"]])
    except ValueError:
        cur_date = datetime.now(TZ).date()
    cur_method = first[COL["Payment Method"]]

    with st.form(f"edit_booking_{sig}"):
        name = st.text_input("Name", value=first[COL["Name"]])
        brand = st.text_input("Brand Name", value=first[COL["Brand Name"]])
        items = st.text_area("Brand Items", value=first[COL["Brand Items"]])
        contact = st.text_input("Contact", value=first[COL["Contact"]])
        bdate = st.date_input("Booking date", value=cur_date, format="DD/MM/YYYY")
        method = st.radio("Payment method", PAYMENT_METHODS, horizontal=True,
                          index=PAYMENT_METHODS.index(cur_method) if cur_method in PAYMENT_METHODS else None)
        received = st.text_input("Received by", value=first[COL["Received By"]])
        st.markdown("**Paid amount for each event (Rs)**")
        amounts = {}
        for c, d in zip(st.columns(len(g)), g):
            with c:
                rid = d["Booking ID"]
                amounts[rid] = st.number_input(f"{d['_event']} ({rid})", min_value=0.0, step=500.0,
                                               value=float(_amount(snap[rid][COL["Paid Amount (This Event)"]])))
        new_logo = st.file_uploader("Add or replace the logo (optional) - saved as PNG",
                                    type=["png", "jpg", "jpeg", "webp"])
        new_receipt = st.file_uploader("Add or replace the payment receipt (optional) - saved as PNG",
                                       type=["png", "jpg", "jpeg", "webp"])
        edited_by = st.text_input("Edited by (your name)")
        reason = st.text_input("Reason for the change (optional)")
        submitted = st.form_submit_button("💾 Save changes", type="primary")
    if not submitted:
        return

    # ---------------- checks ----------------
    name, brand, items, received = (" ".join(x.split()) for x in (name, brand, items, received))
    total = round(sum(amounts.values()), 2)
    problems = []
    for lab, val in (("Name", name), ("Brand Name", brand), ("Brand Items", items),
                     ("Contact", contact.strip()), ("Edited by", " ".join(edited_by.split()))):
        if not val:
            problems.append(f"{lab} cannot be empty.")
    if method is None:
        problems.append("Choose the payment method.")
    elif method == "No payment yet":
        if total != 0:
            problems.append("Payment method is 'No payment yet' but amounts are entered - set them to 0 "
                            "or choose the real payment method.")
    else:
        if total <= 0:
            problems.append("The paid amounts are all 0 - enter them, or choose 'No payment yet'.")
        if not received:
            problems.append("Enter who received the payment.")
        if method != "Cash" and new_receipt is None and not has_ref(first[COL["Payment Receipt"]]):
            problems.append("This payment method needs a receipt photo - please add one.")
    if problems:
        st.error("Nothing was saved:\n\n" + "\n".join(f"- {p}" for p in problems))
        return
    contact_clean, contact_warn = normalize_contact(contact)
    if contact_warn:
        st.warning(contact_warn)

    # ---------------- build the new rows ----------------
    new_logo_ref, new_receipt_ref = first[COL["Logo"]], first[COL["Payment Receipt"]]
    newrows = {}
    for rid in ids:
        row = list(snap[rid])
        row[COL["Name"]], row[COL["Brand Name"]], row[COL["Brand Items"]] = name, brand, items
        row[COL["Booking Date"]] = bdate.strftime("%d-%b-%Y")
        row[COL["Contact"]] = contact_clean
        row[COL["Payment Method"]] = method
        row[COL["Received By"]] = "" if method == "No payment yet" else received
        newrows[rid] = row

    with save_lock():
        # 1) has someone else changed these rows since they were loaded?
        locs = {}
        for d in g:
            rid = d["Booking ID"]
            rownum, cur = fresh_row(sheets[d["_event"]], rid, len(HEADERS))
            if rownum is None or cur != [str(x) for x in snap[rid]][:len(HEADERS)]:
                st.error("This booking was changed by someone else (or removed) while you were editing. "
                         "Nothing was saved. Choose it again to see the latest version.")
                return
            locs[rid] = rownum
        # 2) new photos (before anything is written)
        if new_logo is not None or new_receipt is not None:
            tag = "_".join(ids) + "_edit_" + datetime.now(TZ).strftime("%H%M%S")
            try:
                lref, rref, warn = save_uploads(tag, new_logo, new_receipt)
            except Exception as e:  # noqa
                st.error(f"Nothing was saved. {e}")
                return
            if new_logo is not None:
                new_logo_ref = lref
            if new_receipt is not None:
                new_receipt_ref = rref
            if warn:
                st.warning(warn)
        for rid in ids:
            newrows[rid][COL["Logo"]], newrows[rid][COL["Payment Receipt"]] = new_logo_ref, new_receipt_ref
            newrows[rid][COL["Paid Amount (This Event)"]] = clean_num(amounts[rid])
            newrows[rid][COL["Total Paid (Whole Booking)"]] = clean_num(total)
        # 3) what exactly changed (for the Edit Log)
        old0, new0 = snap[ids[0]], newrows[ids[0]]
        who, joined = " ".join(edited_by.split()), ", ".join(ids)
        changes = [(joined, f, old0[COL[f]], str(new0[COL[f]])) for f in SHARED_FIELDS
                   if old0[COL[f]] != str(new0[COL[f]])]
        for rid in ids:
            j = COL["Paid Amount (This Event)"]
            if float(_amount(snap[rid][j])) != float(newrows[rid][j]):
                changes.append((rid, "Paid Amount (This Event)", snap[rid][j], str(newrows[rid][j])))
        if not changes:
            st.info("Nothing was changed, so nothing was saved.")
            return
        # 4) write
        done = []
        try:
            for d in g:
                rid, r = d["Booking ID"], locs[d["Booking ID"]]
                sheets[d["_event"]].update(range_name=f"A{r}:{LAST_COL}{r}", values=[newrows[rid]],
                                           value_input_option="RAW")
                done.append(rid)
        except Exception as e:  # noqa
            st.error(f"Only some rows were updated ({', '.join(done) or 'none'}) before an error: {e}. "
                     "Open the booking again and check it.")
            return
        try:
            stamp = now_text()
            log_ws.append_rows([[stamp, who, "Booking", i, f, o, n, reason.strip()]
                                for i, f, o, n in changes], value_input_option="RAW")
        except Exception as e:  # noqa
            st.warning(f"The changes were saved but could not be written to the Edit Log: {e}")
    st.cache_data.clear()
    st.session_state.edit_msg = f"Saved the changes to {', '.join(ids)} ({len(changes)} change(s) logged)."
    st.session_state.reset_edit_choice = True
    st.rerun()


def render_expense_editor():
    st.divider()
    st.subheader("Edit or delete an expense")
    if log_ws is None:
        st.error("The 'Edit Log' tab could not be prepared, so editing is switched off.")
        return
    msg = st.session_state.pop("exp_edit_msg", None)
    if msg:
        st.success(msg)
    vals = read_all(expenses_ws, EXPENSE_TITLE)
    width = len(EXPENSE_HEADERS)
    rows = [([str(c) for c in r] + [""] * width)[:width] for r in vals[1:] if r and r[0]]
    if not rows:
        st.caption("No expenses to edit yet.")
        return
    labels = [f"{r[0]} - {r[2]} - Rs {_amount(r[3]):,.0f} - {r[4] or 'no name'}" for r in reversed(rows)]
    by_label = dict(zip(labels, reversed(rows)))
    if st.session_state.pop("reset_exp_choice", False):
        st.session_state["exp_choice"] = None
    choice = st.selectbox("Choose the expense (type to search)", labels, index=None, key="exp_choice")
    if choice is None:
        return
    snap = by_label[choice]
    rid = snap[0]
    sig = hashlib.md5(json.dumps(snap).encode()).hexdigest()[:8]
    ref = snap[5]
    if ref.startswith("http"):
        st.link_button("Open current receipt photo", ref)
    else:
        st.caption(f"Current receipt photo: {ref or '-'}")
    try:
        cur_date = datetime.strptime(snap[1], "%d-%b-%Y").date()
    except ValueError:
        cur_date = datetime.now(TZ).date()
    with st.form(f"edit_expense_{sig}"):
        e_date = st.date_input("Date", value=cur_date, format="DD/MM/YYYY")
        e_what = st.text_input("What was it for?", value=snap[2])
        e_amount = st.number_input("Amount (Rs)", min_value=0.0, step=100.0, value=float(_amount(snap[3])))
        e_taken = st.text_input("Money taken by", value=snap[4])
        e_photo = st.file_uploader("Add or replace the receipt photo (optional) - saved as PNG",
                                   type=["png", "jpg", "jpeg", "webp"])
        e_by = st.text_input("Edited by (your name)")
        e_reason = st.text_input("Reason (optional)")
        c_save, c_del = st.columns(2)
        with c_save:
            do_save = st.form_submit_button("💾 Save changes", type="primary")
        with c_del:
            do_delete = st.form_submit_button("🗑 Delete this expense")
        confirm = st.checkbox("To delete, tick here to confirm (the deleted entry stays in the Edit Log)")
    if not (do_save or do_delete):
        return
    who = " ".join(e_by.split())
    if not who:
        st.error("Please write your name in 'Edited by'. Nothing was changed.")
        return
    if do_delete and not confirm:
        st.error("Tick the confirmation box to delete. Nothing was changed.")
        return
    what, taken = " ".join(e_what.split()), " ".join(e_taken.split())
    if do_save and (not what or not taken or e_amount <= 0):
        st.error("What it was for, the amount (more than 0) and who took the money are all required. "
                 "Nothing was saved.")
        return

    with save_lock():
        rownum, cur = fresh_row(expenses_ws, rid, width)
        if rownum is None or cur != snap:
            st.error("This expense was changed by someone else (or removed) while you were editing. "
                     "Nothing was saved. Choose it again to see the latest version.")
            return
        stamp = now_text()
        if do_delete:
            try:
                expenses_ws.delete_rows(rownum)
                log_ws.append_rows([[stamp, who, "Expense", rid, f, snap[i], "(deleted)", e_reason.strip()]
                                    for i, f in enumerate(EXPENSE_HEADERS)], value_input_option="RAW")
            except Exception as e:  # noqa
                st.error(f"The delete did not finish: {e}. Please check the Expenses tab.")
                return
            st.cache_data.clear()
            st.session_state.exp_edit_msg = f"Deleted {rid}. It is recorded in the Edit Log."
        else:
            new_ref = snap[5]
            if e_photo is not None:
                try:
                    _, new_ref, warn = save_uploads(f"{rid}_edit_{datetime.now(TZ):%H%M%S}", None, e_photo)
                except Exception as e:  # noqa
                    st.error(f"Nothing was saved. {e}")
                    return
                if warn:
                    st.warning(warn)
            new = [rid, e_date.strftime("%d-%b-%Y"), what, clean_num(e_amount), taken, new_ref]
            changes = [(f, snap[i], str(new[i])) for i, f in enumerate(EXPENSE_HEADERS)
                       if i and snap[i] != str(new[i])]
            if not changes:
                st.info("Nothing was changed, so nothing was saved.")
                return
            try:
                expenses_ws.update(range_name=f"A{rownum}:F{rownum}", values=[new], value_input_option="RAW")
            except Exception as e:  # noqa
                st.error(f"Nothing was saved: {e}")
                return
            try:
                log_ws.append_rows([[stamp, who, "Expense", rid, f, o, n, e_reason.strip()]
                                    for f, o, n in changes], value_input_option="RAW")
            except Exception as e:  # noqa
                st.warning(f"The change was saved but could not be written to the Edit Log: {e}")
            st.cache_data.clear()
            st.session_state.exp_edit_msg = f"Saved the changes to {rid} ({len(changes)} change(s) logged)."
    st.session_state.reset_exp_choice = True
    st.rerun()


with tab_exp:
    if expenses_ws is not None:
        render_expense_editor()

with tab_edit:
    st.subheader("Edit a saved booking")
    st.caption("Change or add details of a booking that is already saved. Every change is written to the "
               "'Edit Log' tab, with the old and new value, your name and the time.")
    if log_ws is None:
        st.error("The 'Edit Log' tab could not be prepared in your Google Sheet, so editing is switched off. "
                 "If a tab called 'Edit Log' already exists with other headings, rename it and restart the app.")
    else:
        msg = st.session_state.pop("edit_msg", None)
        if msg:
            st.success(msg)
        if st.session_state.pop("reset_edit_choice", False):
            st.session_state["edit_choice"] = None
        groups = booking_groups()
        if not groups:
            st.caption("No bookings saved yet.")
        else:
            labels = [group_label(g) for g in groups]
            choice = st.selectbox("Choose the booking (type a brand, person, phone number or ID to search)",
                                  labels, index=None, key="edit_choice")
            if choice is not None:
                render_booking_editor(groups[labels.index(choice)])
        with st.expander("Recent changes (Edit Log)"):
            lv = read_all(log_ws, LOG_TITLE)
            if len(lv) > 1:
                ld = [([str(c) for c in r] + [""] * len(LOG_HEADERS))[:len(LOG_HEADERS)] for r in lv[1:]]
                st.dataframe(pd.DataFrame(ld, columns=LOG_HEADERS).iloc[::-1].head(40), hide_index=True)
            else:
                st.caption("No changes recorded yet.")


# ----------------------------------------------------------------------------
# Home-page summary (shown above the tabs)
# ----------------------------------------------------------------------------
def render_dashboard():
    rows = load_all_rows()
    stalls = {stall_key(d) for d in rows if stall_key(d)}
    per = {ev: [0, 0.0] for ev in EVENTS}
    for d in rows:
        if d["_event"] in per:
            per[d["_event"]][0] += 1
            per[d["_event"]][1] += _amount(d.get("Paid Amount (This Event)", ""))
    collected = sum(v[1] for v in per.values())
    spent = 0.0
    if expenses_ws is not None:
        ev = read_all(expenses_ws, EXPENSE_TITLE)
        spent = sum(_amount(r[3]) for r in ev[1:] if len(r) > 3)
    profit = collected - spent
    cards = [("Stalls booked", str(len(stalls)), ""), ("Collected", rs(collected), ""),
             ("Expenses", rs(spent), ""), ("Profit so far", rs(profit), " sbk-pf")]
    html = '<div class="sbk-cards">' + "".join(
        f'<div class="sbk-card{c}"><div class="sbk-k">{k}</div><div class="sbk-v">{v}</div></div>'
        for k, v, c in cards) + "</div>"
    html += '<div class="sbk-evs">'
    for ev_name in EVENTS:
        n, amt = per[ev_name]
        pct = round(100 * amt / collected) if collected > 0 else 0
        color = EVENT_COLORS.get(ev_name, "#B3267F")
        html += (f'<div class="sbk-ev"><div class="sbk-row"><span class="sbk-tag" style="background:{color}">{ev_name}</span>'
                 f'<span class="sbk-n">{n} stall{"s" if n != 1 else ""} · {rs(amt)}</span></div>'
                 f'<div class="sbk-bar"><i style="width:{pct}%;background:{color}"></i></div></div>')
    html += "</div>"
    html += '<div class="sbk-dia">' + "".join(
        f'<i style="background:{c}"></i>' for c in ["#B3267F", "#F5C230", "#1FA39A", "#E8702A"] * 2) + "</div>"
    st.markdown(html, unsafe_allow_html=True)


with dash:
    try:
        render_dashboard()
    except Exception as e:  # noqa - the summary is a bonus, never block the app
        print(f"Dashboard could not be drawn: {e}")
