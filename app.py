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
from PIL import Image, ImageDraw, ImageFont, ImageOps
from google.oauth2.credentials import Credentials as UserCredentials
from google.oauth2.service_account import Credentials

from booking_parser import EVENTS, PAYMENT_METHODS, parse_text, split_amount

# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Stall Booking Records", page_icon="🧾", layout="wide")

BASE = Path(__file__).parent
CRED_FILE = BASE / "credentials.json"      # downloaded from Google Cloud (see guide)
CONFIG_FILE = BASE / "config.json"         # created automatically by the sidebar
FILES_DIR = BASE / "booking_files"         # logos / receipts are always saved here too
TZ = ZoneInfo("Asia/Karachi")
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


def make_receipt_png(rows):
    W, M, LABEL_W, LH = 900, 50, 290, 40
    f_title, f_sub = _font(40, True), _font(24)
    f_label, f_val = _font(24), _font(28, True)
    laid, body_h = [], 0
    for label, lines, _text in rows:
        wrapped = []
        for ln in lines:
            wrapped += textwrap.wrap(str(ln), width=28) or [""]
        laid.append((label, wrapped))
        body_h += len(wrapped) * LH + 22
    H = 190 + body_h + 110
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, W, 150], fill=(31, 111, 74))
    d.text((M, 32), "Booking Confirmation", font=f_title, fill="white")
    d.text((M, 95), EXPO_TITLE, font=f_sub, fill=(220, 240, 228))
    y = 190
    for label, wrapped in laid:
        d.text((M, y + 6), label, font=f_label, fill=(110, 110, 110))
        for i, ln in enumerate(wrapped):
            d.text((M + LABEL_W, y + i * LH), ln, font=f_val, fill=(30, 30, 30))
        y += len(wrapped) * LH + 22
        d.line([M, y - 10, W - M, y - 10], fill=(228, 228, 228), width=2)
    d.text((M, y + 20), "Please keep your booking ID for any future contact.", font=f_sub, fill=(90, 90, 90))
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
SCHEMA_VERSION = "2026-10-27-tolerant-headings"   # bump when the sheet layout changes
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


def next_expense_id(ws):
    top = 0
    for v in ws.col_values(1)[1:]:
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
    st.title("🔒 Stall Booking Records")
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
        ensure_summary_sheet(sh)
    except Exception as e:  # noqa - the summary is a bonus; never stop bookings because of it
        print(f"Summary sheet could not be created: {e}")
    return sh, sheets, expenses_ws


try:
    sh, sheets, expenses_ws = connect(cfg["sheet_url"], SCHEMA_VERSION)
except Exception as e:  # noqa
    st.error(f"Could not open the Google Sheet: {e}")
    st.stop()

st.sidebar.success(f"Connected: {sh.title}")

# ----------------------------------------------------------------------------
# Main UI
# ----------------------------------------------------------------------------
st.title("🧾 Stall Booking Records")
tab_new, tab_view, tab_search, tab_exp = st.tabs(["➕ New booking", "📊 Records", "🔎 Search", "💸 Expenses"])

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
                            exp_id = next_expense_id(expenses_ws)
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
        spent = sum(_amount(r[3]) for r in data)
        collected = total_collected()
        c1, c2, c3 = st.columns(3)
        c1.metric("Collected from stalls", rs(collected))
        c2.metric("Total expenses", rs(spent))
        c3.metric("Profit so far", rs(collected - spent))
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
