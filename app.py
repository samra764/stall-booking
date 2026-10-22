import hmac
import io
import json
import uuid
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
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
           "Total Paid (Whole Booking)", "Discount %", "Booking Date",
           "Contact", "Logo", "Payment Receipt", "Payment Method", "Received By"]
OLD_HEADERS = HEADERS[:14]   # first version of the sheet (before payment method existed)
COL_BRAND, COL_CONTACT, COL_PAID = 3, 11, 7   # 0-based positions used below


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
            ws.format("A1:P1", {"textFormat": {"bold": True},
                                "backgroundColor": {"red": 0.85, "green": 0.92, "blue": 0.85}})
        else:
            ws = existing[ev]
            first = ws.row_values(1)
            if first == OLD_HEADERS:                       # upgrade older tab: add 2 columns at the end
                if ws.col_count < len(HEADERS):
                    ws.add_cols(len(HEADERS) - ws.col_count)
                ws.update(range_name="A1:P1", values=[HEADERS])
            elif first != HEADERS:
                if not any(ws.get_all_values()):          # empty tab -> write headers
                    ws.append_row(HEADERS, value_input_option="RAW")
                else:
                    raise RuntimeError(
                        f"Tab '{ev}' already has data with different column headings. "
                        "Rename or delete that tab so the app can create it correctly.")
        sheets[ev] = ws
    return sheets


@st.cache_data(ttl=15, show_spinner=False)
def read_all(_ws, title):
    return _ws.get_all_values()


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
            refs[label] = "No payment yet"
            continue
        ext = Path(f.name).suffix.lower() or ".bin"
        data = f.getvalue()
        if local_dir:
            (local_dir / f"{label}{ext}").write_bytes(data)
            refs[label] = str(local_dir / f"{label}{ext}")
        if drive:
            try:
                from googleapiclient.http import MediaIoBaseUpload
                media = MediaIoBaseUpload(io.BytesIO(data), mimetype=f.type or "application/octet-stream")
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
            "f_paid": 0.0, "f_discount": 0.0,
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
    st.session_state.f_discount = float(data["discount"] or 0)
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
def connect(url):
    sh = open_spreadsheet(url)
    return sh, ensure_event_sheets(sh)


try:
    sh, sheets = connect(cfg["sheet_url"])
except Exception as e:  # noqa
    st.error(f"Could not open the Google Sheet: {e}")
    st.stop()

st.sidebar.success(f"Connected: {sh.title}")

# ----------------------------------------------------------------------------
# Main UI
# ----------------------------------------------------------------------------
st.title("🧾 Stall Booking Records")
tab_new, tab_view = st.tabs(["➕ New booking", "📊 Records"])

with tab_new:
    st.subheader("Step 1 - Paste the booking details")
    st.text_area("Booking text", key="raw_text", height=230,
                 placeholder="Name:\nBrand Name:\nBrand Items:\nBooked for: Jacranda, AMBLEM, qawali night\n"
                             "Paid Amount:\nOffered Discount %:\nBooking date:\nContact:")
    st.button("🔍 Read details", on_click=do_parse, type="primary",
              disabled=not st.session_state.raw_text.strip())

    if st.session_state.last_saved:
        st.success(st.session_state.last_saved)

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
            st.number_input("Offered Discount %", min_value=0.0, max_value=100.0,
                            step=1.0, key="f_discount")
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
        if method == "Cash":
            st.text_input("Cash received by (name of the team member who took the cash)",
                          key="f_received")
            st.caption("For cash, the receipt photo is optional (e.g. a photo of the handwritten "
                       "receipt or the cash handover).")
        u1, u2 = st.columns(2)
        with u1:
            logo = st.file_uploader("Brand logo", type=["png", "jpg", "jpeg", "webp", "pdf"],
                                    key=f"logo_{st.session_state.uploader_n}")
            if logo and logo.type.startswith("image"):
                st.image(logo, width=160)
        with u2:
            receipt = st.file_uploader(
                "Cash receipt / photo (optional)" if method == "Cash"
                else "No payment yet - nothing to upload" if method == "No payment yet"
                else "Payment receipt", type=["png", "jpg", "jpeg", "webp", "pdf"],
                                       key=f"receipt_{st.session_state.uploader_n}")
            if receipt and receipt.type.startswith("image"):
                st.image(receipt, width=160)

        # ------------------------- validation -------------------------------
        problems = []
        for label, key in (("Name", "f_name"), ("Brand Name", "f_brand"),
                           ("Brand Items", "f_items"), ("Contact", "f_contact")):
            if not str(st.session_state[key]).strip():
                problems.append(f"{label} is empty.")
        if not events:
            problems.append("Select at least one event.")
        if logo is None:
            problems.append("Logo not uploaded.")
        if method is None:
            problems.append("Choose how the payment was made.")
        elif method == "No payment yet":
            if paid != 0:
                problems.append("Method is 'No payment yet' but a paid amount is entered - set it to 0 "
                                "or choose the real payment method.")
        else:
            if paid <= 0:
                problems.append("Paid amount is 0 - enter the amount, or choose 'No payment yet'.")
            if method == "Cash":
                if not st.session_state.f_received.strip():
                    problems.append("Enter who received the cash.")
            elif receipt is None:
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
            booking_id = f"BK{datetime.now(TZ):%y%m%d-%H%M%S}-{uuid.uuid4().hex[:3].upper()}"
            now = datetime.now(TZ).strftime("%d-%b-%Y %H:%M")
            with st.spinner("Saving..."):
                try:
                    logo_ref, receipt_ref, drive_warn = save_uploads(booking_id, logo, receipt)
                    if receipt is None and method == "Cash":
                        receipt_ref = "Cash - no receipt photo"
                except Exception as e:  # noqa
                    st.error(f"Nothing was saved. {e}")
                    st.stop()
                saved_to = []
                try:
                    for ev in events:
                        row = [booking_id, now, st.session_state.f_name.strip(),
                               st.session_state.f_brand.strip(), st.session_state.f_items.strip(),
                               ", ".join(events), ev, amounts[ev], paid,
                               st.session_state.f_discount,
                               st.session_state.f_date.strftime("%d-%b-%Y"),
                               st.session_state.f_contact.strip(),
                               logo_ref, receipt_ref, method,
                               st.session_state.f_received.strip() if method == "Cash" else ""]
                        sheets[ev].append_row(row, value_input_option="RAW")
                        saved_to.append(ev)
                except Exception as e:  # noqa
                    st.error(f"Saved to {', '.join(saved_to) or 'nothing'} but then failed on the next "
                             f"sheet: {e}. Booking ID {booking_id}. Add the missing row manually "
                             "or fix the error and save again after deleting the partial row.")
                    st.stop()
            if drive_warn:
                st.warning(drive_warn)
            msg = (f"Saved {st.session_state.f_brand.strip()} "
                   f"to {', '.join(saved_to)} (ID {booking_id}).")
            st.cache_data.clear()
            st.session_state.last_saved = msg
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
