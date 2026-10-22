"""Parses and validates pasted stall-booking text. No Streamlit / Google code here."""
import re
from datetime import date
from difflib import get_close_matches

from dateutil import parser as dateparser

EVENTS = ["Jacranda", "AMBLEM", "Qawali Night"]

# Every spelling we accept -> the one official event name (= the sheet/tab name)
EVENT_ALIASES = {
    "jacranda": "Jacranda", "jacaranda": "Jacranda", "jakranda": "Jacranda",
    "jacarandah": "Jacranda", "jacranda night": "Jacranda",
    "amblem": "AMBLEM", "emblem": "AMBLEM", "amblum": "AMBLEM", "ambilem": "AMBLEM",
    "qawali night": "Qawali Night", "qawwali night": "Qawali Night",
    "qawali": "Qawali Night", "qawwali": "Qawali Night", "kawali night": "Qawali Night",
    "qawali nite": "Qawali Night", "qawalli night": "Qawali Night",
}

# label written by the user (lower-case, no % sign) -> internal field
FIELD_KEYS = {
    "name": "name",
    "brand name": "brand", "brand": "brand",
    "brand items": "items", "items": "items",
    "booked for": "events", "booked": "events", "event": "events", "events": "events",
    "paid amount": "paid", "paid": "paid", "amount": "paid",
    "offered discount": "discount", "offered discount %": "discount",
    "discount": "discount",
    "booking date": "date", "date": "date",
    "contact": "contact", "phone": "contact", "contact number": "contact",
    "payment method": "method", "payment mode": "method", "payment type": "method",
    "payment": "method", "paid via": "method", "paid by": "method",
    "received by": "received", "collected by": "received",
}

PAYMENT_METHODS = ["Bank transfer", "JazzCash / Easypaisa", "Cash", "Cheque", "No payment yet"]

LINE_RE = re.compile(r"^\s*([^:：\n]{1,40}?)\s*[:：]\s*(.*)$")


def _clean_key(k: str) -> str:
    return re.sub(r"\s+", " ", k.replace("%", "").strip().lower())


def parse_events(text: str):
    """Returns (list_of_official_events, list_of_unrecognised_words)."""
    found, bad = [], []
    for part in re.split(r"[,;/&+\n]|\band\b", text, flags=re.I):
        p = re.sub(r"[^a-z ]", "", part.lower()).strip()
        p = re.sub(r"\s+", " ", p)
        if not p:
            continue
        if p in EVENT_ALIASES:
            ev = EVENT_ALIASES[p]
        else:
            m = get_close_matches(p, EVENT_ALIASES.keys(), n=1, cutoff=0.78)
            ev = EVENT_ALIASES[m[0]] if m else None
        if ev is None:
            bad.append(part.strip())
        elif ev not in found:
            found.append(ev)
    return found, bad


def parse_method(text: str):
    t = re.sub(r"[^a-z ]", " ", text.lower())
    if not t.strip():
        return None
    if any(w in t for w in ("jazz", "easy", "paisa", "sadapay", "nayapay")):
        return "JazzCash / Easypaisa"          # checked first: "JazzCash" contains the word "cash"
    if "cash" in t:
        return "Cash"
    if any(w in t for w in ("bank", "transfer", "raast", "ibft", "online", "account")):
        return "Bank transfer"
    if "cheque" in t or "check" in t:
        return "Cheque"
    if "not paid" in t or "unpaid" in t or "no payment" in t or "pending" in t:
        return "No payment yet"
    return None


def parse_amount(text: str):
    t = text.lower().replace(",", "").replace("/-", "")
    t = re.sub(r"\b(rs|pkr|rupees?)\b\.?", "", t).strip()
    m = re.fullmatch(r"(\d+(?:\.\d+)?)\s*(k|lac|lakh|lacs|lakhs)?", t)
    if not m:
        raise ValueError(f"Could not read paid amount from '{text}'")
    val = float(m.group(1))
    mult = {"k": 1_000, "lac": 100_000, "lakh": 100_000, "lacs": 100_000, "lakhs": 100_000}
    val *= mult.get(m.group(2), 1)
    return int(val) if val == int(val) else round(val, 2)


def parse_discount(text: str):
    t = re.sub(r"(%|percent|percentage)", "", text.lower()).strip()
    if not t:
        return 0
    try:
        val = float(t)
    except ValueError:
        raise ValueError(f"Could not read discount from '{text}'")
    if not 0 <= val <= 100:
        raise ValueError(f"Discount {val}% must be between 0 and 100")
    return int(val) if val == int(val) else val


def parse_date(text: str) -> date:
    try:
        return dateparser.parse(text, dayfirst=True, fuzzy=False).date()
    except (ValueError, OverflowError, TypeError):
        raise ValueError(f"Could not read booking date from '{text}'")


def normalize_contact(text: str):
    """Returns (clean_contact, warning_or_None). Pakistani mobiles -> 03XX-XXXXXXX."""
    digits = re.sub(r"\D", "", text)
    if digits.startswith("0092"):
        digits = "0" + digits[4:]
    elif digits.startswith("92") and len(digits) == 12:
        digits = "0" + digits[2:]
    elif digits.startswith("3") and len(digits) == 10:
        digits = "0" + digits
    if len(digits) == 11 and digits.startswith("03"):
        return f"{digits[:4]}-{digits[4:]}", None
    return text.strip(), f"Contact '{text.strip()}' is not a standard 11-digit mobile (03XX-XXXXXXX). Please check it."


def parse_text(raw: str):
    """
    Turns the pasted block into a dict of clean values.
    Returns (data, errors, warnings). Nothing is guessed silently: anything
    unclear goes into errors/warnings so the human confirms it before saving.
    """
    fields, errors, warnings = {}, [], []
    current = None
    for line in raw.splitlines():
        if not line.strip():
            continue
        m = LINE_RE.match(line)
        if m and _clean_key(m.group(1)) in FIELD_KEYS:
            current = FIELD_KEYS[_clean_key(m.group(1))]
            if current in fields and fields[current]:
                warnings.append(f"'{m.group(1).strip()}' appears more than once - the last one was used.")
                fields[current] = ""
            fields[current] = m.group(2).strip()
        elif m:
            warnings.append(f"Line ignored (unknown label): {line.strip()}")
            current = None
        elif current:  # continuation of a multi-line value (e.g. long item lists)
            prev = fields[current].rstrip()
            if current == "items" and prev and not prev.endswith(","):
                prev += ","
            fields[current] = (prev + " " + line.strip()).strip()
        else:
            warnings.append(f"Line ignored: {line.strip()}")

    data = {}
    data["name"] = fields.get("name", "").strip()
    data["brand"] = fields.get("brand", "").strip()
    data["items"] = fields.get("items", "").strip()
    for key, label in (("name", "Name"), ("brand", "Brand Name"), ("items", "Brand Items")):
        if not data[key]:
            errors.append(f"{label} is missing.")

    evs, bad = parse_events(fields.get("events", ""))
    data["events"] = evs
    if not fields.get("events", "").strip():
        errors.append("'Booked for' is missing.")
    for b in bad:
        errors.append(f"Event '{b}' not recognised. Valid events: {', '.join(EVENTS)}.")
    if fields.get("events", "").strip() and not evs and not bad:
        errors.append("No valid event found in 'Booked for'.")

    for key, fn, label in (("paid", parse_amount, "Paid Amount"),
                           ("date", parse_date, "Booking date")):
        if not fields.get(key, "").strip():
            errors.append(f"{label} is missing.")
            data[key] = None
            continue
        try:
            data[key] = fn(fields[key])
        except ValueError as e:
            errors.append(str(e))
            data[key] = None

    if fields.get("discount", "").strip() == "":
        warnings.append("Offered Discount % was empty - treated as 0%.")
    try:
        data["discount"] = parse_discount(fields.get("discount", ""))
    except ValueError as e:
        errors.append(str(e))
        data["discount"] = 0

    data["method"] = parse_method(fields.get("method", ""))
    if fields.get("method", "").strip() and data["method"] is None:
        warnings.append(f"Payment method '{fields['method'].strip()}' not recognised - please choose it below.")
    data["received_by"] = fields.get("received", "").strip()

    if not fields.get("contact", "").strip():
        errors.append("Contact is missing.")
        data["contact"] = ""
    else:
        data["contact"], w = normalize_contact(fields["contact"])
        if w:
            warnings.append(w)
    return data, errors, warnings


def split_amount(total, n):
    """Equal split that always adds back to exactly `total`."""
    base = round(total / n, 2)
    parts = [base] * n
    parts[0] = round(total - base * (n - 1), 2)
    return [int(p) if float(p).is_integer() else p for p in parts]
