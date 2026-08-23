
#!/usr/bin/env python3
"""
AP payment list parser for the payment-list template.

Update in this version:
- ignores the No. column entirely
- only reads the main left-side payment table
- skips subtotal/total/summary rows
- skips rows that do not have a real Amount
- maps E-form No. only from the real "E-form No." header
- preserves source row order
- writes a clean template workbook
- fix: "Account number" now maps only to the vendor/beneficiary bank
  account column. It previously also matched "Pay from" / "Ordering
  Party" (the company's own paying account), and since "Pay from" sits
  left of "Account number" in the source sheets, every row ended up with
  the company's account instead of the vendor's.
- adds "Entity" and "Source File" columns, guessed from each source
  file's name. This lets the dashboard filter/group by entity even
  though this parser merges every source file into one PAYMENT_LIST
  sheet — without a per-row entity, the dashboard has no way to tell
  VNGSING rows apart from ZPI rows once they are combined.
- fix: Amount now comes from the row's own transaction-amount column,
  not from a currency-converted reporting column. VNGSING's "Citi
  Payment list" sheet carries both "USD Amount" (col C, the USD
  equivalent) and "SỐ TIỀN" (col K, the amount actually paid, which is
  what the CURRENCY column labels). The generic "amount" alias matched
  "USD Amount" first, so every non-USD row got the USD figure stamped
  with its local currency (e.g. IDR 3,731.89 instead of the real IDR
  56,650,000), and every row whose converted column happened to be
  blank was dropped outright. See LOCAL_AMOUNT_HEADERS /
  CONVERTED_AMOUNT_HEADERS and amount_from_row().
- fix: the main-table width is now measured per sheet instead of being
  hardcoded to 15 columns. VNGSING's sheet keeps the vendor account
  ("STK") in col P and "Swiftcode" in col Q, both past the old cap, so
  every VNGSING row exported a blank bank account and blank SWIFT and
  could never satisfy the dashboard's 4-field bank-reconcile check. The
  cap could not simply be raised: the standard template puts an
  auxiliary validation table ("Reason to pay", "Beneficiary Name",
  "Check", "Amount", ...) further right on the same row, which is what
  the cap was protecting against. main_table_width() instead takes the
  contiguous run of populated header cells and stops at the first gap,
  which separates the two cases on every real file seen so far.
- fix: sheets within one workbook that restate the same payments no
  longer double-count. KMZ's and VNGGAMES Sing's files carry both a
  "daily" and a "Total" sheet listing the same payments, so those rows
  landed in the Payment List twice (the copies are not byte-identical —
  "daily" is the fresher one and "Total" leaves value date / bene name /
  account blank on some rows — which is why a plain identical-row check
  did not catch them). See dedupe_across_sheets().
- fix: summary-row detection now only reads the structured columns.
  Matching SUMMARY_KEYWORDS against every cell on the row silently
  discarded real payments whose prose happened to contain one: a VNGSING
  invoice described with the word "balance", and three ZPI GOOGLE ASIA
  PACIFIC invoices whose "Reason to pay" note ends "... applied on the
  total amount". See summary_scan_skip_columns().

Usage:
    python ap_payment_template_parser_fixed_v2.py --input "C:\\path\\to\\input_ap" --output "C:\\path\\to\\AP_PAYMENT_TEMPLATE_OUTPUT.xlsx"
"""

from __future__ import annotations

import argparse
import csv
import re
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook


OUTPUT_HEADERS = [
    "Due Date(Payment date)",
    "Supplier Number",
    "Invoice Number",
    "Due Date(Value date)",
    "Bene name",
    "Amount",
    "Currency",
    "E-form No.",
    "Bank charge",
    "Detail",
    "Account number",
    "Swiftcode",
    "Entity",
    "Source File",
]

SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".csv"}

# Only scan the left-most main table columns. This is an upper bound for the
# header scan, not the width actually used — see main_table_width(), which
# measures each sheet's own header block. It used to be a flat 15 (A:O), which
# cut VNGSING's vendor account ("STK", col P) and "Swiftcode" (col Q) out of
# every row.
MAIN_TABLE_MAX_COLS = 30


def main_table_width(header_cells: List[Any]) -> int:
    """Width of the contiguous populated header block, in columns.

    The standard payment-list template repeats a *validation* table further
    right on the same header row ("Reason to pay", "Beneficiary Name",
    "Check", "Amount", "Curr", "Check", "Beneficiary Account", ...), separated
    from the real table by at least one empty header cell. VNGSING's Citi
    sheet instead continues the real table out to col T ("STK", "Swiftcode",
    "Benificiary Bank", "More detailed information", "Due date") with no gap.

    Taking the contiguous run and stopping at the first gap keeps the real
    table in both shapes and leaves the validation table out — a fixed cap
    cannot do both, which is how VNGSING lost its account/SWIFT columns.
    """
    first = None
    for idx, cell in enumerate(header_cells[:MAIN_TABLE_MAX_COLS]):
        populated = cell is not None and str(cell).strip() != ""
        if populated and first is None:
            first = idx
        elif not populated and first is not None:
            return idx
    return len(header_cells[:MAIN_TABLE_MAX_COLS]) if first is not None else 0


SUMMARY_KEYWORDS = [
    "subtotal",
    "sub total",
    "grand total",
    "control total",
    "total",
    "summary",
    "balance",
    "page total",
    "carry forward",
    "carryforward",
    "brought forward",
    "carried forward",
    "opening balance",
    "closing balance",
    "net total",
    "overall total",
]


def norm(v: Any) -> str:
    return re.sub(r"\s+", " ", str(v or "")).strip().lower()


def norm_key(v: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", norm(v))


def guess_entity(filename: str) -> str:
    """Mirrors the dashboard's guessEntity() so entity names line up exactly."""
    t = filename.lower()
    if "vngsing" in t:
        return "VNGSING"
    if "vngg" in t:
        return "VNGGAMES Sing"
    if "zpi" in t:
        return "ZPI"
    if "vngth" in t or "gameco" in t:
        return "GAMECO"
    if "kmz" in t:
        return "KMZ"
    if "gnth" in t or "gn th" in t:
        return "GN TH"
    if "vngi" in t:
        return "VNG I"
    if "greennode" in t:
        return "GREENNODE SG"
    return "AP"


ALIASES: Dict[str, List[str]] = {
    # NOTE: "due date" alone is deliberately NOT in this list — see
    # payment_date_from_row(). On VNGSING's Citi sheet "Due date" is the
    # invoice's due date (a different figure from the payment date on 94% of
    # rows), while the actual payment date lives in a column headed just
    # "Date". The loose "due date" alias used to grab the wrong column.
    "payment_date": ["due date(payment date)", "payment date"],
    "supplier_number": ["supplier number", "supplier num", "supplier no", "supplier id"],
    "invoice_number": ["invoice number", "ap invoice number", "invoice no", "ap invoice"],
    "value_date": ["due date(value date)", "value date"],
    # "Supplier name" / "tên đối tác" (Vietnamese for "counterparty name")
    # cover an alternate template (seen on VNGSING's and Greennode's Citi/DB
    # payment-list exports) that names the vendor/counterparty column
    # differently from the standard "Bene name" header.
    "beneficiary": ["bene name", "beneficiary name", "beneficiary", "vendor name", "vendor", "supplier name", "tên đối tác"],
    "amount": ["amount", "payment amount", "invoice amount", "group amount"],
    "currency": ["currency", "ccy"],
    "eform_no": ["e-form no.", "e-form no", "eform no.", "eform no", "e-form number", "eform number"],
    "bank_charge": ["bank charge", "charge", "charges"],
    "detail": ["detail", "description", "payment details", "remarks"],
    # Vendor/beneficiary bank account only. Do NOT alias this to "pay from" /
    # "ordering party" / "source account" — those refer to the company's own
    # paying account, not the vendor's bank account, and are a different field
    # (see "pay_from" below). Conflating them here caused the output
    # "Account number" column to be filled with the company's own account
    # number (constant per entity) instead of each vendor's real bank account.
    "account_number": [
        "account number",
        "bank account",
        "beneficiary account",
        "beneficiary account number",
        "bene account",
        "vendor account",
        "vendor bank account",
        "supplier bank account",
        # "STK" = "số tài khoản" (account number) on VNGSING's Citi sheet,
        # which uses no English header for this column.
        "stk",
    ],
    "swiftcode": ["swiftcode", "swift code", "swift", "bic"],
    # The company's own source/ordering account (constant per entity), kept
    # separate from the vendor's "account_number" above.
    "pay_from": ["pay from", "ordering party", "source account", "paying account"],
}


# The amount actually paid, in the row's own currency — i.e. the figure the
# "Currency" column is labelling. Preferred over anything in ALIASES["amount"].
#
# Matched by exact normalized key rather than through header_matches(), because
# norm_key() strips Vietnamese accents down to very short keys that the loose
# substring branch would over-match: norm_key("SỐ TIỀN") is "stin", which is a
# substring of "destination".
LOCAL_AMOUNT_HEADERS = [
    "số tiền",
    "số tiền thanh toán",
    "số tiền ngoại tệ",
]

# Currency-converted / reporting amounts. These are NOT the amount paid, so
# they must never satisfy the plain "amount" lookup for a non-USD row: on
# VNGSING's Citi sheet "USD Amount" is the USD equivalent of a payment whose
# Currency column says IDR/THB/VND, and pairing the two understated every
# non-USD payment by the FX rate. Still usable as a last resort on a USD row,
# where the converted figure and the paid figure are the same number.
CONVERTED_AMOUNT_HEADERS = [
    "usd amount",
    "amount usd",
    "usd amt",
    "tổng số tiền nợ usd",
]


MONTH_NAMES = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def parse_date(v: Any) -> Any:
    if v is None or v == "":
        return ""
    if hasattr(v, "year") and hasattr(v, "month") and hasattr(v, "day"):
        try:
            return v.date() if hasattr(v, "date") else v
        except Exception:
            return v
    s = str(v).strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", s)
    if m:
        return s
    m = re.match(r"^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$", s)
    if m:
        a, b, y = m.groups()
        a_i, b_i = int(a), int(b)
        # This template is normally DD/MM/YYYY, but a source cell stored as
        # plain text can be MM/DD/YYYY instead with no way to tell just by
        # looking at one value - e.g. "06/17/2026" (17 can't be a month) was
        # blindly read as day=06/month=17, producing an invalid "2026-17-06"
        # that then threw off period/date-span logic downstream. When only
        # one of the two numbers could possibly be a month, treat the other
        # one as the day regardless of which position it was in; a value
        # where both numbers are <=12 stays genuinely ambiguous, so keep the
        # existing DD/MM assumption for that case.
        if a_i > 12 and b_i <= 12:
            d, mo = a, b
        elif b_i > 12 and a_i <= 12:
            d, mo = b, a
        else:
            d, mo = a, b
        return f"{y}-{mo.zfill(2)}-{d.zfill(2)}"

    # Month-name text, e.g. "14-Aug-2026" / "Aug 14 2026", in either order.
    # A value with no year at all (real examples: "Aug-14", "Aug 14" on
    # VNGGAMES Sing) is deliberately left as-is rather than having a year
    # guessed for it — this is a payment value date, so an invented year is
    # worse than a visibly unparsed one that a human corrects at source.
    m = re.match(
        r"^(?:(\d{1,2})[\s\-/]+)?([A-Za-z]{3,9})(?:[\s\-/]+(\d{1,2}))?[\s\-/]+(\d{4})$", s
    )
    if m:
        pre_day, month_name, post_day, year = m.groups()
        month = MONTH_NAMES.get(month_name[:3].lower())
        day = pre_day or post_day
        if month and day:
            return f"{year}-{month:02d}-{int(day):02d}"
    return s


def to_number(v: Any) -> Optional[float]:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("$", "").replace(" ", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except Exception:
        return None


def header_matches(header: Any, alias: str) -> bool:
    h = norm_key(header)
    a = norm_key(alias)
    if not h:
        return False
    if h == a:
        return True
    # Allow punctuation / spacing differences.
    if a in h and len(h) >= len(a) - 1:
        return True
    return False


def is_summary_row(values: List[Any], skip_idx: Optional[set] = None) -> bool:
    """Is this row a control / total / subtotal line rather than a payment?

    `skip_idx` holds column positions to leave out of the keyword scan — the
    free-text detail and beneficiary columns. A vendor name or an invoice
    description is allowed to contain words like "total" or "balance" without
    that making the row a control line; matching SUMMARY_KEYWORDS against the
    whole row silently dropped a real VNGSING payment (FA-PM240506042, THB
    137,993.03) whose description happened to include the word "balance".
    """
    skip = skip_idx or set()
    text = " ".join(
        norm(v) for i, v in enumerate(values) if v not in (None, "") and i not in skip
    )
    if not text:
        return False

    # Skip control / total / subtotal / summary lines that AP places in the
    # workbook for reconciliation or control purposes.
    for keyword in SUMMARY_KEYWORDS:
        if keyword in text:
            return True

    # Catch common summary phrases that appear in the first columns.
    compact = re.sub(r"\s+", " ", text).strip()
    if re.match(
        r"^(subtotal|sub total|grand total|control total|page total|"
        r"opening balance|closing balance|brought forward|carried forward|"
        r"balance b/f|balance c/f|net total|overall total)\b",
        compact,
    ):
        return True

    return False


def is_invoice_line(
    invoice_no: Any,
    eform_no: Any,
    amount_raw: Any,
    values: List[Any],
    skip_idx: Optional[set] = None,
) -> bool:
    """Return True only for real invoice lines, not control / summary rows.

    A real payment line always needs a real reference — normally the
    Invoice Number, but some source templates (VNGSING's and Greennode's
    Citi/DB exports) carry no separate Invoice Number column at all and
    use only E-form No. / Transaction No. as the reference. E-form No. is
    also the field the dashboard actually keys the payment reference off
    of everywhere downstream, so accept it as an equally valid anchor
    instead of dropping every row from a file just because it lacks a
    column the standard template happens to also have.
    """
    return invoice_line_reason(invoice_no, eform_no, amount_raw, values, skip_idx) is None


def invoice_line_reason(
    invoice_no: Any,
    eform_no: Any,
    amount_raw: Any,
    values: List[Any],
    skip_idx: Optional[set] = None,
) -> Optional[str]:
    """None for a real invoice line, otherwise the exclusion reason code.

    The reason feeds the PARSE_CONTROL completeness sheet: every candidate
    row a file contributes must end up either kept or counted under one of
    these reasons, so a parsing gap can never be silent.
    """
    inv = str(invoice_no or "").strip()
    eform = str(eform_no or "").strip()
    reference = inv or eform
    if not reference:
        return "no_reference"

    inv_norm = norm(reference)
    if inv_norm in {"-", "n/a", "na", "none", "row", "total", "subtotal"}:
        return "no_reference"
    if re.fullmatch(r"row[- ]?\d+", inv_norm):
        return "no_reference"

    amt = to_number(amount_raw)
    if amt is None or amt <= 0:
        return "no_amount"

    if is_summary_row(values, skip_idx):
        return "summary_row"

    has_business_detail = any(
        str(x).strip()
        for x in values
        if str(x or "").strip() and not re.fullmatch(r"(?:-|0+)", str(x).strip())
    )
    if not has_business_detail:
        return "summary_row"

    return None


HEADER_DETECTION_FIELDS = ["supplier_number", "invoice_number", "amount", "eform_no", "currency", "beneficiary"]


def detect_main_header_row(ws) -> Tuple[Optional[int], Optional[List[str]], int]:
    """
    Detect the row containing the left-side main payment-list headers.
    Returns (row index, headers, width) where width is that sheet's own
    contiguous header-block width from main_table_width() — we deliberately
    stop at the first gap in the header row so we do not pick up auxiliary
    validation tables that appear further right in the same sheet.

    Scores a candidate row by how many of ALIASES' field groups have a
    matching header, using the exact same punctuation-insensitive
    header_matches() that find_value() uses for extraction later — a
    header that would be recognized during extraction must never fail
    detection here only because this used a separate, stricter check.
    ("Eform No." vs "E-form No." previously did just that: extraction
    already tolerated the missing hyphen via header_matches, but this
    function's own hardcoded phrase list did a literal substring check
    that required the hyphen, so header-row detection failed and the
    whole sheet silently produced zero rows.)
    """
    for row_idx in range(1, min(ws.max_row, 12) + 1):
        raw = [c.value for c in ws[row_idx][:MAIN_TABLE_MAX_COLS]]
        width = main_table_width(raw)
        headers = ["" if v is None else str(v).strip() for v in raw[:width]]
        if not any(headers):
            continue
        hits = sum(
            1 for field in HEADER_DETECTION_FIELDS
            if any(header_matches(h, alias) for h in headers for alias in ALIASES[field])
        )
        if hits >= 4:
            return row_idx, headers, width
    return None, None, 0


def load_rows_from_csv(path: Path) -> List[List[Any]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows:
        return []
    headers = rows[0]
    out: List[List[Any]] = [headers]
    for raw in rows[1:]:
        out.append(raw + [""] * max(0, len(headers) - len(raw)))
    return out


def find_value(
    row_vals: List[Any],
    headers: List[str],
    aliases: List[str],
    exclude: Optional[List[str]] = None,
) -> Any:
    """First populated value whose header matches any alias, left to right.

    `headers` is already trimmed to the sheet's own main-table width by
    detect_main_header_row(), so there is no separate column cap here.
    `exclude` lists header phrases that must never satisfy the lookup even if
    an alias would otherwise match them (see CONVERTED_AMOUNT_HEADERS).
    """
    blocked = {norm_key(h) for h in (exclude or [])}
    for idx, header in enumerate(headers):
        if norm_key(header) in blocked:
            continue
        for alias in aliases:
            if header_matches(header, alias):
                if idx < len(row_vals):
                    v = row_vals[idx]
                    if v is not None and str(v).strip() != "":
                        return v
    return ""


def find_value_exact(row_vals: List[Any], headers: List[str], names: List[str]) -> Any:
    """Like find_value(), but the header must equal one of `names` outright.

    Used where the header keys are too short for header_matches()' substring
    branch to be safe — see the note on LOCAL_AMOUNT_HEADERS.
    """
    wanted = {norm_key(n) for n in names}
    for idx, header in enumerate(headers):
        if norm_key(header) in wanted and idx < len(row_vals):
            v = row_vals[idx]
            if v is not None and str(v).strip() != "":
                return v
    return ""


def payment_date_from_row(row_vals: List[Any], headers: List[str]) -> Any:
    """The date the payment is (to be) made, resolved in falling confidence:

    1. An explicit payment-date header ("Due Date (Payment date)" /
       "Payment date") — the standard template.
    2. A column headed exactly "Date" — VNGSING's Citi sheet and Greennode's
       DB sheet keep the payment date there (populated on 99.6% of VNGSING
       rows, vs 77% for its misleadingly-named "Due date" column). Exact
       match only: the loose matcher would also hit "Value date"/"Due date",
       since "date" is a substring of both.
    3. "Due date" as a last resort — on the sheets seen so far this is the
       invoice's due date, not the payment date (they differ on 94% of the
       VNGSING rows that carry both), so it only fills in when nothing
       better exists.
    """
    v = find_value(row_vals, headers, ALIASES["payment_date"])
    if str(v).strip() != "":
        return v
    v = find_value_exact(row_vals, headers, ["date"])
    if str(v).strip() != "":
        return v
    return find_value(row_vals, headers, ["due date"])


def amount_from_row(row_vals: List[Any], headers: List[str], currency: Any) -> Any:
    """The amount actually paid, in the currency the row is labelled with.

    Order matters:
    1. An explicit local transaction-amount column, if the sheet has one.
    2. Otherwise the generic "amount" alias, with converted/reporting columns
       excluded so a USD-equivalent figure can never be paired with a non-USD
       currency.
    3. On a USD row only, fall back to the converted column — there the
       converted figure IS the paid figure, and VNGSING has a few rows that
       carry nothing else. Doing this for any currency is exactly the bug this
       ordering exists to prevent.
    """
    local = find_value_exact(row_vals, headers, LOCAL_AMOUNT_HEADERS)
    if str(local).strip() != "":
        return local

    generic = find_value(row_vals, headers, ALIASES["amount"], exclude=CONVERTED_AMOUNT_HEADERS)
    if str(generic).strip() != "":
        return generic

    if str(currency or "").strip().upper() == "USD":
        return find_value_exact(row_vals, headers, CONVERTED_AMOUNT_HEADERS)
    return ""


# The columns whose values are structured enough that a SUMMARY_KEYWORDS hit in
# them really does mean "this is a control/total line". Everything else on the
# row is prose and gets excluded from that scan — see is_summary_row().
STRUCTURED_FIELDS = [
    "payment_date",
    "supplier_number",
    "invoice_number",
    "value_date",
    "amount",
    "currency",
    "eform_no",
    "bank_charge",
    "account_number",
    "swiftcode",
    "pay_from",
]


def summary_scan_skip_columns(headers: List[str]) -> set:
    """Positions to leave out of the summary-keyword scan.

    Anything that is not one of STRUCTURED_FIELDS — which covers the detail and
    beneficiary columns plus free-text extras the templates carry off to the
    right ("Reason to pay", "More detailed information", "Benificiary Bank").
    Real examples of prose that wrongly matched SUMMARY_KEYWORDS and silently
    dropped a genuine payment: a VNGSING description containing "balance", and
    ZPI's "Reason to pay" note reading "... will be applied on the total
    amount" on three GOOGLE ASIA PACIFIC invoices.

    Dropping the scan on these columns is safe because is_invoice_line() has
    already required a real reference and a positive amount before it runs, so
    an actual subtotal line has no way through anyway.
    """
    structured = set()
    for pos, header in enumerate(headers):
        for field in STRUCTURED_FIELDS:
            if any(header_matches(header, alias) for alias in ALIASES[field]):
                structured.add(pos)
                break
    return set(range(len(headers))) - structured


def extract_row(vals: List[Any], headers: List[str], skip_idx: set) -> Tuple[Optional[List[Any]], Optional[str]]:
    """Map one source row onto the 12 template columns.

    Returns (row, None) for a kept row, or (None, reason) — the reason feeds
    the PARSE_CONTROL completeness sheet. Shared by the xlsx and csv paths
    on purpose: these two used to carry independent copies of this logic,
    which is how a fix applied to one of them could silently miss the other.
    """
    if sum(1 for x in vals if x not in (None, "")) < 2:
        return None, "blank"

    # Skip repeated header rows or helper rows.
    row_text = " ".join(norm(x) for x in vals if isinstance(x, str))
    if "supplier number" in row_text and "invoice number" in row_text and "amount" in row_text:
        return None, "repeated_header"

    payment_date = payment_date_from_row(vals, headers)
    supplier_number = find_value(vals, headers, ALIASES["supplier_number"])
    invoice_number = find_value(vals, headers, ALIASES["invoice_number"])
    value_date = find_value(vals, headers, ALIASES["value_date"])
    beneficiary = find_value(vals, headers, ALIASES["beneficiary"])
    currency = find_value(vals, headers, ALIASES["currency"])
    amount_raw = amount_from_row(vals, headers, currency)
    eform_no = find_value(vals, headers, ALIASES["eform_no"])
    bank_charge = find_value(vals, headers, ALIASES["bank_charge"])
    detail = find_value(vals, headers, ALIASES["detail"])
    account_number = find_value(vals, headers, ALIASES["account_number"])
    swiftcode = find_value(vals, headers, ALIASES["swiftcode"])

    # Keep only real invoice lines.
    reason = invoice_line_reason(invoice_number, eform_no, amount_raw, vals, skip_idx)
    if reason is not None:
        return None, reason

    # Skip if the row is completely empty in the business columns.
    core_present = any(
        str(x).strip()
        for x in (
            supplier_number, invoice_number, beneficiary, amount_raw, currency, eform_no, account_number, swiftcode
        )
    )
    if not core_present:
        return None, "summary_row"

    return [
        parse_date(payment_date),
        str(supplier_number).strip(),
        str(invoice_number).strip(),
        parse_date(value_date),
        str(beneficiary).strip(),
        to_number(amount_raw),
        str(currency).strip(),
        str(eform_no).strip(),
        str(bank_charge).strip(),
        str(detail).strip(),
        str(account_number).strip(),
        str(swiftcode).strip(),
    ], None


def new_control() -> Dict[str, Any]:
    return {"candidates": 0, "blank": 0, "repeated_header": 0, "summary_row": 0,
            "no_amount": 0, "no_reference": 0, "mapped": 0,
            "dedup_dropped": 0, "period_dropped": 0, "window_dropped": 0,
            "kept": 0, "pending_sheets": 0, "no_header_sheets": 0}


def tally(ctl: Dict[str, Any], reason: Optional[str]) -> None:
    if reason == "blank":
        ctl["blank"] += 1
        return
    ctl["candidates"] += 1
    if reason is None:
        ctl["mapped"] += 1
    else:
        ctl[reason] += 1


def parse_sheet(ws, ctl: Optional[Dict[str, Any]] = None) -> List[List[Any]]:
    header_row_idx, headers, width = detect_main_header_row(ws)
    if not headers:
        if ctl is not None:
            ctl["no_header_sheets"] += 1
        return []

    skip_idx = summary_scan_skip_columns(headers)
    output: List[List[Any]] = []
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        parsed, reason = extract_row(list(row[:width]), headers, skip_idx)
        if ctl is not None:
            tally(ctl, reason)
        if parsed is not None:
            output.append(parsed)

    return output


def _payment_identity(row: List[Any]) -> Tuple[str, str, str, str]:
    # E-form, invoice number, amount, currency — enough to identify the same
    # payment restated on another sheet of the same workbook.
    return (
        norm(row[7]),
        norm(row[2]),
        "" if row[5] is None else f"{float(row[5]):.2f}",
        norm(row[6]),
    )


def _populated_fields(row: List[Any]) -> int:
    return sum(1 for v in row[:12] if str(v or "").strip())


def dedupe_across_sheets(per_sheet: List[List[List[Any]]]) -> List[List[Any]]:
    """Collapse the same payment appearing on more than one sheet of a file.

    KMZ's and VNGGAMES Sing's payment lists ship both a "daily" and a "Total"
    sheet covering the same payments, so parsing every sheet counted each of
    those twice. The two copies are not identical — "daily" is the fresher one
    and "Total" leaves value date / bene name / vendor account / SWIFT blank on
    some rows — so the more populated copy wins.

    Deliberately scoped to *across* sheets. Repeats within a single sheet are
    left alone: Greennode's list legitimately holds several distinct
    intercompany transfers that share an amount and carry the literal text
    "TRANSFER" instead of a real E-form, and collapsing those would delete
    real payments.
    """
    out: List[List[Any]] = []
    # key -> (position in `out`, index of the sheet it came from). The sheet
    # index is what keeps this scoped to *across* sheets: a repeat found on the
    # same sheet is always kept, so Greennode's several distinct "TRANSFER"
    # payments survive while KMZ's daily/Total restatements collapse.
    seen: Dict[Tuple[str, str, str, str], Tuple[int, int]] = {}
    for sheet_no, sheet_rows in enumerate(per_sheet):
        for row in sheet_rows:
            key = _payment_identity(row)
            # A row with no reference at all cannot be identified this way.
            if not key[0] and not key[1]:
                out.append(row)
                continue
            if key in seen:
                pos, from_sheet = seen[key]
                if from_sheet != sheet_no:
                    if _populated_fields(row) > _populated_fields(out[pos]):
                        out[pos] = row
                        seen[key] = (pos, sheet_no)
                    continue
            seen[key] = (len(out), sheet_no)
            out.append(row)
    return out


def parse_xlsx(path: Path) -> Tuple[List[List[Any]], Dict[str, Any]]:
    wb = load_workbook(path, data_only=True, read_only=True)
    entity = guess_entity(path.name)
    ctl = new_control()
    per_sheet: List[List[List[Any]]] = []
    for ws in wb.worksheets:
        # "Pending payment" sheets are payments not yet approved/due for this
        # batch - they must not be counted into the current Payment List (this
        # mirrors the same exclusion the legacy tool made by sheet name). They
        # can also be enormous (seen at 1,000,000+ nominal rows in a real
        # file), so skipping by name also avoids scanning that for nothing.
        if "pending" in norm(ws.title):
            ctl["pending_sheets"] += 1
            continue
        per_sheet.append(parse_sheet(ws, ctl))
    deduped = dedupe_across_sheets(per_sheet)
    ctl["dedup_dropped"] = sum(len(sr) for sr in per_sheet) - len(deduped)
    return [row + [entity, path.name] for row in deduped], ctl


def scan_input_files(input_dir: Path) -> List[Path]:
    files = [
        p for p in input_dir.rglob("*")
        if p.is_file() and p.suffix.lower() in SUPPORTED_SUFFIXES and not p.name.startswith("~$")
    ]
    return sorted(files, key=lambda p: (str(p.parent).lower(), p.name.lower()))


# A real payment-list file normally carries a few months of still-pending
# invoices (this batch's + carried-forward overdue ones) — that is expected
# and must not be filtered out. Some source files are instead a running
# multi-year log (every payment ever made, most already settled long ago)
# rather than a single batch; only THOSE are period-filtered, and only when
# --period is explicitly given, so a normal file's legitimate multi-month
# spread of pending invoices is never touched.
RUNNING_LOG_SPAN_MONTHS = 6


def _month_index(date_str: Any) -> Optional[int]:
    m = re.match(r"^(\d{4})-(\d{2})-\d{2}$", str(date_str or ""))
    if not m:
        return None
    return int(m.group(1)) * 12 + int(m.group(2))


def _file_date_span_months(rows: List[List[Any]]) -> int:
    # row[3] = Due Date(Value date), row[0] = Due Date(Payment date).
    idxs = [i for i in (_month_index(r[3]) or _month_index(r[0]) for r in rows) if i is not None]
    return (max(idxs) - min(idxs)) if idxs else 0


def _row_in_period(row: List[Any], period: str) -> bool:
    date_str = str(row[3] or row[0] or "")
    return date_str.startswith(period)


def _payment_date_key(row: List[Any]) -> str:
    """The row's payment date as YYYY-MM-DD text ('' when blank/unparseable)."""
    s = str(row[0] or "")[:10]
    return s if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s) else ""


# ===== Payment-run day (đi lệnh) rules — MUST mirror the dashboard's
# isRunDay()/monthEndLock()/monthStartLockDays() (RUN_WEEKDAY etc. in the
# Processing Calendar tab). Run day = Wednesday, weekly, except during the
# two book-closing windows:
# - month-end lock: (last calendar day − 3) through the last day;
# - month-start lock: the first 4 WORKING days of the month (weekends don't
#   consume the count).
RUN_WEEKDAY = 2  # Python weekday(): Monday=0 → Wednesday=2
MONTH_END_LOCK_DAYS = 3
MONTH_START_LOCK_WDS = 4
# AP teams sometimes enter the payday 1–2 days before the run day (a file
# prepared on Tuesday for Wednesday's run carries Tuesday's date). Rows
# within this lookback still belong to the batch, and their payday is
# normalized to the run day — treasury rule: payday MUST equal the run day.
BATCH_DATE_LOOKBACK_DAYS = 2


def _in_month_end_lock(d: date) -> bool:
    nxt = date(d.year + (d.month == 12), 1 if d.month == 12 else d.month + 1, 1)
    eom = nxt - timedelta(days=1)
    return (eom - timedelta(days=MONTH_END_LOCK_DAYS)) <= d <= eom


def _in_month_start_lock(d: date) -> bool:
    cur, count = d.replace(day=1), 0
    while count < MONTH_START_LOCK_WDS and cur.month == d.month:
        if cur.weekday() < 5:
            if cur == d:
                return True
            count += 1
        cur += timedelta(days=1)
    return False


def is_run_day(d: date) -> bool:
    return d.weekday() == RUN_WEEKDAY and not _in_month_end_lock(d) and not _in_month_start_lock(d)


def nearest_run_day(d: date) -> date:
    """The run day closest to `d` (ties go forward). Normal use is parsing ON
    the run day, so this is only a safety net for off-day runs."""
    for offset in range(0, 22):
        for cand in (d + timedelta(days=offset), d - timedelta(days=offset)):
            if is_run_day(cand):
                return cand
    return d  # unreachable in practice


def filter_batch_date(file_rows: List[List[Any]], batch_date: str, filename: str) -> List[List[Any]]:
    """Keep only this run's payment orders, and force payday = run day.

    Every source file is a running list where earlier batches' rows stay in
    place, so only rows dated at the run day — or up to
    BATCH_DATE_LOOKBACK_DAYS before it (entered on the prep day) — belong to
    this batch. The payment date is the output's first column, which
    payment_date_from_row() already maps per template — VNGSING's "Date"
    column, the standard files' "Due Date (Payment date)" column. Rows with
    a blank or unparseable payment date are old/undated history and drop.

    Kept rows' payment date is NORMALIZED to the run day: treasury rule is
    that the payday must equal the run day, so a row entered 1–2 days early
    comes out dated the run day, not its entry day.
    """
    target = date.fromisoformat(batch_date)
    window_start = (target - timedelta(days=BATCH_DATE_LOOKBACK_DAYS)).isoformat()
    kept, aligned = [], 0
    for r in file_rows:
        key = _payment_date_key(r)
        if key and window_start <= key <= batch_date:
            if key != batch_date:
                r = list(r)
                r[0] = batch_date
                aligned += 1
            kept.append(r)
    dates = sorted({_payment_date_key(r) for r in file_rows if _payment_date_key(r)})
    newest = dates[-1] if dates else "none"
    note = f", {aligned} aligned to the run day" if aligned else ""
    print(f"  {filename}: run day {batch_date} (window {window_start}..{batch_date}) -> kept {len(kept)} of {len(file_rows)} row(s){note} (newest date in file: {newest})")
    return kept


def build_rows(input_dir: Path, period: Optional[str] = None, batch_date: Optional[str] = None) -> Tuple[List[List[Any]], List[Dict[str, Any]]]:
    out: List[List[Any]] = []
    controls: List[Dict[str, Any]] = []
    for file in scan_input_files(input_dir):
        file_rows: List[List[Any]] = []
        ctl = new_control()
        try:
            if file.suffix.lower() == ".csv":
                entity = guess_entity(file.name)
                rows = load_rows_from_csv(file)
                if not rows:
                    continue
                width = main_table_width(rows[0])
                headers = [str(h or "").strip() for h in rows[0][:width]]
                skip_idx = summary_scan_skip_columns(headers)
                for raw in rows[1:]:
                    parsed, reason = extract_row(list(raw[:width]), headers, skip_idx)
                    tally(ctl, reason)
                    if parsed is not None:
                        file_rows.append(parsed + [entity, file.name])
            else:
                file_rows, ctl = parse_xlsx(file)
        except Exception as exc:
            # Keep going so one bad file cannot break the batch — but say so.
            # This used to swallow the exception silently, which meant a file
            # that failed to parse was indistinguishable from a file with no
            # payments in it.
            print(f"  !! {file.name}: SKIPPED - {type(exc).__name__}: {exc}")
            ctl.update({"file": file.name, "entity": guess_entity(file.name), "status": "FILE FAILED"})
            controls.append(ctl)
            continue

        if period and _file_date_span_months(file_rows) > RUNNING_LOG_SPAN_MONTHS:
            before = len(file_rows)
            file_rows = [r for r in file_rows if _row_in_period(r, period)]
            ctl["period_dropped"] = before - len(file_rows)
            print(f"  {file.name}: spans more than {RUNNING_LOG_SPAN_MONTHS} months of data -> filtered to {period}: kept {len(file_rows)} of {before} row(s)")
        if batch_date:
            before_w = len(file_rows)
            file_rows = filter_batch_date(file_rows, batch_date, file.name)
            ctl["window_dropped"] = before_w - len(file_rows)
        ctl["kept"] = len(file_rows)
        ccy_totals: Dict[str, float] = {}
        for r in file_rows:
            ccy = str(r[6] or "?").strip() or "?"
            ccy_totals[ccy] = ccy_totals.get(ccy, 0.0) + float(r[5] or 0)
        ctl["ccy"] = " · ".join(f"{v:,.2f} {c}" for c, v in sorted(ccy_totals.items()))
        # Completeness equations — every candidate row must be accounted for.
        eq1 = ctl["candidates"] == ctl["mapped"] + ctl["repeated_header"] + ctl["summary_row"] + ctl["no_amount"] + ctl["no_reference"]
        eq2 = ctl["kept"] == ctl["mapped"] - ctl["dedup_dropped"] - ctl["period_dropped"] - ctl["window_dropped"]
        ctl["file"] = file.name
        ctl["entity"] = guess_entity(file.name)
        ctl["status"] = "RECONCILED" if (eq1 and eq2) else "CHECK"
        controls.append(ctl)
        out.extend(file_rows)
    return out, controls


def autosize(ws) -> None:
    for col_cells in ws.columns:
        max_len = 0
        col_letter = col_cells[0].column_letter
        for c in col_cells:
            if c.value is None:
                continue
            max_len = max(max_len, len(str(c.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 42)


def write_output(output_path: Path, rows: List[List[Any]], input_dir: Path,
                 controls: Optional[List[Dict[str, Any]]] = None) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "PAYMENT_LIST"
    ws.append(OUTPUT_HEADERS)
    for row in rows:
        ws.append(row)

    if controls:
        ctl_ws = wb.create_sheet("PARSE_CONTROL")
        ctl_ws.append([
            "Source File", "Entity", "Candidate rows", "Kept (this run)",
            "Outside run window", "Restated on other sheet", "Summary/total rows",
            "No positive amount", "No reference", "Repeated headers",
            "Period-filtered", "Blank/formatting rows", "Pending sheets skipped",
            "Kept amount by currency", "Status",
        ])
        for c in controls:
            ctl_ws.append([
                c.get("file", "?"), c.get("entity", "?"), c["candidates"], c["kept"],
                c["window_dropped"], c["dedup_dropped"], c["summary_row"],
                c["no_amount"], c["no_reference"], c["repeated_header"],
                c["period_dropped"], c["blank"], c["pending_sheets"],
                c.get("ccy", ""), c.get("status", "?"),
            ])
        ctl_ws.append([])
        ctl_ws.append(["Control rule: Candidate rows = Kept + every exclusion column (window, restated, summary, no amount, no reference, repeated headers, period). Any file whose equation does not close is flagged CHECK and must be re-parsed or investigated before the batch is used."])

    readme = wb.create_sheet("README")
    readme.append(["Field", "Value"])
    readme.append(["Input folder", str(input_dir)])
    readme.append(["Template", "No. column removed"])
    readme.append(["Rule", "Only the left-most main payment table is parsed; subtotal / total / summary rows are skipped."])
    readme.append(["Row order", "Preserved from source files and source sheets."])
    readme.append(["Amount rule", "Rows without a real positive Amount are ignored."])

    for sheet in wb.worksheets:
        sheet.freeze_panes = "A2"
        autosize(sheet)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Parse AP payment lists into a clean template workbook.")
    parser.add_argument("--input", required=True, help="Folder containing AP payment-list files")
    parser.add_argument("--output", required=True, help="Output .xlsx file path")
    parser.add_argument(
        "--period",
        default=None,
        help=(
            "Optional YYYY-MM filter, e.g. 2026-08. Only applied to a source "
            f"file whose own rows span more than {RUNNING_LOG_SPAN_MONTHS} months "
            "(a running multi-year log, not a single batch) - that file's rows "
            "outside the given month are dropped. Files with a normal, narrower "
            "date spread (this batch's pending invoices) are never filtered, "
            "with or without this flag."
        ),
    )
    parser.add_argument(
        "--date",
        default="today",
        help=(
            "Payment-date filter: keep ONLY rows whose payment date equals "
            "this day - the source files are running lists where earlier "
            "batches' rows stay in place, so the run day's rows are the only "
            "ones that belong to this batch. Accepts YYYY-MM-DD, 'today' "
            "(default: the day you run the parser), or 'all' to disable and "
            "keep every row like before."
        ),
    )
    args = parser.parse_args()

    input_dir = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    if not input_dir.exists() or not input_dir.is_dir():
        raise SystemExit(f"Input folder not found: {input_dir}")

    if args.period and not re.fullmatch(r"\d{4}-\d{2}", args.period):
        raise SystemExit(f"--period must be in YYYY-MM format, got: {args.period}")

    batch_date: Optional[str]
    if args.date.lower() == "all":
        batch_date = None
    elif args.date.lower() == "today":
        today = date.today()
        run = today if is_run_day(today) else nearest_run_day(today)
        if run != today:
            print(f"NOTE: today ({today.isoformat()}) is not a run day per the run-day rule "
                  f"(Wednesday, excluding month-start/month-end closing windows) — using the nearest run day instead.")
        batch_date = run.isoformat()
    elif re.fullmatch(r"\d{4}-\d{2}-\d{2}", args.date):
        batch_date = args.date
        given = date.fromisoformat(batch_date)
        if not is_run_day(given):
            print(f"WARNING: {batch_date} is NOT a run day per the run-day rule "
                  f"(Wednesday, excluding month-start/month-end closing windows). "
                  f"Proceeding with it as given — paydays will be normalized to this date.")
    else:
        raise SystemExit(f"--date must be YYYY-MM-DD, 'today' or 'all', got: {args.date}")

    if batch_date:
        print(f"Batch run day: {batch_date} — keeping rows dated {batch_date} or up to "
              f"{BATCH_DATE_LOOKBACK_DAYS} day(s) before it (entered on the prep day); their payday is "
              f"normalized to the run day. Use --date all to keep everything.")

    rows, controls = build_rows(input_dir, args.period, batch_date)
    write_output(output_path, rows, input_dir, controls)

    print(f"Created: {output_path}")
    print(f"Rows:    {len(rows)}")
    ok = sum(1 for c in controls if c.get("status") == "RECONCILED")
    print(f"PARSE CONTROL: {ok}/{len(controls)} file(s) reconciled — every candidate row is accounted for "
          f"(kept + excluded-by-reason = scanned). Detail: PARSE_CONTROL sheet in the output workbook.")
    for c in controls:
        if c.get("status") != "RECONCILED":
            print(f"  !! CHECK {c.get('file')}: candidates={c['candidates']} kept={c['kept']} — equation does not close, investigate before using this batch")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
