"""
Secure Paper CoD processor — 66 Goulburn Street, Sydney.

Reads Certificate of Destruction (and related) PDFs from a tenant's evidence
folder, deduplicates against what's already logged (by Service Docket # /
Work order #), and writes new rows into the tenant's tab and the Evidence
Tracker. Consolidated and NABERS are then fully REBUILT from the tenant tabs
(they are pure derived views — never hand-edit them, they get wiped and
regenerated every run). Any row that existed in Consolidated/NABERS before a
rebuild but didn't come back out of the tenant tabs (e.g. someone deleted a
row in a tenant tab) is written to the append-only "Sync Flags" tab so it's
never silently lost.

Deterministic by design: PDF text extraction uses format-specific parsers
(one per contractor template seen so far). If a PDF doesn't match a known
template, or has too little extractable text (likely a scan needing OCR),
it is NOT guessed at — it's flagged in the Evidence Tracker for manual
review, and printed to the console so it can be handled (either read
manually, or a new parser added below).

Usage:
    python secure_paper_processor.py --tenant "Property NSW"
    python secure_paper_processor.py --tenant "Property NSW" --dry-run
    python secure_paper_processor.py --all
"""

import argparse
import glob
import os
import re
import shutil
import sys
from collections import Counter
from datetime import datetime

import openpyxl
import pdfplumber
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

# Portable: derived from where this script itself lives, not a hardcoded
# per-machine path. Works for any colleague who has this same "Secure Paper"
# folder synced locally (e.g. via OneDrive/SharePoint), regardless of their
# Windows username, as long as the script stays inside that folder.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EVIDENCE_DIR = os.path.join(BASE_DIR, "1. Secure paper Evidence", "Tenants")
LOG_DIR = os.path.join(BASE_DIR, "2. Visual Coding Information")
MASTER_DIR = os.path.join(BASE_DIR, "3. Secure Paper Masterfile")
BACKUP_DIR = os.path.join(MASTER_DIR, "Back Ups")
UPDATED_DIR = os.path.join(MASTER_DIR, "Updated")
# 2026-09-17: one canonical file from now on. Old version-numbered files
# (v3/v4/v7/v8/"(1)") were two silently diverging lineages - this script was
# writing to "(1).xlsx" while all the real reconciliation work (Bin Tracker,
# BT_BaseBuilding Gap, Pending Queries, etc.) only ever went into the
# hand-versioned files. All of that is now merged into ONE file in Updated/;
# every superseded file was moved to Back Ups/, nothing deleted.
MASTER_FILE = os.path.join(UPDATED_DIR, "Secure_Paper_Bin_Tally_PerTenant.xlsx")

SITE_ADDRESS = "66 Goulburn Street, Sydney, NSW"
SITE_PREMISES_ID = 2235

# ---------------------------------------------------------------------------
# Tenant / folder / tab configuration
# ---------------------------------------------------------------------------

# Evidence folder name -> workbook tab name.
# NCAT has no folder/tab of its own: Iron Mountain's account for the
# Property NSW tenancy is invoiced under the name "NSW CIVIL & ADMINISTRATIVE
# TRIBUNAL" / "NCAT" (per user confirmation, 2026-08-03) — it's a contractor
# label, not a separate tenant. Route it to Property NSW.
TENANT_FOLDER_TO_TAB = {
    "Property NSW": "Property NSW",
    "NCAT": "Property NSW",
    "BDW": "BDW",
    "Novotech": "Novotech",
    "TWP": "TWP",
    "Copyright Agency": "Copyright Agency",
    "Samuel Griffith Chmbrs": "Samuel Griffith Chmbrs",
    "Abbro Pty Ltd": "Abbro Pty Ltd",
    "Indigenous Business Aus": "Indigenous Business Aus",
    # Added 2026-09-17 (audit finding): these two folders/tabs already
    # existed but were never wired in, so --all silently skipped both
    # since inception. See Pending Queries row 6 / secure-paper-auditor.md.
    "Wesfarmers": "Wesfarmers",
    "Barker Evans": "Barker Evans",
}

# Company names that may legitimately appear ON a document for a given tab,
# other than the tab/tenant name itself. Anything else triggers a mismatch
# flag rather than being silently accepted.
KNOWN_ALIASES = {
    "Property NSW": {"NSW CIVIL & ADMINISTRATIVE TRIBUNAL", "NCAT"},
    "TWP": {"THE WORK PROJECT (SYDNEY) PTY LTD"},
    "Samuel Griffith Chmbrs": {"SAMUEL GRIFFITH CHAMBERS"},
    "Abbro Pty Ltd": {"BROUN ABRAHAMS BURREKET LAWYERS", "BROUN ABRAHAMS BURREKETE LAWYERS"},
    "Indigenous Business Aus": {"EVOLVE FM - IBA - SYDNEY", "IBA"},
    # Confirmed with Ram 17/09/2026: Wesfarmers' Shred-X certs are issued
    # under this trading name, address on the cert is 66 Goulburn St, Level 10.
    "Wesfarmers": {"API SERVICES AUSTRALIA PTY LTD"},
}

# Tenants whose evidence source has no floor/level info at all. Level comes
# from the manually-maintained Tenant tracker instead (never auto-read from
# there — user gives the number explicitly). 2026-08-18: Indigenous Business
# Aus' Shred-X Service History export has no address/level column whatsoever.
LEVEL_OVERRIDES = {
    "Indigenous Business Aus": 21,
}

# Tenant tabs that carry an extra "DoC provided" column (holds the Service
# Docket # for traceability). All other tenant tabs are the standard columns.
EXTRA_DOC_COLUMN_TABS = {"Property NSW"}

STANDARD_HEADERS = [
    "Pick-up date\n(dd/mm/yyyy)\n",
    "Site address\n(StreetNumber Street Name, Suburb, State, Postcode)",
    "Waste type",
    "Equipment",
    "Size (m3)",
    "Units collected",
    "Weight (if available)",
    "Tenant",
    "Level",
    "Waste services provider",
    "Verified (Yes/No)",
]
DOC_COLUMN_HEADER = "DoC provided"
UNITS_COL_STANDARD = 6  # "Units collected" position within STANDARD_HEADERS

# Samuel Griffith Chmbrs only: a blanket "tenant statement" style CoD covers
# a date range with no per-collection detail (no date, no docket, no bin
# count) — per user decision 2026-08-04, this never becomes a tenant-tab row.
# Instead it's flagged in Evidence Tracker and annotated onto whichever
# EXISTING tenant-tab rows fall inside the stated range, in an extra column
# specific to this one tab.
STATEMENT_NOTE_HEADER = "Statement note"
STATEMENT_NOTE_TEXT = "TPC Statement provided"
STATEMENT_NOTE_COL = len(STANDARD_HEADERS) + 1  # column L

EVIDENCE_TRACKER_HEADERS = [
    "Tenant",
    "Level",
    "Waste Contractor",
    "CoD",
    "Date",
    "Units Collected",
    "Complies with all information?",
    "Mapped out and incorporated in tenant tab? ",
    "Comment or issue",
]

CONSOLIDATED_HEADERS = STANDARD_HEADERS + [DOC_COLUMN_HEADER, "Calculated weight"]
CONSOLIDATED_UNITS_COL = UNITS_COL_STANDARD  # column F

NABERS_HEADERS = [
    "Pick-up date\n(dd/mm/yyyy)\n",
    "Premises id",
    "Site address\n(StreetNumber Street Name, Suburb, State, Postcode)",
    "Dock\n(can omit if only 1)",
    "Waste type",
    "Equipment",
    "Size (m3)",
    "Units collected",
    "Weight picked up (kg) \n(If available)",
    "Processing Facility sent to\n(Optional)",
    "Bins rejected due to contamination\n(Optional)",
    "Version: 4",
    "Calculated weight",
]
NABERS_UNITS_COL = 8  # column H

SYNC_FLAGS_HEADERS = ["Run timestamp", "Tab", "Issue", "Row values that vanished"]

# Fixed tab order for the front-of-book reference tabs (2026-08-03, per user).
# Tenant tabs keep their existing relative order and sit after these.
# "Sync Flags" always goes last, after every tenant tab.
FRONT_TAB_ORDER = [
    "NABERS",
    "D2 - WD Consolidated",
    "Consolidated Summary",
    "BT_BaseBuilding Gap",
    "Evidence Tracker",
    "Tenant tracker",
    "Method Details",
    "Instruction",
    "Tenant Name Mapping",
    "Pending Queries",
]

# Tabs the rebuild step must never treat as tenant data.
RESERVED_TABS = {
    "NABERS",
    "Instruction",
    "Tenant tracker",
    "Method Details",
    "Evidence Tracker",
    "Consolidated",
    "Sync Flags",
    # Added 2026-09-17 when the two masterfile lineages were merged - these
    # are hand/audit-maintained reconciliation tabs, never tenant data.
    "D2 - WD Consolidated",
    "Consolidated Summary",
    "BT_BaseBuilding Gap",
    "Tenant Name Mapping",
    "Pending Queries",
    "D1 - Bin Tracker",
}

CALC_WEIGHT_FORMULA = "=19.2*0.9*{col}{row}"


# ---------------------------------------------------------------------------
# PDF parsers — one per contractor template observed
# ---------------------------------------------------------------------------

def parse_bin_clauses(desc_text):
    """'Service of 5 x 240L Secure bins' -> [(5, 240)]. Multiple clauses if present."""
    return [
        (int(qty), int(size))
        for qty, size in re.findall(r"(\d+)\s*x\s*(\d+)\s*L", desc_text, re.I)
    ]


def parse_iron_mountain_single(text, filename):
    """Iron Mountain CoD: one A/C#, one Service Docket#, one Service Date,
    one Company Name, one Address, one Description of Service line."""
    m_head = re.search(
        r"A/C\s*#\s*Service Docket\s*#\s*Service Date\s*\n?\s*"
        r"(\S+)\s+(\S+)\s+(\d{2}/\d{2}/\d{2,4})",
        text,
    )
    m_company = re.search(r"Company Name\s*\n(.+)", text)
    m_address = re.search(r"Address\s*(.+)", text)
    m_desc = re.search(r"Description of Service\s*\n?(.+)", text)

    if not (m_head and m_company and m_address and m_desc):
        return None  # not this format

    account, docket, date_raw = m_head.groups()
    company = m_company.group(1).strip()
    address_raw = m_address.group(1).strip()
    desc = m_desc.group(1).strip()
    level_m = re.search(r"LEVEL\s*(\d+)", address_raw, re.I)
    bins = parse_bin_clauses(desc)

    rec = {
        "source_file": filename,
        "parser": "iron_mountain_single",
        "contractor": "Iron Mountain",
        "account_number": account,
        "docket": docket,
        "date_raw": date_raw,
        "company": company,
        "address_raw": address_raw,
        "level": int(level_m.group(1)) if level_m else None,
        "bins": bins,
        "issues": [],
    }
    if not level_m:
        rec["issues"].append(f"Could not find a LEVEL in address '{address_raw}'")
    if not bins:
        rec["issues"].append(f"Could not parse bin size/qty from description '{desc}'")
    return [rec]


def parse_iron_mountain_multi_docket(text, filename):
    """Iron Mountain CoD listing multiple Work order/Date rows under one
    account, with a single shared Description of Service applying to each
    work order (e.g. TWP's D14731 certificate)."""
    if "Work order" not in text or "Description of Service" not in text:
        return None

    m_acc_company = re.search(r"Company Name\s*\n(\S+)\s+(.+)", text)
    m_address = re.search(r"Address\s*(.+)", text)
    m_desc = re.search(r"Description of Service\s*\n?(.+)", text)
    if not (m_acc_company and m_desc):
        return None

    account, company = m_acc_company.groups()
    company = company.strip()
    address_raw = m_address.group(1).strip() if m_address else ""
    level_m = re.search(r"LEVEL\s*(\d+)", address_raw, re.I)
    desc = m_desc.group(1).strip()
    bins = parse_bin_clauses(desc)

    # Isolate the work-order table block so the leftover-line safety net
    # below doesn't pick up unrelated lines elsewhere in the document.
    m_table = re.search(r"Work order.*?Date\s*\n(.*?)\nDescription of Service", text, re.S)
    table_block = m_table.group(1) if m_table else text

    # Tolerate a PDF-extraction glitch that occasionally splits the year with
    # a stray space (e.g. "28/05/20 26" instead of "28/05/2026").
    rows = re.findall(r"^\s*(\d{6,8})\s+(\d{2}/\d{2}/\d{2}\s?\d{0,2})\s*$", table_block, re.M)
    if not rows:
        return None
    matched_dockets = {d for d, _ in rows}

    records = []
    for docket, date_raw in rows:
        rec = {
            "source_file": filename,
            "parser": "iron_mountain_multi_docket",
            "contractor": "Iron Mountain",
            "account_number": account,
            "docket": docket,
            "date_raw": date_raw.replace(" ", ""),
            "company": company,
            "address_raw": address_raw,
            "level": int(level_m.group(1)) if level_m else None,
            "bins": bins,
            "issues": [],
        }
        if not level_m:
            rec["issues"].append(f"Could not find a LEVEL in address '{address_raw}'")
        if not bins:
            rec["issues"].append(f"Could not parse bin size/qty from description '{desc}'")
        records.append(rec)

    # Safety net: any line that looks like a docket + trailing text but
    # didn't match the strict date pattern above must still be surfaced,
    # not silently dropped.
    for docket, rest in re.findall(r"^\s*(\d{6,8})\s+(.+)$", table_block, re.M):
        if docket in matched_dockets:
            continue
        records.append({
            "source_file": filename,
            "parser": "iron_mountain_multi_docket",
            "contractor": "Iron Mountain",
            "account_number": account,
            "docket": docket,
            "date_raw": rest.strip(),
            "company": company,
            "address_raw": address_raw,
            "level": int(level_m.group(1)) if level_m else None,
            "bins": bins,
            "issues": [f"Could not reliably parse a date from '{rest.strip()}' — flagging instead of guessing"],
        })
    return records


def parse_grace_table(text, filename):
    """Grace Worldwide destruction certificate: multi-page table of
    Workorder# / Account / Service Date / Action Code / Item Type / Qty."""
    if "Grace" not in text or "Destruction Certificate" not in text:
        return None

    m_acc = re.search(r"Account Number\s*(\S+)", text)
    m_addr = re.search(r"Collection Address\s*(.+)", text)
    address_raw = m_addr.group(1).strip() if m_addr else ""
    level_m = re.search(r"Level\s*(\d+)", address_raw, re.I)

    rows = re.findall(
        r"^(\d{6,8})\s+(.+?)\s+(\d{2}/\d{2}/\d{4})\s+(\S+)\s+(\d+)L\s+(\d+)\s*$",
        text,
        re.M,
    )
    if not rows:
        return None

    records = []
    for docket, account_name, date_raw, action_code, size, qty in rows:
        rec = {
            "source_file": filename,
            "parser": "grace_table",
            "contractor": "Grace Worldwide (Australia) Pty Ltd",
            "account_number": m_acc.group(1) if m_acc else account_name.strip(),
            "docket": docket,
            "date_raw": date_raw,
            "company": account_name.strip(),
            "address_raw": address_raw,
            "level": int(level_m.group(1)) if level_m else None,
            "bins": [(int(qty), int(size))],
            "issues": [],
        }
        if not level_m:
            rec["issues"].append(f"Could not find a Level in address '{address_raw}'")
        if action_code.lower() != "empty":
            rec["issues"].append(f"Unexpected action code '{action_code}' (expected 'Empty')")
        records.append(rec)
    return records


def parse_ordinal_date(s):
    """'20th of July 2026' -> datetime(2026, 7, 20)."""
    cleaned = re.sub(r"(\d+)(st|nd|rd|th)", r"\1", s).replace(" of ", " ")
    return datetime.strptime(cleaned.strip(), "%d %B %Y")


def parse_grace_single(text, filename):
    """Grace Worldwide destruction certificate, single-collection prose
    style (no Workorder table, no docket at all — e.g. Abbro Pty Ltd's
    Broun Abrahams Burreket cert): 'items collected on the <date>',
    Account Number, ACCOUNT NAME, Collection Address, ITEMS DESTROYED."""
    if "Grace" not in text or "Destruction Certificate" not in text:
        return None
    if "Workorder" in text or "Work order" in text:
        return None  # let parse_grace_table handle the table variant

    m_date = re.search(r"collected on the\s+(.+?)\s+have\b", text, re.I)
    m_acc = re.search(r"Account Number\s*(\S+)", text)
    m_company = re.search(r"ACCOUNT NAME\s*(.+)", text)
    m_addr = re.search(r"Collection Address\s*(.+)", text)
    m_items = re.search(r"ITEMS DESTROYED\s*(.+)", text)
    if not (m_date and m_acc and m_company and m_addr and m_items):
        return None

    account = m_acc.group(1).strip()
    try:
        date_dt = parse_ordinal_date(m_date.group(1))
    except ValueError:
        date_dt = None
    address_raw = m_addr.group(1).strip()
    level_m = re.search(r"Level\s*(\d+)", address_raw, re.I)
    bins = parse_bin_clauses(m_items.group(1))

    # No docket exists on this document at all — per user decision
    # 2026-08-18, use a synthetic account+date key so re-runs stay
    # idempotent instead of guessing or re-adding the same collection.
    docket = f"{account}-{date_dt.strftime('%d%m%Y')}" if date_dt else f"{account}-{filename}"

    rec = {
        "source_file": filename,
        "parser": "grace_single",
        "contractor": "Grace Worldwide (Australia) Pty Ltd",
        "account_number": account,
        "docket": docket,
        "date_raw": date_dt.strftime("%d/%m/%Y") if date_dt else "",
        "company": m_company.group(1).strip(),
        "address_raw": address_raw,
        "level": int(level_m.group(1)) if level_m else None,
        "bins": bins,
        "issues": [],
        "notes": ["Synthetic dedup key (account + date) — this document has no docket/certificate number"],
    }
    if not date_dt:
        rec["issues"].append(f"Could not parse date from 'collected on the {m_date.group(1)}'")
    if not level_m:
        rec["issues"].append(f"Could not find a Level in address '{address_raw}'")
    if not bins:
        rec["issues"].append(f"Could not parse bin size/qty from '{m_items.group(1)}'")
    return [rec]


def parse_shredx(text, filename):
    """Shred-X certificate: one or more 'Certificate Of Destruction' blocks
    (often one per page, e.g. Copyright Agency's 3-in-one-PDF), each with its
    own Certificate #, shared Account #, a qty+size Amount/Description line,
    and a 'Collected From <company> - <address> on <date>' line."""
    if "Certificate Of Destruction" not in text or "Certificate:" not in text:
        return None

    blocks = re.findall(
        r"Certificate:\s*(\S+)\s*\n"
        r"Account:\s*(\S+)\s*\n"
        r"Amount Description\s*\n"
        r"(\d+)\s+(\d+)L[^\n]*\n"
        r"Collected From\s+(.+?)\s+on\s*\n\s*(\d{1,2}/\d{1,2}/\d{4})",
        text,
    )
    if not blocks:
        return None

    records = []
    for docket, account, qty, size, collected_from, date_raw in blocks:
        if " - " in collected_from:
            company, address_raw = collected_from.split(" - ", 1)
        else:
            company, address_raw = collected_from, ""
        company = company.strip()
        address_raw = address_raw.strip()
        level_m = re.search(r"Level\s*(\d+)", address_raw, re.I)

        rec = {
            "source_file": filename,
            "parser": "shredx",
            "contractor": "Shred-X",
            "account_number": account,
            "docket": docket,
            "date_raw": date_raw,
            "company": company,
            "address_raw": address_raw,
            "level": int(level_m.group(1)) if level_m else None,
            "bins": [(int(qty), int(size))],
            "issues": [],
        }
        if not level_m:
            rec["issues"].append(f"Could not find a Level in address '{address_raw}'")
        records.append(rec)
    return records


def parse_complete_shredding(text, filename):
    """Complete Shredding Services certificate (e.g. Barker Evans Lawyers,
    added 2026-09-17): 'Client/Business: <name>' then an address line on the
    next line, 'Type of Destruction:' followed by one or more '<qty> x
    <size>L ...' clauses (possibly split across lines), 'Date Received:',
    'Receipt Docket Number:'. No account number on this template."""
    if "Complete Shredding Services" not in text:
        return None

    m_client = re.search(r"Client/Business:\s*(.+)", text)
    m_addr = re.search(r"Client/Business:\s*.+\n(.+)\n\s*Type of Destruction", text)
    m_type = re.search(r"Type of Destruction:\s*(.+?)\s*Date Received:", text, re.S)
    m_date = re.search(r"Date Received:\s*(\d{1,2}/\d{1,2}/\d{2,4})", text)
    m_docket = re.search(r"Receipt Docket Number:\s*(\S+)", text)

    if not (m_client and m_type and m_date and m_docket):
        return None

    company = m_client.group(1).strip()
    address_raw = m_addr.group(1).strip() if m_addr else ""
    level_m = re.search(r"Level\s*(\d+)", address_raw, re.I)
    bins = parse_bin_clauses(m_type.group(1))

    rec = {
        "source_file": filename,
        "parser": "complete_shredding",
        "contractor": "Complete Shredding Services",
        "account_number": None,
        "docket": m_docket.group(1),
        "date_raw": m_date.group(1),
        "company": company,
        "address_raw": address_raw,
        "level": int(level_m.group(1)) if level_m else None,
        "bins": bins,
        "issues": [],
    }
    if not level_m:
        rec["issues"].append(f"Could not find a Level in address '{address_raw}'")
    if not bins:
        rec["issues"].append(f"Could not parse bin size/qty from '{m_type.group(1)}'")
    return [rec]


PARSERS = [
    parse_iron_mountain_multi_docket,
    parse_iron_mountain_single,
    parse_grace_table,
    parse_grace_single,
    parse_shredx,
    parse_complete_shredding,
]


def detect_blanket_statement(text):
    """Detect a 'tenant statement' style CoD covering a date RANGE with no
    per-collection detail (no date, no docket, no bin count) — e.g. Samuel
    Griffith Chmbrs' Paper Company statement. Returns (date_from, date_to,
    contractor) or None."""
    norm = re.sub(r"\s+", " ", text)
    m_range = re.search(
        r"destroyed documents collected from\s+([A-Za-z]+ \d{4})\s+to\s+([A-Za-z]+ \d{4})",
        norm, re.I,
    )
    if not m_range:
        return None
    m_contractor = re.search(r"This Certificate confirms that (.+?),", norm)
    contractor = m_contractor.group(1).strip() if m_contractor else None

    from calendar import monthrange
    from_dt = datetime.strptime(m_range.group(1), "%B %Y")
    to_dt = datetime.strptime(m_range.group(2), "%B %Y")
    to_dt = to_dt.replace(day=monthrange(to_dt.year, to_dt.month)[1])
    return from_dt, to_dt, contractor


def extract_records(pdf_path):
    """Return (records, flag_reason, raw_text). flag_reason is None on a
    clean parse. raw_text lets the caller run extra checks (e.g. blanket
    statement detection) on documents that didn't match any parser."""
    filename = os.path.basename(pdf_path)
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    if len(text.strip()) < 30:
        return [], f"Almost no extractable text ({len(text.strip())} chars) — likely a scanned image needing OCR. Not auto-processed.", text

    for parser in PARSERS:
        records = parser(text, filename)
        if records is not None:
            return records, None, text

    return [], "No known CoD template matched this document's layout. Needs a new parser or manual review.", text


def extract_records_from_excel(xlsx_path):
    """Shred-X 'Service History' export (e.g. Indigenous Business Aus' Jarrah
    workbook): one real collection per row, Docket/Job Date/Qty/Service Name
    columns. Returns (records, flag_reason, raw_text=None) — same shape as
    extract_records() so the caller doesn't need to branch on file type."""
    filename = os.path.basename(xlsx_path)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    try:
        if "Service History" not in wb.sheetnames:
            return [], "Excel evidence file has no 'Service History' sheet — needs manual review.", None

        ws = wb["Service History"]
        header = None
        records = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                header = [str(h).strip() if h else "" for h in row]
                continue
            if not row or all(v is None for v in row):
                continue
            rowd = dict(zip(header, row))
            if rowd.get("Invoice No.") == "Grand Total" or not rowd.get("CAN"):
                continue

            docket = rowd.get("Docket")
            job_date = rowd.get("Job Date")
            qty = rowd.get("Qty")
            service_name = rowd.get("Service Name") or ""
            weight = rowd.get("Weight (KG)")
            site = (rowd.get("Site") or "").strip()
            size_m = re.search(r"(\d+)L", service_name)

            issues = []
            if not docket:
                issues.append("No Docket # in Service History row")
            if not job_date:
                issues.append("No Job Date in Service History row")
            if not size_m or not qty:
                issues.append(f"Could not parse bin size/qty from Service Name '{service_name}' / Qty '{qty}'")

            date_raw = job_date.strftime("%d/%m/%Y") if hasattr(job_date, "strftime") else str(job_date or "")

            records.append({
                "source_file": filename,
                "parser": "shredx_excel_service_history",
                "contractor": "Shred-X",
                "account_number": rowd.get("CAN"),
                "docket": docket or f"NODOCKET-{filename}-row{i+1}",
                "date_raw": date_raw,
                "company": site,
                "address_raw": "",
                "level": None,
                "bins": [(int(qty), int(size_m.group(1)))] if size_m and qty else [],
                "weight_kg": weight if weight and weight > 0 else None,
                "issues": issues,
            })

        if not records:
            return [], "Excel file recognized but 'Service History' sheet had no data rows.", None
        return records, None, None
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Date parsing / normalisation
# ---------------------------------------------------------------------------

def parse_date(date_raw):
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(date_raw, fmt)
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------------------
# Workbook helpers — backup, tab creation, dedup, tenant/evidence writes
# ---------------------------------------------------------------------------

def backup_workbook():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    # 2026-09-17: Ram asked not to create a fresh backup on every single
    # run. One per calendar day is enough safety margin - reuse today's
    # if one already exists instead of adding another.
    today_stamp = datetime.now().strftime("%Y%m%d")
    existing_today = sorted(glob.glob(
        os.path.join(BACKUP_DIR, f"Secure_Paper_Bin_Tally_PerTenant_backup_{today_stamp}_*.xlsx")
    ))
    if existing_today:
        return existing_today[-1]
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    backup_path = os.path.join(
        BACKUP_DIR, f"Secure_Paper_Bin_Tally_PerTenant_backup_{stamp}.xlsx"
    )
    shutil.copy2(MASTER_FILE, backup_path)
    return backup_path


def first_empty_row(ws, key_col=1):
    r = ws.max_row
    while r >= 1 and ws.cell(row=r, column=key_col).value in (None, ""):
        r -= 1
    return r + 1


def reorder_tabs(wb):
    """Enforce FRONT_TAB_ORDER, then tenant tabs in their existing relative
    order, then anything else (e.g. Sync Flags) last. Runs every time — tab
    order should never need manual fixing again."""
    existing = wb.sheetnames
    front = [n for n in FRONT_TAB_ORDER if n in existing]
    tenant_tabs = [n for n in existing if n not in RESERVED_TABS]
    trailing = [n for n in existing if n not in front and n not in tenant_tabs]
    new_order = front + tenant_tabs + trailing
    if new_order != existing:
        wb._sheets = [wb[name] for name in new_order]
        return True
    return False


def annotate_statement_note(wb, tab_name, date_from, date_to, log_lines):
    """Write STATEMENT_NOTE_TEXT into column L of every EXISTING row in
    tab_name whose Pick-up date falls within [date_from, date_to]. Never
    touches Verified — a blanket statement isn't per-collection evidence.
    Idempotent: skips rows that already carry the note."""
    ws = wb[tab_name]
    if ws.cell(row=1, column=STATEMENT_NOTE_COL).value is None:
        ws.cell(row=1, column=STATEMENT_NOTE_COL, value=STATEMENT_NOTE_HEADER)

    annotated = 0
    for r in range(2, ws.max_row + 1):
        v1 = ws.cell(row=r, column=1).value
        if v1 in (None, "") or str(v1).strip().lower().startswith("example"):
            continue
        row_date = parse_date(str(v1).strip())
        if not row_date or not (date_from <= row_date <= date_to):
            continue
        if ws.cell(row=r, column=STATEMENT_NOTE_COL).value == STATEMENT_NOTE_TEXT:
            continue
        ws.cell(row=r, column=STATEMENT_NOTE_COL, value=STATEMENT_NOTE_TEXT)
        annotated += 1

    log_lines.append(
        f"  STATEMENT NOTE: annotated {annotated} existing row(s) in '{tab_name}' "
        f"covering {date_from.strftime('%d/%m/%Y')}–{date_to.strftime('%d/%m/%Y')} "
        f"with '{STATEMENT_NOTE_TEXT}' (Verified left untouched)"
    )
    return annotated


def ensure_tab(wb, name, headers, index=None):
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name, index) if index is not None else wb.create_sheet(name)
    for c, header in enumerate(headers, start=1):
        ws.cell(row=1, column=c, value=header)
    return ws


def existing_dockets(wb, tab_name):
    """Union of Service Docket #s already present in the Evidence Tracker
    and (if applicable) the tenant tab's DoC-provided column."""
    dockets = set()

    et = wb["Evidence Tracker"]
    for r in range(2, et.max_row + 1):
        v = et.cell(row=r, column=4).value  # CoD column
        if v:
            dockets.add(str(v).strip())

    if tab_name in EXTRA_DOC_COLUMN_TABS and tab_name in wb.sheetnames:
        ws = wb[tab_name]
        doc_col = len(STANDARD_HEADERS) + 1
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=doc_col).value
            if v:
                dockets.add(str(v).strip())

    return dockets


def find_header_col(ws, header_name):
    """1-based column number whose row-1 header matches header_name
    (case/whitespace-insensitive), or None if that tab has no such column.
    Used so extra columns (DoC/Reference, Method, Data Source) get filled
    on whatever tenant tabs actually have them, instead of one hardcoded
    tab special-case."""
    target = header_name.strip().lower()
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v and str(v).strip().lower() == target:
            return c
    return None


def append_tenant_tab_row(wb, tab_name, level, size_m3, qty, contractor, date_dt, docket, weight_kg=None):
    ws = wb[tab_name]
    row = first_empty_row(ws)
    values = [
        date_dt.strftime("%d/%m/%Y"),
        SITE_ADDRESS,
        "Secure paper",
        "Mobile bin",
        size_m3,
        qty,
        weight_kg,
        tab_name,
        level,
        contractor,
        "Yes",
    ]
    for c, v in enumerate(values, start=1):
        ws.cell(row=row, column=c, value=v)
    doc_col = find_header_col(ws, "DoC / Reference") or find_header_col(ws, "DoC provided")
    if doc_col:
        ws.cell(row=row, column=doc_col, value=docket)
    method_col = find_header_col(ws, "Method")
    if method_col:
        ws.cell(row=row, column=method_col, value="Method 1, bin weights (contractor)")
    source_col = find_header_col(ws, "Data Source")
    if source_col:
        ws.cell(row=row, column=source_col, value="Contractor Report")
    return row


def append_evidence_tracker_row(wb, tenant, level, contractor, docket, date_dt, qty, complies, mapped, comment):
    ws = wb["Evidence Tracker"]
    row = first_empty_row(ws)
    values = [
        tenant,
        level,
        contractor,
        docket,
        date_dt.strftime("%d/%m/%Y") if date_dt else None,
        qty,
        complies,
        mapped,
        comment,
    ]
    for c, v in enumerate(values, start=1):
        ws.cell(row=row, column=c, value=v)
    return row


# ---------------------------------------------------------------------------
# Consolidated / NABERS — fully rebuilt from tenant tabs every run
# ---------------------------------------------------------------------------

def snapshot_data_rows(ws, ncols):
    """Existing rows 2..max_row as tuples (first ncols columns), skipping
    blank rows. Used to diff against the post-rebuild content."""
    rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ncols + 1)]
        if all(v in (None, "") for v in vals):
            continue
        rows.append(tuple(vals))
    return rows


def clear_data_rows(ws):
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)


def flag_vanished_rows(wb, tab_name, old_rows, new_rows, log_lines):
    """Anything present in old_rows more times than in new_rows is written
    to the Sync Flags tab (append-only, never rebuilt) plus the run log."""
    missing = Counter(old_rows) - Counter(new_rows)
    if not sum(missing.values()):
        return 0

    ws = ensure_tab(wb, "Sync Flags", SYNC_FLAGS_HEADERS)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    count = 0
    for row_tuple, n in missing.items():
        for _ in range(n):
            row = first_empty_row(ws)
            ws.cell(row=row, column=1, value=stamp)
            ws.cell(row=row, column=2, value=tab_name)
            ws.cell(row=row, column=3, value="Row was present before rebuild, no tenant tab produced it this time")
            ws.cell(row=row, column=4, value=" | ".join("" if v is None else str(v) for v in row_tuple))
            count += 1
    log_lines.append(f"  SYNC FLAG: {count} row(s) that used to be in '{tab_name}' didn't come back out of the tenant tabs — see Sync Flags tab")
    return count


def rebuild_consolidated(wb, log_lines):
    ws = ensure_tab(wb, "Consolidated", CONSOLIDATED_HEADERS)
    ncols_key = len(CONSOLIDATED_HEADERS) - 1  # exclude Calculated weight formula col
    old_rows = snapshot_data_rows(ws, ncols_key)
    clear_data_rows(ws)

    new_rows = []
    for tab_name in wb.sheetnames:
        if tab_name in RESERVED_TABS:
            continue
        tws = wb[tab_name]
        has_doc_col = tab_name in EXTRA_DOC_COLUMN_TABS
        for r in range(2, tws.max_row + 1):
            v1 = tws.cell(row=r, column=1).value
            if v1 in (None, "") or str(v1).strip().lower().startswith("example"):
                continue
            vals = [tws.cell(row=r, column=c).value for c in range(1, len(STANDARD_HEADERS) + 1)]
            docket = tws.cell(row=r, column=len(STANDARD_HEADERS) + 1).value if has_doc_col else None
            new_rows.append(tuple(vals + [docket]))

    for row_vals in new_rows:
        row = first_empty_row(ws)
        for c, v in enumerate(row_vals, start=1):
            ws.cell(row=row, column=c, value=v)
        col_letter = get_column_letter(CONSOLIDATED_UNITS_COL)
        ws.cell(row=row, column=len(CONSOLIDATED_HEADERS), value=CALC_WEIGHT_FORMULA.format(col=col_letter, row=row))

    flag_vanished_rows(wb, "Consolidated", old_rows, new_rows, log_lines)
    log_lines.append(f"  Consolidated rebuilt: {len(new_rows)} row(s) from {len([s for s in wb.sheetnames if s not in RESERVED_TABS])} tenant tab(s)")
    return new_rows


def rebuild_nabers(wb, consolidated_rows, log_lines):
    ws = ensure_tab(wb, "NABERS", NABERS_HEADERS, index=0)
    ncols_key = len(NABERS_HEADERS) - 1  # exclude Calculated weight formula col
    old_rows = snapshot_data_rows(ws, ncols_key)
    clear_data_rows(ws)

    verified_idx = len(STANDARD_HEADERS) - 1  # 0-based index of "Verified (Yes/No)"
    new_rows = []
    for row_vals in consolidated_rows:
        if str(row_vals[verified_idx]).strip().lower() != "yes":
            continue
        date_, site_address, waste_type, equipment, size_m3, units, weight = row_vals[0:7]
        nabers_row = (
            date_, SITE_PREMISES_ID, site_address, None,
            waste_type, equipment, size_m3, units, weight,
            None, None, None,
        )
        new_rows.append(nabers_row)

    for row_vals in new_rows:
        row = first_empty_row(ws)
        for c, v in enumerate(row_vals, start=1):
            ws.cell(row=row, column=c, value=v)
        col_letter = get_column_letter(NABERS_UNITS_COL)
        ws.cell(row=row, column=len(NABERS_HEADERS), value=CALC_WEIGHT_FORMULA.format(col=col_letter, row=row))

    flag_vanished_rows(wb, "NABERS", old_rows, new_rows, log_lines)
    log_lines.append(f"  NABERS rebuilt: {len(new_rows)} verified row(s)")
    return new_rows


# ---------------------------------------------------------------------------
# Main per-tenant processing
# ---------------------------------------------------------------------------

def process_tenant(folder_name, wb, log_lines):
    tab_name = TENANT_FOLDER_TO_TAB.get(folder_name)
    if tab_name is None and folder_name in wb.sheetnames:
        # 2026-09-17 audit fix: a folder with no explicit mapping but an
        # exact-name tab already in the workbook is auto-mapped rather than
        # silently skipped. A folder matching NO tab at all still stops
        # below and needs an explicit decision — this never guesses at an
        # unrelated tab name.
        tab_name = folder_name
        log_lines.append(
            f"NOTE: folder '{folder_name}' has no entry in TENANT_FOLDER_TO_TAB but an "
            f"exact-name tab exists — auto-mapped. Add an explicit entry for clarity."
        )
        print(f"[NOTE] '{folder_name}' auto-mapped to matching tab (no explicit dict entry).")
    if tab_name is None:
        log_lines.append(
            f"STOP: folder '{folder_name}' doesn't map to any known tab. "
            f"Skipping — needs a decision before this can be processed."
        )
        print(f"[STOP] '{folder_name}' has no tab mapping. Skipping.")
        return {"processed": 0, "added": 0, "already_logged": 0, "duplicate_evidence": 0, "flagged": 0}

    if tab_name not in wb.sheetnames:
        log_lines.append(f"STOP: target tab '{tab_name}' does not exist in the workbook. Skipping '{folder_name}'.")
        print(f"[STOP] Tab '{tab_name}' missing. Skipping '{folder_name}'.")
        return {"processed": 0, "added": 0, "already_logged": 0, "duplicate_evidence": 0, "flagged": 0}

    folder_path = os.path.join(EVIDENCE_DIR, folder_name)
    evidence_files = sorted(glob.glob(os.path.join(folder_path, "*.pdf"))) + sorted(
        glob.glob(os.path.join(folder_path, "*.xlsx"))
    )
    # Dockets already logged from a PREVIOUS run. Re-encountering one of
    # these this run is just idempotent re-processing of the same evidence —
    # skip silently, no new row anywhere. A "duplicate" is reserved for two
    # separate pieces of evidence for the same collection found in THIS run.
    pre_existing_dockets = existing_dockets(wb, tab_name)
    processed_this_run = set()

    counts = {"processed": 0, "added": 0, "already_logged": 0, "duplicate_evidence": 0, "flagged": 0}
    all_records = []

    for evidence_path in evidence_files:
        filename = os.path.basename(evidence_path)
        counts["processed"] += 1
        if evidence_path.lower().endswith(".xlsx"):
            records, flag_reason, raw_text = extract_records_from_excel(evidence_path)
        else:
            records, flag_reason, raw_text = extract_records(evidence_path)

        if flag_reason:
            # No real docket exists for these — use the filename as a
            # pseudo dedup key so re-running doesn't re-flag/re-annotate
            # the same file on every future run.
            pseudo_key = f"FILE:{filename}"
            if pseudo_key in pre_existing_dockets:
                log_lines.append(f"  ALREADY LOGGED (skipped, no write): {filename}")
                counts["already_logged"] += 1
                continue

            blanket = detect_blanket_statement(raw_text) if raw_text else None
            if blanket:
                date_from, date_to, contractor = blanket
                append_evidence_tracker_row(
                    wb, tab_name, None, contractor, pseudo_key, None, None,
                    complies="No", mapped="No",
                    comment=(
                        f"Blanket tenant statement ({date_from.strftime('%d/%m/%Y')}–{date_to.strftime('%d/%m/%Y')}), "
                        f"no per-collection date/docket/qty — does not meet NABERS evidence requirements as a "
                        f"standalone record. Annotated onto existing rows in that range instead. (file: {filename})"
                    ),
                )
                annotate_statement_note(wb, tab_name, date_from, date_to, log_lines)
                log_lines.append(f"  FLAGGED (blanket statement, not a collection record): {filename}")
                counts["flagged"] += 1
                processed_this_run.add(pseudo_key)
                continue

            log_lines.append(f"  FLAGGED (unreadable/unknown format): {filename} — {flag_reason}")
            append_evidence_tracker_row(
                wb, tab_name, None, None, pseudo_key, None, None,
                complies="No", mapped="No",
                comment=f"{flag_reason} (file: {filename})",
            )
            counts["flagged"] += 1
            processed_this_run.add(pseudo_key)
            continue

        for rec in records:
            all_records.append(rec)

    # Sort by parsed date for tidy chronological writing.
    def sort_key(r):
        d = parse_date(r["date_raw"])
        return d if d else datetime.max
    all_records.sort(key=sort_key)

    for rec in all_records:
        docket = rec["docket"]
        date_dt = parse_date(rec["date_raw"])
        level = rec["level"]
        contractor = rec["contractor"]
        company = rec["company"]
        bins = rec["bins"]
        weight_kg = rec.get("weight_kg")
        issues = list(rec["issues"])
        notes = list(rec.get("notes", []))

        if level is None and tab_name in LEVEL_OVERRIDES:
            level = LEVEL_OVERRIDES[tab_name]
            notes.append(f"Level {level} taken from Tenant tracker (source evidence has no floor/level info)")

        if not date_dt:
            issues.append(f"Unparseable date '{rec['date_raw']}'")

        alias_ok = (
            tab_name.upper() in company.upper()
            or company.upper() in KNOWN_ALIASES.get(tab_name, set())
        )
        if not alias_ok:
            issues.append(f"Company name on document ('{company}') doesn't match tenant '{tab_name}' or its known aliases")

        if docket in pre_existing_dockets and docket not in processed_this_run:
            # Already logged in a previous run — idempotent re-processing of
            # the same evidence, not a duplicate. No new row anywhere.
            log_lines.append(f"  ALREADY LOGGED (skipped, no write): {rec['source_file']} — docket {docket}")
            counts["already_logged"] += 1
            continue

        if docket in processed_this_run:
            # Genuine duplicate: a second piece of evidence for the same
            # collection found within this same run.
            append_evidence_tracker_row(
                wb, tab_name, level, contractor, docket, date_dt,
                sum(q for q, _ in bins) if bins else None,
                complies="Yes", mapped="No",
                comment=f"Duplicate evidence — Service Docket # {docket} already logged from another file this run. Source file: {rec['source_file']}.",
            )
            log_lines.append(f"  DUPLICATE EVIDENCE: {rec['source_file']} — docket {docket} also found in another file this run")
            counts["duplicate_evidence"] += 1
            continue

        if issues:
            append_evidence_tracker_row(
                wb, tab_name, level, contractor, docket, date_dt,
                sum(q for q, _ in bins) if bins else None,
                complies="No", mapped="No",
                comment="; ".join(issues) + f" (file: {rec['source_file']})",
            )
            log_lines.append(f"  FLAGGED: {rec['source_file']} (docket {docket}) — {'; '.join(issues)}")
            counts["flagged"] += 1
            processed_this_run.add(docket)
            continue

        # Clean record — may still be multiple rows if multiple bin sizes.
        for qty, size_l in bins:
            size_m3 = round(size_l / 1000, 2)
            append_tenant_tab_row(wb, tab_name, level, size_m3, qty, contractor, date_dt, docket, weight_kg=weight_kg)

        total_qty = sum(q for q, _ in bins)
        ok_comment = f"OK — added to '{tab_name}' tab. Source file: {rec['source_file']}."
        if notes:
            ok_comment += " " + "; ".join(notes) + "."
        append_evidence_tracker_row(
            wb, tab_name, level, contractor, docket, date_dt, total_qty,
            complies="Yes", mapped="Yes",
            comment=ok_comment,
        )
        log_lines.append(f"  ADDED: {rec['source_file']} (docket {docket}, {date_dt.strftime('%d/%m/%Y')}, {total_qty} bin(s), Level {level})")
        counts["added"] += 1
        processed_this_run.add(docket)

    return counts


# ---------------------------------------------------------------------------
# Run log
# ---------------------------------------------------------------------------

def write_run_log(log_lines, totals, dry_run, backup_path):
    os.makedirs(LOG_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    log_path = os.path.join(LOG_DIR, f"run_log_{stamp}.md")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"# Secure Paper processing run — {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n")
        f.write(f"Mode: {'DRY RUN (no changes saved)' if dry_run else 'LIVE'}\n\n")
        if backup_path:
            f.write(f"Backup created: `{backup_path}`\n\n")
        f.write("## Detail\n\n")
        for line in log_lines:
            f.write(line + "\n")
        f.write("\n## Totals\n\n")
        f.write(
            f"- PDFs processed: {totals['processed']}\n"
            f"- Rows added (new + compliant): {totals['added']}\n"
            f"- Already logged from a previous run (skipped, no write): {totals['already_logged']}\n"
            f"- Duplicate evidence within this run (flagged): {totals['duplicate_evidence']}\n"
            f"- Flagged for review: {totals['flagged']}\n"
        )
    return log_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tenant", help="Evidence folder name to process (e.g. 'Property NSW')")
    ap.add_argument("--all", action="store_true", help="Process every mapped tenant folder")
    ap.add_argument("--reorder-tabs", action="store_true", help="Only fix tab order (NABERS/Consolidated/Evidence Tracker/Tenant tracker/Method Details/Instruction, then tenant tabs), don't process any evidence")
    ap.add_argument("--dry-run", action="store_true", help="Don't save workbook or write log file changes")
    args = ap.parse_args()

    if not args.tenant and not args.all and not args.reorder_tabs:
        print("Specify --tenant \"<folder name>\", --all, or --reorder-tabs")
        sys.exit(1)

    backup_path = None
    if not args.dry_run:
        backup_path = backup_workbook()
        print(f"Backup: {backup_path}")

    wb = openpyxl.load_workbook(MASTER_FILE)

    if args.reorder_tabs:
        moved = reorder_tabs(wb)
        if not args.dry_run:
            wb.save(MASTER_FILE)
            print(f"Saved: {MASTER_FILE}")
        print("Tab order fixed." if moved else "Tab order already correct — nothing to do.")
        return

    if args.all:
        # 2026-09-17 audit fix: scan the real evidence directory rather than
        # only the folders TENANT_FOLDER_TO_TAB already knows about — a new
        # tenant folder (e.g. Wesfarmers, Barker Evans, both added at some
        # point without updating this dict) is never silently invisible to
        # a run again. Unmapped folders still get a STOP/flag decision
        # inside process_tenant(), they just aren't skipped before that.
        folders = sorted(
            d for d in os.listdir(EVIDENCE_DIR)
            if os.path.isdir(os.path.join(EVIDENCE_DIR, d))
        )
    else:
        folders = [args.tenant]
    # De-dupe folder aliases that route to the same tab but only process each
    # physical folder that actually exists.
    folders = [f for f in dict.fromkeys(folders) if os.path.isdir(os.path.join(EVIDENCE_DIR, f))]

    log_lines = []
    totals = {"processed": 0, "added": 0, "already_logged": 0, "duplicate_evidence": 0, "flagged": 0}
    for folder in folders:
        log_lines.append(f"\n### Folder: {folder}")
        counts = process_tenant(folder, wb, log_lines)
        for k in totals:
            totals[k] += counts[k]

    log_lines.append(
        "\n### NABERS / D2 - WD Consolidated / Consolidated Summary / BT_BaseBuilding Gap"
        "\n  Not auto-rebuilt (changed 2026-09-17) - these are hand/audit-maintained now that"
        "\n  real Base Building weight matching is part of the model. See CLAUDE.md."
    )

    if reorder_tabs(wb):
        log_lines.append("\nTab order was fixed to match the standard front-of-book order.")

    if not args.dry_run:
        wb.save(MASTER_FILE)
        print(f"Saved: {MASTER_FILE}")
    else:
        print("Dry run — workbook not saved.")

    log_path = write_run_log(log_lines, totals, args.dry_run, backup_path)
    print(f"Run log: {log_path}")
    print(totals)


if __name__ == "__main__":
    main()
