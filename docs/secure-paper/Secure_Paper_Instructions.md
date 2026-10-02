# Secure Paper — 66 Goulburn Street, Sydney

Automates turning Certificate of Destruction (CoD) PDFs from secure paper
contractors into rows in the tenant tabs and Evidence Tracker, without
double-counting a collection. As of 17/09/2026 there is exactly ONE
masterfile — see Folder layout below. Everything else (NABERS, D2 - WD
Consolidated, Consolidated Summary, BT_BaseBuilding Gap, Pending Queries)
is hand/audit-maintained, not touched by this script.

## Trigger phrases

**"Run secure paper for 66 Goulburn"** (or a close variant) — ingest new evidence:
1. `cd` into this folder.
2. Run `python secure_paper_processor.py --all --dry-run`.
3. Show the user the run log it prints the path to (in `2. Visual Coding Information\`).
4. If they confirm it looks right, run `python secure_paper_processor.py --all` (same command, no `--dry-run`) to go live.

**"Audit the secure paper masterfile"** / **"empezá a auditar"** (or a close variant) —
run the Secure Paper Auditor: read `secure-paper-auditor.md` in the "66 Goulburn"
claude.ai project first (the full checklist and non-negotiable rules live there,
not here), then work through it sequentially, one tenant at a time. Log findings
in Pending Queries (this workbook) and mirror them in that project doc. Never
edit tenant data without asking first; the only pre-authorised write is the
findings log itself. A scheduled reminder pings the user roughly every 3 months
to run this — it never runs unattended.

Don't make the user type any command themselves unless they ask for it directly.

## Folder layout

- `1. Secure paper Evidence\Tenants\<tenant>\` — raw CoD PDFs, one subfolder per tenant. New evidence just gets dropped in here. The script scans this directory for whatever folders actually exist every run (not a fixed list) — see `TENANT_FOLDER_TO_TAB` below.
- `2. Visual Coding Information\` — dated run logs (`run_log_<timestamp>.md`), auditable history of every run.
- `3. Secure Paper Masterfile\Updated\Secure_Paper_Bin_Tally_PerTenant.xlsx` — **the one live masterfile.** Always the current version; no version number in the filename any more, because history now lives in Back Ups by timestamp instead of in the filename.
- `3. Secure Paper Masterfile\Back Ups\` — timestamped backup, one per calendar day the workbook is touched (not one per run — changed 17/09/2026 at the user's request). Also holds every file superseded by the 17/09/2026 merge (the old `(1).xlsx` the script used to write to, and the old v3/v4/v7/v8 hand-versioned files) — kept for history, none of them are live any more, don't read from them for current state.
- `secure_paper_processor.py` — the script. Portable: it locates itself via `__file__`, so it works from any machine where this whole folder is synced (e.g. OneDrive), regardless of Windows username.

**Why there was ever more than one file**: before 17/09/2026, this script wrote to
`(1).xlsx` while a separate hand-maintained lineage (v3→v4→v7→v8) picked up Bin
Tracker onsite-weight data and the BT_BaseBuilding Gap/Pending Queries reconciliation
work that never fed back into `(1).xlsx`. That's how two real, evidenced Wesfarmers
and Barker Evans collections sat unprocessed indefinitely — the script's tenant
folder list was hardcoded and never updated, so it silently never even looked at
those two folders, and separately nobody was checking `(1).xlsx` against the file
actually used for NABERS reporting. Both problems are fixed: the script now scans
the real evidence directory every run, and there's only one file for it to write to.

## Requirements to run this on another machine

- Python 3.x
- `pip install openpyxl pdfplumber`
- This folder synced/available locally with the same internal structure (paths are all relative to the script's own location — nothing else to configure).

## What the script does

1. Backs up the workbook to `3. Secure Paper Masterfile\Back Ups\` (once per calendar day, reuses today's if one already exists).
2. Scans `1. Secure paper Evidence\Tenants\` for every folder that actually exists on disk (not a fixed list — a brand-new tenant folder is never silently invisible). Each folder maps to a tab via `TENANT_FOLDER_TO_TAB`; if it's not in that dict but a tab with the exact same name already exists in the workbook, it auto-maps to that tab (logged as a NOTE, add an explicit dict entry when convenient). A folder matching no tab at all still stops and flags — it never guesses at an unrelated tab.
3. For each PDF (and `.xlsx`, e.g. a Shred-X "Service History" export), extracts fields (service date, Service Docket #/Work order #, bin count/size, level, contractor) with a contractor-specific parser (see `PARSERS` — Iron Mountain single-record, Iron Mountain multi-docket-per-doc, Grace Worldwide table, Grace Worldwide single-collection prose (no docket — synthetic account+date key), Shred-X one-or-more-certs-per-doc, Shred-X Excel Service History, Complete Shredding Services (added 17/09/2026 for Barker Evans)). `LEVEL_OVERRIDES` covers tenants whose evidence has no floor info at all (level then comes from the manually-maintained Tenant tracker, entered explicitly by the user, never auto-read).
4. Dedup key is the Service Docket #/Work order # — never the account number (repeats across every doc for a tenant). Already-logged dockets from a past run are skipped silently (not a duplicate). A genuine duplicate is two files presenting the same docket within the *same* run — flagged in Evidence Tracker. Files with no docket at all (unreadable, unrecognized format, or a blanket statement) dedupe the same way using a `FILE:<filename>` pseudo-key.
5. New + complete → row in the tenant's own tab + Evidence Tracker. When writing a tenant-tab row, any column literally headered "DoC / Reference" or "DoC provided" gets the docket; any column headered "Method" gets "Method 1, bin weights (contractor)"; any column headered "Data Source" gets "Contractor Report" — looked up by each tab's own actual headers (`find_header_col()`), not a hardcoded single-tenant special case. Every tenant tab's own DoC-linking column gets filled automatically now, not just Property NSW's.
6. New + incomplete/ambiguous/unrecognized format → Evidence Tracker only, flagged with a specific comment — never guessed at silently. A "blanket tenant statement" (date range only, no per-collection detail) never becomes an invented tenant-tab row; it's flagged and annotated onto whichever EXISTING tenant-tab rows fall in range, via `detect_blanket_statement()` + `annotate_statement_note()`. Samuel Griffith Chmbrs has one such annotation column ("Statement note", column L) — bespoke, not part of the standard schema.
7. **Does NOT rebuild NABERS, D2 - WD Consolidated, Consolidated Summary, or BT_BaseBuilding Gap** (changed 17/09/2026). Those tabs used to be fully regenerated from tenant tabs every run under a simpler model (any verified contractor row gets an auto-calculated weight). That model broke once real Base Building weight matching entered the picture — Dock 1 needs a real matched weight, Dock 2 must stay weight-blank per NABERS Rules Section 6.6, and duplicate/orphan/zero-weight cases (like Abbro's 70.1kg entries) need a human judgment call. These tabs are now maintained by hand, or by the Secure Paper Auditor process (see Trigger phrases above) — never by this script. `rebuild_consolidated()`/`rebuild_nabers()` are still in the code for reference but are not called.
8. Fixes tab order automatically every run (`FRONT_TAB_ORDER`): NABERS → D2 - WD Consolidated → Consolidated Summary → BT_BaseBuilding Gap → Evidence Tracker → Tenant tracker → Method Details → Instruction → Tenant Name Mapping → Pending Queries → tenant tabs → D1 - Bin Tracker (last, deliberately not in FRONT_TAB_ORDER) → Sync Flags.
9. Writes a dated run log.

## Key decisions already made (don't re-ask)

- **NCAT** is not a separate tenant — Iron Mountain invoices the Property NSW tenancy as "NSW CIVIL & ADMINISTRATIVE TRIBUNAL"/"NCAT". All such docs route to the **Property NSW** tab.
- **Verified (Yes/No):** Yes for any complete CoD-backed row. No for manually back-reported schedule estimates with no CoD, and for Bin Tracker onsite-weight rows (those get their own Base Building matching treatment in BT_BaseBuilding Gap, not a CoD).
- **Wesfarmers**' Shred-X certificates are issued under "API Services Australia Pty Ltd" — confirmed with Ram 17/09/2026 as Wesfarmers' trading name at 66 Goulburn St, Level 10. In `KNOWN_ALIASES`.
- **Barker Evans**' contractor is Complete Shredding Services (Nicole Birrell) — different certificate format from everything else, own parser added 17/09/2026.
- Invoices/collection reports (non-CoD): logged as existence-only in Evidence Tracker (non-compliant), never extracted, never added to a tenant tab.
- Unknown tenant folder name → script stops and flags rather than guessing or auto-creating a tab.
- Calculated weight formula (internal reference only, never the real NABERS weight): `=19.2*0.9*<units collected>` — the NABERS default density for a 240L secure paper bin.

## If this file or the script's memory of past decisions is missing

The script itself is self-documenting enough to pick back up: read `PARSERS`,
`TENANT_FOLDER_TO_TAB`, `KNOWN_ALIASES`, `RESERVED_TABS`, and `FRONT_TAB_ORDER`
at the top of `secure_paper_processor.py`, run `--all --dry-run` first, and
read the run log it produces before going live. For anything about
BT_BaseBuilding Gap, Pending Queries, or the audit process, read
`secure-paper-auditor.md` in the "66 Goulburn" claude.ai project — that's
the durable copy, kept in sync with what's actually been done here.
