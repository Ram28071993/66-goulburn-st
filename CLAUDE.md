# 66 Goulburn Street, Sydney - EC Focus

Site-level project for 66 Goulburn St (Knight Frank): NABERS Waste rating and advisory services. Each workstream (secure paper, raw data processing, rating, contamination audit, advisory) gets its own folder under `docs/` and `scripts/`.

## Workstreams
- **Secure paper** - `docs/secure-paper/Secure_Paper_Instructions.md` (trigger phrases "Run secure paper for 66 Goulburn", "Audit the secure paper masterfile", folder layout, rules). Scripts versioned in `scripts/secure-paper/`.
- Others (raw data, calculations, rating, workshop, advisory): add a `docs/<workstream>/` file when work starts.

## Where the data lives (NOT in this repo)
SharePoint-synced site root: `C:\Users\RamiroNasta\EC FOCUS PTY. LTD\EC - 4.0 NABERS\4.9 Knight Frank\66 Goulburn St - Sydney\`
- `0. Consultancy\Secure Paper\` - CoD evidence, masterfile, run logs, and the **live copy of the secure paper scripts**
- `1. Raw data from contractors\`, `2. Calculations\`, `3. Ready for upload\`, `4. Uploaded\`, `5. Contamination audit\`, `6. Data verification\`, `7. Performance reports\`, `8. NABERS Certificates\`, `10. Rating Documentation\`, `11. Waste Workshop\`, `09.Proposals\`

## Secure paper scripts: where they run
`secure_paper_processor.py` resolves its folders from its own location (portable for colleagues via SharePoint), so it **runs from the SharePoint `Secure Paper` folder**, not from this repo. After changing a script there, copy it to `scripts/secure-paper/` and commit. Same for the instructions file (SharePoint `CLAUDE.md` -> `docs/secure-paper/Secure_Paper_Instructions.md`).

## Standing rules
- Senior waste / NABERS assessor mindset. Conservative reporting; never take contractor weights at face value.
- Back up before overwriting any workbook. Never edit raw data or tenant data without asking.
- Short, structured, commercially direct. Australian English. User writes in Spanish; client deliverables in English.

## Git workflow (standing authorization from Ramiro)
- Single branch `main`, remote `https://Ram28071993@github.com/Ram28071993/66-goulburn-st.git` (account Ram28071993). No branches, no merges.
- Repo holds only docs, instructions and scripts. **Never commit CoDs, invoices, workbooks, run logs or raw data** (see `.gitignore`).
- After each meaningful milestone, commit and push without asking; tell the user in one line what was pushed.
- Before ending a session with uncommitted or unpushed work, commit + push it (or say why not).
