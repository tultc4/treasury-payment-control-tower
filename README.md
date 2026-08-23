# Treasury Payment Control Tower

**FA AI Hackathon 2026 · Track: Non-Agent (Automation & Analytics) · Individual entry — Lê Thị Cẩm Tú (Stella)**

> From manual payment checking to automated treasury control.

**Live demo:** https://tultc4.github.io/treasury-payment-dashboard/
*(the public page is an empty application shell — all payment data stays on the operator's machine)*

---

## What is it?

A single-screen control tower for the weekly treasury payment run of the Oversea entities
(7 legal entities · 6 currencies · 3 banks). It covers the whole lifecycle:

**AP payment list → Treasury enrichment → ERP CSV export → Bank creation reconciliation
→ Debit / settlement confirmation → Cash sufficiency → Exception management → Audit trail**

Two Python parsers consolidate any AP or bank file into a standard workbook; one offline HTML
dashboard runs every control and produces every output.

## What problem does it solve?

Today each run means: multiple AP Excel files → manual line-by-line checking → manual
enrichment → an ERP upload file built by hand → a bank file compared by eye → statement PDFs
searched per debit → balances checked account by account. The real risk is not the effort — it
is that a wrong account, a duplicate reference or a funding shortfall **looks exactly like a
correct row** until the money has moved.

## Who uses it?

| Who | What they get |
|---|---|
| Treasury operations | One screen, exception-first: the system brings problems to Treasury instead of Treasury hunting for them |
| Treasury manager / CFO | Run value per currency, funding gap, payments at risk, batch progress — before money moves |
| Internal control / audit | Unique-reference control, reasoned overrides, transaction-level settlement evidence, an audit log of every action |

## How it works

**Input** — AP payment lists, bank creation file, bank statements, master data & balances, holiday/run calendar
**Control** — parse & validate every row · enrich bank/method/charges · block duplicates & invalid amounts · 4-field bank match (amount, beneficiary, account, SWIFT) · transaction-level debit matching · cash sufficiency per account
**Output** — clean ERP CSV per entity · live status AP → debit · exception list with owner/age · funding-gap alert · approval email + evidence workbook · audit trail

Nothing pays automatically. Treasury approves every export, resolves every difference (an
override requires a written reason), and sends the approval email before release.

## Key capabilities

- Exception-first home screen ("What needs my attention today?")
- Parser tolerant of real-world files (Vietnamese headers, duplicated sheets, headers on row 5, multi-year running logs)
- Payment-reference uniqueness vs. bank history, with reasoned manual override
- Debit confirmed **only** by transaction-level statement evidence — never by balance movement
- 2026 payment-run calendar built in (Wednesday runs, month-start / month-end close windows: 35 run days, 17 blocked Wednesdays)
- Full audit log: imports, edits, overrides, exports, approvals, exception resolutions
- Parse completeness control: every source row is accounted for (kept + excluded-by-reason = scanned) — reconciliation per file in the PARSE_CONTROL sheet and on the Admin screen

## Business impact (see slide 4)

Before: a day of line-by-line checking per cycle, reactive issue handling.
After: 46 payment groups parsed and validated in under a minute; only true exceptions surface.
Two non-hypothetical catches during testing: a ~5M THB funding shortfall flagged **before** the
run, and a systematic wrong-amount pattern found in a real source file.
*Time-saving KPIs are measured by stopwatching a manual cycle vs. a dashboard cycle on the same
batch — placeholders in the deck are labeled and not fabricated.*

## How to run locally

1. Open `03_Source_Code/dashboard/index.html` in any modern browser (double-click — no install, no server).
2. Load `07_Demo_Data/AP_PAYMENT_SAMPLE_SANITIZED.xlsx` via the upload icon on **Payment List**.
3. Load `07_Demo_Data/BANK_RESULT_SAMPLE.xlsx` on **Bank Creation** (after exporting ERP CSVs).
4. Optional — parsers: `pip install -r requirements.txt` in each parser folder, then run the `.cmd`/`.bat`.

Full walk-through: `DEMO_GUIDE.md`.

## Current status

Working prototype / MVP — the full lifecycle runs on real weekly batches today, QA-audited
(accessibility, responsive, contrast; 14 findings fixed). Single-operator by design; see
`08_Appendix/PRODUCTION_READINESS.md` for the go-live list (SSO, server-side audit store,
system-enforced maker-checker, hosting, security review, UAT).

## Limitations (honest)

See `06_Testing/KNOWN_LIMITATIONS.md`. Highlights: single-user client-side architecture;
maker-checker is procedural (approval email) not system-enforced; currency-mismatch is not a
separate match dimension; statement duplicate-import protection not yet verified.

## Roadmap

Auto approval package → Zalo/Teams alerts → holiday alert agent → bank API integration →
cash-forecast link → Regional Treasury Control Tower.


---

## Repository map

| Folder | Contents |
|---|---|
| 01_Presentation | 6-slide deck (PPTX + PDF), speaker notes embedded |
| 02_Project_Overview | README · one-page summary · demo guide |
| 03_Source_Code | Dashboard (single HTML) · AP parser · bank statement parser |
| 04_Architecture | Architecture & data flow, security posture, track rationale |
| 05_Business_Process | AS-IS vs TO-BE + control matrix |
| 06_Testing | 22 test scenarios (honest PASS / NOT TESTED) + known limitations |
| 07_Demo_Data | Fully synthetic sample files (verified end-to-end) |
| 08_Appendix | Production readiness · 10-persona review & verdict · Q&A defense prep |
