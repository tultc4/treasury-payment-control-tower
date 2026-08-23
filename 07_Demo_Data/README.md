# Demo Data — 100% synthetic

No real vendor, beneficiary, account number, SWIFT, amount or invoice appears in these files.
Entity names and currency mix mirror the real run so every dashboard control fires.

| File | Load on | Contents |
|---|---|---|
| `AP_PAYMENT_SAMPLE_SANITIZED.xlsx` | Payment List (upload icon) | 18 payment groups · 6 entities · 5 currencies. One row is missing its value date and one its account — on purpose, so the "Needs manual fix" workflow has something to find. |
| `BANK_RESULT_SAMPLE.xlsx` | Bank Creation (upload icon, after exporting ERP CSVs) | Matches the sample above: 13 exact matches, 1 amount difference (+100), 1 wrong account, 1 SWIFT branch-code warning (…B01). |

Verified end-to-end on the submitted build: import → enrichment issue surfaced → reconciliation
shows 13 CREATED / 2 DIFF / 1 ⚠ — exactly the demo story in `DEMO_GUIDE.md`.

For the Bank Debit / statement step, present with the team's own T-1 statement file
(statement formats are bank-specific; no synthetic statement is included to avoid fabricating
bank output).
