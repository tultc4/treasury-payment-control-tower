# Testing Evidence

Honest status. **PASS** is claimed only for scenarios actually executed — either on the real
weekly batch (46 payment groups, 333 statement transactions), on the sanitized sample in
`07_Demo_Data/`, or by automated browser tests during QA. Anything else is **NOT TESTED**.

| # | Test | Expected | Actual | Result |
|---|---|---|---|---|
| 1 | Successful payment end-to-end | AP row → enriched → ERP CSV → CREATED at bank → PENDING/CONFIRMED debit | Full chain observed on real batch; ERP CSV schema verified | **PASS** (real batch) |
| 2 | Failed / missing bank record | Exported payment absent from bank result → NOT_CREATED | Observed on real batch (payment missing from bank file was flagged) | **PASS** (real batch) |
| 3 | Pending payment | CREATED but no debit yet → PENDING_DEBIT follow-up | Observed on real batch | **PASS** (real batch) |
| 4 | Duplicate payment reference vs. bank history | REFERENCE_BLOCKED; export impossible; reasoned override available | Exercised on real batch incl. override + undo | **PASS** (real batch) |
| 5 | Duplicate reference within batch | DUPLICATE_IN_BATCH exception; auto-suffix control for multi-invoice groups | Exercised on real batches in earlier cycles | **PASS** (real batch) |
| 6 | Amount mismatch | DIFF with field-level text "bank X vs ours Y" | Sanitized sample row seeded +100 → DIFF with exact amounts shown | **PASS** (sample, automated) |
| 7 | Account mismatch | DIFF naming the account values | Sanitized sample seeded wrong account → DIFF | **PASS** (sample, automated) |
| 8 | SWIFT branch-code anomaly | ⚠ review flag when >8 chars not ending XXX | Sanitized sample seeded "…B01" → warning raised | **PASS** (sample, automated) |
| 9 | Currency mismatch as its own dimension | — | Currency is not separately compared in the bank match (reference+amount+bene+account+SWIFT are) | **NOT TESTED — documented limitation** |
| 10 | Missing treasury enrichment | Export blocked; "Needs manual fix" jump | Seeded missing value date in sample → surfaced + jump works | **PASS** (sample, automated) |
| 11 | Invalid / ambiguous date in source | Parser resolves DD/MM vs MM/DD when only one side can be a month | Real KMZ file with MM/DD text parsed correctly after fix | **PASS** (real file) |
| 12 | Holiday / working-day settlement | IDR/THB/MYR VD+2 working days, weekends skipped, holiday note | Verified against 2026 calendar incl. weekend example | **PASS** (rule verification) |
| 13 | Funding shortfall | SHORTFALL + attention panel + funding-gap KPI | Real ~5M THB shortfall on a GN TH account flagged before the run | **PASS** (real batch) |
| 14 | Duplicate statement transaction / duplicate bank debit | Should not double-confirm a debit | Not yet exercised with a crafted duplicate statement | **NOT TESTED** |
| 15 | No transaction-level settlement evidence | Balance movement alone must never confirm a debit | By-design rule; real accounts with balance-only data stayed unconfirmed | **PASS** (design + real data) |
| 16 | Zero / negative amount | INVALID_AMOUNT blocks ERP export + raises exception | Automated test: amount=0 → blocked + exception listed | **PASS** (automated) |
| 17 | Empty workbook import | No crash; existing data untouched | Automated test: header-only file → 18 rows kept, no error | **PASS** (automated) |
| 18 | Wrong file format (.txt) | Safely ignored | Automated test: no crash, no state change | **PASS** (automated) |
| 19 | Duplicate AP re-import (same file twice) | Replace-all; no double counting | Automated test: 18 → 18 | **PASS** (automated) |
| 20 | Bulk edit safety | Confirmation with scope + export-reset warning; undo restores | Automated test: confirm text verified; undo restored all values | **PASS** (automated) |
| 21 | CSV formula injection | Leading =/+/@ escaped in every CSV export | Automated test: `=cmd()` → `'=cmd()`; `-123.45` untouched | **PASS** (automated) |
| 22 | Parse completeness control | Every candidate source row accounted for: kept + excluded-by-reason = scanned, per file | Real batch: 7/7 files reconciled (e.g. ZPI 83 = 11 kept + 68 window + 1 no-amount + 3 no-ref; VNGSING 8,367 candidates itemized); CHECK path fires the attention panel | **PASS** (real files + automated) |
| 23 | Large batch performance | Thousands of rows | Current largest tested: 46 groups / 333 transactions — instant. Not tested at 1,000+ | **NOT TESTED at scale** |

## UI/accessibility QA (summary)

Separate audit fixed 14 findings: page overflow at 1280px, 210 unlabeled form controls
(now 0), Escape-close for modals, sticky-header regression, WCAG AA contrast in both themes,
faux-bold fonts, mixed table type sizes. All re-verified post-fix with zero console errors.

# Known Limitations

1. **Single-user, client-side** — localStorage is per-browser and clearable; no concurrent use. (Deliberate for data privacy in the MVP.)
2. **Maker-checker is procedural** — the approval email is the four-eyes moment; the system does not yet enforce a second approver before export.
3. **Currency mismatch** is not an independent match dimension (see test 9).
4. **Duplicate statement import** protection unverified (test 14).
5. **Scale** beyond ~1–2k rows unverified; no pagination/virtualization yet (test 22).
6. **Statement formats**: 6 bank formats supported; a new bank format needs a parser mapping.
7. **NOT_CREATED detection** depends on payments being marked exported before the bank-result import.
8. Table-expand opens a popup window — the browser must allow popups once.
