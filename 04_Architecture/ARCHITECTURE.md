# Architecture

**Design principle: the architecture is a control decision.** Everything runs client-side on
the Treasury laptop, so real bank/beneficiary data never touches a server. Run cost ≈ $0.

```
AP Excel / bank statement files (real-world, messy)
        │
        ▼
┌───────────────────────────────────────────────────────────────┐
│  PYTHON PARSERS  (pandas + openpyxl)                          │
│  • AP parser: entity detection from filename, header          │
│    intelligence (aliases, accent-stripping, per-sheet table   │
│    width), dedupe across restated sheets, running-log period  │
│    filter, DD/MM vs MM/DD disambiguation, run-day filter      │
│  • Statement parser: 6 bank formats (Excel/CSV/PDF),          │
│    transaction table + balance derivation, account master     │
└───────────────────────────────────────────────────────────────┘
        │  standard workbooks (PAYMENT_LIST · TRANSACTION_DETAIL + BALANCE_SUMMARY)
        ▼
┌───────────────────────────────────────────────────────────────┐
│  CONTROL TOWER DASHBOARD (single self-contained HTML file)    │
│  SheetJS import · treasury enrichment rules · ERP CSV builder │
│  reference-uniqueness control · 4-field bank match            │
│  transaction-level debit match · cash sufficiency per account │
│  run-day & holiday calendar · exception workflow · audit log  │
│  persistence: browser localStorage (payments, balances,       │
│  overrides, exception metadata, audit events, settings)       │
└───────────────────────────────────────────────────────────────┘
        │
        ▼
ERP CSVs · approval email + XLSX evidence · exception/table CSV exports · audit log
```

## Data flow per weekly run

1. AP files → AP parser → `AP_PAYMENT_TEMPLATE_OUTPUT.xlsx` → **Payment List** (import replaces the previous batch).
2. Treasury enriches (bank / method / charges / value date) — rules pre-fill, human confirms.
3. **ERP Export** builds one CSV per entity; blocked payments (duplicate reference, missing fields, invalid amount) cannot be exported.
4. Bank creation result file → **Bank Creation**: 4-field match; differences fixed inline or overridden with a reason; approval email composed.
5. T-1 statements → statement parser → **Bank Debit**: debits confirmed only by transaction-level rows; balances feed **cash sufficiency**.
6. Everything lands in **Exceptions** (owner / age / status) and the **audit log**.

## Automated vs. human

| Automated | Human approval required |
|---|---|
| Parsing, validation, enrichment defaults, matching, sufficiency, calendar checks, logging | ERP export per entity · every DIFF override (written reason) · approval email L1/L2 · bulk edits (confirmation + one-level undo) |

## Security posture

- 100% client-side; the public GitHub Pages URL serves an **empty shell** — no payment data is hosted.
- No credentials, tokens or API keys anywhere in the codebase (nothing to leak — verified).
- CSV outputs are formula-injection escaped (`=`, `+`, `@` prefixes neutralized).
- Demo data in this package is fully synthetic.

## Why this approach fits (track: Non-Agent)

Payment control needs **deterministic, reproducible, auditable** logic. A generative model
cannot be the arbiter of whether a SWIFT matches; rules can — and every rule here is inspectable
in one file.

## Honest production gaps

Single-user by design today. Production needs: SSO/authentication, role-based access, a
server-side database + audit store (localStorage is per-browser and clearable), system-enforced
maker-checker, secure hosting, monitoring, and a formal security review.
See `08_Appendix/PRODUCTION_READINESS.md`.
