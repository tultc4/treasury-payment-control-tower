# Business Process — AS-IS vs TO-BE

## AS-IS (manual, per weekly run, × 7 entities)

```
AP sends Excel files (formats differ per entity)
  → Treasury checks line by line (references, accounts, amounts, dates)
  → manual enrichment: payment bank, method, charges, value date
  → ERP upload file assembled by hand
  → bank creation file compared against Excel BY EYE
  → each debit searched in statement files/PDFs
  → balances checked account by account
  → issues discovered reactively, evidence reconstructed for audit
```

Pain: fragmented data, no live status, silent failure modes (wrong account, duplicate
reference, wrong value date, holiday miss, funding shortfall, false "settled").

## TO-BE (Control Tower)

```
AP files → automated parsing & validation (every row)
        → rule-based treasury enrichment (human confirms)
        → ERP-ready CSV per entity (blocked payments cannot export)
        → automated 4-field bank creation reconciliation
        → transaction-level settlement confirmation
        → cash sufficiency per account BEFORE the run
        → centralized exceptions (owner · age · status)
        → automatic audit trail
```

**Key message:** the objective is not simply automation — it is moving Treasury from
*reactive checking* to *proactive exception management*.

# Control Matrix (core controls implemented)

| # | Risk | Control in the tool | Type | Evidence |
|---|---|---|---|---|
| 1 | Duplicate payment (reference reused) | Reference uniqueness vs. bank history + within batch; blocked until overridden with reason | Preventive | Reference status pill · override log |
| 2 | Wrong beneficiary account / SWIFT | 4-field bank match (amount, beneficiary, account, SWIFT-8) on every created payment | Detective | DIFF status + field-level difference text |
| 3 | Invalid amount reaches ERP | Zero/negative amount blocks export; amounts validated at parse | Preventive | INVALID_AMOUNT status |
| 4 | Missing enrichment reaches ERP | Method/charge/account/value-date completeness gate | Preventive | "Needs manual fix" panel |
| 5 | Payment on non-run day / close window | 2026 run calendar: Wednesdays only, month-start (4 WD) & month-end (EOM−3) locks | Preventive | Calendar + blocked-day reasons |
| 6 | Holiday settlement miss | Currency SLA rules (e.g. IDR/THB/MYR = VD+2 working days) + holiday calendar | Detective | Expected-completion + holiday note |
| 7 | False settlement confirmation | Debit confirmed ONLY by transaction-level statement match; balance movement never accepted; generic references (TRANSFer…) never auto-match | Preventive | Debit status logic |
| 8 | Funding shortfall | Per-account sufficiency vs. minimum buffer before the run | Detective | SHORTFALL alert + funding-gap KPI |
| 9 | Unauthorized/bulk mistakes | Bulk edits: confirmation with scope + export-reset warning + one-level undo | Preventive | Confirm dialog + audit event |
| 10 | Lost evidence | Audit log: imports, edits, overrides, exports, approvals, resolutions (who + when) | Evidential | Admin → Audit log |
| 11 | Parse completeness (rows silently missed at intake) | PARSE_CONTROL reconciliation: per source file, candidate rows = kept + excluded-by-reason (out-of-window, restated, summary, no amount, no reference); any file that does not close is flagged CHECK, blocks trust in the batch, and is surfaced on the Control Tower attention panel | Preventive | PARSE_CONTROL sheet + Admin table + audit event |
| 12 | Approval bypass | Approval email to L1/L2 with per-entity counts + per-currency totals attachment | Procedural (to be system-enforced in production) | Email + XLSX evidence |
