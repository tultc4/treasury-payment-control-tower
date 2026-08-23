# Project Summary — one page

**Product:** Treasury Payment Control Tower
**Problem:** the weekly multi-entity payment run is controlled by hand across disconnected
Excel/bank/statement files; errors and funding gaps look like normal rows until money moves.
**Solution:** two deterministic parsers + one offline control dashboard covering
validation → enrichment → ERP export → bank reconciliation → settlement evidence →
cash sufficiency → exceptions → audit trail. Exception-first: the home screen answers
*"what needs my attention today?"*
**Positioning:** not a dashboard full of features — a treasury control solution for one
specific business problem, that scales by configuration (entities, banks, currencies are data,
not code).

## One answer per reviewer

| Reviewer | The one thing to look at |
|---|---|
| Treasury operator | Payment List "Needs manual fix → Go to next fix", and the Diff-jump chip on Bank Creation |
| Senior QA | `06_Testing/TEST_CASES.md` — honest PASS / NOT TESTED, plus parser edge-case notes |
| Treasury manager | 4-field bank match + per-account sufficiency + run calendar with close windows |
| CFO | Control Tower KPIs: total run value per currency, funding gap, payments at risk — before release |
| CEO | Slide 6: works today for one team; scale = configuration, not code |
| Auditor | Admin → audit log; reasoned overrides; debit confirmed only by transaction-level evidence |
| IT architect | `04_Architecture/ARCHITECTURE.md` — client-side by design, $0 run cost, explicit go-live gaps |
| UX designer | Exception-first home, one accent + semantic colors, dense financial tables, light/dark |
| Judge | Slide 4: two non-hypothetical catches (5M THB shortfall; systematic wrong-amount pattern) |
| Product manager | `08_Appendix/PRODUCTION_READINESS.md` — working / demonstrated / not implemented / required |

## Deliberately NOT included (scope discipline)

- No payment execution / bank connectivity — the tool prepares, checks and evidences; people pay.
- No multi-user backend in the MVP — chosen so real bank data never leaves the Treasury laptop.
- No AI/LLM component — this problem needs deterministic, auditable logic (hence Non-Agent track).
