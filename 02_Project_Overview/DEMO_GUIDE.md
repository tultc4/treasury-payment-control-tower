# Demo Guide — one story, ~6 minutes

Tell it as a story: **"This is Wednesday morning. We have money to pay and one screen to control it."**
Use `07_Demo_Data/` files (fully synthetic) or the team's real weekly batch.

> Setup: open `03_Source_Code/dashboard/index.html` (or the live URL). Start on **Control Tower** with an empty state — that's intentional.

## STEP 1 — The problem (30s)
Show one raw AP Excel briefly (or describe it): several files, thousands of rows, no status.
Say: *"Today, everything you're about to see is done by hand, file by file."*

## STEP 2 — Load the payment data (45s)
Payment List → upload icon → `AP_PAYMENT_SAMPLE_SANITIZED.xlsx`.
Point at: 18 groups parsed instantly, entity/currency detected per row, Payment Bank & Method
pre-enriched from treasury rules. *"Enrichment that used to be copy-paste is now a rule."*

## STEP 3 — Validation finds the problems for us (45s)
Point at the red **"Needs manual fix"** box → click **Go to next fix** — it jumps straight to the
row missing a value date. Fix it inline.
Say: *"We didn't search for the error. The error came to us."*

## STEP 4 — ERP readiness (30s)
ERP Export tab → per-entity cards → download one CSV.
*"This file is ERP-ready — schema, ordering account, charge codes — and a blocked payment can
never reach it: duplicates, missing fields and zero amounts are stopped at source."*

## STEP 5 — Bank creation reconciliation (60s)
Bank Creation tab → upload `BANK_RESULT_SAMPLE.xlsx`.
Point at the summary chips: **13 matched · 2 DIFF**. Click the red **Diff to review** chip — it
jumps to the first difference and names the exact field (bank amount vs. ours).
Fix one inline → match flips to TRUE live. Override the other with a written reason.
*"Four fields are compared on every payment: amount, beneficiary, account, SWIFT."*

## STEP 6 — Settlement evidence (30s)
Bank Debit tab (with the team's real T-1 statement for a live audience).
Key sentence: *"A balance going down is never proof a specific payment settled. We only confirm
a debit when the transaction itself appears in the statement."*

## STEP 7 — Cash sufficiency (30s)
Control Tower / Cash Position → funding status per account.
*"In testing this flagged a real ~5M THB shortfall before the run — that is the slide-4 story."*

## STEP 8 — Exception Center + close (45s)
Exceptions tab → assign an owner, set one to Resolved with a note → show it leaves the bell count.
Then Admin → **Audit log**: every import, edit, override and resolution from this very demo is
already recorded.

**Closing line:**
> *"Instead of Treasury searching for problems, the system brings problems to Treasury —
> and proves what was done about them."*

---
### If something goes wrong live
- Refresh (F5): data persists in the browser (localStorage).
- Blank state after refresh on another machine: re-upload the two demo files (30 seconds).
- Popup blocked when expanding a table: allow popups once, or skip — not core to the story.
