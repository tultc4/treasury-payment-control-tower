# Prompt Documentation — how AI was used to build this project

**Tool:** Claude (Claude Code, agentic coding assistant) · **Role split:** AI built and tested; the treasury owner decided every business rule, approved every change, and verified outputs against real files.

---

## 1. Working method

Every feature followed the same loop:

> **Business prompt (owner) → AI builds → AI tests on real batch data in a real browser → owner reviews the result → commit.**

Prompts were written in Vietnamese as plain business requirements, not technical specs — the domain knowledge stayed with the treasury owner, the implementation with the AI. Example (verbatim):

> *"chỉ được đi lệnh thứ 4 hàng tuần, nếu tuần cuối trong tháng có ngày cuối tháng − 3 không dính thứ 4 thì không được làm lệnh… làm đến hết năm 2026"*
> → became the 2026 run-day calendar engine (35 run days, 17 blocked Wednesdays) — rules encoded exactly as dictated, then verified date-by-date.

## 2. Prompt patterns that shaped the product

**a. Role-stacked review panels.** Instead of "review my app", the owner prompted AI to act as *ten reviewers at once* — treasury operator, QA tester, treasury manager, CFO, CEO, auditor, IT architect, UX designer, hackathon judge, product manager — each with explicit questions to answer. This produced the 14-finding QA audit (all fixed), the exception-first redesign of the home screen, and the Top-10 change list that was then implemented in full.

**b. Control-first prompts.** The strongest features came from control questions, not feature requests. Example (verbatim):

> *"làm sao tôi biết được bạn parse đủ thông tin từ file AP list hay không, lỡ parse thiếu thì sao — cần cho 1 control chỗ này"*
> → became the **Parse Completeness Control**: every candidate source row must reconcile (kept + excluded-by-reason = scanned), written to a PARSE_CONTROL sheet and surfaced on the dashboard.

**c. Honesty constraints.** Submission-phase prompts explicitly forbade fabrication: *"Do not invent PASS — if chưa test thì ghi NOT TESTED"*, *"impact numbers are not fabricated — use placeholders labeled 'to be validated'"*. The test evidence therefore contains real NOT-TESTED entries, and the deck's unmeasured KPIs are labeled placeholders.

**d. Terminology challenges.** The owner challenged AI outputs the way an auditor would — e.g. rejecting the word *"settlement"* (the tool confirms a **debit on our own account**, not beneficiary settlement) and challenging a "7,000-row" figure until AI measured the physical file (8,368 rows, 2021→2026) and reconciled it with the owner's own volume estimate.

**e. Scope freeze.** A final "SUBMISSION FREEZE" prompt switched the AI from building to hardening: *stability > features, clarity > complexity, remove/merge/hide anything that doesn't support a treasury decision* — producing this cleaned package rather than more functionality.

## 3. What AI produced under those prompts

Two Python parsers (hardened against real-file edge cases: Vietnamese headers, dual amount columns, sheets restating the same payments, headers on row 5, DD/MM-vs-MM/DD ambiguity), the single-file dashboard with its control logic, the automated browser test runs on real batches (46 payments / 333 statement lines), this documentation set, the 6-slide deck, and the narrated demo video (script, screen-drive and voice-over all AI-generated, verified frame-by-frame).

## 4. What stayed human

Every payment rule and bank mapping (dictated, not guessed) · every approval of a change before commit · every override and its written reason · the decision that **nothing pays automatically** · final wording of every judge-facing claim.

## 5. One-line summary for scoring

> AI was used as a build-and-test engine driven by treasury-language prompts; the human supplied the rules, challenged the outputs like an auditor, and kept every money decision. The result is deterministic, evidenced software — not AI output taken on trust.
