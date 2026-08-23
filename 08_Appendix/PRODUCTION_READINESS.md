# Production Readiness — honest statement

## WORKING NOW (used on real weekly batches)
- AP parser (7 entity file formats, edge-case hardened) and bank statement parser (6 bank formats)
- Payment List with enrichment rules, validation, "Needs manual fix" workflow, bulk edit with confirm+undo
- ERP CSV export per entity with export gates (reference, completeness, amount)
- Bank creation reconciliation: 4-field match, field-level differences, inline fix, reasoned override
- Debit confirmation from transaction-level statement data; per-account cash sufficiency & funding status
- 2026 run-day / close-window / currency-SLA calendar
- Exception Center with owner / age / status / note; audit log of every action
- Approval email composer with per-entity counts and per-currency totals attachment
- Light/dark themes, exception-first Control Tower home, WCAG-AA audited UI

## DEMONSTRATED (works, shown with demo or real data)
- End-to-end demo on fully synthetic data (`07_Demo_Data/`) — no real information required
- Funding-shortfall detection (caught a real ~5M THB gap in testing)

## PROTOTYPE ONLY
- Operator identity: a setting, not authentication
- Audit log: complete but stored in the browser (localStorage), not on a server

## NOT IMPLEMENTED (deliberately out of MVP scope)
- Payment execution / bank connectivity (the tool never moves money)
- Multi-user backend, roles, concurrent editing
- System-enforced maker-checker (procedural today via approval email)
- Automated alerting (Zalo/Teams), bank APIs, cash-forecast integration — roadmap

## REQUIRED BEFORE PRODUCTION
1. Authentication (SSO) + role-based access control
2. Backend database + server-side, append-only audit store
3. System-enforced maker-checker on export & override
4. Secure internal hosting + encryption at rest; proper secrets management (none needed today — none exist)
5. Multi-user concurrency model
6. UAT with Treasury users over 2–3 real cycles
7. IT security review; monitoring; backup/disaster recovery
8. Bank/API integration hardening if statement pull is automated
