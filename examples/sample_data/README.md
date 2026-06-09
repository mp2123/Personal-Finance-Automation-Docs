# Danny Bank Sample Data

This directory contains committed synthetic fixtures for product demos, screenshots, tests, and onboarding examples.

The rows are not real financial activity. They are intentionally labeled with `DEMO`, `SAMPLE ONLY`, and `Demo` account names so they cannot be confused with a live user ledger.

Files:
- `demo_transactions.csv`: sample `Transactions!A:G` ledger rows used by Demo Mode in the local control center.
- `demo_income.csv`: sample manual-income CSV shape for documentation and dry-run examples.

Do not copy these rows into a real Google Sheet. Demo Mode is read-only and is not connected to the live Sheet. Real user imports should stay under `src/imports/`, which is ignored by git and requires explicit dry-run review plus confirmation before any append.
