# AGENTS.md

## Mission
This repository is a Public Documentation Shell (`-Docs`) for the Personal Finance Automation ecosystem.

## Boundaries
- **DO NOT** commit real `.env` files, Plaid credentials, or live Google Sheet IDs.
- **DO NOT** copy raw transaction data or live bank identifiers from the private repository.
- Rely strictly on the synthetic CSVs in `examples/sample_data/` for data demonstration.
- Keep all code inside `examples/sanitized-code-excerpts/`.

## Work Handoff
When resuming work on this repository:
1. Review `README.md` to understand the public boundary.
2. Review the sanitized code in `examples/sanitized-code-excerpts/` to understand the Python/Apps Script split.
3. If documenting new features, describe the system architecture rather than uploading the raw local execution scripts.
