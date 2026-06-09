# GEMINI.md

This file establishes durable context for Google AI agents operating within this repository.

## Current Architecture
- This is a Docs Shell.
- The real Google Apps Script deployment ID and Plaid tokens are strictly walled off in a separate private repository.
- We surface the integration pattern via `examples/sanitized-code-excerpts/`.
- We prove the data schema via `examples/sample_data/`.

## Agent Instructions
If asked to "run a sync" or "test the script":
1. Remember that this repo CANNOT run a live sync. It has no `.env` file and no API keys.
2. If asked to write new documentation, ensure you do not hallucinate specific live deployment paths. Stick to architectural concepts (e.g., "The Python Sync Engine normalizes the data").
