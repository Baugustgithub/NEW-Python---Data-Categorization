# CLAUDE.md

## Project Overview
Procurement data categorization pipeline (Python/pandas). Reads CSV procurement data, normalizes the schema, validates inputs, applies a 5-pass rule engine with adaptive confidence scoring, runs QA diagnostics, and outputs categorized CSV + Excel reports.

## Pipeline Architecture (6 stages)
1. **Ingest** (`run_categorization.py`) — read CSV/Excel files
2. **Normalize** (`schema.py`) — map column aliases to canonical names
3. **Validate** (`schema.py`) — fail-fast on missing required columns, warn on recommended
4. **Categorize** (`categorization.py`) — 5-pass rule engine with signal-aware confidence
5. **QA** (`schema.py`) — post-categorization diagnostics, flag suspicious patterns
6. **Export** (`run_categorization.py`, `build_detail_excel_v2.py`) — CSV + Excel output

## Key Files
- `schema.py` — Column alias maps, normalization, validation, diagnostics (runs BEFORE categorization)
- `categorization.py` — Rule engine (5-pass: vendor overrides → commodity codes → vendor map → category metadata → keywords → account fallback)
- `run_categorization.py` — CLI runner / pipeline orchestrator
- `build_detail_excel_v2.py` — Excel report builder
- `aggregate_spend.py` — Spend-by-bucket aggregation
- `launcher.pyw` — Desktop GUI launcher (tkinter)
- `utils.py` — Shared utilities (CSV reading, numeric parsing)
- `reference/` — Source CSVs/XLSX the lookup tables were built from (not read at runtime)

## Schema Rules
- Column aliases are defined in `schema.py:COLUMN_SCHEMA` — update aliases there when export formats change
- **Required columns** (pipeline stops without them): `Vendor Name`, `Extended Price`
- **Recommended columns** (pipeline warns, accuracy degrades): `Commodity Code`, `Product Description`, `Category Level 1`
- **Optional columns** (no warning): `Manufacturer`, `Account`, `Creation Date`, `Contract No`, `Procurement Method`
- NEVER silently convert a missing column to empty strings without at least logging which columns were not found

## Confidence Model
- Base confidence is set by rule pass (0.95 for vendor/commodity, 0.5 for keywords, 0.2 for account fallback)
- A `signal_factor` (0.5–1.0) scales confidence based on how many of 6 input signals are globally available
- Per-row: if only 1 signal is non-blank and classification came from Pass 4+, confidence caps at 0.15
- Confidence labels: Very High (≥0.8), High (≥0.6), Medium (≥0.4), Low (<0.4)

## Resource Discipline
- Inspect only necessary files
- Avoid full-repo scans
- Prefer targeted reads over full file loads
- Do not generate large artifacts unless required
- Prefer incremental edits over full rewrites

## File Exclusions
Never proactively read:
- `__pycache__/`, `.git/`, `*.xlsx`, `*.csv` (data files)
- Large data exports or binaries

## Output Rules
- Default to concise summaries
- No large tables unless necessary
- No duplicate files or version spam
- Prefer compact formats (CSV, JSONL)

## Code Practices
- Shared utilities live in `utils.py` — do not duplicate `safe_num_series` or `read_csv_robust`
- Column resolution lives in `schema.py` — do not hardcode column name hunting elsewhere
- Prefer vectorized pandas operations over `df.apply(..., axis=1)` with row-level Python
- When iterating rows is unavoidable, pre-extract columns as lists and use `range(len(df))` — not `to_dict()` per row
- Avoid `pd.Series([val] * len(df))` — use `pd.Series(val, index=df.index)` instead
- Regex patterns in `KEYWORD_PATTERNS` are pre-compiled at module load (`_COMPILED`) — keep it that way
- `KEYWORD_PATTERNS` supports optional 6th element for exclude patterns
- `VENDOR_OVERRIDES` (Pass 0) takes priority over `VENDOR_MAP` (Pass 2) — do not duplicate entries across both

## Validation
- Start with the cheapest checks (syntax, imports)
- Escalate only if needed (run with sample data)

## Editing Style
- Small diffs over rewrites
- Preserve structure unless change is necessary
- `categorization.py` is large due to lookup tables — that's expected, don't refactor the data into external files without good reason
