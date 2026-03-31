# CLAUDE.md

## Project Overview
Procurement data categorization pipeline (Python/pandas). Reads CSV procurement data, applies a 5-pass rule engine, outputs categorized CSV + Excel reports.

## Key Files
- `categorization.py` — Rule engine (5-pass: vendor overrides → commodity codes → vendor map → category metadata → keywords → account fallback)
- `run_categorization.py` — CLI runner
- `build_detail_excel_v2.py` — Excel report builder
- `aggregate_spend.py` — Spend-by-bucket aggregation
- `launcher.pyw` — Desktop GUI launcher (tkinter)
- `utils.py` — Shared utilities (CSV reading, numeric parsing)
- `reference/` — Source CSVs/XLSX the lookup tables were built from (not read at runtime)

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
- Prefer vectorized pandas operations over `df.apply(..., axis=1)` with row-level Python
- When iterating rows is unavoidable, pre-extract columns as lists and use `range(len(df))` — not `to_dict()` per row
- Avoid `pd.Series([val] * len(df))` — use `pd.Series(val, index=df.index)` instead
- Regex patterns in `KEYWORD_PATTERNS` are pre-compiled at module load (`_COMPILED`) — keep it that way
- `VENDOR_OVERRIDES` (Pass 0) takes priority over `VENDOR_MAP` (Pass 2) — do not duplicate entries across both

## Validation
- Start with the cheapest checks (syntax, imports)
- Escalate only if needed (run with sample data)

## Editing Style
- Small diffs over rewrites
- Preserve structure unless change is necessary
- `categorization.py` is large due to lookup tables — that's expected, don't refactor the data into external files without good reason
