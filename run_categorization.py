"""
Procurement Categorization Runner

Three ways to use:
  1. Double-click / run directly  → auto-finds template_po_search*.csv in same folder
  2. python run_categorization.py myfile.csv          → specific input, default output
  3. python run_categorization.py myfile.csv out.csv  → specific input and output
     (The GUI uses option 2/3 — passing the file you browsed to)
"""

import sys, os, glob
import pandas as pd

script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)
from categorization import categorize_dataframe


def _is_interactive():
    return sys.stdin is not None and sys.stdin.isatty()


def _read_file_robust(path):
    """Read a CSV or Excel file, return a DataFrame."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(path)
    try:
        return pd.read_csv(path, low_memory=False)
    except UnicodeDecodeError:
        return pd.read_csv(path, low_memory=False, encoding="latin-1")


def build_excel_report(df, output_path):
    """Build the Excel report from an already-categorized DataFrame."""
    from build_detail_excel_v2 import build_report
    build_report(df, output_path)


def main():
    # ── Decide input file(s) ─────────────────────────────────────────────────
    if len(sys.argv) > 1:
        input_files = [sys.argv[1]]
        if not os.path.exists(input_files[0]):
            print(f"ERROR: File not found: {input_files[0]}")
            if _is_interactive():
                input("\nPress Enter to exit...")
            sys.exit(1)
    else:
        input_files = glob.glob(os.path.join(script_dir, "template_po_search*.csv"))
        if not input_files:
            print("ERROR: No source CSV found.")
            print("  Either pass a file path as an argument, or place a")
            print("  template_po_search*.csv file in this folder.")
            if _is_interactive():
                input("\nPress Enter to exit...")
            sys.exit(1)

    output_file = sys.argv[2] if len(sys.argv) > 2 else os.path.join(script_dir, "categorized_output.csv")

    # ── Read file(s) ─────────────────────────────────────────────────────────
    dfs = []
    for f in input_files:
        print(f"Reading: {os.path.basename(f)}")
        try:
            dfs.append(_read_file_robust(f))
        except Exception as e:
            print(f"ERROR reading {os.path.basename(f)}: {e}")
            if _is_interactive():
                input("\nPress Enter to exit...")
            sys.exit(1)

    df = pd.concat(dfs, ignore_index=True)
    print(f"Total rows: {len(df):,}")

    # ── Categorize ───────────────────────────────────────────────────────────
    print("Categorizing...")
    results = categorize_dataframe(df)

    # ── Build clean output (handles duplicate column names) ──────────────────
    cat_cols = ["master_bucket", "sub_bucket_l2", "sub_bucket_l3",
                "rule_pass", "rule_pass_label", "rule_hit",
                "confidence_score", "confidence_label", "services_review_flag"]

    base_cols = [c for c in results.columns if c not in cat_cols]
    out = results[base_cols].copy()

    col_positions = {col: [] for col in cat_cols}
    for i, c in enumerate(results.columns):
        if c in col_positions:
            col_positions[c].append(i)

    for col in cat_cols:
        pos = col_positions.get(col, [])
        if pos:
            out[col] = results.iloc[:, pos[-1]].values

    # ── Write output ─────────────────────────────────────────────────────────
    out_dir = os.path.dirname(output_file)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    out.to_csv(output_file, index=False, encoding="utf-8-sig")
    print(f"Written to: {output_file}")

    # ── Bucket summary ───────────────────────────────────────────────────────
    print("\n-- Bucket Summary --")
    summary = out["master_bucket"].value_counts()
    for bucket, count in summary.items():
        pct = count / len(out) * 100
        print(f"  {bucket:<45} {count:>6,}  ({pct:.1f}%)")
    print(f"  {'TOTAL':<45} {len(out):>6,}")

    uncategorized = (out["master_bucket"] == "Uncategorized").sum()
    if uncategorized:
        print(f"\n  WARNING: {uncategorized:,} rows still uncategorized")
    else:
        print("\n  All rows categorized.")

    # ── Services review queue ────────────────────────────────────────────────
    if "services_review_flag" in out.columns:
        review_count = out["services_review_flag"].sum()
        spend_col    = next((c for c in out.columns if "extended" in c.lower() and "price" in c.lower()), None)
        review_spend = out.loc[out["services_review_flag"] == True, spend_col].sum() if spend_col else 0
        print(f"\n-- Services Review Queue --")
        print(f"  {review_count:,} rows need sub-classification  "
              + (f"  ${review_spend:,.0f} spend" if review_spend else ""))

    return out


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nERROR: {e}")
        if _is_interactive():
            input("\nPress Enter to exit...")
        sys.exit(1)
    if _is_interactive():
        input("\nPress Enter to exit...")
