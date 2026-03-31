"""
aggregate_spend.py

Reads categorized_output.csv and writes spend_by_bucket.csv.
Also prints a quick console summary.

Usage:
    py -3.12 aggregate_spend.py
    py -3.12 aggregate_spend.py <categorized_csv> <out_csv>
"""

import os
import sys
import pandas as pd
from utils import safe_num_series as _safe_num_series, read_csv_robust as _read_csv_robust


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    inp = sys.argv[1] if len(sys.argv) > 1 else os.path.join(script_dir, "categorized_output.csv")
    outp = sys.argv[2] if len(sys.argv) > 2 else os.path.join(script_dir, "spend_by_bucket.csv")

    df = _read_csv_robust(inp)

    if "Extended Price" not in df.columns:
        raise ValueError("Expected 'Extended Price' column in categorized file.")

    df["_spend"] = _safe_num_series(df["Extended Price"])

    group = (
        df.groupby("master_bucket", dropna=False)["_spend"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
        .rename(columns={"_spend": "total_spend"})
    )

    group.to_csv(outp, index=False)

    total = float(group["total_spend"].sum())
    print("\n-- Spend by Master Bucket ------------------------------------------")
    for _, r in group.iterrows():
        print(f"{str(r['master_bucket']):<40} ${r['total_spend']:>15,.2f}")
    print(f"{'TOTAL':<40} ${total:>15,.2f}")
    print(f"\nWritten: {outp}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nERROR: {e}")
        if sys.stdin.isatty():
            input("\nPress Enter to exit...")
        sys.exit(1)
    if sys.stdin.isatty():
        input("\nPress Enter to exit...")
