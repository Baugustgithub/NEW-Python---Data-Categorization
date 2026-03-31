"""
build_detail_excel_v2.py

Builds a multi-sheet Excel workbook from categorized_output.csv.

Sheets:
1) Summary
2) Spend by Bucket
3) Top Vendors
4) Services Review Queue
5) Uncategorized

Usage:
    py -3.12 build_detail_excel_v2.py
    py -3.12 build_detail_excel_v2.py <categorized_csv> <out_xlsx>
"""

import os
import sys
from datetime import datetime
import pandas as pd
from utils import safe_num_series as _safe_num_series, read_csv_robust as _read_csv_robust


def _coerce_date(df: pd.DataFrame) -> pd.Series:
    for col in ["Creation Date", "Date", "Invoice Date", "PO Date"]:
        if col in df.columns:
            return pd.to_datetime(df[col], errors="coerce")
    return pd.to_datetime(pd.Series([pd.NaT] * len(df)))


def _infer_month_fy(dates: pd.Series):
    # VCU fiscal year assumed July 1 – June 30
    month = dates.dt.to_period("M").astype(str)
    fy = dates.dt.year
    fy = fy.where(dates.dt.month < 7, fy + 1)  # Jul–Dec -> next FY
    return month, fy


def _infer_on_contract(df: pd.DataFrame) -> pd.Series:
    # Intentionally narrow: only "contract order" counts by method
    ON_CONTRACT_METHODS = {"contract order"}

    contract_cols = [c for c in ["Contract No", "Contract Number", "Contract #", "Contract"] if c in df.columns]
    if contract_cols:
        # Vectorized: check if any contract column has a real value (not blank/nan/none)
        _invalid = {"", "nan", "none"}
        has_contract_num = pd.Series(False, index=df.index)
        for col in contract_cols:
            normed = df[col].astype(str).str.strip().str.lower()
            has_contract_num = has_contract_num | ~normed.isin(_invalid)
    else:
        has_contract_num = pd.Series(False, index=df.index)

    method_col = None
    for c in ["Procurement Method", "Method", "Payment Method", "PO Type", "Order Type",
              "Procurement Method (For Purchasing Use Only)"]:
        if c in df.columns:
            method_col = c
            break

    if method_col:
        method_norm = df[method_col].astype(str).str.strip().str.lower()
        on_method = method_norm.isin(ON_CONTRACT_METHODS)
    else:
        on_method = pd.Series(False, index=df.index)

    return has_contract_num | on_method


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    inp = sys.argv[1] if len(sys.argv) > 1 else os.path.join(script_dir, "categorized_output.csv")
    outp = sys.argv[2] if len(sys.argv) > 2 else os.path.join(script_dir, "Procurement_Detail_Breakdown.xlsx")

    df = _read_csv_robust(inp)

    if "Extended Price" not in df.columns:
        raise ValueError("Expected 'Extended Price' column in categorized file.")
    if "master_bucket" not in df.columns:
        raise ValueError("Expected 'master_bucket' in categorized file. Re-run run_categorization.py first.")

    df["_spend"] = _safe_num_series(df["Extended Price"])

    dates = _coerce_date(df)
    df["_month"], df["_fy"] = _infer_month_fy(dates)

    df["_on_contract"] = _infer_on_contract(df)

    total_spend = float(df["_spend"].sum())
    total_rows = int(len(df))
    on_contract_spend = float(df.loc[df["_on_contract"], "_spend"].sum())
    on_contract_pct = (on_contract_spend / total_spend * 100) if total_spend else 0.0

    by_bucket = (
        df.groupby("master_bucket")["_spend"]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
        .rename(columns={"_spend": "total_spend"})
    )
    by_bucket["pct_of_total"] = by_bucket["total_spend"].apply(lambda x: (x / total_spend * 100) if total_spend else 0)

    # FY-end logic: prefer Apr/May/Jun if present, else last 3 periods
    period_cols = sorted(df["_month"].dropna().unique().tolist())
    fy_end_suffixes = ("-04", "-05", "-06")
    last3 = [p for p in period_cols if str(p).endswith(fy_end_suffixes)]
    if not last3:
        last3 = period_cols[-3:] if len(period_cols) >= 3 else period_cols

    df["_is_fy_end_month"] = df["_month"].isin(last3)
    fy_end_spend = float(df.loc[df["_is_fy_end_month"], "_spend"].sum())
    fy_end_pct = (fy_end_spend / total_spend * 100) if total_spend else 0.0

    vendor_col = None
    for c in ["Vendor Name", "Vendor", "Supplier", "Primary Second Party"]:
        if c in df.columns:
            vendor_col = c
            break

    if vendor_col:
        top_vendors = (
            df.groupby(vendor_col)["_spend"]
            .agg(total_spend="sum", line_count="size")
            .sort_values("total_spend", ascending=False)
            .reset_index()
            .rename(columns={vendor_col: "Vendor"})
            .head(500)
        )
        top_vendors["pct_of_total"] = top_vendors["total_spend"].apply(
            lambda x: (x / total_spend * 100) if total_spend else 0)
    else:
        top_vendors = pd.DataFrame(columns=["Vendor", "total_spend", "line_count", "pct_of_total"])

    if "services_review_flag" in df.columns:
        srvq = df[df["services_review_flag"].fillna(False).astype(bool)].copy()
    else:
        srvq = df.iloc[0:0].copy()

    unc = df[df["master_bucket"].fillna("") == "Uncategorized"].copy()

    with pd.ExcelWriter(outp, engine="openpyxl") as writer:
        summary = pd.DataFrame(
            [
                ["Rows", total_rows],
                ["Total Spend", total_spend],
                ["On-Contract Spend (inferred)", on_contract_spend],
                ["On-Contract %", on_contract_pct],
                ["FY-End Months Used", ", ".join(map(str, last3))],
                ["FY-End Spend", fy_end_spend],
                ["FY-End %", fy_end_pct],
                ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            ],
            columns=["Metric", "Value"],
        )
        summary.to_excel(writer, sheet_name="Summary", index=False)
        by_bucket.to_excel(writer, sheet_name="Spend by Bucket", index=False)
        top_vendors.to_excel(writer, sheet_name="Top Vendors", index=False)
        srvq.to_excel(writer, sheet_name="Services Review", index=False)
        unc.to_excel(writer, sheet_name="Uncategorized", index=False)

    print(f"Written: {outp}")


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
