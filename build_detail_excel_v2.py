"""
build_detail_excel_v2.py

Builds a procurement intelligence workbook from categorized_output.csv.

Sheets:
1) Executive Dashboard       – KPI cards, bar chart, pie chart, trend line
2) How This Was Built        – methodology & data summary
3) Bucket Hierarchy          – 3-level spend taxonomy
4) Top 50 Vendors            – largest vendors per bucket
5) Vendor Concentration      – concentration risk metrics per bucket
6) Spend by Period           – monthly spend heatmap with data bars
7) Single-Txn Vendors        – tail-spend / one-off vendor list

Usage:
    py build_detail_excel_v2.py
    py build_detail_excel_v2.py <categorized_csv> <out_xlsx>
"""

import os
import re
import sys
from datetime import datetime

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import DataBarRule, ColorScaleRule

# ── Constants ────────────────────────────────────────────────────────────────

_ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]")

# Color palette
NAVY       = "1B2A4A"
DARK_BLUE  = "2C3E6B"
ACCENT     = "4472C4"
LIGHT_BLUE = "D6E4F0"
PALE_BLUE  = "EDF2F9"
WHITE      = "FFFFFF"
LIGHT_GRAY = "F2F2F2"
MID_GRAY   = "D9D9D9"
DARK_GRAY  = "595959"
GREEN      = "548235"
AMBER      = "BF8F00"
RED        = "C00000"

# Fonts
FONT_TITLE  = Font(name="Calibri", size=16, bold=True, color=WHITE)
FONT_H2     = Font(name="Calibri", size=12, bold=True, color=NAVY)
FONT_H3     = Font(name="Calibri", size=11, bold=True, color=DARK_BLUE)
FONT_HEADER = Font(name="Calibri", size=10, bold=True, color=WHITE)
FONT_BODY   = Font(name="Calibri", size=10, color="000000")
FONT_BODY_BOLD = Font(name="Calibri", size=10, bold=True, color="000000")
FONT_MUTED  = Font(name="Calibri", size=9, italic=True, color=DARK_GRAY)
FONT_NOTE   = Font(name="Calibri", size=9, color=DARK_GRAY)

# Fills
FILL_TITLE  = PatternFill("solid", fgColor=NAVY)
FILL_HEADER = PatternFill("solid", fgColor=ACCENT)
FILL_ROW_A  = PatternFill("solid", fgColor=WHITE)
FILL_ROW_B  = PatternFill("solid", fgColor=PALE_BLUE)
FILL_LIGHT  = PatternFill("solid", fgColor=LIGHT_BLUE)
FILL_BUCKET = PatternFill("solid", fgColor=LIGHT_GRAY)

# Border
THIN_BORDER = Border(
    bottom=Side(style="thin", color=MID_GRAY),
)
HEADER_BORDER = Border(
    bottom=Side(style="medium", color=NAVY),
)

# Alignment
ALIGN_LEFT   = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_RIGHT  = Alignment(horizontal="right", vertical="center")
ALIGN_WRAP   = Alignment(horizontal="left", vertical="top", wrap_text=True)


# ── Helpers ──────────────────────────────────────────────────────────────────

def _sanitize_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Strip characters that are illegal in Excel/XML cells."""
    df = df.copy()
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].apply(
            lambda v: _ILLEGAL_XML_RE.sub("", v) if isinstance(v, str) else v
        )
    return df


def _safe_num(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0.0)
    )


def _read_csv_robust(path: str) -> pd.DataFrame:
    # Prefer pickle sidecar — avoids all CSV quoting/parsing issues
    pkl_path = path.replace(".csv", ".pkl") if path.endswith(".csv") else None

    # Also accept .pkl path directly
    if path.endswith(".pkl"):
        pkl_path = path

    if pkl_path and os.path.exists(pkl_path):
        try:
            df = pd.read_pickle(pkl_path)
            # Drop phantom columns if present in pickle too
            phantom = [c for c in df.columns if str(c).startswith("Unnamed:")]
            if phantom:
                df = df.drop(columns=phantom)
            print(f"  (loaded from pickle: {os.path.basename(pkl_path)}, "
                  f"{len(df):,} rows x {len(df.columns)} cols)")
            return df
        except Exception as e:
            print(f"  WARNING: pickle load failed ({e}), falling back to CSV")

    # CSV fallback — skip phantom columns during parse
    csv_path = path if path.endswith(".csv") else path.replace(".pkl", ".csv")
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Neither pickle nor CSV found for: {path}")

    _keep = lambda c: not str(c).startswith("Unnamed:")
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            df = pd.read_csv(csv_path, encoding=enc, on_bad_lines="skip",
                             low_memory=False, usecols=_keep)
            print(f"  (loaded from CSV: {len(df):,} rows x {len(df.columns)} cols)")
            return df
        except UnicodeDecodeError:
            continue
        except Exception:
            break
    df = pd.read_csv(csv_path, encoding="latin-1", on_bad_lines="skip",
                     low_memory=False, usecols=_keep)
    print(f"  (loaded from CSV: {len(df):,} rows x {len(df.columns)} cols)")
    return df


def _coerce_date(df: pd.DataFrame) -> pd.Series:
    for col in ["Creation Date", "Date", "Invoice Date", "PO Date"]:
        if col in df.columns:
            return pd.to_datetime(df[col], errors="coerce")
    return pd.to_datetime(pd.Series([pd.NaT] * len(df)))


def _find_vendor_col(df: pd.DataFrame):
    for c in ["Vendor Name", "Vendor", "Supplier", "Primary Second Party"]:
        if c in df.columns:
            return c
    return None


def _fmt_currency(v):
    """Format a number as $X,XXX."""
    try:
        return f"${v:,.0f}"
    except (ValueError, TypeError):
        return str(v)


def _set_col_widths(ws, widths: dict):
    """Set column widths by 1-based column index."""
    for col_idx, w in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = w


def _write_title_banner(ws, title: str, subtitle: str, max_col: int):
    """Write a styled title row spanning max_col columns at row 1."""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    cell = ws.cell(row=1, column=1, value=title)
    cell.font = FONT_TITLE
    cell.fill = FILL_TITLE
    cell.alignment = ALIGN_LEFT
    ws.row_dimensions[1].height = 36

    if subtitle:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
        cell2 = ws.cell(row=2, column=1, value=subtitle)
        cell2.font = FONT_MUTED
        cell2.alignment = ALIGN_LEFT
        ws.row_dimensions[2].height = 20
    # Fill entire banner rows
    for c in range(1, max_col + 1):
        ws.cell(row=1, column=c).fill = FILL_TITLE


def _write_header_row(ws, row: int, headers: list, col_start: int = 1):
    """Write a styled header row."""
    for i, h in enumerate(headers, start=col_start):
        cell = ws.cell(row=row, column=i, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = HEADER_BORDER
    ws.row_dimensions[row].height = 24


def _write_data_rows(ws, start_row: int, data: list, col_start: int = 1,
                     currency_cols=None, pct_cols=None, int_cols=None):
    """Write data rows with alternating fills and number formatting."""
    currency_cols = set(currency_cols or [])
    pct_cols = set(pct_cols or [])
    int_cols = set(int_cols or [])
    for r_idx, row_data in enumerate(data):
        row_num = start_row + r_idx
        fill = FILL_ROW_B if r_idx % 2 == 1 else FILL_ROW_A
        for c_idx, val in enumerate(row_data, start=col_start):
            cell = ws.cell(row=row_num, column=c_idx, value=val)
            cell.font = FONT_BODY
            cell.fill = fill
            cell.border = THIN_BORDER
            cell.alignment = ALIGN_LEFT
            if c_idx in currency_cols:
                cell.number_format = '$#,##0'
                cell.alignment = ALIGN_RIGHT
            elif c_idx in pct_cols:
                cell.number_format = '0.0%'
                cell.alignment = ALIGN_RIGHT
            elif c_idx in int_cols:
                cell.number_format = '#,##0'
                cell.alignment = ALIGN_RIGHT
    return start_row + len(data)


def _freeze_and_filter(ws, freeze_row: int, max_col: int):
    """Freeze panes below freeze_row and add auto-filter."""
    ws.freeze_panes = ws.cell(row=freeze_row, column=1)
    ws.auto_filter.ref = f"A{freeze_row - 1}:{get_column_letter(max_col)}{ws.max_row}"


# ── Sheet builders ───────────────────────────────────────────────────────────

def _build_executive_dashboard(wb, df, total_spend, total_rows, vendor_col, generated):
    """Sheet 1: Executive Dashboard – KPIs, bar chart, pie chart, trend line."""
    ws = wb.create_sheet("Executive Dashboard")
    MAX_COL = 14
    _write_title_banner(ws, "Executive Dashboard",
                        f"Procurement Intelligence Summary  |  Generated {generated}", MAX_COL)

    # ── KPI Cards (row 4-5) ──────────────────────────────────────────────
    vendor_count = df[vendor_col].nunique() if vendor_col else 0
    bucket_count = df["master_bucket"].nunique()

    # Single-txn vendor stats
    single_count = 0
    single_spend = 0.0
    if vendor_col:
        txn_counts = df.groupby(vendor_col).size()
        single_vendors = set(txn_counts[txn_counts == 1].index)
        single_count = len(single_vendors)
        single_spend = float(df[df[vendor_col].isin(single_vendors)]["_spend"].sum())

    avg_txn = total_spend / total_rows if total_rows else 0

    kpis = [
        ("Total Spend", _fmt_currency(total_spend)),
        ("Transactions", f"{total_rows:,}"),
        ("Unique Vendors", f"{vendor_count:,}"),
        ("Categories", f"{bucket_count}"),
        ("Avg Txn Size", _fmt_currency(avg_txn)),
        ("Single-Txn Vendors", f"{single_count:,}"),
        ("Tail Spend", _fmt_currency(single_spend)),
    ]

    for i, (label, value) in enumerate(kpis):
        col = i * 2 + 1
        # Label cell
        cell = ws.cell(row=4, column=col, value=label)
        cell.font = FONT_NOTE
        cell.alignment = ALIGN_CENTER
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        # Value cell
        cell = ws.cell(row=5, column=col, value=value)
        cell.font = Font(name="Calibri", size=14, bold=True, color=NAVY)
        cell.alignment = ALIGN_CENTER
        cell.fill = FILL_LIGHT
        ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)

    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 30

    # ── Chart data area (hidden, rows 30+) ───────────────────────────────
    DATA_START = 30

    # -- Spend by Bucket (for bar chart) --
    bucket_spend = (
        df.groupby("master_bucket")["_spend"].sum()
        .sort_values(ascending=False).head(12)
    )
    ws.cell(row=DATA_START, column=1, value="Category")
    ws.cell(row=DATA_START, column=2, value="Spend")
    for i, (bname, bval) in enumerate(bucket_spend.items()):
        ws.cell(row=DATA_START + 1 + i, column=1, value=str(bname))
        ws.cell(row=DATA_START + 1 + i, column=2, value=float(bval))
    bar_end = DATA_START + len(bucket_spend)

    # Bar chart: Spend by Category
    bar_chart = BarChart()
    bar_chart.type = "col"
    bar_chart.title = "Spend by Category"
    bar_chart.y_axis.title = "Spend ($)"
    bar_chart.y_axis.numFmt = '$#,##0'
    bar_chart.x_axis.title = None
    bar_chart.style = 10
    bar_chart.width = 28
    bar_chart.height = 14
    bar_data = Reference(ws, min_col=2, min_row=DATA_START, max_row=bar_end)
    bar_cats = Reference(ws, min_col=1, min_row=DATA_START + 1, max_row=bar_end)
    bar_chart.add_data(bar_data, titles_from_data=True)
    bar_chart.set_categories(bar_cats)
    bar_chart.legend = None
    if bar_chart.series:
        bar_chart.series[0].graphicalProperties.solidFill = ACCENT
    ws.add_chart(bar_chart, "A7")

    # -- Top 10 Vendors (for pie chart) --
    PIE_COL = 4
    if vendor_col:
        top_vendors = (
            df.groupby(vendor_col)["_spend"].sum()
            .sort_values(ascending=False).head(10)
        )
        ws.cell(row=DATA_START, column=PIE_COL, value="Vendor")
        ws.cell(row=DATA_START, column=PIE_COL + 1, value="Spend")
        for i, (vname, vval) in enumerate(top_vendors.items()):
            clean_name = _ILLEGAL_XML_RE.sub("", str(vname))
            ws.cell(row=DATA_START + 1 + i, column=PIE_COL, value=clean_name)
            ws.cell(row=DATA_START + 1 + i, column=PIE_COL + 1, value=float(vval))
        pie_end = DATA_START + len(top_vendors)

        pie_chart = PieChart()
        pie_chart.title = "Top 10 Vendors by Spend"
        pie_chart.style = 10
        pie_chart.width = 18
        pie_chart.height = 14
        pie_data = Reference(ws, min_col=PIE_COL + 1, min_row=DATA_START,
                             max_row=pie_end)
        pie_cats = Reference(ws, min_col=PIE_COL, min_row=DATA_START + 1,
                             max_row=pie_end)
        pie_chart.add_data(pie_data, titles_from_data=True)
        pie_chart.set_categories(pie_cats)
        pie_chart.dataLabels = DataLabelList()
        pie_chart.dataLabels.showPercent = True
        pie_chart.dataLabels.showVal = False
        ws.add_chart(pie_chart, "H7")

    # -- Monthly Trend (for line chart) --
    TREND_COL = 7
    periods = sorted(df["_month"].dropna().unique().tolist())
    if periods:
        ws.cell(row=DATA_START, column=TREND_COL, value="Period")
        ws.cell(row=DATA_START, column=TREND_COL + 1, value="Spend")
        for i, p in enumerate(periods):
            ws.cell(row=DATA_START + 1 + i, column=TREND_COL, value=str(p))
            ws.cell(row=DATA_START + 1 + i, column=TREND_COL + 1,
                    value=float(df[df["_month"] == p]["_spend"].sum()))
        trend_end = DATA_START + len(periods)

        line_chart = LineChart()
        line_chart.title = "Monthly Spend Trend"
        line_chart.y_axis.title = "Spend ($)"
        line_chart.y_axis.numFmt = '$#,##0'
        line_chart.style = 10
        line_chart.width = 28
        line_chart.height = 12
        line_data = Reference(ws, min_col=TREND_COL + 1, min_row=DATA_START,
                              max_row=trend_end)
        line_cats = Reference(ws, min_col=TREND_COL, min_row=DATA_START + 1,
                              max_row=trend_end)
        line_chart.add_data(line_data, titles_from_data=True)
        line_chart.set_categories(line_cats)
        line_chart.legend = None
        if line_chart.series:
            line_chart.series[0].graphicalProperties.line.solidFill = ACCENT
        ws.add_chart(line_chart, "A23")

    # Column widths
    for c in range(1, MAX_COL + 1):
        ws.column_dimensions[get_column_letter(c)].width = 12
    ws.sheet_properties.tabColor = NAVY


def _build_methodology(wb, df, total_spend, total_rows, vendor_col, generated):
    """Sheet 1: How This Was Built."""
    ws = wb.create_sheet("How This Was Built")
    MAX_COL = 6
    _write_title_banner(ws, "How This Was Built", "Methodology & Data Summary", MAX_COL)

    row = 4
    sections = [
        ("Data Source",
         "This report is built from transaction-level procurement data (PO line detail). "
         "Each row represents a single purchase order line with vendor, commodity, dollar amount, "
         "and date information."),
        ("Categorization Method",
         "A multi-pass rules engine classified each transaction into a three-level bucket hierarchy:\n"
         "  Pass 0 – Vendor hard overrides (known vendor → forced bucket)\n"
         "  Pass 1 – Commodity code crosswalk (NIGP/commodity code → bucket mapping)\n"
         "  Pass 2 – Vendor always-list (vendor name pattern → bucket)\n"
         "  Pass 3 – Category metadata (L1 category field → bucket)\n"
         "  Pass 4 – Keyword / regex matching on description fields\n"
         "  Pass 5 – Account-family fallback (GL account prefix → bucket)\n\n"
         "Each transaction is assigned the first matching rule. Confidence scores reflect "
         "the specificity and reliability of the matching pass."),
        ("Bucket Structure",
         "Spend is organized into three levels:\n"
         "  Level 1 – Master Bucket (e.g. IT, Facilities / MRO, Services)\n"
         "  Level 2 – Sub-Bucket (e.g. IT Software / SaaS, Trades Services)\n"
         "  Level 3 – Detail (e.g. Network Security Equipment, Plumbing & HVAC)\n\n"
         "This hierarchy is defined in the commodity_codes and subcategory_to_major reference files."),
        ("Intended Use",
         "This workbook is a procurement intelligence report for management review. "
         "It communicates spend composition, vendor dependence, concentration risk, "
         "timing patterns, and tail-spend fragmentation to support strategic sourcing decisions."),
    ]

    for title, body in sections:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=MAX_COL)
        cell = ws.cell(row=row, column=1, value=title)
        cell.font = FONT_H2
        ws.row_dimensions[row].height = 22
        row += 1

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=MAX_COL)
        cell = ws.cell(row=row, column=1, value=body)
        cell.font = FONT_BODY
        cell.alignment = ALIGN_WRAP
        ws.row_dimensions[row].height = max(60, body.count("\n") * 15 + 30)
        row += 2

    # Data summary table
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=MAX_COL)
    ws.cell(row=row, column=1, value="Data Summary").font = FONT_H2
    row += 1

    # Confidence breakdown
    conf_counts = {}
    if "confidence_label" in df.columns:
        conf_counts = df["confidence_label"].value_counts().to_dict()

    # Rule pass breakdown
    pass_counts = {}
    if "rule_pass_label" in df.columns:
        pass_counts = df["rule_pass_label"].value_counts().to_dict()

    bucket_count = df["master_bucket"].nunique()
    vendor_count = df[vendor_col].nunique() if vendor_col else 0

    summary_items = [
        ("Total Rows", f"{total_rows:,}"),
        ("Total Spend", _fmt_currency(total_spend)),
        ("Unique Vendors", f"{vendor_count:,}"),
        ("Master Buckets", f"{bucket_count}"),
        ("", ""),
        ("Confidence Breakdown", ""),
    ]
    for label in ["Very High", "High", "Medium", "Low"]:
        cnt = conf_counts.get(label, 0)
        pct = cnt / total_rows * 100 if total_rows else 0
        summary_items.append((f"  {label}", f"{cnt:,}  ({pct:.1f}%)"))

    summary_items.append(("", ""))
    summary_items.append(("Classification Method Breakdown", ""))
    for label, cnt in sorted(pass_counts.items(), key=lambda x: -x[1]):
        pct = cnt / total_rows * 100 if total_rows else 0
        summary_items.append((f"  {label}", f"{cnt:,}  ({pct:.1f}%)"))

    summary_items.append(("", ""))
    summary_items.append(("Report Generated", generated))

    for metric, value in summary_items:
        c1 = ws.cell(row=row, column=1, value=metric)
        c2 = ws.cell(row=row, column=3, value=value)
        if metric and not metric.startswith("  "):
            c1.font = FONT_BODY_BOLD
        else:
            c1.font = FONT_BODY
        c2.font = FONT_BODY
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        row += 1

    _set_col_widths(ws, {1: 36, 2: 4, 3: 20, 4: 20, 5: 16, 6: 16})
    ws.sheet_properties.tabColor = NAVY


def _build_bucket_hierarchy(wb, df, total_spend):
    """Sheet 2: Bucket Hierarchy – three-level spend taxonomy."""
    ws = wb.create_sheet("Bucket Hierarchy")

    headers = ["Master Bucket", "Sub-Bucket", "Detail",
               "Transactions", "Total Spend", "% of Bucket", "% of Total"]
    MAX_COL = len(headers)
    _write_title_banner(ws, "Bucket Hierarchy",
                        "Spend structure from master bucket → sub-bucket → detail", MAX_COL)

    HEADER_ROW = 4
    _write_header_row(ws, HEADER_ROW, headers)

    l2_col = "sub_bucket_l2" if "sub_bucket_l2" in df.columns else None
    l3_col = "sub_bucket_l3" if "sub_bucket_l3" in df.columns else None

    # Build hierarchy data
    rows_data = []
    bucket_order = (
        df.groupby("master_bucket")["_spend"].sum()
        .sort_values(ascending=False).index.tolist()
    )

    for bucket in bucket_order:
        bdf = df[df["master_bucket"] == bucket]
        bucket_spend = bdf["_spend"].sum()
        bucket_count = len(bdf)
        # Master bucket summary row
        rows_data.append({
            "level": 0,
            "values": [bucket, "", "", bucket_count, bucket_spend,
                       1.0, bucket_spend / total_spend if total_spend else 0]
        })

        if l2_col:
            l2_order = (
                bdf.groupby(l2_col)["_spend"].sum()
                .sort_values(ascending=False).index.tolist()
            )
            for sub in l2_order:
                sdf = bdf[bdf[l2_col] == sub]
                sub_spend = sdf["_spend"].sum()
                sub_count = len(sdf)
                rows_data.append({
                    "level": 1,
                    "values": ["", str(sub), "", sub_count, sub_spend,
                               sub_spend / bucket_spend if bucket_spend else 0,
                               sub_spend / total_spend if total_spend else 0]
                })

                if l3_col:
                    l3_order = (
                        sdf.groupby(l3_col)["_spend"].sum()
                        .sort_values(ascending=False).index.tolist()
                    )
                    for detail in l3_order:
                        ddf = sdf[sdf[l3_col] == detail]
                        det_spend = ddf["_spend"].sum()
                        det_count = len(ddf)
                        rows_data.append({
                            "level": 2,
                            "values": ["", "", str(detail), det_count, det_spend,
                                       det_spend / bucket_spend if bucket_spend else 0,
                                       det_spend / total_spend if total_spend else 0]
                        })

    # Write rows with level-based formatting
    row = HEADER_ROW + 1
    for item in rows_data:
        level = item["level"]
        vals = item["values"]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=row, column=c_idx, value=val)
            cell.border = THIN_BORDER
            if level == 0:
                cell.font = FONT_BODY_BOLD
                cell.fill = FILL_BUCKET
            elif level == 1:
                cell.font = FONT_BODY
                cell.fill = FILL_ROW_A
            else:
                cell.font = FONT_NOTE
                cell.fill = FILL_ROW_A

            if c_idx == 4:  # Transactions
                cell.number_format = '#,##0'
                cell.alignment = ALIGN_RIGHT
            elif c_idx == 5:  # Spend
                cell.number_format = '$#,##0'
                cell.alignment = ALIGN_RIGHT
            elif c_idx in (6, 7):  # Percentages
                cell.number_format = '0.0%'
                cell.alignment = ALIGN_RIGHT
        row += 1

    _set_col_widths(ws, {1: 32, 2: 32, 3: 36, 4: 14, 5: 18, 6: 13, 7: 13})

    # Data bars on Total Spend column (col 5) and color scale on % of Total (col 7)
    if rows_data:
        last_row = HEADER_ROW + len(rows_data)
        ws.conditional_formatting.add(
            f"E{HEADER_ROW + 1}:E{last_row}",
            DataBarRule(start_type="min", end_type="max",
                        color=ACCENT, showValue=True)
        )
        ws.conditional_formatting.add(
            f"G{HEADER_ROW + 1}:G{last_row}",
            ColorScaleRule(
                start_type="min", start_color="FFFFFF",
                mid_type="percentile", mid_value=50, mid_color="D6E4F0",
                end_type="max", end_color="4472C4"
            )
        )

    _freeze_and_filter(ws, HEADER_ROW + 1, MAX_COL)
    ws.sheet_properties.tabColor = ACCENT


def _build_top_vendors(wb, df, total_spend, vendor_col):
    """Sheet 3: Top 50 Vendors per bucket."""
    ws = wb.create_sheet("Top 50 Vendors")

    headers = ["Master Bucket", "Rank", "Vendor", "Transactions",
               "Total Spend", "Avg Txn Size", "% of Bucket", "% of Total"]
    MAX_COL = len(headers)
    _write_title_banner(ws, "Top 50 Vendors",
                        "Largest vendors by spend within each category", MAX_COL)

    HEADER_ROW = 4
    _write_header_row(ws, HEADER_ROW, headers)

    if not vendor_col:
        ws.cell(row=HEADER_ROW + 1, column=1, value="No vendor column found in data.").font = FONT_MUTED
        _set_col_widths(ws, {1: 32, 2: 8, 3: 40, 4: 14, 5: 18, 6: 16, 7: 13, 8: 13})
        ws.sheet_properties.tabColor = DARK_BLUE
        return

    bucket_order = (
        df.groupby("master_bucket")["_spend"].sum()
        .sort_values(ascending=False).index.tolist()
    )

    rows_data = []
    for bucket in bucket_order:
        bdf = df[df["master_bucket"] == bucket]
        bucket_spend = bdf["_spend"].sum()

        vendor_agg = (
            bdf.groupby(vendor_col)["_spend"]
            .agg(total="sum", count="size")
            .sort_values("total", ascending=False)
            .head(50)
            .reset_index()
        )

        for rank, (_, vr) in enumerate(vendor_agg.iterrows(), start=1):
            vname = str(vr[vendor_col])
            vname = _ILLEGAL_XML_RE.sub("", vname) if isinstance(vname, str) else vname
            vspend = vr["total"]
            vcount = int(vr["count"])
            avg_txn = vspend / vcount if vcount else 0
            rows_data.append([
                bucket, rank, vname, vcount, vspend, avg_txn,
                vspend / bucket_spend if bucket_spend else 0,
                vspend / total_spend if total_spend else 0,
            ])

    _write_data_rows(ws, HEADER_ROW + 1, rows_data,
                     currency_cols={5, 6}, pct_cols={7, 8}, int_cols={2, 4})

    _set_col_widths(ws, {1: 32, 2: 8, 3: 40, 4: 14, 5: 18, 6: 16, 7: 13, 8: 13})
    _freeze_and_filter(ws, HEADER_ROW + 1, MAX_COL)
    ws.sheet_properties.tabColor = DARK_BLUE


def _build_vendor_concentration(wb, df, total_spend, vendor_col):
    """Sheet 4: Vendor Concentration – risk & opportunity metrics per bucket."""
    ws = wb.create_sheet("Vendor Concentration")

    headers = ["Master Bucket", "Bucket Spend", "Unique Vendors",
               "Top-1 Share", "Top-3 Share", "Top-5 Share", "Top-10 Share",
               "Single-Txn Vendors", "Single-Txn % of Vendors",
               "Single-Txn Spend", "Single-Txn % of Bucket", "Concentration Flag"]
    MAX_COL = len(headers)
    _write_title_banner(ws, "Vendor Concentration",
                        "Supplier dependency risk and consolidation opportunities by category",
                        MAX_COL)

    HEADER_ROW = 4
    _write_header_row(ws, HEADER_ROW, headers)

    if not vendor_col:
        ws.cell(row=HEADER_ROW + 1, column=1, value="No vendor column found.").font = FONT_MUTED
        ws.sheet_properties.tabColor = RED
        return

    bucket_order = (
        df.groupby("master_bucket")["_spend"].sum()
        .sort_values(ascending=False).index.tolist()
    )

    rows_data = []
    for bucket in bucket_order:
        bdf = df[df["master_bucket"] == bucket]
        bucket_spend = bdf["_spend"].sum()

        vendor_spend = (
            bdf.groupby(vendor_col)["_spend"].sum()
            .sort_values(ascending=False)
        )
        unique_vendors = len(vendor_spend)
        if unique_vendors == 0 or bucket_spend == 0:
            continue

        cumulative = vendor_spend.cumsum()
        top1  = float(vendor_spend.iloc[0] / bucket_spend) if unique_vendors >= 1 else 0
        top3  = float(cumulative.iloc[min(2, unique_vendors - 1)] / bucket_spend)
        top5  = float(cumulative.iloc[min(4, unique_vendors - 1)] / bucket_spend)
        top10 = float(cumulative.iloc[min(9, unique_vendors - 1)] / bucket_spend)

        # Single-transaction vendors
        txn_counts = bdf.groupby(vendor_col).size()
        single_vendors_set = set(txn_counts[txn_counts == 1].index)
        single_txn = len(single_vendors_set)
        single_pct = single_txn / unique_vendors if unique_vendors else 0

        # Single-txn vendor spend
        single_spend = float(bdf[bdf[vendor_col].isin(single_vendors_set)]["_spend"].sum())
        single_spend_pct = single_spend / bucket_spend if bucket_spend else 0

        # Concentration flag
        if top1 >= 0.50:
            flag = "High – single vendor dominance"
        elif top3 >= 0.70:
            flag = "Moderate – top 3 control >70%"
        elif single_pct >= 0.60:
            flag = "Fragmented – many one-off vendors"
        else:
            flag = "Balanced"

        rows_data.append([
            bucket, bucket_spend, unique_vendors,
            top1, top3, top5, top10,
            single_txn, single_pct, single_spend, single_spend_pct, flag,
        ])

    row = HEADER_ROW + 1
    for r_idx, rd in enumerate(rows_data):
        fill = FILL_ROW_B if r_idx % 2 == 1 else FILL_ROW_A
        for c_idx, val in enumerate(rd, start=1):
            cell = ws.cell(row=row, column=c_idx, value=val)
            cell.font = FONT_BODY
            cell.fill = fill
            cell.border = THIN_BORDER
            if c_idx in (2, 10):  # Bucket Spend, Single-Txn Spend
                cell.number_format = '$#,##0'
                cell.alignment = ALIGN_RIGHT
            elif c_idx in (3, 8):  # Unique Vendors, Single-Txn Vendors
                cell.number_format = '#,##0'
                cell.alignment = ALIGN_RIGHT
            elif c_idx in (4, 5, 6, 7, 9, 11):  # all percentage columns
                cell.number_format = '0.0%'
                cell.alignment = ALIGN_RIGHT
            elif c_idx == 12:  # Concentration Flag
                cell.alignment = ALIGN_LEFT
                flag_val = str(val)
                if flag_val.startswith("High"):
                    cell.font = Font(name="Calibri", size=10, bold=True, color=RED)
                elif flag_val.startswith("Moderate"):
                    cell.font = Font(name="Calibri", size=10, color=AMBER)
                elif flag_val.startswith("Fragmented"):
                    cell.font = Font(name="Calibri", size=10, color=AMBER)
                elif flag_val.startswith("Balanced"):
                    cell.font = Font(name="Calibri", size=10, color=GREEN)
        row += 1

    _set_col_widths(ws, {1: 32, 2: 18, 3: 16, 4: 12, 5: 12, 6: 12, 7: 12,
                         8: 18, 9: 20, 10: 18, 11: 20, 12: 34})
    _freeze_and_filter(ws, HEADER_ROW + 1, MAX_COL)
    ws.sheet_properties.tabColor = RED


def _build_spend_by_period(wb, df, total_spend):
    """Sheet 5: Spend by Period – monthly heatmap by bucket."""
    ws = wb.create_sheet("Spend by Period")

    periods = sorted(df["_month"].dropna().unique().tolist())
    headers = ["Master Bucket", "Total Spend", "Active Periods"] + [str(p) for p in periods]
    MAX_COL = len(headers)
    _write_title_banner(ws, "Spend by Period",
                        "Monthly spend by category – darker shading indicates higher relative spend",
                        MAX_COL)

    HEADER_ROW = 4
    _write_header_row(ws, HEADER_ROW, headers)

    bucket_order = (
        df.groupby("master_bucket")["_spend"].sum()
        .sort_values(ascending=False).index.tolist()
    )

    pivot = df.pivot_table(index="master_bucket", columns="_month",
                           values="_spend", aggfunc="sum", fill_value=0)
    # Ensure all periods are present
    for p in periods:
        if p not in pivot.columns:
            pivot[p] = 0
    pivot = pivot[periods]

    # For heatmap: find max cell value across pivot
    max_val = pivot.values.max() if pivot.values.size > 0 else 1
    if max_val == 0:
        max_val = 1

    # Heatmap color ramp: white → light blue → accent blue
    def _heat_fill(val):
        if val <= 0:
            return PatternFill("solid", fgColor=WHITE)
        intensity = min(val / max_val, 1.0)
        if intensity < 0.25:
            return PatternFill("solid", fgColor=WHITE)
        elif intensity < 0.50:
            return PatternFill("solid", fgColor=PALE_BLUE)
        elif intensity < 0.75:
            return PatternFill("solid", fgColor=LIGHT_BLUE)
        else:
            return PatternFill("solid", fgColor="B4C6E7")

    PERIOD_START = 4  # period columns start at column 4

    row = HEADER_ROW + 1
    for r_idx, bucket in enumerate(bucket_order):
        fill = FILL_ROW_B if r_idx % 2 == 1 else FILL_ROW_A
        bucket_total = float(pivot.loc[bucket].sum()) if bucket in pivot.index else 0

        # Count active (non-zero) periods for this bucket
        active = int((pivot.loc[bucket] > 0).sum()) if bucket in pivot.index else 0

        cell = ws.cell(row=row, column=1, value=bucket)
        cell.font = FONT_BODY_BOLD
        cell.fill = fill
        cell.border = THIN_BORDER

        cell = ws.cell(row=row, column=2, value=bucket_total)
        cell.font = FONT_BODY_BOLD
        cell.fill = fill
        cell.border = THIN_BORDER
        cell.number_format = '$#,##0'
        cell.alignment = ALIGN_RIGHT

        cell = ws.cell(row=row, column=3, value=active)
        cell.font = FONT_BODY
        cell.fill = fill
        cell.border = THIN_BORDER
        cell.number_format = '#,##0'
        cell.alignment = ALIGN_CENTER

        for p_idx, period in enumerate(periods):
            col = PERIOD_START + p_idx
            val = float(pivot.loc[bucket, period]) if bucket in pivot.index else 0
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = FONT_BODY
            cell.fill = _heat_fill(val)
            cell.border = THIN_BORDER
            cell.number_format = '$#,##0'
            cell.alignment = ALIGN_RIGHT
        row += 1

    # Total row
    ws.cell(row=row, column=1, value="TOTAL").font = FONT_BODY_BOLD
    ws.cell(row=row, column=1).fill = FILL_LIGHT
    total_cell = ws.cell(row=row, column=2, value=total_spend)
    total_cell.font = FONT_BODY_BOLD
    total_cell.fill = FILL_LIGHT
    total_cell.number_format = '$#,##0'
    total_cell.alignment = ALIGN_RIGHT
    ws.cell(row=row, column=3, value=len(periods)).font = FONT_BODY_BOLD
    ws.cell(row=row, column=3).fill = FILL_LIGHT
    ws.cell(row=row, column=3).alignment = ALIGN_CENTER
    for p_idx, period in enumerate(periods):
        col = PERIOD_START + p_idx
        pval = float(df[df["_month"] == period]["_spend"].sum())
        cell = ws.cell(row=row, column=col, value=pval)
        cell.font = FONT_BODY_BOLD
        cell.fill = FILL_LIGHT
        cell.number_format = '$#,##0'
        cell.alignment = ALIGN_RIGHT

    _set_col_widths(ws, {1: 32, 2: 18, 3: 15})
    for i in range(PERIOD_START, MAX_COL + 1):
        ws.column_dimensions[get_column_letter(i)].width = 14

    # Data bars on the Total Spend column (col 2)
    last_data_row = HEADER_ROW + len(bucket_order)
    if last_data_row > HEADER_ROW:
        spend_range = f"B{HEADER_ROW + 1}:B{last_data_row}"
        ws.conditional_formatting.add(
            spend_range,
            DataBarRule(start_type="min", end_type="max",
                        color=ACCENT, showValue=True)
        )

        # Color scale on period cells (green → yellow → red by spend magnitude)
        if periods:
            period_start_letter = get_column_letter(PERIOD_START)
            period_end_letter = get_column_letter(PERIOD_START + len(periods) - 1)
            period_range = (f"{period_start_letter}{HEADER_ROW + 1}:"
                            f"{period_end_letter}{last_data_row}")
            ws.conditional_formatting.add(
                period_range,
                ColorScaleRule(
                    start_type="min", start_color="FFFFFF",
                    mid_type="percentile", mid_value=50, mid_color="D6E4F0",
                    end_type="max", end_color="4472C4"
                )
            )

    _freeze_and_filter(ws, HEADER_ROW + 1, MAX_COL)
    ws.sheet_properties.tabColor = GREEN


def _build_single_txn_vendors(wb, df, total_spend, vendor_col):
    """Sheet 6: Single-Txn Vendors – tail spend."""
    ws = wb.create_sheet("Single-Txn Vendors")

    headers = ["Master Bucket", "Vendor", "Spend", "Description",
               "% of Bucket", "Consolidation Opportunity"]
    MAX_COL = len(headers)
    _write_title_banner(ws, "Single-Transaction Vendors",
                        "One-off vendors that may represent unmanaged tail spend or consolidation candidates",
                        MAX_COL)

    HEADER_ROW = 4
    _write_header_row(ws, HEADER_ROW, headers)

    if not vendor_col:
        ws.cell(row=HEADER_ROW + 1, column=1, value="No vendor column found.").font = FONT_MUTED
        ws.sheet_properties.tabColor = AMBER
        return

    # Find description column
    desc_col = None
    for c in ["Description", "Item Description", "Line Description",
              "Short Description", "PO Description"]:
        if c in df.columns:
            desc_col = c
            break

    txn_counts = df.groupby(vendor_col).size()
    single_vendors = set(txn_counts[txn_counts == 1].index)

    sdf = df[df[vendor_col].isin(single_vendors)].copy()

    bucket_spend = df.groupby("master_bucket")["_spend"].sum().to_dict()

    rows_data = []
    for _, r in sdf.sort_values("_spend", ascending=False).iterrows():
        bucket = r["master_bucket"]
        vname = str(r[vendor_col])
        vname = _ILLEGAL_XML_RE.sub("", vname)
        spend = r["_spend"]
        desc = str(r.get(desc_col, "")) if desc_col else ""
        desc = _ILLEGAL_XML_RE.sub("", desc) if desc else ""
        bspend = bucket_spend.get(bucket, 0)
        pct_bucket = spend / bspend if bspend else 0

        # Simple opportunity flag
        if spend >= 10000:
            opp = "Review – significant one-off spend"
        elif spend >= 2500:
            opp = "Consider blanket PO or p-card"
        else:
            opp = "Low-dollar tail spend"

        rows_data.append([bucket, vname, spend, desc[:100], pct_bucket, opp])

    _write_data_rows(ws, HEADER_ROW + 1, rows_data,
                     currency_cols={3}, pct_cols={5})

    # Summary stats at top-right
    total_single = len(single_vendors)
    total_vendors = df[vendor_col].nunique()
    single_spend = sdf["_spend"].sum()

    stats_col = MAX_COL + 2
    ws.cell(row=4, column=stats_col, value="Tail Spend Summary").font = FONT_H3
    stats = [
        ("Single-txn vendors", f"{total_single:,}"),
        ("Total vendors", f"{total_vendors:,}"),
        ("Single-txn % of vendors", f"{total_single / total_vendors * 100:.1f}%" if total_vendors else "0%"),
        ("Single-txn spend", _fmt_currency(single_spend)),
        ("% of total spend", f"{single_spend / total_spend * 100:.1f}%" if total_spend else "0%"),
    ]
    for i, (label, val) in enumerate(stats, start=5):
        ws.cell(row=i, column=stats_col, value=label).font = FONT_BODY_BOLD
        ws.cell(row=i, column=stats_col + 1, value=val).font = FONT_BODY

    _set_col_widths(ws, {1: 32, 2: 40, 3: 16, 4: 50, 5: 14, 6: 36,
                         stats_col: 24, stats_col + 1: 18})
    _freeze_and_filter(ws, HEADER_ROW + 1, MAX_COL)
    ws.sheet_properties.tabColor = AMBER


def build_excel_from_df(df: pd.DataFrame, outp: str, log=print):
    """Build the Excel report directly from an in-memory DataFrame.

    This avoids the CSV round-trip entirely — no disk I/O, no re-parsing.
    Called by launcher.pyw and launcher.py when the DataFrame is already
    in memory after categorization.
    """
    # Drop any phantom columns that slipped through
    phantom = [c for c in df.columns if str(c).startswith("Unnamed:")]
    if phantom:
        df = df.drop(columns=phantom)

    if "Extended Price" not in df.columns:
        raise ValueError("Expected 'Extended Price' column in categorized file.")
    if "master_bucket" not in df.columns:
        raise ValueError("Expected 'master_bucket' in categorized file. Re-run categorization first.")

    df["_spend"] = _safe_num(df["Extended Price"])
    dates = _coerce_date(df)
    df["_month"], df["_fy"] = (
        dates.dt.to_period("M").astype(str),
        dates.dt.year.where(dates.dt.month < 7, dates.dt.year + 1),
    )

    vendor_col = _find_vendor_col(df)
    total_spend = float(df["_spend"].sum())
    total_rows = len(df)
    generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    log(f"Rows: {total_rows:,}  |  Spend: {_fmt_currency(total_spend)}  |  Vendors: {df[vendor_col].nunique() if vendor_col else 'N/A'}")

    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)

    log("  Building: Executive Dashboard")
    _build_executive_dashboard(wb, df, total_spend, total_rows, vendor_col, generated)

    log("  Building: How This Was Built")
    _build_methodology(wb, df, total_spend, total_rows, vendor_col, generated)

    log("  Building: Bucket Hierarchy")
    _build_bucket_hierarchy(wb, df, total_spend)

    log("  Building: Top 50 Vendors")
    _build_top_vendors(wb, df, total_spend, vendor_col)

    log("  Building: Vendor Concentration")
    _build_vendor_concentration(wb, df, total_spend, vendor_col)

    log("  Building: Spend by Period")
    _build_spend_by_period(wb, df, total_spend)

    log("  Building: Single-Txn Vendors")
    _build_single_txn_vendors(wb, df, total_spend, vendor_col)

    log(f"Saving: {outp}")
    wb.save(outp)
    log(f"Written: {outp}")


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # Accept .pkl or .csv path; default to pkl if available
    if len(sys.argv) > 1:
        inp = sys.argv[1]
    else:
        pkl_default = os.path.join(script_dir, "categorized_output.pkl")
        csv_default = os.path.join(script_dir, "categorized_output.csv")
        inp = pkl_default if os.path.exists(pkl_default) else csv_default
    outp = sys.argv[2] if len(sys.argv) > 2 else os.path.join(script_dir, "Procurement_Detail_Breakdown.xlsx")

    print(f"Reading: {inp}")
    df = _read_csv_robust(inp)

    if "Extended Price" not in df.columns:
        raise ValueError("Expected 'Extended Price' column in categorized file.")
    if "master_bucket" not in df.columns:
        raise ValueError("Expected 'master_bucket' in categorized file. Re-run categorization first.")

    df["_spend"] = _safe_num(df["Extended Price"])
    dates = _coerce_date(df)
    df["_month"], df["_fy"] = (
        dates.dt.to_period("M").astype(str),
        dates.dt.year.where(dates.dt.month < 7, dates.dt.year + 1),
    )

    vendor_col = _find_vendor_col(df)
    total_spend = float(df["_spend"].sum())
    total_rows = len(df)
    generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print(f"Rows: {total_rows:,}  |  Spend: {_fmt_currency(total_spend)}  |  Vendors: {df[vendor_col].nunique() if vendor_col else 'N/A'}")

    # Build workbook with openpyxl directly for full formatting control
    from openpyxl import Workbook
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    print("  Building: Executive Dashboard")
    _build_executive_dashboard(wb, df, total_spend, total_rows, vendor_col, generated)

    print("  Building: How This Was Built")
    _build_methodology(wb, df, total_spend, total_rows, vendor_col, generated)

    print("  Building: Bucket Hierarchy")
    _build_bucket_hierarchy(wb, df, total_spend)

    print("  Building: Top 50 Vendors")
    _build_top_vendors(wb, df, total_spend, vendor_col)


    print("  Building: Vendor Concentration")
    _build_vendor_concentration(wb, df, total_spend, vendor_col)

    print("  Building: Spend by Period")
    _build_spend_by_period(wb, df, total_spend)

    print("  Building: Single-Txn Vendors")
    _build_single_txn_vendors(wb, df, total_spend, vendor_col)

    print(f"Saving: {outp}")
    wb.save(outp)
    print(f"Written: {outp}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print(f"\nERROR: {e}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        if sys.stdin.isatty():
            input("\nPress Enter to exit...")
        sys.exit(1)
    if sys.stdin.isatty():
        input("\nPress Enter to exit...")
