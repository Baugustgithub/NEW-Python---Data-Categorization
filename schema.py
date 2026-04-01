"""
schema.py — Schema normalization, validation, and diagnostics for procurement data.

Pipeline stage 1: runs BEFORE categorization to ensure the DataFrame has
canonical column names and the minimum required signals are present.

Architecture:
  1. normalize_columns(df) — alias matching + content profiling → canonical names
  2. validate_schema(df)  — required/recommended/optional checks → errors or warnings
  3. diagnose_schema(df)  — blanks, coverage, signal strength → diagnostics dict
"""
import re
import pandas as pd

# ═══════════════════════════════════════════════════════════════════════════════
# CANONICAL COLUMN DEFINITIONS
#
# Each entry: canonical_name → {aliases, tier, description}
#   tier: "required"    — script MUST stop if missing
#         "recommended" — script warns + quantifies degradation
#         "optional"    — nice to have, no warning
# ═══════════════════════════════════════════════════════════════════════════════
COLUMN_SCHEMA = {
    "Vendor Name": {
        "aliases": [
            "vendor name", "vendor", "supplier", "supplier name",
            "primary second party", "vendor_name", "vendorname",
        ],
        "tier": "required",
        "desc": "Vendor/supplier name — drives Pass 0 and Pass 2",
    },
    "Extended Price": {
        "aliases": [
            "extended price", "extendedprice", "extended_price",
            "amount", "total amount", "total price", "line amount",
            "spend", "total spend", "line total", "po amount",
            "extended amount", "ext price", "ext. price",
        ],
        "tier": "required",
        "desc": "Dollar amount per line — required for spend analysis",
    },
    "Commodity Code": {
        "aliases": [
            "commodity code", "commoditycode", "commodity_code",
            "comm code", "commodity", "nigp code", "nigp",
            "commodity cd",
        ],
        "tier": "recommended",
        "desc": "6-digit commodity code — drives Pass 1 (highest-confidence structured match)",
    },
    "Product Description": {
        "aliases": [
            "product description", "description", "item description",
            "product_description", "line description", "desc",
            "item desc", "product desc", "po description",
        ],
        "tier": "recommended",
        "desc": "Line item description — drives Pass 4 keyword matching",
    },
    "Category Level 1": {
        "aliases": [
            "category level 1", "category_level_1", "categorylevel1",
            "cat level 1", "cat1", "category 1", "level 1 category",
            "category level1", "unspsc level 1", "category name",
        ],
        "tier": "recommended",
        "desc": "Top-level category metadata — drives Pass 3",
    },
    "Manufacturer": {
        "aliases": [
            "manufacturer", "mfr", "mfg", "manufacturer name",
            "brand", "brand name",
        ],
        "tier": "optional",
        "desc": "Manufacturer/brand — enriches Pass 4 keyword scan text",
    },
    "Account": {
        "aliases": [
            "account", "account code", "account number", "acct",
            "account no", "account #", "gl account", "gl code",
            "account_code", "fund code",
        ],
        "tier": "optional",
        "desc": "Account/GL code — drives Pass 5 fallback",
    },
    # ── Columns used by build_detail_excel_v2.py ──────────────────────
    "Creation Date": {
        "aliases": [
            "creation date", "date", "invoice date", "po date",
            "order date", "transaction date", "created date",
            "create date",
        ],
        "tier": "optional",
        "desc": "Date field — used for fiscal year / month analysis in Excel report",
    },
    "Contract No": {
        "aliases": [
            "contract no", "contract number", "contract #", "contract",
            "contract name", "contract id",
        ],
        "tier": "optional",
        "desc": "Contract identifier — used for on-contract spend analysis",
    },
    "Procurement Method": {
        "aliases": [
            "procurement method", "method", "payment method",
            "po type", "order type",
            "procurement method (for purchasing use only)",
        ],
        "tier": "optional",
        "desc": "Procurement method — used for on-contract inference",
    },
}

# Build a fast lookup: lowercased alias → canonical name
_ALIAS_MAP = {}
for canonical, spec in COLUMN_SCHEMA.items():
    _ALIAS_MAP[canonical.lower()] = canonical
    for alias in spec["aliases"]:
        _ALIAS_MAP[alias.lower()] = canonical


def _fuzzy_match_column(col_name):
    """Try exact alias match, then stripped/normalized match."""
    normed = re.sub(r"[^a-z0-9]", "", col_name.lower())
    # Exact alias match (lowercased)
    if col_name.lower() in _ALIAS_MAP:
        return _ALIAS_MAP[col_name.lower()]
    # Stripped match (remove all non-alphanumeric)
    for alias, canonical in _ALIAS_MAP.items():
        if re.sub(r"[^a-z0-9]", "", alias) == normed:
            return canonical
    return None


def normalize_columns(df):
    """
    Rename DataFrame columns to canonical names using alias matching.

    Returns:
        (df_renamed, mapping_report)
        mapping_report: list of {original, canonical, method} dicts
    """
    rename_map = {}
    report = []
    used_canonicals = set()

    for col in df.columns:
        canonical = _fuzzy_match_column(col)
        if canonical and canonical not in used_canonicals:
            if col != canonical:
                rename_map[col] = canonical
                report.append({"original": col, "canonical": canonical, "method": "alias"})
            else:
                report.append({"original": col, "canonical": canonical, "method": "exact"})
            used_canonicals.add(canonical)
        else:
            report.append({"original": col, "canonical": None, "method": "unmapped"})

    df_out = df.rename(columns=rename_map)
    return df_out, report


def validate_schema(df, report=None):
    """
    Check that required/recommended columns are present after normalization.

    Returns:
        (errors, warnings)
        errors:   list of strings — any of these means STOP
        warnings: list of strings — continue but degrade gracefully
    """
    errors = []
    warnings = []
    present = set(df.columns)

    for canonical, spec in COLUMN_SCHEMA.items():
        tier = spec["tier"]
        if canonical not in present:
            if tier == "required":
                errors.append(
                    f"REQUIRED column missing: '{canonical}' — {spec['desc']}. "
                    f"Known aliases: {spec['aliases'][:5]}"
                )
            elif tier == "recommended":
                warnings.append(
                    f"RECOMMENDED column missing: '{canonical}' — {spec['desc']}. "
                    f"Categorization accuracy will be degraded."
                )

    return errors, warnings


def diagnose_schema(df, mapping_report=None):
    """
    Produce a diagnostics dict describing data quality BEFORE categorization.

    Returns dict with:
        columns_found:    {canonical_name: {present, blank_pct, tier}}
        columns_missing:  {canonical_name: {tier, desc}}
        signals_available: list of pass names that have data
        signals_missing:   list of pass names that lack data
        estimated_degradation: string summary
    """
    diag = {
        "columns_found": {},
        "columns_missing": {},
        "signals_available": [],
        "signals_missing": [],
        "mapping_report": mapping_report or [],
    }

    n = len(df)
    if n == 0:
        return diag

    for canonical, spec in COLUMN_SCHEMA.items():
        if canonical in df.columns:
            blank_count = df[canonical].isna().sum() + (df[canonical].astype(str).str.strip() == "").sum()
            blank_pct = blank_count / n * 100
            diag["columns_found"][canonical] = {
                "present": True,
                "blank_pct": round(blank_pct, 1),
                "tier": spec["tier"],
            }
        else:
            diag["columns_missing"][canonical] = {
                "tier": spec["tier"],
                "desc": spec["desc"],
            }

    # Map columns to rule-engine passes
    pass_signals = {
        "Pass 0 (Vendor Overrides)":  "Vendor Name",
        "Pass 1 (Commodity Codes)":   "Commodity Code",
        "Pass 2 (Vendor Map)":        "Vendor Name",
        "Pass 3 (Category Metadata)": "Category Level 1",
        "Pass 4 (Keywords)":          "Product Description",
        "Pass 5 (Account Fallback)":  "Account",
    }

    for pass_name, col in pass_signals.items():
        found = diag["columns_found"].get(col)
        if found and found["blank_pct"] < 90:
            diag["signals_available"].append(pass_name)
        else:
            diag["signals_missing"].append(pass_name)

    return diag


def print_diagnostics(diag, warnings=None):
    """Print a human-readable diagnostics report to console."""
    SEP = "=" * 70
    print(f"\n{SEP}")
    print("  SCHEMA DIAGNOSTICS REPORT")
    print(SEP)

    # Column mapping
    if diag.get("mapping_report"):
        print("\n  Column Mapping:")
        for entry in diag["mapping_report"]:
            if entry["method"] == "exact":
                print(f"    ✓ {entry['original']}")
            elif entry["method"] == "alias":
                print(f"    → {entry['original']}  ⟶  {entry['canonical']}  (alias match)")
            else:
                print(f"    · {entry['original']}  (not mapped — passed through)")

    # Found columns with blank %
    if diag["columns_found"]:
        print("\n  Column Quality:")
        for col, info in diag["columns_found"].items():
            tier_tag = {"required": "REQ", "recommended": "REC", "optional": "OPT"}[info["tier"]]
            blank_bar = "██" if info["blank_pct"] < 10 else "▓▓" if info["blank_pct"] < 50 else "░░" if info["blank_pct"] < 90 else "  "
            status = "OK" if info["blank_pct"] < 10 else f"{info['blank_pct']:.0f}% blank"
            print(f"    [{tier_tag}] {blank_bar} {col:<30} {status}")

    # Missing columns
    if diag["columns_missing"]:
        print("\n  Missing Columns:")
        for col, info in diag["columns_missing"].items():
            tier_tag = {"required": "REQ", "recommended": "REC", "optional": "OPT"}[info["tier"]]
            print(f"    [{tier_tag}] ✗  {col:<30} {info['desc']}")

    # Signal availability
    print("\n  Rule Engine Signal Strength:")
    for sig in diag["signals_available"]:
        print(f"    ✓ {sig}")
    for sig in diag["signals_missing"]:
        print(f"    ✗ {sig}  — THIS PASS WILL NOT FIRE")

    # Warnings
    if warnings:
        print(f"\n  Warnings:")
        for w in warnings:
            print(f"    ⚠ {w}")

    print(SEP + "\n")


def print_post_categorization_diagnostics(df, total_spend):
    """Print QA diagnostics AFTER categorization runs. Flags suspicious patterns."""
    SEP = "=" * 70
    print(f"\n{SEP}")
    print("  POST-CATEGORIZATION QA CHECKS")
    print(SEP)

    n = len(df)
    flags = []

    # 1. Uncategorized rate
    if "master_bucket" in df.columns:
        unc_count = (df["master_bucket"] == "Uncategorized").sum()
        unc_pct = unc_count / n * 100 if n else 0
        if unc_pct > 50:
            flags.append(f"CRITICAL: {unc_pct:.0f}% of rows are Uncategorized ({unc_count:,} rows). "
                         "Input columns are likely missing or renamed.")
        elif unc_pct > 25:
            flags.append(f"WARNING: {unc_pct:.0f}% of rows are Uncategorized ({unc_count:,} rows). "
                         "Review input column mapping.")
        print(f"\n  Uncategorized: {unc_count:,} / {n:,} rows ({unc_pct:.1f}%)")

    # 2. Confidence distribution
    if "confidence_score" in df.columns:
        avg_conf = df["confidence_score"].mean()
        low_conf = (df["confidence_score"] <= 0.2).sum()
        low_pct = low_conf / n * 100 if n else 0
        print(f"  Avg confidence: {avg_conf:.0%}")
        print(f"  Low confidence (<=0.2): {low_conf:,} rows ({low_pct:.1f}%)")
        if avg_conf < 0.4:
            flags.append(f"CRITICAL: Average confidence is {avg_conf:.0%}. "
                         "Most rows are being classified from weak signals.")

    # 3. Rule pass distribution
    if "rule_pass" in df.columns:
        pass_dist = df["rule_pass"].value_counts().sort_index()
        pass_names = {0: "Vendor Override", 1: "Commodity Code", 2: "Vendor Map",
                      3: "Category Meta", 4: "Keywords", 5: "Account/Fallback"}
        print(f"\n  Classification by pass:")
        for p, count in pass_dist.items():
            pct = count / n * 100
            name = pass_names.get(p, f"Pass {p}")
            marker = " ← WEAK" if p >= 4 and pct > 30 else ""
            print(f"    Pass {p} ({name:<20}): {count:>6,}  ({pct:.1f}%){marker}")

        # Check if too much is keyword-only
        kw_fallback = pass_dist.get(4, 0) + pass_dist.get(5, 0)
        kw_pct = kw_fallback / n * 100 if n else 0
        if kw_pct > 50:
            flags.append(f"WARNING: {kw_pct:.0f}% of rows classified by keywords/fallback only. "
                         "Commodity codes and vendor maps are underperforming.")

    # 4. Spend concentration — top vendor check
    if "Vendor Name" in df.columns and total_spend > 0:
        vendor_spend = df.groupby("Vendor Name")["_spend"].sum().sort_values(ascending=False)
        if len(vendor_spend) > 0:
            top_vendor = vendor_spend.index[0]
            top_pct = vendor_spend.iloc[0] / total_spend * 100
            if top_pct > 40:
                flags.append(f"WARNING: Top vendor '{top_vendor}' is {top_pct:.0f}% of total spend. "
                             "Verify this isn't a column mapping artifact.")

    # 5. Generic bucket concentration
    if "master_bucket" in df.columns and total_spend > 0 and "_spend" in df.columns:
        bucket_spend = df.groupby("master_bucket")["_spend"].sum()
        for bucket in ["Admin & Office", "Services"]:
            if bucket in bucket_spend:
                bpct = bucket_spend[bucket] / total_spend * 100
                if bpct > 35:
                    flags.append(f"WARNING: '{bucket}' captures {bpct:.0f}% of spend — "
                                 "may indicate over-broad keyword matching.")

    # Summary
    if flags:
        print(f"\n  ⚠ QA FLAGS ({len(flags)}):")
        for f in flags:
            print(f"    • {f}")
    else:
        print(f"\n  ✓ No QA flags raised.")

    print(SEP + "\n")
    return flags
