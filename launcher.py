"""
Launcher for Procurement Categorization Pipeline

Works two ways:
  1. Double-click: prompts for input file/folder interactively
  2. Command line:
       python launcher.py --input data.csv
       python launcher.py --input "C:\\path\\to\\periods"    (processes all CSVs in folder)
       python launcher.py --input data.csv --output-dir ./out
       python launcher.py --input data.csv --use-categorized-output existing.csv
"""

import argparse
import csv
import os
import sys
import glob
import subprocess
import traceback
import pandas as pd

# Add current directory to path
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

# Import from existing modules
from categorization import categorize_dataframe


def resolve_input_files(path):
    """Resolve a path to a list of CSV/Excel files.

    If path is a file, return it as a single-item list.
    If path is a directory, return all CSV and Excel files in it.

    Returns:
        tuple: (file_list, error_message)
    """
    path = path.strip().strip('"').strip("'")

    if os.path.isfile(path):
        return [path], None

    if os.path.isdir(path):
        files = []
        for pat in ("*.csv", "*.xlsx", "*.xls"):
            files.extend(glob.glob(os.path.join(path, pat)))
        if not files:
            return [], f"No CSV or Excel files found in folder: {path}"
        files.sort()
        return files, None

    return [], f"Path not found: {path}"


def read_file(filepath):
    """Read a CSV or Excel file into a DataFrame."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(filepath)
    try:
        return pd.read_csv(filepath, low_memory=False)
    except UnicodeDecodeError:
        return pd.read_csv(filepath, low_memory=False, encoding="latin-1")


def validate_categorized_output(filepath):
    """Validate pre-existing categorized output file.

    Returns:
        tuple: (is_valid, error_message)
    """
    if not os.path.exists(filepath):
        return False, f"File not found: {filepath}"

    try:
        df = pd.read_csv(filepath, nrows=1000)

        required_columns = ["master_bucket", "Extended Price"]
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            return False, f"Missing required columns: {missing_columns}. File must contain: master_bucket, Extended Price"
        if len(df) == 0:
            return False, "File is empty"

        return True, None

    except pd.errors.EmptyDataError:
        return False, "File is empty or not a valid CSV"
    except Exception as e:
        return False, f"Error reading file: {e}"


def run_pipeline(input_path, output_dir, use_categorized_output=None, skip_aggregation=False, skip_excel=False):
    """Run the full categorization pipeline.

    Args:
        input_path: Input CSV file path OR folder containing CSVs
        output_dir: Output directory
        use_categorized_output: Optional path to pre-existing categorized output
        skip_aggregation: Whether to skip the aggregation step
        skip_excel: Whether to skip the Excel generation step

    Returns:
        tuple: (success, message)
    """
    # Resolve input to file list
    input_files, error = resolve_input_files(input_path)
    if error:
        return False, error

    # Validate pre-existing categorized output if provided
    if use_categorized_output:
        is_valid, error = validate_categorized_output(use_categorized_output)
        if not is_valid:
            return False, error

    # Create output directory if needed
    os.makedirs(output_dir, exist_ok=True)

    # Step 1: Categorization (skip if pre-existing)
    if use_categorized_output:
        print(f"Using pre-existing categorized output: {use_categorized_output}")
        categorized_file = use_categorized_output
    else:
        print(f"Running categorization on {len(input_files)} file(s)...")
        categorized_file = os.path.join(output_dir, "categorized_output.csv")
        try:
            dfs = []
            for f in input_files:
                print(f"  Reading: {os.path.basename(f)}")
                df = read_file(f)
                print(f"    -> {len(df):,} rows")
                dfs.append(df)

            combined = pd.concat(dfs, ignore_index=True)
            print(f"  Combined: {len(combined):,} total rows")

            results = categorize_dataframe(combined)
            results.to_csv(categorized_file, index=False, encoding="utf-8-sig",
                           quoting=csv.QUOTE_ALL)
            # Write pickle for reliable Excel builder hand-off (no CSV parsing issues)
            results.to_pickle(categorized_file.replace(".csv", ".pkl"))
            print(f"  Saved categorized output to: {categorized_file}")

            # Print summary
            print("\n  -- Bucket Summary --")
            summary = results["master_bucket"].value_counts()
            for bucket, count in summary.items():
                pct = count / len(results) * 100
                print(f"    {bucket:<42} {count:>6,}  ({pct:.1f}%)")

        except Exception as e:
            return False, f"Categorization failed: {e}"

    # Step 2: Aggregation (skip if requested)
    if skip_aggregation:
        print("Skipping aggregation step...")
    else:
        print("Running aggregation step...")
        try:
            agg_output = os.path.join(output_dir, "spend_by_bucket.csv")
            result = subprocess.run(
                [sys.executable, os.path.join(script_dir, "aggregate_spend.py"), categorized_file, agg_output],
                capture_output=True,
                text=True,
                check=True
            )
            print(result.stdout)
        except subprocess.CalledProcessError as e:
            return False, f"Aggregation failed: {e.stderr}"

    # Step 3: Excel generation (skip if requested)
    if skip_excel:
        print("Skipping Excel generation step...")
    else:
        print("Running Excel generation step...")
        try:
            excel_output = os.path.join(output_dir, "Procurement_Detail_Breakdown.xlsx")
            result = subprocess.run(
                [sys.executable, os.path.join(script_dir, "build_detail_excel_v2.py"), categorized_file, excel_output],
                capture_output=True,
                text=True,
                check=True
            )
            print(result.stdout)
        except subprocess.CalledProcessError as e:
            return False, f"Excel generation failed: {e.stderr}"

    return True, "Pipeline completed successfully!"


def main():
    parser = argparse.ArgumentParser(description="Launcher for Procurement Categorization Pipeline")
    parser.add_argument("--input", help="Input CSV file path or folder containing CSVs")
    parser.add_argument("--output-dir", default=".", help="Output directory (default: current directory)")
    parser.add_argument("--use-categorized-output", help="Use pre-existing categorized output file (skip categorization step)")
    parser.add_argument("--skip-aggregation", action="store_true", help="Skip aggregation step")
    parser.add_argument("--skip-excel", action="store_true", help="Skip Excel generation step")

    args = parser.parse_args()

    # If no --input was provided, prompt interactively
    if not args.input:
        print("=" * 60)
        print("  Procurement Categorization Pipeline")
        print("=" * 60)
        print()
        print("No --input argument provided.")
        print("You can enter a path to a single CSV file")
        print("OR a folder containing multiple CSVs.")
        print()
        args.input = input("Enter path to CSV file or folder: ").strip().strip('"')
        if not args.input:
            print("ERROR: No input path provided.")
            input("\nPress Enter to exit...")
            return

        # Also prompt for output dir if running interactively
        out = input(f"Output directory (press Enter for current dir): ").strip().strip('"')
        if out:
            args.output_dir = out
        print()

    try:
        success, message = run_pipeline(
            args.input,
            args.output_dir,
            args.use_categorized_output,
            args.skip_aggregation,
            args.skip_excel
        )

        if success:
            print(f"\n{message}")
        else:
            print(f"\nERROR: {message}")

    except Exception as e:
        print(f"\nERROR: {e}")
        traceback.print_exc()

    # Keep window open if double-clicked (interactive terminal)
    if sys.stdin and sys.stdin.isatty():
        input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
