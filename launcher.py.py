"""
Launcher for Procurement Categorization Pipeline

Features:
  - Full pipeline: categorization → aggregation → excel
  - Use pre-existing categorized output (skip categorization step)
  - Skip individual pipeline steps

Usage:
  python launcher.py --input data.csv
  python launcher.py --input data.csv --use-categorized-output existing.csv
  python launcher.py --input data.csv --use-categorized-output existing.csv --skip-excel
"""

import argparse
import os
import sys
import subprocess
import pandas as pd
from datetime import datetime

# Add current directory to path
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

# Import from existing modules
from categorization import categorize_dataframe


def validate_categorized_output(filepath):
    """Validate pre-existing categorized output file.
    
    Args:
        filepath: Path to the categorized output file
        
    Returns:
        tuple: (is_valid, error_message)
    """
    # Check file exists
    if not os.path.exists(filepath):
        return False, f"File not found: {filepath}"
    
    try:
        # Try to read the file
        df = pd.read_csv(filepath, nrows=1000)  # Read first 100 rows for validation
        
        # Check required columns
        required_columns = ["master_bucket", "Extended Price"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            return False, f"Missing required columns: {missing_columns}. File must contain: master_bucket, Extended Price"
        # Check non-empty
        if len(df) == 0:
            return False, "File is empty"
        
        # Check data types (basic validation)
        if not pd.api.types.is_string_dtype(df["master_bucket"]) or not pd.api.types.is_numeric_dtype(df["Extended Price"]):
            return False, "Invalid data types. Expected: master_bucket (string), Extended Price (numeric)"
        
        return True, None
    
    except pd.errors.EmptyDataError:
        return False, "File is empty or not a valid CSV"
    except Exception as e:
        return False, f"Error reading file: {e}"


def run_pipeline(input_file, output_dir, use_categorized_output=None, skip_aggregation=False, skip_excel=False):
    """Run the full categorization pipeline.
    
    Args:
        input_file: Input CSV file path
        output_dir: Output directory
        use_categorized_output: Optional path to pre-existing categorized output
        skip_aggregation: Whether to skip the aggregation step
        skip_excel: Whether to skip the Excel generation step
        
    Returns:
        tuple: (success, message)
    """
    # Validate inputs
    if not os.path.exists(input_file):
        return False, f"Input file not found: {input_file}"
       
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
        print("Running categorization step...")
        categorized_file = os.path.join(output_dir, "categorized_output.csv")
        # Run categorization (from run_categorization.py logic)
        try:
            df = pd.read_csv(input_file, low_memory=False)
            results = categorize_dataframe(df)
            # Save categorized output
            results.to_csv(categorized_file, index=False)
            print(f"Saved categorized output to: {categorized_file}")
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
    parser.add_argument("--input", required=True, help="Input CSV file path")
    parser.add_argument("--output-dir", default=".", help="Output directory (default: current directory)")
    parser.add_argument("--use-categorized-output", help="Use pre-existing categorized output file (skip categorization step)")
    parser.add_argument("--skip-aggregation", action="store_true", help="Skip aggregation step")
    parser.add_argument("--skip-excel", action="store_true", help="Skip Excel generation step")

    args = parser.parse_args()

    # Run pipeline
    success, message = run_pipeline(args.input, args.output_dir, args.use_categorized_output, args.skip_aggregation, args.skip_excel)

    if not success:
        print(f"ERROR: {message}")
    print(message)


if __name__ == "__main__":
    main()
