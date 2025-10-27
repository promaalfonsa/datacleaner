#!/usr/bin/env python3
"""
Excel Data Cleaner

This script implements the same filtering logic as cleanup.gs:
1. Filter ReturnStatus to keep only 'CR_Returned to Seller' & 'Waiting For RTM'
2. Remove CSState = 'Dhaka North' & 'Dhaka South'
3. Deduplicate by CustomerOrderCode (keep first)
4. For each CSState, keep only first 4 rows

Usage:
    python filter_data.py <input_file> <output_file>
    
Example:
    python filter_data.py "Raw xlsx data/October - 1 to 15 .xlsx" "filtered xlsx data/October - 1 to 15.xlsx"
"""

import pandas as pd
import sys
import os


def filter_excel(input_path, output_path, verbose=True):
    """
    Apply filtering rules to Excel file
    
    Args:
        input_path: Path to input Excel file
        output_path: Path to save filtered Excel file
        verbose: Print progress information
    
    Returns:
        DataFrame with filtered data
    """
    if verbose:
        print(f"Reading: {input_path}")
    
    # Read Excel file
    df = pd.read_excel(input_path, sheet_name=0)
    original_count = len(df)
    
    if verbose:
        print(f"Original rows: {original_count}")
    
    # Required columns check
    required_cols = ['ReturnStatus', 'CSState', 'CustomerOrderCode']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")
    
    # Step 1 & 2: Filter by ReturnStatus and CSState
    if verbose:
        print("\nStep 1 & 2: Filtering by ReturnStatus and CSState...")
    
    # Keep only specific ReturnStatus values
    return_status_keep = ['CR_Returned to Seller', 'Waiting For RTM']
    df_filtered = df[df['ReturnStatus'].isin(return_status_keep)]
    
    if verbose:
        print(f"  After ReturnStatus filter: {len(df_filtered)} rows")
    
    # Remove specific CSState values
    csstate_remove = ['Dhaka North', 'Dhaka South']
    df_filtered = df_filtered[~df_filtered['CSState'].isin(csstate_remove)]
    
    if verbose:
        print(f"  After CSState filter: {len(df_filtered)} rows")
    
    # Step 3: Deduplicate by CustomerOrderCode (keep first)
    if verbose:
        print("\nStep 3: Deduplicating by CustomerOrderCode...")
    
    df_deduped = df_filtered.drop_duplicates(subset='CustomerOrderCode', keep='first')
    
    if verbose:
        print(f"  After deduplication: {len(df_deduped)} rows")
    
    # Step 4: Keep only first 4 rows per CSState
    if verbose:
        print("\nStep 4: Keeping first 4 rows per CSState...")
    
    df_limited = df_deduped.groupby('CSState').head(4).reset_index(drop=True)
    
    if verbose:
        print(f"  After limiting per CSState: {len(df_limited)} rows")
    
    # Show distribution
    if verbose:
        print("\nFinal CSState distribution:")
        csstate_dist = df_limited['CSState'].value_counts()
        print(f"  Unique CSStates: {len(csstate_dist)}")
        print(f"  CSStates with 4 rows: {len(csstate_dist[csstate_dist == 4])}")
        
        print("\nFinal ReturnStatus distribution:")
        for status, count in df_limited['ReturnStatus'].value_counts().items():
            print(f"  - {status}: {count}")
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Save to Excel
    if verbose:
        print(f"\nSaving to: {output_path}")
    
    df_limited.to_excel(output_path, sheet_name='Query result', index=False)
    
    if verbose:
        print(f"Done! Saved {len(df_limited)} rows.")
        print(f"Reduction: {(1 - len(df_limited)/original_count)*100:.1f}%")
    
    return df_limited


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    if not os.path.exists(input_file):
        print(f"Error: Input file not found: {input_file}")
        sys.exit(1)
    
    try:
        filter_excel(input_file, output_file)
        print("\n" + "="*80)
        print("SUCCESS: File has been filtered!")
        print("="*80)
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
