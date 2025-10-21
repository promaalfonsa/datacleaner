#!/usr/bin/env python3
"""
Test script to verify that the data population in Pool 3 Return Pickup sheet is correct.
"""

import pandas as pd
import os
import sys

def test_pool3_data_population():
    """Test that Pool 3 sheet has all required columns filled correctly."""
    
    base_path = "need to fix"
    source_file = os.path.join(base_path, "Customer Survery Not Touched but correct informations.xlsx")
    target_file = os.path.join(base_path, "Customer Survery Partially Completed.xlsx")
    
    print("Testing Pool 3 Return Pickup data population...")
    
    # Read source and target files
    df_source = pd.read_excel(source_file)
    df_target = pd.read_excel(target_file, sheet_name='Pool 3 Return Pickup', header=3)
    
    # Test 1: Check that all required columns are present
    required_columns = ['Order No.', 'Buyer ID', 'Name', 'Phone No.', 'District', 
                       'Category 1', 'Category 2', 'Order Value', 'Product Name']
    
    for col in required_columns:
        assert col in df_target.columns, f"Missing column: {col}"
    print("✓ Test 1 passed: All required columns present")
    
    # Test 2: Check that no values are null in the required columns (except Order No.)
    check_columns = required_columns[1:]  # Skip Order No.
    
    for col in check_columns:
        null_count = df_target[col].isnull().sum()
        assert null_count == 0, f"Column {col} has {null_count} null values"
    print("✓ Test 2 passed: No null values in populated columns")
    
    # Test 3: Verify data count matches
    assert len(df_target) == 95, f"Expected 95 rows, got {len(df_target)}"
    assert len(df_source) == 95, f"Expected 95 source rows, got {len(df_source)}"
    print("✓ Test 3 passed: Correct number of records (95)")
    
    # Test 4: Verify data accuracy by checking a few sample orders
    source_dict = {}
    for _, row in df_source.iterrows():
        order_no = int(row['CustomerOrderCode'])
        source_dict[order_no] = row
    
    sample_orders = [25091619140254, 25091622473323, 25091638845882]
    
    for order_no in sample_orders:
        target_row = df_target[df_target['Order No.'] == order_no].iloc[0]
        source_row = source_dict[order_no]
        
        assert int(target_row['Buyer ID']) == int(source_row['CustomerId'])
        assert target_row['Name'] == source_row['CustomerName']
        assert int(target_row['Phone No.']) == int(source_row['CustomerPhone'])
        assert target_row['District'] == source_row['CSState']
        assert target_row['Category 1'] == source_row['Vertical']
        assert target_row['Category 2'] == source_row['cat1']
        assert int(target_row['Order Value']) == int(source_row['TotalItemPrice'])
        assert target_row['Product Name'] == source_row['ProductName']
    
    print("✓ Test 4 passed: Sample data matches source file exactly")
    
    # Test 5: Verify all orders match
    errors = 0
    for _, target_row in df_target.iterrows():
        order_no = int(target_row['Order No.'])
        if order_no in source_dict:
            source_row = source_dict[order_no]
            
            if int(target_row['Buyer ID']) != int(source_row['CustomerId']):
                errors += 1
            if target_row['Name'] != source_row['CustomerName']:
                errors += 1
            if int(target_row['Phone No.']) != int(source_row['CustomerPhone']):
                errors += 1
            if target_row['District'] != source_row['CSState']:
                errors += 1
            if target_row['Category 1'] != source_row['Vertical']:
                errors += 1
            if target_row['Category 2'] != source_row['cat1']:
                errors += 1
            if int(target_row['Order Value']) != int(source_row['TotalItemPrice']):
                errors += 1
            if target_row['Product Name'] != source_row['ProductName']:
                errors += 1
    
    assert errors == 0, f"Found {errors} data mismatches"
    print("✓ Test 5 passed: All 95 records match source file perfectly")
    
    print("\n" + "=" * 60)
    print("ALL TESTS PASSED! ✓")
    print("=" * 60)
    print(f"Successfully verified {len(df_target)} records")
    print("All required columns are populated correctly")
    return True

if __name__ == "__main__":
    try:
        test_pool3_data_population()
        sys.exit(0)
    except AssertionError as e:
        print(f"\n✗ Test failed: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ Error: {e}")
        sys.exit(1)
