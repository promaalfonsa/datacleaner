# Data Cleaner - Customer Survey Data Population

This repository contains scripts to populate missing customer information in survey data files.

## Problem Statement

The "Customer Survery Partially Completed.xlsx" file in the "need to fix" folder had incomplete data in the "Pool 3 Return Pickup" sheet. The following columns were empty:
- Buyer ID
- Name
- Phone No.
- District
- Category 1
- Category 2
- Order Value
- Product Name

The complete data for these fields exists in "Customer Survery Not Touched but correct informations.xlsx" file.

## Solution

A Python script (`fill_pool3_data.py`) was created to:
1. Read the source file with correct customer information
2. Read the target file's Pool 3 sheet with missing data
3. Match records by Order Number (CustomerOrderCode)
4. Populate all 8 missing columns with data from the source file
5. Save the updated data back to the Excel file

### Column Mapping

| Target Column (Pool 3) | Source Column | Description |
|------------------------|---------------|-------------|
| Buyer ID | CustomerId | Customer unique identifier |
| Name | CustomerName | Customer full name |
| Phone No. | CustomerPhone | Customer phone number |
| District | CSState | Customer district/state |
| Category 1 | Vertical | Main product category |
| Category 2 | cat1 | Sub-category |
| Order Value | TotalItemPrice | Total order price |
| Product Name | ProductName | Full product name |

## Usage

### Running the Data Population Script

```bash
python3 fill_pool3_data.py
```

This will:
- Read both xlsx files from the "need to fix" folder
- Match all 95 orders by Order Number
- Fill in the 8 empty columns
- Save the updated file

### Running the Test

To verify the data population was successful:

```bash
python3 test_data_population.py
```

The test verifies:
- All required columns are present
- No null values remain in populated columns
- Correct number of records (95)
- Data accuracy matches source file 100%

## Results

✅ **Successfully populated all 95 rows** in the Pool 3 Return Pickup sheet:
- Buyer ID: 95/95 filled (100%)
- Name: 95/95 filled (100%)
- Phone No.: 95/95 filled (100%)
- District: 95/95 filled (100%)
- Category 1: 95/95 filled (100%)
- Category 2: 95/95 filled (100%)
- Order Value: 95/95 filled (100%)
- Product Name: 95/95 filled (100%)

All data has been cross-verified for accuracy with 0 errors.

## Requirements

- Python 3.x
- pandas
- openpyxl

Install dependencies:
```bash
pip install pandas openpyxl
```

## File Structure

```
datacleaner/
├── need to fix/
│   ├── Customer Survery Not Touched but correct informations.xlsx (source)
│   └── Customer Survery Partially Completed.xlsx (target - updated)
├── fill_pool3_data.py (main script)
├── test_data_population.py (test script)
└── README.md (this file)
```
