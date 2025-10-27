# Data Cleaning Solution

## Problem Statement
The repository had an October file that needed to be filtered according to the same rules as the August file. The problem statement mentioned "September - 1 to 15.xlsx" but the actual file in the repository was "October - 1 to 15.xlsx".

## Solution Implemented

### Files Filtered
- **Input**: `Raw xlsx data/October - 1 to 15 .xlsx` (47,630 rows)
- **Output**: `filtered xlsx data/October - 1 to 15.xlsx` (63 rows)
- **Reduction**: 99.9%

### Filtering Rules Applied

The filtering follows the logic defined in `cleanup.gs`:

1. **ReturnStatus Filter**
   - Keep only rows where ReturnStatus = 'CR_Returned to Seller' OR 'Waiting For RTM'
   - Result: 680 rows → filtered down from various statuses

2. **CSState Removal**
   - Remove all rows where CSState = 'Dhaka North' OR 'Dhaka South'
   - Result: 305 rows → eliminated 14,097 Dhaka North/South rows from raw data

3. **Deduplication**
   - Remove duplicate entries based on CustomerOrderCode (keep first occurrence)
   - Result: 73 rows → eliminated 29,056 duplicates from raw data

4. **Limit per CSState**
   - Keep maximum 4 rows per unique CSState
   - Result: 63 rows → distributed across 39 unique CSStates

### Verification Results

All filtering rules have been verified and passed:

✓ **Rule 1**: Only contains 'CR_Returned to Seller' (38 rows) and 'Waiting For RTM' (25 rows)
✓ **Rule 2**: No 'Dhaka North' or 'Dhaka South' entries
✓ **Rule 3**: No duplicate CustomerOrderCode entries  
✓ **Rule 4**: No CSState has more than 4 rows (max = 4)

### Comparison with August File

| Metric | August | October |
|--------|--------|---------|
| Raw rows | 42,656 | 47,630 |
| Filtered rows | 126 | 63 |
| Reduction | 99.7% | 99.9% |
| Unique CSStates (filtered) | 49 | 39 |
| CSStates with max 4 rows | 18 | 3 |

### Column Structure

Both filtered files maintain the same 33-column structure from the raw data:
- CustomerOrderCode, SellerOrderCode, OrderDate, SellerName, CustomerName, CustomerPhone, CustomerId, ShippingCost, ApprovalDate, LastStatusUpdatedDate, CSState, CSCity, CSZone, CR_SourceHub, CR_DesHub, Price, Quantity, TotalItemPrice, Vertical, cat1, cat2, cat3, cat4, ProductName, DeliveryDate, DeliveryStatus, CustomerEmail, PaymentMethod, ReturnDate, ReturnStatus, ReturnReason, ReturnDescription, Delivery Attempt Number

### Google Apps Script Files

Two .gs files define the filtering and column sequencing logic:

1. **cleanup.gs** - Implements the 4 filtering rules described above
2. **column_sequence.gs** - Defines target column order (12 columns) for future use

## How to Use

To filter additional Excel files in the future:

1. Place the raw Excel file in `Raw xlsx data/` folder
2. Run the filtering script (or use Google Sheets with cleanup.gs)
3. The filtered file will be saved to `filtered xlsx data/` folder

## Technical Details

- **Language**: Python 3 with pandas and openpyxl libraries
- **Filtering Logic**: Matches cleanup.gs exactly
- **Output Format**: Excel (.xlsx) with original sheet name preserved ("Query result")
