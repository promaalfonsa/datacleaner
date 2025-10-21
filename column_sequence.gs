/**
 * Adds a menu item for the new column-sequencing step.
 * (If you already have an onOpen, merge the addItem lines.)
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cleanup')
    .addItem('Run requested filters', 'runCleanup') // <- keep if you already added earlier
    .addItem('Apply column sequence', 'applyColumnSequence') // <- new item
    .addToUi();
}


/**
 * Apply the exact column sequence, creating the missing ones empty and
 * removing any extra columns.
 *
 * Target sequence:
 * 1. Serial No. (Create)
 * 2. CustomerOrderCode (Existing)
 * 3. Agent Name (Create)
 * 4. Call Date (Create)
 * 5. CustomerId (Existing)
 * 6. CustomerName (Existing)
 * 7. CustomerPhone (Existing)
 * 8. CSState (Existing)
 * 9. Vertical (Existing)
 * 10. cat1 (Existing)
 * 11. TotalItemPrice (Existing)
 * 12. ProductName (Existing)
 */
function applyColumnSequence() {
  const SHEET = SpreadsheetApp.getActiveSheet();
  const HEADER_ROW = 1;


  const TARGET_HEADERS = [
    'Serial No.',
    'CustomerOrderCode',
    'Agent Name',
    'Call Date',
    'CustomerId',
    'CustomerName',
    'CustomerPhone',
    'CSState',
    'Vertical',
    'cat1',
    'TotalItemPrice',
    'ProductName',
  ];


  const lastRow = SHEET.getLastRow();
  const lastCol = SHEET.getLastColumn();


  // If no rows at all, just write headers in the target sequence
  if (lastRow === 0) {
    SHEET.getRange(1, 1, 1, TARGET_HEADERS.length).setValues([TARGET_HEADERS]);
    // Remove any extra columns beyond the target length
    if (lastCol > TARGET_HEADERS.length) {
      SHEET.deleteColumns(TARGET_HEADERS.length + 1, lastCol - TARGET_HEADERS.length);
    }
    return;
  }


  // Read current data (headers + rows)
  const values = SHEET.getRange(HEADER_ROW, 1, Math.max(1, lastRow - HEADER_ROW + 1), lastCol).getValues();
  const header = values[0].map(h => String(h).trim());
  const data = values.slice(1);
  const numRows = Math.max(0, values.length - 1);


  // Map header -> column index
  const idxMap = new Map();
  header.forEach((h, i) => idxMap.set(h, i));


  // Build the new matrix: header row + reordered/created columns
  const output = [];
  output.push(TARGET_HEADERS.slice());


  // For each target header, either copy existing column or create blank
  for (const target of TARGET_HEADERS) {
    let colValues;
    if (idxMap.has(target)) {
      const colIndex = idxMap.get(target);
      colValues = data.map(row => row[colIndex]);
    } else {
      // Create an empty column of same height
      colValues = Array.from({ length: numRows }, () => '');
    }
    // Append column to output (we'll transpose at the end)
    // We'll collect as columns; later we convert to rows.
    if (!output._cols) output._cols = [];
    output._cols.push([target, ...colValues]);
  }


  // Transpose columns -> rows for setValues
  const finalRows = [];
  for (let r = 0; r < (numRows + 1); r++) {
    const row = [];
    for (let c = 0; c < output._cols.length; c++) {
      row.push(output._cols[c][r]);
    }
    finalRows.push(row);
  }


  // Clear entire sheet and write only the target columns
  SHEET.clearContents();
  SHEET.getRange(1, 1, finalRows.length, TARGET_HEADERS.length).setValues(finalRows);


  // If there are leftover columns to the right, delete them
  const newLastCol = SHEET.getLastColumn();
  if (newLastCol > TARGET_HEADERS.length) {
    SHEET.deleteColumns(TARGET_HEADERS.length + 1, newLastCol - TARGET_HEADERS.length);
  }
}


/* ------------------------------------------------------------------
   Optional: keep the cleanup function from earlier here for convenience.
   If you already pasted it before, you can remove this duplicate.
------------------------------------------------------------------ */
function runCleanup() {
  const SHEET_NAME = SpreadsheetApp.getActiveSheet().getName();
  const HEADER_ROW = 1;


  const RETURN_STATUS_KEEP = new Set(['CR_Returned to Seller', 'Waiting For RTM']);
  const CSSTATE_REMOVE = new Set(['Dhaka North', 'Dhaka South']);
  const MAX_PER_CSSTATE = 4;


  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(Sheet "${SHEET_NAME}" not found.);


  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < HEADER_ROW + 1) return;


  const allValues = sh.getRange(HEADER_ROW, 1, lastRow - HEADER_ROW + 1, lastCol).getValues();
  const header = allValues[0].map(v => String(v).trim());
  const rows = allValues.slice(1);


  const idx = name => header.indexOf(name);
  const colReturnStatus = idx('ReturnStatus');
  const colCSState = idx('CSState');
  const colCustomerOrderCode = idx('CustomerOrderCode');


  const missing = [];
  if (colReturnStatus === -1) missing.push('ReturnStatus');
  if (colCSState === -1) missing.push('CSState');
  if (colCustomerOrderCode === -1) missing.push('CustomerOrderCode');
  if (missing.length) throw new Error(Missing required header(s): ${missing.join(', ')});


  const filtered = rows.filter(r => {
    const rs = String(r[colReturnStatus]).trim();
    const cs = String(r[colCSState]).trim();
    return RETURN_STATUS_KEEP.has(rs) && !(cs && (cs === 'Dhaka North' || cs === 'Dhaka South'));
  });


  const seen = new Set();
  const deduped = [];
  for (const r of filtered) {
    const key = String(r[colCustomerOrderCode]).trim();
    if (!seen.has(key)) {
      seen.add(key);
      deduped.push(r);
    }
  }


  const keptCount = new Map();
  const limited = [];
  for (const r of deduped) {
    const cs = String(r[colCSState]).trim();
    const count = keptCount.get(cs) || 0;
    if (count < MAX_PER_CSSTATE) {
      limited.push(r);
      keptCount.set(cs, count + 1);
    }
  }


  const dataRowCount = Math.max(0, lastRow - HEADER_ROW);
  if (dataRowCount > 0) {
    sh.getRange(HEADER_ROW + 1, 1, dataRowCount, lastCol).clearContent();
  }


  if (limited.length > 0) {
    sh.getRange(HEADER_ROW + 1, 1, limited.length, lastCol).setValues(
      limited.map(row => {
        if (row.length === lastCol) return row;
        const copy = row.slice();
        while (copy.length < lastCol) copy.push('');
        return copy.slice(0, lastCol);
      })
    );
  }
}