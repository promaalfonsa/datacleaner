/**
 * Adds a simple menu to run the cleanup with one click.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cleanup')
    .addItem('Run requested filters', 'runCleanup')
    .addToUi();
}


/**
 * Runs all requested steps, in order:
 * 1) Filter ReturnStatus to keep only CR_Returned to Seller & Waiting For RTM
 * 2) Remove CSState = Dhaka North & Dhaka South
 * 3) Deduplicate by CustomerOrderCode (keep first)
 * 4) For each CSState, keep only first 4 rows
 * Writes results back to the same sheet, below the header.
 */
function runCleanup() {
  // === CONFIG ===
  const SHEET_NAME = SpreadsheetApp.getActiveSheet().getName(); // change to a fixed name if you want
  const HEADER_ROW = 1;


  // Values to keep/remove
  const RETURN_STATUS_KEEP = new Set(['CR_Returned to Seller', 'Waiting For RTM']);
  const CSSTATE_REMOVE = new Set(['Dhaka North', 'Dhaka South']);
  const MAX_PER_CSSTATE = 4;


  // === LOAD DATA ===
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(Sheet "${SHEET_NAME}" not found.);


  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < HEADER_ROW + 1) {
    // No data
    return;
  }


  const allValues = sh.getRange(HEADER_ROW, 1, lastRow - HEADER_ROW + 1, lastCol).getValues();
  const header = allValues[0].map(v => String(v).trim());
  const rows = allValues.slice(1);


  // === MAP COLUMNS ===
  const idx = name => header.indexOf(name);
  const colReturnStatus = idx('ReturnStatus');
  const colCSState = idx('CSState');
  const colCustomerOrderCode = idx('CustomerOrderCode');


  const missing = [];
  if (colReturnStatus === -1) missing.push('ReturnStatus');
  if (colCSState === -1) missing.push('CSState');
  if (colCustomerOrderCode === -1) missing.push('CustomerOrderCode');
  if (missing.length) {
    throw new Error(Missing required header(s): ${missing.join(', ')});
  }


  // === 1 & 2: FILTER ===
  const filtered = rows.filter(r => {
    const rs = String(r[colReturnStatus]).trim();
    const cs = String(r[colCSState]).trim();
    const keepRS = RETURN_STATUS_KEEP.has(rs);
    const removeCS = CSSTATE_REMOVE.has(cs);
    return keepRS && !removeCS;
  });


  // === 3: DEDUP BY CustomerOrderCode (keep first) ===
  const seen = new Set();
  const deduped = [];
  for (const r of filtered) {
    const key = String(r[colCustomerOrderCode]).trim();
    if (!seen.has(key)) {
      seen.add(key);
      deduped.push(r);
    }
  }


  // === 4: LIMIT TO 4 PER CSState ===
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


  // === WRITE BACK ===
  // Clear existing data rows
  const dataRowCount = Math.max(0, lastRow - HEADER_ROW);
  if (dataRowCount > 0) {
    sh.getRange(HEADER_ROW + 1, 1, dataRowCount, lastCol).clearContent();
  }


  // Write results (if any)
  if (limited.length > 0) {
    sh.getRange(HEADER_ROW + 1, 1, limited.length, lastCol).setValues(
      // Ensure each row has exactly lastCol columns (pad if needed)
      limited.map(row => {
        if (row.length === lastCol) return row;
        const copy = row.slice();
        while (copy.length < lastCol) copy.push('');
        return copy.slice(0, lastCol);
      })
    );
  }
}