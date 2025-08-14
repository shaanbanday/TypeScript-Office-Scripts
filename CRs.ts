// Shaan Banday – TypeScript

// Updates as of June 17th, 2025

function main(workbook: ExcelScript.Workbook) // define main entry point
{
  type TCell = string | number | boolean | Date; // possible cell value types

  // the code section below, defining constants, COULD probably be optimized more, but it works so I'll keep for now
  const targetName = "CRs"; // sheet to update
  const rawName = "RAW CRs"; // sheet with source data
  const combinedHeader = "CR# and Activity#"; // new combined-key header (search for this)
  const inRawHeader = "In Raw?"; // header for column that indicates row's presence in raw
  const headersToCopy: string[] = ["Days Due", "Activity Due Date", "Activity Lead", "CR Subject", "Activity Subject", "Activity Type", "Sig Level", "Activity Status", "CR Age", "Previous Extensions"]; // rest of the column headers

  // get the CRs sheet
  const target = workbook.getWorksheet(targetName);
  if (!target) // if the script cannot find the CRs sheet
  {
    throw new Error(`Sheet "${targetName}" not found.`); // catch the error and show in console (debugging step)
  }

  // find used range (headers + data) in CRs
  const usedTarget = target.getUsedRange();
  if (!usedTarget) // if the script sees CRs sheet as fully empty
  {
    throw new Error(`No data in "${targetName}".`); // catch the error and show in console (debugging step)
  }
  // record the position and size of that used range
  const firstRow = usedTarget.getRowIndex(); // zero-based row index of header
  const firstCol = usedTarget.getColumnIndex(); // zero-based column index of first header
  let colCount = usedTarget.getColumnCount(); // number of columns in that range. may grow if we insert a column
  const totalRows = usedTarget.getRowCount(); // total rows (header + data)

  // read header row into array
  const headerVals = target
    .getRangeByIndexes(firstRow, firstCol, 1, colCount)
    .getValues() as TCell[][];
  const headerRow: string[] = headerVals[0]
    .map(cell => (cell ?? "").toString().trim());

  // map header to column index in CRs
  const tgtIdx: Record<string, number> = {};
  [combinedHeader, ...headersToCopy, inRawHeader].forEach(hdr => {
    const idx = headerRow.indexOf(hdr);
    if (idx < 0) {
      throw new Error(`Header "${hdr}" missing in "${targetName}".`);
    }
    tgtIdx[hdr] = idx; // record mapping header to column position
  });
  const keyCol = tgtIdx[combinedHeader]; // we will use this to read/write the key

  // build a lookup of existing combined-key to row index in CRs
  const keyRowMap: Record<string, number> = {};
  if (totalRows > 1) {
    // read every value under the combinedHeader column
    const keyVals = target
      .getRangeByIndexes(firstRow + 1, firstCol + keyCol, totalRows - 1, 1)
      .getValues() as TCell[][];
    keyVals.forEach((r, i) => {
      const k = (r[0] ?? "").toString().trim();
      if (k) {
        keyRowMap[k] = firstRow + 1 + i;
      }
    });
  }
  // pointer to the next empty row (for truly new entries)
  let nextRow = firstRow + totalRows;

  // Access the source sheet (RAW CRs)
  const raw = workbook.getWorksheet(rawName);

  if (!raw) // if the raw sheet doesn't exist
  {
    throw new Error(`Sheet "${rawName}" not found.`);
  }

  const usedRaw = raw.getUsedRange();

  if (!usedRaw) return; // nothing to do if raw is empty
  const rawData = usedRaw.getValues() as TCell[][];
  if (rawData.length < 2) return; // skip if no data rows

  // Read and trim the header row in RAW CRs
  const rawHeaders: string[] = rawData[0]
    .map(cell => (cell ?? "").toString().trim());

  // Map each needed header in RAW to its column index
  const rawIdx: Record<string, number> = {};
  ["CR #", "Activity #", ...headersToCopy].forEach(hdr => {
    rawIdx[hdr] = rawHeaders.indexOf(hdr);
  });

  // set to track which keys appear during this run
  const seenKeys = new Set<string>();

  // Loop through each data row in RAW CRs
  for (let i = 1; i < rawData.length; i++) 
  {
    const row = rawData[i] as TCell[];
    // extract and trim the two parts of the key
    const crRaw = (row[rawIdx["CR #"]] ?? "").toString().trim();
    const actRaw = (row[rawIdx["Activity #"]] ?? "").toString().trim();
    if (!crRaw || !actRaw) continue; // skip if either is missing

    // build the combined key exactly as we expect in CRs
    const newKey = `${crRaw}-${actRaw}`;
    seenKeys.add(newKey); // record that we saw this key

    // decide whether to update an existing row or append a new one
    const existingRow = keyRowMap[newKey];
    const outRow = existingRow ?? nextRow;

    // Write the combined key into CRs
    target
      .getCell(outRow, firstCol + keyCol)
      .setValue(newKey);

    // Copy each configured column from raw → CRs
    headersToCopy.forEach(hdr => {
      const tgtC = firstCol + tgtIdx[hdr];  // column in CRs
      const rawC = rawIdx[hdr];             // column in RAW CRs
      const val = rawC >= 0 ? row[rawC] : ""; // fallback empty if missing
      target.getCell(outRow, tgtC).setValue(val);
    });

    // Mark this row as present in raw
    target
      .getCell(outRow, firstCol + tgtIdx[inRawHeader])
      .setValue("Yes");

    // If this was appended, add borders and update our map
    if (!existingRow) 
    {
      keyRowMap[newKey] = nextRow; // record row index

      const brRange = target.getRangeByIndexes(outRow, firstCol, 1, colCount);
      const fmt = brRange.getFormat();
      [
        ExcelScript.BorderIndex.edgeTop,
        ExcelScript.BorderIndex.edgeBottom,
        ExcelScript.BorderIndex.edgeLeft,
        ExcelScript.BorderIndex.edgeRight,
        ExcelScript.BorderIndex.insideVertical
      ].forEach(bi =>
        fmt.getRangeBorder(bi).setStyle(ExcelScript.BorderLineStyle.continuous)
      );

      nextRow++; // move to next available row
    }
  }

  // Mark any rows not seen this run as "No"
  for (let r = firstRow + 1; r < nextRow; r++) 
  {
    const key = (target.getCell(r, firstCol + keyCol).getValue() as string).trim();
    const flag = seenKeys.has(key) ? "Yes" : "No";
    target.getCell(r, firstCol + tgtIdx[inRawHeader]).setValue(flag);
  }

  // Final log for confirmation
  console.log(`Done: updated existing rows and appended new ones.`);
} // end of program
