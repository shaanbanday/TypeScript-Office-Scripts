// Shaan Banday – TypeScript  
// Updates as of June 19th, 2025

function main(workbook: ExcelScript.Workbook) // define main entry point
{
  type TCell = string | number | boolean | Date; // possible cell value types

  // the code section below, defining constants, COULD probably be optimized more, but it works so I'll keep for now
  const targetName = "Turnover"; // sheet we will update
  const rawName = "RAW Turnover"; // sheet we read fresh data from
  const keyHeader = "EC"; // unique identifier column
  const inRawHeader = "In Raw?"; // existing flag column in Turnover
  const headersToCopy: string[] = ["Days Due", "EC", "ATE Task", "ATE Date", "Mod Type", "EC Description", "Owner Name", "FTL Name", "EC Status", "Outage", "Timeline Status", "Status Date"]; // rest of the column headers

  // get the Turnover sheet
  const target = workbook.getWorksheet(targetName);
  if (!target) // if the script cannot find the Turnover sheet
  {
    throw new Error(`Sheet "${targetName}" not found.`); // catch the error and show in console (debugging step)
  }

  // find used range (headers + data) in Turnover
  const usedTarget = target.getUsedRange();
  if (!usedTarget) // if the script sees Turnover sheet as fully empty
  {
    throw new Error(`No data in "${targetName}".`); // catch the error and show in console (debugging step)
  }
  
  const firstRow = usedTarget.getRowIndex();    // index of header row
  const firstCol = usedTarget.getColumnIndex(); // index of first column
  let colCount = usedTarget.getColumnCount(); // how many columns
  const totalRows = usedTarget.getRowCount();    // header + data rows

  // read and header row into array
  const headerVals = target.getRangeByIndexes(firstRow, firstCol, 1, colCount).getValues() as TCell[][];
  const headerRow: string[] = headerVals[0].map(cell => (cell ?? "").toString().trim());

  // map each header to column index
  const tgtIdx: Record<string, number> = {};
  [...headersToCopy, inRawHeader].forEach(hdr => 
  {
    const idx = headerRow.indexOf(hdr);
    if (idx < 0) 
    {
      throw new Error(`Header "${hdr}" missing in "${targetName}".`);
    }
    tgtIdx[hdr] = idx;
  });

  const keyCol = tgtIdx[keyHeader]; // column index of the EC key

  // build a lookup of existing EC to row index for in‐place updates
  const keyRowMap: Record<string, number> = {};
  if (totalRows > 1) 
  {
    const keyVals = target.getRangeByIndexes(firstRow + 1, firstCol + keyCol, totalRows - 1, 1).getValues() as TCell[][];
    keyVals.forEach((r, i) => 
    {
      const k = (r[0] ?? "").toString().trim();
      if (k) keyRowMap[k] = firstRow + 1 + i;
    });
  }
  let nextRow = firstRow + totalRows; // append any new data to next free row

  // get the RAW Turnover sheet
  const raw = workbook.getWorksheet(rawName);
  if (!raw) // if the script cannot find the raw Turnover sheet
  {
    throw new Error(`Sheet "${rawName}" not found.`); // catch the error and show in console (debugging step)
  }

  // find used range (headers + data) in RAW
  const usedRaw = raw.getUsedRange();
  if (!usedRaw) return; // nothing to do if raw is empty
  const rawData = usedRaw.getValues() as TCell[][];
  if (rawData.length < 2) return;     // skip if no data rows
  const rawHeaders: string[] = rawData[0].map(cell => (cell ?? "").toString().trim());

  // map each needed RAW header to its column index
  const rawIdx: Record<string, number> = {};
  headersToCopy.forEach(hdr => 
  {
    rawIdx[hdr] = rawHeaders.indexOf(hdr);
  });
  const seenKeys = new Set<string>();

  // Loop through each data row in RAW Turnover
  for (let i = 1; i < rawData.length; i++) 
  {
    const row = rawData[i] as TCell[];

    // Extract and trim the EC key. skip if blank
    const rawKey = (row[rawIdx[keyHeader]] ?? "").toString().trim();
    if (!rawKey) 
    {
      continue; // never append blank‐key rows
    }
    seenKeys.add(rawKey);

    // Determine if we update in place or append new
    const existingRow = keyRowMap[rawKey];
    const outRow = existingRow !== undefined ? existingRow : nextRow;

    // Copy each configured column from RAW → Turnover
    headersToCopy.forEach(hdr => 
    {
      const tgtC = firstCol + tgtIdx[hdr]; // target column
      const rawC = rawIdx[hdr]; // raw column
      const val = rawC >= 0 ? row[rawC] : ""; // fallback to empty
      target.getCell(outRow, tgtC).setValue(val);
    });

    // mark this row as present in RAW
    target.getCell(outRow, firstCol + tgtIdx[inRawHeader]).setValue("Yes");

    // if newly appended, add borders and update our map
    if (existingRow === undefined) 
    {
      keyRowMap[rawKey] = nextRow;
      const brRange = target.getRangeByIndexes(outRow, firstCol, 1, colCount);
      const brFmt = brRange.getFormat();
      [
        ExcelScript.BorderIndex.edgeTop,
        ExcelScript.BorderIndex.edgeBottom,
        ExcelScript.BorderIndex.edgeLeft,
        ExcelScript.BorderIndex.edgeRight,
        ExcelScript.BorderIndex.insideVertical
      ].forEach(bi =>
        brFmt.getRangeBorder(bi).setStyle(ExcelScript.BorderLineStyle.continuous)
      );
      nextRow++;
    }
  }
  // mark any Turnover rows not seen this run as "No"
  for (let r = firstRow + 1; r < nextRow; r++) 
  {
    const cellVal = target.getCell(r, firstCol + keyCol).getValue();
    const ecKey = (cellVal ?? "").toString().trim();
    const flag = seenKeys.has(ecKey) ? "Yes" : "No";
    target
      .getCell(r, firstCol + tgtIdx[inRawHeader])
      .setValue(flag);
  }
  console.log(`Done: in-place updates + new additions to "${targetName}".`);
} // end of program
