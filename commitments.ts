// Shaan Banday – TypeScript
// Updates as of June 16th, 2025

function main(workbook: ExcelScript.Workbook) // define main entry point
{
  type TCell = string | number | boolean | Date; // possible cell value types

  // the code section below, defining constants, COULD probably be optimized more, but it works so I'll keep for now
  const targetName = "Commitments"; // sheet to update
  const rawName = "RAW Commitments"; // sheet with source data
  const combinedHeader = "Project Number - Activity ID"; // new combined-key header (search for this)
  const projectNameHeader = "Project Name"; // remainder of raw project name
  const inRawHeader = "In Raw?"; // header for column that indicates row's presence in raw
  const headersToCopy = ["Activity Name", "Finish", "Commit. Date", "Variance", "OE", "PCS", "Commit. Type", "Status"]; // rest of the column headers. pretty simple

  // get the Commitments sheet
  const target = workbook.getWorksheet(targetName);
  if (!target) // if the script cannot find the commitments sheet
  {
    throw new Error(`Sheet "${targetName}" not found.`); // catch the error and show in console (debugging step)
  }

  // find used range (headers + data) in Commitments
  const usedTarget = target.getUsedRange();
  if (!usedTarget) // if the script sees commitments sheet as fully empty
  {
    throw new Error(`No data in "${targetName}".`); // catch the error and show in console (debugging step)
  }

  const firstRow = usedTarget.getRowIndex(); // zero-based
  const firstCol = usedTarget.getColumnIndex(); // zero-based
  let colCount = usedTarget.getColumnCount(); // may grow if we insert a column
  const totalRows = usedTarget.getRowCount(); // includes header row

  // read header row into array
  let headerVals = target
    .getRangeByIndexes(firstRow, firstCol, 1, colCount)
    .getValues() as TCell[][];
  let headerRow = headerVals[0].map(cell => (cell ?? "").toString());

  // map header to column index in Commitments
  const tgtIdx: Record<string, number> = {};
  [combinedHeader, projectNameHeader, inRawHeader, ...headersToCopy].forEach(hdr => 
  {
    const idx = headerRow.indexOf(hdr);
    if (idx < 0) 
    {
      throw new Error(`Header "${hdr}" missing in "${targetName}".`);
    }
    tgtIdx[hdr] = idx;
  });

  // use the combinedKey Project Number + Activity ID as the variable you will use to look up
  const keyRowMap: Record<string, number> = {};
  if (totalRows > 1)
  {
    const keyVals = target
      .getRangeByIndexes(firstRow + 1, firstCol + tgtIdx[combinedHeader], totalRows - 1, 1)
      .getValues() as TCell[][];
    keyVals.forEach((r, i) => 
    {
      const k = (r[0] ?? "").toString().trim();
      if (k) keyRowMap[k] = firstRow + 1 + i;
    });
  }
  let nextRow = firstRow + totalRows; // append any new data to next free row

  // get the RAW Commitments sheet
  const raw = workbook.getWorksheet(rawName);
  if (!raw) // if the script cannot find the raw commitments sheet
  {
    throw new Error(`Sheet "${rawName}" not found.`); // catch the error and show in console (debugging step)
  }

  // find used range (headers + data) in RAW
  const usedRaw = raw.getUsedRange();
  if (!usedRaw) return; // nothing to do if empty
  const rawData = usedRaw.getValues() as TCell[][];
  if (rawData.length < 2) return; // skip if no data rows

  // map needed raw headers to column indices
  const rawHeaders = rawData[0].map(c => c.toString());
  const rawIdx: Record<string, number> = {};
  ["Project Name", "Activity ID", ...headersToCopy].forEach(hdr => 
  {
    rawIdx[hdr] = rawHeaders.indexOf(hdr);
  });

  // prepare set of keys seen in this run
  const seenKeys = new Set<string>();

  // process each row in RAW Commitments
  for (let i = 1; i < rawData.length; i++) 
  {
    const row = rawData[i];
    const rawProj = (row[rawIdx["Project Name"]] ?? "").toString();
    const actId = (row[rawIdx["Activity ID"]] ?? "").toString().trim();
    if (!actId) continue; // skip rows lacking an Activity ID

    // split rawProj at first ":" to extract project number + name
    const parts = rawProj.split(":");
    const projectNumber = parts[0].trim();
    const projectName = parts.slice(1).join(":").trim();
    const newKey = projectNumber + "-" + actId;

    // record seen key
    seenKeys.add(newKey);

    // decide row: existing or new
    const existingRow = keyRowMap[newKey];
    const outRow = existingRow ?? nextRow;

    // write combined key
    target
      .getCell(outRow, firstCol + tgtIdx[combinedHeader])
      .setValue(newKey);

    // write project name remainder
    target
      .getCell(outRow, firstCol + tgtIdx[projectNameHeader])
      .setValue(projectName);

    // write other columns
    headersToCopy.forEach(hdr => 
    {
      const tgtCol = firstCol + tgtIdx[hdr];
      const rawCol = rawIdx[hdr];
      const value = rawCol >= 0 ? row[rawCol] : "";
      target.getCell(outRow, tgtCol).setValue(value);
    });

    // mark "In Raw?" = Yes
    target
      .getCell(outRow, firstCol + tgtIdx[inRawHeader])
      .setValue("Yes");

    // append borders and record if new
    if (!existingRow) 
    {
      keyRowMap[newKey] = nextRow;

      const brdRange = target.getRangeByIndexes(outRow, firstCol, 1, colCount);
      const brdFormat = brdRange.getFormat();
      brdFormat.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
      brdFormat.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
      brdFormat.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
      brdFormat.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
      brdFormat.getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.continuous);

      nextRow++;
    }
  }

  // mark rows no longer in raw as "No"
  for (let r = firstRow + 1; r < nextRow; r++) 
  {
    const keyCell = target.getCell(r, firstCol + tgtIdx[combinedHeader]).getValue() as string;
    const inRaw = seenKeys.has(keyCell) ? "Yes" : "No";
    target.getCell(r, firstCol + tgtIdx[inRawHeader]).setValue(inRaw);
  }

  console.log(`Done: updated "In Raw?" status and synced data.`);
} // end of program
