// Shaan Banday – TypeScript  
// Updates as of July 8th, 2025

function main(workbook: ExcelScript.Workbook) // define main entry point
{
    type TCell = string | number | boolean | Date; // possible cell value types

    // the code section below, defining constants, COULD probably be optimized more, but it works so I'll keep for now
    const targetName = "ERs"; // sheet to updateA
    const rawName = "RAW ERs"; // sheet to read fresh data from
    const keyHeader = "ER"; // unique identifier column
    const inRawHeader = "In Raw?"; // flag column, must exist or will be added
    const headersToCopy: string[] = ["ER", "Facility", "ER Title", "Min T-Week", "Set to Ready TCD", "Due Date", "Workflow Status", "SM Due Date", "All Active Outages", "Due Date Type", "Earliest Outage", "All Active WO Types", "OE", "SM(s)", "Earliest WO Age", "Highest WO Priority", "Project"]; // rest of the column headers

    const target = workbook.getWorksheet(targetName); // get the ERs sheet
    if (!target) // if the script cannot find the CRs sheet
    {
        throw new Error(`Sheet "${targetName}" not found.`); // catch the error and show in console (debugging step)
    }

    const usedTarget = target.getUsedRange(); // find used range (headers + data) in CRs
    if (!usedTarget) // if the script sees CRs sheet as fully empty
    {
        throw new Error(`No data in "${targetName}".`); // catch the error and show in console (debugging step)
    }

    // Record where that range begins and its size
    const firstRow = usedTarget.getRowIndex(); // row index of header
    const firstCol = usedTarget.getColumnIndex(); // column index of first header
    let colCount = usedTarget.getColumnCount(); // number of columns
    const totalRows = usedTarget.getRowCount(); // total rows including header

    // read header row into array
    const headerVals = target.getRangeByIndexes(firstRow, firstCol, 1, colCount).getValues() as TCell[][];
    const headerRow: string[] = headerVals[0].map(cell => (cell ?? "").toString().trim());

    const tgtIdx: Record<string, number> = {}; // map header to column index in CRs
    [keyHeader, ...headersToCopy, inRawHeader].forEach(hdr => {
        const idx = headerRow.indexOf(hdr);
        if (idx < 0) {
            throw new Error(`Header "${hdr}" missing in "${targetName}".`);
        }
        tgtIdx[hdr] = idx;
    });
    const keyCol = tgtIdx[keyHeader]; // we will use this to read/write the key

    const keyRowMap: Record<string, number> = {}; // build a lookup of existing key to row index in ERs
    if (totalRows > 1) {
        const keyVals = target.getRangeByIndexes(firstRow + 1, firstCol + keyCol, totalRows - 1, 1).getValues() as TCell[][];
        keyVals.forEach((r, i) => {
            const k = (r[0] ?? "").toString().trim();
            if (k) {
                keyRowMap[k] = firstRow + 1 + i;
            }
        });
    }
    let nextRow = firstRow + totalRows; // pointer for appending new rows

    const raw = workbook.getWorksheet(rawName); // Access the RAW ERs sheet
    if (!raw) // if the raw sheet doesn't exist
    {
        throw new Error(`Sheet "${rawName}" not found.`); // catch the error and show in console (debugging step)
    }

    const usedRaw = raw.getUsedRange();

    if (!usedRaw) return; // nothing to do if raw empty
    const rawData = usedRaw.getValues() as TCell[][];
    if (rawData.length < 2) return; // skip if no data rows

    // Read and trim the RAW header row
    const rawHeaders: string[] = rawData[0].map(cell => (cell ?? "").toString().trim());

    const rawIdx: Record<string, number> = {}; // Map each needed header in RAW to its column index
    headersToCopy.forEach(hdr => {
        rawIdx[hdr] = rawHeaders.indexOf(hdr);
    });

    const seenKeys = new Set<string>(); // Prepare to track which ER keys appear this run

    for (let i = 1; i < rawData.length; i++) // Loop through each data row in RAW ERs
    {
        const row = rawData[i] as TCell[];

        const rawKey = (row[rawIdx[keyHeader]] ?? "").toString().trim(); // Extract and trim the ER key. skip if blank
        if (!rawKey) {
            continue; // do not append blank-key rows
        }
        seenKeys.add(rawKey); // record that we saw this ER

        const existingRow = keyRowMap[rawKey]; // Decide whether to overwrite an existing row or append new
        const outRow = existingRow !== undefined ? existingRow : nextRow;

        headersToCopy.forEach(hdr => // Copy each configured column from raw to CRs
        {
            const tc = firstCol + tgtIdx[hdr]; // target column
            const rc = rawIdx[hdr]; // raw column
            const val = rc >= 0 ? row[rc] : ""; // fallback empty if missing
            target.getCell(outRow, tc).setValue(val);
        });

        target.getCell(outRow, firstCol + tgtIdx[inRawHeader]).setValue("Yes"); // Mark “In Raw?” as Yes 

        if (existingRow === undefined) // If newly appended, add borders and record its index
        {
            keyRowMap[rawKey] = nextRow;
            const brRange = target.getRangeByIndexes(outRow, firstCol, 1, colCount);
            const brFmt = brRange.getFormat();
            [ExcelScript.BorderIndex.edgeTop, ExcelScript.BorderIndex.edgeBottom, ExcelScript.BorderIndex.edgeLeft, ExcelScript.BorderIndex.edgeRight, ExcelScript.BorderIndex.insideVertical].forEach(bi => brFmt.getRangeBorder(bi).setStyle(ExcelScript.BorderLineStyle.continuous));
            nextRow++;
        }
    }

    for (let r = firstRow + 1; r < nextRow; r++) // Mark any ERs not seen in raw as “No”
    {
        const cellVal = target.getCell(r, firstCol + keyCol).getValue();
        const erKey = (cellVal ?? "").toString().trim();
        const flag = seenKeys.has(erKey) ? "Yes" : "No";
        target.getCell(r, firstCol + tgtIdx[inRawHeader]).setValue(flag);
    }
    console.log(`Done: in-place updates + new additions to "${targetName}".`); // Final confirmation log
} // end of program
