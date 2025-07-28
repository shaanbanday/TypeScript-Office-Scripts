// Shaan Banday – TypeScript  
// Updates as of June 16th, 2025

function main(workbook: ExcelScript.Workbook) // define main entry point
{
    type TCell = string | number | boolean | Date; // possible cell value types

    // the code section below, defining constants, COULD probably be optimized more, but it works so I'll keep for now
    const targetName = "EC Closeouts"; // sheet to update
    const rawName = "RAW EC Closeouts"; // sheet with source data
    const keyHeader = "EC Number"; // unique identifier column
    const inRawHeader = "In Raw?"; // flag for presence in raw
    const headersToCopy: string[] = ["Days Due", "Origin Due Date", "EC Number", "Description", "Closeout Status", "Job Plan", "Project", "Engineer", "Task 749 Status", "Task 750 Documents", "EC Age", "Outstanding CRs"];

    //Access the EC Closeouts sheet
    const target = workbook.getWorksheet(targetName);
    if (!target) // if the script cannot find the EC Closeouts sheet
    {
        throw new Error(`Sheet "${targetName}" not found.`); // catch the error and show in console (debugging step)
    }

    // find used range (headers + data) in EC Closeouts
    const usedTarget = target.getUsedRange();
    if (!usedTarget) // if the script sees EC Closeouts sheet as fully empty
    {
        throw new Error(`No data in "${targetName}".`); // catch the error and show in console (debugging step)
    }

    const firstRow = usedTarget.getRowIndex();    // zero-based index of the header row
    const firstCol = usedTarget.getColumnIndex(); // zero-based index of the first column
    let colCount = usedTarget.getColumnCount(); // how many columns are in that range
    const totalRows = usedTarget.getRowCount();    // total rows including header

    // read header row into array
    const headerVals = target.getRangeByIndexes(firstRow, firstCol, 1, colCount).getValues() as TCell[][];
    const headerRow: string[] = headerVals[0].map(cell => (cell ?? "").toString().trim());

    // Match each header in EC Closeouts to its column index
    const tgtIdx: Record<string, number> = {};
    [keyHeader, ...headersToCopy, inRawHeader].forEach(hdr => {
        const idx = headerRow.indexOf(hdr);
        if (idx < 0) {
            throw new Error(`Header "${hdr}" missing in "${targetName}".`);
        }
        tgtIdx[hdr] = idx; // store mapping: header name → column index
    });
    const keyCol = tgtIdx[keyHeader]; // column index of the EC Number

    //Build a lookup of existing EC Number → row index in EC Closeouts
    const keyRowMap: Record<string, number> = {};
    if (totalRows > 1) {
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
    // Pointer to the next empty row for appending new entries
    let nextRow = firstRow + totalRows;

    //Access the RAW EC Closeouts sheet
    const raw = workbook.getWorksheet(rawName);
    if (!raw) // if the script cannot find the raw EC Closeouts sheet
    {
        throw new Error(`Sheet "${rawName}" not found.`);
    }
    const usedRaw = raw.getUsedRange();
    if (!usedRaw) return;               // nothing to do if raw is empty
    const rawData = usedRaw.getValues() as TCell[][];
    if (rawData.length < 2) return;     // skip if no data rows

    // Read and trim the RAW header row
    const rawHeaders: string[] = rawData[0]
        .map(cell => (cell ?? "").toString().trim());

    // Map each needed raw header to its column index
    const rawIdx: Record<string, number> = {};
    ["Days Due", "Origin Due Date", ...headersToCopy.slice(2)].forEach(hdr => {
        rawIdx[hdr] = rawHeaders.indexOf(hdr);
    });

    // Prepare a set to track which EC Numbers we see this run
    const seenKeys = new Set<string>();

    //Loop through every data row in RAW EC Closeouts
    for (let i = 1; i < rawData.length; i++) 
    {
        const row = rawData[i] as TCell[];
        // Extract and trim the EC Number key
        const key = (row[rawIdx[keyHeader]] ?? "").toString().trim();
        if (!key) continue; // skip rows missing EC Number

        // Record that we've seen this key
        seenKeys.add(key);

        // Determine whether to overwrite or append
        const existingRow = keyRowMap[key];
        const outRow = existingRow ?? nextRow;

        // Copy "Days Due" from raw → EC Closeouts
        target
            .getCell(outRow, firstCol + tgtIdx["Days Due"])
            .setValue(row[rawIdx["Days Due"]] ?? "");

        // Copy "Origin Due Date" from raw → EC Closeouts
        target
            .getCell(outRow, firstCol + tgtIdx["Origin Due Date"])
            .setValue(row[rawIdx["Origin Due Date"]] ?? "");

        // Copy the rest of the columns (excluding the first two)
        headersToCopy.slice(2).forEach(hdr => {
            const tgtC = firstCol + tgtIdx[hdr];
            const rawC = rawIdx[hdr];
            const val = rawC >= 0 ? row[rawC] : "";
            target.getCell(outRow, tgtC).setValue(val);
        });

        // Write EC Number into its column (keyHeader)
        target
            .getCell(outRow, firstCol + tgtIdx[keyHeader])
            .setValue(key);

        // Mark "In Raw?" = Yes
        target
            .getCell(outRow, firstCol + tgtIdx[inRawHeader])
            .setValue("Yes");

        // If this was appended (new row), add borders and record its index
        if (!existingRow) {
            keyRowMap[key] = nextRow;

            const brRange = target.getRangeByIndexes(outRow, firstCol, 1, colCount);
            const brdFmt = brRange.getFormat();
            [
                ExcelScript.BorderIndex.edgeTop,
                ExcelScript.BorderIndex.edgeBottom,
                ExcelScript.BorderIndex.edgeLeft,
                ExcelScript.BorderIndex.edgeRight,
                ExcelScript.BorderIndex.insideVertical
            ].forEach(bi =>
                brdFmt.getRangeBorder(bi).setStyle(ExcelScript.BorderLineStyle.continuous)
            );

            nextRow++; // advance append pointer
        }
    }
    // mark any rows no longer in raw as "No"
    for (let r = firstRow + 1; r < nextRow; r++) {
        const cellKey = target
            .getCell(r, firstCol + tgtIdx[keyHeader])
            .getValue();
        const keyStr = (cellKey ?? "").toString().trim();
        const flag = seenKeys.has(keyStr) ? "Yes" : "No";
        target
            .getCell(r, firstCol + tgtIdx[inRawHeader])
            .setValue(flag);
    }
    console.log(`Done: in-place updates + new additions to "${targetName}".`);
} // end of program
