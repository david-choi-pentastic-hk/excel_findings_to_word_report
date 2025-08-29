/**
 * @file script.ts
 * @brief An office script to rearrange the cells for copying to word report.
 * @author David Choi <david.choi@pentastic.hk>
 * @verison 2025.08.08
 * @copyright Copyright (c) 2025 Pentastic Security Limited. All rights reserved.
 */

function main(workbook: ExcelScript.Workbook) {
    // Get the active cell and worksheet.
    let selectedCell = workbook.getActiveCell();
    let selectedSheet = workbook.getActiveWorksheet();

    const summaryWorksheet = workbook.getFirstWorksheet();
    const usedRange = summaryWorksheet.getUsedRange();
    const firstRow = usedRange.getRowIndex() + 1; // inclusive, but skip the first row bc it is the column description
    const lastRow = firstRow + usedRange.getRowCount() - 1; // inclusive
    for (let i = firstRow; i <= lastRow; i++) {
      copyData(summaryWorksheet, workbook, i);
    }
}

function getOrCreateSheet(destWorkbook: ExcelScript.Workbook, sheetName: string): ExcelScript.Worksheet {
  const worksheets: ExcelScript.Worksheet[] =
    destWorkbook.getWorksheets();

  if (!worksheets.map(worksheet => worksheet.getName()).includes(sheetName)) {
    destWorkbook.addWorksheet(sheetName);
  }

  return destWorkbook.getWorksheet(sheetName);
}

function copyData(srcSheet: ExcelScript.Worksheet, destWorkbook: ExcelScript.Workbook, entryRow: number) {
  const findingNumber: string = srcSheet.getCell(entryRow, 0).getValue().toString();
  if (findingNumber == null || findingNumber.trim() === "") {
    return;
  }
  const destSheet: ExcelScript.Worksheet = getOrCreateSheet(destWorkbook, findingNumber);

  // copy W1
  destSheet.getCell(0, 0).setValue(
    findingNumber
  );

  const jointFinding: string =
    srcSheet.getCell(entryRow, 1).getValue().toString();
  
  // copy finding
  destSheet.getCell(0, 1).setValue(
    jointFinding.substring(0, jointFinding.indexOf("\n"))
  );

  // copy risk description
  copySingleColumnToRow(srcSheet, destSheet, entryRow, "Risk Description", "Risk Description", 1);

  // copy owasp
  destSheet.getCell(2, 0).setValue("OWASP Top 10 Vulnerability");

  destSheet.getCell(2, 1).setValue(
      jointFinding.substring(jointFinding.lastIndexOf("\n") + 1)
  );

  // copy risk level
  copySingleColumnToRow(srcSheet, destSheet, entryRow, "Risk Level", "Risk Level", 3);

  // copy likelihood
  copySingleColumnToRow(srcSheet, destSheet, entryRow, "Likelihood", "Likelihood", 4);

  // copy affected asset
  copySingleColumnToRow(srcSheet, destSheet, entryRow, "Affected URL / API", "Affected Asset", 5);

  // copy impact
  copySingleColumnToRow(srcSheet, destSheet, entryRow, "Impact", "Impact", 6);

  // copy evidence for finding
  copySingleColumnToRow(srcSheet, destSheet, entryRow, "Impact", "Evidence for the finding", 7);

  // copy evidence
  copySingleColumnToRow(srcSheet, destSheet, entryRow, "Recommended Safeguards", "Recommended Safeguards:", 8);

  // copy evidence for remedial actions
  copySingleColumnToRow(srcSheet, destSheet, entryRow, "Evidence for the remedial actions", "Evidence for the remedial actions", 9);

  // copy rectification status
  copySingleColumnToRow(srcSheet, destSheet, entryRow, "Rectification Status as at ", "Rectification Status as at ", 10);
}

function copySingleColumnToRow(srcSheet: ExcelScript.Worksheet, destSheet: ExcelScript.Worksheet, entryRow: number, srcFieldName: string, destFieldName: string, destFieldRow: number)
{
  destSheet.getCell(destFieldRow, 0).setValue(destFieldName);

  const likelihoodColumnIndex: number = findColumnByName(srcSheet, srcFieldName);
  const value: string | number | boolean = (likelihoodColumnIndex != -1) ?
    srcSheet.getCell(entryRow, likelihoodColumnIndex).getValue() :
    "";
  destSheet.getCell(destFieldRow, 1).setValue(value);
}

function findColumnByName(worksheet: ExcelScript.Worksheet, columnName: string): number {
  columnName = columnName.trim();
  const usedRange = worksheet.getUsedRange();
  let columnIndex = usedRange.getColumnIndex();
  const columnCount = usedRange.getColumnCount();
  for (let i: number = 0; i < columnCount; i++, columnIndex++) {
    const thisColumnName = worksheet.getCell(0, columnIndex).getValue().toString();
    if (columnName === thisColumnName.trim()) {
      return columnIndex;
    }
  }
  return -1;
}