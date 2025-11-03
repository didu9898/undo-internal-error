/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
});

export async function run() {
  try {
    OfficeExtension.config.extendedErrorLogging = true;
    await Excel.run({ mergeUndoGroup: true }, async (context) => {
      const newSheetName = "NewSheet";
      let newSheet = context.workbook.worksheets.getItemOrNullObject(newSheetName);
      const currentCell = context.workbook.getActiveCell();
      currentCell.load("address");
      await context.sync();
      let newSheetCreated = false;
      let nextFreeCell;
      if (newSheet.isNullObject) {
        context.workbook.worksheets.add(newSheetName);
        await context.sync();
        console.log("New sheet added");
        newSheetCreated = true;
        newSheet = context.workbook.worksheets.getItem(newSheetName);
        nextFreeCell = newSheet.getRange("A1");
      } else {
        newSheet = context.workbook.worksheets.getItem(newSheetName);
        const column = newSheet.getRange("A:A");
        const usedRange = column.getUsedRange();
        usedRange.load("rowCount, rowIndex");
        await context.sync();
        const nextRow = usedRange.rowIndex + usedRange.rowCount;
        nextFreeCell = newSheet.getCell(nextRow, 0);
      }
      nextFreeCell.load("address");
      nextFreeCell.values = [[currentCell.address]];
      await context.sync();
      const formula = `=CONTOSO.FILLANDAPPLYSTYLE('${nextFreeCell.address.split("!")[0]}'!${nextFreeCell.address.split("!")[1]})`;
      // const formula = "=" + 'CONTOSO.FILLANDAPPLYSTYLE("A1:B2")'; // Alternative formula without sheet name = no error
      currentCell.formulas = [[formula]];
      await context.sync();
      console.log("Custom function formula set in cell");
    });
  } catch (error) {
    console.error(error);
  }
}
