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
    await Excel.run(async (context) => {
      const newSheetName = "NewSheet";
      let newSheet = context.workbook.worksheets.getItemOrNullObject(newSheetName);
      await context.sync();
      if (newSheet.isNullObject) {
        context.workbook.worksheets.add(newSheetName);
        await context.sync();
        console.log("New sheet added");
      }
      newSheet = context.workbook.worksheets.getItem(newSheetName);
      const nextFreeCell = newSheet.getCell(0, 0);
      nextFreeCell.load("address");
      await context.sync();
      nextFreeCell.values = [["A1:B2"]];
      await context.sync();
      const formula = "=" + "CONTOSO.FILLANDAPPLYSTYLE" + "('" + newSheetName + "'!A1)";
      // const formula = "=" + 'CONTOSO.FILLANDAPPLYSTYLE("A1:B2")'; // Alternative formula without sheet name = no error
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const targetRange = sheet.getRange("A1");
      targetRange.formulas = [[formula]];
      await context.sync();
      console.log("Custom function formula set in cell");
    });
  } catch (error) {
    console.error(error);
  }
}
