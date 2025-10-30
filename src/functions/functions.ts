/* global clearInterval, console, CustomFunctions, setInterval */
/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Fill and apply style
 * @customfunction fillAndApplyStyle
 * @param address The address of the range to fill and apply style to
 * @returns A 3x2 dynamic array.
 */
export async function fillAndApplyStyle(
  address: string
): Promise<any[][]> {
  await apply(address);
  return [
    [1, 2],
    [1, 2],
    [1, 2],
  ];
}

async function apply(address: string): Promise<void> {
  try {
    OfficeExtension.config.extendedErrorLogging = true;
    await Excel.run({ mergeUndoGroup: true }, async (context) => { // remove mergeUndoGroup option = no error
      context.application.suspendScreenUpdatingUntilNextSync();
      // const excelStyles = context.workbook.styles;
      // excelStyles.load("items/name");
      // await context.sync();
      // const range = sheet.getRange(address);
      // range.format.fill.color = "yellow";
      // range.format.font.bold = true;
      // range.numberFormat = [
      //   ["0.00", "0.00"],
      //   ["0.00", "0.00"],
      // ];
      await context.sync();
      console.log("Applied style to range A1:B2");
    });
  } catch (error) {
    console.error("error: " + error);
  }
}

