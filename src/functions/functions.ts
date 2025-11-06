/* global clearInterval, console, CustomFunctions, setInterval */

import { addCustomStyles } from "./StyleDefinition";

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
 * @param applyStyles Weather to apply styles or not
 * @returns A 3x2 dynamic array.
 */
export async function fillAndApplyStyle(applyStyles?: boolean): Promise<any[][]> {
  setTimeout(() => {
    apply(applyStyles);
  }, 25);
  return [
    [1, 2],
    [1, 2],
    [1, 2],
  ];
}

async function apply(applyStyles?: boolean): Promise<void> {
  try {
    OfficeExtension.config.extendedErrorLogging = true;
    await Excel.run(async (context) => {
      const excelStyles = context.workbook.styles;
      try {
        context.application.suspendScreenUpdatingUntilNextSync();
        excelStyles.load("items/name");
        await context.sync();
      } catch (error) {
        console.error("error: " + error);
      }
      try {
        context.application.suspendScreenUpdatingUntilNextSync();
        if (applyStyles) {
          await fillStylesIfMissing(context, excelStyles);
          await context.sync();
        }
      } catch (error) {
        console.error("error: " + error);
      }
      try {
        context.application.suspendScreenUpdatingUntilNextSync();
        const activeCell = context.workbook.getActiveCell();
        if (applyStyles) {
          const fullRange = activeCell.getResizedRange(2, 1);
          fullRange.style = Excel.BuiltInStyle.normal;
          const r1 = activeCell.getBoundingRect(activeCell.getOffsetRange(1, 1));
          r1.style = "customStyleColsHeaderRow";
          const r2 = r1.getOffsetRange(1, 0);
          r2.style = "customStyleColsHier1Attribute1Row";
        }
      } catch (error) {
        console.error("error: " + error);
      }
      try {
        await context.sync();
      } catch (error) {
        console.error("error: " + error);
      }
      console.log("Applied style to range A1:B2");
    });
  } catch (error) {
    console.error("error: " + error);
  }
}

const stylePrefix = "customStyle";
const stylePrefixLowerCase = stylePrefix.toLowerCase();

function containsCustomStyle(excelStyles: Excel.StyleCollection) {
  if (excelStyles) {
    return excelStyles.items.find((style) => {
      return style.name.toLowerCase().startsWith(stylePrefixLowerCase);
    });
  }
  return false;
}

async function fillStylesIfMissing(
  context: Excel.RequestContext,
  excelStyles: Excel.StyleCollection
) {
  if (!containsCustomStyle(excelStyles)) { // Check if styles are already added
    addCustomStyles(context);
    await context.sync();
  }
}
