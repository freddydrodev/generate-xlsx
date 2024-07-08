import type { Workbook } from "exceljs";

import {
  BOLD_FONT,
  DEFAULT_BORDER,
  DEFAULT_ROW_ALIGNEMENT,
  DEFAULT_ROW_FONT,
  XLS_JUSTIF_HEADING_HEIGHT,
  ROW_HEIGHT,
} from "./constants";
import { generateXLSXGrid } from ".";
import { generalPageHeading } from "./generalPageHeading";

export const sheetGenerator = (
  workbook: Workbook,
  sheetName: string,
  headers: string[],
  columns: { key: string; width: number }[],
  title: string,
  data: any[],
  totalCellAddress: number,
  zoom?: number
) => {
  console.log(totalCellAddress);

  const sheet = workbook.addWorksheet(sheetName, {
    views: [{ style: "pageBreakPreview" }],
    properties: {
      defaultRowHeight: ROW_HEIGHT,
    },
    headerFooter: {
      oddFooter: "&C&A_&F&RPage &P / &N",
    },
    pageSetup: {
      paperSize: 9,
      horizontalCentered: true,
      scale: zoom ?? 60,
      orientation: "portrait",
      margins: {
        top: 0.75,
        bottom: 0.75,
        left: 0.25,
        right: 0.25,
        header: 0.3,
        footer: 0.3,
      },
    },
  });

  sheet.columns = columns;

  generateXLSXGrid(
    sheet,
    headers.length,
    generalPageHeading(headers.length, title),
    DEFAULT_ROW_ALIGNEMENT,
    DEFAULT_ROW_FONT,
    BOLD_FONT,
    30,
    { height: XLS_JUSTIF_HEADING_HEIGHT, fontSize: 28 }
  );

  /**
   * CREATE THE SECTION HEADER
   */
  const headerRow = sheet.addRow(headers);

  headerRow.height = ROW_HEIGHT;
  headerRow.eachCell({ includeEmpty: true }, (cell) => {
    cell.border = DEFAULT_BORDER;
    cell.font = BOLD_FONT;
    cell.alignment = DEFAULT_ROW_ALIGNEMENT;
    cell.style.fill = {
      pattern: "solid",
      type: "pattern",
      fgColor: { argb: "FFD9D9D9" },
    };
  });

  /**
   * CREATE THE SECTION DATA
   */
  const dataRow = sheet.addRows(data);

  if (dataRow.length > 0) {
    const totauxRows: string[] = [];

    dataRow.forEach((r) => {
      r.height = ROW_HEIGHT;

      r.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.border = DEFAULT_BORDER;
        cell.alignment = DEFAULT_ROW_ALIGNEMENT;
        cell.font = DEFAULT_ROW_FONT;

        const isTotal = colNumber === totalCellAddress;

        if (isTotal) {
          cell.numFmt = "#,##;-#,##";
          cell.font = { ...BOLD_FONT, size: 16 };

          totauxRows.push(cell.address);
        }
      });
    });

    /**
     * CREATE THE TOTAL SECTION
     */
    const totalRow = sheet.addRow([
      "TOTAL",
      ...Array(headers.length - 1).fill(""),
    ]);

    //TODO SOLVE TOTAL ISSUE
    if (totalCellAddress > 1) {
      sheet.mergeCells(
        `${totalRow.getCell(1).address}:${
          totalRow.getCell(totalCellAddress - 1).address
        }`
      );
    }

    totalRow.height = ROW_HEIGHT;

    totalRow.eachCell((cell, colNumber) => {
      cell.border = DEFAULT_BORDER;
      cell.alignment = DEFAULT_ROW_ALIGNEMENT;
      cell.font = DEFAULT_ROW_FONT;
      cell.font = { ...BOLD_FONT, size: 16 };

      const addressLetter = cell.address.replace(/[0-9]/g, "");
      const firstRow = addressLetter + dataRow[0].number;
      const lastRow = addressLetter + dataRow[dataRow.length - 1].number;

      if (colNumber === totalCellAddress) {
        cell.numFmt = "#,##;-#,##";

        cell.value = {
          date1904: false,
          formula: `SUM(${firstRow}:${lastRow})`,
        };
      }
    });
  }
};
