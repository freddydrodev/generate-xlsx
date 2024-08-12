import type { Alignment, Font } from "exceljs";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import {
  BOLD_FONT,
  DEFAULT_BORDER,
  DEFAULT_NUM_FMT,
  DEFAULT_ROW_ALIGNEMENT,
  DEFAULT_ROW_FONT,
  ROW_HEIGHT,
} from "./constants";

type GenerateXLSXArgs = {
  fileName: string;
  config: {
    name: string;
    orientation?: "portrait" | "landscape";
    zoom?: number;
    colWidth?: number;
    colHeight?: number;
  };
  data: { [key: string]: string | number }[];
  // headers: Partial<ExcelJS.Column>[];
  headers: ({
    isCurrency?: boolean;
    isNumber?: boolean;
    hasTotal?: boolean;
  } & Pick<
    Partial<ExcelJS.Column>,
    | "key"
    | "header"
    | "border"
    | "alignment"
    | "fill"
    | "numFmt"
    | "values"
    | "width"
    | "style"
    | "font"
  >)[];
  rowAlignment?: Partial<Alignment>;
  defaultFont?: Partial<Font>;
  boldFont?: Partial<Font>;
  height?: number;
  title?: { height?: number; fontSize?: number };
};

export const generateXLSXGrid = async (args: GenerateXLSXArgs) => {
  const { data, config, rowAlignment, defaultFont, height, fileName, headers } =
    args;

  const wb = new ExcelJS.Workbook();

  /**
   * CREATE THE SHEET
   */
  const sheet = wb.addWorksheet(config.name, {
    views: [{ style: "pageBreakPreview" }],
    properties: {
      defaultRowHeight: config.colHeight ?? ROW_HEIGHT,
    },
    headerFooter: {
      oddFooter: "&F&RPage &P / &N",
    },
    pageSetup: {
      paperSize: 9,
      horizontalCentered: true,
      scale: config.zoom ?? 100,
      orientation: config.orientation ?? "portrait",
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

  /**
   * CREATE THE SECTION HEADER
   */
  sheet.columns = headers;

  sheet.eachRow((row) => {
    row.height = height ?? ROW_HEIGHT;
  });

  sheet.columns.forEach((column, index) => {
    column.eachCell?.({ includeEmpty: false }, (cell) => {
      const header = headers.at(index);

      cell.border = header?.border ?? DEFAULT_BORDER;

      cell.font = header?.font ?? BOLD_FONT;

      column.alignment = header?.alignment ?? DEFAULT_ROW_ALIGNEMENT;

      cell.fill = {
        pattern: "solid",
        type: "pattern",
        fgColor: { argb: "FFD9D9D9" },
      };

      if (header?.isCurrency || header?.isNumber) {
        column.numFmt =
          header?.numFmt ?? header.isCurrency ? DEFAULT_NUM_FMT : "#,##;-#,##";

        console.log(column.numFmt);
      }
    });
  });

  /**
   * ADD THE DATA ROWS
   */

  const dataRow = sheet.addRows(data);

  if (dataRow.length > 0) {
    dataRow.forEach((r, i) => {
      r.height = ROW_HEIGHT;

      r.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const header = headers.at(colNumber - 1);

        cell.border = header?.border ?? DEFAULT_BORDER;

        cell.alignment = header?.alignment ?? DEFAULT_ROW_ALIGNEMENT;

        cell.font = header?.font ?? DEFAULT_ROW_FONT;
      });
    });
  }

  /**
   * ADD TOTAL
   */

  const totauxFields: any = {};

  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];

    if (h.hasTotal) {
      totauxFields[h.key ?? "-"] = 0;
    } else {
      totauxFields[h.key ?? "-"] = null;
    }
  }

  const total = sheet.addRow(totauxFields);

  total.height = ROW_HEIGHT;

  total.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const header = headers[colNumber - 1];
    cell.border = header?.border ?? DEFAULT_BORDER;
    cell.alignment = header?.alignment ?? DEFAULT_ROW_ALIGNEMENT;
    cell.font = header?.font ?? BOLD_FONT;

    if (header?.hasTotal) {
      console.log(header);
      const col = cell.address.replace(/[0-9]+/gi, "");

      cell.value = {
        date1904: false,
        formula: `SUM(${col + "1"}:${col + ((cell.row as any) - 1)})`,
      };
    }
  });

  if (data.length <= 0 || headers.length <= 0) return;

  const buffer = await wb.xlsx.writeBuffer();

  const blob = new Blob([buffer], { type: "applicationi/xlsx" });

  saveAs(blob, fileName.replace(/\.xlsx/gi, "") + ".xlsx");
};
