import type { Alignment, Font } from "exceljs";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import {
  DEFAULT_BORDER,
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
  headers: Pick<
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
  >[];
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
      oddFooter: "&C&A_&F&RPage &P / &N",
    },
    pageSetup: {
      paperSize: 9,
      horizontalCentered: true,
      scale: config.zoom ?? 60,
      orientation: config.orientation ?? "portrait",
      margins: {
        top: 0,
        bottom: 0,
        left: 0,
        right: 0,
        header: 0,
        footer: 0,
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
    column.eachCell?.({ includeEmpty: true }, (cell) => {
      cell.border = DEFAULT_BORDER;
      cell.font = defaultFont ?? DEFAULT_ROW_FONT;
      cell.alignment = rowAlignment ?? DEFAULT_ROW_ALIGNEMENT;
      cell.fill = {
        pattern: "solid",
        type: "pattern",
        fgColor: { argb: "FFD9D9D9" },
      };
    });
  });

  /**
   * ADD THE DATA ROWS
   */

  const dataRow = sheet.addRows(data);

  if (dataRow.length > 0) {
    dataRow.forEach((r, i) => {
      r.height = ROW_HEIGHT;

      r.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.border = DEFAULT_BORDER;
        cell.alignment = DEFAULT_ROW_ALIGNEMENT;
        cell.font = DEFAULT_ROW_FONT;
      });
    });
  }

  if (data.length <= 0 || headers.length <= 0) return;

  const buffer = await wb.xlsx.writeBuffer();

  const blob = new Blob([buffer], { type: "applicationi/xlsx" });

  saveAs(blob, fileName);
};
