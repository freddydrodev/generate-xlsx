import type { Font, Borders, Alignment } from "exceljs";

export const DEFAULT_BORDER: Partial<Borders> = {
  bottom: { color: { argb: "FF000000" }, style: "thin" },
  top: { color: { argb: "FF000000" }, style: "thin" },
  left: { color: { argb: "FF000000" }, style: "thin" },
  right: { color: { argb: "FF000000" }, style: "thin" },
};

export const DEFAULT_ROW_ALIGNEMENT: Partial<Alignment> = {
  horizontal: "center",
  vertical: "middle",
  wrapText: true,
};

// const DEFAULT_XLSX_FONT = "Goudy Old Style";

export const ROW_HEIGHT = 50;

export const DEFAULT_ROW_FONT: Partial<Font> = {
  size: 16,
  // name: DEFAULT_XLSX_FONT,
};

// export const BOLD_XLSX_FONT = "Goudy Old Style Bold";

export const BOLD_FONT: Partial<Font> = {
  ...DEFAULT_ROW_FONT,
  bold: true,
  // name: BOLD_XLSX_FONT,
};

export const DEFAULT_NUM_FMT = "# ## [$F CFA-fr-CI];-# ## [$F CFA-fr-CI]";
