var __defProp = Object.defineProperty;
var __defProps = Object.defineProperties;
var __getOwnPropDescs = Object.getOwnPropertyDescriptors;
var __getOwnPropSymbols = Object.getOwnPropertySymbols;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __propIsEnum = Object.prototype.propertyIsEnumerable;
var __defNormalProp = (obj, key, value) => key in obj ? __defProp(obj, key, { enumerable: true, configurable: true, writable: true, value }) : obj[key] = value;
var __spreadValues = (a, b) => {
  for (var prop in b || (b = {}))
    if (__hasOwnProp.call(b, prop))
      __defNormalProp(a, prop, b[prop]);
  if (__getOwnPropSymbols)
    for (var prop of __getOwnPropSymbols(b)) {
      if (__propIsEnum.call(b, prop))
        __defNormalProp(a, prop, b[prop]);
    }
  return a;
};
var __spreadProps = (a, b) => __defProps(a, __getOwnPropDescs(b));
var __async = (__this, __arguments, generator) => {
  return new Promise((resolve, reject) => {
    var fulfilled = (value) => {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    };
    var rejected = (value) => {
      try {
        step(generator.throw(value));
      } catch (e) {
        reject(e);
      }
    };
    var step = (x) => x.done ? resolve(x.value) : Promise.resolve(x.value).then(fulfilled, rejected);
    step((generator = generator.apply(__this, __arguments)).next());
  });
};

// src/index.ts
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

// src/constants.ts
var DEFAULT_BORDER = {
  bottom: { color: { argb: "FF000000" }, style: "thin" },
  top: { color: { argb: "FF000000" }, style: "thin" },
  left: { color: { argb: "FF000000" }, style: "thin" },
  right: { color: { argb: "FF000000" }, style: "thin" }
};
var DEFAULT_ROW_ALIGNEMENT = {
  horizontal: "center",
  vertical: "middle",
  wrapText: true
};
var DEFAULT_XLSX_FONT = "Goudy Old Style";
var ROW_HEIGHT = 50;
var DEFAULT_ROW_FONT = {
  size: 16,
  name: DEFAULT_XLSX_FONT
};
var BOLD_XLSX_FONT = "Goudy Old Style Bold";
var BOLD_FONT = __spreadProps(__spreadValues({}, DEFAULT_ROW_FONT), {
  bold: true,
  name: BOLD_XLSX_FONT
});
var DEFAULT_NUM_FMT = "# ##0 [$F CFA-fr-CI]";

// src/index.ts
var generateXLSXGrid = (args) => __async(void 0, null, function* () {
  var _a, _b, _c, _d, _e;
  const { data, config, rowAlignment, defaultFont, height, fileName, headers } = args;
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet(config.name, {
    views: [{ style: "pageBreakPreview" }],
    properties: {
      defaultRowHeight: (_a = config.colHeight) != null ? _a : ROW_HEIGHT
    },
    headerFooter: {
      oddFooter: "&F&RPage &P / &N"
    },
    pageSetup: {
      paperSize: 9,
      horizontalCentered: true,
      scale: (_b = config.zoom) != null ? _b : 100,
      orientation: (_c = config.orientation) != null ? _c : "portrait",
      margins: {
        top: 0.75,
        bottom: 0.75,
        left: 0.25,
        right: 0.25,
        header: 0.3,
        footer: 0.3
      }
    }
  });
  sheet.columns = headers;
  sheet.eachRow((row) => {
    row.height = height != null ? height : ROW_HEIGHT;
  });
  sheet.columns.forEach((column, index) => {
    var _a2;
    (_a2 = column.eachCell) == null ? void 0 : _a2.call(column, { includeEmpty: false }, (cell) => {
      var _a3, _b2, _c2, _d2;
      const header = headers.at(index);
      cell.border = (_a3 = header == null ? void 0 : header.border) != null ? _a3 : DEFAULT_BORDER;
      cell.font = (_b2 = header == null ? void 0 : header.font) != null ? _b2 : BOLD_FONT;
      column.alignment = (_c2 = header == null ? void 0 : header.alignment) != null ? _c2 : DEFAULT_ROW_ALIGNEMENT;
      cell.fill = {
        pattern: "solid",
        type: "pattern",
        fgColor: { argb: "FFD9D9D9" }
      };
      if ((header == null ? void 0 : header.isCurrency) || (header == null ? void 0 : header.isNumber)) {
        column.numFmt = ((_d2 = header == null ? void 0 : header.numFmt) != null ? _d2 : header.isCurrency) ? DEFAULT_NUM_FMT : "#,##;-#,##";
      }
    });
  });
  const dataRow = sheet.addRows(data);
  if (dataRow.length > 0) {
    dataRow.forEach((r, i) => {
      r.height = ROW_HEIGHT;
      r.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        var _a2, _b2, _c2;
        const header = headers.at(colNumber - 1);
        cell.border = (_a2 = header == null ? void 0 : header.border) != null ? _a2 : DEFAULT_BORDER;
        cell.alignment = (_b2 = header == null ? void 0 : header.alignment) != null ? _b2 : DEFAULT_ROW_ALIGNEMENT;
        cell.font = (_c2 = header == null ? void 0 : header.font) != null ? _c2 : DEFAULT_ROW_FONT;
      });
    });
  }
  const totauxFields = {};
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (h.hasTotal) {
      totauxFields[(_d = h.key) != null ? _d : "-"] = 0;
    } else {
      totauxFields[(_e = h.key) != null ? _e : "-"] = null;
    }
  }
  const total = sheet.addRow(totauxFields);
  total.height = ROW_HEIGHT;
  total.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    var _a2, _b2, _c2;
    const header = headers[colNumber - 1];
    cell.border = (_a2 = header == null ? void 0 : header.border) != null ? _a2 : DEFAULT_BORDER;
    cell.alignment = (_b2 = header == null ? void 0 : header.alignment) != null ? _b2 : DEFAULT_ROW_ALIGNEMENT;
    cell.font = (_c2 = header == null ? void 0 : header.font) != null ? _c2 : BOLD_FONT;
    if (header == null ? void 0 : header.hasTotal) {
      console.log(header);
      const col = cell.address.replace(/[0-9]+/gi, "");
      cell.value = {
        date1904: false,
        formula: `SUM(${col + "1"}:${col + (cell.row - 1)})`
      };
    }
  });
  if (data.length <= 0 || headers.length <= 0) return;
  const buffer = yield wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "applicationi/xlsx" });
  saveAs(blob, fileName.replace(/\.xlsx/gi, "") + ".xlsx");
});
export {
  generateXLSXGrid
};
//# sourceMappingURL=index.mjs.map