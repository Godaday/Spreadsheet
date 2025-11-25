import ExcelJS from "exceljs";
import _ from "lodash";
import tinycolor from "tinycolor2";
/**
 * ==========================================================
 * 样式辅助函数 (Style Helper Functions)
 * ==========================================================
 */

/**
 * 将 ExcelJS 的 ARGB 颜色对象或字符串转换为十六进制 (#RRGGBB)
 */
function getHexColor(color) {
  let argb;
  if (typeof color === "object" && color !== null && color.argb) {
    argb = color.argb;
  } else if (typeof color === "string") {
    argb = color;
  } else {
    return null;
  }

  if (!argb || typeof argb !== "string") return null;
  argb = argb.replace(/^#/, "").toUpperCase();

  if (argb.length === 8) {
    return `#${argb.substring(2)}`; // 忽略 alpha
  }
  if (argb.length === 6) {
    return `#${argb}`;
  }
  return null;
}

/**
 * 将 #RRGGBB 转换为 ExcelJS 需要的 ARGB (FFRRGGBB)
 */
function hexToArgb(hex) {
  // if (!hex) return null;
  // hex = hex.replace("#", "").toUpperCase();
  // if (hex.length === 6) {
  //   return "FF" + hex;
  // }
  // return null;
    var rgb = tinycolor(hex).toRgb()
      var rHex = parseInt(rgb.r).toString(16).padStart(2, '0')
      var gHex = parseInt(rgb.g).toString(16).padStart(2, '0')
      var bHex = parseInt(rgb.b).toString(16).padStart(2, '0')
      var aHex = parseInt(rgb.a).toString(16).padStart(2, '0')
      var res = aHex + rHex + gHex + bHex
      return res;
}

/**
 * 水平对齐映射
 */
function mapXAlignToExcel(xAlign) {
  if (xAlign === "left" || xAlign === "center" || xAlign === "right") return xAlign;
  return undefined;
}

/**
 * 垂直对齐映射
 */
function mapXVAlignToExcel(xVAlign) {
  if (xVAlign === "top" || xVAlign === "middle" || xVAlign === "bottom") return xVAlign;
  return undefined;
}

/**
 * ==========================================================
 * 导入函数 stox (ExcelJS Workbook -> x-spreadsheet Data)
 * ==========================================================
 */
export async function stox(wb) {
  const sheetPromises = [];

  wb.eachSheet((sheet) => {
    sheetPromises.push(new Promise(resolve => {
      const xSheet = {
        name: sheet.name,
        rows: {},
        cols: {},
        merges: [],
        styles: [],
      };

      // 列宽
      sheet.columns.forEach((col, idx) => {
        xSheet.cols[idx] = { width: col.width ? col.width * 8 : 100 };
      });

      // 合并单元格
      const mergeInfoMap = {};
      if (sheet._merges) {
        Object.values(sheet._merges).forEach((m) => {
          const YRange = m.model.bottom - m.model.top;
          const XRange = m.model.right - m.model.left;
          mergeInfoMap[m.tl] = { YRange, XRange };
          xSheet.merges.push(m.shortRange);
        });
      }

      // 行和单元格
      sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const r = rowNumber - 1;
        const rowHeight = row.height ? row.height * 1.333 : undefined;
        xSheet.rows[r] = { cells: {}, height: rowHeight };

        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const c = colNumber - 1;
          let text = "";

          if (cell.value?.formula) text = "=" + cell.value.formula;
          else if (cell.value?.result !== undefined) text = cell.value.result;
          else if (cell.value?.richText) text = cell.value.richText.map(t => t.text).join("");
          else if (cell.value !== undefined && cell.value !== null) text = cell.value;

          const style = {};
          const cellStyle = cell.style;
          const cellFont = cellStyle.font;
  if (cellStyle.border) {
           const { left, right, top, bottom } = cellStyle.border;
              let bd = {};
              if (top) {
                bd.top = _.isArray(top) ? top : [top.style||"thin", hexToArgb(top.color?.argb||"FF000000")];
              }
              if (bottom) {
                bd.bottom = _.isArray(bottom) ? bottom : [bottom.style||"thin", hexToArgb(bottom.color?.argb||"FF000000")];
              }
              if (left) {
                bd.left = _.isArray(left) ? left : [left.style||"thin", hexToArgb(left.color?.argb||"FF000000")];
              }
              if (right) {
                bd.right = _.isArray(right) ? right : [right.style||"thin", hexToArgb(right.color?.argb||"FF000000")];
              }

              cellStyle.border = bd;
            }

          if (cellStyle.fill?.fgColor) {
            const bgColor = getHexColor(cellStyle.fill.fgColor);
            if (bgColor) style.bgcolor = bgColor;
          }

          if (cell.protection?.locked === true) cell.editable = false;

          if (cellStyle.alignment?.horizontal) style.align = cellStyle.alignment.horizontal;
          if (cellStyle.alignment?.vertical) style.valign = cellStyle.alignment.vertical;

          if (cellFont) {
            style.font = {};
            const fontColor = cellFont.color ? getHexColor(cellFont.color) : null;
            if (fontColor) style.color = fontColor;
            if (cellFont.size) style.font.size = cellFont.size;
            if (cellFont.name) style.fontName = cellFont.name;
            if (cellFont.bold) style.font.bold = true;
            if (cellFont.italic) style.font.italic = true;
            if (cellFont.underline) style.font.underline = true;
          }

          const styleJson = JSON.stringify(style);
          let styleIndex = xSheet.styles.findIndex(s => JSON.stringify(s) === styleJson);
          if (styleIndex === -1) {
            styleIndex = xSheet.styles.length;
            xSheet.styles.push(style);
          }

          const cellObject = { text: String(text), style: styleIndex };
          const mergeInfo = mergeInfoMap[cell.address];
          if (mergeInfo) cellObject.merge = [mergeInfo.YRange, mergeInfo.XRange];

          xSheet.rows[r].cells[c] = cellObject;
        });
      });

      resolve(xSheet);
    }));
  });

  return Promise.all(sheetPromises);
}

/**
 * ==========================================================
 * 导出函数 xtos (x-spreadsheet Data -> ExcelJS Workbook)
 * ==========================================================
 */
export async function xtos(sdata) {
  const wb = new ExcelJS.Workbook();

  for (const xSheet of sdata || []) {
    const sheet = wb.addWorksheet(xSheet?.name || "Sheet1");

    // 列宽
    if (xSheet?.cols) {
      const columns = Object.keys(xSheet.cols).map((idx) => {
        const col = xSheet.cols[idx];
        return { width: col?.width ? col.width / 8 : 12 };
      });
      sheet.columns = columns;
    }

    // 行和单元格
    if (xSheet?.rows) {
      Object.entries(xSheet.rows).forEach(([rIdx, rowData]) => {
        if (!rowData) return;

        const rowNumber = parseInt(rIdx, 10) + 1;
        const row = sheet.getRow(rowNumber);

        if (rowData.height) row.height = rowData.height / 1.333;

        if (rowData.cells) {
          Object.entries(rowData.cells).forEach(([cIdx, cellData]) => {
            if (!cellData) return;

            const colNumber = parseInt(cIdx, 10) + 1;
            const cell = row.getCell(colNumber);

            if (cellData.text?.startsWith("=")) {
              cell.value = { formula: cellData.text.slice(1) };
            } else {
              cell.value = cellData.text ?? "";
            }

           
              const styleIndex = cellData.style;
if (styleIndex !== undefined && xSheet.styles?.[styleIndex]) {
  const style = xSheet.styles[styleIndex];
  const cellStyle = {};

  if (style.bgcolor) {
    cellStyle.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: hexToArgb(style.bgcolor) },
    };
  }

  if (style.align || style.valign) {
    cellStyle.alignment = {};
    const hAlign = mapXAlignToExcel(style.align);
    const vAlign = mapXVAlignToExcel(style.valign);
    if (hAlign) cellStyle.alignment.horizontal = hAlign;
    if (vAlign) cellStyle.alignment.vertical = vAlign;
  }

  if (style.font || style.color || style.fontName) {
    cellStyle.font = {};
    if (style.color) cellStyle.font.color = { argb: hexToArgb(style.color) };
    if (style.font?.size) cellStyle.font.size = style.font.size;
    if (style.fontName) cellStyle.font.name = style.fontName;
    if (style.font?.bold) cellStyle.font.bold = true;
    if (style.font?.italic) cellStyle.font.italic = true;
    if (style.font?.underline) cellStyle.font.underline = true;
  }

  // ✅ 边框处理
if (style.border) {
  cellStyle.border = {};
  ["top","left","bottom","right"].forEach(pos => {
    if (style.border[pos]) {
      cellStyle.border[pos] = {
        style: style.border[pos].style || "thin", // ExcelJS 需要合法的边框样式
        color: style.border[pos].color
          ? { argb: hexToArgb(style.border[pos].color) } // 转换为 ARGB
          : undefined
      };
    }
  });
}

  cell.style = cellStyle;


            }

            if (cellData.merge) {
              const [rowSpan, colSpan] = cellData.merge;
              const startRow = rowNumber;
              const startCol = colNumber;
              const endRow = startRow + rowSpan;
              const endCol = startCol + colSpan;
              sheet.mergeCells(startRow, startCol, endRow, endCol);
            }
          });
        }
      });
    }
  }

  return wb;
}
