/*
 * --------------------------------------------------------------------------------
 * 文件名:      ExcelProcessService.cs
 * 作者:        yanwenfei
 * 创建日期:    2025-11-25
 * 描述: 
 * Excel处理服务实现
 * --------------------------------------------------------------------------------
 * 变更日志:
 * --------------------------------------------------------------------------------
 */

using ClosedXML.Excel;
using DocumentFormat.OpenXml.Validation;
using Spreadsheet.ExcelService.ExcelTransfer;
using Spreadsheet.ExcelService.models;
using System.Text.Json;

namespace Spreadsheet.Service
{
    /// <summary>
    /// Excel to  SpreadsheetSheet
    /// </summary>
    public class ExcelProcessService: IExcelProcessService
    {
   
        /// <summary>
        /// 转换模板工作表并填充数据
        /// </summary>
        /// <param name="xLWorksheets">模板文件的所有工作表</param>
        /// <param name="cellValues">工作表Cell的值配置</param>
        /// <returns></returns>
        public List<SpreadsheetSheet> ReadExcelToSpreadsheetSheet(IXLWorksheets xLWorksheets,
                 Dictionary<string, List<CellValue>>? sheetsCellValues=null)
        {
            List<SpreadsheetSheet> result = new List<SpreadsheetSheet>();
        
            foreach (var ws in xLWorksheets)
            {
                //确定填充数据
                List<CellValue>? CellFillData = null;
                var IsHasCellFillData = sheetsCellValues != null && sheetsCellValues.TryGetValue(ws.Name, out CellFillData);
                //单Sheet 自动识别填充数据
                if (sheetsCellValues != null&&!IsHasCellFillData&&sheetsCellValues.Count == 1&& xLWorksheets.Count == 1)
                {
                    IsHasCellFillData = true;
                    CellFillData = sheetsCellValues.Values.ElementAtOrDefault(0);
                }


                var xSheet = new SpreadsheetSheet
                {
                    name = ws.Name,
                    rows = new Dictionary<int, SpreadsheetRow>(),
                    cols = new Dictionary<int, object>(),
                    merges = new List<string>(),
                    styles = new List<object>()
                };

                // 1) 列宽（对齐 ExcelJS：width = Excel 列宽 * 8，兜底 100）
                // 使用 Column(1..LastColumnUsed) 保留空列可选：按需改为 ws.Columns(1, ws.LastColumnUsed().ColumnNumber())
                var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
                for (int colIdx = 1; colIdx <= lastCol; colIdx++)
                {
                    var col = ws.Column(colIdx);
                    double width = col.Width > 0 ? col.Width * 8 : 100;
                    xSheet.cols[colIdx - 1] = new
                    {
                        width
                    };
                }

                // 2) 合并单元格
                var mergeInfoMap = new Dictionary<string, (int YRange, int XRange)>();
                foreach (var range in ws.MergedRanges)
                {
                    var tl = range.FirstCell().Address.ToString(); // TopLeft address
                    int YRange = range.RowCount() - 1;             // 与前端一致：合并的行数 - 1
                    int XRange = range.ColumnCount() - 1;          // 合并的列数 - 1
                    mergeInfoMap[tl] = (YRange, XRange);
                    xSheet.merges.Add(range.RangeAddress.ToString()); // "A1:C3"
                }

                // 3) 行、单元格（includeEmpty 行列）
                var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
                // 预置默认样式，保证 style=0 可用
                if (xSheet.styles.Count == 0)
                    xSheet.styles.Add(new Dictionary<string, object>()); // 默认空样式



                for (int r1 = 1; r1 <= lastRow; r1++)
                {
                    int r = r1 - 1;
                    var rowObj = new SpreadsheetRow
                    {
                        cells = new Dictionary<int, SpreadsheetCell>(),
                        height = ws.Row(r1).Height > 0 ? ws.Row(r1).Height * 1.333 : null
                    };

                    for (int c1 = 1; c1 <= lastCol; c1++)
                    {
                        int c = c1 - 1;
                        var cell = ws.Cell(r1, c1);

                        // 文本/公式
                        string text = !string.IsNullOrEmpty(cell.FormulaA1)
                            ? "=" + cell.FormulaA1
                            : cell.Value.ToString() ?? "";




                        var style = CellTransfer.BuildCellStyle(cell);

                        // 样式去重
                        string styleJson = JsonSerializer.Serialize(style);
                        int styleIndex = xSheet.styles.FindIndex(s => JsonSerializer.Serialize(s) == styleJson);
                        if (styleIndex == -1)
                        {
                            styleIndex = xSheet.styles.Count;
                            xSheet.styles.Add(style);
                        }

                        var cellObj = new SpreadsheetCell
                        {
                            text = text,
                            style = styleIndex,
                            merge = null,
                            editable = true
                        };
                        if (IsHasCellFillData && CellFillData!=null && CellFillData.Count>0)
                        {
                            var currentCellFillData = CellFillData.FirstOrDefault(val=>(val.Row==r1&&val.Col==c1)
                            ||val.Address==cell.Address.ColumnLetter+ cell.Address.RowNumber);
                            if (currentCellFillData!=null)
                            {
                                cellObj.text = currentCellFillData.Value;
                                cellObj.editable = currentCellFillData.IsEditCell;
                            }
                        }

                        // 合并信息：只在左上角单元格标注
                        var addr = cell.Address.ToString();
                        if (mergeInfoMap.TryGetValue(addr, out var m))
                            cellObj.merge = new[] { m.YRange, m.XRange };

                        // ⚠️ 必须放在 rowObj.cells 下
                        rowObj.cells[c] = cellObj;
                    }

                    // ⚠️ 必须保证每行都有 cells
                    xSheet.rows[r] = rowObj;
                }
                result.Add(xSheet);
            }

            return result;


        }


    }
}
