/*
 * --------------------------------------------------------------------------------
 * 文件名:      SpreadsheetRow.cs
 * 作者:        yanwenfei
 * 创建日期:    2025-11-25
 * 描述: 
 * X-Spreadsheet 的展示对象SpreadsheetRow
 * --------------------------------------------------------------------------------
 * 变更日志:
 * --------------------------------------------------------------------------------
 */

namespace Spreadsheet.ExcelService.models
{
    public class SpreadsheetRow
    {
        public Dictionary<int, SpreadsheetCell> cells { get; set; } = new();
        public double? height { get; set; }
    }
}
