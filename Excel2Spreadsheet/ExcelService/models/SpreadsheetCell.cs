/*
 * --------------------------------------------------------------------------------
 * 文件名:      SpreadsheetCell.cs
 * 作者:        yanwenfei
 * 创建日期:    2025-11-25
 * 描述: 
 * X-Spreadsheet 的展示对象SpreadsheetCell
 * --------------------------------------------------------------------------------
 * 变更日志:
 * --------------------------------------------------------------------------------
 */

namespace Spreadsheet.ExcelService.models
{
    public class SpreadsheetCell
    {
        public string text { get; set; } = "";
        public int style { get; set; }
        public int[]? merge { get; set; }
        public bool? editable { get; set; }
    }
}
