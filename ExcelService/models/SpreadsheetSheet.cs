

/*
 * --------------------------------------------------------------------------------
 * 文件名:      SpreadsheetSheet.cs
 * 作者:        yanwenfei
 * 创建日期:    2025-11-25
 * 描述: 
 * X-Spreadsheet 的展示对象SpreadsheetSheet
 * --------------------------------------------------------------------------------
 * 变更日志:
 * --------------------------------------------------------------------------------
 */
namespace Spreadsheet.ExcelService.models
{
    public class SpreadsheetSheet
    {
        public string name { get; set; } = "Sheet1";
        public Dictionary<int, SpreadsheetRow> rows { get; set; } = new();
        public Dictionary<int, object> cols { get; set; } = new();
        public List<string> merges { get; set; } = new();
        public List<object> styles { get; set; } = new();
    }

}
