/*
 * --------------------------------------------------------------------------------
 * 文件名:      ExcelReadResponse.cs
 * 作者:        yanwenfei
 * 创建日期:    2025-11-25
 * 描述: 
 * Excel转换结果对象
 * --------------------------------------------------------------------------------
 * 变更日志:
 * --------------------------------------------------------------------------------
 */

namespace Spreadsheet.ExcelService.models
{
    public class ExcelReadResponse
    {
        public ExcelViewConfig config { get; set; }
        public List<SpreadsheetSheet> data { get; set; } = new List<SpreadsheetSheet>();
    }
}
