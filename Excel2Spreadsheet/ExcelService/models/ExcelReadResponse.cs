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
    public class ExcelReadResponse<T>
    {
        /// <summary>
        /// 报表组件配置
        /// </summary>
        public T Config { get; set; }
        /// <summary>
        /// 报表Id标识
        /// </summary>
        public int ReportId { get; set; }
        /// <summary>
        /// 报表编码
        /// </summary>
        public string ReportCode { get; set; }
        /// <summary>
        /// 报表数据
        /// </summary>
        public List<SpreadsheetSheet> ReportData { get; set; } = new List<SpreadsheetSheet>();
    }
}
