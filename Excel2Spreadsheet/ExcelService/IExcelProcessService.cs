
/*
 * --------------------------------------------------------------------------------
 * 文件名:      IExcelProcessService.cs
 * 作者:        yanwenfei
 * 创建日期:    2025-11-25
 * 描述: 
 * Excel  单元格转换服务接口
 * --------------------------------------------------------------------------------
 * 变更日志:
 * --------------------------------------------------------------------------------
 */
using ClosedXML.Excel;


using Spreadsheet.ExcelService.models;

namespace Spreadsheet.Service
{
    public interface IExcelProcessService
    {
        public List<SpreadsheetSheet> ReadExcelToSpreadsheetSheet(IXLWorksheets xLWorksheets,
            Dictionary<string, List<CellValue>>? sheetsCellValues=null);
    }
}
