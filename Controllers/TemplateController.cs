using ClosedXML.Excel;

using Microsoft.AspNetCore.Mvc;
using Spreadsheet.ExcelService.ExcelTransfer;
using Spreadsheet.ExcelService.models;
using Spreadsheet.Service;
using System.Drawing;
using System.Text.Json;

namespace MyApp.Controllers
{
    [ApiController]
    [Route("api/template")]
    public class TemplateController(IExcelProcessService excelProcessService) : ControllerBase
    {
        [HttpGet("{code}")]
        public IActionResult GetTemplate(string code)
        {
            var filePath = System.IO.Path.Combine("ExcelTemplates", $"{code}.xlsx");
            if (!System.IO.File.Exists(filePath))
                return NotFound($"模板 {code} 不存在");
            using var workbook = new XLWorkbook(filePath);
            ExcelReadResponse excelReadResponse =new ExcelReadResponse ();
            //填充数据
            Dictionary<string, List<CellValue>> sheetCellValues = new Dictionary<string, List<CellValue>>();
            sheetCellValues.Add("Sheet1", new List<CellValue>
            {
               new CellValue("D3","999"),
               new CellValue("D5","888"),
               new CellValue("D7","=D5+D3",false)//公式非编辑项

            });
            excelReadResponse.data= excelProcessService.ReadExcelToSpreadsheetSheet(workbook.Worksheets, sheetCellValues);
            return Ok(excelReadResponse);
        }
    }




}
