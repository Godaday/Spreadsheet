
/*
 * --------------------------------------------------------------------------------
 * 文件名:      CellValue.cs
 * 作者:        yanwenfei
 * 创建日期:    2025-11-25
 * 描述: 
 * Excel 单元格值对象
 * --------------------------------------------------------------------------------
 * 变更日志:
 * --------------------------------------------------------------------------------
 */

using Spreadsheet.ExcelService.ExcelTransfer;

namespace Spreadsheet.ExcelService.models
{
    /// <summary>
    /// 单元格值对象
    /// </summary>
    public class CellValue
    {
        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="address">单元格地址 例如 A1、BB26</param>
        /// <param name="value">值或计算公式</param>
        /// <param name="isReadOnly">是否不允许编辑，默认允许</param>
        public CellValue( string address,string value,
            bool isEditCell = true, 
            IndexMode mode = IndexMode.OneBased)
        {
            Address = address;
            var zeroBase= ExcelCoordinateConverter.A1ToNumeric(address, mode);
            Row = zeroBase.row;
            Col = zeroBase.col;
            Value = value;
            IsEditCell = isEditCell;

        }
       
        /// <summary>
        /// A1 引用地址，例如 "A1", "H28"。
        /// </summary>
        public string Address { get; }

        /// <summary>
        /// 0-based 行索引 (从 0 开始)。
        /// </summary>
        public int? Row { get;}

        /// <summary>
        /// 0-based 列索引 (从 0 开始)。
        /// </summary>
        public int Col { get;  }

        /// <summary>
        /// 具体值或公式
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// 是否编辑项 默认只读
        /// </summary>
        public bool IsEditCell { get; set; }=true;


    }
}
