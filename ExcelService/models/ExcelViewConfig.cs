/*
 * --------------------------------------------------------------------------------
 * 文件名:      ExcelViewConfig.cs
 * 作者:        yanwenfei
 * 创建日期:    2025-11-25
 * 描述: 
 * Excel展示的配置对象
 * --------------------------------------------------------------------------------
 * 变更日志:
 * --------------------------------------------------------------------------------
 */


namespace Spreadsheet.ExcelService.models
{
    public class ExcelViewConfig
    {
        /// <summary>
        /// 表格的行数长度。
        /// 计算公式: (lastRow > 0 ? lastRow : 1) + 2
        /// </summary>
        public int rowLen { get; set; }

        /// <summary>
        /// 表格的列数长度。
        /// 计算公式: (lastCol > 0 ? lastCol : 1) + 2
        /// </summary>
        public int colLen { get; set; }

        /// <summary>
        /// 视图模式（例如："edit", "readonly"）。
        /// 默认值: "edit"
        /// </summary>
        public string mode { get; set; } = "edit";

        /// <summary>
        /// 是否显示工具栏。
        /// 默认值: true
        /// </summary>
        public bool showToolbar { get; set; } = true;

        /// <summary>
        /// 是否显示网格线。
        /// 默认值: true
        /// </summary>
        public bool showGrid { get; set; } = true;

        /// <summary>
        /// 是否显示右键上下文菜单。
        /// 默认值: true
        /// </summary>
        public bool showContextmenu { get; set; } = true;

        /// <summary>
        /// 是否显示底部状态栏/信息栏。
        /// 默认值: true
        /// </summary>
        public bool showBottomBar { get; set; } = true;

        /// <summary>
        /// 构造函数，用于根据实际的行和列使用量计算 rowLen 和 colLen。
        /// </summary>
        /// <param name="lastRow">Excel 文件中使用的最大行号 (1-based)。</param>
        /// <param name="lastCol">Excel 文件中使用的最大列号 (1-based)。</param>
        public ExcelViewConfig(int lastRow, int lastCol)
        {
            // 确保至少有 1 行/列，然后 + 2 作为缓冲区/留白
            this.rowLen = (lastRow > 0 ? lastRow : 1) + 2;
            this.colLen = (lastCol > 0 ? lastCol : 1) + 2;
        }

        // 无参构造函数，方便序列化/反序列化（如果需要）
        public ExcelViewConfig() : this(0, 0) { }
    
}
}
