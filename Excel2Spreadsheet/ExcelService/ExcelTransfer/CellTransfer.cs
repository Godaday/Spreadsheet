/*
 * --------------------------------------------------------------------------------
 * 文件名:      CellTransfer.cs
 * 作者:        yanwenfei
 * 创建日期:    2025-11-25
 * 描述: 
 * Excel  单元格转换
 * --------------------------------------------------------------------------------
 * 变更日志:
 * --------------------------------------------------------------------------------
 */

using ClosedXML.Excel;
using System.Globalization;


namespace Spreadsheet.ExcelService.ExcelTransfer
{
    public class CellTransfer
    {
        public static string GetCellValue(IXLCell cell)
        {
            // 1. 公式优先：如果单元格是公式，返回公式字符串（带'='）
            if (!string.IsNullOrEmpty(cell.FormulaA1))
            {
                // 返回公式本身，而不是公式的计算结果
                return "=" + cell.FormulaA1;
            }

            // 2. 处理不同数据类型的数值
            switch (cell.DataType)
            {
                case XLDataType.DateTime:
                    // 如果 ClosedXML 已经识别为 DateTime 类型，直接获取并格式化
                    return cell.GetDateTime().ToString("yyyy/MM/dd");

                case XLDataType.Number:
                    // 如果数据类型是 Number：
                    // 使用 InvariantCulture 确保数字格式（如小数点）在不同环境下一致
                    return cell.GetDouble().ToString(CultureInfo.InvariantCulture);

                case XLDataType.Text:
                    // 文本类型，直接获取字符串
                    return cell.GetString();

                case XLDataType.Boolean:
                    // 布尔值，转换为字符串（如 "TRUE" 或 "FALSE"）
                    return cell.GetBoolean().ToString().ToUpperInvariant();

                case XLDataType.Blank:
                    // 空白单元格
                    return string.Empty;

                default:
                    // 默认情况，使用 Value.ToString() 作为后备方案
                    return cell.Value.ToString() ?? string.Empty;
            }
        }
        public static Dictionary<string, object> BuildCellStyle(IXLCell cell)
        {
            var style = new Dictionary<string, object>();

            // 背景色
            string? bg = ThemeColor.GetHexColorSafe(cell.Style.Fill.BackgroundColor);
            if (!string.IsNullOrEmpty(bg))
            {
                // 调用 TinyColorHelper 转换成 #RRGGBB
                style["bgcolor"] = bg;
            }

            // 字体颜色
            string? fontHex = ThemeColor.GetHexColorSafe(cell.Style.Font.FontColor);
            if (!string.IsNullOrEmpty(fontHex))
            {
                style["color"] = fontHex;
            }

            // 其他字体属性
            var fontObj = new Dictionary<string, object>();
            if (!string.IsNullOrEmpty(cell.Style.Font.FontName)) fontObj["fontName"] = cell.Style.Font.FontName;
            if (cell.Style.Font.Bold) fontObj["bold"] = true;
            if (cell.Style.Font.Italic) fontObj["italic"] = true;
            if (cell.Style.Font.Underline != XLFontUnderlineValues.None) fontObj["underline"] = true;
            if (cell.Style.Font.FontSize > 0) fontObj["size"] = (int)Math.Round(cell.Style.Font.FontSize);
            if (fontObj.Count > 0) style["font"] = fontObj;

            // 对齐（有限集合）
            var h = cell.Style.Alignment.Horizontal;
            var v = cell.Style.Alignment.Vertical;
            if (h is XLAlignmentHorizontalValues.Left or XLAlignmentHorizontalValues.Center or XLAlignmentHorizontalValues.Right)
                style["align"] = h.ToString().ToLowerInvariant();
            // 默认 middle/top/bottom 映射
            style["valign"] = v switch
            {
                XLAlignmentVerticalValues.Top => "top",
                XLAlignmentVerticalValues.Center => "middle",
                XLAlignmentVerticalValues.Bottom => "bottom",
                _ => "middle"
            };



            // 边框（仅 Thin/Medium/Thick + 允许色）
            var bd = new Dictionary<string, object>();
            AddBorder(bd, "top", cell.Style.Border.TopBorder, cell.Style.Border.TopBorderColor);
            AddBorder(bd, "bottom", cell.Style.Border.BottomBorder, cell.Style.Border.BottomBorderColor);
            AddBorder(bd, "left", cell.Style.Border.LeftBorder, cell.Style.Border.LeftBorderColor);
            AddBorder(bd, "right", cell.Style.Border.RightBorder, cell.Style.Border.RightBorderColor);
            if (bd.Count > 0) style["border"] = bd;

            return style;
        }

        private static void AddBorder(Dictionary<string, object> bd, string side, XLBorderStyleValues styleEnum, XLColor color)
        {
            if (styleEnum == XLBorderStyleValues.None) return;

            string style = styleEnum switch
            {
                XLBorderStyleValues.Thin => "thin",
                XLBorderStyleValues.Medium => "medium",
                XLBorderStyleValues.Thick => "thick",
                _ => "thin" // 其他样式统一降级
            };

            var hex = ThemeColor.GetHexColorSafe(color) ?? "#000000";
            // 再次确保在允许色集合
            hex = MapToAllowedColor(hex);
            var argb = HexToArgb(hex);

            bd[side] = new object[] { style, argb };
        }
        // 仅允许的字体名（按需扩展）
        public static readonly HashSet<string> AllowedFonts = new(StringComparer.OrdinalIgnoreCase)
{
    "Arial","Calibri","Microsoft YaHei","SimSun","宋体","Tahoma","Verdana"
};

        private static readonly HashSet<string> AllowedHexColors = new()
{
    "#000000","#FFFFFF",
    "#FF0000","#00FF00","#0000FF",
    "#FFFF00","#00FFFF","#FF00FF",
    "#C0C0C0","#808080","#800000","#808000",
    "#008000","#800080","#008080","#000080"
};


        private static string MapToAllowedColor(string hex)
        {
            if (AllowedHexColors.Contains(hex)) return hex;


            int r = Convert.ToInt32(hex.Substring(1, 2), 16);
            int g = Convert.ToInt32(hex.Substring(3, 2), 16);
            int b = Convert.ToInt32(hex.Substring(5, 2), 16);

            string best = "#000000";
            double bestDist = double.MaxValue;
            foreach (var allowed in AllowedHexColors)
            {
                int ar = Convert.ToInt32(allowed.Substring(1, 2), 16);
                int ag = Convert.ToInt32(allowed.Substring(3, 2), 16);
                int ab = Convert.ToInt32(allowed.Substring(5, 2), 16);
                double dist = (r - ar) * (r - ar) + (g - ag) * (g - ag) + (b - ab) * (b - ab);
                if (dist < bestDist) { bestDist = dist; best = allowed; }
            }
            return best;
        }


        private static string HexToArgb(string hex)
        {
            var clean = hex.Trim().TrimStart('#');
            return "FF" + clean.ToUpperInvariant();
        }

    }
}
