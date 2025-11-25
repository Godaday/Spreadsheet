
/*
 * --------------------------------------------------------------------------------
 * 文件名:      ThemeColor.cs
 * 作者:        yanwenfei
 * 创建日期:    2025-11-25
 * 描述: 
 * Excel 颜色处理
 * --------------------------------------------------------------------------------
 * 变更日志:
 * --------------------------------------------------------------------------------
 */
using ClosedXML.Excel;
using System.Drawing;

namespace Spreadsheet.ExcelService.ExcelTransfer
{

    public static class ThemeColor
    {


        // Office 默认主题基色表
    private static readonly Dictionary<string, Color> ThemeBaseColors = new()
    {
       { "Text1",       Color.FromArgb(0, 0, 0) },       // 黑色
    { "Text2",       Color.FromArgb(255, 255, 255) }, // 白色
    { "Background1", Color.FromArgb(255, 255, 255) }, // 白色
    { "Background2", Color.FromArgb(0, 0, 0) },       // 黑色

    { "Accent1",     Color.FromArgb(0, 112, 192) },   // 蓝
    { "Accent2",     Color.FromArgb(237, 125, 49) },  // 橙
    { "Accent3",     Color.FromArgb(255, 192, 0) },   // 金
    { "Accent4",     Color.FromArgb(117, 189, 66) },  // 酸橙绿 (#75BD42)
    { "Accent5",     Color.FromArgb(91, 155, 213) },  // 青
    { "Accent6",     Color.FromArgb(255, 0, 0) },     // 红
        { "Hyperlink",          Color.FromArgb(5, 99, 193) },   // 链接蓝
        { "FollowedHyperlink",  Color.FromArgb(149, 79, 114) }  // 已访问链接紫
    };

        /// <summary>
        /// Tint 算法：Excel 用来调整亮度
        /// </summary>
        private static Color ApplyTint(Color baseColor, double tint)
        {
            double r = baseColor.R / 255.0;
            double g = baseColor.G / 255.0;
            double b = baseColor.B / 255.0;

            if (tint < 0)
            {
                r *= 1.0 + tint;
                g *= 1.0 + tint;
                b *= 1.0 + tint;
            }
            else
            {
                r = r + (1.0 - r) * tint;
                g = g + (1.0 - g) * tint;
                b = b + (1.0 - b) * tint;
            }

            return Color.FromArgb(
                (int)(r * 255),
                (int)(g * 255),
                (int)(b * 255)
            );
        }

        /// <summary>
        /// 将 XLColor 转换为 #RRGGBB
        /// </summary>
        public static string? GetHexColorSafe(ClosedXML.Excel.XLColor xlColor)
        {
            if (xlColor == null) return null;

            // 普通 RGB 色
            if (xlColor.ColorType == XLColorType.Color)
            {
                var c = xlColor.Color;
                if (!c.IsEmpty)
                {
                    if (c.A == 0) return null; // 透明色
                    return $"#{c.R:X2}{c.G:X2}{c.B:X2}";
                }
            }
            // Theme 色
            else if (xlColor.ColorType == XLColorType.Theme)
            {
                string themeName = xlColor.ThemeColor.ToString(); // "Accent4", "Text1" 等
                double tint = xlColor.ThemeTint;

                if (ThemeBaseColors.TryGetValue(themeName, out var baseColor))
                {
                    var tinted = ApplyTint(baseColor, tint);
                    return $"#{tinted.R:X2}{tinted.G:X2}{tinted.B:X2}";
                }
            }

            return null;
        }
    }

}
