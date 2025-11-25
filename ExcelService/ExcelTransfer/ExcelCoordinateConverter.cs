
/*
 * --------------------------------------------------------------------------------
 * 文件名:      ZeroBasedCoordinate.cs
 * 作者:        yanwenfei
 * 创建日期:    2025-11-25
 * 描述: 
 * Excel 颜色处理
 * --------------------------------------------------------------------------------
 * 变更日志:
 * --------------------------------------------------------------------------------
 */

using System;
using System.Text;
using System.Text.RegularExpressions;
namespace Spreadsheet.ExcelService.ExcelTransfer
{


    /// <summary>
    /// 定义坐标索引的起始模式。
    /// </summary>
    public enum IndexMode
    {
        /// <summary>
        /// 0-based 索引 (从 0 开始, 例如: A1 -> 0, 0)。
        /// </summary>
        ZeroBased,

        /// <summary>
        /// 1-based 索引 (从 1 开始, 例如: A1 -> 1, 1)。
        /// </summary>
        OneBased
    }

    public static class ExcelCoordinateConverter
    {
        // =========================================================================
        // 1. A1 地址转换为行列索引
        // =========================================================================

        /// <summary>
        /// 将 Excel 的 A1 引用坐标转换为指定模式的行和列索引。
        /// </summary>
        /// <param name="a1Address">Excel 单元格地址，例如 "A1"。</param>
        /// <param name="mode">指定返回的索引是 0-based 还是 1-based。</param>
        /// <returns>包含 (行索引, 列索引) 的元组。</returns>
        /// <exception cref="FormatException">如果地址格式无效。</exception>
        public static (int row, int col) A1ToNumeric(string a1Address, IndexMode mode= IndexMode.OneBased)
        {
            if (string.IsNullOrWhiteSpace(a1Address))
            {
                throw new ArgumentNullException(nameof(a1Address), "A1 address cannot be null or empty.");
            }

            var match = System.Text.RegularExpressions.Regex.Match(a1Address.ToUpperInvariant(), @"^([A-Z]+)([0-9]+)$");
            if (!match.Success)
            {
                throw new FormatException($"Invalid Excel address format: {a1Address}");
            }

            string columnLetters = match.Groups[1].Value;
            string rowString = match.Groups[2].Value;

            // 核心：计算 1-based 索引
            if (!int.TryParse(rowString, out int oneBasedRow) || oneBasedRow < 1)
            {
                throw new FormatException($"Invalid row number in address: {a1Address}");
            }

            int oneBasedCol = 0;
            for (int i = 0; i < columnLetters.Length; i++)
            {
                char c = columnLetters[i];
                oneBasedCol = oneBasedCol * 26 + (c - 'A' + 1);
            }

            if (oneBasedCol < 1)
            {
                throw new FormatException($"Invalid column letters in address: {a1Address}");
            }

            // 根据模式调整结果
            if (mode == IndexMode.ZeroBased)
            {
                return (row: oneBasedRow - 1, col: oneBasedCol - 1);
            }
            else // IndexMode.OneBased
            {
                return (row: oneBasedRow, col: oneBasedCol);
            }
        }


        // =========================================================================
        // 2. 行列索引转换为 A1 地址
        // =========================================================================

        /// <summary>
        /// 将指定模式的行和列索引转换为 Excel 的 A1 引用坐标。
        /// </summary>
        /// <param name="row">行索引。</param>
        /// <param name="col">列索引。</param>
        /// <param name="mode">指定输入的索引是 0-based 还是 1-based。</param>
        /// <returns>Excel 单元格地址，例如 "A1"。</returns>
        /// <exception cref="ArgumentOutOfRangeException">如果索引小于模式的有效最小值。</exception>
        public static string NumericToA1(int row, int col, IndexMode mode= IndexMode.OneBased)
        {
            int minIndex = (mode == IndexMode.ZeroBased) ? 0 : 1;

            if (row < minIndex || col < minIndex)
            {
                throw new ArgumentOutOfRangeException("Row or Column index is less than the minimum allowed index for the specified mode.");
            }

            // 核心：转换为 1-based 索引
            int oneBasedRow = (mode == IndexMode.ZeroBased) ? row + 1 : row;
            int oneBasedCol = (mode == IndexMode.ZeroBased) ? col + 1 : col;

            // 1. 转换列索引为列字母 (26进制反向转换)
            var sb = new StringBuilder();
            int tempCol = oneBasedCol;

            while (tempCol > 0)
            {
                tempCol--; // 关键步骤：先减 1
                char letter = (char)('A' + (tempCol % 26));
                sb.Insert(0, letter);
                tempCol /= 26;
            }

            // 2. 拼接结果：列字母 + 行号
            return $"{sb}{oneBasedRow}";
        }
    }
}
