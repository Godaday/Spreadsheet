using Spreadsheet.ExcelService.models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Spreadsheet.ExcelService
{
    public static class MergeHelper
    {
        /// <summary>
        /// 自动对指定列进行纵向合并（不会破坏模板已有合并）
        /// </summary>
        public static void AutoMergeRows(SpreadsheetSheet sheet, params int[] mergeCols)
        {
            int totalRows = sheet.rows.Count;

            foreach (int col in mergeCols)
            {
                int start = 0;

                for (int r = 1; r <= totalRows; r++)
                {
                    bool isLastRow = (r == totalRows);

                    SpreadsheetCell prevCell = sheet.rows[r - 1].cells.ContainsKey(col)
                                               ? sheet.rows[r - 1].cells[col]
                                               : null;

                    SpreadsheetCell currCell = isLastRow ? null :
                                               (sheet.rows[r].cells.ContainsKey(col)
                                               ? sheet.rows[r].cells[col]
                                               : null);

                    string prev = prevCell?.text;
                    string curr = currCell?.text;

                    bool prevValid = !string.IsNullOrWhiteSpace(prev);
                    bool currValid = !string.IsNullOrWhiteSpace(curr);

                    bool isSame = (!isLastRow && prevValid && currValid && prev == curr);

                    if (!isSame)
                    {
                        // 执行合并 start ~ r-1
                        if (prevValid && (r - 1 > start))
                        {
                            ApplySafeMerge(sheet, start, r - 1, col);
                        }

                        start = r;
                    }
                }
            }
        }

        /// <summary>
        /// 安全合并逻辑：不会破坏模板中已存在的 merge
        /// 1) 在 cell 上设置 top/second/remove-tail（你原有风格）
        /// 2) 同时把合并范围写入 sheet.merges（兼容 List<string> 或 List&lt;SpreadsheetMerge&gt;）
        /// </summary>
        private static void ApplySafeMerge(SpreadsheetSheet sheet, int startRow, int endRow, int col)
        {
            if (endRow <= startRow)
                return;

            // 计算跨度（endRow - startRow 为你原来的语义）
            int rowSpan = endRow - startRow;

            var topCell = sheet.rows[startRow].cells[col];

            // ❗ 如果 topCell 已经被模板合并，不能覆盖
            if (topCell.merge != null)
                return;

            // ========== 设置 top cell ==========
            topCell.merge = new int[] { rowSpan, 0 };

            // ========== 设置 second cell ==========
            //var secondCell = sheet.rows[startRow + 1].cells[col];

            //// 模板已有 merge → 跳过整个合并
            //if (secondCell.merge != null)
            //    return;

            //secondCell.text = "";
            //secondCell.merge = new int[] { rowSpan - 1, 0 };

            // ========== 删除其余行的 cell ==========
            for (int r = startRow + 1; r <= endRow; r++)
            {
                if (sheet.rows[r].cells.ContainsKey(col))
                    sheet.rows[r].cells.Remove(col);
            }

            // ========== 写入 sheet.merges ==========
            // 支持三种情况：
            // 1) sheet.merges == List<string> -> 添加 "A1:A3"
            // 2) sheet.merges == List<SpreadsheetMerge> -> 添加 对象
            // 3) sheet.merges is IList -> 尝试添加字符串形式
            try
            {
                // 1-based coordinates for ranges
                int startRow1 = startRow + 1;
                int endRow1 = endRow + 1;
                string colLetter = ColumnNumberToName(col + 1);
                string rangeStr = $"{colLetter}{startRow1}:{colLetter}{endRow1}";

                if (sheet.merges == null)
                {
                    // Try to initialize as List<string> (safe default)
                    try
                    {
                        sheet.merges = (dynamic)Activator.CreateInstance(typeof(List<string>));
                    }
                    catch
                    {
                        // if cannot create, leave null and return
                        return;
                    }
                }

                // If it's List<string>
                if (sheet.merges is List<string> listStr)
                {
                    if (!listStr.Contains(rangeStr))
                        listStr.Add(rangeStr);
                    return;
                }

           

                // If it's a non-generic IList (or some other IList), try to add the string representation
                if (sheet.merges is System.Collections.IList anyList)
                {
                    // avoid duplicates
                    bool found = false;
                    foreach (var item in anyList)
                    {
                        if (item is string s && s == rangeStr) { found = true; break; }
                    }
                    if (!found)
                        anyList.Add(rangeStr);
                    return;
                }
            }
            catch
            {
                // 忽略写入 merges 失败，不要抛出，避免影响主流程
            }
        }

        // 辅助：把列号(1-based)转为列字母（1->A, 27->AA）
        private static string ColumnNumberToName(int col)
        {
            const int letters = 26;
            string name = String.Empty;
            while (col > 0)
            {
                int rem = (col - 1) % letters;
                name = (char)('A' + rem) + name;
                col = (col - 1) / letters;
            }
            return name;
        }
    }


}
