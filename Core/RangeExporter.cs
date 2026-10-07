using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace GnuplotForExcel.Core
{
    /// <summary>
    /// Excel のセル範囲を gnuplot のデータファイル形式に変換します。
    /// （旧 VBA の eval.txt / utilz.writeMat に相当）
    /// </summary>
    /// <remarks>
    /// 対応する指定形式:
    ///   "A1:B10"             … 現在のシートの連続範囲
    ///   "Sheet2!A1:B10"      … シート指定
    ///   "A1:A10,C1:C10"      … カンマ区切りの飛び飛び範囲（行位置で整列）
    ///   "matrix!Sheet1!A1:D10" … matrix 形式（splot / 等高線用の空行区切りグリッド）
    /// 数値はインバリアント カルチャで出力し、空白セルは NaN 扱い（"?"）にします。
    /// </remarks>
    public static class RangeExporter
    {
        /// <summary>gnuplot の欠損値表現。</summary>
        public const string MissingValue = "?";

        /// <summary>
        /// 範囲指定文字列をデータファイルにエクスポートします。
        /// </summary>
        /// <param name="app">Excel Application。</param>
        /// <param name="spec">範囲指定（例: "Sheet1!A1:A10,C1:C10"）。</param>
        /// <param name="filePath">出力先データファイル。</param>
        /// <param name="matrixMode">true の場合、grid（空行区切り）形式で出力。</param>
        public static void Export(Excel.Application app, string spec, string filePath, bool matrixMode)
        {
            Excel.Range range = ResolveRange(app, spec);
            if (range == null)
            {
                throw new ArgumentException("範囲を解決できません: " + spec);
            }

            if (matrixMode)
            {
                WriteMatrix(range, filePath);
            }
            else
            {
                WriteColumns(range, filePath);
            }
        }

        /// <summary>
        /// 範囲指定文字列を Excel.Range に解決します。カンマ区切りの複数範囲にも対応。
        /// </summary>
        public static Excel.Range ResolveRange(Excel.Application app, string spec)
        {
            string trimmed = (spec ?? string.Empty).Trim().Trim('"', '\'');

            // "Sheet!Range" 形式
            Excel.Worksheet sheet = null;
            string rangePart = trimmed;

            int bang = trimmed.LastIndexOf('!');
            if (bang >= 0)
            {
                string sheetName = trimmed.Substring(0, bang).Trim().Trim('\'');
                string rest = trimmed.Substring(bang + 1);
                sheet = FindSheet(app, sheetName);
                if (sheet == null)
                {
                    return null;
                }
                rangePart = rest;
            }

            Excel.Worksheet targetSheet = sheet ?? (Excel.Worksheet)app.ActiveSheet;
            if (targetSheet == null)
            {
                return null;
            }

            // カンマ区切りの飛び飛び範囲
            string[] parts = rangePart.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            try
            {
                string addresses = string.Join(",", parts.Select(p => p.Trim()));
                return targetSheet.Range[addresses];
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 指定がシート名を含むか（"Sheet!..." 形式）。
        /// </summary>
        public static bool TrySplitSpec(string spec, out string sheetName, out string rangePart)
        {
            sheetName = null;
            rangePart = (spec ?? string.Empty).Trim().Trim('"', '\'');
            int bang = rangePart.LastIndexOf('!');
            if (bang < 0)
            {
                return false;
            }
            sheetName = rangePart.Substring(0, bang).Trim().Trim('\'');
            rangePart = rangePart.Substring(bang + 1);
            return true;
        }

        private static Excel.Worksheet FindSheet(Excel.Application app, string name)
        {
            foreach (Excel.Worksheet ws in app.Worksheets)
            {
                if (string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return ws;
                }
            }
            return null;
        }

        /// <summary>
        /// 列データ（plot 用）を書き出します。飛び飛び範囲は各エリアを行位置で整列して1表にします。
        /// </summary>
        private static void WriteColumns(Excel.Range range, string filePath)
        {
            // 飛び飛び範囲（複数エリア）の場合、各エリアを列方向に結合する。
            // 行数が異なる場合は最大行数に合わせ、足りない分は欠損値にする。
            List<string[,]> areaTables = new List<string[,]>();
            int maxRows = 0;

            foreach (Excel.Range area in range.Areas)
            {
                string[,] table = ToStringTable(area);
                areaTables.Add(table);
                maxRows = Math.Max(maxRows, table.GetLength(0));
            }

            StringBuilder sb = new StringBuilder();
            for (int r = 0; r < maxRows; r++)
            {
                List<string> cells = new List<string>();
                foreach (string[,] table in areaTables)
                {
                    int rows = table.GetLength(0);
                    int cols = table.GetLength(1);
                    for (int c = 0; c < cols; c++)
                    {
                        cells.Add(r < rows ? table[r, c] : MissingValue);
                    }
                }
                sb.AppendLine(string.Join(" ", cells));
            }

            File.WriteAllText(filePath, sb.ToString(), new UTF8Encoding(false));
        }

        /// <summary>
        /// マトリクスデータ（splot / 等高線用）を空行区切りの grid 形式で書き出します。
        /// 先頭行を X ラベル、先頭列を Y ラベルとみなし、残りを Z 値とします。
        /// </summary>
        private static void WriteMatrix(Excel.Range range, string filePath)
        {
            object[,] values = range.Value2 as object[,];
            if (values == null)
            {
                // 単一セル
                File.WriteAllText(filePath, FormatValue(range.Value2) + Environment.NewLine, new UTF8Encoding(false));
                return;
            }

            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            StringBuilder sb = new StringBuilder();
            // 先頭行: Y ラベル列のプレースホルダ + X ラベル群
            sb.Append(MissingValue);
            for (int c = 2; c <= cols; c++)
            {
                sb.Append(' ').Append(FormatValue(values[1, c]));
            }
            sb.AppendLine();

            for (int r = 2; r <= rows; r++)
            {
                // 各行: Y ラベル + Z 値群
                sb.Append(FormatValue(values[r, 1]));
                for (int c = 2; c <= cols; c++)
                {
                    sb.Append(' ').Append(FormatValue(values[r, c]));
                }
                sb.AppendLine();
            }

            File.WriteAllText(filePath, sb.ToString(), new UTF8Encoding(false));
        }

        private static string[,] ToStringTable(Excel.Range area)
        {
            object[,] values = area.Value2 as object[,];
            int rows = area.Rows.Count;
            int cols = area.Columns.Count;
            string[,] table = new string[rows, cols];

            if (values != null)
            {
                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        table[r - 1, c - 1] = FormatValue(values[r, c]);
                    }
                }
            }
            else
            {
                // 単一セル
                table[0, 0] = FormatValue(area.Value2);
            }
            return table;
        }

        private static string FormatValue(object value)
        {
            if (value == null)
            {
                return MissingValue;
            }
            if (value is double || value is float || value is decimal)
            {
                return Convert.ToDouble(value).ToString("R", CultureInfo.InvariantCulture);
            }
            if (value is int || value is long || value is short)
            {
                return Convert.ToString(value, CultureInfo.InvariantCulture);
            }
            if (value is DateTime)
            {
                return ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }
            if (value is bool)
            {
                return (bool)value ? "1" : "0";
            }
            string text = value.ToString();
            return string.IsNullOrWhiteSpace(text) ? MissingValue : text;
        }
    }
}
