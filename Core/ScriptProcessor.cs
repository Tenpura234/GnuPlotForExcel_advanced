using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace GnuplotForExcel.Core
{
    /// <summary>
    /// gnuplot スクリプト内のデータ参照式を解析・展開します。
    /// （旧 VBA の gnuScriptOpen/Close + scriptExpr + contents に相当）
    /// </summary>
    /// <remarks>
    /// 式の構文:
    ///   &lt;&lt;"A1:B10"&gt;&gt;                … 現在のシートの範囲
    ///   &lt;&lt;"Sheet2!A1:B10"&gt;&gt;          … シート指定
    ///   &lt;&lt;"A1:A10,C1:C10"&gt;&gt;          … 飛び飛び範囲
    ///   &lt;&lt;matrix "Sheet1!A1:D10"&gt;&gt;   … マトリクス（splot/等高線用）
    /// 式はデータファイルへのパスに置き換えられます。
    /// </remarks>
    public class ScriptProcessor
    {
        /// <summary>&lt;&lt;"..."&gt;&gt; または &lt;&lt;matrix "..."&gt;&gt; に一致。</summary>
        private static readonly Regex ExpressionPattern = new Regex(
            @"<<\s*(?<matrix>matrix\s+)?(?<quote>[""'])(?<spec>.+?)\k<quote>\s*>>",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// スクリプト内の式から参照している範囲指定の一覧を取得します（動的リンク用）。
        /// </summary>
        public List<RangeReference> ExtractReferences(string script)
        {
            List<RangeReference> list = new List<RangeReference>();
            if (string.IsNullOrEmpty(script))
            {
                return list;
            }
            foreach (Match m in ExpressionPattern.Matches(script))
            {
                list.Add(new RangeReference
                {
                    Spec = m.Groups["spec"].Value,
                    IsMatrix = m.Groups["matrix"].Success
                });
            }
            return list;
        }

        /// <summary>
        /// スクリプト内の式を展開し、データファイルへのパスに置き換えたスクリプトを返します。
        /// 生成したデータファイルは tempFiles に登録されます。
        /// </summary>
        /// <param name="app">Excel Application。</param>
        /// <param name="script">元のスクリプト。</param>
        /// <param name="tempFiles">一時ファイル管理。</param>
        /// <param name="errorMessage">式評価に失敗した場合のメッセージ。</param>
        /// <returns>展開後スクリプト。失敗時は null。</returns>
        public string Expand(Excel.Application app, string script, TempFileManager tempFiles, out string errorMessage)
        {
            errorMessage = null;
            if (string.IsNullOrEmpty(script))
            {
                return script;
            }

            StringBuilder result = new StringBuilder(script);
            int offset = 0;

            foreach (Match m in ExpressionPattern.Matches(script))
            {
                string spec = m.Groups["spec"].Value;
                bool isMatrix = m.Groups["matrix"].Success;

                string dataFile = tempFiles.GetTempFile("gpdata", "dat");
                try
                {
                    RangeExporter.Export(app, spec, dataFile, isMatrix);
                }
                catch (Exception ex)
                {
                    errorMessage = string.Format("範囲 \"{0}\" の取得に失敗しました: {1}", spec, ex.Message);
                    return null;
                }

                // gnuplot ではバックスラッシュはエスケープされるためスラッシュに変換
                string gnuplotPath = "'" + dataFile.Replace('\\', '/') + "'";

                int index = m.Index + offset;
                result.Remove(index, m.Length);
                result.Insert(index, gnuplotPath);
                offset += gnuplotPath.Length - m.Length;
            }

            return result.ToString();
        }
    }

    /// <summary>
    /// スクリプト内のデータ参照（動的リンクの監視対象）。
    /// </summary>
    public class RangeReference
    {
        public string Spec { get; set; }
        public bool IsMatrix { get; set; }
    }
}
