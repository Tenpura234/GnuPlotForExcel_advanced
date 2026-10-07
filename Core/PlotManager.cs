using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace GnuplotForExcel.Core
{
    /// <summary>
    /// シェイプに埋め込むメタ情報のマーカー（旧 VBA の "#!GnuPlot" 相当）。
    /// </summary>
    public static class PlotMarker
    {
        public const string Title = "#!GnuplotForExcel";
        public const string AltPrefix = "#!gnuplot";
    }

    /// <summary>
    /// スクリプトに付随するメタ情報（ヘッダコメントに保持）。
    /// </summary>
    public class PlotMetadata
    {
        /// <summary>出力形式: "png" or "pdf"。</summary>
        public string OutputFormat { get; set; }
        /// <summary>動的リンク（自動更新）の有効/無効。</summary>
        public bool Linked { get; set; }
        /// <summary>生成ファイルをブックと同じディレクトリに保存するか。</summary>
        public bool SaveFile { get; set; }

        public PlotMetadata()
        {
            OutputFormat = "png";
            Linked = true;
            SaveFile = true;
        }

        public string FormatKey { get { return (OutputFormat ?? "png").Trim().ToLowerInvariant(); } }

        /// <summary>メタ情報をヘッダコメント文字列として生成。</summary>
        public string ToHeader()
        {
            return string.Format(CultureInfo.InvariantCulture,
                "{0} format={1} link={2} save={3}",
                PlotMarker.AltPrefix,
                FormatKey,
                Linked ? "1" : "0",
                SaveFile ? "1" : "0");
        }

        /// <summary>ヘッダコメントからメタ情報を解析。存在しなければ既定値。</summary>
        public static PlotMetadata Parse(string script)
        {
            PlotMetadata meta = new PlotMetadata();
            if (string.IsNullOrEmpty(script))
            {
                return meta;
            }
            using (StringReader reader = new StringReader(script))
            {
                string line = reader.ReadLine();
                if (line != null && line.StartsWith(PlotMarker.AltPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    foreach (Match m in Regex.Matches(line, @"(\w+)\s*=\s*([^\s]+)"))
                    {
                        string k = m.Groups[1].Value.ToLowerInvariant();
                        string v = m.Groups[2].Value;
                        if (k == "format") meta.OutputFormat = v;
                        else if (k == "link") meta.Linked = v == "1" || v.Equals("true", StringComparison.OrdinalIgnoreCase);
                        else if (k == "save") meta.SaveFile = v == "1" || v.Equals("true", StringComparison.OrdinalIgnoreCase);
                    }
                }
            }
            return meta;
        }

        /// <summary>スクリプトからメタヘッダ行を除去した本文を返します。</summary>
        public static string StripHeader(string script)
        {
            if (string.IsNullOrEmpty(script))
            {
                return string.Empty;
            }
            using (StringReader reader = new StringReader(script))
            {
                string line = reader.ReadLine();
                if (line != null && line.StartsWith(PlotMarker.AltPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    string rest = reader.ReadToEnd();
                    return rest != null ? rest.TrimStart('\r', '\n') : string.Empty;
                }
            }
            return script;
        }
    }

    /// <summary>
    /// プロット実行結果。
    /// </summary>
    public class PlotOutcome
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string ImagePath { get; set; }
        public string PdfPath { get; set; }
        public string ExpandedScript { get; set; }
        public GnuplotResult GnuplotResult { get; set; }
    }

    /// <summary>
    /// gnuplot スクリプトの実行からシートへの配置、ファイル保存、再描画までを統括します。
    /// （旧 VBA の doShape / doGnu に相当）
    /// </summary>
    public class PlotManager
    {
        private readonly Excel.Application _app;

        public PlotManager(Excel.Application app)
        {
            _app = app;
        }

        /// <summary>
        /// 新しいグラフを作成し、選択セル（またはアクティブセル）の右隣に配置します。
        /// </summary>
        /// <param name="scriptBody">メタヘッダを除くスクリプト本文。</param>
        /// <param name="meta">メタ情報。</param>
        /// <param name="anchor">配置基準セル（null ならアクティブセル）。</param>
        public PlotOutcome CreatePlot(string scriptBody, PlotMetadata meta, Excel.Range anchor)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)_app.ActiveSheet;
            if (sheet == null)
            {
                return Error("アクティブなワークシートがありません。");
            }

            Excel.Range anchorCell = anchor ?? _app.ActiveCell;
            if (anchorCell == null)
            {
                return Error("配置先セルを特定できません。");
            }

            // 選択セルの右隣に配置（作業セルの隣に自動で置く要件）
            Excel.Range target = anchorCell.Offset[0, 1];
            double left = target.Left;
            double top = target.Top;
            double width = 480;
            double height = 320;

            using (TempFileManager tempFiles = new TempFileManager())
            {
                PlotOutcome outcome = RenderToFile(scriptBody, meta, tempFiles, width, height);
                if (!outcome.Success)
                {
                    return outcome;
                }

                // ファイルをブックと同じディレクトリに保存
                if (meta.SaveFile)
                {
                    SaveBesideWorkbook(sheet, outcome, tempFiles);
                }

                // シートに画像を配置（PDF の場合は PNG プレビューを使用）
                string displayImage = !string.IsNullOrEmpty(outcome.ImagePath)
                    ? outcome.ImagePath
                    : outcome.PdfPath; // 通常 PDF は配置不可のため ImagePath がある想定

                if (!string.IsNullOrEmpty(displayImage) && File.Exists(displayImage))
                {
                    PlaceShape(sheet, displayImage, left, top, width, height, scriptBody, meta);
                }

                return outcome;
            }
        }

        /// <summary>
        /// 既存の Gnuplot シェイプを再描画します。
        /// </summary>
        public PlotOutcome RefreshShape(Excel.Shape shape)
        {
            Excel.Worksheet sheet = shape.Parent as Excel.Worksheet;
            if (sheet == null)
            {
                return Error("シェイプの親シートを取得できません。");
            }

            string script = shape.AlternativeText;
            if (string.IsNullOrWhiteSpace(script))
            {
                return Error("シェイプにスクリプトが保持されていません。");
            }

            PlotMetadata meta = PlotMetadata.Parse(script);
            string body = PlotMetadata.StripHeader(script);

            double left = shape.Left;
            double top = shape.Top;
            double width = shape.Width;
            double height = shape.Height;
            string name = shape.Name;

            using (TempFileManager tempFiles = new TempFileManager())
            {
                PlotOutcome outcome = RenderToFile(body, meta, tempFiles, width, height);
                if (!outcome.Success)
                {
                    return outcome;
                }

                if (meta.SaveFile)
                {
                    SaveBesideWorkbook(sheet, outcome, tempFiles);
                }

                string displayImage = !string.IsNullOrEmpty(outcome.ImagePath)
                    ? outcome.ImagePath
                    : outcome.PdfPath;

                if (!string.IsNullOrEmpty(displayImage) && File.Exists(displayImage))
                {
                    Excel.Shape newShape = PlaceShape(sheet, displayImage, left, top, width, height, body, meta);
                    if (newShape != null)
                    {
                        try
                        {
                            shape.Delete();
                            try { newShape.Name = name; } catch { }
                        }
                        catch { }
                    }
                }

                return outcome;
            }
        }

        /// <summary>
        /// ワークシート内の Gnuplot シェイプをすべて取得します。
        /// </summary>
        public List<Excel.Shape> GetPlotShapes(Excel.Worksheet sheet)
        {
            List<Excel.Shape> list = new List<Excel.Shape>();
            if (sheet == null)
            {
                return list;
            }
            foreach (Excel.Shape shape in sheet.Shapes)
            {
                if (IsPlotShape(shape))
                {
                    list.Add(shape);
                }
            }
            return list;
        }

        /// <summary>シェイプが Gnuplot プロットかどうかを判定します。</summary>
        public static bool IsPlotShape(Excel.Shape shape)
        {
            try
            {
                return string.Equals(shape.Title, PlotMarker.Title, StringComparison.Ordinal);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 選択中のシェイプが Gnuplot プロットであれば返します。
        /// </summary>
        public Excel.Shape GetSelectedPlotShape()
        {
            try
            {
                dynamic sel = _app.Selection;
                if (sel == null)
                {
                    return null;
                }
                Excel.ShapeRange sr = sel.ShapeRange as Excel.ShapeRange;
                if (sr == null || sr.Count < 1)
                {
                    return null;
                }
                Excel.Shape shape = sr.Item(1);
                return IsPlotShape(shape) ? shape : null;
            }
            catch
            {
                return null;
            }
        }

        // ----- 内部処理 -----

        /// <summary>
        /// スクリプトを gnuplot で実行し、画像/PDF ファイルを生成します。
        /// </summary>
        private PlotOutcome RenderToFile(string scriptBody, PlotMetadata meta, TempFileManager tempFiles,
            double widthPx, double heightPx)
        {
            ScriptProcessor processor = new ScriptProcessor();
            string expandError;
            string expanded = processor.Expand(_app, scriptBody, tempFiles, out expandError);
            if (expanded == null)
            {
                return Error(expandError);
            }

            Settings settings = Settings.Default;
            bool isPdf = string.Equals(meta.FormatKey, "pdf", StringComparison.OrdinalIgnoreCase);

            string imagePath = null;
            string pdfPath = null;
            string outputPath;

            if (isPdf)
            {
                pdfPath = tempFiles.GetTempFile("gpout", "pdf");
                // シート表示用の PNG も生成する
                imagePath = tempFiles.GetTempFile("gpout", "png");
                outputPath = pdfPath;
            }
            else
            {
                imagePath = tempFiles.GetTempFile("gpout", "png");
                outputPath = imagePath;
            }

            // terminal / output を先頭に付与
            string terminalLine;
            if (isPdf)
            {
                string sizeIn = string.Format(CultureInfo.InvariantCulture,
                    " size {0}in,{1}in", widthPx / 96.0, heightPx / 96.0);
                terminalLine = "set terminal pdfcairo " + settings.PdfTerminalOptions + sizeIn;
            }
            else
            {
                string sizePx = string.Format(CultureInfo.InvariantCulture,
                    " size {0},{1}", (int)Math.Round(widthPx), (int)Math.Round(heightPx));
                terminalLine = "set terminal pngcairo " + settings.PngTerminalOptions + sizePx;
            }

            StringBuilder script = new StringBuilder();
            script.AppendLine(terminalLine);
            script.AppendLine("set output '" + outputPath.Replace('\\', '/') + "'");
            script.AppendLine(expanded);

            GnuplotRunner runner = new GnuplotRunner(settings.GnuplotExePath);
            string scriptPath;
            GnuplotResult result = runner.RunScript(script.ToString(), tempFiles, "gpscript",
                out scriptPath);

            if (result.HasError)
            {
                PlotOutcome err = Error(result.BuildErrorMessage(scriptPath));
                err.GnuplotResult = result;
                err.ExpandedScript = script.ToString();
                return err;
            }

            // PDF の場合、シート表示用 PNG を別途生成
            if (isPdf && imagePath != null)
            {
                StringBuilder pngScript = new StringBuilder();
                pngScript.AppendLine("set terminal pngcairo " + settings.PngTerminalOptions +
                    string.Format(CultureInfo.InvariantCulture, " size {0},{1}", (int)Math.Round(widthPx), (int)Math.Round(heightPx)));
                pngScript.AppendLine("set output '" + imagePath.Replace('\\', '/') + "'");
                pngScript.AppendLine(expanded);
                string pngScriptPath;
                GnuplotResult pngResult = runner.RunScript(pngScript.ToString(), tempFiles, "gpscript_png",
                    out pngScriptPath);
                if (pngResult.HasError)
                {
                    // PNG プレビュー失敗は致命的ではない（PDF は生成済み）
                    imagePath = null;
                }
            }

            return new PlotOutcome
            {
                Success = true,
                ImagePath = imagePath,
                PdfPath = pdfPath,
                ExpandedScript = expanded,
                GnuplotResult = result
            };
        }

        /// <summary>
        /// 生成ファイルをブックと同じディレクトリに保存します。
        /// </summary>
        private void SaveBesideWorkbook(Excel.Worksheet sheet, PlotOutcome outcome, TempFileManager tempFiles)
        {
            try
            {
                Excel.Workbook wb = (Excel.Workbook)sheet.Parent;
                string wbPath = wb.Path;
                if (string.IsNullOrEmpty(wbPath))
                {
                    // 未保存のブックの場合はデスクトップに保存
                    wbPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                }

                string baseName = string.Format("gnuplot_{0:yyyyMMdd_HHmmss}", DateTime.Now);

                if (!string.IsNullOrEmpty(outcome.ImagePath) && File.Exists(outcome.ImagePath))
                {
                    string dest = Path.Combine(wbPath, baseName + ".png");
                    File.Copy(outcome.ImagePath, dest, true);
                    outcome.ImagePath = dest;
                    tempFiles.Detach(dest);
                }
                if (!string.IsNullOrEmpty(outcome.PdfPath) && File.Exists(outcome.PdfPath))
                {
                    string dest = Path.Combine(wbPath, baseName + ".pdf");
                    File.Copy(outcome.PdfPath, dest, true);
                    outcome.PdfPath = dest;
                    tempFiles.Detach(dest);
                }
            }
            catch
            {
                // ファイル保存の失敗はプロット自体を失敗にしない
            }
        }

        /// <summary>
        /// シートに画像シェイプを配置し、メタ情報を埋め込みます。
        /// </summary>
        private Excel.Shape PlaceShape(Excel.Worksheet sheet, string imagePath, double left, double top,
            double width, double height, string scriptBody, PlotMetadata meta)
        {
            try
            {
                Excel.Shape shape = sheet.Shapes.AddPicture(
                    Filename: imagePath,
                    LinkToFile: Office.MsoTriState.msoFalse,
                    SaveWithDocument: Office.MsoTriState.msoTrue,
                    Left: (float)left,
                    Top: (float)top,
                    Width: (float)width,
                    Height: (float)height);

                shape.Title = PlotMarker.Title;
                shape.AlternativeText = meta.ToHeader() + Environment.NewLine + scriptBody;

                // 外枠
                try
                {
                    shape.Line.Visible = Office.MsoTriState.msoTrue;
                    shape.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                }
                catch { }

                // 名前を設定（GnuPlot プレフィックス）
                try
                {
                    string baseName = "GnuPlot";
                    string newName = baseName;
                    int i = 1;
                    while (ShapeExists(sheet, newName))
                    {
                        newName = baseName + (++i);
                    }
                    shape.Name = newName;
                }
                catch { }

                return shape;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("PlaceShape failed: " + ex.Message);
                return null;
            }
        }

        private static bool ShapeExists(Excel.Worksheet sheet, string name)
        {
            try
            {
                Excel.Shape s = sheet.Shapes.Item(name);
                return s != null;
            }
            catch
            {
                return false;
            }
        }

        private static PlotOutcome Error(string message)
        {
            return new PlotOutcome { Success = false, ErrorMessage = message };
        }
    }
}
