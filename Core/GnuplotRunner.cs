using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace GnuplotForExcel.Core
{
    /// <summary>
    /// gnuplot 実行の結果を表します。
    /// </summary>
    public class GnuplotResult
    {
        public int ExitCode { get; set; }
        public string StandardOutput { get; set; }
        public string StandardError { get; set; }
        public bool TimedOut { get; set; }

        public bool HasError
        {
            get
            {
                if (TimedOut || ExitCode != 0)
                {
                    return true;
                }
                if (!string.IsNullOrWhiteSpace(StandardError))
                {
                    // gnuplot の fit ログ等は stderr に流れるため正常時は無視し、エラー時のみ検出
                    if (StandardError.IndexOf("*** FIT ERROR ***", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                    if (System.Text.RegularExpressions.Regex.IsMatch(StandardError, @"line\s+\d+:", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    {
                        return true;
                    }
                }
                return false;
            }
        }

        /// <summary>ユーザー表示用のエラーメッセージを組み立てます。</summary>
        public string BuildErrorMessage(string scriptPath)
        {
            StringBuilder sb = new StringBuilder();
            if (TimedOut)
            {
                sb.AppendLine("gnuplot がタイムアウトしました。");
            }
            if (ExitCode != 0)
            {
                sb.AppendLine("終了コード: " + ExitCode);
            }
            if (!string.IsNullOrWhiteSpace(StandardError))
            {
                string err = StandardError;
                if (!string.IsNullOrEmpty(scriptPath))
                {
                    // スクリプトのフルパス表記を読みやすく置き換え
                    err = err.Replace("\"" + scriptPath + "\", ", string.Empty);
                    err = err.Replace(scriptPath + ", ", string.Empty);
                }
                sb.AppendLine(err.Trim());
            }
            return sb.ToString().Trim();
        }
    }

    /// <summary>
    /// gnuplot プロセスの起動と出力キャプチャを行います。
    /// （旧 VBA の gnu.doGnu 内の WScript.Shell 実行に相当し、System.Diagnostics.Process で実装）
    /// </summary>
    public class GnuplotRunner
    {
        /// <summary>既定のタイムアウト（ミリ秒）。</summary>
        public const int DefaultTimeoutMs = 30000;

        private readonly string _gnuplotExePath;

        public GnuplotRunner(string gnuplotExePath)
        {
            _gnuplotExePath = gnuplotExePath;
        }

        /// <summary>
        /// gnuplot スクリプトファイルを実行し、結果を返します。
        /// </summary>
        public GnuplotResult Run(string scriptPath, int timeoutMs = DefaultTimeoutMs)
        {
            if (string.IsNullOrWhiteSpace(_gnuplotExePath) || !File.Exists(_gnuplotExePath))
            {
                return new GnuplotResult
                {
                    ExitCode = -1,
                    StandardError = "gnuplot.exe が見つかりません: " + _gnuplotExePath +
                                    Environment.NewLine + "設定ダイアログでパスを確認してください。"
                };
            }

            GnuplotResult result = new GnuplotResult();

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = _gnuplotExePath,
                Arguments = "\"" + scriptPath + "\"",
                WorkingDirectory = Path.GetDirectoryName(scriptPath),
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            using (Process process = new Process())
            {
                process.StartInfo = psi;

                StringBuilder stdout = new StringBuilder();
                StringBuilder stderr = new StringBuilder();

                process.OutputDataReceived += (s, e) => { if (e.Data != null) stdout.AppendLine(e.Data); };
                process.ErrorDataReceived += (s, e) => { if (e.Data != null) stderr.AppendLine(e.Data); };

                try
                {
                    process.Start();
                }
                catch (Exception ex)
                {
                    result.ExitCode = -1;
                    result.StandardError = "gnuplot の起動に失敗しました: " + ex.Message;
                    return result;
                }

                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                if (!process.WaitForExit(timeoutMs))
                {
                    result.TimedOut = true;
                    try { process.Kill(); } catch { }
                }

                // 非同期読み取りの残りを確実に取り込む
                process.WaitForExit();

                result.ExitCode = result.TimedOut ? -1 : process.ExitCode;
                result.StandardOutput = stdout.ToString();
                result.StandardError = stderr.ToString();
            }

            return result;
        }

        /// <summary>
        /// スクリプト文字列を一時ファイルに保存して実行します。
        /// </summary>
        public GnuplotResult RunScript(string script, TempFileManager tempFiles, string prefix,
            out string scriptPath)
        {
            return RunScript(script, tempFiles, prefix, DefaultTimeoutMs, out scriptPath);
        }

        /// <summary>
        /// スクリプト文字列を一時ファイルに保存して実行します（タイムアウト指定）。
        /// </summary>
        public GnuplotResult RunScript(string script, TempFileManager tempFiles, string prefix,
            int timeoutMs, out string scriptPath)
        {
            scriptPath = tempFiles.GetTempFile(prefix, "gp");
            tempFiles.WriteFile(scriptPath, script);
            return Run(scriptPath, timeoutMs);
        }
    }
}
