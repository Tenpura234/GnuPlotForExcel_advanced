using System;
using System.Collections.Generic;
using System.IO;

namespace GnuplotForExcel.Core
{
    /// <summary>
    /// 一時ファイル/フォルダの作成と後始末を管理します。
    /// （旧 VBA の tempFile.txt に相当し、Win32 API ではなく .NET の Path/IO を使用します）
    /// </summary>
    public sealed class TempFileManager : IDisposable
    {
        private readonly List<string> _tempFiles = new List<string>();
        private readonly string _baseFolder;

        public TempFileManager()
        {
            _baseFolder = Path.Combine(Path.GetTempPath(), "GnuplotForExcel");
        }

        /// <summary>一時ファイル用ベースフォルダを返し、存在しなければ作成します。</summary>
        public string BaseFolder
        {
            get
            {
                if (!Directory.Exists(_baseFolder))
                {
                    Directory.CreateDirectory(_baseFolder);
                }
                return _baseFolder;
            }
        }

        /// <summary>
        /// 一意の一時ファイルパスを生成します（ファイル自体は作成しません）。
        /// </summary>
        public string GetTempFile(string prefix, string extension)
        {
            string fileName = string.Format("{0}_{1:yyyyMMdd_HHmmss}_{2}.{3}",
                prefix, DateTime.Now, Guid.NewGuid().ToString("N").Substring(0, 8),
                extension.TrimStart('.'));
            string path = Path.Combine(BaseFolder, fileName);
            _tempFiles.Add(path);
            return path;
        }

        /// <summary>テキストを UTF-8 (BOM なし) で書き込みます。</summary>
        public void WriteFile(string path, string content)
        {
            File.WriteAllText(path, content ?? string.Empty, new System.Text.UTF8Encoding(false));
        }

        /// <summary>テキストファイルを読み込みます。存在しない場合は空文字。</summary>
        public string ReadFile(string path)
        {
            return File.Exists(path) ? File.ReadAllText(path) : string.Empty;
        }

        /// <summary>管理下の一時ファイルをすべて削除します。</summary>
        public void Cleanup()
        {
            foreach (string path in _tempFiles)
            {
                TryDelete(path);
            }
            _tempFiles.Clear();
        }

        /// <summary>指定ファイルを削除対象から外します（ディレクトリ保存などで残す場合）。</summary>
        public void Detach(string path)
        {
            _tempFiles.Remove(path);
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // 一時ファイルの削除失敗は致命的ではないため無視
            }
        }

        public void Dispose()
        {
            Cleanup();
        }
    }
}
