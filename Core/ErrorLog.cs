using System;
using System.Text;

namespace GnuplotForExcel.Core
{
    /// <summary>
    /// アプリ内で共有するエラーログの蓄積。
    /// </summary>
    public static class ErrorLog
    {
        private static readonly StringBuilder _buffer = new StringBuilder();
        private static readonly object _lock = new object();

        public static void Append(string source, string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }
            lock (_lock)
            {
                _buffer.AppendLine(string.Format("[{0:yyyy-MM-dd HH:mm:ss}] {1}", DateTime.Now, source));
                _buffer.AppendLine(message.Trim());
                _buffer.AppendLine(new string('-', 60));
            }
        }

        public static void Append(string source, GnuplotResult result, string scriptPath)
        {
            if (result == null)
            {
                return;
            }
            Append(source, result.BuildErrorMessage(scriptPath));
        }

        public static string GetAll()
        {
            lock (_lock)
            {
                return _buffer.ToString();
            }
        }

        public static void Clear()
        {
            lock (_lock)
            {
                _buffer.Length = 0;
            }
        }
    }
}
