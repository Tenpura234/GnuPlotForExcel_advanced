using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace GnuplotForExcel.Core
{
    /// <summary>
    /// アドインの設定を保持します。HKCU\Software\GnuplotForExcel に保存します。
    /// （旧 VBA の registry.txt / gnu.txt の設定処理に相当します）
    /// </summary>
    public class Settings
    {
        private const string RootKey = @"Software\GnuplotForExcel";
        private const string PresetKey = RootKey + @"\Presets";

        private static readonly Lazy<Settings> _instance = new Lazy<Settings>(() => new Settings());

        public static Settings Default
        {
            get { return _instance.Value; }
        }

        private Settings()
        {
            Load();
        }

        /// <summary>gnuplot 実行ファイルのパス。</summary>
        public string GnuplotExePath { get; set; }

        /// <summary>PNG 出力時の追加 terminal オプション。</summary>
        public string PngTerminalOptions { get; set; }

        /// <summary>PDF 出力時の追加 terminal オプション。</summary>
        public string PdfTerminalOptions { get; set; }

        /// <summary>生成画像/PDF をブックと同じディレクトリに保存するか。</summary>
        public bool SaveFileNextToWorkbook { get; set; }

        /// <summary>セル変更に連動した自動更新（動的リンク）を有効にするか。</summary>
        public bool AutoRefreshEnabled { get; set; }

        /// <summary>既定の出力形式 ("png" / "pdf")。</summary>
        public string DefaultOutputFormat { get; set; }

        private void Load()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RootKey))
            {
                GnuplotExePath = GetString(key, "GnuplotExePath", @"C:\Program Files\gnuplot\bin\gnuplot.exe");
                PngTerminalOptions = GetString(key, "PngTerminalOptions", "color font 'Arial,11'");
                if (PngTerminalOptions != null && PngTerminalOptions.StartsWith("cairo ", StringComparison.OrdinalIgnoreCase))
                {
                    PngTerminalOptions = PngTerminalOptions.Substring("cairo ".Length).TrimStart();
                }
                PdfTerminalOptions = GetString(key, "PdfTerminalOptions", "color font 'Arial,11'");
                DefaultOutputFormat = GetString(key, "DefaultOutputFormat", "png");
                SaveFileNextToWorkbook = GetBool(key, "SaveFileNextToWorkbook", true);
                AutoRefreshEnabled = GetBool(key, "AutoRefreshEnabled", true);
            }
        }

        public void Save()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RootKey))
            {
                key.SetValue("GnuplotExePath", GnuplotExePath ?? string.Empty, RegistryValueKind.String);
                key.SetValue("PngTerminalOptions", PngTerminalOptions ?? string.Empty, RegistryValueKind.String);
                key.SetValue("PdfTerminalOptions", PdfTerminalOptions ?? string.Empty, RegistryValueKind.String);
                key.SetValue("DefaultOutputFormat", DefaultOutputFormat ?? "png", RegistryValueKind.String);
                key.SetValue("SaveFileNextToWorkbook", SaveFileNextToWorkbook ? 1 : 0, RegistryValueKind.DWord);
                key.SetValue("AutoRefreshEnabled", AutoRefreshEnabled ? 1 : 0, RegistryValueKind.DWord);
            }
        }

        // ----- プリセット -----

        /// <summary>書式プリセット名の一覧を取得します。</summary>
        public List<string> GetPresetNames()
        {
            List<string> names = new List<string>();
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(PresetKey))
            {
                if (key != null)
                {
                    names.AddRange(key.GetSubKeyNames());
                }
            }
            names.Sort(StringComparer.OrdinalIgnoreCase);
            return names;
        }

        /// <summary>プリセットを読み込みます。存在しない場合は null。</summary>
        public string LoadPreset(string name)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(PresetKey + @"\" + name))
            {
                if (key == null)
                {
                    return null;
                }
                return key.GetValue("Script") as string;
            }
        }

        /// <summary>プリセットを保存します。</summary>
        public void SavePreset(string name, string script)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(PresetKey + @"\" + name))
            {
                key.SetValue("Script", script ?? string.Empty, RegistryValueKind.String);
            }
        }

        /// <summary>プリセットを削除します。</summary>
        public void DeletePreset(string name)
        {
            Registry.CurrentUser.DeleteSubKeyTree(PresetKey + @"\" + name, false);
        }

        private static string GetString(RegistryKey key, string name, string defaultValue)
        {
            if (key == null)
            {
                return defaultValue;
            }
            object value = key.GetValue(name);
            return value != null ? value.ToString() : defaultValue;
        }

        private static bool GetBool(RegistryKey key, string name, bool defaultValue)
        {
            if (key == null)
            {
                return defaultValue;
            }
            object value = key.GetValue(name);
            if (value == null)
            {
                return defaultValue;
            }
            int intValue;
            return int.TryParse(value.ToString(), out intValue) ? intValue != 0 : defaultValue;
        }
    }
}
