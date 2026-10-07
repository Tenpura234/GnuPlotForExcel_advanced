using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using GnuplotForExcel.Core;

namespace GnuplotForExcel.Forms
{
    /// <summary>
    /// gnuplot.exe のパスや出力オプションを設定するダイアログ。
    /// （旧 VBA の GnuPlotSetup フォームに相当）
    /// </summary>
    public class SetupForm : Form
    {
        private readonly TextBox _txtExePath;
        private readonly TextBox _txtPngOptions;
        private readonly TextBox _txtPdfOptions;
        private readonly ComboBox _cmbDefaultFormat;
        private readonly CheckBox _chkSaveFile;
        private readonly CheckBox _chkAutoRefresh;

        public SetupForm()
        {
            Text = "GnuplotForExcel 設定";
            StartPosition = FormStartPosition.CenterParent;
            Width = 560;
            Height = 340;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            Font = SystemFonts.MessageBoxFont;

            int labelWidth = 150;
            int rowHeight = 30;
            int y = 12;

            // gnuplot パス
            Label lblExe = MakeLabel("gnuplot.exe のパス:", labelWidth, y);
            _txtExePath = new TextBox { Left = labelWidth + 8, Top = y, Width = 300 };
            Button btnBrowse = new Button { Text = "参照...", Left = labelWidth + 312, Top = y - 1, Width = 70 };
            btnBrowse.Click += (s, e) => BrowseExe();

            y += rowHeight + 8;
            Label lblPng = MakeLabel("PNG オプション:", labelWidth, y);
            _txtPngOptions = new TextBox { Left = labelWidth + 8, Top = y, Width = 382 };

            y += rowHeight + 8;
            Label lblPdf = MakeLabel("PDF オプション:", labelWidth, y);
            _txtPdfOptions = new TextBox { Left = labelWidth + 8, Top = y, Width = 382 };

            y += rowHeight + 8;
            Label lblFmt = MakeLabel("既定の出力形式:", labelWidth, y);
            _cmbDefaultFormat = new ComboBox { Left = labelWidth + 8, Top = y, Width = 90, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbDefaultFormat.Items.AddRange(new object[] { "png", "pdf" });

            y += rowHeight + 8;
            _chkSaveFile = new CheckBox { Left = labelWidth + 8, Top = y, Width = 300, Text = "生成ファイルをブックと同じ場所に保存" };

            y += rowHeight;
            _chkAutoRefresh = new CheckBox { Left = labelWidth + 8, Top = y, Width = 300, Text = "セル変更に連動して自動更新（動的リンク）" };

            // ボタン
            Panel bottom = new Panel { Dock = DockStyle.Bottom, Height = 44 };
            Button btnOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Dock = DockStyle.Right, Width = 95 };
            Button btnCancel = new Button { Text = "キャンセル", DialogResult = DialogResult.Cancel, Dock = DockStyle.Right, Width = 95 };
            bottom.Controls.Add(btnOk);
            bottom.Controls.Add(btnCancel);

            Controls.Add(lblExe);
            Controls.Add(_txtExePath);
            Controls.Add(btnBrowse);
            Controls.Add(lblPng);
            Controls.Add(_txtPngOptions);
            Controls.Add(lblPdf);
            Controls.Add(_txtPdfOptions);
            Controls.Add(lblFmt);
            Controls.Add(_cmbDefaultFormat);
            Controls.Add(_chkSaveFile);
            Controls.Add(_chkAutoRefresh);
            Controls.Add(bottom);

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            Load += (s, e) => LoadSettings();
            FormClosing += SetupForm_FormClosing;
        }

        private static Label MakeLabel(string text, int width, int top)
        {
            return new Label { Text = text, Left = 8, Top = top + 3, Width = width, TextAlign = ContentAlignment.MiddleLeft };
        }

        private void LoadSettings()
        {
            Settings s = Settings.Default;
            _txtExePath.Text = s.GnuplotExePath;
            _txtPngOptions.Text = s.PngTerminalOptions;
            _txtPdfOptions.Text = s.PdfTerminalOptions;
            _cmbDefaultFormat.SelectedItem = s.DefaultOutputFormat == "pdf" ? "pdf" : "png";
            if (_cmbDefaultFormat.SelectedIndex < 0) _cmbDefaultFormat.SelectedIndex = 0;
            _chkSaveFile.Checked = s.SaveFileNextToWorkbook;
            _chkAutoRefresh.Checked = s.AutoRefreshEnabled;
        }

        private void SetupForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult != DialogResult.OK)
            {
                return;
            }

            string exe = (_txtExePath.Text ?? string.Empty).Trim();
            if (exe.Length == 0 || !File.Exists(exe))
            {
                MessageBox.Show("gnuplot.exe のパスが正しくありません。", "GnuplotForExcel",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
                return;
            }

            Settings s = Settings.Default;
            s.GnuplotExePath = exe;
            s.PngTerminalOptions = _txtPngOptions.Text;
            s.PdfTerminalOptions = _txtPdfOptions.Text;
            s.DefaultOutputFormat = _cmbDefaultFormat.SelectedItem != null ? _cmbDefaultFormat.SelectedItem.ToString() : "png";
            s.SaveFileNextToWorkbook = _chkSaveFile.Checked;
            s.AutoRefreshEnabled = _chkAutoRefresh.Checked;
        }

        private void BrowseExe()
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Title = "gnuplot.exe を選択";
                dlg.Filter = "gnuplot.exe|gnuplot.exe|実行ファイル (*.exe)|*.exe|すべてのファイル (*.*)|*.*";
                if (File.Exists(_txtExePath.Text))
                {
                    dlg.InitialDirectory = Path.GetDirectoryName(_txtExePath.Text);
                    dlg.FileName = Path.GetFileName(_txtExePath.Text);
                }
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _txtExePath.Text = dlg.FileName;
                }
            }
        }
    }
}
