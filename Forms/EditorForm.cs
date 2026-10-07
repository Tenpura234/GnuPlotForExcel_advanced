using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using GnuplotForExcel.Core;

namespace GnuplotForExcel.Forms
{
    /// <summary>
    /// Gnuplot スクリプトの内蔵エディタとプレビュー。
    /// </summary>
    /// <remarks>
    /// - 既存グラフの編集（shape を渡す）と新規作成（null を渡す）の両方に対応。
    /// - プレビューは一時的に png で実行して PictureBox に表示。
    /// - データ挿入ボタンで選択範囲の式（&lt;&lt;"..."&gt;&gt; / &lt;&lt;matrix "..."&gt;&gt;）を挿入。
    /// </remarks>
    public class EditorForm : Form
    {
        private readonly Excel.Application _app;
        private readonly Excel.Shape _editingShape;

        private readonly RichTextBox _txtScript;
        private readonly PictureBox _preview;
        private readonly ComboBox _cmbFormat;
        private readonly CheckBox _chkLinked;
        private readonly CheckBox _chkSaveFile;
        private readonly ComboBox _cmbPreset;
        private readonly Label _lblStatus;

        private string _insertedExpression;

        /// <summary>編集対象シェイプのスクリプト本文（メタヘッダ除く）。</summary>
        public string ScriptBody
        {
            get { return _txtScript.Text; }
        }

        /// <summary>ユーザーが選択したメタ情報。</summary>
        public PlotMetadata Metadata
        {
            get
            {
                return new PlotMetadata
                {
                    OutputFormat = _cmbFormat.SelectedItem != null ? _cmbFormat.SelectedItem.ToString() : "png",
                    Linked = _chkLinked.Checked,
                    SaveFile = _chkSaveFile.Checked
                };
            }
        }

        public EditorForm(Excel.Application app, Excel.Shape editingShape)
        {
            _app = app;
            _editingShape = editingShape;

            Text = editingShape != null ? "Gnuplot スクリプト編集" : "Gnuplot 新規グラフ";
            StartPosition = FormStartPosition.CenterParent;
            Width = 980;
            Height = 640;
            FormBorderStyle = FormBorderStyle.Sizable;
            ShowIcon = false;
            Font = SystemFonts.MessageBoxFont;

            // 上部ツール
            Panel top = new Panel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(6, 4, 6, 4) };

            Label lblPreset = new Label { Text = "プリセット:", AutoSize = true, Dock = DockStyle.Left, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(0, 6, 0, 0), Width = 70 };
            _cmbPreset = new ComboBox { Dock = DockStyle.Left, Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };
            Button btnLoadPreset = new Button { Text = "読込", Dock = DockStyle.Left, Width = 60 };
            Button btnSavePreset = new Button { Text = "保存", Dock = DockStyle.Left, Width = 60 };

            Label lblFormat = new Label { Text = "出力:", AutoSize = true, Dock = DockStyle.Left, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(12, 6, 0, 0), Width = 46 };
            _cmbFormat = new ComboBox { Dock = DockStyle.Left, Width = 70, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbFormat.Items.AddRange(new object[] { "png", "pdf" });
            _cmbFormat.SelectedIndex = 0;

            _chkLinked = new CheckBox { Text = "自動更新", Dock = DockStyle.Left, AutoSize = true, Padding = new Padding(12, 6, 0, 0), Checked = true };
            _chkSaveFile = new CheckBox { Text = "ファイル保存", Dock = DockStyle.Left, AutoSize = true, Padding = new Padding(12, 6, 0, 0), Checked = true };

            top.Controls.Add(_chkSaveFile);
            top.Controls.Add(_chkLinked);
            top.Controls.Add(_cmbFormat);
            top.Controls.Add(lblFormat);
            top.Controls.Add(btnSavePreset);
            top.Controls.Add(btnLoadPreset);
            top.Controls.Add(_cmbPreset);
            top.Controls.Add(lblPreset);

            // 中央: 左にエディタ、右にプレビュー
            SplitContainer split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 520
            };

            _txtScript = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F),
                AcceptsTab = true,
                WordWrap = false,
                DetectUrls = false
            };

            Panel rightPanel = new Panel { Dock = DockStyle.Fill };
            _preview = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.White };

            Panel insertPanel = new Panel { Dock = DockStyle.Top, Height = 36, Padding = new Padding(4) };
            Button btnInsertColumn = new Button { Text = "列データ挿入", Dock = DockStyle.Left, Width = 100 };
            Button btnInsertMatrix = new Button { Text = "マトリクス挿入", Dock = DockStyle.Left, Width = 110 };
            Button btnPreview = new Button { Text = "プレビュー", Dock = DockStyle.Right, Width = 90 };
            insertPanel.Controls.Add(btnPreview);
            insertPanel.Controls.Add(btnInsertMatrix);
            insertPanel.Controls.Add(btnInsertColumn);

            rightPanel.Controls.Add(_preview);
            rightPanel.Controls.Add(insertPanel);

            split.Panel1.Controls.Add(_txtScript);
            split.Panel2.Controls.Add(rightPanel);

            // 下部: ステータスとボタン
            Panel bottom = new Panel { Dock = DockStyle.Bottom, Height = 44 };
            _lblStatus = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(8, 0, 0, 0), ForeColor = Color.DarkRed };

            Panel buttons = new Panel { Dock = DockStyle.Right, Width = 200 };
            Button btnOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Dock = DockStyle.Right, Width = 95 };
            Button btnCancel = new Button { Text = "キャンセル", DialogResult = DialogResult.Cancel, Dock = DockStyle.Right, Width = 95 };
            buttons.Controls.Add(btnOk);
            buttons.Controls.Add(btnCancel);

            bottom.Controls.Add(_lblStatus);
            bottom.Controls.Add(buttons);

            Controls.Add(split);
            Controls.Add(top);
            Controls.Add(bottom);

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            // イベント
            btnPreview.Click += (s, e) => RunPreview();
            btnInsertColumn.Click += (s, e) => InsertSelection(false);
            btnInsertMatrix.Click += (s, e) => InsertSelection(true);
            btnLoadPreset.Click += (s, e) => LoadSelectedPreset();
            btnSavePreset.Click += (s, e) => SaveAsPreset();
            FormClosing += EditorForm_FormClosing;

            Load += (s, e) =>
            {
                LoadPresetList();
                InitializeContent();
            };
        }

        /// <summary>外部から式を挿入します（リボンのデータ挿入ボタン用）。</summary>
        public void InsertExpression(string expr)
        {
            _insertedExpression = expr;
        }

        private void InitializeContent()
        {
            if (_editingShape != null)
            {
                string script = _editingShape.AlternativeText;
                PlotMetadata meta = PlotMetadata.Parse(script);
                _cmbFormat.SelectedItem = meta.FormatKey == "pdf" ? "pdf" : "png";
                if (_cmbFormat.SelectedIndex < 0) _cmbFormat.SelectedIndex = 0;
                _chkLinked.Checked = meta.Linked;
                _chkSaveFile.Checked = meta.SaveFile;
                _txtScript.Text = PlotMetadata.StripHeader(script);
            }
            else if (!string.IsNullOrEmpty(_insertedExpression))
            {
                _txtScript.Text = DefaultTemplate(_insertedExpression);
                _insertedExpression = null;
            }
            else
            {
                _txtScript.Text = DefaultTemplate(null);
            }
        }

        private static string DefaultTemplate(string dataExpr)
        {
            if (string.IsNullOrEmpty(dataExpr))
            {
                return
                    "set title \"Gnuplot\"" + Environment.NewLine +
                    "set xlabel \"X\"" + Environment.NewLine +
                    "set ylabel \"Y\"" + Environment.NewLine +
                    "plot sin(x) title \"sin(x)\" with lines";
            }
            return
                "set title \"Gnuplot\"" + Environment.NewLine +
                "set xlabel \"X\"" + Environment.NewLine +
                "set ylabel \"Y\"" + Environment.NewLine +
                "plot " + dataExpr + " using 1:2 title \"data\" with linespoints";
        }

        private void EditorForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                if (string.IsNullOrWhiteSpace(_txtScript.Text))
                {
                    MessageBox.Show("スクリプトが空です。", "GnuplotForExcel",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                }
            }
        }

        // ----- プレビュー -----

        private void RunPreview()
        {
            _lblStatus.ForeColor = Color.DarkRed;
            _lblStatus.Text = "プレビュー生成中...";
            Application.DoEvents();

            using (TempFileManager tempFiles = new TempFileManager())
            {
                ScriptProcessor processor = new ScriptProcessor();
                string error;
                string expanded = processor.Expand(_app, _txtScript.Text, tempFiles, out error);
                if (expanded == null)
                {
                    _lblStatus.Text = error;
                    ErrorLog.Append("Preview", error);
                    return;
                }

                string imagePath = tempFiles.GetTempFile("gppreview", "png");
                string script =
                    "set terminal pngcairo " + Settings.Default.PngTerminalOptions + " size 640,420" + Environment.NewLine +
                    "set output '" + imagePath.Replace('\\', '/') + "'" + Environment.NewLine +
                    expanded;

                GnuplotRunner runner = new GnuplotRunner(Settings.Default.GnuplotExePath);
                string scriptPath;
                GnuplotResult result = runner.RunScript(script, tempFiles, "gppreview",
                    out scriptPath);

                if (result.HasError)
                {
                    string msg = result.BuildErrorMessage(scriptPath);
                    _lblStatus.Text = msg;
                    ErrorLog.Append("Preview", result, scriptPath);
                    return;
                }

                try
                {
                    // 既存イメージを破棄してから読み込み
                    if (_preview.Image != null)
                    {
                        _preview.Image.Dispose();
                        _preview.Image = null;
                    }
                    // ファイルロックを避けるためストリーム経由でコピー
                    using (FileStream fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                    {
                        _preview.Image = Image.FromStream(fs);
                    }
                    _lblStatus.ForeColor = Color.DarkGreen;
                    _lblStatus.Text = "プレビュー更新: " + DateTime.Now.ToString("HH:mm:ss");
                }
                catch (Exception ex)
                {
                    _lblStatus.Text = "画像の読み込みに失敗: " + ex.Message;
                }
            }
        }

        // ----- データ挿入 -----

        private void InsertSelection(bool matrix)
        {
            string spec = GetSelectionSpec();
            if (spec == null)
            {
                MessageBox.Show("セル範囲を選択してください。", "GnuplotForExcel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string expr = matrix
                ? string.Format("<<matrix \"{0}\">>", spec)
                : string.Format("<<\"{0}\">>", spec);
            _txtScript.SelectedText = expr;
            _txtScript.Focus();
        }

        private string GetSelectionSpec()
        {
            try
            {
                Excel.Range sel = _app.Selection as Excel.Range;
                if (sel == null)
                {
                    return null;
                }
                string address = sel.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);
                Excel.Worksheet ws = (Excel.Worksheet)sel.Worksheet;
                return ws.Name + "!" + address;
            }
            catch
            {
                return null;
            }
        }

        // ----- プリセット -----

        private void LoadPresetList()
        {
            _cmbPreset.Items.Clear();
            foreach (string name in Settings.Default.GetPresetNames())
            {
                _cmbPreset.Items.Add(name);
            }
            if (_cmbPreset.Items.Count > 0)
            {
                _cmbPreset.SelectedIndex = 0;
            }
        }

        private void LoadSelectedPreset()
        {
            string name = _cmbPreset.SelectedItem as string;
            if (string.IsNullOrEmpty(name))
            {
                return;
            }
            string script = Settings.Default.LoadPreset(name);
            if (script != null)
            {
                _txtScript.Text = script;
            }
        }

        private void SaveAsPreset()
        {
            string name = PromptForPresetName();
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }
            Settings.Default.SavePreset(name.Trim(), _txtScript.Text);
            LoadPresetList();
            _cmbPreset.SelectedItem = name.Trim();
        }

        private static string PromptForPresetName()
        {
            using (Form prompt = new Form())
            {
                prompt.Text = "プリセット名";
                prompt.StartPosition = FormStartPosition.CenterParent;
                prompt.Width = 360;
                prompt.Height = 140;
                prompt.FormBorderStyle = FormBorderStyle.FixedDialog;
                prompt.MaximizeBox = false;
                prompt.MinimizeBox = false;
                prompt.ShowIcon = false;

                TextBox input = new TextBox { Dock = DockStyle.Top, Margin = new Padding(8) };
                Button ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Dock = DockStyle.Bottom, Height = 32 };
                prompt.Controls.Add(input);
                prompt.Controls.Add(ok);
                prompt.AcceptButton = ok;

                return prompt.ShowDialog() == DialogResult.OK ? input.Text : null;
            }
        }
    }
}
