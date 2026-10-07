using System;
using System.Drawing;
using System.Windows.Forms;
using GnuplotForExcel.Core;

namespace GnuplotForExcel.Forms
{
    /// <summary>
    /// Gnuplot のエラーログを表示するダイアログ（デバッグ機能）。
    /// </summary>
    public class ErrorLogForm : Form
    {
        private readonly TextBox _txtLog;
        private readonly Button _btnClose;
        private readonly Button _btnClear;
        private readonly Label _lblTitle;

        public ErrorLogForm()
        {
            Text = "Gnuplot エラーログ";
            StartPosition = FormStartPosition.CenterParent;
            Width = 720;
            Height = 480;
            MinimizeBox = false;
            MaximizeBox = true;
            FormBorderStyle = FormBorderStyle.Sizable;
            ShowIcon = false;
            ShowInTaskbar = false;
            Font = SystemFonts.MessageBoxFont;

            _lblTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 28,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                Text = "Gnuplot の実行結果とエラーを表示します。"
            };

            _txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                ReadOnly = true,
                WordWrap = false,
                Font = new Font("Consolas", 9F),
                BackColor = Color.White
            };

            Panel bottom = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            _btnClear = new Button { Text = "クリア", Width = 90, Dock = DockStyle.Left };
            _btnClose = new Button { Text = "閉じる", Width = 90, Dock = DockStyle.Right, DialogResult = DialogResult.OK };

            _btnClear.Click += (s, e) => { _txtLog.Clear(); ErrorLog.Clear(); };

            bottom.Controls.Add(_btnClear);
            bottom.Controls.Add(_btnClose);

            Controls.Add(_txtLog);
            Controls.Add(_lblTitle);
            Controls.Add(bottom);

            AcceptButton = _btnClose;
            CancelButton = _btnClose;

            Load += (s, e) => { _txtLog.Text = ErrorLog.GetAll(); };
        }
    }
}
