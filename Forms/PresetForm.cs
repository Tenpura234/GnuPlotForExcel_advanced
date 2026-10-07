using System;
using System.Drawing;
using System.Windows.Forms;
using GnuplotForExcel.Core;

namespace GnuplotForExcel.Forms
{
    /// <summary>
    /// 書式プリセットの保存・適用・削除を行う管理ダイアログ。
    /// </summary>
    public class PresetForm : Form
    {
        private readonly ListBox _list;
        private readonly TextBox _txtScript;
        private readonly TextBox _txtName;

        public PresetForm()
        {
            Text = "書式プリセット管理";
            StartPosition = FormStartPosition.CenterParent;
            Width = 760;
            Height = 520;
            FormBorderStyle = FormBorderStyle.Sizable;
            ShowIcon = false;
            Font = SystemFonts.MessageBoxFont;

            SplitContainer split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 220
            };

            _list = new ListBox { Dock = DockStyle.Fill };
            split.Panel1.Controls.Add(_list);

            Panel right = new Panel { Dock = DockStyle.Fill, Padding = new Padding(6) };

            Panel namePanel = new Panel { Dock = DockStyle.Top, Height = 30 };
            Label lblName = new Label { Text = "名前:", Dock = DockStyle.Left, Width = 44, TextAlign = ContentAlignment.MiddleLeft };
            _txtName = new TextBox { Dock = DockStyle.Fill };
            namePanel.Controls.Add(_txtName);
            namePanel.Controls.Add(lblName);

            _txtScript = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Font = new Font("Consolas", 9.5F),
                AcceptsReturn = true,
                AcceptsTab = true
            };

            Panel btnPanel = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            Button btnNew = new Button { Text = "新規", Width = 80, Dock = DockStyle.Left };
            Button btnSave = new Button { Text = "保存", Width = 80, Dock = DockStyle.Left };
            Button btnDelete = new Button { Text = "削除", Width = 80, Dock = DockStyle.Left };
            Button btnClose = new Button { Text = "閉じる", Width = 80, Dock = DockStyle.Right, DialogResult = DialogResult.OK };

            btnNew.Click += (s, e) => NewPreset();
            btnSave.Click += (s, e) => SaveCurrent();
            btnDelete.Click += (s, e) => DeleteCurrent();

            btnPanel.Controls.Add(btnDelete);
            btnPanel.Controls.Add(btnSave);
            btnPanel.Controls.Add(btnNew);
            btnPanel.Controls.Add(btnClose);

            right.Controls.Add(_txtScript);
            right.Controls.Add(namePanel);
            right.Controls.Add(btnPanel);

            split.Panel2.Controls.Add(right);

            Controls.Add(split);
            AcceptButton = btnClose;
            CancelButton = btnClose;

            _list.SelectedIndexChanged += (s, e) => LoadSelected();
            Load += (s, e) => RefreshList(null);
        }

        private void RefreshList(string selectName)
        {
            _list.Items.Clear();
            foreach (string name in Settings.Default.GetPresetNames())
            {
                _list.Items.Add(name);
            }
            if (selectName != null)
            {
                int idx = _list.Items.IndexOf(selectName);
                if (idx >= 0) _list.SelectedIndex = idx;
            }
            else if (_list.Items.Count > 0)
            {
                _list.SelectedIndex = 0;
            }
        }

        private void LoadSelected()
        {
            string name = _list.SelectedItem as string;
            if (name == null)
            {
                _txtName.Clear();
                _txtScript.Clear();
                return;
            }
            _txtName.Text = name;
            _txtScript.Text = Settings.Default.LoadPreset(name) ?? string.Empty;
        }

        private void NewPreset()
        {
            _list.ClearSelected();
            _txtName.Text = "新しいプリセット";
            _txtScript.Text = "set title \"\"\r\nset grid\r\n";
            _txtName.Focus();
            _txtName.SelectAll();
        }

        private void SaveCurrent()
        {
            string name = (_txtName.Text ?? string.Empty).Trim();
            if (name.Length == 0)
            {
                MessageBox.Show("プリセット名を入力してください。", "GnuplotForExcel",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Settings.Default.SavePreset(name, _txtScript.Text);
            RefreshList(name);
        }

        private void DeleteCurrent()
        {
            string name = _list.SelectedItem as string;
            if (name == null)
            {
                return;
            }
            if (MessageBox.Show("プリセット \"" + name + "\" を削除しますか？", "GnuplotForExcel",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Settings.Default.DeletePreset(name);
                RefreshList(null);
            }
        }
    }
}
