using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using GnuplotForExcel.Core;
using GnuplotForExcel.Forms;

namespace GnuplotForExcel
{
    /// <summary>
    /// Gnuplot アドインのリボン (XML)。
    /// </summary>
    [ComVisible(true)]
    public class GnuplotRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("GnuplotForExcel.Ribbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        // ----- グラフ -----

        public void OnNewPlot(Office.IRibbonControl control)
        {
            using (EditorForm form = new EditorForm(Globals.ThisAddIn.Application, null))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    CreatePlot(form.ScriptBody, form.Metadata);
                }
            }
        }

        public void OnEditPlot(Office.IRibbonControl control)
        {
            Excel.Shape shape = GetSelectedPlot();
            if (shape == null)
            {
                MessageBox.Show("Gnuplot グラフを選択してください。", "GnuplotForExcel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (EditorForm form = new EditorForm(Globals.ThisAddIn.Application, shape))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    string script = form.Metadata.ToHeader() + Environment.NewLine + form.ScriptBody;
                    shape.AlternativeText = script;
                    RefreshShape(shape, true);
                }
            }
        }

        public void OnRedraw(Office.IRibbonControl control)
        {
            Excel.Shape shape = GetSelectedPlot();
            if (shape == null)
            {
                MessageBox.Show("Gnuplot グラフを選択してください。", "GnuplotForExcel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            RefreshShape(shape, true);
        }

        // ----- データ挿入 -----

        public void OnInsertColumnData(Office.IRibbonControl control)
        {
            InsertDataReference(false);
        }

        public void OnInsertMatrixData(Office.IRibbonControl control)
        {
            InsertDataReference(true);
        }

        private void InsertDataReference(bool matrix)
        {
            string spec = GetSelectionSpec();
            if (spec == null)
            {
                MessageBox.Show("セル範囲を選択してください。", "GnuplotForExcel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (EditorForm form = new EditorForm(Globals.ThisAddIn.Application, null))
            {
                string expr = matrix
                    ? string.Format("<<matrix \"{0}\">>", spec)
                    : string.Format("<<\"{0}\">>", spec);
                form.InsertExpression(expr);
                if (form.ShowDialog() == DialogResult.OK)
                {
                    CreatePlot(form.ScriptBody, form.Metadata);
                }
            }
        }

        private static string GetSelectionSpec()
        {
            try
            {
                Excel.Range sel = Globals.ThisAddIn.Application.Selection as Excel.Range;
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

        public void OnManagePresets(Office.IRibbonControl control)
        {
            using (PresetForm form = new PresetForm())
            {
                form.ShowDialog();
            }
        }

        // ----- ツール -----

        public void OnShowErrorLog(Office.IRibbonControl control)
        {
            using (ErrorLogForm form = new ErrorLogForm())
            {
                form.ShowDialog();
            }
        }

        public void OnToggleAutoRefresh(Office.IRibbonControl control, bool pressed)
        {
            Settings.Default.AutoRefreshEnabled = pressed;
            Settings.Default.Save();
            Globals.ThisAddIn.SetAutoRefresh(pressed);
        }

        public bool OnGetAutoRefreshPressed(Office.IRibbonControl control)
        {
            return Settings.Default.AutoRefreshEnabled;
        }

        public void OnRefreshAll(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.RefreshAll();
        }

        // ----- 設定 -----

        public void OnSetup(Office.IRibbonControl control)
        {
            using (SetupForm form = new SetupForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    Settings.Default.Save();
                    if (_ribbon != null)
                    {
                        _ribbon.InvalidateControl("tglAutoRefresh");
                    }
                }
            }
        }

        // ----- 内部 -----

        private static Excel.Shape GetSelectedPlot()
        {
            PlotManager pm = new PlotManager(Globals.ThisAddIn.Application);
            return pm.GetSelectedPlotShape();
        }

        private void CreatePlot(string scriptBody, PlotMetadata meta)
        {
            PlotManager pm = new PlotManager(Globals.ThisAddIn.Application);
            PlotOutcome outcome = pm.CreatePlot(scriptBody, meta, null);
            if (!outcome.Success)
            {
                ErrorLog.Append("CreatePlot", outcome.ErrorMessage);
                MessageBox.Show(outcome.ErrorMessage, "GnuplotForExcel",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshShape(Excel.Shape shape, bool showError)
        {
            PlotManager pm = new PlotManager(Globals.ThisAddIn.Application);
            PlotOutcome outcome = pm.RefreshShape(shape);
            if (!outcome.Success)
            {
                ErrorLog.Append("RefreshShape", outcome.ErrorMessage);
                if (showError)
                {
                    MessageBox.Show(outcome.ErrorMessage, "GnuplotForExcel",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    return null;
                }
                using (StreamReader reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
