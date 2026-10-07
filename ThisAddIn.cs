using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace GnuplotForExcel
{
    public partial class ThisAddIn
    {
        private Core.LinkManager _linkManager;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _linkManager = new Core.LinkManager(Application);
            _linkManager.EnableAutoRefresh(Core.Settings.Default.AutoRefreshEnabled);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (_linkManager != null)
            {
                _linkManager.Dispose();
                _linkManager = null;
            }
        }

        /// <summary>
        /// Ribbon (XML) によるリボン拡張を返します。
        /// </summary>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new GnuplotRibbon();
        }

        /// <summary>
        /// セル変更に連動したグラフの自動更新を有効/無効にします。
        /// </summary>
        internal void SetAutoRefresh(bool enabled)
        {
            if (_linkManager != null)
            {
                _linkManager.EnableAutoRefresh(enabled);
            }
        }

        /// <summary>
        /// シート上のすべての Gnuplot グラフを再描画します。
        /// </summary>
        internal void RefreshAll()
        {
            if (_linkManager != null)
            {
                _linkManager.RefreshAll();
            }
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
