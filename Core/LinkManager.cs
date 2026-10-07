using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GnuplotForExcel.Core
{
    /// <summary>
    /// セルデータの変更に連動したグラフの自動更新（動的リンク）を管理します。
    /// </summary>
    /// <remarks>
    /// SheetChange イベントで変更範囲を取得し、シート上の Gnuplot シェイプが参照する
    /// 範囲と交差する場合に、デバウンス（約 1 秒）して再描画します。
    /// 再描画中の再入は抑止します。
    /// </remarks>
    public sealed class LinkManager : IDisposable
    {
        private const int DebounceMs = 1000;

        private readonly Excel.Application _app;
        private readonly PlotManager _plotManager;
        private readonly ScriptProcessor _scriptProcessor;
        private readonly Timer _debounceTimer;
        private readonly HashSet<Excel.Shape> _pendingShapes = new HashSet<Excel.Shape>();

        private bool _enabled;
        private bool _isRefreshing;

        public LinkManager(Excel.Application app)
        {
            _app = app;
            _plotManager = new PlotManager(app);
            _scriptProcessor = new ScriptProcessor();
            _debounceTimer = new Timer { Interval = DebounceMs };
            _debounceTimer.Tick += DebounceTimer_Tick;
        }

        /// <summary>自動更新の有効/無効を切り替えます。</summary>
        public void EnableAutoRefresh(bool enabled)
        {
            if (_enabled == enabled)
            {
                return;
            }
            _enabled = enabled;
            if (enabled)
            {
                _app.SheetChange += App_SheetChange;
            }
            else
            {
                _app.SheetChange -= App_SheetChange;
                StopDebounce();
            }
        }

        /// <summary>現在のブック内の全 Gnuplot グラフを再描画します。</summary>
        public void RefreshAll()
        {
            if (_isRefreshing)
            {
                return;
            }
            _isRefreshing = true;
            try
            {
                foreach (Excel.Worksheet sheet in _app.Worksheets)
                {
                    foreach (Excel.Shape shape in _plotManager.GetPlotShapes(sheet))
                    {
                        _plotManager.RefreshShape(shape);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("RefreshAll failed: " + ex.Message);
            }
            finally
            {
                _isRefreshing = false;
            }
        }

        private void App_SheetChange(object sheetObject, Excel.Range target)
        {
            if (!_enabled || _isRefreshing)
            {
                return;
            }

            Excel.Worksheet sheet = sheetObject as Excel.Worksheet;
            if (sheet == null)
            {
                return;
            }

            try
            {
                // 変更範囲と交差するリンク済みシェイプを収集
                foreach (Excel.Shape shape in _plotManager.GetPlotShapes(sheet))
                {
                    string script = shape.AlternativeText;
                    PlotMetadata meta = PlotMetadata.Parse(script);
                    if (!meta.Linked)
                    {
                        continue;
                    }

                    string body = PlotMetadata.StripHeader(script);
                    List<RangeReference> refs = _scriptProcessor.ExtractReferences(body);
                    bool intersects = refs.Any(r => RangeIntersects(sheet, r.Spec, target));
                    if (intersects)
                    {
                        lock (_pendingShapes)
                        {
                            _pendingShapes.Add(shape);
                        }
                    }
                }

                if (HasPending())
                {
                    RestartDebounce();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SheetChange handling failed: " + ex.Message);
            }
        }

        /// <summary>変更範囲が参照指定と交差するか判定します。</summary>
        private bool RangeIntersects(Excel.Worksheet changedSheet, string spec, Excel.Range target)
        {
            try
            {
                string sheetName;
                string rangePart;
                if (RangeExporter.TrySplitSpec(spec, out sheetName, out rangePart))
                {
                    if (!string.Equals(changedSheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }
                }

                Excel.Range refRange = RangeExporter.ResolveRange(_app, spec);
                if (refRange == null)
                {
                    return false;
                }
                Excel.Range intersection = _app.Intersect(refRange, target);
                return intersection != null;
            }
            catch
            {
                return false;
            }
        }

        // ----- デバウンス -----

        private bool HasPending()
        {
            lock (_pendingShapes)
            {
                return _pendingShapes.Count > 0;
            }
        }

        private void RestartDebounce()
        {
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        private void StopDebounce()
        {
            _debounceTimer.Stop();
            lock (_pendingShapes)
            {
                _pendingShapes.Clear();
            }
        }

        private void DebounceTimer_Tick(object sender, EventArgs e)
        {
            _debounceTimer.Stop();
            FlushPending();
        }

        private void FlushPending()
        {
            if (_isRefreshing)
            {
                return;
            }

            List<Excel.Shape> shapes;
            lock (_pendingShapes)
            {
                shapes = _pendingShapes.ToList();
                _pendingShapes.Clear();
            }

            if (shapes.Count == 0)
            {
                return;
            }

            _isRefreshing = true;
            try
            {
                foreach (Excel.Shape shape in shapes)
                {
                    // 削除済みシェイプの再描画はスキップ
                    try
                    {
                        string name = shape.Name; // アクセスで存在確認
                        _plotManager.RefreshShape(shape);
                    }
                    catch
                    {
                        // 既に削除されている等
                    }
                }
            }
            finally
            {
                _isRefreshing = false;
            }
        }

        public void Dispose()
        {
            try
            {
                _app.SheetChange -= App_SheetChange;
            }
            catch { }
            StopDebounce();
            if (_debounceTimer != null)
            {
                _debounceTimer.Dispose();
            }
        }
    }
}
