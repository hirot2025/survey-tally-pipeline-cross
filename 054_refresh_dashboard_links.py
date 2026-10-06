"""
054_refresh_dashboard_links.py
ダッシュボード用Excelの外部リンクを自動更新し、さらに
【全グラフのフォントを游ゴシック・8ptに完全固定して】自身に上書き保存するスクリプト。
（所内共有用の完成品053.xlsxを生成するための最終処理）
"""

import gc
import os
import sys
import time
import datetime
import logging

import psutil
import pythoncom
import win32com.client
import win32process
from pywintypes import com_error

def setup_logger():
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    log_file = os.path.join(log_dir, f"054_refresh_{today_str}.log")
    
    logger = logging.getLogger("DashboardRefresher")
    logger.setLevel(logging.INFO)
    
    if not logger.handlers:
        fh = logging.FileHandler(log_file, encoding='utf-8')
        sh = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        fh.setFormatter(formatter)
        sh.setFormatter(formatter)
        logger.addHandler(fh)
        logger.addHandler(sh)
    return logger

logger = setup_logger()

def _lock_file_path(abs_path):
    dir_name = os.path.dirname(abs_path)
    base_name = os.path.basename(abs_path)
    return os.path.join(dir_name, f"~${base_name}")

def _is_file_locked(abs_path):
    return os.path.exists(_lock_file_path(abs_path))

def _kill_orphan_excel(pid):
    try:
        if pid and psutil.pid_exists(pid):
            proc = psutil.Process(pid)
            if "excel" in proc.name().lower():
                proc.kill()
                logger.info(f"🧹 残存していたExcelプロセス(PID={pid})を強制終了しました。")
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        pass

def _force_chart_font(chart, font_name="游ゴシック", target_size=8):
    """【追加】グラフの全要素に対して、フォントとサイズを徹底的に強制設定する"""
    # 1. 全体
    try:
        chart.ChartArea.Font.Name = font_name
        chart.ChartArea.Font.Size = target_size
        chart.ChartArea.Font.Bold = False
        try:
            tf2 = chart.ChartArea.Format.TextFrame2.TextRange.Font
            tf2.NameComplexScript = font_name
            tf2.NameFarEast = font_name
            tf2.Name = font_name
            tf2.Size = target_size
        except Exception: pass
    except Exception: pass

    # 2. タイトル
    if chart.HasTitle:
        try:
            chart.ChartTitle.Font.Name = font_name
            chart.ChartTitle.Font.Size = target_size
            chart.ChartTitle.Font.Bold = False
        except Exception: pass

    # 3. 軸
    try:
        for axis_type in [1, 2, 3]:
            for axis_group in [1, 2]:
                try:
                    axis = chart.Axes(axis_type, axis_group)
                    axis.TickLabels.Font.Name = font_name
                    axis.TickLabels.Font.Size = target_size
                    axis.TickLabels.Font.Bold = False
                    if axis.HasTitle:
                        axis.AxisTitle.Font.Name = font_name
                        axis.AxisTitle.Font.Size = target_size
                        axis.AxisTitle.Font.Bold = False
                except Exception: pass
    except Exception: pass

    # 4. 凡例
    if chart.HasLegend:
        try:
            chart.Legend.Font.Name = font_name
            chart.Legend.Font.Size = target_size
            chart.Legend.Font.Bold = False
        except Exception: pass

    # 5. データラベル
    try:
        for series in chart.SeriesCollection():
            if series.HasDataLabels:
                try:
                    series.DataLabels().Font.Name = font_name
                    series.DataLabels().Font.Size = target_size
                    series.DataLabels().Font.Bold = False
                except Exception: pass
    except Exception: pass


def refresh_excel_links(target_path):
    abs_target_path = os.path.abspath(target_path)

    if not os.path.exists(abs_target_path):
        logger.error(f"❌ エラー: 対象のExcelファイルが見つかりません。 パス: {abs_target_path}")
        sys.exit(1)

    logger.info("=" * 60)
    logger.info("🔄 ダッシュボードの値自動更新を開始します")
    logger.info(f"対象ファイル: {os.path.basename(abs_target_path)}")
    logger.info("=" * 60)

    if _is_file_locked(abs_target_path):
        logger.warning(f"⚠️ スキップ: 対象ファイル ({os.path.basename(abs_target_path)}) が現在開かれています。閉じてから再実行してください。")
        sys.exit(2)

    start_time = time.time()
    pythoncom.CoInitialize()

    excel = None
    excel_pid = None
    wb = None

    try:
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except com_error as e:
            logger.error(f"❌ エラー: Excelの起動に失敗しました。詳細: {e}")
            sys.exit(1)

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        
        # ※ファイルを開く前は設定できないため、ここで手動計算モードにするのはやめます
        excel.AskToUpdateLinks = False       
        excel.ScreenUpdating = False         
        excel.EnableEvents = False           
        excel.Interactive = False            
        excel.AlertBeforeOverwriting = False

        try:
            _, excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        except Exception:
            excel_pid = None

        logger.info("⏳ Excelファイルを読み込み、外部データを取得しています...")
        
        wb = excel.Workbooks.Open(
            abs_target_path,
            UpdateLinks=3,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
        )

        if wb.ReadOnly:
            logger.warning("⚠️ スキップ: 対象ファイルが読み取り専用で開かれました。保存せず終了します。")
            wb.Close(SaveChanges=False)
            wb = None
            sys.exit(2)

        # ここで再計算を実行（データ更新）
        excel.Calculate()

        # =========================================================================
        # ★追加：データ更新が終わった直後に、全グラフのフォントとサイズを強制的に8ptで固定する
        # =========================================================================
        logger.info("🎨 共有に向けて、グラフの全フォントを「游ゴシック・8pt」に整形しています...")
        chart_count = 0
        for ws in wb.Worksheets:
            for cho in ws.ChartObjects():
                orig_w = cho.Width
                orig_h = cho.Height
                
                # キャッシュ更新
                cho.Chart.Refresh()
                
                # 隅々まで8pt・游ゴシックに強制上書き
                _force_chart_font(cho.Chart, "游ゴシック", 8)
                
                # Excelグラフ自体のサイズ崩れを防ぐ
                cho.Width = orig_w
                cho.Height = orig_h
                chart_count += 1
        logger.info(f"   -> {chart_count} 個のグラフの書式固定が完了しました。")
        # =========================================================================

        logger.info(f"💾 {os.path.basename(abs_target_path)} に上書き保存しています...")
        wb.Save()
        wb.Close(SaveChanges=False)  
        wb = None

        elapsed = time.time() - start_time
        logger.info(f"✅ 更新と書式固定、上書き保存が完了しました！ (所要時間: {elapsed:.1f}秒)")

    except com_error as e:
        elapsed = time.time() - start_time
        logger.error(f"❌ COMエラーが発生しました（所要時間: {elapsed:.1f}秒） 詳細: {e}")
        sys.exit(1)
    except Exception as e:
        elapsed = time.time() - start_time
        logger.exception(f"❌ 予期しないエラーが発生しました（所要時間: {elapsed:.1f}秒）")
        sys.exit(1)
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception: pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception: pass

        wb = None
        del excel
        gc.collect()
        pythoncom.CoUninitialize()

        for _ in range(5):
            if not (excel_pid and psutil.pid_exists(excel_pid)):
                break
            time.sleep(1)
        _kill_orphan_excel(excel_pid)

if __name__ == '__main__':
    target_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'053_batch_analysis_master_cross_shisho_all.xlsx')
    refresh_excel_links(target_file)