"""
055_ja100_word2021_update_cross.py
「リンク貼り付け」されたWordレポートの図表を一括更新するスクリプト
※ リアルタイム進捗マーカー（走査・更新同期）対応版
※ doc.Fields.Update() は見えないダイアログでのデッドロックを避けるためスキップ（Fields件数はログに記録）
※ Type値による事前フィルタをやめ、LinkFormat.Update()を実際に試して成否で判定する方式に変更
   （msoChart型でもOLEリンクを持つ場合は更新対象に含める）
※ 【追加】更新時の座標（Top/Left）固定処理を追加し、Wordの仕様によるレイアウト崩れ（図の吹き飛び）を防止。
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

wdAlertsNone = 0
wdDoNotSaveChanges = 0
msoGroup = 6
msoChart = 3
xlCalculationManual = -4135
xlCalculationAutomatic = -4105


def fix_excel_charts_font(excel_path):
    """★追加：Excel側のグラフのフォントを游ゴシックに固定し、サイズを維持する。
    Word側のリンク更新（update_word_links）とは独立した、Excel側だけの処理。"""
    print("=" * 60)
    print("📊 [事前処理] Excelグラフのフォント固定処理を開始します...")
    print("=" * 60)

    if not os.path.exists(excel_path):
        print(f"❌ エラー: Excelファイルが見つかりません: {excel_path}")
        return False

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.ScreenUpdating = False

    wb = None
    try:
        wb = excel.Workbooks.Open(excel_path)
        # ※ファイルを開いた後に計算モードを手動にして高速化（エラー回避策）
        excel.Calculation = xlCalculationManual
        chart_count = 0

        for ws in wb.Worksheets:
            for cho in ws.ChartObjects():
                original_width = cho.Width
                original_height = cho.Height

                cho.Chart.Refresh()

                chart_font = cho.Chart.ChartArea.Font
                chart_font.Name = "Yu Gothic"
                chart_font.Bold = False
                chart_font.FontStyle = "標準"

                try:
                    tf2_font = cho.Chart.ChartArea.Format.TextFrame2.TextRange.Font
                    tf2_font.NameComplexScript = "Yu Gothic"
                    tf2_font.NameFarEast = "Yu Gothic"
                    tf2_font.Name = "Yu Gothic"
                except Exception:
                    pass

                try:
                    cho.Chart.ChartTitle.Font.Bold = False
                except Exception:
                    pass

                cho.Width = original_width
                cho.Height = original_height
                chart_count += 1

        wb.Save()
        print(f"✅ Excel完了: {chart_count} 個のグラフを「Yu Gothic・太字なし」に固定しました。")
        return True

    except Exception as e:
        print(f"❌ Excel処理中にエラーが発生しました: {e}")
        return False

    finally:
        try:
            excel.Calculation = xlCalculationAutomatic
            excel.ScreenUpdating = True
        except Exception:
            pass
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        time.sleep(1)

def setup_logger():
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    log_file = os.path.join(log_dir, f"055_{today_str}.log")

    logger = logging.getLogger("WordUpdater")
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

def _kill_orphan_process(pid, proc_name):
    try:
        if pid and psutil.pid_exists(pid):
            proc = psutil.Process(pid)
            if proc_name in proc.name().lower():
                proc.kill()
                logger.info(f"🧹 残存していた{proc_name}プロセス(PID={pid})を強制終了しました。")
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        pass

def _try_update_link(shape_obj):
    """LinkFormatが存在し、更新できればTrue。存在しない/失敗すればFalse
    ※更新前後でWidth/Heightに加え、Top/Left（座標位置）も強制保持し、
    Word特有の「リンク更新でグラフが吹き飛ぶ現象」を完全に防ぐ。"""
    try:
        lf = shape_obj.LinkFormat
        if lf is None:
            return False

        # 1. 更新前のサイズと【座標】を記憶
        original_width = shape_obj.Width
        original_height = shape_obj.Height
        
        original_left = None
        original_top = None
        try:
            # 浮動図表（Shapes）のみ座標を取得
            original_left = shape_obj.Left
            original_top = shape_obj.Top
        except Exception:
            pass # インライン図表（InlineShapes）には位置座標がないのでスキップ

        try:
            original_lock_aspect = shape_obj.LockAspectRatio
            shape_obj.LockAspectRatio = False  # 縦横比連動による干渉を防ぐ
        except Exception:
            original_lock_aspect = None

        # 2. リンク更新の実行
        lf.Update()

        # 3. サイズと【座標】の強制復元（Wordの暴走を上書きしてピン留め）
        shape_obj.Width = original_width
        shape_obj.Height = original_height
        
        if original_left is not None:
            try:
                shape_obj.Left = original_left
            except Exception:
                pass
        if original_top is not None:
            try:
                shape_obj.Top = original_top
            except Exception:
                pass

        if original_lock_aspect is not None:
            try:
                shape_obj.LockAspectRatio = original_lock_aspect
            except Exception:
                pass

        return True
    except Exception:
        return False

def _is_linked_chart_unrefreshed(shape_obj):
    """Chart型で、かつChartDataがIsLinked=Trueなのに上のLinkFormat.Update()が
    効かなかったものを検知するための補助判定。"""
    try:
        if shape_obj.Type == msoChart and shape_obj.HasChart:
            return bool(shape_obj.Chart.ChartData.IsLinked)
    except Exception:
        pass
    return False

def _update_inline_shapes(doc):
    try:
        total = doc.InlineShapes.Count
    except Exception:
        return 0, 0, []

    ok, fail = 0, 0
    chart_unrefreshed = []
    if total == 0:
        return ok, fail, chart_unrefreshed

    print(f"  👉 インライン図表の更新を開始します (全 {total} 個のオブジェクトを走査)")
    for i in range(1, total + 1):
        try:
            shape = doc.InlineShapes(i)
            if _try_update_link(shape):
                ok += 1
            elif _is_linked_chart_unrefreshed(shape):
                chart_unrefreshed.append(("Inline", i, getattr(shape, "Range", None) and "" or ""))
            # それ以外（テキストボックス・単なる画像等）はカウントしない
        except Exception:
            pass

        sys.stdout.write(f"\r    🔄 走査・更新中... {i}/{total} 個完了")
        sys.stdout.flush()

    sys.stdout.write("\n")
    return ok, fail, chart_unrefreshed

def _update_floating_shapes(doc):
    try:
        total = doc.Shapes.Count
    except Exception:
        return 0, 0, []

    ok, fail = 0, 0
    chart_unrefreshed = []
    if total == 0:
        return ok, fail, chart_unrefreshed

    def process_one(shape, idx_label):
        nonlocal ok, fail
        if _try_update_link(shape):
            ok += 1
        elif _is_linked_chart_unrefreshed(shape):
            try:
                name = shape.Name
            except Exception:
                name = ""
            chart_unrefreshed.append((idx_label, name))

    print(f"  👉 浮動図表の更新を開始します (全 {total} 個のオブジェクトを走査)")
    for i in range(1, total + 1):
        try:
            shape = doc.Shapes(i)
            if shape.Type == msoGroup:
                for j in range(1, shape.GroupItems.Count + 1):
                    try:
                        sub_shape = shape.GroupItems(j)
                        process_one(sub_shape, f"{i}-G{j}")
                    except Exception:
                        pass
            else:
                process_one(shape, str(i))
        except Exception:
            pass

        sys.stdout.write(f"\r    🔄 走査・更新中... {i}/{total} 個完了")
        sys.stdout.flush()

    sys.stdout.write("\n")
    return ok, fail, chart_unrefreshed

def update_word_links(excel_path, word_path):
    abs_excel = os.path.abspath(excel_path)
    abs_word = os.path.abspath(word_path)

    if not os.path.exists(abs_excel):
        logger.error(f"❌ エラー: 参照元のExcelファイルが見つかりません: {abs_excel}")
        sys.exit(1)
    if not os.path.exists(abs_word):
        logger.error(f"❌ エラー: 対象のWordファイルが見つかりません: {abs_word}")
        sys.exit(1)

    logger.info("=" * 60)
    logger.info("🔄 Wordレポートのリンク図表・自動更新を開始します")
    logger.info("=" * 60)

    if _is_file_locked(abs_excel):
        logger.warning(f"⚠️ スキップ: Excel ({os.path.basename(abs_excel)}) が開かれています。")
        sys.exit(2)
    if _is_file_locked(abs_word):
        logger.warning(f"⚠️ スキップ: Word ({os.path.basename(abs_word)}) が開かれています。")
        sys.exit(2)

    start_time = time.time()
    pythoncom.CoInitialize()

    excel, word, wb, doc = None, None, None, None
    excel_pid, word_pid = None, None
    original_update_links_setting = None

    try:
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AutomationSecurity = 3
            try:
                _, excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
            except Exception:
                pass
            logger.info("⏳ 参照元のExcelをバックグラウンドで準備しています...")
            wb = excel.Workbooks.Open(abs_excel, ReadOnly=True, UpdateLinks=False)
        except com_error as e:
            logger.error(f"❌ Excel起動失敗: {e}")
            sys.exit(1)

        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = wdAlertsNone
            word.ScreenUpdating = False
            word.AutomationSecurity = 3
            original_update_links_setting = word.Options.UpdateLinksAtOpen
            word.Options.UpdateLinksAtOpen = False
            try:
                _, word_pid = win32process.GetWindowThreadProcessId(word.Hwnd)
            except Exception:
                pass
        except com_error as e:
            logger.error(f"❌ Word起動失敗: {e}")
            sys.exit(1)

        logger.info("⏳ Wordレポートを開いています...")
        doc = word.Documents.Open(abs_word, ReadOnly=False)

        if doc.ReadOnly:
            logger.warning("⚠️ スキップ: Wordファイルが読み取り専用です。")
            sys.exit(2)

        logger.info("📊 図表データの更新を実行中... (プログレス表示)")

        sys.stdout.write("  👉 全体リンク(Fields)の更新中... (スキップします)\n")
        sys.stdout.flush()
        try:
            field_count = doc.Fields.Count
            logger.info(f"ℹ️ 文書内のFields数: {field_count}件（一括更新はスキップ。グラフ以外のリンク貼り付けがあれば要目視確認）")
        except Exception:
            pass

        inline_ok, inline_fail, inline_chart_stuck = _update_inline_shapes(doc)
        float_ok, float_fail, float_chart_stuck = _update_floating_shapes(doc)

        update_count = inline_ok + float_ok
        fail_count = inline_fail + float_fail
        chart_stuck = inline_chart_stuck + float_chart_stuck

        msg = f"✨ 計 {update_count} 個の図表リンクを最新化しました。"
        if fail_count > 0:
            msg += f"（⚠️ 更新失敗: {fail_count}件）"
        logger.info(msg)

        if chart_stuck:
            logger.warning(f"⚠️ 【要確認】リンク済みのはずのグラフ(Chart)で自動更新できなかったもの: {len(chart_stuck)}件")
            for label, name in chart_stuck:
                logger.warning(f"     - 番号:{label} 名前:{name}")
            logger.warning("   これらは各グラフを右クリック→「データの編集」を開いて手動更新するか、"
                           "個別にご相談ください。自動更新の対象外のまま保存されています。")
        else:
            logger.info("✅ グラフ(Chart)型のオブジェクトについて、更新できなかったリンク済みチャートはありませんでした。")

        logger.info("💾 Wordレポートを自身に上書き保存しています...")
        doc.Save()

        elapsed = time.time() - start_time
        logger.info(f"✅ Word出力が完了しました！ (所要時間: {elapsed:.1f}秒)")

    except com_error as e:
        logger.error(f"❌ COMエラーが発生しました 詳細: {e}")
        sys.exit(1)
    except Exception as e:
        logger.exception(f"❌ 予期しないエラーが発生しました: {e}")
        sys.exit(1)
    finally:
        try:
            if doc is not None: doc.Close(SaveChanges=wdDoNotSaveChanges)
        except Exception: pass
        try:
            if word is not None:
                try: word.NormalTemplate.Saved = True
                except Exception: pass
                if original_update_links_setting is not None:
                    word.Options.UpdateLinksAtOpen = original_update_links_setting
                word.Quit()
        except Exception: pass
        try:
            if wb is not None: wb.Close(SaveChanges=False)
        except Exception: pass
        try:
            if excel is not None: excel.Quit()
        except Exception: pass

        doc, wb, word, excel = None, None, None, None
        gc.collect()
        pythoncom.CoUninitialize()

        for _ in range(5):
            word_alive = word_pid and psutil.pid_exists(word_pid)
            excel_alive = excel_pid and psutil.pid_exists(excel_pid)
            if not word_alive and not excel_alive:
                break
            time.sleep(1)

        _kill_orphan_process(word_pid, "winword")
        _kill_orphan_process(excel_pid, "excel")


if __name__ == '__main__':
    excel_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'053_batch_analysis_master_cross_shisho_all.xlsx')
    word_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'055_ja100_word2021_update_cross.docx')

    # ① まずExcel側のグラフフォントを固定
    excel_ok = fix_excel_charts_font(excel_file)

    # ② 続けてWordのリンク図表を更新（座標ズレ防止追加版）
    if excel_ok:
        update_word_links(excel_file, word_file)
        print("\n🎉 全ての自動更新パイプラインが正常に完了しました！")
    else:
        print("\n⚠️ Excelのフォント固定処理に失敗したため、Wordの更新を中止しました。")