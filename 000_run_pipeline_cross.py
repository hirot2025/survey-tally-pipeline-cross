# -*- coding: utf-8 -*-
"""
000_run_pipeline_cross.py
クロス集計パイプライン 一気通貫バッチ（データ整形 → 分析 → マスター出力 → Excel/Word更新）

使い方（PowerShell）:
  python 000_run_pipeline_cross.py                 … 最初から最後まで
  python 000_run_pipeline_cross.py --list          … 実行順の一覧を表示して終了
  python 000_run_pipeline_cross.py 8               … 一覧の 8 番から再開
  python 000_run_pipeline_cross.py 060_pca_ward_grouping.py   … そのスクリプトから再開（最初に出てくる位置）
  python 000_run_pipeline_cross.py --until 14      … 14 番まで実行して終了（Excel/Word 更新を行わない等）

旧版からの主な変更（2026-09-29）:
  1. 分析スクリプト（060〜086）を追加し、入力待ちのあるスクリプトには答えを自動で渡す
     （060: k=2 / 062: クラスター番号 1 と 2 / 064: 1:3 2:5）
  2. タイムアウトと Office 後始末の対象を実際のファイル名（054/055）に合わせた
     ※ 旧版は 052/053 のままだったため、055 は既定の60秒で打ち切られ、Wordが強制終了されていた
  3. Word を使う前に、残っている WINWORD.EXE / EXCEL.EXE がないか確認する
     （強制終了されたWordが残ると、次回「重大なエラーが発生しました」の確認が非表示で出て止まるため）
  4. 実行順の番号で再開・途中終了ができる。最後に各ステップの所要時間を一覧表示する
  5. 年鑑の S層・サブクラスタは 089_yearbook_layers.py（方法a）を 075・083・086・087 が共通で読む
     （089 は読み込まれるだけなので実行順には入れていないが、作業フォルダに置く必要がある）
"""
import io
import os
import subprocess
import sys
import time

if sys.platform == "win32":
    if sys.stdout is not None:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True)
    if sys.stderr is not None:
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace", line_buffering=True)

# ============================================================================
# 実行順（上から順に実行）
#   script : スクリプト名
#   args   : コマンドライン引数
#   stdin  : 入力待ちに自動で渡す文字列（改行区切り）。"" は「入力なし」（誤って入力待ちになっても止まらない）
#            指定なしのステップは、これまでどおりコンソールの入力を引き継ぐ
#   timeout: 秒
#   office : タイムアウト・異常終了時に後始末する Office プロセス
#   confirm: 実行前に人の確認を求める（Word操作の前など）
#   note   : 一覧表示用の説明
# ============================================================================
STEPS = [
    {"script": "010_generate_csvs_from_rawxlsx_cross.py", "timeout": 180, "note": "元のxlsx → CSV"},
    {"script": "011_format_csv_data_cross.py",          "timeout": 60,  "note": "CSV整形"},
    {"script": "020_step1_merge_cross.py",              "timeout": 60,  "note": "結合"},
    {"script": "050_horizon_conv_cross_100ja.py",       "timeout": 60,  "note": "横持ち100JA"},
    {"script": "051_horizon_conv_cross_100ja_cleaned.py", "timeout": 120, "note": "クリーニング（True/False→1/0 等）"},
    {"script": "052_master_batch_processor_cross.py", "timeout": 900, "note": "クロス集計(設定の CSV ごとに xlsx)"},
]


def running_office(names):
    """起動中の Office プロセス名の一覧（Windows 以外や tasklist がない環境では空）"""
    found = []
    for n in names:
        try:
            r = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {n}"], capture_output=True, text=True)
            if n.lower() in r.stdout.lower():
                found.append(n)
        except FileNotFoundError:
            return []
    return found


def kill_office_safely(targets):
    print("\n⚠️ 停止・タイムアウトを検知しました。対象の Office プロセスを終了します...")
    for t in targets:
        try:
            r = subprocess.run(["taskkill", "/IM", t, "/F"], capture_output=True, text=True)
            if r.returncode == 0:
                print(f"  ✅ {t} を強制終了しました。")
            else:
                print(f"  ✅ {t} は起動していません。")
        except FileNotFoundError:
            pass
    if "WINWORD.EXE" in targets:
        print("  ⚠️ Word を強制終了した場合、次回の起動時に「前回開いた時に重大なエラーが発生しました」と確認されることがあります。")
        print("     再実行の前に Word を一度手動で起動し、確認が出たら「はい」で開いて保存せずに閉じてください。")


def list_steps():
    print("実行順:")
    for i, s in enumerate(STEPS, 1):
        extra = []
        if s.get("args"):
            extra.append("引数 " + " ".join(s["args"]))
        if s.get("stdin"):
            extra.append("入力 " + s["stdin"].strip().replace("\n", ","))
        print(f"  {i:>2}. {s['script']}  （{s.get('note', '')}{'／' + '／'.join(extra) if extra else ''}）")


def parse_args(argv):
    start, until = 1, len(STEPS)
    args = list(argv)
    if "--list" in args:
        list_steps()
        sys.exit(0)
    if "--until" in args:
        i = args.index("--until")
        until = int(args[i + 1])
        del args[i:i + 2]
    if args:
        a = args[0]
        if a.isdigit():
            start = int(a)
        else:
            names = [s["script"].lower() for s in STEPS]
            if a.lower() not in names:
                print(f"❌ 一覧にないスクリプト名です: {a}（--list で一覧を確認）")
                sys.exit(1)
            start = names.index(a.lower()) + 1
    if not (1 <= start <= until <= len(STEPS)):
        print(f"❌ 範囲が正しくありません: {start} 〜 {until}（1 〜 {len(STEPS)}）")
        sys.exit(1)
    return start, until


def run_batch():
    print("========== クロス集計 パイプライン（一気通貫） ==========")
    print("💡 実行中は対象の Excel / Word ファイルを開かず、コンソールの文字もクリックしないでください。")
    start, until = parse_args(sys.argv[1:])
    base_dir = os.path.dirname(os.path.abspath(__file__))
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    targets = list(enumerate(STEPS, 1))[start - 1:until]
    if start > 1 or until < len(STEPS):
        print(f"⏩ {start} 番〜{until} 番を実行します（{STEPS[start - 1]['script']} から）。")

    missing = [s["script"] for _, s in targets if not os.path.exists(os.path.join(base_dir, s["script"]))]
    needs = {"062_sub_clustering.py": ["064_sub_clustering_multi.py"],          # 064 は 062 の関数を読み込む
             "089_yearbook_layers.py": ["075_export_master_data.py", "083_survey_strict_statistics.py",
                                        "086_yearbook_subcluster_features.py", "087_yearbook_dendrograms.py"]}
    for mod, users in needs.items():
        if any(s["script"] in users for _, s in targets) and not os.path.exists(os.path.join(base_dir, mod)):
            missing.append(mod)
    if missing:
        print("❌ 次のスクリプトが作業フォルダにありません。置いてから再実行してください:")
        for m in dict.fromkeys(missing):
            print(f"   - {m}")
        sys.exit(1)

    times, total_start, cur = [], time.time(), None
    try:
        for no, step in targets:
            cur = (no, step)
            name = step["script"]

            if step.get("check_office"):
                left = running_office(["WINWORD.EXE", "EXCEL.EXE"])
                if left:
                    print(f"\n🛑 起動中の Office があります: {', '.join(left)}")
                    print("   保存して閉じるか、スクリプトが残したものなら Stop-Process で終了してから、再開してください:")
                    print(f"   python 000_run_pipeline_cross.py {no}")
                    sys.exit(1)
            if step.get("confirm"):
                pause = time.time()
                try:
                    ans = input(f"\n🛑 次は {name}（{step.get('note', '')}）です。未保存のWordファイルがあれば保存して閉じてください。\n"
                                "👉 続行は Enter / 中止は q : ")
                except EOFError:
                    ans = ""
                total_start += time.time() - pause
                if ans.strip().lower() in ("q", "quit", "exit"):
                    print("中止しました。")
                    sys.exit(130)

            extra = []
            if step.get("args"):
                extra.append("引数 " + " ".join(step["args"]))
            if step.get("stdin"):
                extra.append("自動入力 " + step["stdin"].strip().replace("\n", ",")) if step.get("stdin") else None
            print(f"\n▶ [{no}/{len(STEPS)}] {name}  {step.get('note', '')}" + (f"（{'／'.join(extra)}）" if extra else ""))
            t0 = time.time()
            subprocess.run(
                [sys.executable, "-u", os.path.join(base_dir, name)] + step.get("args", []),
                cwd=base_dir, check=True, timeout=step.get("timeout", 60), env=env,
                input=step.get("stdin"), text=True if step.get("stdin") is not None else None,
            )
            el = time.time() - t0
            times.append((no, name, el))
            print(f"✅ {name} 完了（{el:.1f} 秒）")

    except subprocess.TimeoutExpired:
        no, step = cur
        print(f"\n❌ 【タイムアウト】{no}. {step['script']} が {step.get('timeout', 60)} 秒を超えました。")
        if step.get("office"):
            kill_office_safely(step["office"])
        print(f"   原因を確認後、次のコマンドで再開できます: python 000_run_pipeline_cross.py {no}")
        sys.exit(1)
    except subprocess.CalledProcessError as e:
        no, step = cur
        if e.returncode == 2:
            print(f"\n⚠️ 【スキップ】{step['script']} が排他ロック等により安全にスキップされました。")
            sys.exit(0)
        print(f"\n❌ 【エラー】{no}. {step['script']} が異常終了しました（終了コード {e.returncode}）。上のメッセージを確認してください。")
        if step.get("office"):
            kill_office_safely(step["office"])
        print(f"   修正後、次のコマンドで再開できます: python 000_run_pipeline_cross.py {no}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n🛑 手動で中断されました（Ctrl+C）。")
        if cur and cur[1].get("office"):
            kill_office_safely(cur[1]["office"])
        sys.exit(130)

    print("\n================ 所要時間 ================")
    for no, name, el in times:
        print(f"  {no:>2}. {name:<45} {el:>7.1f} 秒")
    print(f"  合計 {time.time() - total_start:.1f} 秒")
    print("🎉 すべてのステップが正常に完了しました。")


if __name__ == "__main__":
    run_batch()
