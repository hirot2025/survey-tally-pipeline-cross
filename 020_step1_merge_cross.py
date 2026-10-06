"""
survey-tally-pipeline-simple
Copyright (c) 2026 HIROI Takayuki
Released under the MIT license
https://github.com/hirot2025/survey-tally-pipeline-simple/blob/main/LICENSE.md
"""

import os
import glob
import re
import pandas as pd
import numpy as np

# ==========================================
# 設定
# ==========================================
INPUT_DIR = "011_ja100_csv_formatted_cross"
OUTPUT_DIR = "020_ja100_step1_merge_cross"

# =========================================================
# 最終確定版 セルマッピング定義
# =========================================================
CHECKBOX_CONFIGS = [
    {"range": "G14:G24", "link_start": "'集計用'!$L$9"},
    {"range": "H14:H24", "link_start": "'集計用'!$L$21"},
    {"range": "F31:F34", "link_start": "'集計用'!$L$37"},
    {"range": "G77:G87", "link_start": "'集計用'!$L$103"},
    {"range": "H77:H87", "link_start": "'集計用'!$L$115"},
    {"range": "E94:E107", "link_start": "'集計用'!$L$131"},
    {"range": "F94:F107", "link_start": "'集計用'!$L$146"},
    {"range": "F115:F128", "link_start": "'集計用'!$L$166"},
    {"range": "F136:F140", "link_start": "'集計用'!$L$190"},
    {"range": "H157:H169", "link_start": "'集計用'!$L$208"},
    {"range": "D196:D201", "link_start": "'集計用'!$L$302"},
    {"range": "E196:E201", "link_start": "'集計用'!$L$309"},
    {"range": "F196:F201", "link_start": "'集計用'!$L$316"},
    {"range": "H242:H254", "link_start": "'集計用'!$L$355"},
    {"range": "G299:G310", "link_start": "'集計用'!$L$401"}
]

RADIO_CONFIGS = [
    {"range": "D45:D47", "link": "'集計用'!$M$47", "options": 3, "opt_dir": "V"},
    {"range": "E53:E57", "link": "'集計用'!$M$54", "options": 5, "opt_dir": "V"},
    {"range": "D64:H64", "link": "'集計用'!$M$62", "options": 5, "opt_dir": "H"},
    {"range": "D65:H65", "link": "'集計用'!$M$68", "options": 5, "opt_dir": "H"},
    {"range": "D66:H66", "link": "'集計用'!$M$74", "options": 5, "opt_dir": "H"},
    {"range": "D67:H67", "link": "'集計用'!$M$80", "options": 5, "opt_dir": "H"},
    {"range": "D68:H68", "link": "'集計用'!$M$86", "options": 5, "opt_dir": "H"},
    {"range": "D69:H69", "link": "'集計用'!$M$92", "options": 5, "opt_dir": "H"},
    {"range": "H147:H152", "link": "'集計用'!$M$200", "options": 5, "opt_dir": "V"},
    {"range": "C179:C183", "link": "'集計用'!$M$228", "options": 5, "opt_dir": "V"},
    {"range": "D179:D183", "link": "'集計用'!$M$234", "options": 5, "opt_dir": "V"},
    {"range": "E179:E183", "link": "'集計用'!$M$240", "options": 5, "opt_dir": "V"},
    {"range": "F179:F183", "link": "'集計用'!$M$246", "options": 5, "opt_dir": "V"},
    {"range": "G179:G183", "link": "'集計用'!$M$252", "options": 5, "opt_dir": "V"},
    {"range": "H179:H183", "link": "'集計用'!$M$258", "options": 5, "opt_dir": "V"},
    {"range": "C186:C190", "link": "'集計用'!$M$264", "options": 5, "opt_dir": "V"},
    {"range": "D186:D190", "link": "'集計用'!$M$270", "options": 5, "opt_dir": "V"},
    {"range": "E186:E190", "link": "'集計用'!$M$276", "options": 5, "opt_dir": "V"},
    {"range": "F186:F190", "link": "'集計用'!$M$282", "options": 5, "opt_dir": "V"},
    {"range": "G186:G190", "link": "'集計用'!$M$288", "options": 5, "opt_dir": "V"},
    {"range": "H186:H190", "link": "'集計用'!$M$294", "options": 5, "opt_dir": "V"},
    {"range": "D227:D230", "link": "'集計用'!$M$341", "options": 4, "opt_dir": "V"},
    {"range": "E264:E266", "link": "'集計用'!$M$374", "options": 3, "opt_dir": "V"},
    {"range": "G277:G279", "link": "'集計用'!$M$384", "options": 4, "opt_dir": "V"},
    {"range": "G290:G294", "link": "'集計用'!$M$394", "options": 5, "opt_dir": "V"}
]

TEXT_CONFIGS = [
    {"range": "D27", "link": "'集計用'!$L$34"},
    {"range": "D37", "link": "'集計用'!$L$43"},
    {"range": "D72", "link": "'集計用'!$L$99"},
    {"range": "D90", "link": "'集計用'!$L$128"},
    {"range": "D110", "link": "'集計用'!$L$162"},
    {"range": "D131", "link": "'集計用'!$L$187"},
    {"range": "D143", "link": "'集計用'!$L$197"},
    {"range": "D173", "link": "'集計用'!$L$224"},
    {"range": "D205", "link": "'集計用'!$L$325"},
    {"range": "B209", "link": "'集計用'!$L$328"},
    {"range": "B210", "link": "'集計用'!$L$329"},
    {"range": "D234", "link": "'集計用'!$L$348"},
    {"range": "D237", "link": "'集計用'!$L$351"},
    {"range": "D257", "link": "'集計用'!$L$370"},
    {"range": "D283", "link": "'集計用'!$L$390"},
    {"range": "D313", "link": "'集計用'!$L$415"},
    {"range": "D316", "link": "'集計用'!$L$418"}
]

# ==========================================
# マッピング設定からの動的行番号計算関数
# ==========================================
def parse_excel_row(link_str):
    """ '$L$9' のような文字列から行番号(9)を抽出 """
    match = re.search(r'\$[A-Z]+\$(\d+)', link_str)
    return int(match.group(1)) if match else None

def get_target_rows():
    """ 抽出対象となるExcelの行番号リストを自動生成 """
    cb_rows = []
    for c in CHECKBOX_CONFIGS:
        start_row = parse_excel_row(c["link_start"])
        m = re.search(r'[A-Z](\d+):[A-Z](\d+)', c["range"])
        count = int(m.group(2)) - int(m.group(1)) + 1
        cb_rows.extend(range(start_row, start_row + count))
        
    # ▼【数式崩壊の対策】ラジオボタンはL列の数式ではなく、M列の値を見てPython側で演算する
    rb_logic = []
    for r in RADIO_CONFIGS:
        start_row = parse_excel_row(r["link"])
        for i in range(r["options"]):
            rb_logic.append({
                "out_row": start_row + i,
                "m_row": start_row,          # M列の参照行
                "opt_val": str(i + 1)        # 判定する選択肢の番号(1, 2, 3...)
            })
        
    txt_rows = [parse_excel_row(t["link"]) for t in TEXT_CONFIGS]
    
    return cb_rows, rb_logic, txt_rows

def step1_merge_csvs():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    csv_files = glob.glob(os.path.join(INPUT_DIR, "*.csv"))
    if not csv_files:
        print(f"エラー: '{INPUT_DIR}' フォルダ内にCSVファイルが見つかりません。")
        return
        
    print(f"【開始】{len(csv_files)}件のCSVファイルからデータを抽出し、マージします...")

    # 動的に行番号リストを生成
    checkbox_rows, radio_logic, text_rows = get_target_rows()

    def get_val(df, row_idx, col_idx):
        try:
            val = df.iat[row_idx, col_idx]
            return val if pd.notna(val) else ""
        except IndexError:
            return ""

    all_records = []
    success_count = 0

    for file_idx, file_path in enumerate(csv_files, 1):
        filename = os.path.basename(file_path)
        
        try:
            df = pd.read_csv(file_path, header=None, dtype=str, encoding="cp932")
        except Exception:
            try:
                df = pd.read_csv(file_path, header=None, dtype=str, encoding="utf-8")
            except Exception as e:
                print(f"  [スキップ] {filename} の読み込みに失敗: {e}")
                continue

        # ▼【列数拡張】L列(11), M列(12)が欠落しないよう、さらにO〜R列等の
        # ガイドラインテキストも切り捨てられないように余裕を持って19列(S列相当)まで拡張する
        if len(df.columns) < 19:
            df = df.reindex(columns=range(19))

        # ▼【E列混入対策】H4のメールアドレスが下へコピー(ffill)されるのを防ぐため、
        # メタデータの本体が始まる9行目(インデックス8)以降にのみ前詰め処理を限定する
        df.iloc[8:, 7:11] = df.iloc[8:, 7:11].replace('空欄', np.nan).ffill()

        # ４）その他１：基本情報の取得 (指定されたセル番地から直取り)
        ja_name    = get_val(df, 1, 1) # B2
        dep_name   = get_val(df, 1, 4) # E2
        person     = get_val(df, 2, 4) # E3
        mail_left  = get_val(df, 3, 4) # E4
        mail_right = get_val(df, 3, 7) # H4

        if not ja_name:
            ja_name = f"未記入JA_{file_idx}"
            
        file_id = filename.replace(".csv", "")
        base = {"通し番号": file_idx, "元ファイル名": filename, "集計用ID": file_id, "JA名": ja_name}
        
        records = []
        
        # 基本情報をレコードに追加
        records.append({**base, "行番号": 2, "対象列": "B", "データ種別": "基本情報(JA名)", "回答値": ja_name})
        records.append({**base, "行番号": 2, "対象列": "E", "データ種別": "基本情報(部署)", "回答値": dep_name})
        records.append({**base, "行番号": 3, "対象列": "E", "データ種別": "基本情報(氏名)", "回答値": person})
        records.append({**base, "行番号": 4, "対象列": "E", "データ種別": "基本情報(メール左)", "回答値": mail_left})
        records.append({**base, "行番号": 4, "対象列": "H", "データ種別": "基本情報(メール右)", "回答値": mail_right})
        
        # １）チェックボックス抽出 (L列 = index 11)
        for r in checkbox_rows:
            records.append({**base, "行番号": r, "対象列": "L", "データ種別": "チェックボックス", "回答値": get_val(df, r-1, 11)})
            
        # ２）ラジオボタン抽出 (▼数式ではなく、Python側でM列から判定して直接テキストの「1/0」を出力する)
        for rb in radio_logic:
            m_val = get_val(df, rb["m_row"] - 1, 12) # M列(インデックス12)の値を取得
            
            # 選択された値(m_val)と、この行が担当する選択肢番号(opt_val)を比較
            try:
                # "1" と "1.0" などの表記ブレを吸収して安全に比較し、テキストとして確定
                is_selected = "1" if str(int(float(m_val))) == rb["opt_val"] else "0"
            except (ValueError, TypeError):
                is_selected = "0" # 空欄やエラー値の場合は0
                
            records.append({**base, "行番号": rb["out_row"], "対象列": "L", "データ種別": "ラジオボタン", "回答値": is_selected})
            
        # ３）記述回答抽出 (L列 = index 11)
        for r in text_rows:
            records.append({**base, "行番号": r, "対象列": "L", "データ種別": "記述回答", "回答値": get_val(df, r-1, 11)})
            
        df_records = pd.DataFrame(records)
        
        # ５）その他２：メタデータ(大項目, 設問, 中項目, 小項目)の結合
        def get_meta(row):
            # 基本情報の行には、メタデータ(大項目IDなど)を紐付けずに空欄にする
            if "基本情報" in str(row["データ種別"]):
                return pd.Series([np.nan]*4, index=["大項目ID", "設問ID", "中項目ID", "小項目ID"])
                
            r_idx = row["行番号"] - 1
            if r_idx < len(df) and len(df.columns) >= 11:
                return pd.Series(df.iloc[r_idx, [7, 8, 9, 10]].values, index=["大項目ID", "設問ID", "中項目ID", "小項目ID"])
            return pd.Series([np.nan]*4, index=["大項目ID", "設問ID", "中項目ID", "小項目ID"])
            
        df_meta = df_records.apply(get_meta, axis=1)
        df_final = pd.concat([df_records, df_meta], axis=1)
        df_final["セル位置"] = df_final["対象列"] + df_final["行番号"].astype(str)
        
        all_records.append(df_final)
        success_count += 1
        print(f"  -> {filename} を抽出・マージしました")

    # 全データの統合と出力
    if all_records:
        df_all = pd.concat(all_records, ignore_index=True)
        # 列の並び順を整理
        output_columns = ["通し番号", "元ファイル名", "集計用ID", "JA名", "大項目ID", "設問ID", "中項目ID", "小項目ID", "データ種別", "セル位置", "回答値"]
        
        master_csv = os.path.join(OUTPUT_DIR, "020_ja100_master_merged_cross.csv")
        df_all[output_columns].to_csv(master_csv, index=False, encoding="utf-8-sig")
        print(f"\n【完了】全{success_count}件のデータを統合し '{master_csv}' を作成しました。")
        
        # 確認出力
        print("\n■ 次のステップに向けた確認 (大項目IDごとのデータ件数)")
        print(df_all["大項目ID"].value_counts(dropna=False))

if __name__ == "__main__":
    step1_merge_csvs()