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

# ==========================================
# 設定
# ==========================================
INPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'010_ja100_csv_cross')
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'011_ja100_csv_formatted_cross')
TEMPLATE_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'080_questionnaire_template_real_cross.csv')

# ==========================================
# 入力範囲の定義リスト
# ==========================================
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
    {"range": "D45:D47", "link": "'集計用'!$M$47", "options": 3},
    {"range": "E53:E57", "link": "'集計用'!$M$54", "options": 5},
    {"range": "D64:H64", "link": "'集計用'!$M$62", "options": 5},
    {"range": "D65:H65", "link": "'集計用'!$M$68", "options": 5},
    {"range": "D66:H66", "link": "'集計用'!$M$74", "options": 5},
    {"range": "D67:H67", "link": "'集計用'!$M$80", "options": 5},
    {"range": "D68:H68", "link": "'集計用'!$M$86", "options": 5},
    {"range": "D69:H69", "link": "'集計用'!$M$92", "options": 5},
    {"range": "H147:H152", "link": "'集計用'!$M$200", "options": 5},
    {"range": "C179:C183", "link": "'集計用'!$M$228", "options": 5},
    {"range": "D179:D183", "link": "'集計用'!$M$234", "options": 5},
    {"range": "E179:E183", "link": "'集計用'!$M$240", "options": 5},
    {"range": "F179:F183", "link": "'集計用'!$M$246", "options": 5},
    {"range": "G179:G183", "link": "'集計用'!$M$252", "options": 5},
    {"range": "H179:H183", "link": "'集計用'!$M$258", "options": 5},
    {"range": "C186:C190", "link": "'集計用'!$M$264", "options": 5},
    {"range": "D186:D190", "link": "'集計用'!$M$270", "options": 5},
    {"range": "E186:E190", "link": "'集計用'!$M$276", "options": 5},
    {"range": "F186:F190", "link": "'集計用'!$M$282", "options": 5},
    {"range": "G186:G190", "link": "'集計用'!$M$288", "options": 5},
    {"range": "H186:H190", "link": "'集計用'!$M$294", "options": 5},
    {"range": "D227:D230", "link": "'集計用'!$M$341", "options": 4},
    {"range": "E264:E266", "link": "'集計用'!$M$374", "options": 3},
    {"range": "G277:G279", "link": "'集計用'!$M$384", "options": 4},
    {"range": "G290:G294", "link": "'集計用'!$M$394", "options": 5}
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

OTHER_CONFIGS = [
    {"range": "B2", "link": "'集計用'!$B$2"},
    {"range": "F2", "link": "'集計用'!$E$2"},
    {"range": "F3", "link": "'集計用'!$E$3"},
    {"range": "F4", "link": "'集計用'!$E$4"},
    {"range": "H4", "link": "'集計用'!$H$4"}
]

# ==========================================
# 処理関数
# ==========================================
def parse_link(link_str):
    """ '$M$47' のような文字列から行・列のインデックス(0始まり)を取得 """
    match = re.search(r'\$([A-Z]+)\$(\d+)', link_str)
    if match:
        col_str = match.group(1)
        col_idx = 0
        for char in col_str:
            col_idx = col_idx * 26 + (ord(char) - ord('A')) + 1
        return int(match.group(2)) - 1, col_idx - 1
    return None, None

def safe_get(df, r, c):
    """ データフレームから安全に値を取得し、欠損値なら空文字を返す """
    if r < len(df) and c < len(df.columns):
        val = df.iat[r, c]
        return "" if pd.isna(val) else val
    return ""

def main():
    print("【開始】CSVデータのピンポイント抽出・翻訳処理を開始します...")

    if not os.path.exists(INPUT_DIR):
        print(f"エラー: 入力フォルダ '{INPUT_DIR}' が見つかりません。")
        return
    if not os.path.exists(TEMPLATE_CSV):
        print(f"エラー: 雛型CSV '{TEMPLATE_CSV}' が見つかりません。")
        return

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    input_csv_files = glob.glob(os.path.join(INPUT_DIR, "*.csv"))
    
    if not input_csv_files:
        print(f"'{INPUT_DIR}' 内にCSVファイルが見つかりません。")
        return

    # 雛型の読み込み (001などが数値にならないよう str で読み込み)
    try:
        df_template = pd.read_csv(TEMPLATE_CSV, header=None, dtype=str, encoding="utf-8-sig")
    except UnicodeDecodeError:
        df_template = pd.read_csv(TEMPLATE_CSV, header=None, dtype=str, encoding="cp932")

    success_count = 0
    error_count = 0
    
    for file_path in input_csv_files:
        try:
            # 入力元CSVの読み込み
            try:
                df_in = pd.read_csv(file_path, header=None, dtype=str, encoding="utf-8-sig")
            except UnicodeDecodeError:
                df_in = pd.read_csv(file_path, header=None, dtype=str, encoding="cp932")

            # 雛型のコピーを作成し、枠を十分に広げる (M列=インデックス12まで担保)
            max_rows = max(len(df_in), len(df_template))
            max_cols = max(len(df_in.columns), len(df_template.columns), 13)
            df_out = df_template.copy()
            df_out = df_out.reindex(index=range(max_rows), columns=range(max_cols))
            df_out = df_out.fillna("")

            # ==========================================
            # 1) チェックボックス (範囲抽出)
            # ==========================================
            for item in CHECKBOX_CONFIGS:
                m = re.search(r'[A-Z](\d+):[A-Z](\d+)', item["range"])
                if m:
                    count = int(m.group(2)) - int(m.group(1)) + 1
                    r_start, c_idx = parse_link(item["link_start"])
                    if r_start is not None and c_idx is not None:
                        for i in range(count):
                            df_out.iat[r_start + i, c_idx] = safe_get(df_in, r_start + i, c_idx)

            # ==========================================
            # 2) ラジオボタン (M列抽出 + L列への数式翻訳)
            # ==========================================
            for item in RADIO_CONFIGS:
                r_idx, c_idx = parse_link(item["link"])  # c_idxはM列(12)
                if r_idx is not None and c_idx is not None:
                    # ① M列の値をピンポイントで上書き
                    df_out.iat[r_idx, c_idx] = safe_get(df_in, r_idx, c_idx)
                    
                    # ② L列(インデックス11)に数式を展開して翻訳
                    base_row_excel = r_idx + 1
                    for i in range(item["options"]):
                        df_out.iat[r_idx + i, 11] = f"=--($M${base_row_excel}={i+1})"

            # ==========================================
            # 3) 記述部分 (単一セル抽出)
            # ==========================================
            for item in TEXT_CONFIGS:
                r_idx, c_idx = parse_link(item["link"])
                if r_idx is not None and c_idx is not None:
                    df_out.iat[r_idx, c_idx] = safe_get(df_in, r_idx, c_idx)

            # ==========================================
            # 4) その他 (単一セル抽出)
            # ==========================================
            for item in OTHER_CONFIGS:
                r_idx, c_idx = parse_link(item["link"])
                if r_idx is not None and c_idx is not None:
                    df_out.iat[r_idx, c_idx] = safe_get(df_in, r_idx, c_idx)

            # ==========================================
            # 保存処理
            # ==========================================
            out_name = os.path.basename(file_path) # 入力ファイル名(ID名)をそのまま維持
            out_path = os.path.join(OUTPUT_DIR, out_name)
            
            df_out.to_csv(out_path, index=False, header=False, encoding="utf-8-sig")
            
            print(f"出力完了: {out_name}")
            success_count += 1
            
        except Exception as e:
            print(f"エラー ({os.path.basename(file_path)}): {e}")
            error_count += 1
            
    print("\n" + "=" * 40)
    print(f"処理完了: {success_count} 件成功, {error_count} 件エラー")
    print(f"出力先: '{OUTPUT_DIR}'")
    print("=" * 40)

if __name__ == "__main__":
    main()