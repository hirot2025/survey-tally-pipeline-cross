import os
import pandas as pd
import re
import numpy as np
import unicodedata

# =========================================================
# 1. ファイルパスの設定
# =========================================================
file_020 = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'020_ja100_step1_merge_cross\020_ja100_master_merged_cross.csv')
file_081 = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'081_question_id_mapping_real_cross.csv')
file_080 = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'080_questionnaire_template_real_cross.csv')
output_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), r'050_horizon_conv_cross_100ja.csv')

# =========================================================
# 2. 020データの読み込みとマトリクス化（ソート完備）
# =========================================================
print("1) 020データを読み込み中...")
try:
    df_020 = pd.read_csv(file_020, encoding="cp932").fillna("")
except UnicodeDecodeError:
    df_020 = pd.read_csv(file_020, encoding="utf-8-sig").fillna("")

df_020["セル位置"] = df_020["セル位置"].astype(str).str.strip()

# JA名 × セル位置 × 回答値
df_matrix = df_020.pivot_table(
    index="JA名",
    columns="セル位置",
    values="回答値",
    aggfunc="first"
).fillna("")

# カラムを「文字」と「数字」に分けて正しくソートする
def sort_cell_position(col_name):
    match = re.match(r'([A-Za-z]+)(\d+)', str(col_name))
    if match:
        return (match.group(1).upper(), int(match.group(2)))
    return (str(col_name), 0)

sorted_columns = sorted(df_matrix.columns, key=sort_cell_position)
df_matrix = df_matrix[sorted_columns]

# =========================================================
# 2.5. 081マッピングデータの読み込みとJA名照合の準備
# =========================================================
print("2) エリア・支所・G列マッピング (081) を準備中...")
try:
    # 8行目(skiprows=7)以降から B列(1), C列(2), E列(4), ★G列(6) を抽出
    df_map = pd.read_csv(file_081, skiprows=7, usecols=[1, 2, 4, 6], names=['エリア名', '支所名', 'JA名', '集計用ID'], encoding="cp932").fillna("不明")
except UnicodeDecodeError:
    df_map = pd.read_csv(file_081, skiprows=7, usecols=[1, 2, 4, 6], names=['エリア名', '支所名', 'JA名', '集計用ID'], encoding="utf-8-sig").fillna("不明")

def normalize_ja_name(name):
    if pd.isna(name): return ""
    val = unicodedata.normalize('NFKC', str(name))
    val = re.sub(r'\s+', '', val)
    val = val.upper()
    val = val.replace('農業協同組合', 'JA').replace('農協', 'JA')
    return val

# 照合用キーを作成して辞書化
df_map['__MatchKey__'] = df_map['JA名'].apply(normalize_ja_name)
df_map = df_map.drop_duplicates(subset='__MatchKey__', keep='first')
map_dict_area = dict(zip(df_map['__MatchKey__'], df_map['エリア名']))
map_dict_branch = dict(zip(df_map['__MatchKey__'], df_map['支所名']))
map_dict_g_col = dict(zip(df_map['__MatchKey__'], df_map['集計用ID'])) # ★追加

# 020データ(df_matrix)のJA名(インデックス)に合わせて配列を作成
matrix_match_keys = [normalize_ja_name(ja) for ja in df_matrix.index]
area_col = [map_dict_area.get(k, '不明') for k in matrix_match_keys]
branch_col = [map_dict_branch.get(k, '不明') for k in matrix_match_keys]
g_col = [map_dict_g_col.get(k, '不明') for k in matrix_match_keys] # ★追加

# =========================================================
# 3. 080の読み込みとテキスト抽出
# =========================================================
print("3) テンプレート (080) を解析中...")
try:
    df_080 = pd.read_csv(file_080, header=None, encoding="cp932").fillna("")
except UnicodeDecodeError:
    df_080 = pd.read_csv(file_080, header=None, encoding="utf-8-sig").fillna("")

def get_cell_text(value):
    text = str(value).strip()
    if not text: return ""
    match = re.fullmatch(r"([A-Z]+)(\d+)", text.upper())
    if not match: return text
    
    col_str, row_str = match.groups()
    idx = 0
    for ch in col_str:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    col_idx = idx - 1
    row_idx = int(row_str) - 1
    
    try:
        return str(df_080.iloc[row_idx, col_idx]).strip()
    except IndexError:
        return text

dict_headers = {}
for idx in range(8, len(df_080)):
    key_cell = str(df_080.iloc[idx, 17]).strip()
    if not key_cell: continue
    
    match = re.search(r'\d+', key_cell)
    if match:
        row_num = match.group()
        dict_headers[row_num] = [
            str(df_080.iloc[idx, 7]).strip(),
            str(df_080.iloc[idx, 8]).strip(),
            str(df_080.iloc[idx, 9]).strip(),
            str(df_080.iloc[idx, 10]).strip(),
            get_cell_text(df_080.iloc[idx, 14]),
            get_cell_text(df_080.iloc[idx, 15]),
            get_cell_text(df_080.iloc[idx, 16]),
            get_cell_text(df_080.iloc[idx, 17])
        ]

# =========================================================
# 4. 8段ヘッダー構築
# =========================================================
print("4) 8段ヘッダーを構築中...")
row_H, row_I, row_J, row_K, row_O, row_P, row_Q, row_R = [], [], [], [], [], [], [], []

for col in df_matrix.columns:
    match = re.search(r'\d+', col)
    row_num = match.group() if match else ""
    
    if row_num and row_num in dict_headers:
        vals = dict_headers[row_num]
        row_H.append(vals[0])
        row_I.append(vals[1])
        row_J.append(vals[2])
        row_K.append(vals[3])
        row_O.append(vals[4])
        row_P.append(vals[5])
        row_Q.append(vals[6])
        row_R.append(vals[7])
    else:
        row_H.append("未照合")
        row_I.append("未照合")
        row_J.append("未照合")
        row_K.append("未照合")
        row_O.append("未照合")
        row_P.append("未照合")
        row_Q.append("未照合")
        row_R.append("未照合")

index_names = ["大項目", "設問ID", "中項目", "小項目", "大項目名", "設問ID名", "中項目名", "小項目名"]
df_headers = pd.DataFrame(
    [row_H, row_I, row_J, row_K, row_O, row_P, row_Q, row_R],
    columns=df_matrix.columns,
    index=index_names
)

# =========================================================
# 5. 10行の基礎統計量・空白行の構築 (厳密なフィルタリング適用)
# =========================================================
print("5) PCA・クラスタリングに向けた統計データ(10行)を計算中...")
df_num = df_matrix.replace("", np.nan).apply(pd.to_numeric, errors='coerce')

def calc_stats(col_name):
    match = re.search(r'\d+', col_name)
    row_num = match.group() if match else ""
    is_matched = bool(row_num and row_num in dict_headers)
    
    s_num = df_num[col_name]
    
    if not is_matched or s_num.isna().all():
        return ["", "", "", "", "", "", "", "", "", ""]
    
    s_orig = df_matrix[col_name].astype(str).str.strip()
    s_null = s_orig.eq("").sum()
    
    s_mean = round(s_num.mean(), 3)
    s_var = round(s_num.var(ddof=1), 3) if s_num.count() > 1 else 0
    s_std = round(s_num.std(ddof=1), 3) if s_num.count() > 1 else 0
    s_med = round(s_num.median(), 3)
    s_min = s_num.min()
    s_25 = round(s_num.quantile(0.25), 3)
    s_75 = round(s_num.quantile(0.75), 3)
    s_max = s_num.max()
    s_blank = ""
    
    return [s_mean, s_var, s_std, s_med, s_min, s_25, s_75, s_max, s_null, s_blank]

stats_data = []
for col in df_matrix.columns:
    stats_data.append(calc_stats(col))

stats_rows = list(map(list, zip(*stats_data)))

index_stats = [
    "【統計】平均",
    "【統計】分散",
    "【統計】標準偏差",
    "【統計】中央値",
    "【統計】最小値",
    "【統計】25パーセンタイル",
    "【統計】75パーセンタイル",
    "【統計】最大値",
    "【統計】欠損値(未回答)数",
    "【空白】"
]
df_stats = pd.DataFrame(stats_rows, columns=df_matrix.columns, index=index_stats)

# =========================================================
# 6. 「エリア名」「支所名」「集計用ID」の挿入と全データの結合・出力
# =========================================================
print("6) エリア名・支所名・集計用IDの列を挿入し、最終出力中...")

# ヘッダー行と統計行の先頭に空列を追加して位置を合わせる
df_headers.insert(0, 'エリア名', '')
df_headers.insert(1, '支所名', '')
df_headers.insert(2, '集計用ID', '')  # ★追加

df_stats.insert(0, 'エリア名', '')
df_stats.insert(1, '支所名', '')
df_stats.insert(2, '集計用ID', '')    # ★追加

# 実データ(df_matrix)の先頭に取得したエリア名と支所名、集計用IDを挿入
df_matrix.insert(0, 'エリア名', area_col)
df_matrix.insert(1, '支所名', branch_col)
df_matrix.insert(2, '集計用ID', g_col) # ★追加

# 縦に結合
df_final = pd.concat([df_stats, df_headers, df_matrix])
df_final.index.name = "JA名 / 項目"

df_final.to_csv(output_file, encoding="utf-8-sig")
print(f"\n✨ 処理完了！マスタと結合済みの完璧な分析用データを出力しました:\n{output_file}")