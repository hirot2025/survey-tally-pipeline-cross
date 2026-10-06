"""
051_horizon_conv_cross_100ja_cleaned.py
050 → 051 のクリーニング（True/False→1/0、0人列の削除、スキップロジック補正、その他記述の抽出）
※ 2026-09-29 修正: メタ行を「先頭30行固定」から「先頭から続くメタ行を自動判定」に変更。
   修正前は先頭12件のJA（あ〜こ）が変換・補正から漏れていた。
"""
import pandas as pd
import warnings
import re
warnings.filterwarnings('ignore')

def main():
    input_file = '050_horizon_conv_cross_100ja.csv'
    output_file = '051_horizon_conv_cross_100ja_cleaned.csv'
    audit_file = '051_audit.csv'
    other_answers_file = '051_other_answers.txt'

    print(f"[{input_file}] を読み込んでいます...")
    try:
        df_raw = pd.read_csv(input_file, encoding='utf-8')
    except UnicodeDecodeError:
        df_raw = pd.read_csv(input_file, encoding='cp932')

    ja_col = df_raw.columns[0]

    exclude_keywords = ['項目', '設問', '【', '値', 'パーセンタイル', '平均', '最大', '最小', '標準偏差', '分散', '全体', 'ID']
    pattern = '|'.join(exclude_keywords)

    # ★修正(2026-09-29): メタ行の数を「先頭30行」と固定していたため、メタ行が18行しかない050では
    #   先頭12件のJA行がメタ行扱いになり、True/False→1/0変換とスキップロジック補正から漏れていた。
    #   先頭から「JA名が空欄またはメタ用キーワードを含む行」が続く範囲をメタ行とする。
    ja_str = df_raw[ja_col].astype(str).str.strip()
    is_meta_row = df_raw[ja_col].isna() | (ja_str == '') | ja_str.str.contains(pattern, na=False, regex=True)
    n_meta = int(is_meta_row.values.argmin()) if (~is_meta_row).any() else len(df_raw)
    df_meta = df_raw.head(n_meta).copy()
    df_ja = df_raw.iloc[n_meta:].dropna(subset=[ja_col])
    df_ja = df_ja[~df_ja[ja_col].astype(str).str.contains(pattern, na=False, regex=True)].copy()
    late_meta = int(is_meta_row.iloc[n_meta:].sum())
    print(f"  メタ行: 先頭 {n_meta} 行 / JA行: {len(df_ja)} 件" + (f"（JA行の後ろにメタ扱いの行 {late_meta} 行あり・除外）" if late_meta else ""))
    for r in (11, 15, 16, 17):
        if r >= n_meta:
            print(f"  [警告] ラベル取得に使う {r} 行目がメタ行の範囲外です（label_of・0人列の判定を確認してください）")

    for col in df_ja.columns:
        if col != ja_col:
            df_ja[col] = df_ja[col].apply(
                lambda x: 1 if str(x).strip().upper() in ['TRUE', '1', '1.0', '〇', '○'] 
                else (0 if str(x).strip().upper() in ['FALSE', '0', '0.0', '-1', '×', 'NAN'] else x)
            )

    audit_records = []
    other_answers = []

    def find_cols(keyword1, keyword2=None):
        matched = []
        for col in df_raw.columns:
            if col == ja_col: continue
            meta_str = " ".join(df_meta[col].fillna("").astype(str))
            if keyword1 in meta_str:
                if keyword2 is None or keyword2 in meta_str:
                    matched.append(col)
        return matched

    def find_single_col(keyword1, keyword2=None):
        cols = find_cols(keyword1, keyword2)
        return cols[0] if cols else None

    def label_of(col):
        target_rows = [15, 11, 16, 17]
        texts = []
        for r in target_rows:
            if r in df_meta.index:
                val = df_meta.at[r, col]
                if pd.notna(val) and str(val).strip() != '' and str(val).strip() != 'nan':
                    texts.append(str(val).strip())
        
        unique_texts = []
        for t in texts:
            # 💡 修正：「0人」が「10人」に吸い込まれる過剰な重複除去を廃止し、完全一致のみ除外
            if t not in unique_texts:
                unique_texts.append(t)
                
        txt = " ＞ ".join(unique_texts)
        return txt if txt else "(ラベル取得不可)"

    # ============================================================================
    # ★修正：「0人」列の削除を、ラベル文字列ではなくメタデータの「セル値」で直接判定
    # ============================================================================
    cols_to_drop = []
    for c in df_ja.columns:
        if c == ja_col: continue
        
        # 11行目(設問ID)と17行目(小項目)のセルの文字を直接確認
        q_id_val = str(df_meta.at[11, c]).strip() if 11 in df_meta.index else ""
        sub_item_val = str(df_meta.at[17, c]).strip() if 17 in df_meta.index else ""
        
        # 設問がQ5-1かQ12-1であり、かつ小項目が「0人」であれば無条件で削除
        if ("Q5-1" in q_id_val or "Q12-1" in q_id_val) and sub_item_val == "0人":
            cols_to_drop.append(c)
            
    if cols_to_drop:
        print("\n" + "=" * 60)
        print("【事前処理：分析ノイズとなる「0人」選択肢列の削除】")
        print("=" * 60)
        for c in cols_to_drop:
            print(f"  - 削除済: {c} : {label_of(c)}")
            
        df_ja = df_ja.drop(columns=cols_to_drop, errors='ignore')
        df_meta = df_meta.drop(columns=cols_to_drop, errors='ignore')
        df_raw = df_raw.drop(columns=cols_to_drop, errors='ignore')

    # ============================================================================
    # Error A / B：スキップロジック
    # ============================================================================
    rules = [
        {
            "error_type": "A",
            "name": "Q5-1全滅 -> Q6〜Q10, Q12の一括無効(Q11除く)",
            "condition_type": "group_sum",
            "parent_cols": find_cols("Q5-1"),
            "child_cols": (
                [c for c in find_cols("Q6-1") if "Q6-2" not in " ".join(df_meta[c].fillna("").astype(str))] +
                [c for c in find_cols("Q7-1") if "Q7-2" not in " ".join(df_meta[c].fillna("").astype(str))] +
                [c for c in find_cols("Q8-1") if "Q8-2" not in " ".join(df_meta[c].fillna("").astype(str))] +
                find_cols("Q10") +
                find_cols("Q12-1") +
                find_cols("Q12-2") +
                find_cols("Q12-4")
            ),
        },
        {
            "error_type": "B",
            "name": "Q8-1(住居の提供)未選択 -> Q9-1無効",
            "condition_type": "single",
            "parent_col": find_single_col("Q8-1", "住居の提供"),
            "child_cols": [c for c in find_cols("Q9-1") if "Q9-2" not in " ".join(df_meta[c].fillna("").astype(str))],
        },
        {
            "error_type": "B",
            "name": "Q13-1(②)未選択 -> Q13-2(②)無効",
            "condition_type": "single",
            "parent_col": find_single_col("Q13-1", "②"),
            "child_cols": [c for c in find_cols("Q13-2") if "②" in " ".join(df_meta[c].fillna("").astype(str))],
        },
        {
            "error_type": "B",
            "name": "Q13-1(④)未選択 -> Q13-2(④)無効",
            "condition_type": "single",
            "parent_col": find_single_col("Q13-1", "④"),
            "child_cols": [c for c in find_cols("Q13-2") if "④" in " ".join(df_meta[c].fillna("").astype(str))],
        },
        {
            "error_type": "B",
            "name": "Q13-1(①)未選択 -> Q14-1無効(⑬除く)",
            "condition_type": "single",
            "parent_col": find_single_col("Q13-1", "①"),
            "child_cols": [c for c in find_cols("Q14-1") if "Q14-2" not in " ".join(df_meta[c].fillna("").astype(str)) and "⑬" not in " ".join(df_meta[c].fillna("").astype(str))],
        },
        {
            "error_type": "B",
            "name": "Q15-1(②)未選択 -> Q15-2無効",
            "condition_type": "single",
            "parent_col": find_single_col("Q15-1", "②"),
            "child_cols": find_cols("Q15-2"),
        }
    ]

    # ============================================================================
    # Error C-1 / C-2：「その他」複数親（OR条件）完全対応版
    # ============================================================================
    other_rules = []
    text_cols = find_cols("「その他」を選択した場合") + find_cols("「その他」の具体的内容")
    text_cols = list(set(text_cols))
    columns_list = list(df_raw.columns)

    for txt_col in text_cols:
        txt_idx = columns_list.index(txt_col)
        
        txt_meta_str = " ".join(df_meta[txt_col].fillna("").astype(str))
        q_match = re.search(r'Q(\d+)', txt_meta_str)
        txt_q_num = q_match.group(1) if q_match else None

        matched_chk_cols = []
        for i in range(txt_idx - 1, 0, -1):
            candidate = columns_list[i]
            meta_str = " ".join(df_meta[candidate].fillna("").astype(str))
            
            if txt_q_num:
                cand_match = re.search(r'Q(\d+)', meta_str)
                if cand_match and cand_match.group(1) != txt_q_num:
                    break
            
            if "その他" in meta_str and "具体的内容" not in meta_str and "記述" not in meta_str:
                matched_chk_cols.append(candidate)

        if matched_chk_cols:
            matched_chk_cols.reverse()
            other_rules.append({
                "name": f"Q{txt_q_num} その他" if txt_q_num else "その他",
                "check_cols": matched_chk_cols,
                "text_col": txt_col
            })

    print("\n" + "=" * 70)
    print("【事前検証：ルールが実際にどの列にマッチしたか】")
    print("=" * 70)
    for rule in rules:
        if rule["error_type"] == "A":
            p_list = rule.get("parent_cols", [])
        else:
            p_list = [rule.get("parent_col")] if rule.get("parent_col") else []

        print(f"\n[{rule['error_type']}] {rule['name']}")
        print(f"  親列（トリガー）({len(p_list)}件):")
        for c in p_list:
            print(f"    - {c} : {label_of(c)}")

        c_list = rule.get("child_cols", [])
        print(f"  子列（消去対象）({len(c_list)}件):")
        for c in c_list:
            print(f"    - {c} : {label_of(c)}")

    print(f"\n[C1/C2] その他ペアリング 自動検出 {len(other_rules)} 件")
    for r in other_rules:
        print(f"  - {r['name']} 記述列: {r['text_col']}")

    summary_counts = {"A": {"ja": 0, "cell": 0}, "B": {"ja": 0, "cell": 0},
                       "C1": {"ja": 0, "cell": 0}, "C2": {"ja": 0, "cell": 0}}

    print("\n" + "=" * 60)
    print("【実行処理：スキップロジックと自由記述の自動補正】")
    print("=" * 60)

    for rule in rules:
        etype = rule["error_type"]
        c_cols = rule.get("child_cols", [])

        if etype == "A":
            p_cols = rule.get("parent_cols", [])
            if not p_cols or not c_cols:
                continue
            df_ja[p_cols] = df_ja[p_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
            cond = df_ja[p_cols].sum(axis=1) == 0
            trigger_lbl = f"{rule['name']} [{label_of(p_cols[0])} 等]"
        else:
            p_col = rule.get("parent_col")
            if not p_col or not c_cols:
                continue
            df_ja[p_col] = pd.to_numeric(df_ja[p_col], errors='coerce').fillna(0)
            cond = df_ja[p_col] == 0
            trigger_lbl = label_of(p_col)

        df_ja[c_cols] = df_ja[c_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
        child_sum = df_ja[c_cols].sum(axis=1)
        errors = df_ja[cond & (child_sum > 0)]

        if len(errors) > 0:
            for idx, row in errors.iterrows():
                for c in c_cols:
                    before = row[c]
                    if before != 0 and pd.notna(before) and before != '':
                        audit_records.append({
                            "JA_ID": row[ja_col],
                            "エラー種別": f"{etype} (親設問未選択による関連設問無効)",
                            "トリガー条件（親）": trigger_lbl,
                            "補正対象（子）": label_of(c),
                            "変更前": before,
                            "変更後": 0
                        })
                        summary_counts[etype]["cell"] += 1
                        df_ja.at[idx, c] = 0
                summary_counts[etype]["ja"] += 1

    for r in other_rules:
        chk_cols, txt_col = r.get("check_cols"), r.get("text_col")
        df_ja[chk_cols] = df_ja[chk_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
        
        any_checked = df_ja[chk_cols].sum(axis=1) > 0

        try:
            trigger_summary = " OR ".join([label_of(c).split(' ＞ ')[-2] if len(label_of(c).split(' ＞ ')) >= 2 else label_of(c).split(' ＞ ')[-1] for c in chk_cols])
            trigger_full_lbl = f"{label_of(chk_cols[0]).split(' ＞ ')[0]} ＞ [対象: {trigger_summary}]"
        except Exception:
            trigger_full_lbl = label_of(chk_cols[0])

        for idx, row in df_ja.iterrows():
            checked = any_checked[idx]
            text_val = row[txt_col]
            has_text = pd.notna(text_val) and str(text_val).strip().upper() not in ('', '0', 'NAN', '0.0', 'FALSE')

            if checked and has_text:
                other_answers.append(f"[{r['name']}] {row[ja_col]} : {text_val}")
            elif checked and not has_text:
                summary_counts["C2"]["ja"] += 1
                audit_records.append({
                    "JA_ID": row[ja_col],
                    "エラー種別": "C2 (選択あり・記述なしの記載漏れ)",
                    "トリガー条件（親）": trigger_full_lbl,
                    "補正対象（子）": label_of(txt_col),
                    "変更前": text_val,
                    "変更後": text_val
                })
            elif (not checked) and has_text:
                summary_counts["C1"]["ja"] += 1
                summary_counts["C1"]["cell"] += 1
                audit_records.append({
                    "JA_ID": row[ja_col],
                    "エラー種別": "C1 (チェック忘れ → 生の声を保護)",
                    "トリガー条件（親）": trigger_full_lbl,
                    "補正対象（子）": label_of(txt_col),
                    "変更前": text_val,
                    "変更後": text_val
                })
                other_answers.append(f"[{r['name']}] {row[ja_col]} : {text_val} (※チェック忘れを保護)")

    print("\n" + "=" * 60)
    print("【補正サマリー】")
    print("=" * 60)
    print(f"  Error A (親設問の全滅による配下設問の一括無効) : 対象JA {summary_counts['A']['ja']}件 / 対象セル {summary_counts['A']['cell']}件")
    print(f"  Error B (特定選択肢の未選択による関連設問無効) : 対象JA {summary_counts['B']['ja']}件 / 対象セル {summary_counts['B']['cell']}件")
    print(f"  Error C1(記述あり・チェック忘れのデータを保護)   : 対象JA {summary_counts['C1']['ja']}件 / 対象セル {summary_counts['C1']['cell']}件")
    print(f"  Error C2(選択済なのに記述なし・要確認)         : 対象JA {summary_counts['C2']['ja']}件 / 対象セル {summary_counts['C2']['cell']}件")

    df_final = pd.concat([df_meta, df_ja], ignore_index=True)

    dup_count = df_final[ja_col].duplicated().sum()
    if dup_count > 0:
        print(f"\n[警告] {ja_col} に重複が{dup_count}件見つかりました。")
    else:
        print(f"\n[OK] {ja_col} の重複なし")

    df_final.to_csv(output_file, encoding='utf-8', index=False)

    # ★追加: 変換漏れの確認（出力のJA行にTrue/Falseが残っていないか）
    tf_left = df_ja.drop(columns=[ja_col]).astype(str).apply(lambda s: s.str.strip().str.upper().isin(['TRUE', 'FALSE'])).values.sum()
    print(f"[確認] JA行に残ったTrue/Falseのセル: {int(tf_left)} 件（0であれば変換漏れなし）")

    audit_cols = ["JA_ID", "エラー種別", "トリガー条件（親）", "補正対象（子）", "変更前", "変更後"]
    if audit_records:
        pd.DataFrame(audit_records)[audit_cols].to_csv(audit_file, encoding='utf-8-sig', index=False)
    else:
        pd.DataFrame(columns=audit_cols).to_csv(audit_file, encoding='utf-8-sig', index=False)

    with open(other_answers_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(other_answers) if other_answers else "有効な「その他」記述はありませんでした。")

    print("\n処理完了：以下のファイルを出力しました。")
    print(f"  - {output_file} (分析用本番データ)")
    print(f"  - {audit_file} (証跡ログ・階層ラベル付き)")
    print(f"  - {other_answers_file} (テキスト抽出)")

if __name__ == "__main__":
    main()