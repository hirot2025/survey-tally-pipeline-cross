# survey-tally-pipeline-cross

アンケートの回答 Excel(100 件)を CSV 化・整形・結合し、1 行 1 回答者の横持ちデータに変換・クリーニングしたうえで、
設定の CSV(行列)に書いた組み合わせごとにクロス集計の Excel を作るパイプラインです。
(survey-tally-pipeline のクロス集計版。単純集計版は survey-tally-pipeline-simple)

このリポジトリには **実データ(回答・団体名・地域名・設問文)を含みません**。
同梱の **ダミーデータ** で 010 → 052 まで動かせます。

---

## 1. 処理の流れ

```
000_ja100_cross\*.xlsx(回答原本 100 件)
  → 010 回答原本の「集計用」シートを CSV に
  → 011 雛型 080 に合わせて整形
  → 020 100 件を1本の CSV に結合
  → 050 1 行 1 回答者の横持ちに変換(エリア名・支所名を 081 から付ける)
  → 051 クリーニング(True/False → 1/0、スキップロジックの補正、証跡ログ)
  → 052 設定の CSV ごとに、クロス集計の Excel を作る
  →(054 ダッシュボード 053 のリンク更新、055 Word の図表更新:053 と Word 文書は同梱していません)
```

実行用:`000_run_pipeline_cross.py`(010 → 052 を順に実行。`--list` で一覧、番号やスクリプト名で途中から再開)

| ファイル | 役割 |
|---|---|
| `010_generate_csvs_from_rawxlsx_cross.py` | 回答原本(xlsx)の「集計用」シートを CSV にする |
| `011_format_csv_data_cross.py` | 雛型 080 をもとに、回答だけを同じ番地に写す |
| `020_step1_merge_cross.py` | 100 件を1本の CSV に結合する |
| `050_horizon_conv_cross_100ja.py` | 横持ちに変換し、8 段の見出しとエリア名・支所名を付ける |
| `051_horizon_conv_cross_100ja_cleaned.py` | クリーニングと補正、証跡ログ(051_audit.csv)の出力 |
| `052_master_batch_processor_cross.py` | 設定の CSV(`*batch_processing_cross_*.csv`)ごとにクロス集計の xlsx を作る |
| `054_refresh_dashboard_links.py` | ダッシュボード(053.xlsx)のリンク更新(053 は同梱なし) |
| `055_ja100_word2021_update_cross.py` | Word 文書の図表更新(Word 文書は同梱なし) |
| `080_questionnaire_template_cross.csv` | 調査票の雛型(公開用。設問文と選択肢は空欄) |
| `081_question_id_mapping_dummy_cross.csv` | ID・表示名・エリア名・支所名の対応表(ダミー) |
| `sample\000_ja100_cross\` | ダミーの回答原本 100 件 |
| `sample\batch\` | 052 の設定の CSV の例(エリア名・支所名はダミー) |

---

## 2. ダミーデータで動かす

必要なもの:Windows、Microsoft Excel、Python 3.13

```powershell
pip install -r requirements.txt
Copy-Item .\080_questionnaire_template_cross.csv .\080_questionnaire_template_real_cross.csv
Copy-Item .\081_question_id_mapping_dummy_cross.csv .\081_question_id_mapping_real_cross.csv
Copy-Item .\sample\000_ja100_cross .\000_ja100_cross -Recurse
Copy-Item .\sample\batch\*.csv .
python .\000_run_pipeline_cross.py
```

スクリプトは本番用のファイル名(`_real_cross`)を読むため、ダミーをその名前で置いてから実行します。
これらと、実行で作られるファイル・フォルダは `.gitignore` で除外しています。

---

## 3. 052 の設定の CSV(クロス集計の組み合わせ)

- 1〜3 行目に横軸(受け手)の 設問ID・中項目・小項目、A〜C 列に縦軸(仕手)の 設問ID・中項目・小項目を並べる。
- 交わるセルに印を書く:`11`(縦軸を選択 かつ 横軸を選択)、`10`(選択 × 未選択)、`0`(未選択 × 選択)、`-1`(未選択 × 未選択)。
  例:Q3① の行と Q4① の列の交点に `11` → Q3① と Q4① の両方を選んだ回答者を数える。
- 縦軸・横軸には、設問のほかにエリア名・支所名(081 の B・C 列の名前)も書ける。
- 出力は `052_batch_analysis_master_cross<設定 CSV の名前の後半>.xlsx`。シートは縦軸ごと。

---

## 4. 既知の点・今後

- ダミーでは、`sample\batch\082_batch_processing_cross_011_shisho.csv`(縦軸が支所名)からシートが作られない(0 シート)。
  エリア名・支所名の照合がダミーで成立していない可能性がある(未確認)。設問 × 設問の `021_koyou` ではシートが作られる。
- 公開版 080 は設問名が空欄のため、051 の「その他」の自由記述の補正は、ダミーでは対象の列が見つからず何もしない。
- 今後:調査データの分析(060〜064、PCA・Ward 法)、ダッシュボード 053 のひな形。
  農協年鑑のデータを使う分析(070〜089)は、データを公開できないため含めない予定。

---

## ライセンス

MIT License(詳細は `LICENSE.md`)

Copyright (c) 2026 HIROI Takayuki
