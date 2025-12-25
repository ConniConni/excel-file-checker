# Excel File Checker

指定したディレクトリから特定の条件に合致するExcel/CSVファイルを探索し、指定されたセルの値を抽出して整形出力するツールです。画像（電子捺印）の存在判定機能も備えています。

## 機能

- **再帰的ファイル探索**: 指定ディレクトリ配下を再帰的に探索
- **セル値抽出**: Excel (.xlsx, .xls) および CSV ファイルから指定セルの値を抽出
- **画像判定**: 指定セルに画像（電子捺印等）が存在するかを判定
- **整形出力**: カンマ位置が揃った視認性の高い出力フォーマット

## 必要要件

- Python 3.10以上
- 依存ライブラリ: pandas, openpyxl

## インストール

```bash
pip install -r requirements.txt
```

## 使い方

### 1. 設定ファイルの作成

設定ファイル（.ini）を作成します（`config/sample.ini` を参考）。

```ini
[SETTINGS]
# 探索を開始するルートディレクトリのパス
target_dir = ./input
# 抽出対象とするファイル名に含まれるキーワード
search_keyword = 日経平均
# 抽出したいセルの指定 (例: A1, B2 等) ※複数指定はカンマ区切り
target_cells = A1, B1, C1
# 画像判定対象セルの指定 (例: D1, E1 等) ※複数指定はカンマ区切り
image_check_cells = D1, E1
# 出力ファイル名
output_filename = extraction_result.txt
```

### 2. 実行

```bash
# 設定ファイルを指定して実行
python run.py -i config.ini

# または config/sample.ini を使用する場合
python run.py -i config/sample.ini

# -i オプションを省略した場合、カレントディレクトリのconfig.iniを使用
python run.py
```

**オプション:**
- `-i, --ini <ファイルパス>`: 設定ファイル（.ini）のパスを指定（デフォルト: config.ini）

### 3. 結果の確認

設定ファイルと同じディレクトリに指定した名前の結果ファイルが生成されます。

## 出力フォーマット例

```text
Filename             , A1        , B1      , C1  , D1(画像), E1(画像)
root_data.xlsx       , 2023/12/25, 33000.5 , 確定, ○      , ○
sub/folder1/data.csv , 2023/12/26, 33100.2 , 暫定, -      , -
sub/folder2/info.xlsx, 2023/12/27, 33200.0 , 完了, ○      , ×
```

**凡例:**
- `○`: 画像あり
- `×`: 画像なし
- `-`: 画像判定非対応（CSVファイル等）

## テスト

```bash
# 全テスト実行
pytest

# カバレッジ付きテスト実行
pytest --cov=src --cov-report=html

# HTMLレポート確認
open htmlcov/index.html
```

## プロジェクト構造

```
excel-file-checker/
├── config/
│   └── sample.ini          # 設定ファイルサンプル
├── src/
│   ├── config_loader.py    # 設定ファイル読み込み
│   ├── file_searcher.py    # ファイル探索
│   ├── cell_extractor.py   # セル値抽出
│   ├── image_checker.py    # 画像判定
│   ├── output_formatter.py # 出力整形
│   └── main.py             # メインプログラム
├── tests/
│   ├── fixtures/           # テスト用ファイル
│   └── test_*.py           # 各種テストコード
├── config.ini              # 設定ファイル（要作成）
├── run.py                  # 実行スクリプト
├── requirements.txt        # 依存ライブラリ
└── README.md               # このファイル
```

## 開発

このプロジェクトはテスト駆動開発（TDD）で開発されています。

- テストフレームワーク: pytest
- カバレッジ: pytest-cov
- コーディング規約: `CLAUDE.md` を参照
- 詳細仕様: `SPEC.md` を参照

## ライセンス

MIT License
