# excel-splitter

Excel ファイルを分割するための Python CLI ツールです。

- 複数シートを **1ファイル1シート** に分割
- 必要に応じて **巨大な1シートを複数ファイルに分割**
- セルの **シート間ハイパーリンクを壊さず** に外部リンクへ変換

業務でよくある「巨大 Excel を安全に分割したい」というニーズに対応します。

---

## 主な機能

### 1. Workbook Split

1つの Excel（複数シート）を、シートごとに 1 ファイルずつに分割します。
元ファイルの書式を維持するため「Delete Other Sheets」方式を採用しています。

- 出力ファイル名: `{元ファイル名}__SHEET__{シート名}.xlsx`

### 2. Hyperlink Rewrite

ブック内リンク（`#OtherSheet!A1`）を、分割後も有効な外部ファイルリンクに変換します。

- 絶対参照（`$A$1`）・範囲参照（`A1:B10`）にも対応
- 範囲が複数 Part にまたがる場合は開始セルに縮退

### 3. Sheet Split（オプション）

`--max-rows` 指定時、巨大なシートを行数基準で複数ファイルに分割します。

- ヘッダー（1行目）は全 Part にコピーされます
- セルのスタイル（フォント・塗り・罫線・配置・表示形式）を維持
- 出力ファイル名: `{元ファイル名}__SHEET__{シート名}_PART{N}.xlsx`

---

## 動作環境

- Python 3.10+
- [uv](https://docs.astral.sh/uv/) (パッケージマネージャー)
- 依存ライブラリ: `openpyxl`, `click`

## インストール

### uv のインストール

```bash
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows (PowerShell)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### 依存ライブラリのインストール

```bash
uv sync
```

---

## 使い方

```
uv run excel-splitter [OPTIONS] INPUT_FILE
```

### 引数・オプション

| 引数/オプション | 必須 | デフォルト | 説明 |
|:---|:---:|:---:|:---|
| `INPUT_FILE` | Yes | - | 分割対象の `.xlsx` ファイルパス |
| `-o`, `--output-dir` | No | `./dist` | 出力先ディレクトリ（存在しない場合は自動作成） |
| `--max-rows` | No | `0`（無効） | 1シートの最大データ行数（ヘッダー除外）。超過時に Part 分割 |
| `--dry-run` | No | `False` | 実際の書き込みを行わず、分割予定を表示 |
| `--verbose` | No | `False` | 詳細なログ出力を有効化 |

### 実行例

```bash
# 基本: シートごとに分割（出力先: ./dist）
uv run excel-splitter input.xlsx

# 出力先ディレクトリを指定
uv run excel-splitter input.xlsx -o output/

# 行数制限付き分割（1シートあたり最大 50000 データ行）
uv run excel-splitter input.xlsx --max-rows 50000 -o output/

# ドライラン（ファイルを生成せずに分割結果をプレビュー）
uv run excel-splitter input.xlsx --dry-run

# 詳細ログ付き実行
uv run excel-splitter input.xlsx --verbose -o output/
```

---

## 開発

### テストの実行

```bash
uv sync --extra dev
uv run pytest tests/ -v
```

---

## 制約事項

- `.xlsx` 形式のみ対応（`.xlsm` マクロ付きファイルは非対応）
- Excel の数式内シート参照（例: `=SUM(Sheet2!A1:A10)`）の書き換えは対象外（ハイパーリンクのみ）
- 条件付き書式の完全な維持は保証外

---

## License

MIT License
