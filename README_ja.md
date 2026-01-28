# excel-splitter

[English](./README.md) | [日本語](./README_ja.md)

[![Elvez](https://img.shields.io/badge/Elvez-Product-3F61A7?style=flat-square)](https://elvez.co.jp/)
[![IXV Ecosystem](https://img.shields.io/badge/IXV-Ecosystem-3F61A7?style=flat-square)](https://elvez.co.jp/ixv/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.10+-blue?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![Stars](https://img.shields.io/github/stars/elvezjp/excel-splitter?style=social)](https://github.com/elvezjp/excel-splitter/stargazers)

Excel ファイルを分割するための Python CLI ツール。複数シートを個別ファイルに分割し、シート間ハイパーリンクを外部リンクへ自動変換します。

## ユースケース

- **大規模 Excel の分割**: 巨大な Excel ファイルをシート単位・行数単位で分割
- **ファイル共有の効率化**: 必要なシートだけを抽出して共有
- **データ移行**: システム間でのデータ移行時に Excel を分割
- **バックアップ**: シートごとに個別ファイルとして保存

## 開発の背景

本ツールは、日本語の開発文書・仕様書を対象とした開発支援AI **IXV（イクシブ）** の開発過程で生まれた小さな実用品です。

IXVでは、システム開発における日本語の文書について、理解・構造化・活用という課題に取り組んでおり、本リポジトリでは、その一部を切り出して公開しています。

## 特徴

- **シート分割**: 複数シートを含む Excel を、1シート1ファイルに分割
- **行数分割**: `--max-rows` オプションで巨大シートを複数ファイルに分割
- **ハイパーリンク維持**: シート間リンクを分割後の外部ファイルリンクに自動変換
- **書式維持**: 「他シート削除」方式により、元のスタイル・書式を可能な限り保持
- **ドライラン**: `--dry-run` で実際の書き込みなしに分割結果をプレビュー

## ドキュメント

- [CHANGELOG.md](CHANGELOG.md) - 変更履歴
- [CONTRIBUTING.md](CONTRIBUTING.md) - コントリビューション方法
- [SECURITY.md](SECURITY.md) - セキュリティポリシー
- [spec.md](spec.md) - 技術仕様書

## セットアップ

### 必要環境

- Python 3.10 以上
- [uv](https://docs.astral.sh/uv/) パッケージマネージャー

### uv のインストール

```bash
# macOS / Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows (PowerShell)
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### 依存ライブラリのインストール

```bash
git clone https://github.com/elvezjp/excel-splitter.git
cd excel-splitter
uv sync
```

## 使い方

```bash
uv run excel-splitter [OPTIONS] INPUT_FILE
```

### 最小限の使用例

```bash
# シートごとに分割（出力先: ./dist）
uv run excel-splitter input.xlsx
```

### その他の使用例

```bash
# 出力先ディレクトリを指定
uv run excel-splitter input.xlsx -o output/

# 行数制限付き分割（1シートあたり最大 50000 データ行）
uv run excel-splitter input.xlsx --max-rows 50000 -o output/

# ドライラン（ファイルを生成せずに分割結果をプレビュー）
uv run excel-splitter input.xlsx --dry-run

# 詳細ログ付き実行
uv run excel-splitter input.xlsx --verbose -o output/
```

### サンプルファイルで試す

動作確認用の Excel ファイルを用意しています。このファイルには複数シート、スタイル、ハイパーリンクが含まれており、主要機能を一通り確認できます。

```bash
uv run excel-splitter tests/fixtures/manual_test.xlsx -o dist --max-rows 50
```

### 注意事項

- **別シート参照**: 数式（`=Sheet2!A1`）やグラフのデータソースで別シートを参照している場合、分割後に参照先が失われます
- **行分割時の図形・画像**: `--max-rows` による行分割が発生したシートでは、Shape（図形）や画像は失われます

## 主要オプション

| オプション | デフォルト | 説明 |
|:---|:---:|:---|
| `INPUT_FILE` | - | 分割対象の `.xlsx` ファイルパス（必須） |
| `-o`, `--output-dir` | `./dist` | 出力先ディレクトリ（存在しない場合は自動作成） |
| `--max-rows` | `0`（無効） | 1シートの最大データ行数（ヘッダー除外）。超過時に Part 分割 |
| `--dry-run` | `False` | 実際の書き込みを行わず、分割予定を表示 |
| `--verbose` | `False` | 詳細なログ出力を有効化 |

## 出力例

### シート分割

入力: `report.xlsx`（Sheet1, Sheet2, Sheet3 を含む）

```
dist/
├── report__SHEET__Sheet1.xlsx
├── report__SHEET__Sheet2.xlsx
└── report__SHEET__Sheet3.xlsx
```

### 行数分割（--max-rows 指定時）

入力: `large_data.xlsx`（10万行のデータを含む Data シート）

```bash
uv run excel-splitter large_data.xlsx --max-rows 50000
```

```
dist/
├── large_data__SHEET__Data_PART1.xlsx  # 1〜50000行
└── large_data__SHEET__Data_PART2.xlsx  # 50001〜100000行
```

## ディレクトリ構成

```
excel-splitter/
├── src/
│   └── excel_splitter/
│       ├── __init__.py
│       ├── cli.py           # CLI エントリーポイント
│       ├── splitter.py      # ワークブック分割
│       ├── row_splitter.py  # 行数分割
│       ├── hyperlinks.py    # ハイパーリンク処理
│       └── utils.py         # ユーティリティ
├── tests/                   # テストコード
├── docs/                    # ドキュメント
├── pyproject.toml
├── README.md
├── CHANGELOG.md
├── CONTRIBUTING.md
├── SECURITY.md
├── spec.md
└── LICENSE
```

## 制約事項

- `.xlsx` 形式のみ対応（`.xlsm` マクロ付きファイルは非対応）
- Excel の数式内シート参照（例: `=SUM(Sheet2!A1:A10)`）の書き換えは対象外（ハイパーリンクのみ）
- 条件付き書式の完全な維持は保証外
- Shape（図形）・画像は、行分割（`--max-rows`）が発生したシートでは失われます

## セキュリティ

セキュリティに関する詳細は [SECURITY.md](SECURITY.md) を参照してください。

- 入力ファイルは信頼できるソースからのものに限定してください
- 出力ディレクトリへの書き込み権限を適切に設定してください

## コントリビューション

コントリビューションを歓迎します。詳細は [CONTRIBUTING.md](CONTRIBUTING.md) を参照してください。

- バグ報告: [GitHub Issues](https://github.com/elvezjp/excel-splitter/issues)
- 機能提案: [GitHub Issues](https://github.com/elvezjp/excel-splitter/issues)
- プルリクエスト: [GitHub Pull Requests](https://github.com/elvezjp/excel-splitter/pulls)

## 変更履歴

詳細は [CHANGELOG.md](CHANGELOG.md) を参照してください。

## ライセンス

MIT License - 詳細は [LICENSE](LICENSE) を参照してください。

## 問い合わせ先

- **GitHub Issues**: [https://github.com/elvezjp/excel-splitter/issues](https://github.com/elvezjp/excel-splitter/issues)
- **メールアドレス**: info@elvez.co.jp
- **宛先**: 株式会社エルブズ
