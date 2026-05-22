# Contributing to excel-splitter

[English](./CONTRIBUTING.md) | [日本語](./CONTRIBUTING_ja.md)

excel-splitter への貢献に興味を持っていただきありがとうございます。このドキュメントでは、貢献の方法について説明します。

## 貢献の方法

### バグの報告

バグを見つけた場合は、GitHub Issue を作成してください。

### 機能改善の提案

新機能や改善のアイデアがある場合は、GitHub Issue で提案してください。

### プルリクエスト

コードの修正や新機能の追加は、プルリクエストを通じて行います。

## バグ報告の必須情報

Issue を作成する際は、以下の情報を含めてください：

- **明確で説明的なタイトル**: 問題を端的に表すタイトル
- **再現手順**: 問題を再現するための具体的な手順
- **期待される動作**: 本来どのように動作すべきか
- **実際の動作**: 実際にどのような動作をしたか
- **サンプルファイル**: 可能であれば、問題を再現できる Excel ファイル
- **バージョン情報**:
  - excel-splitter のバージョン
  - Python のバージョン
  - OS とバージョン

### バグ報告の例

```
タイトル: ハイパーリンクが日本語シート名で正しく書き換えられない

再現手順:
1. 日本語名のシート（例: "売上データ"）を含む Excel ファイルを作成
2. 他のシートから日本語シートへのハイパーリンクを設定
3. excel-splitter で分割を実行

期待される動作:
ハイパーリンクが分割後のファイルを正しく参照する

実際の動作:
ハイパーリンクが壊れ、クリックしてもエラーが発生する

バージョン情報:
- excel-splitter: 0.1.0
- Python: 3.12.0
- OS: macOS 14.0
```

## 機能改善提案の必須情報

機能提案の Issue には以下を含めてください：

- **明確で説明的なタイトル**: 提案する機能を端的に表すタイトル
- **機能の詳細な説明**: どのような機能か、どのように動作するか
- **ユースケースとメリット**: なぜこの機能が必要か、どのような問題を解決するか
- **関連する例やモックアップ**: 可能であれば、具体的な使用例や画面イメージ

## プルリクエストの手順

### 1. リポジトリのフォークとブランチ作成

```bash
# リポジトリをフォーク（GitHub UI から）

# フォークしたリポジトリをクローン
git clone https://github.com/your-username/excel-splitter.git
cd excel-splitter

# 新しいブランチを作成
git checkout -b {ユーザー名}/{日付YYYYMMDD}-{内容}
```

**ブランチ命名規則**: `{ユーザー名}/{日付YYYYMMDD}-{内容}`

例:
- `taro/20260127-fix-hyperlink-bug`
- `hanako/20260128-add-csv-export`

### 2. コーディングスタイルへの準拠

Python コードは PEP 8 スタイルガイドラインに従ってください。

### 3. テストの作成と実行

新機能やバグ修正には、対応するテストを追加してください。

```bash
# テストの実行
uv run pytest tests/ -v
```

### 4. ドキュメントの更新

必要に応じて、README.md や関連ドキュメントを更新してください。

### 5. コミットメッセージの書き方

コミットメッセージは日本語または英語で、変更内容を明確に記述してください。

### 6. プッシュと PR 送信

```bash
# 変更をプッシュ
git push origin {ブランチ名}

# GitHub UI から Pull Request を作成
```

### 7. レビュー待ちとフィードバック対応

- レビューコメントには迅速に対応してください
- 必要に応じてコードを修正し、追加コミットを行ってください

### PR 送信前のチェックリスト

- [ ] テストがすべてパスする
- [ ] 新機能にはテストを追加した
- [ ] ドキュメントを更新した（必要な場合）
- [ ] コミットメッセージが明確である

## 開発環境のセットアップ

### 前提条件

- Python 3.10 以上
- [uv](https://docs.astral.sh/uv/) パッケージマネージャー

### インストール手順

```bash
# リポジトリをクローン
git clone https://github.com/elvezjp/excel-splitter.git
cd excel-splitter

# uv のインストール（未インストールの場合）
curl -LsSf https://astral.sh/uv/install.sh | sh

# 依存関係のインストール（開発用含む）
uv sync --extra dev
```

## テストの実行方法

```bash
# すべてのテストを実行
uv run pytest tests/ -v

# 特定のテストファイルを実行
uv run pytest tests/test_utils.py -v

# 特定のテストケースを実行
uv run pytest tests/test_utils.py::TestSanitizeFilename -v
```

## コーディングガイドライン

### スタイルガイド

- [PEP 8](https://peps.python.org/pep-0008/) に準拠
- インデントは 4 スペース
- 行の最大長は 88 文字（Black のデフォルト）

### 命名規則

- 関数名・変数名: `snake_case`
- クラス名: `PascalCase`
- 定数: `UPPER_SNAKE_CASE`

### ドキュメント

- 公開関数には docstring を記述
- 複雑なロジックにはコメントを追加

```python
def split_workbook_by_sheet(input_file: str, output_dir: str) -> list[str]:
    """
    Excel ワークブックをシートごとに分割する。

    Args:
        input_file: 入力 Excel ファイルのパス
        output_dir: 出力ディレクトリのパス

    Returns:
        生成されたファイルパスのリスト
    """
    ...
```

## コミットメッセージのルール

### 形式

```
{変更の種類}: {変更の概要}

{詳細な説明（必要に応じて）}

{関連 Issue への参照（必要に応じて）}
```

### 良い例

```
fix: 日本語シート名のハイパーリンク書き換えを修正

シート名にマルチバイト文字が含まれる場合、エンコーディングが
正しく処理されていなかった問題を修正。

Fixes #123
```

```
feat: CSV エクスポート機能を追加

--format csv オプションを追加し、Excel だけでなく CSV 形式での
出力をサポート。
```

### 悪い例

```
修正  # 何を修正したか不明
```

```
いろいろ直した  # 具体性がない
```

## 問い合わせ先

- **一般的な質問**: GitHub Issue を作成し、`question` ラベルを付けてください
- **機能提案**: GitHub Issue を作成し、`enhancement` ラベルを付けてください
- **バグ報告**: GitHub Issue を作成し、`bug` ラベルを付けてください

## 参考リンク

- [GitHub Contributing Guidelines](https://github.com/github/docs/blob/main/CONTRIBUTING.md)
- [PEP 8 -- Style Guide for Python Code](https://peps.python.org/pep-0008/)
