# Plan (Implementation Roadmap)

`Spec.md` に基づいた実装計画。

## Phase 1: Project Initialization & POC

開発環境のセットアップと、コア技術（分割・リンク書き換え）の実現性検証を行う。

- [×] **1.1. Project Setup**
  - [×] `poetry init` && `pyproject.toml` 作成
  - [×] 依存ライブラリ追加 (`openpyxl`, `click`) → `poetry add`
  - [×] Git リポジトリ初期化・`.gitignore` 設定 (Mac/Python用)
  - [×] `src/` ディレクトリと基本パッケージ構造作成
- [×] **1.2. Core Logic Prototype (Delete Other Sheets)**
  - [×] 入力 `.xlsx` を読み込み、シート名を列挙するスクリプト作成。
  - [×] "Delete Other Sheets" 方式で、特定シートだけを残して別名保存する関数実装 (`workbook_splitter.py`)。
  - [×] **ユニットテスト**: スタイル（色、フォント、結合セル）が維持されているか確認。
  - [×] **テストデータ**: シンプルな 3 シート `.xlsx` ファイルを作成。
- [×] **1.3. Dependency: 1.1 → 1.2 | 1.4 (以降の全フェーズ)**
- [×] **1.4. ファイル名サニタイズ関数**
  - [×] 禁止文字 (`/`, `\\`, `:`, `*`, `?`, `"`, `<`, `>`, `|`) を `_` に置換。
  - [×] セパレーター `__SHEET__` でシート名と part 分割を区別するロジック。

## Phase 2: Hyperlink Handling (MVP Core)

最大の難所であるハイパーリンク書き換えロジックを実装する。

**難所注記**: 複雑な数式内のシート参照（`=SUM(Sheet2!A1:A10)`）は v1.0 では対象外。ハイパーリンク (Hyperlink オブジェクト) のみを処理する。

- [×] **2.1. Link Parsing & Detection**
  - [×] セルから `Hyperlink` オブジェクトを抽出する関数 (`extract_hyperlinks.py`)。
  - [×] リンク先文字列を解析し、内部参照 (`#SheetName!A1`) かどうかを判定する正規表現実装。
  - [×] **ユニットテスト**: 内部リンク・外部リンク・URL が正しく分類されることを確認。
- [×] **2.2. Link Target Mapping**
  - [×] 参照先シート名から、分割後のファイル名 (`{base}__SHEET__{sheet}.xlsx`) をマッピングする関数。
  - [×] セパレーター `__SHEET__` 形式に対応。
- [×] **2.3. Link Rewriting Engine**
  - [×] `openpyxl` でハイパーリンクを External Link に更新する処理の実装 (`hyperlink_rewriter.py`)。
  - [×] `external:./filename.xlsx#SheetName!A1` 相対パス形式で記述。
  - [×] **ユニットテスト**: 書き換え後の Excel をプログラムで開き、リンク先が正確に更新されていることを確認。
- [×] **2.4. Dependency: 1.2 → 2.1, 2.2 | 2.3 (リンク書き換え前提)**

## Phase 3: Large Sheet Splitting (Optional Spec)

行数制限機能（`--max-rows`）の実装。行数は **ヘッダー除外** でカウント。

- [ ] **3.1. Row Counter & Strategy**
  - [ ] シートのデータ行数をカウント（ヘッダー 1 行を除く）。
  - [ ] `--max-rows` と比較し、分割が必要か判定するロジック。
- [ ] **3.2. Row Chunking & Header Replication**
  - [ ] ヘッダー（1 行目）をメモリ上に保存。
  - [ ] 指定行数ごとに新しい Workbook を作成し、ヘッダー + データチャンク を転記する処理 (`row_splitter.py`)。
  - [ ] ファイル名形式 `{base}__SHEET__{sheet}_PART1.xlsx`, `_PART2.xlsx` ...
- [ ] **3.3. Style Handling Decision**
  - [ ] 値 + セル書式（背景色・フォント・境界線）をコピー（完全な条件付き書式維持は v2.0 以降検討）。
  - [ ] **ユニットテスト**: 複数 part ファイルが生成され、各ファイルがヘッダーを持つことを確認。
- [ ] **3.4. Dependency: 1.4 → 3.1, 3.2**

## Phase 4: CLI Implementation & Refinement

`click` を導入し、ユーザーが利用可能な CLI ツールとして仕上げる。

- [ ] **4.1. CLI Skeleton**
  - [ ] `click` を用いたコマンド定義 (`cli.py`)。
  - [ ] オプション解析: `INPUT_FILE`, `-o/--output-dir` (default: `./dist`), `--max-rows` (default: 0), `--dry-run`, `--verbose`。
  - [ ] `--help` メッセージの整備。
- [ ] **4.2. Orchestration**
  - [ ] 1.2 (Workbook Splitting) → 2.3 (Link Rewriting) → 3.2 (Row Splitting) を順序立てて呼び出すメインロジック (`orchestrator.py`)。
  - [ ] `--dry-run` オプション時は、分割予定を表示するが実ファイルは書き込まない。
  - [ ] `--verbose` オプション時は、ステップごとのログ出力。
- [ ] **4.3. Robustness & Error Handling**
  - [ ] 出力ディレクトリの自動作成 (`os.makedirs`)。
  - [ ] 入力ファイル不在時の親切なエラーメッセージ。
  - [ ] `.xlsx` 拡張子チェック。
  - [ ] 出力ディレクトリへの書き込み権限確認。
  - [ ] Windows ロック競合エラーの処理。
- [ ] **4.4. Dependency: 1.4, 2.3, 3.3 → 4.1, 4.2**

## Phase 5: Integration Testing & Final Verification

統合テストと品質担保。

- [ ] **5.1. Comprehensive Test Data**
  - [ ] シンプルケース: 2 シート、シンプルなデータ、ハイパーリンク 1 つ。
  - [ ] 複雑ケース: 結合セル、日本語シート名 (`営業`, `経費報告`)、色付きセル、複数ハイパーリンク、`--max-rows` 分割が必要なデータ。
  - [ ] エッジケース: シート名に禁止文字 (`A/B`, `C:D`)、ハイパーリンク先がないシート、存在しないシートへのリンク。
- [ ] **5.2. Integration Test Suite**
  - [ ] ユニットテストの統合実行 (pytest)。
  - [ ] 各 Phase のモジュール間の連携確認。
  - [ ] **期待結果**: テストケースすべてが PASS。
- [ ] **5.3. Manual End-to-End Verification**
  - [ ] CLI 実行: `excel-splitter complex_test.xlsx -o output/ --max-rows 500 --verbose`
  - [ ] 生成ファイルが Excel で開けるか確認。
  - [ ] ハイパーリンクをクリックして、正しいファイル・位置に飛ぶか確認。
  - [ ] `--dry-run` 実行時のログ出力が適切か確認。
- [ ] **5.4. Dependency: 4.3 → 5.1, 5.2, 5.3**

---

## Phase 6: Documentation & Release Prep

外部ドキュメント整備と配布準備。

- [ ] **6.1. README.md 更新**
  - [ ] プロジェクト説明、インストール方法 (`pip install -e .` or `poetry install`)。
  - [ ] **Usage** セクション: CLI 基本操作、各オプション説明、実行例。
  - [ ] **Examples** セクション: コマンド実行例（シート分割、行数制限、ハイパーリンク書き換え）。
  - [ ] **Known Limitations**: 数式内参照は v1.0 では対象外、etc.
- [ ] **6.2. CHANGELOG.md 作成**
  - [ ] v1.0.0 の機能リスト、既知の制限事項。
- [ ] **6.3. Package Metadata**
  - [ ] `pyproject.toml` に author, description, license, keywords 追加。
  - [ ] バージョン: `1.0.0`。
- [ ] **6.4. Dependency: 5.3 → 6.1, 6.2**

---

## Dependency Graph (Summary)

```
1.1 (Setup)
  └─→ 1.2 (Workbook Splitter)
       └─→ 1.4 (Sanitize)
            ├─→ 2.1 (Link Parsing)
            │    └─→ 2.2 (Link Mapping)
            │         └─→ 2.3 (Link Rewriting)
            │              └─→ 4.2 (Orchestration)
            ├─→ 3.1 (Row Counter)
            │    └─→ 3.2 (Row Chunking)
            │         └─→ 4.2
            └─→ 4.1 (CLI)
                 └─→ 4.2
                      └─→ 4.3 (Error Handling)
                           └─→ 5.2 (Integration Test)
                                └─→ 5.3 (End-to-End)
                                     └─→ 6.1 (Documentation)
```

---

## Estimated Task Breakdown

| Phase | 推定時間 | 優先度 | 難度 |
| :--- | :---: | :---: | :---: |
| 1 (Setup & POC) | 2-3h | 🔴 High | Low |
| 2 (Hyperlink) | 4-5h | 🔴 High | **Very High** |
| 3 (Row Splitting) | 2-3h | 🟡 Medium | Medium |
| 4 (CLI) | 2h | 🟡 Medium | Low |
| 5 (Integration Test) | 2-3h | 🟡 Medium | Medium |
| 6 (Documentation) | 1h | 🟢 Low | Low |
| **合計** | **13-17h** | - | - |
