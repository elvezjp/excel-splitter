# excel-splitter Specification

## 1. 目的
業務で利用される巨大・多シート Excel ファイルを、
構造とリンクを極力維持したまま分割することを目的とする。

## 2. 対象ファイル
- .xlsx（.xlsm は非対応）

## 3. 機能仕様

### 3.1 Workbook Split
- 各シートを 1 ファイルとして出力
- シート名を元にファイル名を生成

### 3.2 Sheet Split（オプション）
- --max-rows 指定時のみ有効
- 行数ベースで分割
- {SheetName}_part01.xlsx の形式で命名

### 3.3 ハイパーリンク
- セルに設定されたハイパーリンクのみ対象
- #Sheet!A1 → Sheet.xlsx#Sheet!A1 に変換

## 4. 制約
- openpyxl 使用
- 図形・ボタンのリンクは対象外
