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

### 2. Hyperlink Rewrite
ブック内リンク（#OtherSheet!A1）を、分割後も有効な外部ファイルリンクに変換します。

### 3. Sheet Split（オプション）
巨大なシートを行数基準で複数ファイルに分割します。

---

## 使い方（例）

```bash
excel-splitter split input.xlsx -o out/
excel-splitter split input.xlsx --max-rows 50000 -o out/
```

---

## License
MIT License
