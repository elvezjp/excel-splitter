# Changelog

All notable changes to this project will be documented in this file.

## [1.0.0] - 2026-01-26

### Added

- **Workbook Splitter**: Split a multi-sheet Excel file into individual `.xlsx` files per sheet.
- **Hyperlink Rewriter**: Detects internal links (e.g., `#Sheet2!A1`) and rewrites them to point to the newly created external files (e.g., `external:./Base__SHEET__Sheet2.xlsx#Sheet2!A1`).
- **Row Splitter**: Support for splitting large sheets into multiple files based on `--max-rows`.
- **CLI**: Robust command-line interface using `click` with options for output directory and dry-run.
- **Safety**: "Delete Other Sheets" strategy to preserve original styling and formatting as much as possible.
