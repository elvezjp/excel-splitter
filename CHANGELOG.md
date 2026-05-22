# Changelog

[English](./CHANGELOG.md) | [日本語](./CHANGELOG_ja.md)

All notable changes to this project will be recorded in this file.

## [0.1.0] - 2026-01-27

### Added

- **Workbook splitting**: Split Excel files containing multiple sheets into individual `.xlsx` files, one per sheet
- **Hyperlink rewriting**: Detect internal links (e.g., `#Sheet2!A1`) and rewrite them as links to the split external files (e.g., `./Base__SHEET__Sheet2.xlsx#Sheet2!A1`)
- **Row splitting**: Support for splitting large sheets into multiple files via the `--max-rows` option
- **CLI**: Robust command-line interface built with `click` (output directory, dry-run, and more)
- **Format preservation**: Preserve original styles and formatting as much as possible by using the "delete other sheets" approach
