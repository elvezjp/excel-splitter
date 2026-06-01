# Contributing to excel-splitter

[English](./CONTRIBUTING.md) | [日本語](./CONTRIBUTING_ja.md)

Thank you for your interest in contributing to excel-splitter. This document explains how you can contribute.

## How to Contribute

### Reporting Bugs

If you find a bug, please create a GitHub Issue.

### Proposing Improvements

If you have ideas for new features or improvements, please propose them in a GitHub Issue.

### Pull Requests

Code changes and new features are submitted via Pull Requests.

## Required Information for Bug Reports

When creating an Issue, please include the following:

- **Clear and descriptive title**: A title that succinctly conveys the problem
- **Reproduction steps**: Concrete steps to reproduce the issue
- **Expected behavior**: How the tool should behave
- **Actual behavior**: What actually happened
- **Sample file**: If possible, an Excel file that reproduces the issue
- **Version information**:
  - excel-splitter version
  - Python version
  - OS and version

### Bug Report Example

```
Title: Hyperlinks are not rewritten correctly for Japanese sheet names

Reproduction steps:
1. Create an Excel file containing a Japanese-named sheet (e.g., "売上データ")
2. Add a hyperlink from another sheet to the Japanese-named sheet
3. Run excel-splitter to split the file

Expected behavior:
The hyperlinks correctly reference the split files.

Actual behavior:
The hyperlinks are broken and clicking them produces an error.

Version information:
- excel-splitter: 0.1.0
- Python: 3.12.0
- OS: macOS 14.0
```

## Required Information for Feature Proposals

When creating a feature proposal Issue, please include:

- **Clear and descriptive title**: A title that succinctly conveys the proposed feature
- **Detailed feature description**: What the feature does and how it works
- **Use cases and benefits**: Why the feature is needed and what problems it solves
- **Related examples or mockups**: Concrete usage examples or screen mockups, if possible

## Pull Request Procedure

### 1. Fork the Repository and Create a Branch

```bash
# Fork the repository (from the GitHub UI)

# Clone your fork
git clone https://github.com/your-username/excel-splitter.git
cd excel-splitter

# Create a new branch
git checkout -b {username}/{YYYYMMDD}-{description}
```

**Branch naming convention**: `{username}/{YYYYMMDD}-{description}`

Examples:
- `taro/20260127-fix-hyperlink-bug`
- `hanako/20260128-add-csv-export`

### 2. Follow the Coding Style

Python code should follow the PEP 8 style guidelines.

### 3. Write and Run Tests

Add tests for new features and bug fixes.

```bash
# Run tests
uv run pytest tests/ -v
```

### 4. Update Documentation

Update README.md and related documentation as needed.

### 5. Commit Message Style

Write commit messages in English or Japanese, clearly describing the change.

### 6. Push and Open a PR

```bash
# Push your changes
git push origin {branch-name}

# Open a Pull Request from the GitHub UI
```

### 7. Review and Feedback

- Respond to review comments promptly
- Update your code and push additional commits as needed

### Pre-submission Checklist

- [ ] All tests pass
- [ ] Tests have been added for new functionality
- [ ] Documentation has been updated (if necessary)
- [ ] Commit messages are clear

## Development Environment Setup

### Prerequisites

- Python 3.10 or higher
- [uv](https://docs.astral.sh/uv/) package manager

### Installation

```bash
# Clone the repository
git clone https://github.com/elvezjp/excel-splitter.git
cd excel-splitter

# Install uv (if not already installed)
curl -LsSf https://astral.sh/uv/install.sh | sh

# Install dependencies (including development)
uv sync --extra dev
```

## Running Tests

```bash
# Run all tests
uv run pytest tests/ -v

# Run a specific test file
uv run pytest tests/test_utils.py -v

# Run a specific test case
uv run pytest tests/test_utils.py::TestSanitizeFilename -v
```

## Coding Guidelines

### Style Guide

- Follow [PEP 8](https://peps.python.org/pep-0008/)
- Use 4 spaces for indentation
- Maximum line length is 88 characters (Black default)

### Naming Conventions

- Functions and variables: `snake_case`
- Classes: `PascalCase`
- Constants: `UPPER_SNAKE_CASE`

### Documentation

- Write docstrings for public functions
- Add comments for complex logic

```python
def split_workbook_by_sheet(input_file: str, output_dir: str) -> list[str]:
    """
    Split an Excel workbook by sheet.

    Args:
        input_file: Path to the input Excel file
        output_dir: Path to the output directory

    Returns:
        List of generated file paths
    """
    ...
```

## Commit Message Rules

### Format

```
{change type}: {summary of the change}

{detailed description (if needed)}

{related Issue references (if applicable)}
```

### Good Examples

```
fix: correct hyperlink rewriting for Japanese sheet names

When sheet names contained multibyte characters, encoding was not
handled correctly. This commit fixes that issue.

Fixes #123
```

```
feat: add CSV export

Add a --format csv option to support CSV output in addition to Excel.
```

### Bad Examples

```
fix  # Unclear what was fixed
```

```
misc fixes  # Lacks specificity
```

## Release Procedure

### When to Bump the Version

Bump the version when the repository has meaningful changes, such as new features, bug fixes, or major documentation additions or updates.
Do not bump the version for dependency updates only, such as security patches from Dependabot.
In that case, record the change under `[Unreleased]` and include it together with the next release that contains meaningful changes.

When bumping the version, update the following files consistently:

- `pyproject.toml`: update the `version` value
- `CHANGELOG.md` and `CHANGELOG_ja.md`: move `[Unreleased]` changes to the release version and date
- `SECURITY.md` and `SECURITY_ja.md`: update the supported version range if it changes

### How to Create a Tag

After the version update commit has been merged into `main`, create a tag on that commit and push it to the remote repository.

```bash
git tag -a v1.0.0 -m "Release v1.0.0"
git push origin v1.0.0
```

## Where to Reach Out

- **General questions**: Open a GitHub Issue with the `question` label
- **Feature proposals**: Open a GitHub Issue with the `enhancement` label
- **Bug reports**: Open a GitHub Issue with the `bug` label

## References

- [GitHub Contributing Guidelines](https://github.com/github/docs/blob/main/CONTRIBUTING.md)
- [PEP 8 -- Style Guide for Python Code](https://peps.python.org/pep-0008/)
