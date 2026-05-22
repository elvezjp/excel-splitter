# Security Policy

[English](./SECURITY.md) | [日本語](./SECURITY_ja.md)

## Supported Versions

Security updates are provided for the following versions. We recommend using the latest version.

| Version | Supported          |
| ------- | ------------------ |
| 0.1.x   | :white_check_mark: |
| < 0.1   | :x:                |

## How to Report a Vulnerability

If you discover a security vulnerability, **please do not open a public Issue**.

### Reporting Methods

Please use one of the following methods:

1. **GitHub Security Advisories (recommended)**
   - Create a private report from the repository's [Security Advisories](https://github.com/elvezjp/excel-splitter/security/advisories/new).

2. **Email**
   - Contact the security team directly by email.

### What to Include in Your Report

- **Vulnerability description**: A summary of the issue and technical details
- **Reproduction steps**: Concrete steps to reproduce the issue
- **Potential impact and severity**: What impact the vulnerability could have and how severe it is
- **Suggested fix or mitigation**: A proposed fix or mitigation, if available
- **Contact information**: Your contact for follow-up (optional)

### Report Example

```
Subject: [SECURITY] Path traversal vulnerability via malicious Excel file

Description:
When processing certain Excel files, files may be created outside the
specified output directory.

Reproduction steps:
1. Create an Excel file with a sheet name containing "../"
2. Process the file with excel-splitter
3. Files are created outside the output directory

Impact:
An attacker may be able to create files in arbitrary directories.
Severity: High

Suggested fix:
Strip path-traversal sequences during sheet-name sanitization.
```

## Response Schedule

- **Initial response**: Within 48 hours
- **Status updates**: Within 7 days
- **Resolution**: Based on severity
  - Critical: Within 14 days
  - High: Within 30 days
  - Medium: Within 60 days
  - Low: Next release cycle

## Security Considerations

### File Processing

- Process only files from trusted sources
- Set appropriate access permissions on the output directory
- Verify file size before processing; be cautious with extremely large files

### Input Validation

- Only `.xlsx` files are processed
- Forbidden characters in sheet names (`/`, `\`, `:`, `*`, `?`, `"`, `<`, `>`, `|`) are sanitized
- Write permissions on the output directory are checked

### Output Security

- Output filenames are sanitized to prevent path-traversal attacks
- Output is written only inside the specified directory

### Dependency Security

- Dependencies (`openpyxl`, `click`) are updated regularly
- `uv` is used to manage dependency integrity

### Dependabot Alerts Policy

**Malware tab**: Always remediate, regardless of where it appears.

**Vulnerable**: Follow the table below.

| Target | Action |
|--------|--------|
| Latest version | **Remediate** (update dependency / open PR). When the fix is possible via a dependency version bump alone, do not change the project version. |

## Security Best Practices

1. **Use the latest version**: Always use the latest version of excel-splitter
2. **Validate input files**: Inspect Excel files from untrusted sources before processing
3. **Run in a sandbox**: Process untrusted files in an isolated environment
4. **Verify the output**: Check that generated files are created in the expected location
5. **Restrict permissions**: Grant only the minimum necessary permissions on the output directory

## Known Security Limitations

- Macro-enabled Excel files (`.xlsm`) are not processed
- Encrypted Excel files are not supported
- Memory usage during processing depends on the input file size

## Other Security Questions

For non-vulnerability security questions:

- Open a GitHub Issue with the `security` label
- General questions about security best practices are also welcome

## References

- [GitHub Security Advisories](https://docs.github.com/en/code-security/security-advisories)
- [OWASP Top 10](https://owasp.org/www-project-top-ten/)
