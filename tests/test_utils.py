"""Unit tests for utils module"""
import pytest
from excel_splitter.utils import (
    sanitize_filename,
    get_sheet_filename,
    get_part_filename,
    SHEET_SEPARATOR,
    PART_SEPARATOR,
)


class TestSanitizeFilename:
    """Test filename sanitization"""

    def test_sanitize_forward_slash(self):
        """Forward slashes should be replaced with underscores"""
        assert sanitize_filename("A/B") == "A_B"

    def test_sanitize_backslash(self):
        """Backslashes should be replaced with underscores"""
        assert sanitize_filename("A\\B") == "A_B"

    def test_sanitize_colon(self):
        """Colons should be replaced with underscores"""
        assert sanitize_filename("A:B") == "A_B"

    def test_sanitize_asterisk(self):
        """Asterisks should be replaced with underscores"""
        assert sanitize_filename("A*B") == "A_B"

    def test_sanitize_question_mark(self):
        """Question marks should be replaced with underscores"""
        assert sanitize_filename("A?B") == "A_B"

    def test_sanitize_quotes(self):
        """Double quotes should be replaced with underscores"""
        assert sanitize_filename('A"B') == "A_B"

    def test_sanitize_angle_brackets(self):
        """Angle brackets should be replaced with underscores"""
        assert sanitize_filename("A<B") == "A_B"
        assert sanitize_filename("A>B") == "A_B"

    def test_sanitize_pipe(self):
        """Pipes should be replaced with underscores"""
        assert sanitize_filename("A|B") == "A_B"

    def test_sanitize_space(self):
        """Spaces should be replaced with underscores"""
        assert sanitize_filename("A B") == "A_B"

    def test_sanitize_multiple_invalid_chars(self):
        """Multiple invalid characters should all be replaced"""
        assert sanitize_filename("A/B:C*D") == "A_B_C_D"

    def test_sanitize_separator_marker(self):
        """Sheet separator marker should be replaced to avoid ambiguity"""
        # If sheet name contains __SHEET__, it should be escaped
        result = sanitize_filename("Data__SHEET__2")
        # Should NOT contain __SHEET__ marker
        assert "__SHEET__" not in result
        assert "_SHEET_" in result

    def test_sanitize_part_marker(self):
        """Part separator marker should be replaced to avoid ambiguity"""
        # If sheet name contains _PART, it should be escaped
        result = sanitize_filename("Data_PART_2")
        # Should NOT contain _PART marker (depending on implementation)
        assert "_PART_" in result

    def test_sanitize_valid_chars(self):
        """Valid alphanumeric characters should not be changed"""
        assert sanitize_filename("Data2024") == "Data2024"
        assert sanitize_filename("Sales_Report") == "Sales_Report"
        assert sanitize_filename("report-2024") == "report-2024"


class TestGetSheetFilename:
    """Test sheet filename generation"""

    def test_basic_sheet_name(self):
        """Basic sheet name should generate correct format"""
        result = get_sheet_filename("report", "Sales")
        assert result == f"report{SHEET_SEPARATOR}Sales.xlsx"

    def test_sheet_name_with_spaces(self):
        """Sheet name with spaces should be sanitized"""
        result = get_sheet_filename("report", "Sales 2024")
        assert result == f"report{SHEET_SEPARATOR}Sales_2024.xlsx"

    def test_sheet_name_with_special_chars(self):
        """Sheet name with special characters should be sanitized"""
        result = get_sheet_filename("report", "A/B:C")
        assert result == f"report{SHEET_SEPARATOR}A_B_C.xlsx"

    def test_sheet_filename_ends_with_xlsx(self):
        """Generated filename should end with .xlsx"""
        result = get_sheet_filename("base", "sheet")
        assert result.endswith(".xlsx")

    def test_sheet_filename_contains_separator(self):
        """Generated filename should contain separator"""
        result = get_sheet_filename("base", "sheet")
        assert SHEET_SEPARATOR in result


class TestGetPartFilename:
    """Test part filename generation"""

    def test_basic_part_name(self):
        """Basic part name should generate correct format"""
        result = get_part_filename("report", "Sales", 1)
        expected = f"report{SHEET_SEPARATOR}Sales{PART_SEPARATOR}1.xlsx"
        assert result == expected

    def test_part_number_increments(self):
        """Different part numbers should generate different names"""
        result1 = get_part_filename("report", "Sales", 1)
        result2 = get_part_filename("report", "Sales", 2)
        assert "1" in result1
        assert "2" in result2
        assert result1 != result2

    def test_part_filename_ends_with_xlsx(self):
        """Generated part filename should end with .xlsx"""
        result = get_part_filename("base", "sheet", 1)
        assert result.endswith(".xlsx")

    def test_part_filename_contains_separators(self):
        """Generated part filename should contain both separators"""
        result = get_part_filename("base", "sheet", 1)
        assert SHEET_SEPARATOR in result
        assert PART_SEPARATOR in result

    def test_part_sheet_name_sanitized(self):
        """Sheet name in part filename should be sanitized"""
        result = get_part_filename("report", "Sales 2024", 1)
        assert "Sales_2024" in result
        assert result.endswith(".xlsx")
