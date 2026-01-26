import re

# Separators defined in Spec
SHEET_SEPARATOR = "__SHEET__"
PART_SEPARATOR = "_PART"

def sanitize_filename(name: str) -> str:
    """
    Sanitize a string to be safe for use as a filename.
    Replaces restricted characters with underscores.
    Restricted: / \\ : * ? " < > |
    Also replaces the separator markers to avoid ambiguity.
    """
    # Windows/Unix restricted characters
    # / \ : * ? " < > | and Space
    cleaned = re.sub(r'[\\/:*?"<>| ]', '_', name)
    
    # Replace separator markers to avoid confusion
    # This prevents sheet names like "Data__SHEET__2" from becoming ambiguous
    cleaned = cleaned.replace(SHEET_SEPARATOR, '_SHEET_')
    cleaned = cleaned.replace(PART_SEPARATOR, '_PART_')
    
    return cleaned

def get_sheet_filename(base_name: str, sheet_name: str) -> str:
    """
    Generate the filename for a split sheet.
    Format: {base}__SHEET__{sanitized_sheet}.xlsx
    """
    sanitized = sanitize_filename(sheet_name)
    return f"{base_name}{SHEET_SEPARATOR}{sanitized}.xlsx"

def get_part_filename(base_name: str, sheet_name: str, part_num: int) -> str:
    """
    Generate the filename for a row-split part.
    Format: {base}__SHEET__{sanitized_sheet}_PART{N}.xlsx
    """
    sanitized = sanitize_filename(sheet_name)
    return f"{base_name}{SHEET_SEPARATOR}{sanitized}{PART_SEPARATOR}{part_num}.xlsx"
