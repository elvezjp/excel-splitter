import re
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl import Workbook
from typing import Optional, Dict

from .utils import get_sheet_filename, SHEET_SEPARATOR

# Regex to detect internal sheet links.
# Examples: "#Sheet2!A1", "'Sheet 2'!B5", "Sheet2!A1" (sometimes internal refs don't start with # in the logic, but Hyperlink targets usually do if they are location based)
# Actually, openpyxl hyperlink.target might look like "SheetName!A1" or "#SheetName!A1".
# We need to handle quoted sheet names: 'My Sheet'!A1
INTERNAL_LINK_PATTERN = re.compile(r"^#?('?)(.+?)\1!([A-Za-z0-9]+)$")

def is_internal_link(target: str) -> bool:
    """Check if the hyperlink target is an internal reference to another sheet."""
    if not target:
        return False
    return bool(INTERNAL_LINK_PATTERN.match(target))

def parse_internal_link(target: str) -> tuple[Optional[str], Optional[str]]:
    """
    Parse internal link to extract SheetName and CellAddress.
    Returns (sheet_name, cell_address) or (None, None).
    """
    match = INTERNAL_LINK_PATTERN.match(target)
    if match:
        # group 2 is sheet name, group 3 is cell address. 
        # CAUTION: if sheet name contained single quotes escaped, standard Excel is 'Sheet''s Name'!A1.
        # Our regex is simple: '(.+?)'. 
        sheet_name = match.group(2)
        cell_addr = match.group(3)
        # simplistic unescaping for '
        sheet_name = sheet_name.replace("''", "'")
        return sheet_name, cell_addr
    return None, None

def generate_external_link(base_name: str, target_sheet: str, cell_addr: str) -> str:
    """
    Generate external link string.
    Format: external:./{filename}#{sheet}!{cell}
    """
    filename = get_sheet_filename(base_name, target_sheet)
    # Relative path syntax for Excel external links
    return f"external:./{filename}#{target_sheet}!{cell_addr}"

def rewrite_hyperlinks_in_workbook(wb: Workbook, base_name: str, current_sheet_name: str, verbose: bool = False):
    """
    Iterate through all hyperlinks in the given single-sheet workbook and rewrite internal links.
    """
    ws = wb[current_sheet_name]
    
    # Access private attribute _hyperlinks as public property seems missing in openpyxl 3.1.5
    # Also, it seems sometimes the collection is not populated even if cells have links.
    # To be robust, we will iterate over used cells.
    
    # Strategy: iterate used range.
    has_rewritten = False
    
    # We can iterate iter_rows()
    for row in ws.iter_rows():
        for cell in row:
            if cell.hyperlink and cell.hyperlink.target:
                target = cell.hyperlink.target
                if verbose:
                    # Too noisy to print every cell, only print match
                    if is_internal_link(target):
                        print(f"DEBUG: Found internal link at {cell.coordinate}: {target}")
                
                if is_internal_link(target):
                    target_sheet, cell_addr = parse_internal_link(target)
                    
                    if target_sheet and target_sheet != current_sheet_name:
                        # Rewrite
                        new_target = generate_external_link(base_name, target_sheet, cell_addr)
                        if verbose:
                            print(f"  [Link Rewrite] {cell.coordinate}: {target} -> {new_target}")
                        cell.hyperlink.target = new_target
                        has_rewritten = True
    
    # Note: modifying cell.hyperlink.target is sufficient for openpyxl to save it correctly.


