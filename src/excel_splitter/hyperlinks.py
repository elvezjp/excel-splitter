import re
from dataclasses import dataclass
from typing import Optional
from openpyxl import Workbook
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.utils.cell import coordinate_from_string, range_boundaries

from .utils import get_sheet_filename, get_part_filename

# Regex to detect internal sheet links.
# Examples: "#Sheet2!A1", "'Sheet 2'!B5", "Sheet2!A1" (sometimes internal refs don't start with # in the logic, but Hyperlink targets usually do if they are location based)
# Actually, openpyxl hyperlink.target might look like "SheetName!A1" or "#SheetName!A1".
# We need to handle quoted sheet names: 'My Sheet'!A1
INTERNAL_LINK_PATTERN = re.compile(r"^#?('?)(.+?)\1!(.+)$")

@dataclass(frozen=True)
class InternalLink:
    sheet_name: str
    ref: str
    start_row: int
    end_row: int
    is_range: bool

def quote_sheet_name(sheet_name: str) -> str:
    """
    Quote sheet names that contain spaces/special chars.
    Excel uses single quotes and escapes single quote by doubling it.
    """
    if re.search(r"[^A-Za-z0-9_]", sheet_name):
        escaped = sheet_name.replace("'", "''")
        return f"'{escaped}'"
    return sheet_name

def is_internal_link(target: str) -> bool:
    """Check if the hyperlink target is an internal reference to another sheet."""
    if not target:
        return False
    return bool(INTERNAL_LINK_PATTERN.match(target))

def split_internal_link_target(target: str) -> tuple[Optional[str], Optional[str]]:
    """
    Split internal link into (sheet_name, ref).
    Returns (None, None) if not internal link.
    """
    match = INTERNAL_LINK_PATTERN.match(target)
    if not match:
        return None, None
    sheet_name = match.group(2)
    ref = match.group(3)
    sheet_name = sheet_name.replace("''", "'")
    return sheet_name, ref

def parse_ref_rows(ref: str) -> tuple[Optional[int], Optional[int], bool]:
    """
    Parse A1-style cell ref or range to extract row bounds.
    Returns (start_row, end_row, is_range), or (None, None, False) if unsupported.
    """
    if not ref:
        return None, None, False
    if ":" in ref:
        try:
            _, start_row, _, end_row = range_boundaries(ref)
            return start_row, end_row, True
        except ValueError:
            return None, None, False
    try:
        _, row = coordinate_from_string(ref.replace("$", ""))
        return row, row, False
    except ValueError:
        return None, None, False

def parse_internal_link(target: str) -> Optional[InternalLink]:
    """
    Parse internal link to extract sheet, ref, and row bounds.
    Returns InternalLink or None.
    """
    sheet_name, ref = split_internal_link_target(target)
    if not sheet_name or not ref:
        return None
    start_row, end_row, is_range = parse_ref_rows(ref)
    if start_row is None or end_row is None:
        return None
    return InternalLink(sheet_name=sheet_name, ref=ref, start_row=start_row, end_row=end_row, is_range=is_range)

def extract_internal_link(hyperlink: Hyperlink) -> Optional[InternalLink]:
    """
    Extract internal link info from a hyperlink target/location.
    Returns InternalLink or None.
    """
    if not hyperlink:
        return None
    raw = hyperlink.target or hyperlink.location
    if not raw:
        return None
    return parse_internal_link(raw)

def generate_external_link(base_name: str, target_sheet: str, ref: str) -> str:
    """
    Generate external link string.
    Format: external:./{filename}#{sheet}!{ref}
    """
    filename = get_sheet_filename(base_name, target_sheet)
    # Relative path syntax for Excel external links
    sheet_ref = quote_sheet_name(target_sheet)
    return f"./{filename}#{sheet_ref}!{ref}"

def generate_part_external_link(base_name: str, target_sheet: str, ref: str, part_num: int) -> str:
    """
    Generate external link to a specific part file.
    Format: external:./{part_filename}#{sheet}!{ref}
    """
    part_filename = get_part_filename(base_name, target_sheet, part_num)
    sheet_ref = quote_sheet_name(target_sheet)
    return f"./{part_filename}#{sheet_ref}!{ref}"

def get_part_number_for_row(row: int, max_rows: int) -> int:
    """
    Convert an absolute row to part number based on max_rows.
    """
    if row < 1:
        return 1
    return ((row - 1) // max_rows) + 1

def get_range_part_decision(link: InternalLink, max_rows: int) -> tuple[int, str, bool]:
    """
    Return (target_part, ref_for_link, spans_parts)
    """
    start_part = get_part_number_for_row(link.start_row, max_rows)
    end_part = get_part_number_for_row(link.end_row, max_rows)
    if start_part == end_part:
        return start_part, link.ref, False
    # Span across parts: link to top-left cell only (recommended safe fallback).
    start_ref = link.ref.split(":", 1)[0]
    return start_part, start_ref, True

def rewrite_hyperlinks_in_workbook(
    wb: Workbook,
    base_name: str,
    current_sheet_name: str,
    split_map: dict[str, int],
    verbose: bool = False,
):
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
            if cell.hyperlink:
                raw_target = cell.hyperlink.target or cell.hyperlink.location or ""
                if verbose and is_internal_link(raw_target):
                    print(f"  [Link] Found internal link at {cell.coordinate}: {raw_target}")

                link = extract_internal_link(cell.hyperlink)
                if link and link.sheet_name != current_sheet_name:
                    target_max_rows = split_map.get(link.sheet_name, 0)
                    if target_max_rows > 0:
                        part_num, ref_for_link, spans_parts = get_range_part_decision(link, target_max_rows)
                        new_target = generate_part_external_link(base_name, link.sheet_name, ref_for_link, part_num)
                        if verbose:
                            note = " (range spans parts)" if spans_parts else ""
                            print(f"  [Link Rewrite] {cell.coordinate}: {raw_target} -> {new_target}{note}")
                    else:
                        new_target = generate_external_link(base_name, link.sheet_name, link.ref)
                        if verbose:
                            print(f"  [Link Rewrite] {cell.coordinate}: {raw_target} -> {new_target}")
                    cell.hyperlink.target = new_target
                    cell.hyperlink.location = None
                    has_rewritten = True
    
    # Note: modifying cell.hyperlink.target is sufficient for openpyxl to save it correctly.
