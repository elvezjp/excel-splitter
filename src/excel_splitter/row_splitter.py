import os
from copy import copy
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from .utils import get_part_filename

def copy_row_style(source_cell, target_cell):
    """
    Copy basic style and hyperlink from source to target.
    Deepcopying style objects is safer than assigning references if across workbooks.
    """
    if source_cell.has_style:
        # Assigning style objects directly can work if openpyxl version supports it, 
        # but often copying attributes is safer for different workbooks.
        # For simplicity/speed in 'large sheet' context, we might only copy values if too slow.
        # Let's try to copy essentials: font, fill, border, alignment.
        try:
            target_cell.font = copy(source_cell.font)
            target_cell.fill = copy(source_cell.fill)
            target_cell.border = copy(source_cell.border)
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.number_format = source_cell.number_format
        except:
            # Fallback if copy fails
            pass
    
    # Copy hyperlink if present
    if source_cell.hyperlink:
        try:
            target_cell.hyperlink = copy(source_cell.hyperlink)
        except Exception as e:
            # If hyperlink copy fails, log but don't crash
            pass

def split_sheet_by_rows(
    origin_wb_path: str,
    sheet_name: str,
    output_dir: str,
    max_rows: int,
    base_name: str,
    verbose: bool = False
) -> list[str]:
    """
    Split a specific sheet into multiple files based on max_rows.
    Returns list of generated file paths.
    """
    
    # Load source. Read-only might be faster for reading, 
    # but we need cell styles. So standard load.
    wb_source = load_workbook(origin_wb_path, read_only=False, data_only=False)
    ws_source = wb_source[sheet_name]
    
    # 1. Analyze Header
    # Assuming Row 1 is header.
    header_cells = list(ws_source[1]) # Tuple of cells in row 1
    
    # 2. Count rows
    # ws.max_row includes empty rows often.
    total_rows = ws_source.max_row
    
    # If total_rows <= max_rows + 1 (header), no split needed? 
    # Logic: "Split if data rows > max_rows"
    data_rows_count = total_rows - 1
    if data_rows_count <= max_rows:
        return [] # No split happened here. Caller handles "Single file" case.
        
    generated_files = []
    
    # 3. Chunking
    # data starts at row 2
    current_row = 2
    part_num = 1
    
    while current_row <= total_rows:
        if verbose:
            print(f"  Creating Part {part_num} for {sheet_name} (Row {current_row}...)")
            
        new_wb = Workbook()
        new_ws = new_wb.active
        new_ws.title = sheet_name
        
        # A. Write Header
        for col_idx, cell in enumerate(header_cells, 1):
            new_cell = new_ws.cell(row=1, column=col_idx, value=cell.value)
            copy_row_style(cell, new_cell)
            
        # B. Write Data Chunk
        # Calculate end row for this chunk
        end_row = min(current_row + max_rows - 1, total_rows)
        
        # Iterate rows in source
        # openpyxl slicing: ws[row_result]
        # ws.iter_rows is generator
        
        row_write_idx = 2
        for row in ws_source.iter_rows(min_row=current_row, max_row=end_row, values_only=False):
            for col_idx, cell in enumerate(row, 1):
                new_cell = new_ws.cell(row=row_write_idx, column=col_idx, value=cell.value)
                copy_row_style(cell, new_cell)
            row_write_idx += 1
            
        # Save
        filename = get_part_filename(base_name, sheet_name, part_num)
        out_path = os.path.join(output_dir, filename)
        new_wb.save(out_path)
        new_wb.close()
        
        generated_files.append(out_path)
        
        current_row = end_row + 1
        part_num += 1
        
    wb_source.close()
    return generated_files
