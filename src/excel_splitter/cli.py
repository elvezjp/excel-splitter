import os
import sys
import click
from openpyxl import load_workbook

from .utils import sanitize_filename
from .splitter import split_workbook_by_sheet
from .hyperlinks import rewrite_hyperlinks_in_workbook
from .row_splitter import split_sheet_by_rows

@click.command()
@click.argument('input_file', type=click.Path(exists=True, dir_okay=False))
@click.option('-o', '--output-dir', default='./dist', help='Directory to save split files.', type=click.Path())
@click.option('--max-rows', default=0, help='Maximum rows per sheet (exclude header). Splits into parts if exceeded.', type=int)
@click.option('--dry-run', is_flag=True, help='Simulate split without writing files.')
@click.option('--verbose', is_flag=True, help='Enable verbose output.')
def main(input_file, output_dir, max_rows, dry_run, verbose):
    """
    Excel Splitter: Splits .xlsx files into individual sheets or row chunks.
    Preserves hyperlinks by rewriting them to external relative links.
    """
    try:
        if verbose:
            click.echo(f"Input: {input_file}")
            click.echo(f"Output: {output_dir}")
            click.echo(f"Max Rows: {max_rows}")
            click.echo("---")

        # Validation: .xlsx extension check
        if not input_file.lower().endswith('.xlsx'):
            click.echo("Error: Only .xlsx files are supported.", err=True)
            sys.exit(1)

        # Validation: Check if input file is readable
        if not os.access(input_file, os.R_OK):
            click.echo(f"Error: Cannot read input file: {input_file}", err=True)
            sys.exit(1)

        # Validation: max_rows should be non-negative
        if max_rows < 0:
            click.echo("Error: --max-rows must be >= 0", err=True)
            sys.exit(1)

        base_name = os.path.splitext(os.path.basename(input_file))[0]

        # DRY RUN LOGIC
        if dry_run:
            click.echo("[Dry Run] Simulating split...")
            try:
                wb = load_workbook(input_file, read_only=True)
                for sheet in wb.sheetnames:
                    click.echo(f"- Would create file for sheet: {sheet}")
                wb.close()
            except Exception as e:
                click.echo(f"Error: Failed to read workbook: {e}", err=True)
                sys.exit(1)
            return

        # REAL RUN
        # Create output directory with proper error handling
        try:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
        except PermissionError:
            click.echo(f"Error: No write permission for directory: {output_dir}", err=True)
            sys.exit(1)
        except Exception as e:
            click.echo(f"Error: Failed to create output directory: {e}", err=True)
            sys.exit(1)

        # Validate write permission
        if not os.access(output_dir, os.W_OK):
            click.echo(f"Error: No write permission for directory: {output_dir}", err=True)
            sys.exit(1)

        # Step 1: Workbook Split & Link Rewrite
        click.echo(">> Phase 1: Splitting sheets...")
        try:
            generated_files = split_workbook_by_sheet(input_file, output_dir, verbose)
        except Exception as e:
            click.echo(f"Error: Phase 1 failed (Workbook Split): {e}", err=True)
            sys.exit(1)
        
        # Step 2: Rewrite Hyperlinks in generated files
        click.echo(">> Phase 2: Rewriting hyperlinks...")
        
        final_file_list = []

        for sheet_name, file_path in generated_files:
            try:
                # Load the split file
                wb_split = load_workbook(file_path)
                
                # Rewrite links
                rewrite_hyperlinks_in_workbook(wb_split, base_name, sheet_name, verbose)
                
                # Check if we need Row Split (Phase 3)
                ws = wb_split[sheet_name]
                data_rows = ws.max_row - 1
                
                perform_row_split = (max_rows > 0) and (data_rows > max_rows)
                
                if perform_row_split:
                    if verbose:
                        click.echo(f"  [Row Split] Sheet '{sheet_name}' has {data_rows} rows (> {max_rows}). Splitting...")
                    
                    # Save the link-rewritten version, then feed it to row splitter
                    wb_split.save(file_path)
                    wb_split.close()
                    
                    # Produce parts
                    try:
                        part_files = split_sheet_by_rows(file_path, sheet_name, output_dir, max_rows, base_name, verbose)
                    except Exception as e:
                        click.echo(f"Error: Phase 3 failed (Row Split) for sheet '{sheet_name}': {e}", err=True)
                        sys.exit(1)
                    
                    if part_files:
                        # Delete the whole sheet file to avoid duplication
                        try:
                            os.remove(file_path)
                        except Exception as e:
                            click.echo(f"Warning: Failed to delete original file after split: {e}")
                        final_file_list.extend(part_files)
                        click.echo(f"  -> Created {len(part_files)} parts for '{sheet_name}'.")
                    else:
                        final_file_list.append(file_path)
                else:
                    # Just save the link rewrites
                    wb_split.save(file_path)
                    wb_split.close()
                    final_file_list.append(file_path)
                    
            except Exception as e:
                click.echo(f"Error: Phase 2 failed for sheet '{sheet_name}': {e}", err=True)
                sys.exit(1)

        click.echo(f"\nDone. Generated {len(final_file_list)} files in {output_dir}")

    except Exception as e:
        click.echo(f"Unexpected error: {e}", err=True)
        sys.exit(1)

if __name__ == '__main__':
    main()
