"""
ZIP/XML操作によるシート分割モジュール。

図形・画像を保持するため、openpyxlではなくZIP/XMLを直接操作する。
"""
import os
import re
import zipfile
import xml.etree.ElementTree as ET

from .utils import get_sheet_filename

# XML名前空間
NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}


def get_sheet_info(zip_file: zipfile.ZipFile) -> list[dict]:
    """
    workbook.xmlとworkbook.xml.relsからシート情報を取得する。

    Returns:
        List of dicts with keys: name, sheet_id, r_id, target (sheet file path)
    """
    # workbook.xmlを解析
    wb_xml = zip_file.read('xl/workbook.xml').decode('utf-8')
    wb_root = ET.fromstring(wb_xml)

    sheets = []
    for sheet_elem in wb_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
        sheet_info = {
            'name': sheet_elem.get('name'),
            'sheet_id': sheet_elem.get('sheetId'),
            'r_id': sheet_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'),
        }
        sheets.append(sheet_info)

    # workbook.xml.relsを解析してtargetを取得
    rels_xml = zip_file.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rels_root = ET.fromstring(rels_xml)

    r_id_to_target = {}
    for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        r_id_to_target[rel.get('Id')] = rel.get('Target')

    for sheet in sheets:
        target = r_id_to_target.get(sheet['r_id'], '')
        # 相対パスを正規化（worksheets/sheet1.xml -> xl/worksheets/sheet1.xml）
        if not target.startswith('xl/'):
            target = 'xl/' + target.lstrip('/')
        sheet['target'] = target

    return sheets


def split_workbook_by_sheet(input_path: str, output_dir: str, verbose: bool = False) -> list[tuple[str, str]]:
    """
    ZIP/XML操作でワークブックを各シートごとのファイルに分割する。
    図形・画像を保持するため、openpyxlではなくZIPを直接操作する。

    Returns:
        List of (sheet_name, output_file_path) tuples.
    """
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    generated_files = []

    with zipfile.ZipFile(input_path, 'r') as src_zip:
        sheet_info_list = get_sheet_info(src_zip)

        if verbose:
            print(f"Found {len(sheet_info_list)} sheets in workbook")

        for target_sheet_info in sheet_info_list:
            target_sheet_name = target_sheet_info['name']

            if verbose:
                print(f"Processing sheet: {target_sheet_name}")

            # 出力ファイルパス
            filename = get_sheet_filename(base_name, target_sheet_name)
            out_path = os.path.join(output_dir, filename)

            # 削除するシートのリスト（対象シート以外）
            sheets_to_remove = [s for s in sheet_info_list if s['name'] != target_sheet_name]

            # 新しいZIPを作成
            _create_split_workbook(src_zip, out_path, target_sheet_info, sheets_to_remove)

            generated_files.append((target_sheet_name, out_path))

    return generated_files


def _create_split_workbook(
    src_zip: zipfile.ZipFile,
    out_path: str,
    target_sheet: dict,
    sheets_to_remove: list[dict]
) -> None:
    """
    対象シート以外を削除した新しいワークブックを作成する。
    """
    # 削除対象のファイルパス
    files_to_skip = set()
    for sheet in sheets_to_remove:
        # シートファイル
        files_to_skip.add(sheet['target'])
        # シートのリレーションファイル
        sheet_filename = os.path.basename(sheet['target'])
        sheet_rels_path = f"xl/worksheets/_rels/{sheet_filename}.rels"
        files_to_skip.add(sheet_rels_path)

    # 削除対象シートのr_idリスト
    r_ids_to_remove = {s['r_id'] for s in sheets_to_remove}
    # 削除対象シートの名前リスト
    names_to_remove = {s['name'] for s in sheets_to_remove}
    # 削除対象シートのtargetパスリスト
    targets_to_remove = {s['target'] for s in sheets_to_remove}

    # 新しいZIPを作成
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as dst_zip:
        for item in src_zip.namelist():
            # スキップ対象のファイル
            if item in files_to_skip:
                continue

            data = src_zip.read(item)

            # XMLファイルの編集（正規表現ベース - 名前空間を保持）
            if item == 'xl/workbook.xml':
                data = _edit_workbook_xml(data, names_to_remove)
            elif item == 'xl/_rels/workbook.xml.rels':
                data = _edit_workbook_rels(data, r_ids_to_remove)
            elif item == '[Content_Types].xml':
                data = _edit_content_types(data, targets_to_remove)

            dst_zip.writestr(item, data)


def _edit_workbook_xml(data: bytes, names_to_remove: set[str]) -> bytes:
    """
    workbook.xmlから不要なシート要素を削除する。
    正規表現ベースで元のXML構造を保持する。
    """
    content = data.decode('utf-8')

    # 各シート名に対して<sheet>要素を削除
    for name in names_to_remove:
        # シート名をエスケープ（正規表現の特殊文字対策）
        escaped_name = re.escape(name)
        # <sheet ... name="シート名" ... /> または <sheet ... name="シート名" ...></sheet> を削除
        pattern = rf'<sheet\s+[^>]*name="{escaped_name}"[^>]*/>'
        content = re.sub(pattern, '', content)
        # 閉じタグ形式の場合
        pattern = rf'<sheet\s+[^>]*name="{escaped_name}"[^>]*>\s*</sheet>'
        content = re.sub(pattern, '', content)

    # activeTab属性を0にリセット（シートが減っているため）
    content = re.sub(r'activeTab="\d+"', 'activeTab="0"', content)

    return content.encode('utf-8')


def _edit_workbook_rels(data: bytes, r_ids_to_remove: set[str]) -> bytes:
    """
    workbook.xml.relsから不要なリレーションを削除する。
    正規表現ベースで元のXML構造を保持する。
    """
    content = data.decode('utf-8')

    # 各r_idに対して<Relationship>要素を削除
    for r_id in r_ids_to_remove:
        escaped_r_id = re.escape(r_id)
        # <Relationship Id="rIdX" ... /> を削除
        pattern = rf'<Relationship\s+[^>]*Id="{escaped_r_id}"[^>]*/>'
        content = re.sub(pattern, '', content)

    return content.encode('utf-8')


def _edit_content_types(data: bytes, targets_to_remove: set[str]) -> bytes:
    """
    [Content_Types].xmlから不要なOverride要素を削除する。
    正規表現ベースで元のXML構造を保持する。
    """
    content = data.decode('utf-8')

    # 各targetに対して<Override>要素を削除
    for target in targets_to_remove:
        # /xl/worksheets/sheet1.xml の形式に変換
        part_name = '/' + target if not target.startswith('/') else target
        escaped_part_name = re.escape(part_name)
        # <Override PartName="/xl/worksheets/sheetX.xml" ... /> を削除
        pattern = rf'<Override\s+[^>]*PartName="{escaped_part_name}"[^>]*/>'
        content = re.sub(pattern, '', content)

    return content.encode('utf-8')
