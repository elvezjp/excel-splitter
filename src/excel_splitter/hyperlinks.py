import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Optional
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.utils.cell import coordinate_from_string, range_boundaries

from .utils import get_sheet_filename, get_part_filename

# XML名前空間
NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

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


def get_part_number_from_boundaries(row: int, boundaries: list[tuple[int, int]]) -> int:
    """
    Find which part contains the given row based on actual part boundaries.
    Returns 1-indexed part number.
    If row is not in any boundary (e.g., in a skipped blank range),
    returns the next available part or the last part.
    """
    for i, (start_row, end_row) in enumerate(boundaries, 1):
        if start_row <= row <= end_row:
            return i
    # Row not in any part - find the closest part
    # This can happen if the row falls in a skipped blank region
    for i, (start_row, end_row) in enumerate(boundaries, 1):
        if row < start_row:
            return i  # Return the next part after the gap
    # Row is beyond all parts, return the last part
    return len(boundaries) if boundaries else 1


def get_range_part_decision_from_boundaries(
    link: InternalLink, boundaries: list[tuple[int, int]]
) -> tuple[int, str, bool]:
    """
    Return (target_part, ref_for_link, spans_parts) using actual part boundaries.
    """
    start_part = get_part_number_from_boundaries(link.start_row, boundaries)
    end_part = get_part_number_from_boundaries(link.end_row, boundaries)
    if start_part == end_part:
        return start_part, link.ref, False
    # Span across parts: link to top-left cell only (recommended safe fallback).
    start_ref = link.ref.split(":", 1)[0]
    return start_part, start_ref, True

def has_internal_links(ws) -> bool:
    """
    ワークシートに内部リンク（他シートへのリンク）が含まれているかを確認する。

    Args:
        ws: openpyxl worksheet object

    Returns:
        True if the worksheet contains internal links to other sheets
    """
    current_sheet_name = ws.title

    for row in ws.iter_rows():
        for cell in row:
            if cell.hyperlink:
                link = extract_internal_link(cell.hyperlink)
                if link and link.sheet_name != current_sheet_name:
                    return True
    return False


def rewrite_hyperlinks_zip(
    file_path: str,
    base_name: str,
    current_sheet_name: str,
    split_map: dict[str, list[tuple[int, int]]],
    verbose: bool = False,
) -> None:
    """
    ZIP/XML操作でハイパーリンクを書き換える。
    図形・画像を保持したまま内部リンクを外部リンクに変換する。

    Args:
        file_path: 対象のExcelファイルパス
        base_name: 元ファイルのベース名
        current_sheet_name: 現在のシート名
        split_map: シート分割情報 {sheet_name: [(start_row, end_row), ...]}
        verbose: 詳細出力フラグ
    """
    import tempfile
    import shutil
    import os

    # 一時ファイルを作成
    temp_fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(temp_fd)

    try:
        with zipfile.ZipFile(file_path, 'r') as src_zip:
            # シートファイルのパスを特定
            sheet_xml_path = _find_sheet_xml_path(src_zip, current_sheet_name)
            if not sheet_xml_path:
                if verbose:
                    print(f"  Warning: Could not find sheet XML for '{current_sheet_name}'")
                return

            sheet_rels_path = _get_rels_path(sheet_xml_path)

            # 既存のリレーションを読み込み
            existing_rels = {}
            max_rid = 0
            if sheet_rels_path in src_zip.namelist():
                rels_content = src_zip.read(sheet_rels_path).decode('utf-8')
                existing_rels, max_rid = _parse_existing_rels(rels_content)

            # シートXMLを解析して内部リンクを収集
            sheet_content = src_zip.read(sheet_xml_path).decode('utf-8')
            internal_links = _extract_internal_links_from_xml(sheet_content, current_sheet_name, existing_rels)

            if not internal_links:
                if verbose:
                    print(f"  No internal links found in '{current_sheet_name}'")
                return

            # 新しいリレーションを生成
            new_rels = {}
            updated_rels = {}  # 既存のr:idを更新するもの
            rid_counter = max_rid + 1

            for cell_ref, link_info in internal_links.items():
                target_sheet = link_info['sheet_name']
                ref = link_info['ref']
                existing_rid = link_info.get('existing_rid')

                # リンク先を決定
                boundaries = split_map.get(target_sheet, [])
                if boundaries:
                    link = parse_internal_link(f"#{target_sheet}!{ref}")
                    if link:
                        part_num, ref_for_link, spans_parts = get_range_part_decision_from_boundaries(link, boundaries)
                        new_target = generate_part_external_link(base_name, target_sheet, ref_for_link, part_num)
                        if verbose and spans_parts:
                            print(f"  [Link Rewrite] {cell_ref}: range spans parts")
                    else:
                        new_target = generate_external_link(base_name, target_sheet, ref)
                else:
                    new_target = generate_external_link(base_name, target_sheet, ref)

                if existing_rid:
                    # 既存のr:idがある場合は更新
                    updated_rels[existing_rid] = {
                        'rid': existing_rid,
                        'target': new_target,
                        'cell_ref': cell_ref,
                    }
                else:
                    # 新しいr:idを割り当て
                    rid = f"rId{rid_counter}"
                    new_rels[cell_ref] = {
                        'rid': rid,
                        'target': new_target,
                    }
                    rid_counter += 1

                if verbose:
                    print(f"  [Link Rewrite] {cell_ref}: #{target_sheet}!{ref} -> {new_target}")

            # シートXMLを書き換え（新しいr:idが必要なものだけ）
            new_sheet_content = _rewrite_sheet_xml_hyperlinks(sheet_content, new_rels)

            # リレーションXMLを生成/更新
            new_rels_content = _generate_rels_xml(
                existing_rels, new_rels, updated_rels, sheet_rels_path in src_zip.namelist()
            )

            # 新しいZIPを作成
            with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as dst_zip:
                for item in src_zip.namelist():
                    if item == sheet_xml_path:
                        dst_zip.writestr(item, new_sheet_content.encode('utf-8'))
                    elif item == sheet_rels_path:
                        dst_zip.writestr(item, new_rels_content.encode('utf-8'))
                    else:
                        dst_zip.writestr(item, src_zip.read(item))

                # リレーションファイルが存在しなかった場合は新規作成
                if sheet_rels_path not in src_zip.namelist() and new_rels:
                    dst_zip.writestr(sheet_rels_path, new_rels_content.encode('utf-8'))
                    # Content_Types.xmlの更新は不要（relsファイルは自動認識される）

        # 一時ファイルで元ファイルを置き換え
        shutil.move(temp_path, file_path)

    except Exception as e:
        # エラー時は一時ファイルを削除
        if os.path.exists(temp_path):
            os.remove(temp_path)
        raise e


def _find_sheet_xml_path(zip_file: zipfile.ZipFile, sheet_name: str) -> Optional[str]:
    """
    シート名に対応するXMLファイルのパスを見つける。
    """
    # workbook.xmlからシート情報を取得
    wb_xml = zip_file.read('xl/workbook.xml').decode('utf-8')
    wb_root = ET.fromstring(wb_xml)

    # シート名からr:idを取得
    target_rid = None
    for sheet_elem in wb_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
        if sheet_elem.get('name') == sheet_name:
            target_rid = sheet_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            break

    if not target_rid:
        return None

    # workbook.xml.relsからファイルパスを取得
    rels_xml = zip_file.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rels_root = ET.fromstring(rels_xml)

    for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        if rel.get('Id') == target_rid:
            target = rel.get('Target')
            # 絶対パス（/xl/...）の場合は先頭の/を除去
            target = target.lstrip('/')
            # 相対パス（worksheets/...）の場合はxl/を追加
            if not target.startswith('xl/'):
                target = 'xl/' + target
            return target

    return None


def _get_rels_path(sheet_xml_path: str) -> str:
    """
    シートXMLパスに対応するリレーションファイルのパスを返す。
    """
    import os
    dir_path = os.path.dirname(sheet_xml_path)
    filename = os.path.basename(sheet_xml_path)
    return f"{dir_path}/_rels/{filename}.rels"


def _parse_existing_rels(rels_content: str) -> tuple[dict, int]:
    """
    既存のリレーションXMLを解析する。

    Returns:
        (existing_rels dict, max_rid number)
    """
    existing_rels = {}
    max_rid = 0

    root = ET.fromstring(rels_content)
    for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        rid = rel.get('Id')
        existing_rels[rid] = {
            'type': rel.get('Type'),
            'target': rel.get('Target'),
            'target_mode': rel.get('TargetMode'),
        }
        # rIdの番号を抽出
        if rid and rid.startswith('rId'):
            try:
                rid_num = int(rid[3:])
                max_rid = max(max_rid, rid_num)
            except ValueError:
                pass

    return existing_rels, max_rid


def _extract_internal_links_from_xml(
    sheet_content: str,
    current_sheet_name: str,
    existing_rels: dict
) -> dict:
    """
    シートXMLから内部リンクを抽出する。
    location属性を持つものと、r:id経由で内部リンクを参照するものの両方に対応。

    Returns:
        {cell_ref: {'sheet_name': str, 'ref': str, 'location': str, 'existing_rid': str|None}}
    """
    internal_links = {}

    root = ET.fromstring(sheet_content)

    # hyperlinks要素を探す
    for hyperlink in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}hyperlink'):
        cell_ref = hyperlink.get('ref')
        location = hyperlink.get('location')
        r_id = hyperlink.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')

        # location属性がある場合
        if location and cell_ref:
            link = parse_internal_link(f"#{location}")
            if link and link.sheet_name != current_sheet_name:
                internal_links[cell_ref] = {
                    'sheet_name': link.sheet_name,
                    'ref': link.ref,
                    'location': location,
                    'existing_rid': None,
                }

        # r:id属性がある場合（openpyxlで作成されたファイル）
        elif r_id and cell_ref and r_id in existing_rels:
            rel_info = existing_rels[r_id]
            target = rel_info.get('target', '')

            # Targetが#で始まる場合は内部リンク
            if target.startswith('#'):
                link = parse_internal_link(target)
                if link and link.sheet_name != current_sheet_name:
                    internal_links[cell_ref] = {
                        'sheet_name': link.sheet_name,
                        'ref': link.ref,
                        'location': target.lstrip('#'),
                        'existing_rid': r_id,
                    }

    return internal_links


def _rewrite_sheet_xml_hyperlinks(sheet_content: str, new_rels: dict) -> str:
    """
    シートXML内のhyperlink要素を書き換える。
    location属性をr:id属性に変換する。
    """
    # 正規表現で各hyperlinkを書き換え
    for cell_ref, rel_info in new_rels.items():
        rid = rel_info['rid']

        # location属性を持つhyperlinkを見つけてr:idに変換
        # パターン: <hyperlink ref="A1" location="Sheet2!B5"/>
        # または: <hyperlink ref="A1" location="Sheet2!B5" display="..."/>

        # まず、該当するhyperlink要素全体を見つける
        pattern = rf'(<hyperlink\s+[^>]*ref="{re.escape(cell_ref)}"[^>]*)(location="[^"]*")([^>]*/?>)'

        def replace_location(match):
            prefix = match.group(1)
            suffix = match.group(3)
            # location属性をr:id属性に置き換え
            # 名前空間プレフィックスを追加
            return f'{prefix}r:id="{rid}"{suffix}'

        sheet_content = re.sub(pattern, replace_location, sheet_content)

    return sheet_content


def _generate_rels_xml(
    existing_rels: dict, new_rels: dict, updated_rels: dict, has_existing_file: bool
) -> str:
    """
    リレーションXMLを生成する。

    Args:
        existing_rels: 既存のリレーション {rid: {type, target, target_mode}}
        new_rels: 新規追加するリレーション {cell_ref: {rid, target}}
        updated_rels: 更新するリレーション {rid: {rid, target, cell_ref}}
        has_existing_file: 既存のリレーションファイルが存在するか
    """
    lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    lines.append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')

    # 既存のリレーションを追加（updated_relsで更新するものは新しいターゲットを使用）
    for rid, rel_info in existing_rels.items():
        if rid in updated_rels:
            # このリレーションは更新される
            new_target = _escape_xml_attr(updated_rels[rid]['target'])
            lines.append(f'<Relationship Id="{rid}" Type="{HYPERLINK_REL_TYPE}" Target="{new_target}" TargetMode="External"/>')
        else:
            # そのまま維持
            target_mode = f' TargetMode="{rel_info["target_mode"]}"' if rel_info.get('target_mode') else ''
            lines.append(f'<Relationship Id="{rid}" Type="{rel_info["type"]}" Target="{rel_info["target"]}"{target_mode}/>')

    # 新しいハイパーリンクリレーションを追加
    for cell_ref, rel_info in new_rels.items():
        rid = rel_info['rid']
        target = _escape_xml_attr(rel_info['target'])
        lines.append(f'<Relationship Id="{rid}" Type="{HYPERLINK_REL_TYPE}" Target="{target}" TargetMode="External"/>')

    lines.append('</Relationships>')
    return '\n'.join(lines)


def _escape_xml_attr(value: str) -> str:
    """
    XML属性値をエスケープする。
    """
    return (value
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))
