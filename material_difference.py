"""汇总机柜设计中所有 -MX 工作表的整件物料。"""

from collections import OrderedDict
from copy import copy
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter


def output_path_for(source, output_dir=None):
    source = Path(source)
    directory = Path(output_dir) if output_dir else Path(__file__).resolve().parent
    stem = source.stem.replace("机柜设计", "差异表", 1)
    if stem == source.stem:
        stem += "-差异表"
    return directory / f"{stem}.xlsx"


def _headers(ws):
    wanted = {"代号", "名称", "数量", "备注", "更改"}
    found = {}
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 8)):
        for cell in row:
            key = "".join(str(cell.value or "").split())
            if key in wanted:
                found[key] = (cell.row, cell.column)
    if set(found) != wanted or len({r for r, _ in found.values()}) != 1:
        raise ValueError(f"{ws.title} 缺少标准表头：{', '.join(sorted(wanted - set(found)))}")
    return found


def _marker_row(ws):
    for row in ws.iter_rows():
        for cell in row:
            if "".join(str(cell.value or "").split()) == "整件" and cell.font.italic:
                return cell.row
    raise ValueError(f"{ws.title} 找不到斜体的“整件”行")


def _copy_rows(source, target, first, last, destination, first_column=1, last_column=None, column_offset=0):
    offset = destination - first
    last_column = last_column or source.max_column
    for row in source.iter_rows(min_row=first, max_row=last,
                                min_col=first_column, max_col=last_column):
        for old in row:
            new = target.cell(old.row + offset, old.column + column_offset)
            new.value = old.value
            _copy_style(old, new)
        old_dim = source.row_dimensions[row[0].row]
        new_dim = target.row_dimensions[row[0].row + offset]
        new_dim.height = old_dim.height
        new_dim.hidden = old_dim.hidden
    for merged in source.merged_cells.ranges:
        if (first <= merged.min_row and merged.max_row <= last
                and first_column <= merged.min_col and merged.max_col <= last_column):
            target.merge_cells(
                start_row=merged.min_row + offset,
                end_row=merged.max_row + offset,
                start_column=merged.min_col + column_offset,
                end_column=merged.max_col + column_offset,
            )


def _copy_style(old, new):
    if old.has_style:
        new.font = copy(old.font)
        new.fill = copy(old.fill)
        new.border = copy(old.border)
        new.alignment = copy(old.alignment)
        new.number_format = old.number_format
        new.protection = copy(old.protection)



def _xls_color(book, index):
    """把 xlrd 调色板颜色转换为 openpyxl 的 ARGB。"""
    rgb = book.colour_map.get(index)
    if not rgb:
        return None
    return "FF{:02X}{:02X}{:02X}".format(*rgb)


def _xls_side(book, border, edge):
    styles = {
        0: None, 1: "thin", 2: "medium", 3: "dashed", 4: "dotted",
        5: "thick", 6: "double", 7: "hair", 8: "mediumDashed",
        9: "dashDot", 10: "mediumDashDot", 11: "dashDotDot",
        12: "mediumDashDotDot", 13: "slantDashDot",
    }
    style = styles.get(getattr(border, f"{edge}_line_style", 0))
    color = _xls_color(book, getattr(border, f"{edge}_colour_index", 0))
    return Side(style=style, color=color)


def _apply_xls_style(book, source_sheet, row, column, target_cell):
    """复制旧版 .xls 单元格的常用显示格式。"""
    xf = book.xf_list[source_sheet.cell_xf_index(row, column)]
    old_font = book.font_list[xf.font_index]
    target_cell.font = Font(
        name=old_font.name or "宋体",
        size=old_font.height / 20 if old_font.height else None,
        bold=bool(old_font.bold),
        italic=bool(old_font.italic),
        underline="single" if old_font.underlined else None,
        strike=bool(old_font.struck_out),
        color=_xls_color(book, old_font.colour_index),
    )
    background = xf.background
    fill_color = _xls_color(book, background.pattern_colour_index)
    target_cell.fill = PatternFill(
        fill_type="solid" if background.fill_pattern else None,
        fgColor=fill_color or "00000000",
    )
    border = xf.border
    target_cell.border = Border(
        left=_xls_side(book, border, "left"),
        right=_xls_side(book, border, "right"),
        top=_xls_side(book, border, "top"),
        bottom=_xls_side(book, border, "bottom"),
    )
    horizontal = {
        0: "general", 1: "left", 2: "center", 3: "right",
        4: "fill", 5: "justify", 6: "centerContinuous", 7: "distributed",
    }.get(xf.alignment.hor_align, "general")
    vertical = {
        0: "top", 1: "center", 2: "bottom", 3: "justify", 4: "distributed",
    }.get(xf.alignment.vert_align, "bottom")
    rotation = xf.alignment.rotation
    target_cell.alignment = Alignment(
        horizontal=horizontal,
        vertical=vertical,
        wrap_text=bool(xf.alignment.text_wrapped),
        shrink_to_fit=bool(xf.alignment.shrink_to_fit),
        text_rotation=rotation if 0 <= rotation <= 180 else 0,
    )
    number_format = book.format_map.get(xf.format_key)
    if number_format:
        target_cell.number_format = number_format.format_str
    target_cell.protection = Protection(
        locked=bool(xf.protection.cell_locked),
        hidden=bool(xf.protection.formula_hidden),
    )


def _load_xls(source_path):
    """读取旧版 .xls，并转换为差异表生成器使用的工作簿对象。"""
    import xlrd

    legacy = xlrd.open_workbook(str(source_path), formatting_info=True)
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    for old_sheet in legacy.sheets():
        if not old_sheet.name.upper().endswith("-MX"):
            continue
        sheet = workbook.create_sheet(old_sheet.name)
        for row in range(old_sheet.nrows):
            row_info = old_sheet.rowinfo_map.get(row)
            if row_info:
                sheet.row_dimensions[row + 1].height = row_info.height / 20
                sheet.row_dimensions[row + 1].hidden = bool(row_info.hidden)
            for column in range(old_sheet.ncols):
                old_cell = old_sheet.cell(row, column)
                value = old_cell.value
                if old_cell.ctype == xlrd.XL_CELL_DATE:
                    value = xlrd.xldate_as_datetime(value, legacy.datemode)
                elif old_cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    value = bool(value)
                elif old_cell.ctype == xlrd.XL_CELL_ERROR:
                    value = xlrd.error_text_from_code.get(value, "#ERROR!")
                target = sheet.cell(row + 1, column + 1, value)
                _apply_xls_style(legacy, old_sheet, row, column, target)
        for column, column_info in old_sheet.colinfo_map.items():
            dimension = sheet.column_dimensions[get_column_letter(column + 1)]
            dimension.width = column_info.width / 256
            dimension.hidden = bool(column_info.hidden)
        for row_start, row_end, col_start, col_end in old_sheet.merged_cells:
            sheet.merge_cells(
                start_row=row_start + 1,
                end_row=row_end,
                start_column=col_start + 1,
                end_column=col_end,
            )
    return workbook


def _load_source(source_path):
    if source_path.suffix.lower() == ".xls":
        workbook = _load_xls(source_path)
        return workbook, workbook
    return (
        openpyxl.load_workbook(source_path, data_only=False),
        openpyxl.load_workbook(source_path, data_only=True),
    )


def generate_difference(source_path, output_path=None):
    """生成差异表，返回 (输出路径, 不同代号数)。"""
    source_path = Path(source_path)
    if source_path.suffix.lower() not in {".xls", ".xlsx", ".xlsm"}:
        raise ValueError("差异表需要 .xls、.xlsx 或 .xlsm 机柜设计文件")
    output_path = Path(output_path) if output_path else output_path_for(source_path)
    if source_path.resolve() == output_path.resolve():
        raise ValueError("输出路径不能与源文件相同")

    workbook, cached = _load_source(source_path)
    mx_sheets = [ws for ws in workbook.worksheets if ws.title.upper().endswith("-MX")]
    if not mx_sheets:
        raise ValueError("源文件中没有以 -MX 结尾的工作表")

    materials = OrderedDict()
    for ws in mx_sheets:
        headers = _headers(ws)
        marker = _marker_row(ws)
        code_col = headers["代号"][1]
        name_col = headers["名称"][1]
        qty_col = headers["数量"][1]
        for row in range(marker + 1, ws.max_row + 1):
            code_value = ws.cell(row, code_col).value
            if code_value is None or not str(code_value).strip():
                continue
            code = str(code_value).strip()
            raw_qty = cached[ws.title].cell(row, qty_col).value
            try:
                quantity = Decimal(str(raw_qty))
                if not quantity.is_finite():
                    raise InvalidOperation
            except (InvalidOperation, TypeError):
                raise ValueError(f"{ws.title}!{get_column_letter(qty_col)}{row} 数量无效：{raw_qty!r}") from None
            if code not in materials:
                materials[code] = [ws.cell(row, name_col).value, Decimal(0)]
            materials[code][1] += quantity

    template = mx_sheets[0]
    header = _headers(template)
    section = _marker_row(template)
    shift = -5  # 删除模板未使用的 A:E，F 列移到输出 A 列
    output_columns = 52  # A:AZ，原因只占一个普通单元格

    result = openpyxl.Workbook()
    sheet = result.active
    sheet.title = "机柜差异表"
    sheet.sheet_format = copy(template.sheet_format)
    sheet.sheet_properties = copy(template.sheet_properties)
    sheet.page_margins = copy(template.page_margins)
    sheet.print_options = copy(template.print_options)
    sheet.page_setup = copy(template.page_setup)
    sheet.sheet_view.view = "normal"
    sheet.sheet_view.showGridLines = True

    _copy_rows(template, sheet, 2, 3, 2, 6, 56, shift)
    title_stem = source_path.stem.replace("机柜设计", "差异表", 1)
    if title_stem == source_path.stem:
        title_stem += "-差异表"
    title = re.sub(r"-V\d+(?:\.\d+)?$", "", title_stem, flags=re.IGNORECASE)
    sheet.merge_cells(start_row=1, end_row=1, start_column=1, end_column=output_columns)
    sheet.cell(1, 1).value = title
    sheet.cell(1, 1).font = Font(name="宋体", size=14, bold=True)
    sheet.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 24

    for index, (code, (name, quantity)) in enumerate(materials.items(), 1):
        row = 4 + 2 * (index - 1)
        _copy_rows(template, sheet, section + 2, section + 3, row, 6, 56, shift)
        sheet.cell(row, 1).value = index
        sheet.cell(row, header["代号"][1] + shift).value = code
        sheet.cell(row, header["名称"][1] + shift).value = name
        sheet.cell(row, header["数量"][1] + shift).value = int(quantity) if quantity == int(quantity) else float(quantity)
        sheet.cell(row, header["备注"][1] + shift).value = None
        sheet.cell(row, header["更改"][1] + shift).value = f'=IF(AL{row}="","",AH{row}-AL{row})'
        _copy_style(template.cell(section + 2, header["备注"][1]), sheet.cell(row, output_columns))
        sheet.cell(row, output_columns).value = None
        sheet.merge_cells(start_row=row, end_row=row + 1,
                          start_column=output_columns, end_column=output_columns)

    sheet.cell(2, header["备注"][1] + shift).value = "合同数量"
    sheet.cell(2, header["更改"][1] + shift).value = "数量差异"
    _copy_style(template.cell(2, header["备注"][1]), sheet.cell(2, output_columns))
    sheet.cell(2, output_columns).value = "原因"
    sheet.merge_cells(start_row=2, end_row=3,
                      start_column=output_columns, end_column=output_columns)
    sheet.column_dimensions["AZ"].width = 18
    sheet.freeze_panes = None
    sheet.sheet_properties.pageSetUpPr.autoPageBreaks = False
    sheet.sheet_properties.pageSetUpPr.fitToPage = False
    sheet.print_area = None
    sheet.print_title_rows = None

    thin = Side(style="thin", color="000000")
    all_lines = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row,
                               min_col=1, max_col=output_columns):
        for cell in row:
            cell.border = all_lines

    output_path.parent.mkdir(parents=True, exist_ok=True)
    result.save(output_path)
    return output_path, len(materials)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="从机柜设计文件生成机柜差异表")
    parser.add_argument("source", help="机柜设计 .xls/.xlsx/.xlsm 文件")
    parser.add_argument("--output", help="输出路径，默认在程序目录")
    arguments = parser.parse_args()
    path, count = generate_difference(arguments.source, arguments.output)
    print(f"已生成 {path}，共 {count} 种代号")
