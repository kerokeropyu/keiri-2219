"""openpyxl の操作をまとめたヘルパー。"""

from copy import copy
from typing import cast

from openpyxl.cell.cell import Cell
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from keiri_2219.constants import RED


def get_cell(sheet: Worksheet, row: int, column: int) -> Cell:
    """結合セルのないシートから Cell を取得する。"""
    return cast(Cell, sheet.cell(row=row, column=column))


def to_int(value: object) -> int | None:
    """セルの値を整数に変換する。数値でなければ None を返す。"""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, int | float):
        return int(value)
    if isinstance(value, str):
        try:
            return int(float(value.strip()))
        except ValueError:
            return None
    return None


def is_empty(value: object) -> bool:
    """セルが空かどうかを返す。"""
    return value is None or value == ""


def write_red(cell: Cell, value: int | str | None) -> None:
    """値を赤文字で書き込む。"""
    font = cell.font
    cell.value = value
    cell.font = Font(
        name=font.name,
        size=font.size,
        bold=font.bold,
        italic=font.italic,
        underline=font.underline,
        color=RED,
    )


def fill_red(cell: Cell) -> None:
    """セルの背景を赤にする。"""
    cell.fill = PatternFill(fill_type="solid", start_color=RED, end_color=RED)


def insert_rows_below(sheet: Worksheet, row: int, count: int) -> None:
    """指定行の直下に行を追加し、書式・行の高さ・非表示状態を引き継ぐ。

    openpyxl の insert_rows は行の高さや非表示状態、フィルター範囲を
    ずらさないため、ここで合わせて更新する。
    """
    sheet.insert_rows(row + 1, count)
    _shift_row_dimensions(sheet, row, count)
    _extend_auto_filter(sheet, count)

    source_dimension = sheet.row_dimensions[row]
    for new_row in range(row + 1, row + count + 1):
        sheet.row_dimensions[new_row].height = source_dimension.height
        sheet.row_dimensions[new_row].hidden = source_dimension.hidden
        for column in range(1, sheet.max_column + 1):
            source = get_cell(sheet, row, column)
            target = get_cell(sheet, new_row, column)
            target._style = copy(source._style)  # type: ignore[attr-defined]


def _shift_row_dimensions(sheet: Worksheet, row: int, count: int) -> None:
    """指定行より下の行情報を count 行ずらす。"""
    rows_to_move = sorted((r for r in sheet.row_dimensions if r > row), reverse=True)
    for current_row in rows_to_move:
        dimension = sheet.row_dimensions.pop(current_row)
        dimension.index = current_row + count
        sheet.row_dimensions[current_row + count] = dimension


def _extend_auto_filter(sheet: Worksheet, count: int) -> None:
    """フィルター範囲の最終行を count 行広げる。"""
    if not sheet.auto_filter.ref:
        return
    start, end = sheet.auto_filter.ref.split(":")
    last_column = end.rstrip("0123456789")
    last_row = int(end[len(last_column) :])
    sheet.auto_filter.ref = f"{start}:{last_column}{last_row + count}"
