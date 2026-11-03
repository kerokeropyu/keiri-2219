from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from keiri_2219.constants import RED, SOURCE_SHEET, TARGET_SHEET
from keiri_2219.excel_utils import get_cell
from keiri_2219.processor import process_workbook


def _source_row(voucher: int, subject: int, parent: int, debit: int, credit: int) -> list[object]:
    return [2219, 250731, 7, voucher, subject, "", "07ﾃｽﾄ", debit, credit, 0, "課10", parent]


def _target_row(voucher: int, debit: int, credit: int, tax_class: str = "課10") -> list[object]:
    row: list[object] = [384, 250731, 7, voucher, 2219, "預かり金その他", "07ﾃｽﾄ"]
    row += [None, None, debit, None, None, None, None, credit, None, None, None, None, 0, tax_class]
    return row


def _build_workbook(source_rows: list[list[object]], target_rows: list[list[object]]) -> Workbook:
    workbook = Workbook()
    source = workbook.active
    assert isinstance(source, Worksheet)
    source.title = SOURCE_SHEET
    target = workbook.create_sheet(TARGET_SHEET)
    for sheet, rows in ((source, source_rows), (target, target_rows)):
        for _ in range(3):
            sheet.append(["見出し"])
        for row in rows:
            sheet.append(row)
    return workbook


def _target(workbook: Workbook) -> Worksheet:
    sheet = workbook[TARGET_SHEET]
    assert isinstance(sheet, Worksheet)
    return sheet


def test_single_subject_debit_is_written_in_red() -> None:
    workbook = _build_workbook(
        [_source_row(100, 384, 100, 0, 1000), _source_row(101, 6523, 100, 0, 10000)],
        [_target_row(100, 1000, 0)],
    )

    summary = process_workbook(workbook)

    target = _target(workbook)
    assert summary.full_entries == 1
    assert get_cell(target, 4, 8).value == 10
    assert get_cell(target, 4, 11).value == 6523
    assert get_cell(target, 4, 14).value == 1000
    assert get_cell(target, 4, 14).font.color is not None
    assert get_cell(target, 4, 14).font.color.rgb == RED


def test_credit_uses_columns_p_and_s() -> None:
    workbook = _build_workbook(
        [_source_row(100, 384, 100, 1000, 0), _source_row(101, 6523, 100, 10000, 0)],
        [_target_row(100, 0, 1000)],
    )

    process_workbook(workbook)

    target = _target(workbook)
    assert get_cell(target, 4, 16).value == 6523
    assert get_cell(target, 4, 19).value == 1000


def test_multiple_subjects_insert_row_with_red_a_column() -> None:
    workbook = _build_workbook(
        [
            _source_row(100, 384, 100, 0, 3000),
            _source_row(101, 6523, 100, 0, 20000),
            _source_row(102, 243, 100, 0, 10000),
        ],
        [_target_row(100, 3000, 0), _target_row(200, 500, 0)],
    )

    summary = process_workbook(workbook)

    target = _target(workbook)
    assert summary.inserted_rows == 1
    assert get_cell(target, 4, 11).value == 6523
    assert get_cell(target, 4, 14).value == 2000
    assert get_cell(target, 5, 4).value == 100
    assert get_cell(target, 5, 10).value is None
    assert get_cell(target, 5, 11).value == 243
    assert get_cell(target, 5, 14).value == 1000
    assert get_cell(target, 5, 1).fill.start_color.rgb == RED
    assert get_cell(target, 6, 4).value == 200


def test_filled_row_only_gets_tax_amount() -> None:
    workbook = _build_workbook([], [_target_row(100, 1000, 0, "課08")])
    target = _target(workbook)
    get_cell(target, 4, 8).value = "8"
    get_cell(target, 4, 11).value = "7199"

    summary = process_workbook(workbook)

    assert summary.tax_only_entries == 1
    assert get_cell(target, 4, 11).value == "7199"
    assert get_cell(target, 4, 13).value == 1000


def test_row_with_tax_amount_is_not_changed() -> None:
    workbook = _build_workbook([], [_target_row(100, 1000, 0)])
    target = _target(workbook)
    get_cell(target, 4, 14).value = 999

    summary = process_workbook(workbook)

    assert summary.already_done == 1
    assert get_cell(target, 4, 14).value == 999


def test_detail_above_tax_row_is_ignored() -> None:
    workbook = _build_workbook(
        [_source_row(99, 7199, 99, 0, 10000), _source_row(100, 384, 99, 0, 1000)],
        [_target_row(100, 1000, 0)],
    )

    summary = process_workbook(workbook)

    assert summary.skipped == {"税行の下に明細がない": 1}
    assert get_cell(_target(workbook), 4, 14).value is None


def test_tax_mismatch_is_not_written() -> None:
    workbook = _build_workbook(
        [_source_row(100, 384, 100, 0, 1500), _source_row(101, 6523, 100, 0, 10000)],
        [_target_row(100, 1500, 0)],
    )

    summary = process_workbook(workbook)

    assert summary.skipped == {"明細合計×税率と税額が一致しない": 1}
    assert get_cell(_target(workbook), 4, 14).value is None


def test_tax_class_08_is_written_as_8_percent() -> None:
    workbook = _build_workbook(
        [_source_row(100, 384, 100, 0, 800), _source_row(101, 6523, 100, 0, 10000)],
        [_target_row(100, 800, 0, "課08")],
    )

    summary = process_workbook(workbook)

    target = _target(workbook)
    assert summary.full_entries == 1
    assert get_cell(target, 4, 8).value == 8
    assert get_cell(target, 4, 13).value == 800


def test_tax_only_row_with_different_rate_is_not_written() -> None:
    workbook = _build_workbook([], [_target_row(100, 1000, 0, "課10")])
    target = _target(workbook)
    get_cell(target, 4, 8).value = 8
    get_cell(target, 4, 11).value = 7199

    summary = process_workbook(workbook)

    assert summary.skipped == {"課税区分と入力済みの税率が一致しない": 1}
    assert get_cell(target, 4, 13).value is None
