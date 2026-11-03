"""2219処理の本体。入力内容を決めてから「0384仮払い消費税」シートへ書き込む。"""

import logging
from dataclasses import dataclass, field

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from keiri_2219.constants import (
    ACCOUNT_2219,
    HEADER_ROWS,
    SOURCE_SHEET,
    TARGET_SHEET,
    TAX_CLASS_PERCENTS,
    TGT_COPY_LAST_COLUMN,
    TGT_CREDIT,
    TGT_CREDIT_SUBJECT,
    TGT_CREDIT_TAX_COLUMNS,
    TGT_DEBIT,
    TGT_DEBIT_SUBJECT,
    TGT_DEBIT_TAX_COLUMNS,
    TGT_RATE,
    TGT_SUBJECT_CODE,
    TGT_TAX_CLASS,
    TGT_VOUCHER,
)
from keiri_2219.excel_utils import (
    fill_red,
    get_cell,
    insert_rows_below,
    is_empty,
    to_int,
    write_red,
)
from keiri_2219.ledger import TaxSet, build_tax_sets, read_ledger_rows
from keiri_2219.tax import TaxLine, allocate_tax, choose_percent, is_tax_valid

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class Side:
    """借方または貸方の列。"""

    name: str
    amount_column: int
    subject_column: int
    tax_columns: dict[int, int]


DEBIT_SIDE = Side("借方", TGT_DEBIT, TGT_DEBIT_SUBJECT, TGT_DEBIT_TAX_COLUMNS)
CREDIT_SIDE = Side("貸方", TGT_CREDIT, TGT_CREDIT_SUBJECT, TGT_CREDIT_TAX_COLUMNS)


@dataclass(frozen=True)
class FullEntry:
    """税率・科目・税額を入力する。科目が複数なら行を追加する。"""

    row: int
    side: Side
    percent: int
    lines: list[TaxLine]


@dataclass(frozen=True)
class TaxOnlyEntry:
    """税率と科目が入力済みの行に、税額だけを入力する。"""

    row: int
    side: Side
    percent: int
    tax: int


@dataclass
class Summary:
    """処理結果の件数。"""

    target_rows: int = 0
    full_entries: int = 0
    tax_only_entries: int = 0
    inserted_rows: int = 0
    already_done: int = 0
    skipped: dict[str, int] = field(default_factory=dict)

    def add_skip(self, reason: str) -> None:
        self.skipped[reason] = self.skipped.get(reason, 0) + 1


def process_workbook(workbook: Workbook) -> Summary:
    """ブック全体に2219処理を行う。"""
    source = _get_sheet(workbook, SOURCE_SHEET)
    target = _get_sheet(workbook, TARGET_SHEET)

    tax_sets = build_tax_sets(read_ledger_rows(source))
    summary = Summary()
    full_entries: list[FullEntry] = []
    tax_only_entries: list[TaxOnlyEntry] = []

    for row in range(HEADER_ROWS + 1, target.max_row + 1):
        if to_int(get_cell(target, row, TGT_SUBJECT_CODE).value) != ACCOUNT_2219:
            continue
        summary.target_rows += 1
        _plan_row(target, row, tax_sets, summary, full_entries, tax_only_entries)

    for tax_only_entry in tax_only_entries:
        _write_tax_only_entry(target, tax_only_entry)

    # 行を追加すると下の行番号がずれるため、下の行から書き込む
    for full_entry in reversed(full_entries):
        _write_full_entry(target, full_entry)

    summary.full_entries = len(full_entries)
    summary.tax_only_entries = len(tax_only_entries)
    summary.inserted_rows = sum(len(entry.lines) - 1 for entry in full_entries)
    return summary


def _get_sheet(workbook: Workbook, name: str) -> Worksheet:
    if name not in workbook.sheetnames:
        raise ValueError(f"シート「{name}」が見つかりません。")
    sheet = workbook[name]
    if not isinstance(sheet, Worksheet):
        raise ValueError(f"「{name}」はワークシートではありません。")
    return sheet


def _plan_row(
    target: Worksheet,
    row: int,
    tax_sets: dict[int, TaxSet],
    summary: Summary,
    full_entries: list[FullEntry],
    tax_only_entries: list[TaxOnlyEntry],
) -> None:
    """1行分の処理内容を決める。シートへの書き込みはしない。"""
    if _has_tax_value(target, row):
        summary.already_done += 1
        return

    side = _find_side(target, row)
    if side is None:
        _skip(summary, row, "金額が借方・貸方のどちらか一方に入っていない")
        return

    if _has_rate_or_subject(target, row):
        _plan_tax_only_entry(target, row, side, summary, tax_only_entries)
    else:
        _plan_full_entry(target, row, side, tax_sets, summary, full_entries)


def _plan_tax_only_entry(
    target: Worksheet,
    row: int,
    side: Side,
    summary: Summary,
    entries: list[TaxOnlyEntry],
) -> None:
    percent = to_int(get_cell(target, row, TGT_RATE).value)
    if percent not in side.tax_columns:
        _skip(summary, row, "入力済みの税率が5・8・10ではない")
        return
    if is_empty(get_cell(target, row, side.subject_column).value):
        _skip(summary, row, f"{side.name}の科目が入っていない")
        return

    tax = _read_amount(target, row, side)
    entries.append(TaxOnlyEntry(row, side, percent, tax))
    logger.info("税額のみ入力: 元の行=%d %s 税率=%d%% 税額=%d", row, side.name, percent, tax)


def _plan_full_entry(
    target: Worksheet,
    row: int,
    side: Side,
    tax_sets: dict[int, TaxSet],
    summary: Summary,
    entries: list[FullEntry],
) -> None:
    voucher = to_int(get_cell(target, row, TGT_VOUCHER).value) or 0
    tax_set = tax_sets.get(voucher)
    if tax_set is None:
        _skip(summary, row, "科目シートに同じ伝票№の税行がない")
        return
    if not tax_set.details:
        _skip(summary, row, "税行の下に明細がない")
        return
    if tax_set.has_mixed_sides:
        _skip(summary, row, "明細の貸借が税行と揃っていない")
        return

    tax_amount = _read_amount(target, row, side)
    if tax_amount != tax_set.tax_row.debit + tax_set.tax_row.credit:
        _skip(summary, row, "科目シートと0384シートで税額が違う")
        return

    tax_class = str(get_cell(target, row, TGT_TAX_CLASS).value)
    candidates = TAX_CLASS_PERCENTS.get(tax_class)
    if candidates is None:
        _skip(summary, row, f"想定していない課税区分: {tax_class}")
        return

    details = tax_set.details
    base_amount = sum(detail.amount for detail in details)
    percent = choose_percent(candidates, base_amount, tax_amount)
    if not is_tax_valid(details, percent, tax_amount):
        _skip(summary, row, "明細合計×税率と税額が一致しない")
        logger.warning(
            "  金額確認NG: 伝票№=%d 明細合計=%d 税率=%d%% 税額=%d",
            voucher,
            base_amount,
            percent,
            tax_amount,
        )
        return

    lines = allocate_tax(details, percent, tax_amount)
    entries.append(FullEntry(row, side, percent, lines))
    logger.info(
        "入力: 元の行=%d 伝票№=%d %s 税率=%d%% 科目別税額=%s",
        row,
        voucher,
        side.name,
        percent,
        {line.subject: line.tax for line in lines},
    )


def _write_tax_only_entry(target: Worksheet, entry: TaxOnlyEntry) -> None:
    tax_column = entry.side.tax_columns[entry.percent]
    write_red(get_cell(target, entry.row, tax_column), entry.tax)


def _write_full_entry(target: Worksheet, entry: FullEntry) -> None:
    extra_rows = len(entry.lines) - 1
    if extra_rows > 0:
        insert_rows_below(target, entry.row, extra_rows)

    for offset, line in enumerate(entry.lines):
        row = entry.row + offset
        if offset > 0:
            _copy_row_header(target, entry.row, row)
        write_red(get_cell(target, row, TGT_RATE), entry.percent)
        write_red(get_cell(target, row, entry.side.subject_column), line.subject)
        write_red(get_cell(target, row, entry.side.tax_columns[entry.percent]), line.tax)


def _copy_row_header(target: Worksheet, source_row: int, new_row: int) -> None:
    """元の行のA〜G列を追加行へ複写し、A列を赤くする。"""
    first_cell = get_cell(target, new_row, 1)
    first_cell.value = get_cell(target, source_row, 1).value
    fill_red(first_cell)

    for column in range(2, TGT_COPY_LAST_COLUMN + 1):
        value = get_cell(target, source_row, column).value
        write_red(get_cell(target, new_row, column), _as_cell_value(value))


def _as_cell_value(value: object) -> int | str | None:
    if value is None:
        return None
    if isinstance(value, int | str):
        return value
    number = to_int(value)
    return number if number is not None else str(value)


def _find_side(target: Worksheet, row: int) -> Side | None:
    """金額が入っている側を返す。両方または両方なしの場合は None。"""
    has_debit = bool(to_int(get_cell(target, row, TGT_DEBIT).value))
    has_credit = bool(to_int(get_cell(target, row, TGT_CREDIT).value))
    if has_debit == has_credit:
        return None
    return DEBIT_SIDE if has_debit else CREDIT_SIDE


def _read_amount(target: Worksheet, row: int, side: Side) -> int:
    return to_int(get_cell(target, row, side.amount_column).value) or 0


def _has_tax_value(target: Worksheet, row: int) -> bool:
    """税額列のどこかに値があれば True。"""
    columns = [*TGT_DEBIT_TAX_COLUMNS.values(), *TGT_CREDIT_TAX_COLUMNS.values()]
    return any(not is_empty(get_cell(target, row, column).value) for column in columns)


def _has_rate_or_subject(target: Worksheet, row: int) -> bool:
    """税率・科目のどれかに値があれば True。"""
    columns = [TGT_RATE, TGT_DEBIT_SUBJECT, TGT_CREDIT_SUBJECT]
    return any(not is_empty(get_cell(target, row, column).value) for column in columns)


def _skip(summary: Summary, row: int, reason: str) -> None:
    summary.add_skip(reason)
    logger.warning("未入力: 元の行=%d 理由=%s", row, reason)
