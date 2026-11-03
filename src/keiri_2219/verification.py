"""保存前の確認。シートに書き込まれた税額の合計が、元の金額と一致するかを調べる。"""

import logging
from collections.abc import Iterable
from dataclasses import dataclass, field

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from keiri_2219.constants import (
    ACCOUNT_2219,
    HEADER_ROWS,
    TARGET_SHEET,
    TGT_CREDIT,
    TGT_CREDIT_TAX_COLUMNS,
    TGT_DEBIT,
    TGT_DEBIT_TAX_COLUMNS,
    TGT_SUBJECT_CODE,
    TGT_VOUCHER,
)
from keiri_2219.excel_utils import get_cell, is_empty, to_int

logger = logging.getLogger(__name__)


@dataclass
class VoucherTotal:
    """1つの伝票№について、税額と元の金額を借方・貸方別に合計したもの。"""

    rows: list[int] = field(default_factory=list)
    debit_tax: int = 0
    debit_amount: int = 0
    credit_tax: int = 0
    credit_amount: int = 0


@dataclass(frozen=True)
class Mismatch:
    """税額の合計が元の金額と一致しない伝票。"""

    voucher: int
    rows: list[int]
    side_name: str
    tax_total: int
    amount: int


@dataclass(frozen=True)
class VerificationResult:
    checked_vouchers: int
    mismatches: list[Mismatch]


def verify_tax_amounts(workbook: Workbook) -> VerificationResult:
    """「0384仮払い消費税」シートで、E列が2219で税額が入っている行を伝票№ごとに確認する。

    行を追加した伝票は、元の行と追加行が同じ伝票№なので合算して比べる。
    """
    target = workbook[TARGET_SHEET]
    if not isinstance(target, Worksheet):
        raise ValueError(f"「{TARGET_SHEET}」はワークシートではありません。")
    totals = _sum_by_voucher(target)

    mismatches: list[Mismatch] = []
    for voucher, total in totals.items():
        if total.debit_tax != total.debit_amount:
            mismatches.append(
                Mismatch(voucher, total.rows, "借方", total.debit_tax, total.debit_amount)
            )
        if total.credit_tax != total.credit_amount:
            mismatches.append(
                Mismatch(voucher, total.rows, "貸方", total.credit_tax, total.credit_amount)
            )

    for mismatch in mismatches:
        logger.error(
            "一致チェックNG: 伝票№=%d 行=%s %s 税額合計=%d 元の金額=%d",
            mismatch.voucher,
            mismatch.rows,
            mismatch.side_name,
            mismatch.tax_total,
            mismatch.amount,
        )
    logger.info("一致チェック: 伝票%d件 不一致%d件", len(totals), len(mismatches))
    return VerificationResult(len(totals), mismatches)


def _sum_by_voucher(target: Worksheet) -> dict[int, VoucherTotal]:
    totals: dict[int, VoucherTotal] = {}
    for row in range(HEADER_ROWS + 1, target.max_row + 1):
        if to_int(get_cell(target, row, TGT_SUBJECT_CODE).value) != ACCOUNT_2219:
            continue
        if not _has_tax_value(target, row):
            continue

        voucher = _read_int(target, row, TGT_VOUCHER)
        total = totals.setdefault(voucher, VoucherTotal())
        total.rows.append(row)
        total.debit_tax += _sum_columns(target, row, TGT_DEBIT_TAX_COLUMNS.values())
        total.debit_amount += _read_int(target, row, TGT_DEBIT)
        total.credit_tax += _sum_columns(target, row, TGT_CREDIT_TAX_COLUMNS.values())
        total.credit_amount += _read_int(target, row, TGT_CREDIT)
    return totals


def _has_tax_value(target: Worksheet, row: int) -> bool:
    columns = [*TGT_DEBIT_TAX_COLUMNS.values(), *TGT_CREDIT_TAX_COLUMNS.values()]
    return any(not is_empty(get_cell(target, row, column).value) for column in columns)


def _sum_columns(target: Worksheet, row: int, columns: Iterable[int]) -> int:
    return sum(_read_int(target, row, column) for column in columns)


def _read_int(target: Worksheet, row: int, column: int) -> int:
    return to_int(get_cell(target, row, column).value) or 0
