"""「科目２０００～５０００未満」シートから税行と明細のセットを読み込む。"""

from dataclasses import dataclass

from openpyxl.worksheet.worksheet import Worksheet

from keiri_2219.constants import (
    ACCOUNT_2219,
    ACCOUNT_TEMP_TAX,
    HEADER_ROWS,
    SRC_ACCOUNT,
    SRC_CREDIT,
    SRC_DEBIT,
    SRC_PARENT_VOUCHER,
    SRC_SUBJECT,
    SRC_VOUCHER,
)
from keiri_2219.excel_utils import get_cell, to_int
from keiri_2219.tax import Detail


@dataclass(frozen=True)
class LedgerRow:
    """元帳の1行。"""

    row: int
    voucher: int
    subject: int
    parent_voucher: int
    debit: int
    credit: int


@dataclass(frozen=True)
class TaxSet:
    """仮払消費税の行と、親伝票№で紐づく明細のセット。"""

    tax_row: LedgerRow
    detail_rows: list[LedgerRow]

    @property
    def has_mixed_sides(self) -> bool:
        """明細に、税行と逆側の金額があれば True。"""
        if self.tax_row.credit:
            return any(row.debit for row in self.detail_rows)
        return any(row.credit for row in self.detail_rows)

    @property
    def details(self) -> list[Detail]:
        return [Detail(row.subject, row.debit + row.credit) for row in self.detail_rows]


def read_ledger_rows(sheet: Worksheet) -> list[LedgerRow]:
    """A列が2219の行を読み込む。"""
    rows: list[LedgerRow] = []
    for row in range(HEADER_ROWS + 1, sheet.max_row + 1):
        if to_int(get_cell(sheet, row, SRC_ACCOUNT).value) != ACCOUNT_2219:
            continue
        rows.append(
            LedgerRow(
                row=row,
                voucher=_read_int(sheet, row, SRC_VOUCHER),
                subject=_read_int(sheet, row, SRC_SUBJECT),
                parent_voucher=_read_int(sheet, row, SRC_PARENT_VOUCHER),
                debit=_read_int(sheet, row, SRC_DEBIT),
                credit=_read_int(sheet, row, SRC_CREDIT),
            )
        )
    return rows


def build_tax_sets(rows: list[LedgerRow]) -> dict[int, TaxSet]:
    """税行の伝票№をキーに、税行と下の明細をセットにする。

    明細の親伝票№（L列）が、税行の伝票№と同じものをセットとして扱う。
    """
    details_by_parent: dict[int, list[LedgerRow]] = {}
    for row in rows:
        if row.subject != ACCOUNT_TEMP_TAX:
            details_by_parent.setdefault(row.parent_voucher, []).append(row)

    tax_sets: dict[int, TaxSet] = {}
    for row in rows:
        if row.subject != ACCOUNT_TEMP_TAX:
            continue
        related_details = details_by_parent.get(row.voucher, [])
        details_below = [detail for detail in related_details if detail.row > row.row]
        tax_sets[row.voucher] = TaxSet(row, details_below)
    return tax_sets


def _read_int(sheet: Worksheet, row: int, column: int) -> int:
    return to_int(get_cell(sheet, row, column).value) or 0
