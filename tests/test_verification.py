from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from keiri_2219.constants import TARGET_SHEET
from keiri_2219.verification import verify_tax_amounts

# 列: D=伝票№, E=相手科目, J=借方, N=借方10%, O=貸方, S=貸方10%


def _build_target(rows: list[dict[str, int]]) -> Workbook:
    workbook = Workbook()
    sheet = workbook.active
    assert isinstance(sheet, Worksheet)
    sheet.title = TARGET_SHEET
    for _ in range(3):
        sheet.append(["見出し"])
    for row_number, values in enumerate(rows, start=4):
        sheet[f"E{row_number}"] = 2219
        for column, value in values.items():
            sheet[f"{column}{row_number}"] = value
    return workbook


def test_debit_credit_and_inserted_row_pass() -> None:
    workbook = _build_target(
        [
            {"D": 100, "J": 3000, "N": 2000},
            {"D": 100, "N": 1000},  # 行追加した行（J列は空）
            {"D": 200, "O": 500, "S": 500},
            {"D": 300, "J": 700},  # 税額が空の行は対象外
        ]
    )

    result = verify_tax_amounts(workbook)

    assert result.mismatches == []
    assert result.checked_vouchers == 2


def test_mismatch_is_detected() -> None:
    workbook = _build_target(
        [
            {"D": 100, "J": 3000, "N": 2000},
            {"D": 100, "N": 999},
        ]
    )

    result = verify_tax_amounts(workbook)

    assert len(result.mismatches) == 1
    mismatch = result.mismatches[0]
    assert mismatch.voucher == 100
    assert mismatch.rows == [4, 5]
    assert mismatch.side_name == "借方"
    assert mismatch.tax_total == 2999
    assert mismatch.amount == 3000
