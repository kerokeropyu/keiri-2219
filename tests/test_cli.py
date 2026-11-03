from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from keiri_2219.cli import main
from keiri_2219.constants import SOURCE_SHEET, TARGET_SHEET


def test_xls_input_is_rejected(tmp_path: Path) -> None:
    input_path = tmp_path / "test.xls"
    input_path.write_bytes(b"dummy")

    exit_code = main([str(input_path), "--log-dir", str(tmp_path / "logs")])

    assert exit_code == 1


def test_mismatch_does_not_create_output(tmp_path: Path) -> None:
    workbook = Workbook()
    source = workbook.active
    assert isinstance(source, Worksheet)
    source.title = SOURCE_SHEET
    target = workbook.create_sheet(TARGET_SHEET)
    for _ in range(3):
        target.append(["見出し"])
    # 入力済みの行（税額999）が借方1000と一致しない
    target.append([384, 250731, 7, 100, 2219, "", "", 10, None, 1000, 7199, None, None, 999])
    input_path = tmp_path / "test.xlsx"
    workbook.save(input_path)

    exit_code = main([str(input_path), "--log-dir", str(tmp_path / "logs")])

    assert exit_code == 1
    assert not (tmp_path / "test_2219処理済み.xlsx").exists()
