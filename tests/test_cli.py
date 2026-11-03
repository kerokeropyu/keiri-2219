from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from keiri_2219.cli import main
from keiri_2219.constants import SOURCE_SHEET, TARGET_SHEET


def test_xls_input_is_rejected(tmp_path: Path) -> None:
    input_path = tmp_path / "test.xls"
    input_path.write_bytes(b"dummy")

    exit_code = main([str(input_path), "--log-dir", str(tmp_path / "logs")])

    assert exit_code == 1


def _save_workbook(path: Path, target_rows: list[list[object]]) -> None:
    workbook = Workbook()
    source = workbook.active
    assert isinstance(source, Worksheet)
    source.title = SOURCE_SHEET
    target = workbook.create_sheet(TARGET_SHEET)
    for _ in range(3):
        target.append(["見出し"])
    for row in target_rows:
        target.append(row)
    workbook.save(path)


def test_mismatch_does_not_create_output(tmp_path: Path) -> None:
    input_path = tmp_path / "test.xlsx"
    # 入力済みの行（税額999）が借方1000と一致しない
    _save_workbook(
        input_path,
        [[384, 250731, 7, 100, 2219, "", "", 10, None, 1000, 7199, None, None, 999]],
    )

    exit_code = main([str(input_path), "--log-dir", str(tmp_path / "logs")])

    assert exit_code == 1
    assert not (tmp_path / "test_2219処理済み.xlsx").exists()


def test_existing_output_gets_numbered_name(tmp_path: Path) -> None:
    input_path = tmp_path / "test.xlsx"
    _save_workbook(input_path, [])
    (tmp_path / "test_2219処理済み.xlsx").write_bytes(b"old")
    (tmp_path / "test_2219処理済み(2).xlsx").write_bytes(b"old")

    exit_code = main([str(input_path), "--log-dir", str(tmp_path / "logs")])

    assert exit_code == 0
    assert (tmp_path / "test_2219処理済み(3).xlsx").exists()
    assert (tmp_path / "test_2219処理済み.xlsx").read_bytes() == b"old"


def test_permission_error_shows_excel_hint(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    input_path = tmp_path / "test.xlsx"
    _save_workbook(input_path, [])

    def raise_permission_error(self: Workbook, filename: str) -> None:
        raise PermissionError(13, "Permission denied", filename)

    monkeypatch.setattr(Workbook, "save", raise_permission_error)

    exit_code = main([str(input_path), "--log-dir", str(tmp_path / "logs")])

    assert exit_code == 1
    assert "Excelでファイルを開いていませんか" in capsys.readouterr().out


def test_unexpected_error_is_logged(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    # .xlsの中身のまま拡張子だけ.xlsxにしたファイルを想定
    input_path = tmp_path / "test.xlsx"
    input_path.write_bytes(b"not a zip file")
    log_dir = tmp_path / "logs"

    exit_code = main([str(input_path), "--log-dir", str(log_dir)])

    assert exit_code == 1
    assert "開発担当者に連絡してください" in capsys.readouterr().out
    log_text = (log_dir / "keiri_2219.log").read_text(encoding="utf-8")
    assert "想定外のエラーが発生しました" in log_text
