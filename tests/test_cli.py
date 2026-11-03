from pathlib import Path

from keiri_2219.cli import main


def test_xls_input_is_rejected(tmp_path: Path) -> None:
    input_path = tmp_path / "test.xls"
    input_path.write_bytes(b"dummy")

    exit_code = main([str(input_path), "--log-dir", str(tmp_path / "logs")])

    assert exit_code == 1
