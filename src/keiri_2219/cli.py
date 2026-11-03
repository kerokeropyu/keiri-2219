"""コマンドラインの入口。"""

import argparse
import logging
from pathlib import Path

from openpyxl import load_workbook

from keiri_2219.logging_setup import setup_logging
from keiri_2219.processor import Summary, process_workbook

logger = logging.getLogger(__name__)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    input_path: Path = args.input
    output_path: Path = args.output or _default_output_path(input_path)

    setup_logging(args.log_dir)

    error = _validate_paths(input_path, output_path)
    if error:
        print(f"エラー: {error}")
        return 1

    logger.info("処理開始: %s", input_path)
    print("ファイルを読み込んでいます...")
    try:
        workbook = load_workbook(input_path)
        summary = process_workbook(workbook)
        print("ファイルを保存しています...")
        workbook.save(output_path)
    except (OSError, ValueError) as error_detail:
        logger.exception("処理に失敗しました")
        print(f"エラー: {error_detail}")
        return 1

    logger.info("処理終了: %s", output_path)
    _print_summary(summary, output_path, args.log_dir)
    return 0


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="経理の2219処理を行います。")
    parser.add_argument("input", type=Path, help="入力ファイル（.xlsx）")
    parser.add_argument("-o", "--output", type=Path, help="出力ファイル（省略時は自動で命名）")
    parser.add_argument("--log-dir", type=Path, default=Path("logs"), help="ログの保存先")
    return parser.parse_args(argv)


def _default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_2219処理済み{input_path.suffix}")


def _validate_paths(input_path: Path, output_path: Path) -> str | None:
    """問題があればエラー文を返す。"""
    if not input_path.is_file():
        return f"ファイルが見つかりません: {input_path}"
    if input_path.suffix.lower() != ".xlsx":
        return "入力は.xlsxのみ対応です。Excelで.xlsx形式に保存してから実行してください。"
    if output_path.exists():
        return f"出力先がすでに存在します: {output_path}"
    return None


def _print_summary(summary: Summary, output_path: Path, log_dir: Path) -> None:
    print(f"完了: {output_path}")
    print(f"  対象行（E列2219）: {summary.target_rows}件")
    print(f"  税率・科目・税額を入力: {summary.full_entries}件")
    print(f"  税額のみ入力: {summary.tax_only_entries}件")
    print(f"  追加した行: {summary.inserted_rows}行")
    print(f"  税額入力済みのため変更なし: {summary.already_done}件")
    for reason, count in summary.skipped.items():
        print(f"  未入力（{reason}）: {count}件")
    print(f"未入力の行はログを確認してください: {log_dir}")


if __name__ == "__main__":
    raise SystemExit(main())
