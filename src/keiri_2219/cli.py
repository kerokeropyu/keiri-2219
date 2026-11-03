"""コマンドラインの入口。"""

import argparse
import logging
import sys
from pathlib import Path

from openpyxl import load_workbook

from keiri_2219.logging_setup import setup_logging
from keiri_2219.processor import Summary, process_workbook
from keiri_2219.verification import VerificationResult, verify_tax_amounts

logger = logging.getLogger(__name__)


def main(argv: list[str] | None = None) -> int:
    if not _is_exe():
        return _run(argv)

    # exeはドラッグ＆ドロップで起動されるため、結果を読めるように終了前に待つ
    exe_args = sys.argv[1:] if argv is None else argv
    try:
        if not exe_args:
            _print_exe_usage()
            return 1
        return _run(exe_args)
    finally:
        _wait_for_enter()


def _run(argv: list[str] | None) -> int:
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
        verification = verify_tax_amounts(workbook)
        if verification.mismatches:
            _print_mismatch_error(verification, args.log_dir)
            return 1
        print("ファイルを保存しています...")
        workbook.save(output_path)
    except (OSError, ValueError) as error_detail:
        logger.exception("処理に失敗しました")
        print(f"エラー: {error_detail}")
        return 1

    logger.info("処理終了: %s", output_path)
    _print_summary(summary, verification, output_path, args.log_dir)
    return 0


def _is_exe() -> bool:
    """PyInstallerで作ったexeから実行していれば True。"""
    return bool(getattr(sys, "frozen", False))


def _default_log_dir() -> Path:
    """exeならexeと同じフォルダの logs、それ以外は実行場所の logs。"""
    if _is_exe():
        return Path(sys.executable).parent / "logs"
    return Path("logs")


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="keiri-2219", description="経理の2219処理を行います。")
    parser.add_argument("input", type=Path, help="入力ファイル（.xlsx）")
    parser.add_argument("-o", "--output", type=Path, help="出力ファイル（省略時は自動で命名）")
    parser.add_argument("--log-dir", type=Path, default=_default_log_dir(), help="ログの保存先")
    return parser


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    return _build_parser().parse_args(argv)


def _print_exe_usage() -> None:
    print("使い方: 処理する.xlsxファイルを keiri-2219.exe にドラッグ＆ドロップしてください。")
    print()
    _build_parser().print_help()


def _wait_for_enter() -> None:
    print()
    try:
        input("Enterキーで終了します")
    except EOFError:
        pass


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


def _print_mismatch_error(verification: VerificationResult, log_dir: Path) -> None:
    print(
        f"エラー: 税額の合計が元の金額と一致しない伝票が{len(verification.mismatches)}件あります。"
    )
    for mismatch in verification.mismatches:
        print(
            f"  伝票№={mismatch.voucher} 行={mismatch.rows} {mismatch.side_name}"
            f" 税額合計={mismatch.tax_total} 元の金額={mismatch.amount}"
        )
    print("ファイルは保存していません。")
    print(f"詳しくはログを確認してください: {log_dir}")


def _print_summary(
    summary: Summary, verification: VerificationResult, output_path: Path, log_dir: Path
) -> None:
    print(f"完了: {output_path}")
    print(f"  対象行（E列2219）: {summary.target_rows}件")
    print(f"  税率・科目・税額を入力: {summary.full_entries}件")
    print(f"  税額のみ入力: {summary.tax_only_entries}件")
    print(f"  追加した行: {summary.inserted_rows}行")
    print(f"  税額入力済みのため変更なし: {summary.already_done}件")
    print(f"  税額の一致チェック: {verification.checked_vouchers}伝票すべて一致")
    for reason, count in summary.skipped.items():
        print(f"  未入力（{reason}）: {count}件")
    print(f"未入力の行はログを確認してください: {log_dir}")


if __name__ == "__main__":
    raise SystemExit(main())
