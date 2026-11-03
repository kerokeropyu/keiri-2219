"""税額の確認と、科目ごとの税額配分。"""

from dataclasses import dataclass


@dataclass(frozen=True)
class Detail:
    """仮払消費税の行に対応する明細行。"""

    subject: int
    amount: int


@dataclass(frozen=True)
class TaxLine:
    """科目ごとの税額。"""

    subject: int
    tax: int


def choose_percent(candidates: list[int], base_amount: int, tax_amount: int) -> int:
    """候補の税率から、税額に最も近いものを返す。"""
    return min(candidates, key=lambda percent: abs(base_amount * percent - tax_amount * 100))


def is_tax_valid(details: list[Detail], percent: int, tax_amount: int) -> bool:
    """明細合計×税率と税額の差が、明細の行数（円）以内なら正しいとみなす。

    明細ごとに端数処理されている場合、1行につき最大1円ずれるため。
    """
    base_amount = sum(detail.amount for detail in details)
    difference = abs(base_amount * percent - tax_amount * 100)
    return difference <= len(details) * 100


def allocate_tax(details: list[Detail], percent: int, tax_amount: int) -> list[TaxLine]:
    """税額を科目別に配分する。端数は金額が最大の科目で調整する。"""
    subject_totals: dict[int, int] = {}
    for detail in details:
        subject_totals[detail.subject] = subject_totals.get(detail.subject, 0) + detail.amount

    subject_taxes = {subject: total * percent // 100 for subject, total in subject_totals.items()}

    largest_subject = max(subject_totals, key=lambda subject: subject_totals[subject])
    subject_taxes[largest_subject] += tax_amount - sum(subject_taxes.values())

    return [TaxLine(subject, tax) for subject, tax in subject_taxes.items()]
