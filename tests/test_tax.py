from keiri_2219.tax import Detail, TaxLine, allocate_tax, choose_percent, is_tax_valid


def test_is_tax_valid_with_issue_example() -> None:
    details = [Detail(6523, amount) for amount in [108900] * 6 + [253600]]
    assert is_tax_valid(details, 10, 90700)


def test_is_tax_valid_rejects_large_difference() -> None:
    assert not is_tax_valid([Detail(7112, 25900)], 10, 2592)


def test_choose_percent_for_outside_tax() -> None:
    assert choose_percent([10, 8], 1000, 80) == 8


def test_allocate_tax_adjusts_fraction_on_largest_subject() -> None:
    details = [Detail(7117, 11844), Detail(7199, 3417)]
    assert allocate_tax(details, 10, 1525) == [TaxLine(7117, 1184), TaxLine(7199, 341)]
    assert allocate_tax(details, 10, 1526) == [TaxLine(7117, 1185), TaxLine(7199, 341)]
