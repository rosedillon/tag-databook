from __future__ import annotations

import itertools

import pytest

from tag_databook import TagDataBook, defaults


@pytest.fixture
def book() -> TagDataBook:
    return TagDataBook.with_defaults()


def test_defaults_need_no_workbook(book: TagDataBook) -> None:
    assert book.price_year == defaults.PRICE_YEAR
    assert book.vot_work_rail == pytest.approx(38.22)
    assert book.fallbacks == ()
    assert "default" in book.source


def test_discount_bands() -> None:
    bands = defaults.discount_rates().standard
    for (_, previous_high, previous_rate), (low, _, rate) in itertools.pairwise(bands):
        assert low == previous_high + 1, "bands must not leave a gap"
        assert rate < previous_rate, "Green Book rates step downward"


def test_health_variant_step_down() -> None:
    rates = defaults.discount_rates()
    for series in (rates.health, rates.variant):
        assert [rate for _, _, rate in series] == sorted(
            (rate for _, _, rate in series), reverse=True
        )
        assert series[-1][2] < series[0][2]


def test_health_variant_bands() -> None:
    rates = defaults.discount_rates()
    spans = {
        tuple((low, high) for low, high, _ in series)
        for series in (rates.standard, rates.health, rates.variant)
    }
    assert len(spans) == 1


def test_carbon_span() -> None:
    carbon = defaults.carbon_values(first_year=2025, last_year=2030)
    assert min(carbon.factor_cost) == 2025
    assert max(carbon.factor_cost) == 2030


def test_carbon_low_central_high() -> None:
    carbon = defaults.carbon_values()
    low, central, high = carbon.factor_cost[2030]
    assert low < central < high


def test_carbon_base_year() -> None:
    carbon = defaults.carbon_values()
    assert carbon.factor_cost[defaults.PRICE_YEAR][1] == pytest.approx(
        defaults.CARBON_CENTRAL_AT_BASE
    )


def test_no_deflation_on_defaults(book: TagDataBook) -> None:
    # There is no deflator history without the workbook; the error says
    # so rather than returning a plausible-looking number.
    with pytest.raises(KeyError, match="2015"):
        book.deflate(100.0, 2015, 2023)


def test_defaults_interface(book: TagDataBook) -> None:
    assert book.discount_factor(10) < 1.0
    assert book.carbon_price(2030) > 0.0
    assert book.car_vot_by_distance(100.0) > 0.0
