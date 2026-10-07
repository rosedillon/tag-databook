"""Representative TAG v2.03 values, for running without the workbook.

These are figures from the November 2023 TAG Data Book, written in so
that a screening appraisal can run without the ``.xlsm``. They are also
the fallback for any cell that cannot be read from a workbook that is
loaded.

A formal business case should still read the current workbook. These
values are frozen at one release and go out of date as DfT publishes
new ones.
"""

from __future__ import annotations

from tag_databook.tables import CarbonValues, DiscountRates, GDPDeflator, ValueOfTime

PRICE_YEAR = 2023

# Green Book declining long-run discount bands, in years from the base
# year.
DISCOUNT_BANDS: tuple[tuple[int, int], ...] = (
    (0, 30),
    (31, 75),
    (76, 125),
    (126, 200),
    (201, 300),
    (301, 999),
)

# Standard Green Book rate for each band in `DISCOUNT_BANDS`.
STANDARD_RATES: tuple[float, ...] = (0.035, 0.030, 0.025, 0.020, 0.015, 0.010)

# Rates for health effects, for each band in `DISCOUNT_BANDS`.
HEALTH_RATES: tuple[float, ...] = (0.015, 0.0129, 0.0107, 0.0086, 0.0064, 0.0043)

# The Green Book "variant" (risk to life) rates, for each band.
VARIANT_RATES: tuple[float, ...] = (0.030, 0.0257, 0.0214, 0.0171, 0.0129, 0.0086)

# Real-terms uplift applied to values of time each year (TAG Unit
# A1.1).
VOT_ANNUAL_GROWTH = 0.015

VALUE_OF_TIME = ValueOfTime(
    work_car_driver=23.17,
    work_car_passenger=23.17,
    work_rail=38.22,
    work_bus=13.13,
    work_walk=13.13,
    work_cycle=13.13,
    work_average=25.23,
    car_u_factor_cost=38.65895977419787,
    car_xmid=66.5258,
    car_k=67.0174,
    commuting=13.04,
    other=5.95,
    price_year=PRICE_YEAR,
)


def discount_rates() -> DiscountRates:
    """Green Book stepped discount rates built from the tables above."""

    def bands(rates: tuple[float, ...]) -> tuple[tuple[int, int, float], ...]:
        return tuple(
            (low, high, rate) for (low, high), rate in zip(DISCOUNT_BANDS, rates, strict=True)
        )

    return DiscountRates(
        standard=bands(STANDARD_RATES),
        health=bands(HEALTH_RATES),
        variant=bands(VARIANT_RATES),
    )


# Central carbon value in `PRICE_YEAR`, pounds per tCO2e at factor
# cost.
CARBON_CENTRAL_AT_BASE = 284.79

# Real growth in the carbon value per year. Grown from the base value
# at this rate, the series stays within 0.4% of the published one from
# 2020 to 2100.
CARBON_ANNUAL_GROWTH = 0.015

# The low and high values TAG gives either side of the central one.
CARBON_LOW_MULTIPLIER = 0.5
CARBON_HIGH_MULTIPLIER = 1.5


def carbon_values(
    first_year: int = 2020,
    last_year: int = 2100,
) -> CarbonValues:
    """Build a carbon series by growing the base-year central value.

    The published series grows at almost exactly 1.5% a year, so this
    stays within 0.4% of it between 2020 and 2100. It is still an
    approximation, so read the workbook when carbon matters to the
    result.
    """
    factor_cost: dict[int, tuple[float, float, float]] = {}
    for year in range(first_year, last_year + 1):
        central = CARBON_CENTRAL_AT_BASE * (CARBON_ANNUAL_GROWTH + 1.0) ** (year - PRICE_YEAR)
        factor_cost[year] = (
            central * CARBON_LOW_MULTIPLIER,
            central,
            central * CARBON_HIGH_MULTIPLIER,
        )
    return CarbonValues(factor_cost=factor_cost, price_year=PRICE_YEAR)


def gdp_deflator() -> GDPDeflator:
    """A single-point deflator series: the base year, indexed at 100.

    Without the workbook there is no deflator history, so nothing can be
    converted between price years. Check for the year you need before
    relying on the series.
    """
    return GDPDeflator(series={PRICE_YEAR: 100.0}, base_year=PRICE_YEAR)
