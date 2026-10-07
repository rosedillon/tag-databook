"""Value types for the parameters read out of the TAG Data Book.

Each class mirrors one table in the workbook. They are plain frozen
dataclasses with no behaviour and no link to the workbook, so a parsed
data book can be pickled, compared or built by hand in a test without
openpyxl.
"""

from __future__ import annotations

from dataclasses import dataclass, field

# (years_from_base_low, years_from_base_high, rate)
DiscountBand = tuple[int, int, float]


@dataclass(frozen=True, slots=True)
class ValueOfTime:
    """Values of time from TAG Table A 1.3.1, in its price year.

    Figures are factor costs in pounds per hour. Working time values are
    per person travelling on employer's business, and ``commuting`` and
    ``other`` are the two non-work purposes TAG distinguishes.
    """

    work_car_driver: float
    work_car_passenger: float
    work_rail: float
    work_bus: float
    work_walk: float
    work_cycle: float
    work_average: float

    # Parameters of TAG's logistic car VoT-by-distance curve (see
    # `TagDataBook.car_vot_by_distance`), not a per-hour value on their
    # own.
    car_u_factor_cost: float
    car_xmid: float
    car_k: float

    commuting: float
    other: float

    price_year: int


@dataclass(frozen=True, slots=True)
class DiscountRates:
    """Green Book stepped discount rates from TAG Table A 1.1.1.

    Each series is a tuple of ``(from_year, to_year, rate)`` bands, with
    years counted from the appraisal base year. The Green Book steps the
    rate down over very long horizons, so it is a set of bands rather
    than one number.
    """

    standard: tuple[DiscountBand, ...]
    health: tuple[DiscountBand, ...]
    variant: tuple[DiscountBand, ...]


@dataclass(frozen=True, slots=True)
class CarbonValues:
    """Carbon values from TAG Table A 3.4.1, in pounds per tCO2e.

    Keyed by calendar year, each entry is ``(low, central, high)``. The
    series rises by about 1.5% a year, so the central value is about 2.4
    times higher at the end of a 60-year appraisal than at the start.
    """

    factor_cost: dict[int, tuple[float, float, float]]
    price_year: int


@dataclass(frozen=True, slots=True)
class GDPDeflator:
    """GDP deflator index from the workbook's *Annual Parameters* sheet.

    Used to move money between price years. The series is indexed so
    that ``base_year`` equals 100.
    """

    series: dict[int, float] = field(default_factory=dict)
    base_year: int = 2023
