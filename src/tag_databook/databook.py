"""Read a DfT TAG Data Book workbook into plain Python values.

The workbook is laid out for people to read. Values sit at fixed cells
under merged headers and footnotes, and there is no machine-readable
schema, so this module hard-codes a layout (the sheet name and cell for
each parameter) and reads against it.

The layout was checked against TAG Data Book v2.03 (November 2023).
DfT rearranges the workbook between releases, so in another version a
value may have moved without any error. Two things guard against that.
Every parameter has a fallback in :mod:`tag_databook.defaults`, so a
missing cell gives a published figure instead of a crash. And every
fallback taken is recorded on :attr:`TagDataBook.fallbacks`, so a
changed layout shows up as a long list instead of as plausible numbers.

Check ``databook.fallbacks`` before trusting a run against an
unfamiliar release. If it is not empty, the layout needs updating.
"""

from __future__ import annotations

import logging
import math
from dataclasses import dataclass
from typing import TYPE_CHECKING, Literal

import openpyxl

from tag_databook import defaults
from tag_databook.tables import CarbonValues, DiscountRates, GDPDeflator, ValueOfTime

if TYPE_CHECKING:
    import pathlib
    from collections.abc import Iterator

logger = logging.getLogger(__name__)

CarbonScenario = Literal["low", "central", "high"]

SHEET_VALUE_OF_TIME = "A1.3.1"
SHEET_DISCOUNT = "A1.1.1"
SHEET_CARBON = "A3.4"
SHEET_ANNUAL_PARAMETERS = "Annual Parameters"

# Cell holding the price year, on both the VoT and carbon sheets.
PRICE_YEAR_CELL = "B14"

# Value-of-time field -> cell on `SHEET_VALUE_OF_TIME`. Column D is the
# factor-cost column; column I holds the car distance-curve parameters.
VALUE_OF_TIME_CELLS: dict[str, str] = {
    "work_car_driver": "D28",
    "work_car_passenger": "D29",
    "work_bus": "D34",
    "work_rail": "D37",
    "work_walk": "D39",
    "work_cycle": "D40",
    "work_average": "D42",
    "car_u_factor_cost": "I40",
    "car_xmid": "I43",
    "car_k": "I44",
    "commuting": "D47",
    "other": "D48",
}

# First row of the discount-rate band table on `SHEET_DISCOUNT`. One
# row per band in `defaults.DISCOUNT_BANDS`, in the same order.
DISCOUNT_FIRST_ROW = 25
DISCOUNT_COLUMNS = {"standard": "D", "health": "F", "variant": "H"}

# Row span of the carbon year table, and its year/low/central/high
# columns.
CARBON_ROWS = range(26, 118)
CARBON_YEAR_COLUMN = "B"
CARBON_VALUE_COLUMNS = ("D", "E", "F")

# Row span of the GDP deflator series, and its year/index columns.
DEFLATOR_ROWS = range(45, 160)
DEFLATOR_YEAR_COLUMN = "B"
DEFLATOR_INDEX_COLUMN = "D"


@dataclass(frozen=True, slots=True)
class Fallback:
    """A parameter that was not readable from the workbook.

    Recorded rather than raised, so one moved cell does not block an
    otherwise usable appraisal, while a moved sheet still shows up as a
    long list of these.
    """

    parameter: str
    sheet: str
    coordinate: str
    value_used: float

    def __str__(self) -> str:
        return (
            f"{self.parameter}: {self.sheet}!{self.coordinate} unreadable, "
            f"used {self.value_used:g}"
        )


class _Reader:
    """Reads numeric cells, recording any it cannot read."""

    def __init__(self, workbook: openpyxl.Workbook) -> None:
        self._workbook = workbook
        self.fallbacks: list[Fallback] = []

    def sheet_names(self) -> set[str]:
        return set(self._workbook.sheetnames)

    def number(self, sheet: str, coordinate: str, *, parameter: str, fallback: float) -> float:
        """Read one numeric cell, recording a fallback on failure."""
        value = self._raw(sheet, coordinate)
        if value is None:
            self.fallbacks.append(Fallback(parameter, sheet, coordinate, fallback))
            return fallback
        return value

    def optional_number(self, sheet: str, coordinate: str) -> float | None:
        """Read one numeric cell; ``None`` if it is blank or text.

        Used when scanning a table of unknown length, where a blank
        means the table has ended.
        """
        return self._raw(sheet, coordinate)

    def _raw(self, sheet: str, coordinate: str) -> float | None:
        if sheet not in self._workbook.sheetnames:
            return None
        raw = self._workbook[sheet][coordinate].value
        if raw is None or isinstance(raw, bool):
            return None
        try:
            return float(raw)
        except (TypeError, ValueError):
            return None


class TagDataBook:
    """Appraisal parameters read from a TAG Data Book workbook.

    Construct with :meth:`from_workbook` to read a real ``.xlsm``, or
    :meth:`with_defaults` to run without one.

    Attributes
    ----------
    vot, discount, carbon, gdp
        The four parsed tables. See :mod:`tag_databook.tables`.
    vot_growth_rate
        Real-terms annual growth applied to values of time (TAG Unit
        A1.1).
    source
        Where the values came from, for reporting.
    fallbacks
        Parameters that could not be read and fell back to a published
        figure. Empty for a workbook whose layout matches; see the
        module docstring.
    """

    def __init__(
        self,
        *,
        vot: ValueOfTime,
        discount: DiscountRates,
        carbon: CarbonValues,
        gdp: GDPDeflator,
        source: str,
        fallbacks: tuple[Fallback, ...] = (),
        vot_growth_rate: float = defaults.VOT_ANNUAL_GROWTH,
    ) -> None:
        self.vot = vot
        self.discount = discount
        self.carbon = carbon
        self.gdp = gdp
        self.source = source
        self.fallbacks = fallbacks
        self.vot_growth_rate = vot_growth_rate

    @classmethod
    def from_workbook(cls, path: str | pathlib.Path) -> TagDataBook:
        """Parse a TAG Data Book ``.xlsm``.

        The workbook is opened read-only, with formulas resolved to
        their cached values, so Excel is not needed and nothing is
        written.
        """
        logger.info("Reading TAG Data Book: %s", path)
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            reader = _Reader(workbook)
            expected = {SHEET_VALUE_OF_TIME, SHEET_DISCOUNT, SHEET_CARBON}
            missing = expected - reader.sheet_names()
            if missing:
                logger.warning(
                    "Workbook is missing expected sheet(s) %s; it is probably not a "
                    "TAG Data Book, or is a release with a different layout",
                    ", ".join(sorted(missing)),
                )
            databook = cls(
                vot=_read_value_of_time(reader),
                discount=_read_discount_rates(reader),
                carbon=_read_carbon_values(reader),
                gdp=_read_gdp_deflator(reader),
                source=str(path),
                fallbacks=tuple(reader.fallbacks),
            )
        finally:
            workbook.close()

        if databook.fallbacks:
            logger.warning(
                "%d parameter(s) fell back to built-in defaults; the workbook layout may "
                "differ from v2.03. First: %s",
                len(databook.fallbacks),
                databook.fallbacks[0],
            )
        logger.info(
            "Loaded TAG Data Book (%d prices): rail work VoT £%.2f/hr, commuting £%.2f/hr, "
            "other £%.2f/hr",
            databook.price_year,
            databook.vot.work_rail,
            databook.vot.commuting,
            databook.vot.other,
        )
        return databook

    @classmethod
    def with_defaults(cls) -> TagDataBook:
        """Build from the built-in v2.03 figures, with no workbook.

        Suitable for screening and for tests.
        :mod:`tag_databook.defaults` says how each figure was obtained;
        the carbon series, for one, is grown from a single year's value.
        """
        logger.info("Using built-in TAG v2.03 defaults (no workbook)")
        return cls(
            vot=defaults.VALUE_OF_TIME,
            discount=defaults.discount_rates(),
            carbon=defaults.carbon_values(),
            gdp=defaults.gdp_deflator(),
            source="built-in defaults (TAG v2.03)",
        )

    @property
    def price_year(self) -> int:
        """The price year of every monetary value here."""
        return self.vot.price_year

    @property
    def vot_work_rail(self) -> float:
        """Employer's business value of time by rail, £/hour."""
        return self.vot.work_rail

    @property
    def vot_commuting(self) -> float:
        """Commuting value of time, £/hour."""
        return self.vot.commuting

    @property
    def vot_other(self) -> float:
        """Other non-work value of time, £/hour."""
        return self.vot.other

    def discount_rate_for_year(self, years_ahead: int) -> float:
        """The standard Green Book rate *years_ahead* of the base."""
        for low, high, rate in self.discount.standard:
            if low <= years_ahead <= high:
                return rate
        return self.discount.standard[-1][2]

    def discount_factor(self, years_ahead: int) -> float:
        """Discount factor for a value *years_ahead* of the base year.

        Compounds year by year instead of applying one rate over the
        whole period, because the Green Book rate steps down partway
        through a long appraisal.
        """
        factor = 1.0
        for year in range(1, years_ahead + 1):
            factor /= 1.0 + self.discount_rate_for_year(year)
        return factor

    def carbon_price(self, year: int, scenario: CarbonScenario = "central") -> float:
        """Carbon value for *year*, £/tCO2e at factor cost.

        Years after the published series are extrapolated at 1.5% a
        year, and years before it take the earliest published value.
        Both are logged, since TAG prescribes neither.
        """
        index = {"low": 0, "central": 1, "high": 2}[scenario]
        series = self.carbon.factor_cost
        if not series:
            raise ValueError("no carbon values loaded")
        if year in series:
            return series[year][index]

        latest = max(series)
        if year > latest:
            logger.debug("Carbon value for %d extrapolated beyond published %d", year, latest)
            return series[latest][index] * (1.015 ** (year - latest))

        earliest = min(series)
        logger.debug(
            "Carbon value for %d predates published %d; using %d", year, earliest, earliest
        )
        return series[earliest][index]

    def vot_for_year(self, base_vot: float, base_year: int, target_year: int) -> float:
        """Grow *base_vot* in real terms from one year to another.

        TAG uplifts values of time by 1.5% a year (Unit A1.1). Going
        back in time is not defined, so an earlier target year returns
        the value unchanged.
        """
        years = target_year - base_year
        if years <= 0:
            return base_vot
        return base_vot * (1.0 + self.vot_growth_rate) ** years

    def car_vot_by_distance(self, distance_km: float) -> float:
        """Employer's business car value of time by trip length.

        TAG gives this as a logistic curve in distance rather than one
        rate, because longer business trips are made by people whose
        time is valued more highly::

            VoT = U / (1 + exp((xmid - distance) / k))
        """
        exponent = (self.vot.car_xmid - distance_km) / self.vot.car_k
        try:
            denominator = 1.0 + math.exp(exponent)
        except OverflowError:
            # A very short trip makes the exponent large and positive,
            # and the curve tends to zero there.
            return 0.0
        return self.vot.car_u_factor_cost / denominator

    def deflate(self, amount: float, from_year: int, to_year: int) -> float:
        """Convert *amount* from *from_year* prices to *to_year* prices.

        Raises ``KeyError`` if either year is absent from the deflator
        series, which is the common case when running on built-in
        defaults.
        """
        series = self.gdp.series
        if not series:
            raise KeyError("GDP deflator series is empty (running on built-in defaults?)")
        missing = [year for year in (from_year, to_year) if year not in series]
        if missing:
            wanted = ", ".join(str(year) for year in missing)
            raise KeyError(
                f"GDP deflator has no entry for {wanted}; "
                f"series covers {min(series)}-{max(series)}"
            )
        return amount * series[to_year] / series[from_year]

    def __repr__(self) -> str:
        return (
            f"<TagDataBook {self.price_year} prices from {self.source!r}"
            f"{f', {len(self.fallbacks)} fallbacks' if self.fallbacks else ''}>"
        )


# Per-table readers.


def _read_value_of_time(reader: _Reader) -> ValueOfTime:
    """Read Table A 1.3.1, values of time per person."""
    year = reader.optional_number(SHEET_VALUE_OF_TIME, PRICE_YEAR_CELL)
    fields = {
        name: reader.number(
            SHEET_VALUE_OF_TIME,
            coordinate,
            parameter=f"vot.{name}",
            fallback=getattr(defaults.VALUE_OF_TIME, name),
        )
        for name, coordinate in VALUE_OF_TIME_CELLS.items()
    }
    return ValueOfTime(
        price_year=int(year) if year is not None else defaults.PRICE_YEAR,
        **fields,
    )


def _read_discount_rates(reader: _Reader) -> DiscountRates:
    """Read Table A 1.1.1, the Green Book stepped discount rates."""
    fallback = defaults.discount_rates()
    series: dict[str, tuple[tuple[int, int, float], ...]] = {}
    for name, column in DISCOUNT_COLUMNS.items():
        default_bands = getattr(fallback, name)
        bands = tuple(
            (
                low,
                high,
                reader.number(
                    SHEET_DISCOUNT,
                    f"{column}{DISCOUNT_FIRST_ROW + offset}",
                    parameter=f"discount.{name}[{low}-{high}]",
                    fallback=default_rate,
                ),
            )
            for offset, ((low, high), (_, _, default_rate)) in enumerate(
                zip(defaults.DISCOUNT_BANDS, default_bands, strict=True)
            )
        )
        series[name] = bands
    return DiscountRates(**series)


def _read_carbon_values(reader: _Reader) -> CarbonValues:
    """Read Table A 3.4.1, carbon values by year."""
    year = reader.optional_number(SHEET_CARBON, PRICE_YEAR_CELL)
    factor_cost = dict(_scan_carbon_rows(reader))
    if not factor_cost:
        logger.warning(
            "No carbon values found on sheet %r; falling back to the synthesised series",
            SHEET_CARBON,
        )
        return defaults.carbon_values()
    return CarbonValues(
        factor_cost=factor_cost,
        price_year=int(year) if year is not None else defaults.PRICE_YEAR,
    )


def _scan_carbon_rows(reader: _Reader) -> Iterator[tuple[int, tuple[float, float, float]]]:
    for row in CARBON_ROWS:
        year = reader.optional_number(SHEET_CARBON, f"{CARBON_YEAR_COLUMN}{row}")
        if year is None:
            continue
        low, central, high = (
            reader.optional_number(SHEET_CARBON, f"{column}{row}") or 0.0
            for column in CARBON_VALUE_COLUMNS
        )
        yield int(year), (low, central, high)


def _read_gdp_deflator(reader: _Reader) -> GDPDeflator:
    series: dict[int, float] = {}
    for row in DEFLATOR_ROWS:
        year = reader.optional_number(SHEET_ANNUAL_PARAMETERS, f"{DEFLATOR_YEAR_COLUMN}{row}")
        index = reader.optional_number(SHEET_ANNUAL_PARAMETERS, f"{DEFLATOR_INDEX_COLUMN}{row}")
        if year is None or index is None:
            continue
        series[int(year)] = index
    if not series:
        logger.warning("No GDP deflator series found on sheet %r", SHEET_ANNUAL_PARAMETERS)
    return GDPDeflator(series=series, base_year=defaults.PRICE_YEAR)
