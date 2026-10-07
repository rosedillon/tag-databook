# tag-databook

Read appraisal parameters out of the DfT TAG Data Book without opening
Excel.

The [TAG Data Book](https://www.gov.uk/government/publications/tag-data-book)
is the Department for Transport's reference workbook for transport
appraisal in Great Britain: values of time, Green Book discount rates,
carbon values and the GDP deflator. Business cases across the country
take their numbers from it, and it comes as a 6 MB macro-enabled
spreadsheet with the values under merged headers and footnotes at fixed
cells.

Projects often end up with a constant copied by hand into a Python
file, out of date as soon as DfT publishes the next release. This
reads the workbook directly.

```python
from tag_databook import TagDataBook

book = TagDataBook.from_workbook("tag-data-book-v2-03.xlsm")

book.vot_work_rail  # 38.215157722171526   £/hour, 2023 prices
book.discount_factor(30)  # 0.3562784106023024   stepped Green Book rates
book.carbon_price(2050)  # 427.33920325320685   £/tCO2e, central scenario
book.deflate(1_000_000, 2019, 2023)  # 1186690.5702068459
```

Values come back at the workbook's full precision. Most of these cells
are formulas, so the values are not the rounded figures printed in the
guidance. Round them when you present them, as the CLI does.

## Install

Python 3.11 or later.

```sh
pip install git+https://github.com/rosedillon/tag-databook
```

or, from a clone, `pip install .`. The only runtime dependency is
openpyxl.

## Getting the data book

The workbook is not included here. DfT publishes it, a new release
comes out a few times a year, and it is 6 MB.

1. Go to <https://www.gov.uk/government/publications/tag-data-book>.
2. Download the current TAG Data Book (`.xlsm`).
3. Pass its path to `TagDataBook.from_workbook` or the CLI.

There is no login and no licence to accept. Use the current release,
not an old copy, because the values change between releases.

## Command line

Every subcommand takes the workbook path. Leave it out to use the
built-in figures (see below).

```console
$ tag-databook show tag-data-book-v2-03.xlsm
Source        tag-data-book-v2-03.xlsm
Price year    2023

Values of time (£/hour, factor cost)
  Rail, employer's business   £38.22
  Car driver, business        £23.17
  Average, business           £25.23
  Commuting                   £13.04
  Other non-work              £5.95
  Real growth                 1.5% a year

Discount rate, year 1         3.5%
Carbon values                 2010-2100
GDP deflator                  1990-2100
```

Carbon values for the years you need:

```console
$ tag-databook carbon tag-data-book-v2-03.xlsm --years 2025 2030 2040 2050
Carbon value, central scenario (£/tCO2e, 2023 prices)
  2025  £293.53
  2030  £316.57
  2040  £368.22
  2050  £427.34
```

Discount rates and the cumulative factors that follow from them:

```console
$ tag-databook discount tag-data-book-v2-03.xlsm --years 10 30 60
Green Book standard rates
  years   0-30   3.5%
  years  31-75   3.0%
  years  76-125  2.5%
  years 126-200  2.0%
  years 201-300  1.5%
  years 301-999  1.0%

Cumulative discount factors
  year  10      0.7089
  year  30      0.3563
  year  60      0.1468
```

Moving money between price years:

```console
$ tag-databook deflate tag-data-book-v2-03.xlsm --amount 1000000 --from-year 2019 --to-year 2023
£1,000,000.00 in 2019 prices = £1,186,690.57 in 2023 prices
```

Add `--json` to any of these for machine-readable output.

## What it reads

| Parameter | Where in the workbook | In Python |
|---|---|---|
| values of time, working and non-working | Table A1.3.1 | `book.vot`, a [`ValueOfTime`](src/tag_databook/tables.py) |
| car value of time by distance (a logistic curve) | Table A1.3.1 | `book.car_vot_by_distance(km)` |
| Green Book stepped discount rates (standard, health, risk to life) | Table A1.1.1 | `book.discount`, `book.discount_factor(n)` |
| carbon values, low, central and high by year | Table A3.4.1 | `book.carbon`, `book.carbon_price(year)` |
| GDP deflator index | Annual Parameters | `book.gdp`, `book.deflate(x, from, to)` |

Values come back in the workbook's own price year (2023 in v2.03),
which is on `book.price_year`. Nothing is converted for you.

## Newer releases

The cell layout was checked against TAG Data Book v2.03 (November 2023).

The workbook has no schema, so the parser reads fixed cells such as
`A1.3.1!D37`. DfT rearranges sheets between releases, so a newer
workbook may keep a value somewhere else. When a cell cannot be read,
the parameter takes its built-in v2.03 value and the fallback is
recorded:

```python
book = TagDataBook.from_workbook("tag-data-book-v2-99.xlsm")

for fallback in book.fallbacks:
    print(fallback)
# vot.work_rail: A1.3.1!D37 unreadable, used 38.22
```

An empty `book.fallbacks` means every value came from your workbook. A
long list means the layout has moved and the cells in
[`databook.py`](src/tag_databook/databook.py) need updating. They are
listed together at the top of that module, so updating them is a data
change, not a code change.

The CLI prints the same warning to stderr, so it shows up in a
pipeline too.

## Running without the workbook

For screening work, or a CI job that should not carry a 6 MB file:

```python
book = TagDataBook.with_defaults()
```

This has the same interface, backed by v2.03 figures written into
[`defaults.py`](src/tag_databook/defaults.py). Values of time are
rounded to the penny. Two other differences:

- The carbon series is grown from the 2023 central value at 1.5% a
  year, which keeps it within 0.4% of the published series from 2020
  to 2100. Use the workbook if carbon matters to your result.
- There is no deflator history, so `deflate()` raises an error rather
  than guess.

## Limitations

- Only the tables above are read. The data book has many more; adding
  one means adding its cells to the layout table.
- Values are at factor cost. Conversion to market prices is not
  applied.
- `carbon_price()` extends the published series past its last year at
  1.5% a year. That is a convenience rather than DfT guidance, and it is
  logged at debug level.
- `vot_for_year()` only grows values forward. TAG does not define
  growing them backwards, so an earlier target year returns the value
  unchanged.
- It reads the data book and nothing more. There is no BCR, NPV or
  scheme modelling.

## Python API

| Name | |
|---|---|
| `TagDataBook.from_workbook(path)` | read a workbook |
| `TagDataBook.with_defaults()` | the built-in v2.03 figures |
| `.vot`, `.discount`, `.carbon`, `.gdp` | the four parsed tables |
| `.fallbacks` | cells that could not be read |
| `.discount_factor(years)` | cumulative factor over stepped rates |
| `.carbon_price(year, scenario="central")` | £/tCO2e for one year |
| `.vot_for_year(vot, base_year, target_year)` | grow a value of time in real terms |
| `.car_vot_by_distance(km)` | employer's business car value of time |
| `.deflate(amount, from_year, to_year)` | change price year |

## Development

```sh
git clone https://github.com/rosedillon/tag-databook
cd tag-databook
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
pytest
ruff check . && ruff format --check . && mypy src
```

The tests build synthetic workbooks with values at exactly the cells the
parser reads (see [`tests/conftest.py`](tests/conftest.py)), so moving a
cell in the parser without moving it in the fixture fails the suite.
The tests need no network and no copy of the real workbook.

## Licence

GPL-3.0. See [LICENSE](LICENSE).

Not affiliated with or endorsed by the Department for Transport.
