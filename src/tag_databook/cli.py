"""Command-line access to TAG Data Book parameters.

Every subcommand takes an optional workbook path. Omit it and the
built-in v2.03 figures are used instead, with a note on stderr so the
output is never mistaken for a real reading of a real workbook.
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from typing import TYPE_CHECKING, Any

from tag_databook.databook import TagDataBook

if TYPE_CHECKING:
    from collections.abc import Sequence

_MONEY = "£{:,.2f}"


def _load(args: argparse.Namespace) -> TagDataBook:
    if args.workbook is None:
        print(
            "No workbook given; using built-in TAG v2.03 defaults. "
            "Pass a .xlsm path for a formal appraisal.",
            file=sys.stderr,
        )
        return TagDataBook.with_defaults()
    book = TagDataBook.from_workbook(args.workbook)
    if book.fallbacks:
        print(
            f"warning: {len(book.fallbacks)} parameter(s) could not be read and fell back "
            "to built-in defaults. The workbook layout may differ from v2.03:",
            file=sys.stderr,
        )
        for fallback in book.fallbacks:
            print(f"  {fallback}", file=sys.stderr)
    return book


def _emit(payload: dict[str, Any], lines: list[str], *, as_json: bool) -> int:
    if as_json:
        json.dump(payload, sys.stdout, indent=2)
        sys.stdout.write("\n")
    else:
        print("\n".join(lines))
    return 0


def _cmd_show(args: argparse.Namespace) -> int:
    book = _load(args)
    payload: dict[str, Any] = {
        "source": book.source,
        "price_year": book.price_year,
        "values_of_time": {
            "work_rail": book.vot.work_rail,
            "work_car_driver": book.vot.work_car_driver,
            "work_average": book.vot.work_average,
            "commuting": book.vot.commuting,
            "other": book.vot.other,
        },
        "vot_annual_growth": book.vot_growth_rate,
        "discount_rate_year_1": book.discount_rate_for_year(1),
        "carbon_years": (
            [min(book.carbon.factor_cost), max(book.carbon.factor_cost)]
            if book.carbon.factor_cost
            else []
        ),
        "deflator_years": (
            [min(book.gdp.series), max(book.gdp.series)] if book.gdp.series else []
        ),
        "fallbacks": [str(f) for f in book.fallbacks],
    }
    carbon: list[int] = payload["carbon_years"]
    deflator: list[int] = payload["deflator_years"]
    lines = [
        f"Source        {book.source}",
        f"Price year    {book.price_year}",
        "",
        "Values of time (£/hour, factor cost)",
        f"  Rail, employer's business   {_MONEY.format(book.vot.work_rail)}",
        f"  Car driver, business        {_MONEY.format(book.vot.work_car_driver)}",
        f"  Average, business           {_MONEY.format(book.vot.work_average)}",
        f"  Commuting                   {_MONEY.format(book.vot.commuting)}",
        f"  Other non-work              {_MONEY.format(book.vot.other)}",
        f"  Real growth                 {book.vot_growth_rate:.1%} a year",
        "",
        f"Discount rate, year 1         {book.discount_rate_for_year(1):.1%}",
        f"Carbon values                 {carbon[0]}-{carbon[1]}"
        if carbon
        else "Carbon values                 none",
        f"GDP deflator                  {deflator[0]}-{deflator[1]}"
        if deflator
        else "GDP deflator                  none",
    ]
    return _emit(payload, lines, as_json=args.json)


def _cmd_vot(args: argparse.Namespace) -> int:
    book = _load(args)
    vot = book.vot
    payload: dict[str, Any] = {
        "price_year": vot.price_year,
        "work_car_driver": vot.work_car_driver,
        "work_car_passenger": vot.work_car_passenger,
        "work_rail": vot.work_rail,
        "work_bus": vot.work_bus,
        "work_walk": vot.work_walk,
        "work_cycle": vot.work_cycle,
        "work_average": vot.work_average,
        "commuting": vot.commuting,
        "other": vot.other,
    }
    if args.year is not None:
        payload = {
            key: book.vot_for_year(value, vot.price_year, args.year)
            if isinstance(value, float)
            else value
            for key, value in payload.items()
        }
        payload["grown_to_year"] = args.year
    width = max(len(key) for key in payload)
    lines = [
        f"{key.replace('_', ' '):<{width}}  "
        + (_MONEY.format(value) if isinstance(value, float) else str(value))
        for key, value in payload.items()
    ]
    return _emit(payload, lines, as_json=args.json)


def _cmd_carbon(args: argparse.Namespace) -> int:
    book = _load(args)
    years = args.years or [book.price_year]
    payload: dict[str, Any] = {
        "price_year": book.carbon.price_year,
        "scenario": args.scenario,
        "values": {str(year): book.carbon_price(year, args.scenario) for year in years},
    }
    lines = [f"Carbon value, {args.scenario} scenario (£/tCO2e, {book.carbon.price_year} prices)"]
    lines += [
        f"  {year}  {_MONEY.format(book.carbon_price(year, args.scenario))}" for year in years
    ]
    return _emit(payload, lines, as_json=args.json)


def _cmd_discount(args: argparse.Namespace) -> int:
    book = _load(args)
    payload: dict[str, Any] = {
        "bands": [
            {"from_year": low, "to_year": high, "rate": rate}
            for low, high, rate in book.discount.standard
        ],
        "factors": {str(year): book.discount_factor(year) for year in args.years},
    }
    lines = ["Green Book standard rates"]
    lines += [
        f"  years {low:>3}-{high:<3}  {rate:.1%}" for low, high, rate in book.discount.standard
    ]
    if args.years:
        lines += ["", "Cumulative discount factors"]
        lines += [f"  year {year:>3}      {book.discount_factor(year):.4f}" for year in args.years]
    return _emit(payload, lines, as_json=args.json)


def _cmd_deflate(args: argparse.Namespace) -> int:
    book = _load(args)
    try:
        converted = book.deflate(args.amount, args.from_year, args.to_year)
    except KeyError as error:
        print(f"error: {error.args[0]}", file=sys.stderr)
        return 1
    payload: dict[str, Any] = {
        "amount": args.amount,
        "from_year": args.from_year,
        "to_year": args.to_year,
        "converted": converted,
    }
    lines = [
        f"{_MONEY.format(args.amount)} in {args.from_year} prices "
        f"= {_MONEY.format(converted)} in {args.to_year} prices"
    ]
    return _emit(payload, lines, as_json=args.json)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="tag-databook",
        description=(
            "Read appraisal parameters from a DfT TAG Data Book workbook. "
            "Omit the workbook to use built-in TAG v2.03 figures."
        ),
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="log what is being read")
    subparsers = parser.add_subparsers(dest="command", required=True)

    def add(name: str, help_text: str) -> argparse.ArgumentParser:
        sub = subparsers.add_parser(name, help=help_text, description=help_text)
        sub.add_argument("workbook", nargs="?", help="path to the TAG Data Book .xlsm")
        sub.add_argument("--json", action="store_true", help="emit JSON instead of a table")
        return sub

    add("show", "Summarise the whole data book.").set_defaults(func=_cmd_show)

    vot = add("vot", "Values of time, in £ per hour.")
    vot.add_argument(
        "--year",
        type=int,
        help="grow the values to this year at the TAG real-terms rate",
    )
    vot.set_defaults(func=_cmd_vot)

    carbon = add("carbon", "Carbon values, in £ per tCO2e.")
    carbon.add_argument(
        "--years", type=int, nargs="+", help="years to report (default: price year)"
    )
    carbon.add_argument(
        "--scenario",
        choices=("low", "central", "high"),
        default="central",
        help="which published sensitivity to use (default: central)",
    )
    carbon.set_defaults(func=_cmd_carbon)

    discount = add("discount", "Green Book discount rates and cumulative factors.")
    discount.add_argument(
        "--years",
        type=int,
        nargs="*",
        default=[],
        help="report the cumulative discount factor at these year offsets",
    )
    discount.set_defaults(func=_cmd_discount)

    deflate = add("deflate", "Convert an amount between price years.")
    # An option rather than a positional: `workbook` is already an
    # optional positional, so a bare `deflate 100` would bind 100 to the
    # path.
    deflate.add_argument("--amount", type=float, required=True, help="amount of money to convert")
    deflate.add_argument("--from-year", type=int, required=True, dest="from_year")
    deflate.add_argument("--to-year", type=int, required=True, dest="to_year")
    deflate.set_defaults(func=_cmd_deflate)

    # -v also works after the subcommand.
    for subcommand in subparsers.choices.values():
        subcommand.add_argument(
            "-v",
            "--verbose",
            action="store_true",
            default=argparse.SUPPRESS,
            help="log progress",
        )

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING,
        format="%(levelname)s %(message)s",
    )
    return int(args.func(args))
