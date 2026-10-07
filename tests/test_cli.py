from __future__ import annotations

import json
from typing import TYPE_CHECKING

import pytest

from tag_databook import databook as db
from tag_databook.cli import build_parser, main

from .conftest import PRICE_YEAR, VOT_VALUES, build_workbook, carbon_row

if TYPE_CHECKING:
    import pathlib

    from _pytest.capture import CaptureFixture


SMOKE_ARGS = ["show"]


def test_show_reports_the_workbook_values(
    workbook_path: pathlib.Path, capsys: CaptureFixture[str]
) -> None:
    assert main(["show", str(workbook_path)]) == 0
    out = capsys.readouterr().out
    assert str(PRICE_YEAR) in out
    assert f"{VOT_VALUES['work_rail']:,.2f}" in out


def test_show_as_json_is_machine_readable(
    workbook_path: pathlib.Path, capsys: CaptureFixture[str]
) -> None:
    assert main(["show", str(workbook_path), "--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["price_year"] == PRICE_YEAR
    assert payload["values_of_time"]["work_rail"] == pytest.approx(VOT_VALUES["work_rail"])
    assert payload["fallbacks"] == []


def test_no_workbook_warns(capsys: CaptureFixture[str]) -> None:
    assert main(["show"]) == 0
    captured = capsys.readouterr()
    assert "built-in" in captured.err
    assert "2023" in captured.out


def test_vot_growth(workbook_path: pathlib.Path, capsys: CaptureFixture[str]) -> None:
    assert main(["vot", str(workbook_path), "--year", str(PRICE_YEAR + 10), "--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["grown_to_year"] == PRICE_YEAR + 10
    assert payload["work_rail"] == pytest.approx(VOT_VALUES["work_rail"] * 1.015**10)


def test_carbon_reports_requested_years(
    workbook_path: pathlib.Path, capsys: CaptureFixture[str]
) -> None:
    assert main(["carbon", str(workbook_path), "--years", "2021", "2022", "--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["values"]["2021"] == pytest.approx(carbon_row(2021)[1])
    assert payload["values"]["2022"] == pytest.approx(carbon_row(2022)[1])


def test_carbon_scenario(workbook_path: pathlib.Path, capsys: CaptureFixture[str]) -> None:
    assert (
        main(["carbon", str(workbook_path), "--years", "2021", "--scenario", "high", "--json"])
        == 0
    )
    payload = json.loads(capsys.readouterr().out)
    assert payload["values"]["2021"] == pytest.approx(carbon_row(2021)[2])


def test_discount_reports_bands_and_factors(
    workbook_path: pathlib.Path, capsys: CaptureFixture[str]
) -> None:
    assert main(["discount", str(workbook_path), "--years", "10", "20", "--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert len(payload["bands"]) == 6
    assert 0.0 < payload["factors"]["20"] < payload["factors"]["10"] < 1.0


def test_deflate(workbook_path: pathlib.Path, capsys: CaptureFixture[str]) -> None:
    exit_code = main(
        [
            "deflate",
            str(workbook_path),
            "--amount",
            "100",
            "--from-year",
            "2020",
            "--to-year",
            "2025",
            "--json",
        ]
    )
    assert exit_code == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["converted"] > 100.0


def test_deflate_missing_year(workbook_path: pathlib.Path, capsys: CaptureFixture[str]) -> None:
    exit_code = main(
        [
            "deflate",
            str(workbook_path),
            "--amount",
            "100",
            "--from-year",
            "1970",
            "--to-year",
            "2025",
        ]
    )
    assert exit_code == 1
    assert "1970" in capsys.readouterr().err


def test_degraded_workbook_warns(tmp_path: pathlib.Path, capsys: CaptureFixture[str]) -> None:
    path = build_workbook(tmp_path / "partial.xlsx", omit={db.SHEET_VALUE_OF_TIME})
    assert main(["show", str(path)]) == 0
    assert "fell back" in capsys.readouterr().err


def test_no_subcommand_is_an_error() -> None:
    with pytest.raises(SystemExit) as exit_info:
        main([])
    assert exit_info.value.code != 0


@pytest.mark.parametrize(
    ("argv", "verbose"),
    [(SMOKE_ARGS, False), (["-v", *SMOKE_ARGS], True), ([*SMOKE_ARGS, "-v"], True)],
)
def test_verbose_flag(argv: list[str], verbose: bool) -> None:
    assert build_parser().parse_args(argv).verbose is verbose
