import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import verify_dashboard as vd  # noqa: E402


@pytest.fixture(scope="module")
def workbook():
    return vd.load_workbook()


def failing(checks):
    return {c.name for c in checks if not c.passed and c.blocking}


def test_workbook_passes_every_blocking_check():
    assert vd.main([]) == 0


def test_sheets_loaded_without_empty_table_rows(workbook):
    original, data, _ = workbook
    assert len(original) == len(data) == 10329


def test_headline_kpis(workbook):
    _, data, _ = workbook
    k = vd.recalculate_kpis(data)
    assert k["total_sales"] == pytest.approx(561085.51, abs=0.01)
    assert k["total_units"] == 20699
    assert k["units_excl_cancelled"] == 19566
    assert k["platform_share"]["Shopify"] == pytest.approx(0.627, abs=0.001)
    assert k["returning_rate"] == pytest.approx(0.169, abs=0.001)
    assert k["cancel_rate"] == pytest.approx(0.054, abs=0.001)
    assert k["avg_rating"] == pytest.approx(4.12, abs=0.005)


def test_monthly_pivot_keeps_both_years(workbook):
    monthly = vd.pivot_monthly(workbook[2])
    assert (2019, "Jan") in monthly and (2020, "Jan") in monthly
    assert monthly[(2019, "Jan")] != monthly[(2020, "Jan")]


def test_detects_a_changed_sale(workbook):
    original, data, grid = workbook
    data = data.copy()
    data.loc[0, "Sales(Won)"] += 1000
    assert "Sales(Won) matches original 확정매출" in failing(vd.data_checks(original, data))


def test_detects_a_wrong_usd_conversion(workbook):
    original, data, _ = workbook
    data = data.copy()
    data.loc[5, "Sales($)"] += 1
    assert "Sales($) = Sales(Won) × 0.00078" in failing(vd.data_checks(original, data))


def test_detects_inconsistent_translation(workbook):
    original, data, _ = workbook
    data = data.copy()
    first_amazon = data.index[data["Platform"] == "Amazon"][0]
    data.loc[first_amazon, "Platform"] = "Shopify"
    assert "Platform: each Korean value → one English value" in failing(vd.data_checks(original, data))


def test_detects_kpi_drift(workbook):
    _, data, grid = workbook
    data = data.copy()
    data = data[data["Platform"] != "Amazon"]   # dropping rows changes every total
    assert "total sales ($)" in failing(vd.kpi_checks(vd.recalculate_kpis(data), grid))


def test_spelling_variants_are_a_warning_not_a_failure(workbook):
    original, data, _ = workbook
    spelling = [c for c in vd.data_checks(original, data) if c.name == "Sub-Category: consistent spelling/case"][0]
    assert not spelling.passed and not spelling.blocking
