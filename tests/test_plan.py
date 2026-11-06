"""Checks against the recalculated workbook in output/, plus model logic that needs no Excel."""
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

from fpa.plan import ACTUAL_MONTHS, DRIVERS, actuals

BOOK = Path(__file__).resolve().parent.parent / "output" / "FY2026_Operating_Plan.xlsx"


@pytest.fixture(scope="module")
def wb():
    wb = load_workbook(BOOK, data_only=True)
    if wb["Budget"]["O20"].value is None:
        pytest.skip("workbook has no calculated values; open and save it in Excel first")
    return wb


def budget_month(m):
    d = DRIVERS
    t = d["traffic"][m]
    units = t * d["conversion"] * d["upt"]
    rev = sum(units * d["mix"][i] * d["aur"][i] for i in range(4))
    cogs = sum(units * d["mix"][i] * d["aur"][i] * (1 - d["margin"][i]) for i in range(4))
    sga = (d["fixed_hours"] + d["hours_per_k_visits"] * t) * d["wage"] + rev * d["marketing_pct"] + d["occupancy"] + d["other_sga"]
    return rev, rev - cogs - units * d["freight_per_unit"] - sga


def test_excel_budget_matches_an_independent_calculation(wb):
    rev = sum(budget_month(m)[0] for m in range(12))
    ebitda = sum(budget_month(m)[1] for m in range(12))
    assert wb["Budget"]["O20"].value == pytest.approx(rev, rel=1e-9)
    assert wb["Budget"]["O37"].value == pytest.approx(ebitda, rel=1e-9)


def test_plan_ratios_are_realistic_for_a_large_specialty_retailer(wb):
    b = wb["Budget"]
    rev = b["O20"].value
    assert 0.30 < b["O30"].value < 0.36          # gross margin
    assert 0.08 < b["O32"].value / rev < 0.12    # store payroll
    assert 0.09 < b["O37"].value / rev < 0.13    # EBITDA margin


def test_every_bridge_and_consistency_check_reads_true(wb):
    assert wb["Bridge"]["B21"].value is True and wb["Bridge"]["B35"].value is True
    assert wb["Forecast"]["O57"].value is True and wb["Scenarios"]["C10"].value is True


def test_price_volume_mix_adds_up_to_the_revenue_variance(wb):
    br = wb["Bridge"]
    effects = sum(br[f"B{r}"].value for r in range(15, 19))
    assert br["B14"].value + effects == pytest.approx(br["B19"].value, abs=1e-6)


def test_scenarios_are_ordered_downside_base_upside(wb):
    s = wb["Scenarios"]
    assert s["E7"].value < s["C7"].value < s["D7"].value


def test_forecast_keeps_actual_months_and_forecasts_the_rest(wb):
    f, a = wb["Forecast"], wb["Actuals"]
    assert f["J20"].value == pytest.approx(a["J20"].value)   # September is actual
    assert a["K20"].value is None and f["K20"].value > 0      # October is forecast


def test_commentary_is_generated_from_the_numbers(wb):
    first = wb["Commentary"]["A4"].value
    assert first.startswith("Revenue was $") and "plan" in first


def test_power_bi_export_matches_the_workbook(wb):
    fact = pd.read_csv(BOOK.parent.parent / "powerbi" / "fact_pnl.csv")
    act = fact[(fact.version == "Actual") & (fact.line == "Net revenue")].amount.sum()
    assert act == pytest.approx(wb["Actuals"]["O20"].value, rel=1e-6)


def test_actuals_generator_is_deterministic():
    assert actuals()["revenue"] == actuals()["revenue"]
    assert len(actuals()["traffic"]) == ACTUAL_MONTHS
