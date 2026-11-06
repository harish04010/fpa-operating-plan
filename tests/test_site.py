"""The web dashboard's outlook arithmetic must reproduce the workbook's Scenarios sheet."""
import json
from pathlib import Path

import pytest

DATA = Path(__file__).resolve().parent.parent / "site" / "data.json"


def outlook(o, t, a, f, w):
    """Same formula as site/index.html and the Scenarios sheet."""
    c = o["components"]
    rev = sum(c["rev"]) * t * a
    cost = sum(c["cost"]) * t * a
    freight = sum(c["units"]) * t * o["freight_per_unit"] * o["freight_runrate"] * f
    payroll = (sum(c["pay_fixed"]) + sum(c["pay_var"]) * t) * w
    return o["ytd_revenue"] + rev, o["ytd_ebitda"] + rev - cost - freight - payroll - rev * o["marketing_pct"] - 4 * o["fixed_monthly"]


@pytest.mark.parametrize("scenario", ["Base", "Upside", "Downside"])
def test_page_outlook_matches_excel(scenario):
    d = json.loads(DATA.read_text())["outlook"]
    rev, ebitda = outlook(d, *d["scenarios"][scenario])
    assert rev == pytest.approx(d["excel_results"][scenario]["revenue"], abs=1e-6)
    assert ebitda == pytest.approx(d["excel_results"][scenario]["ebitda"], abs=1e-6)
