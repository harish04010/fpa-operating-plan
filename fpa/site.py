"""Exports the recalculated plan to site/data.json for the web view, including the forecast
components, so the page can re-run the scenario math live."""
import json
from pathlib import Path

from openpyxl import load_workbook

from .plan import ACTUAL_MONTHS, DIVISIONS, MONTHS, SCENARIOS
from .workbook import FC, FY, MC, ROWS as R

ROOT = Path(__file__).resolve().parent.parent


def export() -> str:
    wb = load_workbook(ROOT / "output" / "FY2026_Operating_Plan.xlsx", data_only=True)
    B, A, F, V, BR, SC, AS, CM = (wb[n] for n in ("Budget", "Actuals", "Forecast", "BvA", "Bridge", "Scenarios", "Assumptions", "Commentary"))
    if B["O20"].value is None:
        raise RuntimeError("Recalculate the workbook first.")
    r1 = lambda v: round(float(v), 1)
    comp = {k: round(sum(F[f"{c}{row}"].value for c in FC), 3) for k, row in
            (("rev", 50), ("cost", 51), ("units", 52), ("pay_fixed", 53), ("pay_var", 54))}
    ytd = {k: r1(A[f"O{R[k]}"].value) for k in ("revenue", "gross_profit", "ebitda")}
    data = {
        "months": MONTHS, "actual_months": ACTUAL_MONTHS,
        "revenue": {"budget": [r1(B[f"{c}{R['revenue']}"].value) for c in MC],
                    "actual": [r1(A[f"{c}{R['revenue']}"].value) for c in MC[:ACTUAL_MONTHS]],
                    "forecast": [r1(F[f"{c}{R['revenue']}"].value) for c in MC]},
        "kpi": {"ytd_rev": r1(V["D9"].value), "ytd_rev_var": round(V["F9"].value, 4),
                "ytd_ebitda": r1(V["D19"].value), "ytd_ebitda_bud": r1(V["C19"].value),
                "gm": round(V["D13"].value, 4), "gm_bud": round(V["C13"].value, 4),
                "fy_budget_ebitda": r1(B[f"{FY}{R['ebitda']}"].value), "fy_budget_rev": r1(B[f"{FY}{R['revenue']}"].value)},
        "bva": [{"line": V[f"A{r}"].value, "bud": round(V[f"C{r}"].value, 4), "act": round(V[f"D{r}"].value, 4),
                 "var": round(V[f"E{r}"].value, 4), "fav": V[f"G{r}"].value, "pct": "%" in (V[f"A{r}"].value or "")}
                for r in range(5, 22)],
        "rev_bridge": [{"step": BR[f"A{r}"].value, "v": r1(BR[f"B{r}"].value)} for r in range(14, 20)],
        "ebitda_bridge": [{"step": BR[f"A{r}"].value, "v": r1(BR[f"B{r}"].value)} for r in range(26, 34)],
        "divisions": [{"name": d, "aur_b": round(BR[f"D{5+i}"].value, 2), "aur_a": round(BR[f"E{5+i}"].value, 2),
                       "mix_b": round(BR[f"F{5+i}"].value, 4), "mix_a": round(BR[f"G{5+i}"].value, 4)} for i, d in enumerate(DIVISIONS)],
        "scenario_model": {"ytd": ytd, "components": comp, "fpu": AS["C9"].value, "freight_rr": round(AS["C32"].value, 5),
                           "marketing_pct": AS["C13"].value, "fixed_4m": r1(4 * (AS["C14"].value + AS["C15"].value)),
                           "scenarios": SCENARIOS},
        "scenario_check": {k: r1(SC[f"{c}7"].value) for k, c in (("Base", "C"), ("Upside", "D"), ("Downside", "E"))},
        "commentary": [CM[f"A{r}"].value for r in range(4, 9)],
    }
    out = ROOT / "site" / "data.json"
    out.write_text(json.dumps(data, separators=(",", ":")))
    return f"{out.name}: {out.stat().st_size:,} bytes"


if __name__ == "__main__":
    print(export())
