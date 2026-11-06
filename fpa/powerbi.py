"""
Exports the plan as a star schema for Power BI, from the recalculated workbook,
so Power BI and Excel show the same numbers by construction.

    powerbi/fact_pnl.csv       one row per month x line x division x version (Budget, Actual, Forecast)
    powerbi/dim_month.csv      fiscal calendar: fiscal month number, quarter, half, actual/forecast flag
    powerbi/dim_line.csv       P&L lines with their order, group and sign (+1 good when higher, -1 for costs)
    powerbi/dim_division.csv   merchandise divisions
    powerbi/measures.dax       the measures for the report pages
"""
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from .plan import ACTUAL_MONTHS, DIVISIONS, MONTHS
from .workbook import MC, ROWS

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "powerbi"

LINES = [
    # key, label, group, sign, division index or None
    *[(f"r{i}", "Net revenue", "Revenue", 1, i) for i in range(4)],
    *[(f"c{i}", "Product cost", "Cost of sales", -1, i) for i in range(4)],
    ("freight", "Freight", "Cost of sales", -1, None),
    ("payroll", "Store payroll", "SG&A", -1, None),
    ("marketing", "Marketing", "SG&A", -1, None),
    ("occupancy", "Occupancy", "SG&A", -1, None),
    ("other_sga", "Other SG&A", "SG&A", -1, None),
    ("da", "Depreciation & amortization", "Below EBITDA", -1, None),
    ("interest", "Interest expense", "Below EBITDA", -1, None),
    *[(f"u{i}", "Units", "Volume", 1, i) for i in range(4)],
    ("traffic", "Store traffic", "Volume", 1, None),
]
ORDER = ["Net revenue", "Product cost", "Freight", "Store payroll", "Marketing", "Occupancy", "Other SG&A",
         "Depreciation & amortization", "Interest expense", "Units", "Store traffic"]


def export(workbook=ROOT / "output" / "FY2026_Operating_Plan.xlsx") -> None:
    OUT.mkdir(exist_ok=True)
    wb = load_workbook(workbook, data_only=True)
    if wb["Budget"]["O20"].value is None:
        raise RuntimeError("Recalculate the workbook first (Excel or scripts/recalc.py); it has no cached values yet.")
    rows = []
    for version, sheet, months in (("Budget", "Budget", range(12)), ("Actual", "Actuals", range(ACTUAL_MONTHS)),
                                   ("Forecast", "Forecast", range(12))):
        ws = wb[sheet]
        for m in months:
            for key, label, group, sign, div in LINES:
                v = ws[f"{MC[m]}{ROWS[key]}"].value
                rows.append({"fiscal_month": m + 1, "line": label, "division": DIVISIONS[div] if div is not None else "All",
                             "version": version, "amount": round(float(v or 0), 6)})
    pd.DataFrame(rows).to_csv(OUT / "fact_pnl.csv", index=False)
    pd.DataFrame({"fiscal_month": range(1, 13), "month": MONTHS,
                  "calendar_month": [f"{2026 + (i >= 11)}-{(i + 1) % 12 + 1:02d}" for i in range(12)],
                  "quarter": [f"Q{i // 3 + 1}" for i in range(12)], "half": ["H1" if i < 6 else "H2" for i in range(12)],
                  "is_actual": [i < ACTUAL_MONTHS for i in range(12)]}).to_csv(OUT / "dim_month.csv", index=False)
    seen = {}
    for _, label, group, sign, _ in LINES:
        seen.setdefault(label, (group, sign))
    pd.DataFrame([{"line": k, "group": g, "sign": s, "sort": ORDER.index(k) + 1} for k, (g, s) in seen.items()]
                 ).to_csv(OUT / "dim_line.csv", index=False)
    pd.DataFrame({"division": DIVISIONS + ["All"], "sort": range(1, 6)}).to_csv(OUT / "dim_division.csv", index=False)
