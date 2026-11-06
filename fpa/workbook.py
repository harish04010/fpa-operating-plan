"""
Builds the FY2026 operating plan and variance pack.

Sheets, in reading order: Read Me, Dashboard, Commentary, BvA, Bridge,
Forecast, Scenarios, Budget, Actuals, Assumptions.

Budget, BvA, Bridge, Forecast, Scenarios, the dashboard figures and even the
commentary sentences are formulas. Change a driver on Assumptions, or a number
on Actuals, and the whole pack moves.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

from .plan import ACTUAL_MONTHS, DIVISIONS, DRIVERS, MONTHS, SCENARIOS, actuals

FONT = "Arial"
BLUE, GREEN, BLACK, GREY = "0000FF", "008000", "000000", "5B6677"
HEAD = PatternFill("solid", fgColor="1F3A5F")
SUB = PatternFill("solid", fgColor="E8EDF3")
YELLOW = PatternFill("solid", fgColor="FFFF00")
FAV = PatternFill("solid", fgColor="D9EAD3")
UNFAV = PatternFill("solid", fgColor="F4CCCC")
M = '#,##0.0;(#,##0.0);"-"'          # $ millions, one decimal
M2 = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
X = '0.000"x"'

C0 = 3                                  # first month column (C)
MC = [L(C0 + i) for i in range(12)]     # month columns C..N
FY = L(C0 + 12)                         # O
YTD_END = MC[ACTUAL_MONTHS - 1]         # J (Sep)
FC = MC[ACTUAL_MONTHS:]                 # K..N forecast months


def font(color=BLACK, bold=False, size=10, italic=False):
    return Font(name=FONT, color=color, bold=bold, size=size, italic=italic)


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = font(bold=True, size=14)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = font(GREY, italic=True)


def month_header(ws, row, extra=("FY2026",), label="$ millions"):
    ws.cell(row=row, column=1, value=label)
    for i, m in enumerate(MONTHS + list(extra)):
        c = ws.cell(row=row, column=C0 + i, value=m if i < 12 else extra[i - 12])
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = Alignment(horizontal="center")
    for c in (1, 2):
        ws.cell(row=row, column=c).fill = HEAD
        ws.cell(row=row, column=c).font = Font(name=FONT, bold=True, color="FFFFFF")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 9
    for c in MC + [FY]:
        ws.column_dimensions[c].width = 11


# P&L layout shared by Budget, Actuals and Forecast so rows line up everywhere.
ROWS = {
    "traffic": 5, "conversion": 6, "transactions": 7, "upt": 8, "units": 9,
    "u0": 11, "u1": 12, "u2": 13, "u3": 14,
    "r0": 16, "r1": 17, "r2": 18, "r3": 19, "revenue": 20,
    "c0": 22, "c1": 23, "c2": 24, "c3": 25, "product_cost": 26, "freight": 27, "cost_of_sales": 28,
    "gross_profit": 29, "gm_pct": 30,
    "payroll": 32, "marketing": 33, "occupancy": 34, "other_sga": 35, "sga": 36,
    "ebitda": 37, "ebitda_pct": 38, "da": 39, "ebit": 40, "interest": 41, "pretax": 42, "tax": 43, "net_income": 44,
}
LABELS = {
    "traffic": "Store traffic (M visits)", "conversion": "Conversion", "transactions": "Transactions (M)",
    "upt": "Units per transaction", "units": "Units (M)",
    **{f"u{i}": f"  Units: {d}" for i, d in enumerate(DIVISIONS)},
    **{f"r{i}": f"  {d}" for i, d in enumerate(DIVISIONS)},
    "revenue": "Net revenue", **{f"c{i}": f"  Product cost: {d}" for i, d in enumerate(DIVISIONS)},
    "product_cost": "Product cost", "freight": "Freight", "cost_of_sales": "Cost of sales",
    "gross_profit": "Gross profit", "gm_pct": "Gross margin %", "payroll": "Store payroll", "marketing": "Marketing",
    "occupancy": "Occupancy", "other_sga": "Other SG&A", "sga": "Total SG&A", "ebitda": "EBITDA",
    "ebitda_pct": "EBITDA margin %", "da": "Depreciation & amortization", "ebit": "Operating income (EBIT)",
    "interest": "Interest expense", "pretax": "Pretax income", "tax": "Income tax", "net_income": "Net income",
}
SECTIONS = {4: "Drivers", 10: "Units by division (M)", 15: "Revenue by division", 21: "Cost of sales", 31: "SG&A"}
BOLD = {"revenue", "gross_profit", "ebitda", "net_income", "sga", "cost_of_sales"}
RATE = {"conversion", "gm_pct", "ebitda_pct"}


def labels(ws):
    for k, r in ROWS.items():
        ws.cell(row=r, column=1, value=LABELS[k]).font = font(bold=k in BOLD)
    for r, s in SECTIONS.items():
        c = ws.cell(row=r, column=1, value=s)
        c.font = font(bold=True)
        for col in range(1, C0 + 13):
            ws.cell(row=r, column=col).fill = SUB


def totals_and_profit(ws, cols, fy=True):
    """Sum and profit lines that are the same formula on every P&L sheet."""
    R = ROWS
    for c in cols:
        ws[f"{c}{R['units']}"] = f"=SUM({c}{R['u0']}:{c}{R['u3']})" if ws.title != "Budget" else ws[f"{c}{R['units']}"].value
        ws[f"{c}{R['revenue']}"] = f"=SUM({c}{R['r0']}:{c}{R['r3']})"
        ws[f"{c}{R['product_cost']}"] = f"=SUM({c}{R['c0']}:{c}{R['c3']})"
        ws[f"{c}{R['cost_of_sales']}"] = f"={c}{R['product_cost']}+{c}{R['freight']}"
        ws[f"{c}{R['gross_profit']}"] = f"={c}{R['revenue']}-{c}{R['cost_of_sales']}"
        ws[f"{c}{R['gm_pct']}"] = f"=IF({c}{R['revenue']}=0,0,{c}{R['gross_profit']}/{c}{R['revenue']})"
        ws[f"{c}{R['sga']}"] = f"=SUM({c}{R['payroll']}:{c}{R['other_sga']})"
        ws[f"{c}{R['ebitda']}"] = f"={c}{R['gross_profit']}-{c}{R['sga']}"
        ws[f"{c}{R['ebitda_pct']}"] = f"=IF({c}{R['revenue']}=0,0,{c}{R['ebitda']}/{c}{R['revenue']})"
        ws[f"{c}{R['ebit']}"] = f"={c}{R['ebitda']}-{c}{R['da']}"
        ws[f"{c}{R['pretax']}"] = f"={c}{R['ebit']}-{c}{R['interest']}"
        ws[f"{c}{R['tax']}"] = f"=MAX(0,{c}{R['pretax']})*Assumptions!$C$27"
        ws[f"{c}{R['net_income']}"] = f"={c}{R['pretax']}-{c}{R['tax']}"
    if fy:
        for k, r in R.items():
            if k in RATE or k in ("upt",):
                continue
            ws[f"{FY}{r}"] = f"=SUM({MC[0]}{r}:{MC[-1]}{r})"
        ws[f"{FY}{R['gm_pct']}"] = f"=IF({FY}{R['revenue']}=0,0,{FY}{R['gross_profit']}/{FY}{R['revenue']})"
        ws[f"{FY}{R['ebitda_pct']}"] = f"=IF({FY}{R['revenue']}=0,0,{FY}{R['ebitda']}/{FY}{R['revenue']})"
        ws[f"{FY}{R['conversion']}"] = f"=IF({FY}{R['traffic']}=0,0,{FY}{R['transactions']}/{FY}{R['traffic']})"
        ws[f"{FY}{R['upt']}"] = f"=IF({FY}{R['transactions']}=0,0,{FY}{R['units']}/{FY}{R['transactions']})"


def fmt_pl(ws, cols):
    for k, r in ROWS.items():
        for c in cols:
            cell = ws[f"{c}{r}"]
            cell.number_format = PCT if k in RATE else (M2 if k in ("upt",) else M)
            if cell.font.color is None or cell.font.color.rgb in (None, "FF000000"):
                cell.font = font(bold=k in BOLD)


def build(path: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    names = ["Read Me", "Dashboard", "Commentary", "BvA", "Bridge", "Forecast", "Scenarios", "Budget", "Actuals", "Assumptions",
             "Chart Data"]
    S = {n: wb.create_sheet(n) for n in names}
    d, a = DRIVERS, actuals()

    # ================================================================== Assumptions
    ws = S["Assumptions"]
    title(ws, "Assumptions and drivers", "Blue = input. Change any blue cell and the budget, forecast, scenarios and dashboard update.")
    month_header(ws, 4, extra=(), label="Monthly drivers")
    ws["A5"] = "Store traffic (M visits)"
    for i, c in enumerate(MC):
        ws[f"{c}5"] = d["traffic"][i]
        ws[f"{c}5"].font = font(BLUE)
        ws[f"{c}5"].number_format = M
    scal = [
        (7, "Conversion (transactions / visits)", d["conversion"], PCT),
        (8, "Units per transaction", d["upt"], M2),
        (9, "Freight cost per unit ($)", d["freight_per_unit"], M2),
        (10, "Store labor: fixed hours per month (M)", d["fixed_hours"], M2),
        (11, "Store labor: variable hours per M visits (M)", d["hours_per_k_visits"], '0.000'),
        (12, "Average hourly wage incl. burden ($)", d["wage"], M2),
        (13, "Marketing, % of revenue", d["marketing_pct"], PCT),
        (14, "Occupancy per month ($M)", d["occupancy"], M),
        (15, "Other SG&A per month ($M)", d["other_sga"], M),
        (16, "D&A per month ($M)", d["da"], M),
        (17, "Interest per month ($M)", d["interest"], M),
    ]
    for r, lab, v, fm in scal:
        ws[f"A{r}"], ws[f"C{r}"] = lab, v
        ws[f"C{r}"].font = font(BLUE)
        ws[f"C{r}"].number_format = fm
    ws["A19"] = "Division"
    for j, h in enumerate(["Unit mix", "AUR ($)", "Merch margin"]):
        ws.cell(row=19, column=3 + j, value=h).font = font(bold=True)
    for i, dv in enumerate(DIVISIONS):
        r = 20 + i
        ws[f"A{r}"] = dv
        for j, (v, fm) in enumerate([(d["mix"][i], PCT), (d["aur"][i], M2), (d["margin"][i], PCT)]):
            cell = ws.cell(row=r, column=3 + j, value=v)
            cell.font = font(BLUE)
            cell.number_format = fm
    ws["A24"], ws["C24"] = "Mix check (should be 100%)", "=SUM(C20:C23)"
    ws["C24"].number_format = PCT
    ws["A27"], ws["C27"] = "Tax rate", d["tax_rate"]
    ws["C27"].font = font(BLUE)
    ws["C27"].number_format = PCT
    ws["C7"].fill = ws["C20"].fill = YELLOW
    ws["C7"].comment = Comment("Key driver. Conversion and AUR move revenue most; see Scenarios for sensitivity.", "FP&A")

    # Forecast run-rates, measured from actuals: analyst judgment written as formulas.
    ws["A30"] = "Forecast run-rates (from actuals)"
    ws["A30"].font = font(bold=True)
    rr = [
        (31, "Traffic vs plan, Jun-Sep (spring weather treated as one-off)",
         f"=SUM(Actuals!G5:J5)/SUM(Budget!G5:J5)", X),
        (32, "Freight per unit vs plan, September", f"=(Actuals!J27/Actuals!J9)/$C$9", X),
        (33, "Wage vs plan, September", f"=(Actuals!J32/($C$10+$C$11*Actuals!J5))/$C$12", X),
    ]
    for r, lab, v, fm in rr:
        ws[f"A{r}"], ws[f"C{r}"] = lab, v
        ws[f"C{r}"].number_format = fm
        ws[f"C{r}"].font = font(GREEN)
    ws["A35"] = "Division"
    ws["C35"], ws["D35"] = "YTD AUR vs plan", "YTD unit mix"
    for i, dv in enumerate(DIVISIONS):
        r = 36 + i
        ws[f"A{r}"] = dv
        ws[f"C{r}"] = (f"=(SUM(Actuals!C{ROWS['r0'] + i}:{YTD_END}{ROWS['r0'] + i})/SUM(Actuals!C{ROWS['u0'] + i}:{YTD_END}{ROWS['u0'] + i}))"
                       f"/$D${20 + i}")
        ws[f"D{r}"] = f"=SUM(Actuals!C{ROWS['u0'] + i}:{YTD_END}{ROWS['u0'] + i})/SUM(Actuals!C{ROWS['units']}:{YTD_END}{ROWS['units']})"
        ws[f"C{r}"].number_format, ws[f"D{r}"].number_format = X, PCT
        ws[f"C{r}"].font = ws[f"D{r}"].font = font(GREEN)

    ws["A42"] = "Scenarios (multipliers on the forecast months)"
    ws["A42"].font = font(bold=True)
    for j, h in enumerate(["Traffic", "AUR", "Freight", "Wage"]):
        ws.cell(row=43, column=3 + j, value=h).font = font(bold=True)
    for k, (name, mults) in enumerate(SCENARIOS.items()):
        r = 44 + k
        ws[f"A{r}"] = name
        for j, v in enumerate(mults):
            cell = ws.cell(row=r, column=3 + j, value=v)
            cell.font = font(BLUE)
            cell.number_format = X
    ws["A48"], ws["C48"] = "Active scenario", "Base"
    ws["C48"].font = font(BLUE, bold=True)
    ws["C48"].fill = YELLOW
    dv_ = DataValidation(type="list", formula1='"Base,Upside,Downside"', allow_blank=False)
    ws.add_data_validation(dv_)
    dv_.add("C48")
    ws["D48"] = "Pick Base, Upside or Downside. The Forecast sheet and Dashboard follow this cell."
    ws["D48"].font = font(GREY, italic=True)
    for c in "CDEF":
        ws[f"{c}49"] = f"=INDEX({c}44:{c}46,MATCH($C$48,$A$44:$A$46,0))"
        ws[f"{c}49"].number_format = X
    ws["A49"] = "Active multipliers"
    ws.column_dimensions["A"].width = 52
    for c in "CDEF":
        ws.column_dimensions[c].width = 14

    # ================================================================== Budget (driver formulas)
    ws = S["Budget"]
    title(ws, "FY2026 budget (driver-based)", "Every cell is a formula over Assumptions.")
    month_header(ws, 3)
    labels(ws)
    R = ROWS
    for i, c in enumerate(MC):
        ws[f"{c}{R['traffic']}"] = f"=Assumptions!{c}5"
        ws[f"{c}{R['conversion']}"] = "=Assumptions!$C$7"
        ws[f"{c}{R['transactions']}"] = f"={c}{R['traffic']}*{c}{R['conversion']}"
        ws[f"{c}{R['upt']}"] = "=Assumptions!$C$8"
        ws[f"{c}{R['units']}"] = f"={c}{R['transactions']}*{c}{R['upt']}"
        for k in range(4):
            ws[f"{c}{R['u0'] + k}"] = f"={c}{R['units']}*Assumptions!$C${20 + k}"
            ws[f"{c}{R['r0'] + k}"] = f"={c}{R['u0'] + k}*Assumptions!$D${20 + k}"
            ws[f"{c}{R['c0'] + k}"] = f"={c}{R['r0'] + k}*(1-Assumptions!$E${20 + k})"
        ws[f"{c}{R['freight']}"] = f"={c}{R['units']}*Assumptions!$C$9"
        ws[f"{c}{R['payroll']}"] = f"=(Assumptions!$C$10+Assumptions!$C$11*{c}{R['traffic']})*Assumptions!$C$12"
        ws[f"{c}{R['marketing']}"] = f"={c}{R['revenue']}*Assumptions!$C$13"
        ws[f"{c}{R['occupancy']}"] = "=Assumptions!$C$14"
        ws[f"{c}{R['other_sga']}"] = "=Assumptions!$C$15"
        ws[f"{c}{R['da']}"] = "=Assumptions!$C$16"
        ws[f"{c}{R['interest']}"] = "=Assumptions!$C$17"
    totals_and_profit(ws, MC)
    fmt_pl(ws, MC + [FY])
    for c in MC:
        for k in ("traffic", "conversion", "upt", "occupancy", "other_sga", "da", "interest"):
            ws[f"{c}{R[k]}"].font = font(GREEN)
    ws.freeze_panes = "C4"

    # ================================================================== Actuals (data, Feb-Sep)
    ws = S["Actuals"]
    title(ws, "FY2026 actuals, February to September",
          "Blue = actuals as loaded from the ledger (case data). Totals and profit lines are formulas. Oct-Jan are blank until they close.")
    month_header(ws, 3, extra=("YTD",))
    labels(ws)
    act = MC[:ACTUAL_MONTHS]
    for m, c in enumerate(act):
        vals = {"traffic": a["traffic"][m], "freight": a["freight"][m], "payroll": a["payroll"][m],
                "marketing": a["marketing"][m], "occupancy": a["occupancy"][m], "other_sga": a["other_sga"][m],
                "da": a["da"][m], "interest": a["interest"][m]}
        for k in range(4):
            vals[f"u{k}"], vals[f"r{k}"], vals[f"c{k}"] = a["units"][m][k], a["revenue"][m][k], a["cogs"][m][k]
        for k, v in vals.items():
            ws[f"{c}{R[k]}"] = v
            ws[f"{c}{R[k]}"].font = font(BLUE)
        ws[f"{c}{R['upt']}"] = "=Assumptions!$C$8"
        ws[f"{c}{R['transactions']}"] = f"={c}{R['units']}/{c}{R['upt']}"
        ws[f"{c}{R['conversion']}"] = f"={c}{R['transactions']}/{c}{R['traffic']}"
        ws[f"{c}{R['upt']}"].font = font(GREEN)
    ws["A46"] = ("Transactions are implied from units at plan UPT, so conversion here is a derived figure, "
                 "not a counted one: worth saying when presenting it.")
    ws["A46"].font = font(GREY, italic=True)
    totals_and_profit(ws, act, fy=False)
    for k, r in R.items():
        if k in RATE or k == "upt":
            continue
        ws[f"{FY}{r}"] = f"=SUM({MC[0]}{r}:{YTD_END}{r})"
    ws[f"{FY}{R['gm_pct']}"] = f"=IF({FY}{R['revenue']}=0,0,{FY}{R['gross_profit']}/{FY}{R['revenue']})"
    ws[f"{FY}{R['ebitda_pct']}"] = f"=IF({FY}{R['revenue']}=0,0,{FY}{R['ebitda']}/{FY}{R['revenue']})"
    ws[f"{FY}{R['conversion']}"] = f"=IF({FY}{R['traffic']}=0,0,{FY}{R['transactions']}/{FY}{R['traffic']})"
    fmt_pl(ws, act + [FY])
    ws.freeze_panes = "C4"

    # ================================================================== Forecast (8 + 4)
    ws = S["Forecast"]
    title(ws, "FY2026 forecast: 8 months actual + 4 months forecast",
          "Feb-Sep link to Actuals. Oct-Jan run the budget drivers at measured run-rates, times the active scenario on Assumptions.")
    month_header(ws, 3)
    labels(ws)
    for c in act:
        for k, r in R.items():
            ws[f"{c}{r}"] = f"=Actuals!{c}{r}"
            ws[f"{c}{r}"].font = font(GREEN)
    # Pre-scenario forecast components (rows 50+), reused by Scenarios and the sensitivity grid.
    P = {"rev": 50, "cost": 51, "units": 52, "pay_fixed": 53, "pay_var": 54, "traffic": 55}
    ws["A49"] = "Forecast components before scenario multipliers (used by Scenarios)"
    ws["A49"].font = font(bold=True)
    for lab, r in (("Revenue", 50), ("Product cost", 51), ("Units", 52), ("Payroll: fixed hours", 53),
                   ("Payroll: traffic-driven", 54), ("Traffic", 55)):
        ws[f"A{r}"] = lab
    for c in FC:
        t = f"Budget!{c}{R['traffic']}*Assumptions!$C$31"
        ws[f"{c}{P['traffic']}"] = f"={t}"
        ws[f"{c}{P['units']}"] = f"={c}{P['traffic']}*Assumptions!$C$7*Assumptions!$C$8"
        ws[f"{c}{P['rev']}"] = "=" + "+".join(
            f"{c}{P['units']}*Assumptions!$D${36 + k}*Assumptions!$D${20 + k}*Assumptions!$C${36 + k}" for k in range(4))
        ws[f"{c}{P['cost']}"] = "=" + "+".join(
            f"{c}{P['units']}*Assumptions!$D${36 + k}*Assumptions!$D${20 + k}*Assumptions!$C${36 + k}*(1-Assumptions!$E${20 + k})"
            for k in range(4))
        ws[f"{c}{P['pay_fixed']}"] = "=Assumptions!$C$10*Assumptions!$C$12*Assumptions!$C$33"
        ws[f"{c}{P['pay_var']}"] = f"=Assumptions!$C$11*{c}{P['traffic']}*Assumptions!$C$12*Assumptions!$C$33"
        for r in P.values():
            ws[f"{c}{r}"].number_format = M
        # The visible forecast: components times the active scenario.
        tm, am, fm, wm = "Assumptions!$C$49", "Assumptions!$D$49", "Assumptions!$E$49", "Assumptions!$F$49"
        ws[f"{c}{R['traffic']}"] = f"={c}{P['traffic']}*{tm}"
        ws[f"{c}{R['conversion']}"] = "=Assumptions!$C$7"
        ws[f"{c}{R['transactions']}"] = f"={c}{R['traffic']}*{c}{R['conversion']}"
        ws[f"{c}{R['upt']}"] = "=Assumptions!$C$8"
        for k in range(4):
            ws[f"{c}{R['u0'] + k}"] = f"={c}{R['transactions']}*{c}{R['upt']}*Assumptions!$D${36 + k}"
            ws[f"{c}{R['r0'] + k}"] = f"={c}{R['u0'] + k}*Assumptions!$D${20 + k}*Assumptions!$C${36 + k}*{am}"
            ws[f"{c}{R['c0'] + k}"] = f"={c}{R['r0'] + k}*(1-Assumptions!$E${20 + k})"
        ws[f"{c}{R['freight']}"] = f"={c}{R['units']}*Assumptions!$C$9*Assumptions!$C$32*{fm}"
        ws[f"{c}{R['payroll']}"] = f"=({c}{P['pay_fixed']}+{c}{P['pay_var']}*{tm})*{wm}"
        ws[f"{c}{R['marketing']}"] = f"={c}{R['revenue']}*Assumptions!$C$13"
        for k, ref in (("occupancy", 14), ("other_sga", 15), ("da", 16), ("interest", 17)):
            ws[f"{c}{R[k]}"] = f"=Assumptions!$C${ref}"
    totals_and_profit(ws, FC)
    fmt_pl(ws, MC + [FY])
    ws["A57"], ws[f"{FY}57"] = "Check: forecast revenue rows rebuild the components at Base", \
        f"=IF(Assumptions!$C$48<>\"Base\",TRUE,ABS(SUM({FC[0]}{R['revenue']}:{FC[-1]}{R['revenue']})-SUM({FC[0]}50:{FC[-1]}50))<0.001)"
    for c in MC:
        ws[f"{c}3"].fill = PatternFill("solid", fgColor="1F3A5F" if c in act else "3D6A99")
    ws.freeze_panes = "C4"

    # ================================================================== BvA (YTD)
    ws = S["BvA"]
    title(ws, "Budget vs actual, year to date (February to September)",
          "Favorable is green: revenue and profit above plan, costs below it.")
    heads = ["Line", "", "YTD budget", "YTD actual", "Variance $", "Variance %", "Fav / Unfav", "Sign"]
    for j, h in enumerate(heads, start=1):
        cell = ws.cell(row=4, column=j, value=h)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = HEAD
    lines = [("r0", 1), ("r1", 1), ("r2", 1), ("r3", 1), ("revenue", 1), ("product_cost", -1), ("freight", -1),
             ("gross_profit", 1), ("gm_pct", 1), ("payroll", -1), ("marketing", -1), ("occupancy", -1),
             ("other_sga", -1), ("sga", -1), ("ebitda", 1), ("ebitda_pct", 1), ("net_income", 1),
             ("traffic", 1), ("units", 1)]
    BVA = {}
    for i, (k, sign) in enumerate(lines, start=5):
        BVA[k] = i
        r = R[k]
        ws[f"A{i}"] = LABELS[k].strip()
        ws[f"A{i}"].font = font(bold=k in BOLD)
        if k in RATE:
            num = {"gm_pct": "gross_profit", "ebitda_pct": "ebitda"}[k]
            ws[f"C{i}"] = f"=IF(C{BVA['revenue']}=0,0,C{BVA[num]}/C{BVA['revenue']})"
            ws[f"D{i}"] = f"=IF(D{BVA['revenue']}=0,0,D{BVA[num]}/D{BVA['revenue']})"
            ws[f"E{i}"] = f"=D{i}-C{i}"
            ws[f"F{i}"] = ""
            for c in "CDE":
                ws[f"{c}{i}"].number_format = PCT2
        else:
            ws[f"C{i}"] = f"=SUM(Budget!{MC[0]}{r}:Budget!{YTD_END}{r})"
            ws[f"D{i}"] = f"=SUM(Actuals!{MC[0]}{r}:Actuals!{YTD_END}{r})"
            ws[f"E{i}"] = f"=D{i}-C{i}"
            ws[f"F{i}"] = f"=IF(C{i}=0,0,E{i}/ABS(C{i}))"
            for c in "CDE":
                ws[f"{c}{i}"].number_format = M
            ws[f"F{i}"].number_format = PCT
        ws[f"H{i}"] = sign
        ws[f"G{i}"] = f"=IF(ABS(E{i})<0.0005,\"-\",IF(E{i}*H{i}>0,\"Fav\",\"Unfav\"))"
        for c in "CD":
            ws[f"{c}{i}"].font = font(GREEN, bold=k in BOLD)
    last = 4 + len(lines)
    ws.conditional_formatting.add(f"E5:G{last}", FormulaRule(formula=[f"$G5=\"Fav\""], fill=FAV))
    ws.conditional_formatting.add(f"E5:G{last}", FormulaRule(formula=[f"$G5=\"Unfav\""], fill=UNFAV))
    ws["A26"] = "Sign: +1 means a higher number is good (revenue, profit); -1 means lower is good (costs)."
    ws["A26"].font = font(GREY, italic=True)
    ws.column_dimensions["A"].width = 30
    for c, w in zip("BCDEFGH", [2, 14, 14, 13, 12, 11, 6]):
        ws.column_dimensions[c].width = w
    ws.column_dimensions["H"].hidden = False

    # ================================================================== Bridge
    ws = S["Bridge"]
    title(ws, "What moved revenue and EBITDA, year to date", "Each bridge adds up exactly; the check cells prove it.")
    hd = ["Division", "Budget units", "Actual units", "Budget AUR", "Actual AUR", "Budget mix", "Actual mix",
          "Mix effect", "Price effect"]
    for j, h in enumerate(hd, start=1):
        cell = ws.cell(row=4, column=j, value=h)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = HEAD
    for i, dv in enumerate(DIVISIONS):
        r = 5 + i
        ws[f"A{r}"] = dv
        ws[f"B{r}"] = f"=SUM(Budget!{MC[0]}{R['u0'] + i}:Budget!{YTD_END}{R['u0'] + i})"
        ws[f"C{r}"] = f"=SUM(Actuals!{MC[0]}{R['u0'] + i}:Actuals!{YTD_END}{R['u0'] + i})"
        ws[f"D{r}"] = f"=SUM(Budget!{MC[0]}{R['r0'] + i}:Budget!{YTD_END}{R['r0'] + i})/B{r}"
        ws[f"E{r}"] = f"=SUM(Actuals!{MC[0]}{R['r0'] + i}:Actuals!{YTD_END}{R['r0'] + i})/C{r}"
        ws[f"F{r}"] = f"=B{r}/$B$9"
        ws[f"G{r}"] = f"=C{r}/$C$9"
        ws[f"H{r}"] = f"=(C{r}-$C$9*F{r})*D{r}"
        ws[f"I{r}"] = f"=C{r}*(E{r}-D{r})"
        for c, fm in zip("BCDEFGHI", [M, M, M2, M2, PCT, PCT, M, M]):
            ws[f"{c}{r}"].number_format = fm
    ws["A9"], ws["B9"], ws["C9"], ws["H9"], ws["I9"] = "Total", "=SUM(B5:B8)", "=SUM(C5:C8)", "=SUM(H5:H8)", "=SUM(I5:I8)"
    for c in "BCHI":
        ws[f"{c}9"].number_format = M
        ws[f"{c}9"].font = font(bold=True)
    ws["D9"] = "=SUMPRODUCT(B5:B8,D5:D8)/B9"
    ws["D9"].number_format = M2
    # Revenue bridge rows (also the chart data)
    ws["A12"] = "Revenue bridge, budget to actual ($M)"
    ws["A12"].font = font(bold=True)
    for j, h in enumerate(["Step", "Value", "Base", "Up", "Down"], start=1):
        ws.cell(row=13, column=j, value=h).font = font(bold=True)
    steps = [
        ("Budget revenue", f"=BvA!C{BVA['revenue']}"),
        ("Traffic", f"=(SUM(Actuals!C5:{YTD_END}5)-SUM(Budget!C5:{YTD_END}5))*(B9/SUM(Budget!C5:{YTD_END}5))*D9"),
        ("Conversion x UPT", f"=(C9-B9)*D9-B15"),
        ("Mix", "=H9"),
        ("Price (AUR)", "=I9"),
        ("Actual revenue", f"=BvA!D{BVA['revenue']}"),
    ]
    for i, (lab, fm) in enumerate(steps, start=14):
        ws[f"A{i}"], ws[f"B{i}"] = lab, fm
        ws[f"B{i}"].number_format = M
    ws["A21"], ws["B21"] = "Check: budget + effects = actual", "=ABS(B14+B15+B16+B17+B18-B19)<0.001"
    # EBITDA bridge
    ws["A24"] = "EBITDA bridge, budget to actual ($M)"
    ws["A24"].font = font(bold=True)
    for j, h in enumerate(["Step", "Value", "Base", "Up", "Down"], start=1):
        ws.cell(row=25, column=j, value=h).font = font(bold=True)
    rb, ra = BVA["revenue"], BVA["revenue"]
    eb = [
        ("Budget EBITDA", f"=BvA!C{BVA['ebitda']}"),
        ("Sales volume, mix and price", f"=(BvA!D{rb}-BvA!C{rb})*(1-BvA!C{BVA['product_cost']}/BvA!C{rb})"),
        ("Merchandise margin rate", f"=BvA!D{rb}*((1-BvA!D{BVA['product_cost']}/BvA!D{rb})-(1-BvA!C{BVA['product_cost']}/BvA!C{rb}))"),
        ("Freight", f"=-BvA!E{BVA['freight']}"),
        ("Store payroll", f"=-BvA!E{BVA['payroll']}"),
        ("Marketing", f"=-BvA!E{BVA['marketing']}"),
        ("Occupancy and other", f"=-(BvA!E{BVA['occupancy']}+BvA!E{BVA['other_sga']})"),
        ("Actual EBITDA", f"=BvA!D{BVA['ebitda']}"),
    ]
    for i, (lab, fm) in enumerate(eb, start=26):
        ws[f"A{i}"], ws[f"B{i}"] = lab, fm
        ws[f"B{i}"].number_format = M
    ws["A35"], ws["B35"] = "Check: budget + effects = actual", "=ABS(SUM(B26:B32)-B33)<0.001"
    # Waterfall helper columns: hidden base, up and down bars.
    for first, lastr in ((14, 19), (26, 33)):
        for i in range(first, lastr + 1):
            if i in (first, lastr):
                ws[f"C{i}"], ws[f"D{i}"], ws[f"E{i}"] = 0, f"=B{i}", 0
            else:
                ws[f"C{i}"] = f"=SUM($B${first}:B{i - 1})-MAX(0,-B{i})"
                ws[f"D{i}"] = f"=MAX(0,B{i})"
                ws[f"E{i}"] = f"=MAX(0,-B{i})"
            for c in "CDE":
                ws[f"{c}{i}"].number_format = M
    for i in range(14, 20):
        ws[f"C{i}"].font = font(GREY)
    # Axis floors: a waterfall from zero hides $50M steps on a $3.3B base. The floor sits a little below the
    # lowest point the bridge reaches, and the chart title says where it starts. Set at build time.
    floors = {"Revenue": 3100, "EBITDA": 320}
    for anchor, first, lastr, ttl, lo in (("G12", 14, 19, "Revenue: budget to actual, YTD ($M)", "Revenue"),
                                           ("G26", 26, 33, "EBITDA: budget to actual, YTD ($M)", "EBITDA")):
        ttl = ttl + f" (axis starts at {floors[lo]:,})"
        ch = BarChart()
        ch.type, ch.grouping, ch.overlap = "col", "stacked", 100
        ch.title = ttl
        ch.add_data(Reference(ws, min_col=3, max_col=5, min_row=first - 1, max_row=lastr), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=first, max_row=lastr))
        ch.series[0].graphicalProperties.noFill = True
        ch.series[0].graphicalProperties.line.noFill = True
        ch.series[1].graphicalProperties.solidFill = "2F6B4F"
        ch.series[2].graphicalProperties.solidFill = "B3263E"
        ch.legend = None
        ch.height, ch.width = 7.5, 17
        ch.y_axis.numFmt = "#,##0"
        ch.y_axis.scaling.min = floors[lo]
        ws.add_chart(ch, anchor)
    ws.column_dimensions["A"].width = 30
    for c in "BCDEFGHI":
        ws.column_dimensions[c].width = 13
    ws.conditional_formatting.add("B21", CellIsRule(operator="equal", formula=["TRUE"], fill=FAV))
    ws.conditional_formatting.add("B35", CellIsRule(operator="equal", formula=["TRUE"], fill=FAV))

    # ================================================================== Scenarios + sensitivity
    ws = S["Scenarios"]
    title(ws, "Full-year outlook by scenario", "Same forecast components as the Forecast sheet, multiplied by each scenario's levers.")
    for j, h in enumerate(["", "Budget", "Base", "Upside", "Downside"], start=1):
        cell = ws.cell(row=4, column=j, value=h)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = HEAD
    fc = f"Forecast!{FC[0]}{{r}}:Forecast!{FC[-1]}{{r}}"
    ytd = lambda k: f"SUM(Actuals!{MC[0]}{R[k]}:Actuals!{YTD_END}{R[k]})"
    for j, (name, col) in enumerate((("Base", "C"), ("Upside", "D"), ("Downside", "E"))):
        sr = 44 + j
        t, au, fr, wg = (f"Assumptions!$C${sr}", f"Assumptions!$D${sr}", f"Assumptions!$E${sr}", f"Assumptions!$F${sr}")
        rev = f"SUM({fc.format(r=50)})*{t}*{au}"
        cost = f"SUM({fc.format(r=51)})*{t}*{au}"
        freight = f"SUM({fc.format(r=52)})*{t}*Assumptions!$C$9*Assumptions!$C$32*{fr}"
        payroll = f"(SUM({fc.format(r=53)})+SUM({fc.format(r=54)})*{t})*{wg}"
        fixed = f"4*(Assumptions!$C$14+Assumptions!$C$15)"
        ws[f"{col}5"] = f"={ytd('revenue')}+{rev}"
        ws[f"{col}6"] = f"={ytd('gross_profit')}+({rev})-({cost})-({freight})"
        ws[f"{col}7"] = (f"={ytd('ebitda')}+({rev})-({cost})-({freight})-({payroll})"
                         f"-({rev})*Assumptions!$C$13-{fixed}")
        ws[f"{col}8"] = f"=IF({col}5=0,0,{col}7/{col}5)"
        ws[f"{col}9"] = f"={col}7-$B$7"
    ws["A5"], ws["A6"], ws["A7"], ws["A8"], ws["A9"] = "Revenue", "Gross profit", "EBITDA", "EBITDA margin", "EBITDA vs budget"
    ws["B5"], ws["B6"], ws["B7"] = f"=Budget!{FY}{R['revenue']}", f"=Budget!{FY}{R['gross_profit']}", f"=Budget!{FY}{R['ebitda']}"
    ws["B8"] = "=IF(B5=0,0,B7/B5)"
    for c in "BCDE":
        for r in (5, 6, 7, 9):
            ws[f"{c}{r}"].number_format = M
        ws[f"{c}8"].number_format = PCT
    ws["A10"], ws["C10"] = "Check: Base matches the Forecast sheet at Base", \
        f"=IF(Assumptions!$C$48<>\"Base\",TRUE,ABS(C7-Forecast!{FY}{R['ebitda']})<0.001)"
    ws["A13"] = "Sensitivity: full-year EBITDA ($M) by forecast-month traffic and AUR, freight and wages at Base"
    ws["A13"].font = font(bold=True)
    ws["A14"] = "Traffic \u2193  /  AUR \u2192"
    aurs = [0.97, 0.98, 0.99, 1.00, 1.01, 1.02, 1.03]
    trs = [0.94, 0.96, 0.98, 1.00, 1.02, 1.04, 1.06]
    for j, av in enumerate(aurs):
        cell = ws.cell(row=14, column=2 + j, value=av)
        cell.font = font(BLUE, bold=True)
        cell.number_format = X
    for i, tv in enumerate(trs):
        r = 15 + i
        cell = ws.cell(row=r, column=1, value=tv)
        cell.font = font(BLUE, bold=True)
        cell.number_format = X
        for j in range(len(aurs)):
            c = L(2 + j)
            t, au = f"$A{r}", f"{c}$14"
            rev = f"SUM({fc.format(r=50)})*{t}*{au}"
            ws[f"{c}{r}"] = (f"={ytd('ebitda')}+{rev}-SUM({fc.format(r=51)})*{t}*{au}"
                             f"-SUM({fc.format(r=52)})*{t}*Assumptions!$C$9*Assumptions!$C$32"
                             f"-(SUM({fc.format(r=53)})+SUM({fc.format(r=54)})*{t})-{rev}*Assumptions!$C$13"
                             f"-4*(Assumptions!$C$14+Assumptions!$C$15)")
            ws[f"{c}{r}"].number_format = M
    ws.conditional_formatting.add("B15:H21", FormulaRule(formula=["B15<$B$7"], fill=UNFAV))
    ws.conditional_formatting.add("B15:H21", FormulaRule(formula=["B15>=$B$7"], fill=FAV))
    ws["A23"] = "Green cells meet or beat the full-year EBITDA budget; red cells miss it."
    ws["A23"].font = font(GREY, italic=True)
    ws.column_dimensions["A"].width = 44
    for c in "BCDEFGH":
        ws.column_dimensions[c].width = 12

    # ================================================================== Commentary (live sentences)
    ws = S["Commentary"]
    title(ws, "Variance commentary, September year to date",
          "Each sentence is a formula, so the commentary rewrites itself when actuals or drivers change.")
    m = lambda ref: f"TEXT(ABS({ref}),\"$#,##0.0\")&\"M\""
    p = lambda ref: f"TEXT(ABS({ref}),\"0.0%\")"
    ud = lambda ref, up="above", down="below": f"IF({ref}>=0,\"{up}\",\"{down}\")"
    b = BVA
    sentences = [
        f"=\"Revenue was \"&{m(f'BvA!D{b['revenue']}')}&\", \"&{m(f'BvA!E{b['revenue']}')}&\" (\"&{p(f'BvA!F{b['revenue']}')}&\") \"&{ud(f'BvA!E{b['revenue']}')}&\" plan.\"",
        f"=\"Traffic cost \"&{m('Bridge!B15')}&\" and conversion/basket \"&{ud('Bridge!B16', 'added ', 'cost ')}&{m('Bridge!B16')}&\", while price added \"&{m('Bridge!B18')}&\" and mix toward higher-priced Footwear added \"&{m('Bridge!B17')}&\".\"",
        f"=\"Traffic ran \"&{p(f'BvA!F{b['traffic']}')}&\" \"&{ud(f'BvA!E{b['traffic']}')}&\" plan, almost all of it in a cold, wet February-May; June-September ran \"&TEXT(Assumptions!C31-1,\"+0.0%;-0.0%\")&\" against plan.\"",
        f"=\"EBITDA was \"&{m(f'BvA!D{b['ebitda']}')}&\", \"&{m(f'BvA!E{b['ebitda']}')}&\" \"&{ud(f'BvA!E{b['ebitda']}')}&\" plan, at a \"&TEXT(BvA!D{b['ebitda_pct']},\"0.0%\")&\" margin against \"&TEXT(BvA!C{b['ebitda_pct']},\"0.0%\")&\" planned.\"",
        f"=\"Freight cost \"&{m(f'BvA!E{b['freight']}')}&\" more than plan as ocean rates spiked in May-July; September freight per unit was \"&TEXT(Assumptions!C32,\"0.00x\")&\" plan.\"",
        f"=\"Store payroll was \"&{m(f'BvA!E{b['payroll']}')}&\" \"&{ud(f'BvA!E{b['payroll']}', 'over', 'under')}&\" plan, mostly overtime in the July-August back-to-school ramp; marketing was \"&{m(f'BvA!E{b['marketing']}')}&\" \"&{ud(f'BvA!E{b['marketing']}', 'over', 'under')}&\", from the August campaign.\"",
        f"=\"The \"&Assumptions!C48&\" forecast lands full-year EBITDA at \"&{m(f'Forecast!{FY}{R['ebitda']}')}&\", \"&{m(f'Forecast!{FY}{R['ebitda']}-Budget!{FY}{R['ebitda']}')}&\" \"&{ud(f'Forecast!{FY}{R['ebitda']}-Budget!{FY}{R['ebitda']}')}&\" budget. The range across scenarios is \"&{m('Scenarios!E7')}&\" to \"&{m('Scenarios!D7')}&\".\"",
    ]
    for i, s in enumerate(sentences, start=4):
        ws[f"A{i}"] = s
        ws[f"A{i}"].font = font(size=11)
        ws[f"A{i}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = 34
    ws.column_dimensions["A"].width = 130

    # ================================================================== Dashboard
    ws = S["Dashboard"]
    title(ws, "FY2026 performance: September year to date", "Ridgeline Outfitters (case study). $ millions.")
    tiles = [
        ("B4", "YTD revenue", f"=BvA!D{b['revenue']}", f"=BvA!F{b['revenue']}", "vs plan"),
        ("D4", "YTD EBITDA", f"=BvA!D{b['ebitda']}", f"=BvA!E{b['ebitda']}/BvA!C{b['ebitda']}", "vs plan"),
        ("F4", "YTD gross margin", f"=BvA!D{b['gm_pct']}", f"=BvA!E{b['gm_pct']}", "pts vs plan"),
        ("H4", "FY EBITDA forecast", f"=Forecast!{FY}{R['ebitda']}", f"=Forecast!{FY}{R['ebitda']}/Budget!{FY}{R['ebitda']}-1", "vs budget"),
    ]
    for anchor, lab, val, delta, note in tiles:
        col, row = anchor[0], int(anchor[1:])
        ws[f"{col}{row}"] = lab
        ws[f"{col}{row}"].font = font(GREY, bold=True)
        ws[f"{col}{row + 1}"] = val
        ws[f"{col}{row + 1}"].font = font(bold=True, size=18)
        ws[f"{col}{row + 1}"].number_format = PCT if "margin" in lab else '"$"#,##0.0"M"'
        ws[f"{col}{row + 2}"] = delta
        ws[f"{col}{row + 2}"].number_format = '+0.0%;-0.0%' if "margin" not in lab else '+0.00%;-0.00%'
        ws[f"{chr(ord(col) + 1)}{row + 2}"] = note
        ws[f"{chr(ord(col) + 1)}{row + 2}"].font = font(GREY)
        ws.conditional_formatting.add(f"{col}{row + 2}", CellIsRule(operator="greaterThanOrEqual", formula=["0"], font=Font(name=FONT, color="2F6B4F", bold=True)))
        ws.conditional_formatting.add(f"{col}{row + 2}", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, color="B3263E", bold=True)))
    for c, w in zip("ABCDEFGHI", [2, 18, 10, 18, 10, 18, 10, 20, 12]):
        ws.column_dimensions[c].width = w
    # Chart series live on a visible sheet: Excel doesn't plot hidden cells by default.
    cd = S["Chart Data"]
    title(cd, "Chart data", "Feeds the Dashboard charts. Formulas only; nothing to edit here.")
    cd["A4"], cd["B4"], cd["C4"], cd["D4"] = "Month", "Budget", "Actual", "Forecast"
    for i, mth in enumerate(MONTHS):
        r = 5 + i
        c = MC[i]
        cd[f"A{r}"] = mth
        cd[f"B{r}"] = f"=Budget!{c}{R['revenue']}"
        cd[f"C{r}"] = f"=Actuals!{c}{R['revenue']}" if i < ACTUAL_MONTHS else None
        cd[f"D{r}"] = f"=Forecast!{c}{R['revenue']}" if i >= ACTUAL_MONTHS - 1 else None
        for col in "BCD":
            cd[f"{col}{r}"].number_format = M
    lc = LineChart()
    lc.title = "Net revenue by month ($M)"
    lc.add_data(Reference(cd, min_col=2, max_col=4, min_row=4, max_row=16), titles_from_data=True)
    lc.set_categories(Reference(cd, min_col=1, min_row=5, max_row=16))
    lc.series[0].graphicalProperties.line.solidFill = "9AA5AE"
    lc.series[0].graphicalProperties.line.dashStyle = "dash"
    lc.series[1].graphicalProperties.line.solidFill = "1F3A5F"
    lc.series[2].graphicalProperties.line.solidFill = "B7791F"
    for s in lc.series:
        s.smooth = False  # straight segments: a smoothed curve invents values between months
    lc.height, lc.width = 8, 18
    ws.add_chart(lc, "B9")
    bc = BarChart()
    bc.type = "bar"
    bc.title = "YTD cost variances vs plan ($M, positive = over plan)"
    cd["F4"], cd["G4"] = "Line", "Over (under) plan"
    for i, (lab, k) in enumerate((("Freight", "freight"), ("Store payroll", "payroll"), ("Marketing", "marketing"),
                                  ("Occupancy", "occupancy"), ("Other SG&A", "other_sga"))):
        cd[f"F{5 + i}"], cd[f"G{5 + i}"] = lab, f"=BvA!E{b[k]}"
        cd[f"G{5 + i}"].number_format = M
    bc.add_data(Reference(cd, min_col=7, min_row=4, max_row=9), titles_from_data=True)
    bc.set_categories(Reference(cd, min_col=6, min_row=5, max_row=9))
    bc.series[0].graphicalProperties.solidFill = "B3263E"
    bc.legend = None
    bc.height, bc.width = 8, 14
    ws.add_chart(bc, "B26")
    cd.column_dimensions["A"].width = 10
    cd.column_dimensions["F"].width = 14
    ws["H26"] = "Revenue and EBITDA bridges are on the Bridge sheet; scenario range on Scenarios."
    ws["H26"].font = font(GREY, italic=True)

    # ================================================================== Read Me
    ws = S["Read Me"]
    title(ws, "FY2026 operating plan and variance pack")
    rows = [
        ("Company", "Ridgeline Outfitters, a fictional sporting goods retailer used as a case study. All figures are generated."),
        ("Fiscal year", "February 2026 to January 2027 (a retail year). Actuals through September; October to January forecast."),
        ("Units", "$ millions unless the row says otherwise. Traffic and units in millions."),
        ("Reading order", "Dashboard > Commentary > BvA > Bridge > Forecast > Scenarios. Budget, Actuals and Assumptions underneath."),
        ("Colors", "Blue text = input. Black = formula. Green = formula pulling from another sheet. Yellow fill = key assumption."),
        ("Change a driver", "Edit any blue cell on Assumptions; Budget, Forecast, Scenarios, Dashboard and Commentary all update."),
        ("Switch scenario", "Assumptions!C48: Base, Upside or Downside. Forecast, Dashboard and Commentary follow."),
        ("Checks", "Bridge!B21 and B35 (bridges add up), Forecast!O57 and Scenarios!C10 (scenario logic consistent). All should read TRUE."),
        ("Power BI", "powerbi/ holds the same data as a star schema plus DAX measures, for the same views in Power BI."),
    ]
    for r, (k, v) in enumerate(rows, start=3):
        ws[f"A{r}"], ws[f"B{r}"] = k, v
        ws[f"A{r}"].font = font(bold=True)
        ws[f"B{r}"].font = font()
        ws[f"B{r}"].alignment = Alignment(wrap_text=True)
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 110
    for w in wb.worksheets:
        w.sheet_view.showGridLines = False
    wb.save(path)
