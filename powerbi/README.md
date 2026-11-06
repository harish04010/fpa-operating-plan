# Power BI version

The same plan as a star schema, exported from the recalculated workbook so both tools show identical numbers.

## Build it (about 20 minutes)

1. **Get data > Text/CSV:** load `fact_pnl.csv`, `dim_month.csv`, `dim_line.csv` and `dim_division.csv`.
2. **Model view:** relate `fact_pnl[fiscal_month]` to `dim_month[fiscal_month]`, `fact_pnl[line]` to `dim_line[line]`, and `fact_pnl[division]` to `dim_division[division]`. All are many-to-one, single direction.
3. **Sort columns:** sort `dim_month[month]` by `fiscal_month`, `dim_line[line]` by `sort`, and `dim_division[division]` by `sort`. Otherwise February won't come first.
4. **New measures:** paste in the measures from `measures.dax`.
5. **Pages to build:**
   - **Overview:** cards for Revenue (Actual), EBITDA (Actual) and EBITDA (Outlook FY). Add a line chart of Outlook vs Budget by month, filtered to Net revenue.
   - **Variance:** a matrix with lines on rows, and Budget, Actual, Variance, Variance % and Fav / Unfav as values. Add conditional formatting on Variance (favorable).
   - **Divisions:** a clustered bar of Revenue (Actual) vs Revenue (Budget YTD) by division, plus AUR (Actual) vs AUR (Budget YTD).

The `.pbix` file isn't in the repository, because it can't be produced from code. The data and measures are exactly what it's built from.
