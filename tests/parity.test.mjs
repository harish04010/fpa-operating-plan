// The browser model must reproduce the workbook.
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";
import { budgetMonth, outlook } from "../site/model.js";

const data = JSON.parse(readFileSync(new URL("../site/data.json", import.meta.url)));
const close = (a, b, tol = 1e-6) => assert.ok(Math.abs(a - b) <= tol, `${a} vs ${b}`);

test("budget months match the Budget sheet", () => {
  data.months.forEach((_, m) => {
    const b = budgetMonth(data.drivers, m);
    close(b.revenue, data.budget_excel.revenue[m]);
    close(b.ebitda, data.budget_excel.ebitda[m]);
  });
});

test("each scenario's full-year EBITDA and revenue match the Scenarios sheet", () => {
  for (const [name, [traffic, aur, freight, wage]] of Object.entries(data.scenarios)) {
    const o = outlook(data, { traffic, aur, freight, wage });
    close(o.oFY.ebitda, data.excel.scenario_ebitda[name]);
    close(o.oFY.revenue, data.excel.scenario_revenue[name]);
  }
});

test("every cell of the sensitivity grid matches", () => {
  const s = data.excel.sensitivity;
  s.traffic.forEach((t, i) => s.aur.forEach((a, j) => {
    close(outlook(data, { traffic: t, aur: a, freight: 1, wage: 1 }).oFY.ebitda, s.ebitda[i][j]);
  }));
});

test("the outlook bridge adds up exactly", () => {
  const o = outlook(data, { traffic: 0.97, aur: 1.02, freight: 1.1, wage: 1.03 });
  const parts = o.steps.slice(0, -1).reduce((s, [, v]) => s + v, 0);
  close(parts, o.steps.at(-1)[1], 1e-9);
});
