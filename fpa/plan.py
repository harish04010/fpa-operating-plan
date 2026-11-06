"""
The plan's drivers, and the actuals the year actually delivered.

Fiscal year 2026 runs February 2026 to January 2027, the usual retail year,
which ends after the holiday season and January clearance. Actuals are in
through September (eight months); October to January are forecast.

The plan is driver-based, as an FP&A team would build it:
  transactions   = store traffic x conversion
  units          = transactions x units per transaction x division mix
  revenue        = units x average unit retail (AUR), by division
  product cost   = revenue x (1 - merchandise margin), by division
  freight        = units x freight cost per unit
  store payroll  = (fixed hours + variable hours per 1,000 visits x traffic) x wage
  marketing      = % of revenue
  occupancy, other SG&A, D&A: fixed by month

Every number is fictional and generated for this case study. The cost
structure is calibrated to ratios typical of large US specialty retailers:
merchandise margin in the low-to-mid 30s, store payroll near 10% of sales,
occupancy near 5%, and EBITDA near 11%.
"""
import random

MONTHS = ["Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan"]
ACTUAL_MONTHS = 8
DIVISIONS = ["Outdoors", "Sports & Recreation", "Apparel", "Footwear"]

DRIVERS = {
    # store visits (millions) by month: spring sports, summer, back-to-school, holiday
    "traffic": [9.8, 11.2, 12.0, 12.6, 12.9, 12.4, 13.8, 11.0, 10.6, 13.4, 17.8, 10.2],
    "conversion": 0.46,
    "upt": 2.3,                         # units per transaction
    "mix": [0.27, 0.27, 0.27, 0.19],    # share of units by division
    "aur": [31.50, 27.80, 22.60, 54.20],        # average unit retail, $
    "margin": [0.330, 0.315, 0.395, 0.345],     # merchandise margin rate by division
    "freight_per_unit": 0.62,
    "fixed_hours": 1.30,                # store labor hours (millions) per month regardless of traffic
    "hours_per_k_visits": 0.120,        # variable hours (millions) per million visits
    "wage": 15.40,
    "marketing_pct": 0.032,
    "occupancy": 22.0,                  # $M per month
    "other_sga": 12.0,                  # $M per month
    "da": 12.8,                         # $M per month
    "interest": 3.1,                    # $M per month
    "tax_rate": 0.235,
}

SCENARIOS = {
    #            traffic, AUR, freight, wage
    "Base":     [1.00, 1.00, 1.00, 1.00],
    "Upside":   [1.03, 1.01, 0.97, 1.00],
    "Downside": [0.94, 0.99, 1.08, 1.02],
}


def actuals(seed: int = 26) -> dict:
    """
    What happened, February to September, with the variances a retailer of
    this kind would recognize:
      - a cold, wet spring: traffic 4-7% under plan Feb-May, then a little over
      - price increases: AUR 2-3% above plan, strongest in Footwear
      - mix drifting to Footwear and away from Outdoors
      - ocean freight spiking in the second quarter
      - overtime in the July-August back-to-school ramp
      - marketing over plan in August for the back-to-school campaign
    """
    rng = random.Random(seed)
    d = DRIVERS
    out = {"traffic": [], "units": [], "revenue": [], "cogs": [], "freight": [], "payroll": [], "marketing": [],
           "occupancy": [], "other_sga": [], "da": [], "interest": []}
    for m in range(ACTUAL_MONTHS):
        t_adj = [0.94, 0.93, 0.95, 0.96, 1.01, 1.02, 1.02, 1.01][m] * rng.uniform(0.995, 1.005)
        traffic = d["traffic"][m] * t_adj
        conv = d["conversion"] * rng.uniform(0.99, 1.01)
        mix = [d["mix"][0] - 0.012, d["mix"][1], d["mix"][2] - 0.004, d["mix"][3] + 0.016]
        aur_adj = [1.018, 1.020, 1.022, 1.034]
        units_total = traffic * conv * d["upt"] * rng.uniform(0.995, 1.005)
        units = [units_total * s for s in mix]
        revenue = [u * d["aur"][i] * aur_adj[i] for i, u in enumerate(units)]
        margin = [d["margin"][i] - (0.004 if i == 0 else 0) for i in range(4)]  # Outdoors clearance in spring
        cogs = [r * (1 - margin[i]) for i, r in enumerate(revenue)]
        fpu = d["freight_per_unit"] * ([1.0, 1.0, 1.06, 1.17, 1.21, 1.12, 1.04, 1.0][m])
        hours = d["fixed_hours"] + d["hours_per_k_visits"] * traffic
        wage = d["wage"] * (1.055 if m in (5, 6) else 1.01)  # overtime in Jul-Aug
        rev_total = sum(revenue)
        out["traffic"].append(round(traffic, 4))
        out["units"].append([round(u, 4) for u in units])
        out["revenue"].append([round(r, 3) for r in revenue])
        out["cogs"].append([round(c, 3) for c in cogs])
        out["freight"].append(round(sum(units) * fpu, 3))
        out["payroll"].append(round(hours * wage, 3))
        out["marketing"].append(round(rev_total * (d["marketing_pct"] * (1.18 if m == 6 else 1.0)), 3))
        out["occupancy"].append(round(d["occupancy"] * rng.uniform(0.995, 1.01), 3))
        out["other_sga"].append(round(d["other_sga"] * rng.uniform(0.97, 1.05), 3))
        out["da"].append(d["da"])
        out["interest"].append(d["interest"])
    return out
