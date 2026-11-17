"""Calculation layer: reads raw inputs from the workbook and recomputes the model.

Every formula mirrors the corresponding Excel formula, so the app and the
workbook always agree. Only input cells are read; results are recalculated here,
which is what allows the scenario page to re-run the whole model.
"""
import math
from statistics import NormalDist

import numpy as np
import pandas as pd
from openpyxl import load_workbook


def load_inputs(path):
    wb = load_workbook(path, data_only=True)
    a = wb["Assumptions"]
    params = {
        "alpha": a["B4"].value,
        "order_cost": a["B5"].value,
        "holding": a["B6"].value,
        "days": a["B7"].value,
        "a_limit": a["B8"].value,
        "b_limit": a["B9"].value,
        "x_limit": a["B10"].value,
        "y_limit": a["B11"].value,
        "service": {k: a[f"B{r}"].value for k, r in zip("ABC", (14, 15, 16))},
        "z": {k: a[f"C{r}"].value for k, r in zip("ABC", (14, 15, 16))},
        "w_ontime": a["B19"].value,
        "w_quality": a["B20"].value,
        "w_lt": a["B21"].value,
        "w_price": a["B22"].value,
        "tier_a": a["B26"].value,
        "tier_b": a["B27"].value,
        "conc": a["B28"].value,
        "lt_delay": 0,
        "demand_uplift": 0.0,
    }

    it = wb["Items"]
    items = []
    r = 5
    while it.cell(r, 1).value:
        items.append([it.cell(r, c).value for c in range(1, 10)])
        r += 1
    items = pd.DataFrame(
        items,
        columns=["id", "item", "unit", "cost", "lead", "lt_sd", "supplier", "stock", "on_order"],
    )

    dm = wb["Demand"]
    months = [dm.cell(4, c).value for c in range(3, 27)]
    demand = np.array(
        [[dm.cell(5 + i, c).value for c in range(3, 27)] for i in range(len(items))],
        dtype=float,
    )

    sp = wb["Suppliers"]
    sups = []
    r = 5
    while sp.cell(r, 1).value:
        sups.append([sp.cell(r, c).value for c in range(1, 9)])
        r += 1
    sups = pd.DataFrame(
        sups,
        columns=["supplier", "pos", "pos_ontime", "units_rec", "units_acc",
                 "lt_promised", "lt_actual", "price_index"],
    )
    return params, items, months, demand, sups


def z_for(level):
    return round(NormalDist().inv_cdf(level), 2)


def es_forecast(d, alpha):
    f = np.full(len(d), np.nan)
    f[1] = d[0]
    for t in range(2, len(d)):
        f[t] = alpha * d[t - 1] + (1 - alpha) * f[t - 1]
    return f, alpha * d[-1] + (1 - alpha) * f[-1]


def ma_forecast(d, n=3):
    f = np.full(len(d), np.nan)
    for t in range(n, len(d)):
        f[t] = d[t - n:t].mean()
    return f, d[-n:].mean()


def mape(d, f):
    return float(np.mean(np.abs(d[12:] - f[12:]) / d[12:]))


def run(params, items, demand, sups):
    p = params
    n = len(items)
    rows = []
    series = {}
    for i in range(n):
        d = demand[i]
        fe, ne = es_forecast(d, p["alpha"])
        fm, nm = ma_forecast(d)
        m_es, m_ma = mape(d, fe), mape(d, fm)
        use_es = m_es <= m_ma
        series[items.id[i]] = {"es": fe, "ma": fm}
        last = d[12:]
        rows.append({
            "id": items.id[i], "item": items.item[i], "supplier": items.supplier[i],
            "cost": items.cost[i], "annual_demand": last.sum(),
            "avg_m": last.mean(), "sd_m": last.std(ddof=1),
            "mape_es": m_es, "mape_ma": m_ma,
            "method": "Exp. smoothing" if use_es else "Moving avg (3)",
            "mape": min(m_es, m_ma),
            "forecast": (ne if use_es else nm) * (1 + p["demand_uplift"]),
            "lead": items.lead[i] + p["lt_delay"], "lt_sd": items.lt_sd[i],
            "stock": items.stock[i], "on_order": items.on_order[i],
        })
    a = pd.DataFrame(rows)
    a["annual_value"] = a.annual_demand * a.cost
    total = a.annual_value.sum()
    a["share"] = a.annual_value / total
    a["rank"] = a.annual_value.rank(ascending=False, method="min").astype(int)
    a["cum_share"] = [a.annual_value[a.annual_value >= v].sum() / total for v in a.annual_value]
    prior = a.cum_share - a.share
    a["abc"] = np.where(prior < p["a_limit"], "A", np.where(prior < p["b_limit"], "B", "C"))
    a["cv"] = a.sd_m / a.avg_m
    a["xyz"] = np.where(a.cv <= p["x_limit"], "X", np.where(a.cv <= p["y_limit"], "Y", "Z"))
    a["class"] = a.abc + a.xyz
    a["z"] = a.abc.map(p["z"])
    lt_m = a.lead / p["days"]
    a["safety"] = np.ceil(a.z * np.sqrt(lt_m * a.sd_m ** 2 + a.forecast ** 2 * (a.lt_sd / p["days"]) ** 2) - 1e-9)
    a["rop"] = np.ceil(a.forecast * lt_m + a.safety - 1e-9)
    a["eoq"] = np.ceil(np.sqrt(2 * a.forecast * 12 * p["order_cost"] / (a.cost * p["holding"])) - 1e-9)
    a["max_level"] = a.rop + a.eoq
    a["position"] = a.stock + a.on_order
    a["cover"] = a.stock / a.forecast
    a["status"] = np.where(a.position < a.rop, "REORDER",
                           np.where(a.stock > a.max_level, "OVERSTOCK", "OK"))
    a["order_qty"] = np.where(a.status == "REORDER", a.max_level - a.position, 0)
    a["excess_value"] = np.maximum(0, a.stock - a.max_level) * a.cost
    a["order_value"] = a.order_qty * a.cost

    s = sups.copy()
    s["ontime"] = s.pos_ontime / s.pos
    s["quality"] = s.units_acc / s.units_rec
    spend = a.groupby("supplier").annual_value.sum()
    s["spend_share"] = s.supplier.map(spend).fillna(0) / total
    s["lt_rel"] = np.minimum(1, s.lt_promised / s.lt_actual)
    s["price_score"] = np.minimum(1, 1 / s.price_index)
    s["score"] = 100 * (s.ontime * p["w_ontime"] + s.quality * p["w_quality"]
                        + s.lt_rel * p["w_lt"] + s.price_score * p["w_price"])
    s["tier"] = np.where(s.score >= p["tier_a"], "A", np.where(s.score >= p["tier_b"], "B", "C"))

    def action(r):
        if r.score < p["tier_b"]:
            return "Corrective action plan"
        if r.spend_share > p["conc"]:
            return "Develop second source"
        return "Strategic partner" if r.tier == "A" else "Monitor monthly"

    s["action"] = s.apply(action, axis=1)
    s["n_items"] = s.supplier.map(a.supplier.value_counts()).fillna(0).astype(int)

    inv_value = float((a.stock * a.cost).sum())
    kpi = {
        "spend": float(total),
        "inventory_value": inv_value,
        "turns": float(total / inv_value),
        "reorder": int((a.status == "REORDER").sum()),
        "overstock": int((a.status == "OVERSTOCK").sum()),
        "excess_value": float(a.excess_value.sum()),
        "order_value": float(a.order_value.sum()),
        "safety_value": float((a.safety * a.cost).sum()),
        "mape": float(a.mape.mean()),
        "supplier_score": float((s.score * s.spend_share).sum()),
        "over_conc": int((s.spend_share > p["conc"]).sum()),
    }
    return a, s, series, kpi
