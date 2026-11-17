import copy
import html
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

import engine

AUTHOR = "Mohamed Ibrahiem"
WORKBOOK = Path(__file__).parent / "SupplyChain_Interview_Project.xlsx"

st.set_page_config(
    page_title="Network Materials Supply Planning",
    layout="wide",
    initial_sidebar_state="collapsed",
)

INK, MUTED, LINE, ACCENT = "#1B2430", "#5B6675", "#E3E6EA", "#1F3A5F"
CLASS_COLORS = {"A": "#1F3A5F", "B": "#6C8EBF", "C": "#B9C4D3"}
STATUS_COLORS = {"REORDER": "#B42318", "OVERSTOCK": "#B54708", "OK": "#3B7A57"}

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"], .stApp { font-family: 'Inter', 'Segoe UI', sans-serif; color: #1B2430; }
#MainMenu, footer, [data-testid="stHeader"], [data-testid="stToolbar"],
[data-testid="stDecoration"], .stDeployButton { display: none !important; visibility: hidden; }
.block-container { padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1240px; }
h1, h2, h3, h4 { font-family: 'Inter', sans-serif; letter-spacing: -0.01em; color: #1B2430; }
.hdr { display:flex; justify-content:space-between; align-items:flex-end; gap:16px;
       border-bottom: 1px solid #E3E6EA; padding-bottom: 14px; margin-bottom: 6px; flex-wrap: wrap; }
.hdr h1 { font-size: 1.55rem; font-weight: 650; margin: 0; padding: 0; }
.hdr .sub { color:#5B6675; font-size: 0.92rem; margin-top: 4px; }
.tag { font-size: 0.74rem; font-weight: 600; color:#5B6675; border:1px solid #CFD5DC; border-radius: 4px;
       padding: 3px 8px; text-transform: uppercase; letter-spacing: .04em; white-space: nowrap; }
.stTabs [data-baseweb="tab-list"] { gap: 26px; border-bottom: 1px solid #E3E6EA; }
.stTabs [data-baseweb="tab"] { padding: 10px 0; font-weight: 500; color:#5B6675; background: transparent; }
.stTabs [aria-selected="true"] { color: #1F3A5F; }
.sec { font-size: 0.95rem; font-weight: 600; margin: 22px 0 8px 0; color:#1B2430; }
.kpi { background:#fff; border:1px solid #E3E6EA; border-radius: 6px; padding: 14px 16px; height: 100%; }
.kpi-label { font-size: 0.76rem; color:#5B6675; font-weight: 500; text-transform: uppercase; letter-spacing: .04em; }
.kpi-value { font-size: 1.55rem; font-weight: 650; margin-top: 4px; font-variant-numeric: tabular-nums; }
.kpi-sub { font-size: 0.78rem; color:#5B6675; margin-top: 2px; min-height: 1.1em; }
.panel { background:#fff; border:1px solid #E3E6EA; border-radius: 6px; padding: 6px 18px 12px 18px; }
.find { margin: 0; padding: 0; list-style: none; }
.find li { padding: 10px 0; border-bottom: 1px solid #EEF0F3; font-size: 0.9rem; line-height: 1.5; }
.find li:last-child { border-bottom: none; }
table.t { width: 100%; border-collapse: collapse; background:#fff; border:1px solid #E3E6EA; font-size: 0.86rem; }
table.t th { text-align:left; font-weight:600; font-size:0.74rem; color:#5B6675; text-transform: uppercase;
             letter-spacing:.04em; padding: 9px 12px; background:#FAFBFC; border-bottom:1px solid #E3E6EA; }
table.t td { padding: 9px 12px; border-bottom: 1px solid #EEF0F3; font-variant-numeric: tabular-nums; }
table.t tr:last-child td { border-bottom: none; }
table.t .r { text-align: right; }
.badge { display:inline-block; font-size:0.72rem; font-weight:600; padding:2px 8px; border-radius: 4px; letter-spacing:.03em; }
.note { font-size: 0.82rem; color:#5B6675; line-height: 1.55; }
.foot { margin-top: 36px; padding-top: 12px; border-top: 1px solid #E3E6EA; font-size: 0.78rem; color:#5B6675; }
.mx td, .mx th { text-align:center; }
.mx td { height: 64px; }
</style>
""",
    unsafe_allow_html=True,
)


# ---------- helpers ----------
def money(x):
    x = float(x)
    if abs(x) >= 1e6:
        return f"{x / 1e6:,.1f}M"
    if abs(x) >= 1e3:
        return f"{x / 1e3:,.0f}K"
    return f"{x:,.0f}"


def kpi(label, value, sub=""):
    return (
        f'<div class="kpi"><div class="kpi-label">{html.escape(label)}</div>'
        f'<div class="kpi-value">{value}</div><div class="kpi-sub">{sub}</div></div>'
    )


def badge(text):
    color = STATUS_COLORS.get(text)
    if color is None:
        color = {"A": "#1F3A5F", "B": "#4D6F9E", "C": "#6B7686"}.get(text, "#4A5565")
    return (
        f'<span class="badge" style="color:{color};background:{color}14;'
        f'border:1px solid {color}40">{html.escape(str(text))}</span>'
    )


def html_table(df, right=(), badges=(), raw=()):
    head = "".join(
        f'<th class="{"r" if c in right else ""}">{html.escape(str(c))}</th>' for c in df.columns
    )
    body = ""
    for _, row in df.iterrows():
        cells = ""
        for c in df.columns:
            v = row[c]
            if c in badges:
                txt = badge(v)
            elif c in raw:
                txt = str(v)
            else:
                txt = html.escape(str(v))
            cells += f'<td class="{"r" if c in right else ""}">{txt}</td>'
        body += f"<tr>{cells}</tr>"
    return f'<table class="t"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>'


def style_chart(ch):
    return (
        ch.configure_view(strokeWidth=0)
        .configure_axis(
            labelFont="Inter", titleFont="Inter", labelColor=MUTED, titleColor=MUTED,
            gridColor="#EEF0F3", domainColor="#CFD5DC", tickColor="#CFD5DC", labelFontSize=11,
            titleFontSize=11, titleFontWeight="normal",
        )
        .configure_legend(labelFont="Inter", titleFont="Inter", labelColor=MUTED, titleColor=MUTED, orient="top")
    )


def section(title):
    st.markdown(f'<div class="sec">{html.escape(title)}</div>', unsafe_allow_html=True)


# ---------- data ----------
if not WORKBOOK.exists():
    st.error(f"Data file not found: {WORKBOOK.name}. Place it in the same folder as app.py.")
    st.stop()


@st.cache_data
def load(path):
    return engine.load_inputs(path)


params, items, months, demand, sups = load(str(WORKBOOK))
A, S, SERIES, K = engine.run(params, items, demand, sups)

# ---------- header ----------
st.markdown(
    """
<div class="hdr">
  <div>
    <h1>Network Materials Supply Planning</h1>
    <div class="sub">Demand forecasting, inventory policy and supplier performance &nbsp;|&nbsp; EGP &nbsp;|&nbsp;
    24 months of demand, 12 items, 5 suppliers</div>
  </div>
  <div class="tag">Illustrative data</div>
</div>
""",
    unsafe_allow_html=True,
)

tab_ov, tab_inv, tab_fc, tab_sup, tab_sc, tab_m = st.tabs(
    ["Overview", "Inventory", "Demand forecast", "Suppliers", "Scenarios", "Method"]
)

# ======================= OVERVIEW =======================
with tab_ov:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    cols = st.columns(6)
    cards = [
        kpi("Annual spend", f"{money(K['spend'])}", "EGP, last 12 months"),
        kpi("Inventory value", f"{money(K['inventory_value'])}", "EGP, on hand"),
        kpi("Inventory turns", f"{K['turns']:.1f}x", "spend / inventory value"),
        kpi("Below reorder point", f"{K['reorder']}", f"orders worth {money(K['order_value'])} EGP"),
        kpi("Excess stock", f"{money(K['excess_value'])}", f"EGP, {K['overstock']} items above max"),
        kpi("Forecast error", f"{K['mape']:.0%}", "avg MAPE, best method"),
    ]
    for c, h in zip(cols, cards):
        c.markdown(h, unsafe_allow_html=True)

    left, right = st.columns([3, 2], gap="large")
    with left:
        section("Spend concentration by item (Pareto)")
        d = A.sort_values("annual_value", ascending=False).copy()
        d["value_m"] = d.annual_value / 1e6
        order = d["item"].tolist()
        base = alt.Chart(d).encode(
            x=alt.X("item:N", sort=order, title=None, axis=alt.Axis(labelAngle=-35, labelLimit=150))
        )
        bars = base.mark_bar(size=22).encode(
            y=alt.Y("value_m:Q", title="Annual value (EGP m)"),
            color=alt.Color(
                "abc:N",
                scale=alt.Scale(domain=list(CLASS_COLORS), range=list(CLASS_COLORS.values())),
                legend=alt.Legend(title="Class"),
            ),
            tooltip=[
                alt.Tooltip("item:N", title="Item"),
                alt.Tooltip("value_m:Q", title="EGP m", format=",.1f"),
                alt.Tooltip("cum_share:Q", title="Cumulative", format=".0%"),
            ],
        )
        line = base.mark_line(color="#1B2430", strokeWidth=1.5, point=alt.OverlayMarkDef(color="#1B2430", size=28)).encode(
            y=alt.Y("cum_share:Q", title="Cumulative share", axis=alt.Axis(format="%"), scale=alt.Scale(domain=[0, 1]))
        )
        st.altair_chart(
            style_chart(alt.layer(bars, line).resolve_scale(y="independent").properties(height=330)),
            use_container_width=True,
        )
    with right:
        section("Key findings")
        nA = int((A.abc == "A").sum())
        shA = A.loc[A.abc == "A", "share"].sum()
        worst_cover = A.sort_values("cover", ascending=False).iloc[0]
        top_sup = S.sort_values("spend_share", ascending=False).iloc[0]
        worst_mape = A.sort_values("mape", ascending=False).iloc[0]
        items_txt = [
            f"<b>{nA} of {len(A)} items</b> are Class A and account for <b>{shA:.0%}</b> of annual spend. "
            "This is where planning effort pays back.",
            f"<b>{K['reorder']} items</b> are below their reorder point. Suggested orders total "
            f"<b>EGP {money(K['order_value'])}</b>.",
            f"<b>{K['overstock']} items</b> sit above their maximum level, tying up "
            f"<b>EGP {money(K['excess_value'])}</b>. {html.escape(worst_cover['item'])} holds "
            f"{worst_cover['cover']:.1f} months of cover.",
        ]
        if top_sup.spend_share > params["conc"]:
            items_txt.append(
                f"<b>{html.escape(top_sup.supplier)}</b> supplies <b>{top_sup.spend_share:.0%}</b> of spend "
                f"(limit {params['conc']:.0%}) with on-time delivery of {top_sup.ontime:.0%}. "
                "A second source is recommended."
            )
        else:
            items_txt.append("No supplier exceeds the spend concentration limit.")
        items_txt.append(
            f"Back-test error averages <b>{K['mape']:.0%}</b>. The highest is "
            f"{html.escape(worst_mape['item'])} at {worst_mape['mape']:.0%}; items like this are better planned "
            "from the rollout schedule than from history."
        )
        st.markdown(
            '<div class="panel"><ul class="find">' + "".join(f"<li>{t}</li>" for t in items_txt) + "</ul></div>",
            unsafe_allow_html=True,
        )

    section("Action list")
    act = A[A.status != "OK"].copy()
    act["_o"] = act.status.map({"REORDER": 0, "OVERSTOCK": 1})
    act = act.sort_values(["_o", "annual_value"], ascending=[True, False])
    tbl = pd.DataFrame({
        "Item": act["item"], "Class": act["class"], "Status": act.status,
        "On hand": act.stock.map("{:,.0f}".format), "Reorder point": act.rop.map("{:,.0f}".format),
        "Suggested order": act.order_qty.map(lambda v: f"{v:,.0f}" if v else "-"),
        "Excess (EGP)": act.excess_value.map(lambda v: f"{v:,.0f}" if v else "-"),
    })
    st.markdown(
        html_table(tbl, right=("On hand", "Reorder point", "Suggested order", "Excess (EGP)"), badges=("Status", "Class")),
        unsafe_allow_html=True,
    )

# ======================= INVENTORY =======================
with tab_inv:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    c1, c2 = st.columns([2, 3], gap="large")
    with c1:
        section("ABC-XYZ matrix (items and share of spend)")
        mx = '<table class="t mx"><thead><tr><th></th><th>X  stable</th><th>Y  variable</th><th>Z  intermittent</th></tr></thead><tbody>'
        for abc in "ABC":
            mx += f"<tr><th>{abc}</th>"
            for xyz in "XYZ":
                sub = A[(A.abc == abc) & (A.xyz == xyz)]
                if len(sub):
                    mx += f'<td><b>{len(sub)}</b><br><span class="note">{sub.share.sum():.0%} of spend</span></td>'
                else:
                    mx += '<td><span class="note">-</span></td>'
            mx += "</tr>"
        st.markdown(mx + "</tbody></table>", unsafe_allow_html=True)
        st.markdown(
            '<p class="note" style="margin-top:10px">Class A items with stable demand (AX) are candidates for '
            "automated replenishment. Intermittent items (Z) should be planned from project schedules.</p>",
            unsafe_allow_html=True,
        )
    with c2:
        section("Months of cover by item")
        d = A.sort_values("cover", ascending=False)
        ch = (
            alt.Chart(d)
            .mark_bar(size=18)
            .encode(
                y=alt.Y("item:N", sort=d["item"].tolist(), title=None, axis=alt.Axis(labelLimit=200)),
                x=alt.X("cover:Q", title="Months of cover (on hand / forecast monthly demand)"),
                color=alt.Color(
                    "status:N",
                    scale=alt.Scale(domain=list(STATUS_COLORS), range=list(STATUS_COLORS.values())),
                    legend=alt.Legend(title="Status"),
                ),
                tooltip=[alt.Tooltip("item:N", title="Item"), alt.Tooltip("cover:Q", title="Months", format=".1f"),
                         alt.Tooltip("status:N", title="Status")],
            )
            .properties(height=330)
        )
        st.altair_chart(style_chart(ch), use_container_width=True)

    section("Inventory policy by item")
    pol = pd.DataFrame({
        "Item": A["item"], "Class": A["class"], "Forecast / month": A.forecast.round(0),
        "Lead time (d)": A.lead, "Safety stock": A.safety, "Reorder point": A.rop, "EOQ": A.eoq,
        "Max level": A.max_level, "On hand": A.stock, "On order": A.on_order,
        "Cover (months)": A.cover.round(1), "Status": A.status, "Suggested order": A.order_qty,
    })
    st.dataframe(
        pol, use_container_width=True, hide_index=True,
        column_config={
            "Forecast / month": st.column_config.NumberColumn(format="%d"),
            "Lead time (d)": st.column_config.NumberColumn(format="%d"),
            "Safety stock": st.column_config.NumberColumn(format="%d"),
            "Reorder point": st.column_config.NumberColumn(format="%d"),
            "EOQ": st.column_config.NumberColumn(format="%d"),
            "Max level": st.column_config.NumberColumn(format="%d"),
            "On hand": st.column_config.NumberColumn(format="%d"),
            "On order": st.column_config.NumberColumn(format="%d"),
            "Cover (months)": st.column_config.NumberColumn(format="%.1f"),
            "Suggested order": st.column_config.NumberColumn(format="%d"),
        },
    )

    section("Item detail")
    pick = st.selectbox("Item", A["item"].tolist(), label_visibility="collapsed")
    r = A[A["item"] == pick].iloc[0]
    dc = st.columns(5)
    dc[0].markdown(kpi("Reorder point", f"{r.rop:,.0f}", f"safety stock {r.safety:,.0f}"), unsafe_allow_html=True)
    dc[1].markdown(kpi("Order quantity (EOQ)", f"{r.eoq:,.0f}", f"max level {r.max_level:,.0f}"), unsafe_allow_html=True)
    dc[2].markdown(kpi("Inventory position", f"{r.position:,.0f}", f"on hand {r.stock:,.0f} + on order {r.on_order:,.0f}"), unsafe_allow_html=True)
    dc[3].markdown(kpi("Service target", f"{params['service'][r.abc]:.0%}", f"Class {r.abc}, z = {r.z:.2f}"), unsafe_allow_html=True)
    dc[4].markdown(kpi("Status", badge(r.status), f"{r.cover:.1f} months of cover"), unsafe_allow_html=True)
    if r.status == "REORDER":
        msg = f"Place an order for {r.order_qty:,.0f} {items.unit[A.index[A['item'] == pick][0]]} now: position is below the reorder point of {r.rop:,.0f}."
    elif r.status == "OVERSTOCK":
        msg = f"On-hand stock exceeds the maximum level by {r.stock - r.max_level:,.0f} units (EGP {r.excess_value:,.0f}). Hold orders and review redistribution."
    else:
        msg = "Position is within policy. No action needed."
    st.markdown(f'<p class="note" style="margin-top:10px">{html.escape(msg)}</p>', unsafe_allow_html=True)

# ======================= FORECAST =======================
with tab_fc:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    fpick = st.selectbox("Item", A["item"].tolist(), key="fc_item")
    fi = A.index[A["item"] == fpick][0]
    fr = A.loc[fi]
    f = SERIES[fr["id"]]
    k4 = st.columns(4)
    k4[0].markdown(kpi("MAPE, exponential smoothing", f"{fr.mape_es:.1%}", f"alpha = {params['alpha']}"), unsafe_allow_html=True)
    k4[1].markdown(kpi("MAPE, moving average", f"{fr.mape_ma:.1%}", "3-month window"), unsafe_allow_html=True)
    k4[2].markdown(kpi("Selected method", fr.method, "lowest back-test error"), unsafe_allow_html=True)
    k4[3].markdown(kpi("Forecast, next month", f"{fr.forecast:,.0f}", f"{items.unit[fi]} (Nov-2026)"), unsafe_allow_html=True)

    section("Actual demand and back-tested forecasts")
    n = len(months)
    long = pd.DataFrame({
        "Month": months * 3,
        "Series": ["Actual"] * n + ["Exponential smoothing"] * n + ["Moving average (3)"] * n,
        "Units": list(demand[fi]) + list(f["es"]) + list(f["ma"]),
    })
    xenc = alt.X("Month:N", sort=months, title=None, axis=alt.Axis(labelAngle=-45))
    lines = (
        alt.Chart(long)
        .mark_line(strokeWidth=2)
        .encode(
            x=xenc,
            y=alt.Y("Units:Q", title=f"Units ({items.unit[fi]})", scale=alt.Scale(zero=False)),
            color=alt.Color(
                "Series:N",
                scale=alt.Scale(domain=["Actual", "Exponential smoothing", "Moving average (3)"],
                                range=["#1B2430", "#2F6FB3", "#C2761A"]),
                legend=alt.Legend(title=None),
            ),
            strokeDash=alt.condition(alt.datum.Series == "Actual", alt.value([1, 0]), alt.value([5, 3])),
        )
    )
    rule = alt.Chart(pd.DataFrame({"Month": [months[12]]})).mark_rule(color="#9AA5B1", strokeDash=[2, 3]).encode(x=alt.X("Month:N", sort=months))
    st.altair_chart(style_chart(alt.layer(lines, rule).properties(height=340)), use_container_width=True)
    st.markdown(
        f'<p class="note">The dotted vertical line marks the start of the 12-month back-test window ({months[12]}). '
        "MAPE is measured over that window only.</p>",
        unsafe_allow_html=True,
    )

    section("Forecast accuracy, all items")
    acc = pd.DataFrame({
        "Item": A["item"], "Class": A["class"], "MAPE, exp. smoothing": A.mape_es.map("{:.1%}".format),
        "MAPE, moving avg": A.mape_ma.map("{:.1%}".format), "Selected": A.method,
        "Next month": A.forecast.map("{:,.0f}".format),
    })
    st.markdown(html_table(acc, right=("MAPE, exp. smoothing", "MAPE, moving avg", "Next month"), badges=("Class",)), unsafe_allow_html=True)

# ======================= SUPPLIERS =======================
with tab_sup:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    section("Supplier scorecard, last 12 months")
    sc = pd.DataFrame({
        "Supplier": S.supplier, "Items": S.n_items, "On-time": S.ontime.map("{:.0%}".format),
        "Quality": S.quality.map("{:.1%}".format), "Lead-time reliability": S.lt_rel.map("{:.0%}".format),
        "Price vs market": S.price_index.map("{:.2f}".format), "Spend share": S.spend_share.map("{:.0%}".format),
        "Score": S.score.map("{:.1f}".format), "Tier": S.tier, "Recommended action": S.action,
    })
    st.markdown(
        html_table(sc, right=("Items", "On-time", "Quality", "Lead-time reliability", "Price vs market", "Spend share", "Score"), badges=("Tier",)),
        unsafe_allow_html=True,
    )
    c1, c2 = st.columns(2, gap="large")
    with c1:
        section("Share of spend")
        ch = alt.Chart(S).mark_bar(size=26, color=ACCENT).encode(
            x=alt.X("supplier:N", title=None, axis=alt.Axis(labelAngle=0)),
            y=alt.Y("spend_share:Q", title=None, axis=alt.Axis(format="%")),
            tooltip=[alt.Tooltip("supplier:N", title="Supplier"), alt.Tooltip("spend_share:Q", title="Spend share", format=".0%")],
        )
        lim = alt.Chart(pd.DataFrame({"y": [params["conc"]]})).mark_rule(color="#B42318", strokeDash=[5, 3]).encode(y="y:Q")
        st.altair_chart(style_chart(alt.layer(ch, lim).properties(height=260)), use_container_width=True)
        st.markdown(f'<p class="note">Dashed line: concentration limit of {params["conc"]:.0%} of spend.</p>', unsafe_allow_html=True)
    with c2:
        section("Performance score")
        ch = alt.Chart(S).mark_bar(size=26, color="#6C8EBF").encode(
            x=alt.X("supplier:N", title=None, axis=alt.Axis(labelAngle=0)),
            y=alt.Y("score:Q", title=None, scale=alt.Scale(domain=[50, 100])),
            tooltip=[alt.Tooltip("supplier:N", title="Supplier"), alt.Tooltip("score:Q", title="Score", format=".1f")],
        )
        rules = alt.Chart(pd.DataFrame({"y": [params["tier_a"], params["tier_b"]]})).mark_rule(color="#5B6675", strokeDash=[5, 3]).encode(y="y:Q")
        st.altair_chart(style_chart(alt.layer(ch, rules).properties(height=260)), use_container_width=True)
        st.markdown(f'<p class="note">Dashed lines: Tier A from {params["tier_a"]}, Tier B from {params["tier_b"]}.</p>', unsafe_allow_html=True)
    st.markdown(
        f'<p class="note">Score weights: on-time delivery {params["w_ontime"]:.0%}, quality {params["w_quality"]:.0%}, '
        f'lead-time reliability {params["w_lt"]:.0%}, price vs market {params["w_price"]:.0%}.</p>',
        unsafe_allow_html=True,
    )

# ======================= SCENARIOS =======================
DEFAULTS = {
    "sc_svc": 100 * params["service"]["A"], "sc_alpha": float(params["alpha"]),
    "sc_oc": int(params["order_cost"]), "sc_h": float(params["holding"]),
    "sc_delay": 0, "sc_up": 0,
}


def reset_scenario():
    for k_, v_ in DEFAULTS.items():
        st.session_state[k_] = v_


with tab_sc:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    st.markdown(
        '<p class="note">Change the assumptions below. The whole model is recalculated and compared against the base case from the workbook.</p>',
        unsafe_allow_html=True,
    )
    c1, c2, c3 = st.columns(3, gap="large")
    svc = c1.slider("Service level, Class A (%)", 90.0, 99.5, DEFAULTS["sc_svc"], 0.5, key="sc_svc")
    delay = c1.slider("Supplier delay on all lead times (days)", 0, 30, DEFAULTS["sc_delay"], 1, key="sc_delay")
    up = c2.slider("Demand uplift, e.g. rollout surge (%)", 0, 30, DEFAULTS["sc_up"], 1, key="sc_up")
    alpha = c2.slider("Smoothing factor (alpha)", 0.10, 0.90, DEFAULTS["sc_alpha"], 0.05, key="sc_alpha")
    oc = c3.slider("Ordering cost per PO (EGP)", 500, 3000, DEFAULTS["sc_oc"], 100, key="sc_oc")
    hc = c3.slider("Annual holding cost (% of unit cost)", 0.10, 0.35, DEFAULTS["sc_h"], 0.01, key="sc_h")
    st.button("Reset to base case", on_click=reset_scenario)

    q = copy.deepcopy(params)
    q.update(alpha=alpha, order_cost=oc, holding=hc, lt_delay=delay, demand_uplift=up / 100)
    q["z"]["A"] = engine.z_for(svc / 100)
    A2, S2, _, K2 = engine.run(q, items, demand, sups)

    section("Impact on key measures")
    rows = [
        ("Items below reorder point", K["reorder"], K2["reorder"], "{:,.0f}"),
        ("Items above maximum level", K["overstock"], K2["overstock"], "{:,.0f}"),
        ("Suggested order value (EGP)", K["order_value"], K2["order_value"], "{:,.0f}"),
        ("Safety stock value (EGP)", K["safety_value"], K2["safety_value"], "{:,.0f}"),
        ("Excess stock value (EGP)", K["excess_value"], K2["excess_value"], "{:,.0f}"),
    ]
    cmp_df = pd.DataFrame({
        "Measure": [r_[0] for r_ in rows],
        "Base case": [r_[3].format(r_[1]) for r_ in rows],
        "Scenario": [r_[3].format(r_[2]) for r_ in rows],
        "Change": [("+" if r_[2] - r_[1] > 0 else "") + r_[3].format(r_[2] - r_[1]) if r_[2] != r_[1] else "-" for r_ in rows],
    })
    st.markdown(html_table(cmp_df, right=("Base case", "Scenario", "Change")), unsafe_allow_html=True)

    cc1, cc2 = st.columns([3, 2], gap="large")
    with cc1:
        section("Safety stock value by item (EGP m)")
        cmpd = pd.concat([
            pd.DataFrame({"Item": A["item"], "Case": "Base", "Value": A.safety * A.cost / 1e6}),
            pd.DataFrame({"Item": A2["item"], "Case": "Scenario", "Value": A2.safety * A2.cost / 1e6}),
        ])
        order = A.assign(v=A.safety * A.cost).sort_values("v", ascending=False)["item"].tolist()
        ch = alt.Chart(cmpd).mark_bar().encode(
            x=alt.X("Item:N", sort=order, title=None, axis=alt.Axis(labelAngle=-35, labelLimit=140)),
            xOffset="Case:N",
            y=alt.Y("Value:Q", title=None),
            color=alt.Color("Case:N", scale=alt.Scale(domain=["Base", "Scenario"], range=["#B9C4D3", ACCENT]), legend=alt.Legend(title=None)),
            tooltip=[alt.Tooltip("Item:N"), alt.Tooltip("Case:N"), alt.Tooltip("Value:Q", format=",.2f", title="EGP m")],
        ).properties(height=300)
        st.altair_chart(style_chart(ch), use_container_width=True)
    with cc2:
        section("Items whose status changes")
        chg = A[["id", "item", "status"]].merge(A2[["id", "status", "order_qty"]], on="id", suffixes=("_b", "_s"))
        chg = chg[chg.status_b != chg.status_s]
        if chg.empty:
            st.markdown('<p class="note">No item changes status under these assumptions.</p>', unsafe_allow_html=True)
        else:
            t = pd.DataFrame({"Item": chg["item"], "Base": chg.status_b, "Scenario": chg.status_s})
            st.markdown(html_table(t, badges=("Base", "Scenario")), unsafe_allow_html=True)

# ======================= METHOD =======================
with tab_m:
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    m1, m2 = st.columns(2, gap="large")
    with m1:
        section("How each result is calculated")
        st.markdown("**Forecast (exponential smoothing)**")
        st.latex(r"F_{t+1} = \alpha \, D_t + (1-\alpha) \, F_t")
        st.markdown("**Forecast error**")
        st.latex(r"\mathrm{MAPE} = \frac{1}{n}\sum \frac{|D_t - F_t|}{D_t}")
        st.markdown("**Safety stock**")
        st.latex(r"SS = z \sqrt{\tfrac{L}{30}\,\sigma_D^{2} + F^{2}\left(\tfrac{\sigma_L}{30}\right)^{2}}")
        st.markdown("**Reorder point and order quantity**")
        st.latex(r"ROP = F \cdot \tfrac{L}{30} + SS \qquad EOQ = \sqrt{\frac{2 \cdot 12F \cdot S}{c \cdot h}}")
        st.markdown(
            '<p class="note">F: forecast monthly demand. L: lead time in days. S: ordering cost per PO. '
            "c: unit cost. h: annual holding cost rate. An item is reordered when its inventory position "
            "(on hand plus on order) falls below ROP, up to ROP + EOQ.</p>",
            unsafe_allow_html=True,
        )
    with m2:
        section("Assumptions in use")
        asm = pd.DataFrame([
            ("Smoothing factor (alpha)", f"{params['alpha']}"),
            ("Ordering cost per PO", f"EGP {params['order_cost']:,}"),
            ("Annual holding cost", f"{params['holding']:.0%} of unit cost"),
            ("ABC limits", f"A to {params['a_limit']:.0%}, B to {params['b_limit']:.0%} of spend"),
            ("XYZ limits", f"X up to CV {params['x_limit']}, Y up to {params['y_limit']}"),
            ("Service level (A / B / C)", " / ".join(f"{params['service'][k_]:.0%}" for k_ in "ABC")),
            ("Supplier tiers", f"A from {params['tier_a']}, B from {params['tier_b']}"),
            ("Concentration limit", f"{params['conc']:.0%} of spend"),
        ], columns=["Parameter", "Value"])
        st.markdown(html_table(asm), unsafe_allow_html=True)
        section("Limitations")
        st.markdown(
            """
<ul class="note" style="padding-left:18px">
<li>Figures are illustrative and were not taken from company records. The method is ready to run on ERP data.</li>
<li>Forecasting methods are deliberately simple baselines. Production forecasts should be tied to rollout and budget plans, and any machine-learning model should be compared against these baselines on the same MAPE.</li>
<li>Intermittent items (class Z) need schedule-based planning or a method such as Croston.</li>
<li>Quantity discounts, warehouse capacity and transport constraints are not modelled.</li>
</ul>
""",
            unsafe_allow_html=True,
        )

st.markdown(
    f'<div class="foot">Prepared by {html.escape(AUTHOR)} &nbsp;|&nbsp; Source: {WORKBOOK.name} '
    "&nbsp;|&nbsp; All figures are illustrative.</div>",
    unsafe_allow_html=True,
)
