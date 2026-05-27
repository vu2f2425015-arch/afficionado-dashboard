"""
Afficionado Coffee Roasters — Forecasting & Peak Demand Dashboard
Coffee-themed Streamlit app with high-contrast text, simpler charts, and per-report downloads.
"""
import base64
from pathlib import Path
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# -------- Page config & theme --------
st.set_page_config(
    page_title="Afficionado Coffee Roasters — Forecasting Dashboard",
    page_icon="☕",
    layout="wide",
)

# -------- Dark coffee theme (single, polished) --------
# Coffee palette — shared brand colors
ESPRESSO = "#3E2723"
MOCHA    = "#6F4E37"
CARAMEL  = "#C9853F"   # brighter caramel for visibility
LATTE    = "#E8CDA8"   # lighter latte
CREAM    = "#FFF3DC"   # bright cream for text
GOLD     = "#F2B65A"   # bright gold accent

PAPER         = "#160B07"      # app background — deep espresso
SURFACE       = "#231410"      # cards / plot area
SIDEBAR_BG    = "linear-gradient(180deg, #2A1812 0%, #110704 100%)"
INK           = "#FFF3DC"      # primary text — bright cream (high contrast)
SUBTEXT       = "#E8CDA8"
BORDER        = "#7A4A28"
GRID          = "rgba(255,243,220,0.10)"
HEADING       = "#FFD89B"      # warm cream-gold for headings
TAB_ACTIVE    = GOLD
LEGEND_BG     = "rgba(15,8,5,0.92)"
METRIC_BG     = "linear-gradient(145deg, #2C1A14 0%, #170B07 100%)"
INSIGHT_BG    = "rgba(242,182,90,0.10)"
BTN_BG        = GOLD
BTN_FG        = "#1A0E08"
BTN_HOVER     = LATTE

PALETTE = [GOLD, LATTE, CARAMEL, CREAM, "#E0C097", MOCHA, "#8B5E3C", "#5D3A1F"]

# Plotly layout helper — themed
def style_fig(fig, height=380, legend=True):
    fig.update_layout(
        height=height,
        paper_bgcolor=SURFACE,
        plot_bgcolor=SURFACE,
        font=dict(family="Georgia, serif", color=INK, size=13),
        title=dict(text="", font=dict(family="Georgia, serif", color=HEADING, size=16)),
        margin=dict(l=40, r=20, t=40, b=40),
        xaxis=dict(gridcolor=GRID, linecolor=BORDER, tickfont=dict(color=INK), title_font=dict(color=INK)),
        yaxis=dict(gridcolor=GRID, linecolor=BORDER, tickfont=dict(color=INK), title_font=dict(color=INK)),
        legend=dict(
            orientation="h", y=-0.18, x=0.5, xanchor="center",
            bgcolor=LEGEND_BG,
            bordercolor=BORDER, borderwidth=1,
            font=dict(color=INK, size=12, family="Georgia, serif"),
        ) if legend else dict(),
    )
    return fig

st.markdown(f"""
<style>
.stApp {{ background: {PAPER}; color: {INK}; font-family: Georgia, 'Times New Roman', serif; }}
section[data-testid="stSidebar"] {{ background: {SIDEBAR_BG}; }}
section[data-testid="stSidebar"] * {{ color: {INK} !important; }}
h1, h2, h3, h4 {{ font-family: Georgia, serif; color: {HEADING} !important; letter-spacing: .3px; }}
p, span, label, li {{ color: {INK} !important; }}
.metric-card {{
  background: {METRIC_BG};
  border: 1px solid {BORDER}; border-radius: 16px; padding: 18px 22px;
  box-shadow: 0 4px 14px rgba(0,0,0,.18);
}}
.metric-card .label {{ color: {SUBTEXT} !important; font-size: 12px; text-transform: uppercase; letter-spacing: 1.2px; font-weight: 600; }}
.metric-card .value {{ color: {HEADING} !important; font-size: 28px; font-weight: 700; margin-top: 6px; }}
.hero {{
  color: {CREAM}; padding: 36px 40px; border-radius: 20px; margin-bottom: 22px;
  display: flex; align-items: center; gap: 28px;
  box-shadow: 0 10px 30px rgba(0,0,0,.35);
}}
.hero h1 {{ color: #FFF6E6 !important; margin: 0; font-size: 34px; }}
.hero p  {{ color: {LATTE} !important; margin: 6px 0 0 0; font-size: 15px; }}
.hero .emoji {{ font-size: 72px; line-height: 1; filter: drop-shadow(0 4px 8px rgba(0,0,0,.4)); }}
div[data-testid="stTabs"] button {{ font-family: Georgia, serif; color: {INK} !important; font-size: 15px; }}
div[data-testid="stTabs"] button[aria-selected="true"] {{ color: {TAB_ACTIVE} !important; border-bottom: 3px solid {TAB_ACTIVE} !important; }}
.stDataFrame {{ background: {SURFACE}; border-radius: 12px; }}
.section-card {{
  background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 16px;
  padding: 18px 22px; margin-bottom: 18px; box-shadow: 0 2px 8px rgba(0,0,0,.12);
}}
.insight {{
  background: {INSIGHT_BG}; border-left: 4px solid {TAB_ACTIVE}; padding: 12px 16px;
  border-radius: 8px; color: {INK} !important; font-size: 14px; margin-top: 10px;
}}
.stDownloadButton button, .stButton button {{
  background: {BTN_BG} !important; color: {BTN_FG} !important;
  border-radius: 10px !important; border: none !important; font-family: Georgia, serif !important; font-weight: 600 !important;
}}
.stDownloadButton button:hover, .stButton button:hover {{ background: {BTN_HOVER} !important; }}
</style>
""", unsafe_allow_html=True)

# -------- Data loading --------
DATA_FILE = Path(__file__).parent / "data" / "Afficionado_Coffee_Roasters.xlsx"

@st.cache_data(show_spinner=True)
def load_data() -> pd.DataFrame:
    df = pd.read_excel(DATA_FILE)
    df["transaction_time"] = df["transaction_time"].astype(str)
    df["hour"] = pd.to_datetime(df["transaction_time"], format="%H:%M:%S", errors="coerce").dt.hour
    df = df.sort_values("transaction_id").reset_index(drop=True)
    n_days = 181
    bin_edges = np.linspace(0, len(df), n_days + 1, dtype=int)
    day_idx = np.zeros(len(df), dtype=int)
    for i in range(n_days):
        day_idx[bin_edges[i]:bin_edges[i+1]] = i
    start = pd.Timestamp("2025-01-01")
    df["transaction_date"] = start + pd.to_timedelta(day_idx, unit="D")
    df["dow"] = df["transaction_date"].dt.day_name()
    df["revenue"] = df["transaction_qty"] * df["unit_price"]
    return df

df_all = load_data()

# -------- Sidebar --------
with st.sidebar:
    st.markdown("### ☕ Filters")
    st.caption("Quick-select chips")

    locs = sorted(df_all["store_location"].unique().tolist())
    cats = sorted(df_all["product_category"].unique().tolist())

    loc_sel = st.pills("Store location", locs, default=locs, selection_mode="multi")
    cat_sel = st.pills("Product category", cats, default=cats, selection_mode="multi")

    date_min = df_all["transaction_date"].min().date()
    date_max = df_all["transaction_date"].max().date()
    date_range = st.date_input("Date range", (date_min, date_max),
                               min_value=date_min, max_value=date_max)

    horizon = st.slider("Forecast horizon (days)", 1, 30, 7)
    metric_mode = st.radio("Forecast metric", ["Revenue", "Quantity"], horizontal=True)

    st.markdown("---")
    st.caption("Models: Moving Avg · Exp. Smoothing · Linear Trend · Seasonal Naive")

# Apply filters
mask = (
    df_all["store_location"].isin(loc_sel or locs) &
    df_all["product_category"].isin(cat_sel or cats)
)
if isinstance(date_range, tuple) and len(date_range) == 2:
    mask &= (df_all["transaction_date"].dt.date >= date_range[0]) & \
            (df_all["transaction_date"].dt.date <= date_range[1])
df = df_all[mask].copy()

# -------- Hero --------
_hero_img = Path(__file__).parent / "assets" / "coffee_hero.jpg"
_img_b64 = ""
if _hero_img.exists():
    _img_b64 = base64.b64encode(_hero_img.read_bytes()).decode()

st.markdown(
    f"""
<div class="hero" style="background:
    linear-gradient(135deg, rgba(26,15,12,.92) 0%, rgba(62,39,35,.85) 50%, rgba(111,78,55,.75) 100%),
    url('data:image/jpeg;base64,{_img_b64}'); background-size: cover; background-position: center;">
  <div class="emoji">☕</div>
  <div>
    <h1>Afficionado Coffee Roasters</h1>
    <p>Data-driven forecasting & peak demand prediction · {len(df):,} transactions in view</p>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

# -------- KPI row --------
total_rev = df["revenue"].sum()
total_qty = df["transaction_qty"].sum()
n_tx = len(df)
avg_ticket = total_rev / max(n_tx, 1)
n_days_view = df["transaction_date"].nunique()
daily_rev = total_rev / max(n_days_view, 1)

cols = st.columns(5)
for c, (label, value) in zip(cols, [
    ("Total Revenue", f"${total_rev:,.0f}"),
    ("Items Sold",    f"{total_qty:,}"),
    ("Transactions",  f"{n_tx:,}"),
    ("Avg Ticket",    f"${avg_ticket:,.2f}"),
    ("Avg Daily Rev", f"${daily_rev:,.0f}"),
]):
    c.markdown(f"<div class='metric-card'><div class='label'>{label}</div><div class='value'>{value}</div></div>",
               unsafe_allow_html=True)

st.write("")

# -------- Helper: download buttons --------
def csv_dl(df_out: pd.DataFrame, name: str, label: str = "⬇ Download data (CSV)"):
    st.download_button(label, df_out.to_csv(index=False).encode("utf-8"),
                       f"{name}.csv", "text/csv", key=f"dl_{name}")

# -------- Tabs --------
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📈 Forecast", "🔥 Peak Demand", "🏪 Stores", "📦 Products", "🗂 Data", "📥 Reports"
])

target_col = "revenue" if metric_mode == "Revenue" else "transaction_qty"
target_label = "Revenue ($)" if metric_mode == "Revenue" else "Quantity"

# ===== Tab 1: Forecast =====
with tab1:
    st.subheader("📈 Daily sales — history & forecast")
    st.caption("Solid brown line = past sales. Orange line = predicted next days. Shaded band = uncertainty range.")

    daily = df.groupby("transaction_date")[target_col].sum().sort_index()
    if len(daily) < 14:
        st.warning("Need at least 14 days of data for forecasting. Widen filters.")
    else:
        last_idx = daily.index
        future_idx = pd.date_range(last_idx[-1] + pd.Timedelta(days=1), periods=horizon, freq="D")

        ma7 = np.repeat(daily.tail(7).mean(), horizon)
        alpha = 0.3
        level = daily.iloc[0]
        for v in daily.iloc[1:]:
            level = alpha * v + (1 - alpha) * level
        es = np.repeat(level, horizon)
        x = np.arange(len(daily))
        a, b = np.polyfit(x, daily.values, 1)
        lin = a * np.arange(len(daily), len(daily) + horizon) + b
        dow_means = daily.groupby(daily.index.dayofweek).mean()
        seasonal = np.array([dow_means.loc[d.dayofweek] for d in future_idx])

        ensemble = np.mean([ma7, es, lin, seasonal], axis=0)
        residuals = (daily - daily.rolling(7, min_periods=1).mean()).dropna()
        sigma = residuals.std()
        upper = ensemble + 1.96 * sigma
        lower = np.maximum(0, ensemble - 1.96 * sigma)

        # SIMPLIFIED CHART: only History + Forecast + Confidence band
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=future_idx, y=upper, name="Upper bound",
                                 line=dict(color="rgba(0,0,0,0)"), showlegend=False, hoverinfo="skip"))
        fig.add_trace(go.Scatter(x=future_idx, y=lower, name="Uncertainty range",
                                 fill="tonexty", fillcolor="rgba(198,134,66,0.25)",
                                 line=dict(color="rgba(0,0,0,0)")))
        fig.add_trace(go.Scatter(x=daily.index, y=daily.values, name="Past sales (history)",
                                 line=dict(color=ESPRESSO, width=3)))
        fig.add_trace(go.Scatter(x=future_idx, y=ensemble, name="Predicted sales",
                                 line=dict(color=GOLD, width=4)))
        # connector
        fig.add_trace(go.Scatter(x=[daily.index[-1], future_idx[0]],
                                 y=[daily.values[-1], ensemble[0]],
                                 line=dict(color=GOLD, width=2, dash="dot"),
                                 showlegend=False, hoverinfo="skip"))
        style_fig(fig, height=460)
        fig.update_yaxes(title_text=target_label)
        fig.update_xaxes(title_text="Date")
        st.plotly_chart(fig, width="stretch")

        # Forecast summary
        sm_cols = st.columns(3)
        sm_cols[0].markdown(f"<div class='metric-card'><div class='label'>Forecast total ({horizon}d)</div>"
                            f"<div class='value'>{('$' if metric_mode=='Revenue' else '')}{ensemble.sum():,.0f}</div></div>",
                            unsafe_allow_html=True)
        sm_cols[1].markdown(f"<div class='metric-card'><div class='label'>Avg per day</div>"
                            f"<div class='value'>{('$' if metric_mode=='Revenue' else '')}{ensemble.mean():,.0f}</div></div>",
                            unsafe_allow_html=True)
        peak_day = future_idx[int(np.argmax(ensemble))].strftime("%a %b %d")
        sm_cols[2].markdown(f"<div class='metric-card'><div class='label'>Busiest day predicted</div>"
                            f"<div class='value' style='font-size:22px'>{peak_day}</div></div>",
                            unsafe_allow_html=True)

        st.markdown(f"<div class='insight'>💡 <b>Insight:</b> Expected total of "
                    f"{('$' if metric_mode=='Revenue' else '')}{ensemble.sum():,.0f} over the next {horizon} days. "
                    f"Plan stock & staffing peaks for <b>{peak_day}</b>.</div>", unsafe_allow_html=True)

        # Backtest table — simpler
        st.markdown("#### How accurate is the forecast?")
        st.caption("Lower MAPE % = more accurate. Tested on the most recent 14 days.")
        bt_n = min(14, len(daily) // 4)
        train, test = daily.iloc[:-bt_n], daily.iloc[-bt_n:]
        def _eval(pred):
            err = test.values - pred
            mae = np.mean(np.abs(err))
            mape = np.mean(np.abs(err) / np.maximum(test.values, 1)) * 100
            return mae, mape
        models_bt = {
            "Moving Avg (7d)":  np.repeat(train.tail(7).mean(), bt_n),
            "Exp. Smoothing":   np.repeat(train.ewm(alpha=0.3).mean().iloc[-1], bt_n),
            "Linear Trend":     np.polyval(np.polyfit(np.arange(len(train)), train.values, 1),
                                           np.arange(len(train), len(train) + bt_n)),
            "Seasonal Naive":   np.array([train.groupby(train.index.dayofweek).mean()
                                          .loc[d.dayofweek] for d in test.index]),
        }
        bt_df = pd.DataFrame([{"Model": k, "Avg Error (MAE)": round(_eval(v)[0],1),
                               "Accuracy %": round(100 - _eval(v)[1], 2)}
                              for k, v in models_bt.items()])
        st.dataframe(bt_df, width="stretch", hide_index=True)

        forecast_df = pd.DataFrame({
            "date": future_idx, "forecast": ensemble.round(2),
            "lower_bound": lower.round(2), "upper_bound": upper.round(2)
        })
        csv_dl(forecast_df, "forecast_report")

# ===== Tab 2: Peak Demand =====
with tab2:
    st.subheader("🔥 When are the rushes?")
    st.caption("Darker = busier. Quickly spot which day & hour drive demand.")

    heat = df.groupby(["dow", "hour"])[target_col].sum().reset_index()
    dow_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    heat_p = heat.pivot(index="dow", columns="hour", values=target_col).reindex(dow_order).fillna(0)
    fig = px.imshow(heat_p,
                    color_continuous_scale=[[0, "#231410"], [0.3, MOCHA], [0.6, CARAMEL], [1, GOLD]],
                    aspect="auto", labels=dict(color=target_label, x="Hour of day", y=""))
    style_fig(fig, height=380, legend=False)
    st.plotly_chart(fig, width="stretch")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Busiest hours")
        hourly = df.groupby("hour")[target_col].sum().reset_index()
        fig = px.bar(hourly, x="hour", y=target_col, color_discrete_sequence=[CARAMEL],
                     text_auto=".2s")
        fig.update_traces(textposition="outside", textfont=dict(color=INK, size=10))
        style_fig(fig, height=340, legend=False)
        fig.update_yaxes(title_text=target_label)
        fig.update_xaxes(title_text="Hour", dtick=1)
        st.plotly_chart(fig, width="stretch")
        top_hour = int(hourly.sort_values(target_col, ascending=False).iloc[0]["hour"])
        st.markdown(f"<div class='insight'>☕ Busiest hour: <b>{top_hour:02d}:00</b> — staff up before this window.</div>",
                    unsafe_allow_html=True)

    with c2:
        st.markdown("#### Busiest day of week")
        dow_d = df.groupby("dow")[target_col].sum().reindex(dow_order).reset_index()
        fig = px.bar(dow_d, x="dow", y=target_col, color_discrete_sequence=[MOCHA],
                     text_auto=".2s")
        fig.update_traces(textposition="outside", textfont=dict(color=INK, size=10))
        style_fig(fig, height=340, legend=False)
        fig.update_yaxes(title_text=target_label)
        fig.update_xaxes(title_text="")
        st.plotly_chart(fig, width="stretch")
        top_dow = dow_d.sort_values(target_col, ascending=False).iloc[0]["dow"]
        st.markdown(f"<div class='insight'>📅 Busiest day: <b>{top_dow}</b></div>", unsafe_allow_html=True)

    csv_dl(heat_p.reset_index(), "peak_demand_heatmap")

# ===== Tab 3: Stores =====
with tab3:
    st.subheader("🏪 Per-store performance")

    # Use weekly resample to make line clearer
    store_weekly = (df.groupby([pd.Grouper(key="transaction_date", freq="W"), "store_location"])[target_col]
                      .sum().reset_index())
    fig = px.line(store_weekly, x="transaction_date", y=target_col, color="store_location",
                  color_discrete_sequence=PALETTE, markers=True)
    fig.update_traces(line=dict(width=3))
    style_fig(fig, height=380)
    fig.update_yaxes(title_text=f"Weekly {target_label}")
    fig.update_xaxes(title_text="")
    st.plotly_chart(fig, width="stretch")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Revenue share by store")
        store_rev = df.groupby("store_location")["revenue"].sum().reset_index()
        fig = px.pie(store_rev, names="store_location", values="revenue", hole=0.55,
                     color_discrete_sequence=PALETTE)
        fig.update_traces(textfont=dict(color=INK, size=13, family="Georgia"),
                          textposition="outside", textinfo="label+percent")
        style_fig(fig, height=380, legend=False)
        st.plotly_chart(fig, width="stretch")
    with c2:
        st.markdown("#### Store totals")
        sr = store_rev.sort_values("revenue").copy()
        fig = px.bar(sr, x="revenue", y="store_location", orientation="h",
                     color_discrete_sequence=[CARAMEL], text_auto=".2s")
        fig.update_traces(textposition="outside", textfont=dict(color=INK))
        style_fig(fig, height=380, legend=False)
        fig.update_xaxes(title_text="Revenue ($)")
        fig.update_yaxes(title_text="")
        st.plotly_chart(fig, width="stretch")

    csv_dl(store_rev, "stores_report")

# ===== Tab 4: Products =====
with tab4:
    st.subheader("📦 Category & product mix")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Revenue by category")
        cat = df.groupby("product_category")["revenue"].sum().sort_values().reset_index()
        fig = px.bar(cat, x="revenue", y="product_category", orientation="h",
                     color_discrete_sequence=[MOCHA], text_auto=".2s")
        fig.update_traces(textposition="outside", textfont=dict(color=INK))
        style_fig(fig, height=420, legend=False)
        fig.update_xaxes(title_text="Revenue ($)")
        fig.update_yaxes(title_text="")
        st.plotly_chart(fig, width="stretch")
    with c2:
        st.markdown("#### Top 12 products")
        top = (df.groupby("product_type")["revenue"].sum()
               .sort_values(ascending=True).tail(12).reset_index())
        fig = px.bar(top, x="revenue", y="product_type", orientation="h",
                     color_discrete_sequence=[CARAMEL], text_auto=".2s")
        fig.update_traces(textposition="outside", textfont=dict(color=INK))
        style_fig(fig, height=420, legend=False)
        fig.update_xaxes(title_text="Revenue ($)")
        fig.update_yaxes(title_text="")
        st.plotly_chart(fig, width="stretch")

    st.markdown("#### Category trend (weekly)")
    cd = (df.groupby([pd.Grouper(key="transaction_date", freq="W"), "product_category"])["revenue"]
            .sum().reset_index())
    fig = px.area(cd, x="transaction_date", y="revenue", color="product_category",
                  color_discrete_sequence=PALETTE)
    style_fig(fig, height=380)
    fig.update_yaxes(title_text="Revenue ($)")
    fig.update_xaxes(title_text="")
    st.plotly_chart(fig, width="stretch")

    csv_dl(cat, "category_revenue")

# ===== Tab 5: Data =====
with tab5:
    st.subheader("🗂 Filtered transactions")
    st.dataframe(df.head(2000), width="stretch", height=460)
    csv_dl(df, "afficionado_filtered_transactions", "⬇ Download full filtered CSV")

# ===== Tab 6: Reports (downloads hub) =====
with tab6:
    st.subheader("📥 Download reports")
    st.caption("One-click exports for every section of the dashboard.")

    # Build report frames
    daily_full = df.groupby("transaction_date").agg(
        revenue=("revenue", "sum"),
        quantity=("transaction_qty", "sum"),
        transactions=("transaction_id", "count"),
    ).reset_index()

    hourly_full = df.groupby("hour").agg(
        revenue=("revenue", "sum"),
        quantity=("transaction_qty", "sum"),
    ).reset_index()

    store_full = df.groupby("store_location").agg(
        revenue=("revenue", "sum"),
        quantity=("transaction_qty", "sum"),
        transactions=("transaction_id", "count"),
    ).reset_index().sort_values("revenue", ascending=False)

    product_full = df.groupby(["product_category", "product_type"]).agg(
        revenue=("revenue", "sum"),
        quantity=("transaction_qty", "sum"),
    ).reset_index().sort_values("revenue", ascending=False)

    reports = [
        ("Daily sales summary", daily_full, "daily_sales_report"),
        ("Hourly demand summary", hourly_full, "hourly_demand_report"),
        ("Store performance", store_full, "store_performance_report"),
        ("Product performance", product_full, "product_performance_report"),
        ("Full filtered transactions", df, "full_transactions_report"),
    ]

    for title, frame, fname in reports:
        cA, cB = st.columns([3, 1])
        with cA:
            st.markdown(f"<div class='section-card'><b style='color:{ESPRESSO};font-size:15px'>📄 {title}</b>"
                        f"<br><span style='color:{MOCHA};font-size:12px'>{len(frame):,} rows</span></div>",
                        unsafe_allow_html=True)
        with cB:
            st.write("")
            st.download_button("⬇ CSV", frame.to_csv(index=False).encode("utf-8"),
                               f"{fname}.csv", "text/csv", key=f"rpt_{fname}")

    # Excel workbook with all reports
    try:
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xl:
            for title, frame, fname in reports:
                frame.to_excel(xl, sheet_name=title[:31], index=False)
        st.download_button("⬇ Download ALL reports as Excel workbook",
                           buf.getvalue(),
                           "afficionado_full_report.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.caption(f"(Excel export unavailable: {e})")

st.markdown(
    f"<div style='text-align:center;color:{MOCHA};padding:22px 0;font-style:italic;font-size:13px'>"
    f"Brewed with ☕ for Afficionado Coffee Roasters · Forecasting Intelligence Suite</div>",
    unsafe_allow_html=True,
)
