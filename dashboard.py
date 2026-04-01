# dashboard.py
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import numpy as np
from screener import run_screener, FUND_UNIVERSE, fetch_metrics, composite_score
from pitch_generator import build_pitch_slide
import os, io

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AllianzGI Multi-Asset Screener",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    .metric-label { font-size: 12px !important; }
</style>
""", unsafe_allow_html=True)

# ── SIDEBAR FILTERS ───────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Filters")
    fund_type   = st.selectbox("Asset class", ["All", "Balanced", "Defensive", "Growth", "Absolute Return"])
    min_sharpe  = st.slider("Min Sharpe ratio", 0.0, 1.5, 0.4, 0.1)
    min_esg     = st.slider("Min ESG score",    0,   100, 60,  5)
    max_dd      = st.slider("Max drawdown (%)", -40.0, 0.0, -25.0, 1.0)
    top_n       = st.slider("Generate pitch for top N funds", 1, 5, 3)
    run_btn     = st.button("Run screener", type="primary", use_container_width=True)

# ── SESSION STATE ─────────────────────────────────────────────────────────────
if "funds" not in st.session_state:
    st.session_state.funds = []
    st.session_state.auto_run_done = False

# ── RUN SCREENER ──────────────────────────────────────────────────────────────
if run_btn or not st.session_state.auto_run_done:
    with st.spinner("Fetching live data and screening funds..."):
        filters = {"min_sharpe": min_sharpe, "min_esg": min_esg, "max_drawdown": max_dd}
        funds = run_screener(filters)
        if fund_type != "All":
            funds = [f for f in funds if f["type"] == fund_type]
        st.session_state.funds = funds
        st.session_state.auto_run_done = True

funds = st.session_state.funds

# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("## AllianzGI Multi-Asset Fund Screener")
st.caption("Quantitative screening tool · Product Specialist team · Hong Kong")
st.divider()

# ── PORTFOLIO OVERVIEW METRICS ────────────────────────────────────────────────
if funds:
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Funds passing filters", len(funds), f"of {len(FUND_UNIVERSE)} screened")
    col2.metric("Avg Sharpe ratio",  round(np.mean([f["sharpe"] for f in funds]), 2))
    col3.metric("Avg ESG score",     round(np.mean([f["esg"] for f in funds]), 1))
    col4.metric("Top composite score", max(f["score"] for f in funds), "/ 100")
    st.divider()

# ── MAIN TABS ────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["Fund table", "Analytics", "Generate reports"])

# ── TAB 1: FUND TABLE ─────────────────────────────────────────────────────────
with tab1:
    if not funds:
        st.info("Set your filters in the sidebar and click 'Run screener' to begin.")
    else:
        df = pd.DataFrame(funds)[
            ["name","type","region","return_1y","sharpe","max_drawdown","esg","score"]
        ].rename(columns={
            "name":         "Fund name",
            "type":         "Type",
            "region":       "Region",
            "return_1y":    "1Y return (%)",
            "sharpe":       "Sharpe",
            "max_drawdown": "Max DD (%)",
            "esg":          "ESG score",
            "score":        "Composite score",
        })

        # Colour-map the score column
        st.dataframe(
            df.style
              .background_gradient(subset=["Composite score"], cmap="RdYlGn", vmin=0, vmax=100)
              .background_gradient(subset=["1Y return (%)"],   cmap="Greens", vmin=0)
              .background_gradient(subset=["ESG score"],       cmap="YlGn",   vmin=0, vmax=100)
              .format({"1Y return (%)": "{:.1f}%", "Max DD (%)": "{:.1f}%",
                       "Sharpe": "{:.2f}", "Composite score": "{:.0f}"}),
            use_container_width=True,
            height=350,
        )

        # Fund detail expander
        selected = st.selectbox("View fund detail", [f["name"] for f in funds])
        fund_detail = next(f for f in funds if f["name"] == selected)

        with st.expander(f"Detail: {selected}", expanded=True):
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("1Y return",    f"{fund_detail['return_1y']}%")
            c2.metric("Sharpe",       fund_detail["sharpe"])
            c3.metric("Max drawdown", f"{fund_detail['max_drawdown']}%")
            c4.metric("ESG score",    f"{fund_detail['esg']} / 100")
            c5.metric("Score",        f"{fund_detail['score']} / 100")

            # Allocation donut
            alloc = fund_detail["allocation"]
            fig_donut = go.Figure(go.Pie(
                labels=list(alloc.keys()),
                values=list(alloc.values()),
                hole=0.55,
                marker_colors=["#378ADD","#1D9E75","#D85A30","#888780"],
                textinfo="label+percent",
            ))
            fig_donut.update_layout(
                height=260, margin=dict(t=10,b=10,l=10,r=10),
                showlegend=False, paper_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(fig_donut, use_container_width=True)

# ── TAB 2: ANALYTICS ──────────────────────────────────────────────────────────
with tab2:
    if not funds:
        st.info("Run the screener first.")
    else:
        col_a, col_b = st.columns(2)

        with col_a:
            # Risk-return scatter
            df_plot = pd.DataFrame(funds)
            fig_scatter = px.scatter(
                df_plot,
                x="max_drawdown", y="return_1y",
                size="esg", color="score",
                color_continuous_scale="RdYlGn",
                range_color=[0, 100],
                hover_name="name",
                labels={"max_drawdown": "Max drawdown (%)",
                        "return_1y":    "1Y return (%)",
                        "score":        "Composite score"},
                title="Risk vs return (bubble size = ESG score)",
            )
            fig_scatter.update_layout(
                height=320, paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0.02)"
            )
            st.plotly_chart(fig_scatter, use_container_width=True)

        with col_b:
            # Composite score bar chart
            df_sorted = pd.DataFrame(funds).sort_values("score", ascending=True)
            fig_bar = px.bar(
                df_sorted,
                x="score", y="name",
                orientation="h",
                color="score",
                color_continuous_scale="RdYlGn",
                range_color=[0,100],
                labels={"score": "Composite score", "name": ""},
                title="Composite score ranking",
            )
            fig_bar.update_layout(
                height=320, paper_bgcolor="rgba(0,0,0,0)",
                coloraxis_showscale=False, showlegend=False
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        # ESG vs Sharpe
        fig_esg = px.scatter(
            pd.DataFrame(funds),
            x="esg", y="sharpe",
            color="type",
            hover_name="name",
            size="return_1y",
            labels={"esg":"ESG score","sharpe":"Sharpe ratio","type":"Fund type"},
            title="ESG score vs Sharpe ratio (bubble size = 1Y return)",
        )
        fig_esg.update_layout(height=300, paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_esg, use_container_width=True)

# ── TAB 3: GENERATE REPORTS ───────────────────────────────────────────────────
with tab3:
    if not funds:
        st.info("Run the screener first.")
    else:
        st.markdown(f"**Generating pitch one-pagers for the top {top_n} funds by composite score.**")

        for fund in funds[:top_n]:
            col_info, col_btn = st.columns([4, 1])
            col_info.markdown(f"**{fund['name']}** · Score: {fund['score']}/100 · {fund['type']}")

            filename = f"{fund['name'].replace(' ','_')}_Pitch.pptx"
            filepath = os.path.join("output", filename)
            os.makedirs("output", exist_ok=True)
            build_pitch_slide(fund, filepath)

            with open(filepath, "rb") as f:
                col_btn.download_button(
                    label="Download",
                    data=f,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key=fund["name"],
                )

        st.success(f"{min(top_n, len(funds))} pitch one-pagers ready for download.")