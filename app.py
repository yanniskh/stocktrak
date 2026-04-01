import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import os

# ─────────────────────────────────────────────────────────────────────────────
# CHEMIN VERS VOTRE FICHIER EXCEL
# ─────────────────────────────────────────────────────────────────────────────
EXCEL_PATH = r"\data\Professional_Trading_Journal.xlsx"
CSV_PATH   = r"\data\historicalportfoliovalues.csv"
EXPORT_PATH = r"\data\assetlist march.csv"
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Trading Journal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
:root {
    --bg:     #0b0e13;
    --panel:  #12181f;
    --border: #1c2633;
    --green:  #4cff72;
    --red:    #ff4444;
    --cyan:   #00cfff;
    --orange: #ffaa00;
    --white:  #e8f0f8;
    --muted:  #4a6680;
    --muted2: #7a9ab8;
}
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui, sans-serif;
    background-color: var(--bg) !important;
    color: var(--white) !important;
}
.stApp { background: var(--bg) !important; }
.main .block-container { padding: 0.4rem 0.8rem 2rem 0.8rem; max-width: 100%; }
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }

.top-bar {
    background: #0d1420; border: 1px solid var(--border);
    border-radius: 4px; padding: 7px 16px;
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 8px;
}
.logo-icon {
    background: var(--orange); width: 32px; height: 32px;
    border-radius: 3px; display: flex; align-items: center;
    justify-content: center; font-weight: 900; font-size: 13px;
    color: #000; margin-right: 10px;
}
.app-title { font-size: 13px; font-weight: 700; color: var(--white); letter-spacing: 1px; }
.app-sub   { font-size: 10px; color: var(--muted); letter-spacing: 2px; text-transform: uppercase; }

.kpi-bar { background: #0d1420; border: 1px solid var(--border); border-radius: 4px; display: flex; margin-bottom: 8px; }
.kpi-cell { flex: 1; padding: 8px 14px; border-right: 1px solid var(--border); }
.kpi-cell:last-child { border-right: none; }
.kpi-lbl { font-size: 9px; color: var(--muted); text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 2px; }
.kpi-val { font-size: 15px; font-weight: 700; font-variant-numeric: tabular-nums; }
.c-green  { color: var(--green); }
.c-red    { color: var(--red); }
.c-cyan   { color: var(--cyan); }
.c-white  { color: var(--white); }
.c-orange { color: var(--orange); }
.c-muted2 { color: var(--muted2); }

.sec-title {
    text-align: center; font-size: 10px; font-weight: 600;
    letter-spacing: 5px; color: var(--cyan);
    border-top: 1px solid var(--border); border-bottom: 1px solid var(--border);
    padding: 4px 0; margin: 6px 0; text-transform: uppercase;
}
.pnl {
    background: var(--panel); border: 1px solid var(--border);
    border-radius: 4px; padding: 10px 12px; margin-bottom: 8px;
}
.pnl-title {
    font-size: 9px; color: var(--muted); text-transform: uppercase;
    letter-spacing: 1.5px; padding-bottom: 5px; margin-bottom: 6px;
    border-bottom: 1px solid var(--border);
}
.sr { display: flex; justify-content: space-between; align-items: center; padding: 3px 0; border-bottom: 1px solid #161e29; }
.sr:last-child { border-bottom: none; }
.sr-lbl { font-size: 11px; color: var(--muted2); }
.sr-g   { font-size: 12px; font-weight: 600; color: var(--green); font-variant-numeric: tabular-nums; }
.sr-r   { font-size: 12px; font-weight: 600; color: var(--red);   font-variant-numeric: tabular-nums; }
.sr-w   { font-size: 12px; font-weight: 600; color: var(--white); font-variant-numeric: tabular-nums; }
.sr-o   { font-size: 12px; font-weight: 600; color: var(--orange); font-variant-numeric: tabular-nums; }

.ts-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 6px; margin-top: 8px; }
.ts-cell { background: #0a0f16; border-radius: 3px; padding: 5px 4px; text-align: center; }
.ts-val  { font-size: 13px; font-weight: 700; font-variant-numeric: tabular-nums; }
.ts-lbl  { font-size: 8px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; margin-top: 1px; }

.trader-yannis { color: #00cfff; font-weight: 600; }
.trader-jade   { color: #ff88cc; font-weight: 600; }
.trader-wang   { color: #ffaa00; font-weight: 600; }
.trader-wanchun{ color: #88ffcc; font-weight: 600; }

.stTabs [data-baseweb="tab-list"] {
    background: #0a0e14 !important; border-bottom: 1px solid var(--border) !important;
    gap: 0 !important; padding: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    font-size: 11px !important; font-weight: 600 !important; color: var(--muted) !important;
    background: transparent !important; border: none !important;
    padding: 8px 14px !important; text-transform: uppercase; letter-spacing: 0.5px;
}
.stTabs [aria-selected="true"] {
    color: var(--cyan) !important; background: #0f1822 !important;
    border-bottom: 2px solid var(--cyan) !important;
}
.stTabs [data-baseweb="tab-panel"] { background: var(--bg) !important; padding: 8px 0 !important; }

[data-testid="stDataFrame"] { background: var(--panel) !important; }
::-webkit-scrollbar { width: 3px; height: 3px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 2px; }
div[data-testid="stFileUploadDropzone"] {
    background: var(--panel) !important; border: 1px dashed var(--cyan) !important; border-radius: 4px !important;
}
</style>
""", unsafe_allow_html=True)

SYS_FONT = "-apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif"

# ─────────────────────────────────────────────────────────────────────────────
# LECTURE EXCEL
# ─────────────────────────────────────────────────────────────────────────────
def load_excel(path):
    try:
        xl = pd.ExcelFile(path)
        tj = xl.parse("Trading_Journal")
        tj.columns = [str(c).strip() for c in tj.columns]
        tj = tj[tj["Trade ID"].notna()].copy()
        tj["Trade date"]         = pd.to_datetime(tj["Trade date"], errors="coerce")
        tj["Amount($)"]          = pd.to_numeric(tj["Amount($)"],          errors="coerce")
        tj["Amount(€)"]          = pd.to_numeric(tj["Amount(€)"],          errors="coerce")
        tj["Risk (€)"]           = pd.to_numeric(tj["Risk (€)"],           errors="coerce")
        tj["% Portfolio Risked"] = pd.to_numeric(tj["% Portfolio Risked"], errors="coerce")
        tj["Price paid"]         = pd.to_numeric(tj["Price paid"],         errors="coerce")
        tj["Qty"]                = pd.to_numeric(tj["Qty"],                errors="coerce")

        db_raw = xl.parse("Dashboard")
        db = db_raw.copy().reset_index(drop=True)
        db.columns = ["Full Trade cycle", "Date", "Instrument", "Asset class", "Trader",
                      "Net P&L", "Net P&L $", "Portfolio wallet currently",
                      "Portfolio $", "Col10", "Col11"]
        db["Full Trade cycle"]           = pd.to_numeric(db["Full Trade cycle"],           errors="coerce")
        db["Net P&L"]                    = pd.to_numeric(db["Net P&L"],                    errors="coerce")
        db["Net P&L $"]                  = pd.to_numeric(db["Net P&L $"],                  errors="coerce")
        db["Portfolio wallet currently"] = pd.to_numeric(db["Portfolio wallet currently"], errors="coerce")
        db = db[db["Full Trade cycle"].notna() & db["Net P&L"].notna()].copy().reset_index(drop=True)
        db["Date"] = pd.to_datetime(db["Date"], format="%d/%m/%Y %H:%M", errors="coerce")
        initial = 1_000_000.0
        return tj, db, initial, None
    except FileNotFoundError:
        return None, None, None, f"Fichier introuvable :\n{path}"
    except Exception as e:
        return None, None, None, f"Erreur lors de la lecture : {e}"


def load_csv(path):
    h = pd.read_csv(path, sep=";", decimal=",")
    h.columns = [c.strip() for c in h.columns]
    h["Date"] = pd.to_datetime(h["Date"].str.strip(), format="%m/%d/%Y", errors="coerce")
    h = h.dropna(subset=["Date"]).sort_values("Date").reset_index(drop=True)
    for col in [h.columns[3], h.columns[4]]:
        h[col] = h[col].astype(str).str.replace("%","").str.replace(",",".").str.strip()
        h[col] = pd.to_numeric(h[col], errors="coerce").fillna(0)
    return h


tj, db, initial, load_error = load_excel(EXCEL_PATH)

# ─────────────────────────────────────────────────────────────────────────────
# UTILITAIRES
# ─────────────────────────────────────────────────────────────────────────────
def bar_colors(series):
    return ["#4cff72" if v >= 0 else "#ff4444" for v in series]

def gauge(val, color):
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=val,
        number=dict(suffix="%", font=dict(size=18, color=color, family=SYS_FONT)),
        gauge=dict(
            axis=dict(range=[0, 100], visible=False),
            bar=dict(color=color, thickness=0.28),
            bgcolor="#0b0e13", bordercolor="#1c2633",
            steps=[dict(range=[0, 100], color="#0d1218")],
            threshold=dict(line=dict(color=color, width=2), thickness=0.6, value=val)
        )
    ))
    fig.update_layout(paper_bgcolor="#12181f", height=150, margin=dict(l=0, r=0, t=10, b=0))
    return fig

# ─────────────────────────────────────────────────────────────────────────────
# BARRE DU HAUT
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="top-bar">
  <div style="display:flex;align-items:center;">
    <div class="logo-icon">TJ</div>
    <div>
      <div class="app-title">Professional Trading Journal</div>
      <div class="app-sub">STJ — Ver. 1.0</div>
    </div>
  </div>
  <div style="font-size:11px;color:#4a6680;">{datetime.now().strftime("%Y-%m-%d  %H:%M")}</div>
</div>
""", unsafe_allow_html=True)

if load_error:
    st.error(f"❌  {load_error}")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# MÉTRIQUES GLOBALES
# ─────────────────────────────────────────────────────────────────────────────
total_pnl    = -(1000000 - 924684.08)
net_pos      = db[db["Net P&L"] > 0]["Net P&L"].sum()
net_neg      = db[db["Net P&L"] < 0]["Net P&L"].sum()
n_wins       = int((db["Net P&L"] > 0).sum())
n_losses     = int((db["Net P&L"] < 0).sum())
n_open       = 115
win_rate     = (n_wins / n_open) * 100 if n_open else 0
current_wallet = db["Portfolio wallet currently"].iloc[-1] if n_open > 0 else initial
avg_win      = db[db["Net P&L"] > 0]["Net P&L"].mean() if n_wins   else 0
avg_loss     = abs(db[db["Net P&L"] < 0]["Net P&L"].mean()) if n_losses else 0
total_return = (((current_wallet - initial) / initial) * 100) if initial else 0

# ─────────────────────────────────────────────────────────────────────────────
# ONGLETS
# ─────────────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "📊 Dashboard",
    "📋 All trades",
    "📈 Analysis",
    "📉 Long-Short Strategy",
    "🔍 Bonds Strategy",
])

# ═══════════════════════════════════════════════════════════════════════════
# TAB 0 — DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════
with tabs[0]:
    pnl_cls = "c-green" if total_pnl >= 0 else "c-red"
    ret_cls = "c-green" if total_return >= 0 else "c-red"

    st.markdown(f"""
    <div class="kpi-bar">
      <div class="kpi-cell"><div class="kpi-lbl">Initial Portfolio</div><div class="kpi-val c-white">€{initial:,.0f}</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Net P&amp;L Total</div><div class="kpi-val {pnl_cls}">€{total_pnl:+,.2f}</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Portfolio Actual</div><div class="kpi-val c-white">€924,684.08</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Total Return</div><div class="kpi-val {ret_cls}">-7.53%</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Total Trades</div><div class="kpi-val c-cyan">232</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="sec-title">OVERALL SUMMARY</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1.7, 1.3])

    with c1:
        net_cls2 = "sr-g" if total_pnl >= 0 else "sr-r"
        top3 = db.nlargest(3, "Net P&L").set_index("Instrument")["Net P&L"]
        bot3 = db.nsmallest(3, "Net P&L").set_index("Instrument")["Net P&L"]
        max_abs = db["Net P&L"].abs().max()

        st.markdown(f"""
        <div class="pnl">
        <div class="pnl-title">OVERALL PERFORMANCE</div>
        <div class="sr"><span class="sr-lbl">Total Gains</span><span class="sr-g">€{net_pos:+,.2f}</span></div>
        <div class="sr"><span class="sr-lbl">Total Losses</span><span class="sr-r">€{net_neg:+,.2f}</span></div>
        <div class="sr"><span class="sr-lbl">Net P&amp;L</span><span class="{net_cls2}">€{total_pnl:+,.2f}</span></div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="pnl-title" style="margin-top:10px;">🏆 TOP 3 — BEST SYMBOLS</div>', unsafe_allow_html=True)
        for symbol, pnl in top3.items():
            pct = (pnl / initial) * 100
            bar_w = (abs(pnl) / max_abs) * 100 if max_abs else 0
            st.markdown(f"""
            <div style="margin-bottom:6px;">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:3px;">
                <span style="font-size:11px;color:#7a9ab8;font-weight:600;">{symbol}</span>
                <span style="font-size:12px;font-weight:700;color:#4cff72;">€{pnl:+,.2f} ({pct:+.2f}%)</span>
            </div>
            <div style="background:#0a0f16;border-radius:2px;height:5px;overflow:hidden;">
                <div style="width:{bar_w:.1f}%;height:100%;background:linear-gradient(90deg,#2aff5e,#4cff72);border-radius:2px;"></div>
            </div></div>""", unsafe_allow_html=True)

        st.markdown('<div class="pnl-title" style="margin-top:10px;">⚠️ TOP 3 — WORST SYMBOLS</div>', unsafe_allow_html=True)
        for symbol, pnl in bot3.items():
            pct = (pnl / initial) * 100
            bar_w = (abs(pnl) / max_abs) * 100 if max_abs else 0
            st.markdown(f"""
            <div style="margin-bottom:6px;">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:3px;">
                <span style="font-size:11px;color:#7a9ab8;font-weight:600;">{symbol}</span>
                <span style="font-size:12px;font-weight:700;color:#ff4444;">€{pnl:+,.2f} ({pct:+.2f}%)</span>
            </div>
            <div style="background:#0a0f16;border-radius:2px;height:5px;overflow:hidden;">
                <div style="width:{bar_w:.1f}%;height:100%;background:linear-gradient(90deg,#ff2222,#ff4444);border-radius:2px;"></div>
            </div></div>""", unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)

        now = datetime.now()
        db["month_num"] = db["Date"].dt.month
        db["year"] = db["Date"].dt.year
        current_year_db = db[db["year"] == now.year]
        monthly_pnl = current_year_db.groupby("month_num")["Net P&L"].sum()
        jan_pnl = monthly_pnl.get(1, 0)
        feb_pnl = monthly_pnl.get(2, 0)
        mar_pnl = monthly_pnl.get(3, 0)

        def period_row(label, value):
            cls = "sr-g" if value >= 0 else "sr-r"
            pct = (value / initial) * 100 if initial else 0
            arrow = "▲" if value >= 0 else "▼"
            return (f'<div class="sr"><span class="sr-lbl">{label}</span>'
                    f'<span class="{cls}">€{value:+,.2f} {arrow} {abs(pct):.2f}%</span></div>')

        st.markdown(
            '<div class="pnl"><div class="sec-title" style="margin-bottom:8px;">PERFORMANCE</div>'
            + period_row("January", jan_pnl)
            + period_row("February", feb_pnl)
            + period_row("March", mar_pnl)
            + f'<div style="font-size:8px;color:#4a6680;margin-top:6px;">Report as of: {now.strftime("%d-%m-%Y")}</div></div>',
            unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="pnl"><div class="pnl-title">DISTRIBUTION OF GAINS AND LOSSES</div>', unsafe_allow_html=True)
        db["pct_pnl"] = (db["Net P&L"] / initial) * 100
        bins = list(range(-20, 22, 2))
        labels_bins = [f"{bins[i]}%/{bins[i+1]}%" for i in range(len(bins)-1)]
        db["pnl_bin"] = pd.cut(db["pct_pnl"], bins=bins, labels=labels_bins, include_lowest=True)
        dist = pd.DataFrame({"pnl_bin": labels_bins})
        counts = db.groupby("pnl_bin", observed=True).size().reset_index(name="count")
        dist = dist.merge(counts, on="pnl_bin", how="left").fillna(0)
        bin_colors = ["#4cff72" if float(str(b).split("%/")[0]) >= 0 else "#ff4444" for b in dist["pnl_bin"]]
        fig_dist = go.Figure()
        fig_dist.add_trace(go.Bar(x=dist["pnl_bin"], y=dist["count"], marker_color=bin_colors,
            hovertemplate="%{x}<br>%{y} trades<extra></extra>"))
        fig_dist.update_layout(
            paper_bgcolor="#12181f", plot_bgcolor="#0b0e13",
            font=dict(family=SYS_FONT, color="#4a6680", size=10),
            margin=dict(l=32, r=8, t=8, b=48), height=200, bargap=0.1,
            xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=8, color="#4a6680"), tickangle=-45),
            yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"),
                      range=[0, max(dist["count"].max() + 3, 20)],
                      title=dict(text="NUMBER OF TRADES", font=dict(size=8, color="#4a6680"))))
        st.plotly_chart(fig_dist, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="pnl"><div class="pnl-title">PERFORMANCE CURVE</div>', unsafe_allow_html=True)
        wallet_vals = [initial] + db["Portfolio wallet currently"].tolist()
        n_trades = len(wallet_vals)
        x_axis = list(range(n_trades))
        fig_curve = go.Figure()
        fig_curve.add_trace(go.Scatter(x=x_axis, y=wallet_vals, mode="none", fill="tozeroy",
            fillcolor="rgba(0,207,255,0.04)", showlegend=False, hoverinfo="skip"))
        fig_curve.add_trace(go.Scatter(x=x_axis, y=wallet_vals, mode="lines",
            line=dict(color="#00cfff", width=2), name="Performance",
            hovertemplate="Trade %{x}<br>€%{y:,.0f}<extra></extra>"))
        total_ret = ((wallet_vals[-1] - initial) / initial) * 100
        fig_curve.add_annotation(x=0.98, y=0.95, xref="paper", yref="paper",
            text=f"Total Return: <b><span style='color:#ff4444'>{total_ret:+.1f}%</span></b>",
            showarrow=False, font=dict(size=11, color="#7a9ab8", family=SYS_FONT), align="right")
        fig_curve.add_annotation(x=0.5, y=0.95, xref="paper", yref="paper",
            text=f"{n_trades - 1} Round trips", showarrow=False,
            font=dict(size=10, color="#4a6680", family=SYS_FONT), align="right")
        fig_curve.update_layout(
            paper_bgcolor="#12181f", plot_bgcolor="#0b0e13",
            font=dict(family=SYS_FONT, color="#4a6680", size=10),
            margin=dict(l=60, r=16, t=16, b=40), height=220, showlegend=True,
            legend=dict(orientation="h", x=0.01, y=0.01, font=dict(size=9, color="#7a9ab8"),
                       bgcolor="rgba(0,0,0,0)", borderwidth=0),
            xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"),
                      tickmode="linear", dtick=max(1, n_trades // 6)),
            yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickformat="€,.0f",
                      tickfont=dict(size=9, color="#4a6680"), range=[850000, 1200000]))
        st.plotly_chart(fig_curve, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    with c3:
        if len(db) > 1:
            returns = db["Net P&L"] / initial
        else:
            returns = pd.Series([0])

        st.markdown('<div class="pnl"><div class="pnl-title">TRADE STATISTICS</div>', unsafe_allow_html=True)
        st.plotly_chart(gauge(win_rate, "#ffaa00"), use_container_width=True, config={"displayModeBar": False})
        st.markdown(f'<div style="text-align:center;font-size:9px;color:#4a6680;margin-top:-12px;">'
                    f'<span style="color:#ffaa00;font-weight:700;">{win_rate:.1f}%</span> Win Rate</div>',
                    unsafe_allow_html=True)

        st.markdown(f"""
        <div class="ts-grid" style="grid-template-columns:repeat(3,1fr);margin-top:6px;">
        <div class="ts-cell"><div class="ts-val c-green">{n_wins}</div><div class="ts-lbl">Wins</div></div>
        <div class="ts-cell"><div class="ts-val c-red">{n_losses}</div><div class="ts-lbl">Losses</div></div>
        <div class="ts-cell"><div class="ts-val c-cyan">{len(db)}</div><div class="ts-lbl">Clôturés</div></div>
        </div>""", unsafe_allow_html=True)

        st.markdown(f"""
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:6px;margin-top:8px;">
        <div class="ts-cell"><div class="ts-val c-green">€{avg_win:,.0f}</div><div class="ts-lbl">Avg Win</div></div>
        <div class="ts-cell"><div class="ts-val c-red">€{avg_loss:,.0f}</div><div class="ts-lbl">Avg Loss</div></div>
        <div class="ts-cell"><div class="ts-val c-cyan">-4.18</div><div class="ts-lbl">Sharpe</div></div>
        </div>""", unsafe_allow_html=True)

        st.markdown('<div class="pnl"><div class="pnl-title">WEEKLY PERFORMANCE CURVE</div>', unsafe_allow_html=True)
        db["week"] = db["Date"].dt.isocalendar().week.astype(int)
        db["year_w"] = db["Date"].dt.year
        weekly = db.groupby(["year_w", "week"])["Net P&L"].sum().reset_index()
        weekly["label"] = "W" + weekly["week"].astype(str)
        weekly["cumul"] = weekly["Net P&L"].cumsum()

        available_weeks = weekly["label"].tolist()
        now = datetime.now()
        current_week = now.isocalendar()[1]
        current_label = f"W{current_week}"
        default_idx = available_weeks.index(current_label) if current_label in available_weeks else len(available_weeks) - 1

        selected_week = st.selectbox("Semaine", available_weeks, index=default_idx, key="week_selector")
        selected_idx = available_weeks.index(selected_week)
        weekly_filtered = weekly.iloc[:selected_idx + 1].copy()
        weekly_filtered["cumul"] = weekly_filtered["Net P&L"].cumsum()

        sel_week_pnl  = weekly_filtered.iloc[-1]["Net P&L"]
        prev_week_pnl = weekly_filtered.iloc[-2]["Net P&L"] if len(weekly_filtered) >= 2 else 0
        sel_week_pct  = (sel_week_pnl / initial) * 100
        prev_week_pct = (prev_week_pnl / initial) * 100
        tw_cls = "c-green" if sel_week_pnl >= 0 else "c-red"
        pw_cls = "c-green" if prev_week_pnl >= 0 else "c-red"

        st.markdown(f"""
        <div style="display:flex;gap:24px;margin-bottom:6px;font-size:11px;">
        <div><span style="color:#4a6680;">Selected Week:</span>
        <span class="{tw_cls}" style="margin-left:8px;font-weight:700;">€{sel_week_pnl:+,.0f}</span>
        <span class="{tw_cls}" style="margin-left:6px;">{sel_week_pct:+.2f}%</span></div>
        <div><span style="color:#4a6680;">Prev. Week:</span>
        <span class="{pw_cls}" style="margin-left:8px;font-weight:700;">€{prev_week_pnl:+,.0f}</span>
        <span class="{pw_cls}" style="margin-left:6px;">{prev_week_pct:+.2f}%</span></div>
        </div>""", unsafe_allow_html=True)

        smooth_weekly = pd.Series(weekly_filtered["cumul"].tolist()).rolling(window=2, min_periods=1).mean().tolist()
        fig_week = go.Figure()
        fig_week.add_trace(go.Scatter(x=weekly_filtered["label"], y=weekly_filtered["cumul"],
            mode="none", fill="tozeroy", fillcolor="rgba(0,207,255,0.04)", showlegend=False, hoverinfo="skip"))
        fig_week.add_trace(go.Scatter(x=weekly_filtered["label"], y=smooth_weekly, mode="lines",
            line=dict(color="#00cfff", width=2, shape="spline", smoothing=1.2), name="Weekly Cumul",
            hovertemplate="%{x}<br>Cumul: €%{y:+,.0f}<extra></extra>"))
        fig_week.update_layout(
            paper_bgcolor="#12181f", plot_bgcolor="#0b0e13",
            font=dict(family=SYS_FONT, color="#4a6680", size=10),
            margin=dict(l=60, r=16, t=8, b=40), height=180, showlegend=False,
            xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"), tickangle=-30),
            yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickformat="€,.0f", tickfont=dict(size=9, color="#4a6680")))
        st.plotly_chart(fig_week, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 1 — TOUS LES TRADES
# ═══════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.markdown('<div class="sec-title">All trades - Trading_Journal</div>', unsafe_allow_html=True)
    top1, top2 = st.columns(2)

    with top1:
        st.markdown('<div class="pnl"><div class="pnl-title">EXPOSURE BY ASSET CLASS (Risk €)</div>', unsafe_allow_html=True)
        exp = tj[tj["Risk (€)"].notna()].groupby("Asset Class")["Risk (€)"].sum().reset_index()
        exp["Asset Class"] = exp["Asset Class"].replace({"Call": "Option", "Put": "Option"})
        exp = exp.groupby("Asset Class")["Risk (€)"].sum().reset_index()
        fig_exp = go.Figure(go.Pie(
            labels=exp["Asset Class"], values=exp["Risk (€)"], hole=0.5,
            marker=dict(colors=["#30acc8","#272f80","#bb8b2a","#b03030","#bf448a","#21ff9f","#7a9ab8","#a0522d"]),
            textfont=dict(size=12, color="#e8f0f8")))
        fig_exp.update_layout(paper_bgcolor="#12181f", height=250, margin=dict(l=10,r=10,t=10,b=10),
            legend=dict(font=dict(size=15, color="#7a9ab8"), orientation="v"))
        st.markdown('<div style="overflow:hidden;">', unsafe_allow_html=True)
        st.plotly_chart(fig_exp, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    with top2:
        st.markdown('<div class="pnl"><div class="pnl-title">WINNING TRADES — P&L PAR ASSET CLASS</div>', unsafe_allow_html=True)
        win_exp = db.groupby("Asset class")["Net P&L"].sum().reset_index()
        win_exp.columns = ["Asset Class", "Net P&L"]
        win_exp = win_exp[win_exp["Net P&L"] > 0].sort_values("Net P&L", ascending=False)
        fig_win = go.Figure(go.Pie(
            labels=win_exp["Asset Class"], values=win_exp["Net P&L"], hole=0.5,
            marker=dict(colors=["#272f80","#b03030","#bf448a","#30acc8","#21ff9f","#7a9ab8","#a0522d","#bb8b2a"]),
            textfont=dict(size=12, color="#e8f0f8")))
        fig_win.update_layout(paper_bgcolor="#12181f", height=250, margin=dict(l=10,r=10,t=10,b=10),
            legend=dict(font=dict(size=15, color="#7a9ab8"), orientation="v"))
        st.markdown('<div style="overflow:hidden;">', unsafe_allow_html=True)
        st.plotly_chart(fig_win, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        sel_t = st.selectbox("Trader", ["Tous"] + sorted(tj["Trader"].dropna().unique().tolist()))
    with fc2:
        sel_a = st.selectbox("Asset Class", ["Tous"] + sorted(tj["Asset Class"].dropna().unique().tolist()))
    with fc3:
        sel_d = st.selectbox("Direction", ["Tous"] + sorted(tj["Direction"].dropna().unique().tolist()))

    tf = tj.copy()
    if sel_t != "Tous": tf = tf[tf["Trader"] == sel_t]
    if sel_a != "Tous": tf = tf[tf["Asset Class"] == sel_a]
    if sel_d != "Tous": tf = tf[tf["Direction"] == sel_d]

    cols_show = ["Trade ID", "Trade date", "Trader", "Instrument", "Underlying asset",
                 "Asset Class", "Type", "Trading Session", "Option direction", "Direction",
                 "Price paid", "Qty", "Amount($)", "Amount(€)", "Stop Loss Set",
                 "Take Profit Set", "Commission + Fees", "Risk (€)", "% Portfolio Risked",
                 "Setup Type", "Strategy"]
    disp_tj = tf[cols_show].copy()
    disp_tj["Trade date"]         = disp_tj["Trade date"].dt.strftime("%Y-%m-%d %H:%M").fillna("—")
    disp_tj["Amount($)"]          = disp_tj["Amount($)"].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "—")
    disp_tj["Amount(€)"]          = disp_tj["Amount(€)"].apply(lambda x: f"€{x:,.2f}" if pd.notna(x) else "—")
    disp_tj["Risk (€)"]           = disp_tj["Risk (€)"].apply(lambda x: f"€{x:,.2f}" if pd.notna(x) else "—")
    disp_tj["% Portfolio Risked"] = disp_tj["% Portfolio Risked"].apply(lambda x: f"{x*100:.2f}%" if pd.notna(x) else "—")
    disp_tj["Price paid"]         = disp_tj["Price paid"].apply(lambda x: f"{x:,.4f}" if pd.notna(x) else "—")
    disp_tj["Qty"]                = disp_tj["Qty"].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "—")

    def highlight_direction(val):
        if str(val).strip().upper() == "BUY":  return "color: #4cff72; font-weight: bold"
        elif str(val).strip().upper() == "SELL": return "color: #ff4444; font-weight: bold"
        return ""

    styled = disp_tj.style.applymap(highlight_direction, subset=["Direction"])
    st.dataframe(styled, use_container_width=True, hide_index=True, height=480)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 2 — ANALYSE
# ═══════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.markdown('<div class="sec-title">PERFORMANCE & RISK ANALYTICS</div>', unsafe_allow_html=True)

    returns = db["Net P&L"] / initial
    sharpe           = -4.18
    Ourperformance   = -7.53
    Benchmarkperformance = -7.49

    # ── Chargement CSV + calculs volatilité/VaR ──────────────────────────
    hist_vol = load_csv(CSV_PATH)
    port_col_v   = hist_vol.columns[3]
    daily_returns = hist_vol[port_col_v] / 100
    n_days        = len(daily_returns)
    vol_daily     = daily_returns.std()
    vol_period    = vol_daily * np.sqrt(n_days) * 100
    vol_annual    = vol_daily * np.sqrt(252) * 100

    # VaR paramétrique : VaR = z × σ_daily × V
    V            = 924684.08
    z95, z99     = 1.65, 2.33
    var_95_param = z95 * vol_daily * V
    var_99_param = z99 * vol_daily * V

    # ── KPI Bar ──────────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="kpi-bar">
      <div class="kpi-cell"><div class="kpi-lbl">Max Drawdown</div><div class="kpi-val c-red">-12.60%</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Sharpe Ratio</div><div class="kpi-val c-cyan">{sharpe:.2f}</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Our Performance</div><div class="kpi-val c-cyan">{Ourperformance:.2f}%</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Benchmark (SP500)</div><div class="kpi-val c-orange">{Benchmarkperformance:.2f}%</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Total Commissions</div><div class="kpi-val c-red">€264.00</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Total Dividends</div><div class="kpi-val c-green">€235.79</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">VaR 95% </div><div class="kpi-val c-orange">€{var_95_param:,.0f}</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">VaR 99% </div><div class="kpi-val c-red">€{var_99_param:,.0f}</div></div>
    </div>
    """, unsafe_allow_html=True)

    # ── Row 1 : Drawdown + Portfolio vs SP500 ────────────────────────────
    r1a, r1b = st.columns(2)

    with r1a:
        st.markdown('<div class="pnl"><div class="pnl-title">DRAWDOWN CURVE (WITH CLOSED POSITIONS)</div>', unsafe_allow_html=True)
        wallet_series = pd.Series([initial] + db["Portfolio wallet currently"].tolist())
        drawdown_real = (wallet_series / initial - 1) * 100
        max_dd_real   = drawdown_real.min()
        fig_dd = go.Figure()
        fig_dd.add_trace(go.Scatter(x=list(range(len(drawdown_real))), y=drawdown_real.tolist(),
            mode="lines", fill="tozeroy", fillcolor="rgba(255,68,68,0.15)",
            line=dict(color="#ff4444", width=1.5),
            hovertemplate="Trade %{x}<br>Return: %{y:.2f}%<extra></extra>"))
        fig_dd.add_annotation(x=0.02, y=0.05, xref="paper", yref="paper",
            text=f"Max DD: <b>{max_dd_real:.2f}%</b>",
            showarrow=False, font=dict(size=11, color="#ff4444", family=SYS_FONT))
        fig_dd.update_layout(
            paper_bgcolor="#12181f", plot_bgcolor="#0b0e13",
            font=dict(family=SYS_FONT, color="#4a6680", size=10),
            margin=dict(l=50, r=16, t=8, b=32), height=220, showlegend=False,
            xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680")),
            yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"), ticksuffix="%"))
        st.plotly_chart(fig_dd, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    with r1b:
        st.markdown('<div class="pnl"><div class="pnl-title">PORTFOLIO VS S&P 500 — TOTAL RETURN (INCL. UNREALIZED P&L)</div>', unsafe_allow_html=True)
        hist = load_csv(CSV_PATH)
        port_col = hist.columns[3]
        spy_col  = hist.columns[4]
        hist["cumul_port"] = ((1 + hist[port_col] / 100).cumprod() - 1) * 100
        hist["cumul_spy"]  = ((1 + hist[spy_col]  / 100).cumprod() - 1) * 100
        fig_bench = go.Figure()
        fig_bench.add_hline(y=0, line_color="#2a2a3a", line_width=1.5)
        fig_bench.add_trace(go.Scatter(x=hist["Date"], y=hist["cumul_spy"], mode="lines",
            line=dict(color="#4040a0", width=2), name="SPY %",
            hovertemplate="%{x|%d %b}<br>SPY: %{y:.2f}%<extra></extra>"))
        fig_bench.add_trace(go.Scatter(x=hist["Date"], y=hist["cumul_port"], mode="lines",
            line=dict(color="#e060c0", width=2), name="Portfolio %",
            hovertemplate="%{x|%d %b}<br>Portfolio: %{y:.2f}%<extra></extra>"))
        fig_bench.update_layout(
            paper_bgcolor="#12181f", plot_bgcolor="#12181f",
            font=dict(family=SYS_FONT, color="#4a6680", size=10),
            margin=dict(l=60, r=16, t=8, b=50), height=220, showlegend=True,
            legend=dict(orientation="h", x=0.25, y=-0.3, font=dict(size=9, color="#7a9ab8"), bgcolor="rgba(0,0,0,0)"),
            xaxis=dict(gridcolor="#1e2535", linecolor="#1e2535", tickfont=dict(size=8, color="#4a6680"), tickformat="%d %b", tickangle=-45),
            yaxis=dict(gridcolor="#1e2535", linecolor="#1e2535", tickfont=dict(size=9, color="#4a6680"), ticksuffix="%",
                      title=dict(text="% return vs Benchmark", font=dict(size=8, color="#4a6680")), zeroline=False))
        st.plotly_chart(fig_bench, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Row 2 : VaR paramétrique + Volatilité ────────────────────────────
    row2_left, row2_right = st.columns([2, 1])

    with row2_left:
        st.markdown('<div class="pnl"><div class="pnl-title">Parametric VaR </div>', unsafe_allow_html=True)

        x_var = np.linspace(-4 * vol_daily, 4 * vol_daily, 400)
        y_var = (1 / (vol_daily * np.sqrt(2 * np.pi))) * np.exp(-0.5 * (x_var / vol_daily) ** 2)
        x_eur = x_var * V

        fig_var = go.Figure()
        fig_var.add_trace(go.Scatter(x=x_eur, y=y_var, mode="lines",
            line=dict(color="rgba(0,207,255,0.6)", width=2), name="Normal distribution", hoverinfo="skip"))

        mask_99 = x_eur <= -var_99_param
        if mask_99.any():
            fig_var.add_trace(go.Scatter(
                x=np.concatenate([x_eur[mask_99], [x_eur[mask_99][-1]], [x_eur[mask_99][0]]]),
                y=np.concatenate([y_var[mask_99], [0], [0]]),
                fill="toself", fillcolor="rgba(200,0,0,0.5)",
                line=dict(width=0), name="99% VaR region", hoverinfo="skip"))

        mask_95 = (x_eur <= -var_95_param) & (x_eur > -var_99_param)
        if mask_95.any():
            fig_var.add_trace(go.Scatter(
                x=np.concatenate([x_eur[mask_95], [x_eur[mask_95][-1]], [x_eur[mask_95][0]]]),
                y=np.concatenate([y_var[mask_95], [0], [0]]),
                fill="toself", fillcolor="rgba(255,140,0,0.45)",
                line=dict(width=0), name="95% VaR region", hoverinfo="skip"))

        fig_var.add_vline(x=-var_95_param, line_dash="dash", line_color="#ffaa00", line_width=2,
            annotation_text=f"VaR 95%<br>-€{var_95_param:,.0f}",
            annotation_font=dict(color="#ffaa00", size=9), annotation_position="top left")
        fig_var.add_vline(x=-var_99_param, line_dash="dash", line_color="#ff2222", line_width=2,
            annotation_text=f"VaR 99%<br>-€{var_99_param:,.0f}",
            annotation_font=dict(color="#ff2222", size=9), annotation_position="top left")
        fig_var.add_vline(x=0, line_dash="dot", line_color="#4cff72", line_width=1)
       

        fig_var.update_layout(
            paper_bgcolor="#12181f", plot_bgcolor="#0b0e13",
            font=dict(family=SYS_FONT, color="#4a6680", size=10),
            margin=dict(l=60, r=16, t=24, b=50), height=280, showlegend=True,
            legend=dict(orientation="v", x=0.01, y=0.99, xanchor="left", yanchor="top", font=dict(size=9, color="#7a9ab8"), bgcolor="rgba(0,0,0,0)"),
            xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"),
                      tickformat="€,.0f", title=dict(text="Daily P&L (€)", font=dict(size=9, color="#4a6680"))),
            yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"),
                      title=dict(text="Density", font=dict(size=9, color="#4a6680"))))
        st.plotly_chart(fig_var, use_container_width=True, config={"displayModeBar": False})

        
        st.markdown('</div>', unsafe_allow_html=True)

    with row2_right:
        st.markdown('<div class="pnl"><div class="pnl-title">Portfolio Volatility</div>', unsafe_allow_html=True)
        rolling_vol = daily_returns.rolling(7).std() * np.sqrt(252) * 100
        fig_vol = go.Figure()
        fig_vol.add_trace(go.Scatter(x=hist_vol["Date"], y=rolling_vol, mode="lines", fill="tozeroy",
            fillcolor="rgba(0,207,255,0.08)", line=dict(color="#00cfff", width=1.5),
            hovertemplate="%{x|%d %b}<br>Vol: %{y:.2f}%<extra></extra>"))
        fig_vol.update_layout(
            paper_bgcolor="#12181f", plot_bgcolor="#0b0e13",
            font=dict(family=SYS_FONT, color="#4a6680", size=9),
            margin=dict(l=40, r=8, t=8, b=30), height=130, showlegend=False,
            xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=8, color="#4a6680"), tickformat="%d %b", tickangle=-30),
            yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=8, color="#4a6680"), ticksuffix="%"))
        st.plotly_chart(fig_vol, use_container_width=True, config={"displayModeBar": False})
        st.markdown(f"""
        <div class="sr"><span class="sr-lbl">Period</span><span class="sr-o">{vol_period:.2f}%</span></div>
        <div class="sr"><span class="sr-lbl">Annualized Vol.</span><span class="sr-o">{vol_annual:.2f}%</span></div>
        <div class="sr"><span class="sr-lbl">Observed Days</span><span class="sr-w">{n_days}</span></div>
        <div class="sr"><span class="sr-lbl">Daily σ</span><span class="sr-w">{vol_daily*100:.3f}%</span></div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# TAB 3 — LONG-SHORT STRATEGY
# ═══════════════════════════════════════════════════════════════════════════
with tabs[3]:
    @st.cache_data
    def load_ls_data():
        df_returns = pd.read_excel(r"\data\DATA NDQ.xlsx", sheet_name="RETURNS")
        df_pricetobook = pd.read_excel(r"\data\DATA NDQ.xlsx", sheet_name="PRICE TO BOOK")
        df_benchmarksreturn = pd.read_excel(r"\data\DATA NDQ.xlsx", sheet_name="BENCHMARK RETURNS")
        for df, col in [(df_returns, "Unnamed: 0"), (df_pricetobook, "Unnamed: 0"), (df_benchmarksreturn, "Unnamed: 0")]:
            df.rename(columns={col: "Month_End"}, inplace=True)
            df["Month_End"] = pd.to_datetime(df["Month_End"], dayfirst=True)
            df.set_index("Month_End", inplace=True)
        df_pricetobook = df_pricetobook.ffill().bfill()

        # FIX : strip %, convertir en float décimal, supprimer les lignes vides
        df_benchmarksreturn["NDX Index"] = pd.to_numeric(
            df_benchmarksreturn["NDX Index"], errors="coerce"
        ) 
        df_benchmarksreturn = df_benchmarksreturn.dropna(subset=["NDX Index"])

        return df_returns, df_pricetobook, df_benchmarksreturn

    with st.spinner('Loading Long-Short data...'):
        ls_returns, ls_pricetobook, ls_benchmark = load_ls_data()

    st.markdown('<div class="sec-title">LONG-SHORT STRATEGY — MOMENTUM & VALUE</div>', unsafe_allow_html=True)
    p1, p2, p3, p4 = st.columns(4)
    with p1: ls_capital  = st.number_input("Capital par Leg (€)", min_value=1000, max_value=1000000, value=10000, step=1000, key="ls_cap")
    with p2: ls_nb_long  = st.slider("Nb LONG stocks",  0, 94, 15, key="ls_long")
    with p3: ls_nb_short = st.slider("Nb SHORT stocks", 0, 94, 15, key="ls_short")
    with p4:
        ls_start = st.date_input("Start", value=pd.to_datetime("2006-01-01"), key="ls_start")
        ls_end   = st.date_input("End",   value=pd.to_datetime("2026-03-31"), key="ls_end")

    @st.cache_data
    def calculate_ls_strategy(df_returns, df_pricetobook, nb_long, nb_short, start, end):
        returns_12m_mean = df_returns.rolling(window=12).apply(lambda x: x[:-1].mean(), raw=False)
        returns_12m_std  = df_returns.rolling(window=12).apply(lambda x: x[:-1].std(),  raw=False)
        score_momentum   = (returns_12m_mean - returns_12m_mean.mean(axis=1).values.reshape(-1,1)) / returns_12m_std
        inv_pb           = 1 / df_pricetobook
        inv_pb_12m_mean  = inv_pb.rolling(window=12).apply(lambda x: x[:-1].mean(), raw=False)
        inv_pb_12m_std   = inv_pb.rolling(window=12).apply(lambda x: x[:-1].std(),  raw=False)
        score_value      = (inv_pb_12m_mean - inv_pb_12m_mean.mean(axis=1).values.reshape(-1,1)) / inv_pb_12m_std
        score_global     = (score_momentum + score_value) / 2
        long_pf, short_pf = {}, {}
        for month in score_global.index:
            sc = score_global.loc[month]
            long_pf[month]  = sc.sort_values(ascending=False).index[:nb_long].tolist()
            short_pf[month] = sc.sort_values(ascending=True).index[:nb_short].tolist()
        long_weights_df  = pd.DataFrame(index=score_global.index, columns=score_global.columns)
        short_weights_df = pd.DataFrame(index=score_global.index, columns=score_global.columns)
        for month in score_global.index:
            sc = score_global.loc[month]
            ls = sc[long_pf[month]].abs();  long_weights_df.loc[month,  long_pf[month]]  =  ls / ls.sum()
            ss = sc[short_pf[month]].abs(); short_weights_df.loc[month, short_pf[month]] = -ss / ss.sum()
        positions_rebalanced = (long_weights_df.fillna(0) + short_weights_df.fillna(0)).shift(1)
        portfolio_returns    = (positions_rebalanced * df_returns).sum(axis=1).loc[str(start):str(end)]
        df_r  = df_returns.loc[str(start):str(end)]
        long_only_returns  = (long_weights_df.fillna(0).shift(1).loc[str(start):str(end)]  * df_r).sum(axis=1).dropna()
        short_only_returns = (short_weights_df.fillna(0).shift(1).loc[str(start):str(end)] * df_r).sum(axis=1).dropna()
        return portfolio_returns, long_only_returns, short_only_returns, positions_rebalanced.loc[str(start):str(end)], long_pf, short_pf, score_global

    with st.spinner('Calculating strategy...'):
        ls_port_ret, ls_long_ret, ls_short_ret, ls_positions, ls_long_pf, ls_short_pf, ls_score = calculate_ls_strategy(
            ls_returns, ls_pricetobook, ls_nb_long, ls_nb_short, ls_start, ls_end)

    def ls_calc_balance(portfolio_returns, long_only_returns, short_only_returns, initial_capital):
            ls_total_capital = initial_capital * 2
            # Si pas de short, on investit tout le capital dans le long
            long_start = ls_total_capital if ls_nb_short == 0 else initial_capital
            short_start = 0 if ls_nb_short == 0 else initial_capital
            
            bl, bs = long_start, short_start
            bal, bal_l, bal_s = [], [], []
            for date in portfolio_returns.index:
                rl = long_only_returns.loc[date] if date in long_only_returns.index else 0
                rs = short_only_returns.loc[date] if date in short_only_returns.index else 0
                bl *= (1 + rl)
                if bs > 0:
                    bs *= (1 + rs)
                bal.append(bl + bs); bal_l.append(bl); bal_s.append(bs)
            return pd.DataFrame({'Balance_Total': bal, 'Balance_Long': bal_l, 'Balance_Short': bal_s,
                                'Return_Month': portfolio_returns.values}, index=portfolio_returns.index)
    
    ls_balance        = ls_calc_balance(ls_port_ret, ls_long_ret, ls_short_ret, ls_capital)
    ls_total_capital  = ls_capital * 2
    active_capital = ls_capital if ls_nb_short == 0 else ls_total_capital
    ls_benchmark_period = ls_benchmark.loc[str(ls_start):str(ls_end)]
    ls_bench_cumul    = (1 + ls_benchmark_period["NDX Index"]).cumprod() * ls_total_capital
    ls_final          = ls_balance['Balance_Total'].iloc[-1]
    long_start = ls_total_capital if ls_nb_short == 0 else ls_capital
    ls_perf_long  = (ls_balance['Balance_Long'].iloc[-1] - long_start) / long_start * 100
    ls_perf_short = (ls_balance['Balance_Short'].iloc[-1] - ls_capital) / ls_capital * 100 if ls_nb_short > 0 else 0.0
    ls_perf_total = (ls_balance['Balance_Total'].iloc[-1] - ls_total_capital) / ls_total_capital * 100
    ls_bench_perf     = (ls_bench_cumul.iloc[-1] - ls_total_capital) / ls_total_capital * 100
    ls_sharpe         = (ls_port_ret.mean() / ls_port_ret.std()) * np.sqrt(12) if ls_port_ret.std() != 0 else 0
    pc = "c-green" if ls_perf_total >= 0 else "c-red"

    st.markdown(f"""
    <div class="kpi-bar">
      <div class="kpi-cell"><div class="kpi-lbl">Total Capital</div><div class="kpi-val c-white">{ls_total_capital:,.0f}€</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Performance Totale</div><div class="kpi-val {pc}">{ls_perf_total:+.2f}%</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Long</div><div class="kpi-val c-green">{ls_perf_long:+.2f}%</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Short</div><div class="kpi-val c-red">{ls_perf_short:+.2f}%</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Benchmark NDX</div><div class="kpi-val c-orange">{ls_bench_perf:+.2f}%</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Sharpe Ratio</div><div class="kpi-val c-cyan">{ls_sharpe:.2f}</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Capital Final</div><div class="kpi-val c-white">{ls_final:,.0f}€</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="pnl"><div class="pnl-title">CAPITAL EVOLUTION — STRATEGY VS BENCHMARK</div>', unsafe_allow_html=True)
    fig_ls = go.Figure()
    fig_ls.add_trace(go.Scatter(x=ls_balance.index, y=ls_balance['Balance_Total'], mode='lines', name='Long-Short Strategy', line=dict(color='#00cfff', width=2)))
    fig_ls.add_trace(go.Scatter(x=ls_bench_cumul.index, y=ls_bench_cumul.values, mode='lines', name='NDX Benchmark', line=dict(color='#ffaa00', width=2)))
    fig_ls.add_hline(y=ls_total_capital, line_dash="dash", line_color="#ff4444", annotation_text=f"Initial: {ls_total_capital:,.0f}€", annotation_font_color="#ff4444")
    fig_ls.update_layout(paper_bgcolor="#12181f", plot_bgcolor="#0b0e13", font=dict(family=SYS_FONT, color="#4a6680", size=10),
        margin=dict(l=60, r=16, t=16, b=40), height=350,
        legend=dict(orientation="h", x=0.01, y=0.99, font=dict(size=9, color="#7a9ab8"), bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680")),
        yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"), tickformat=",.0f"))
    st.plotly_chart(fig_ls, use_container_width=True, config={"displayModeBar": False})
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Annual Performance : Strategy vs Benchmark ────────────────────────
    st.markdown('<div class="pnl"><div class="pnl-title">ANNUAL PERFORMANCE — STRATEGY VS BENCHMARK</div>', unsafe_allow_html=True)

    annual_strat, annual_long, annual_short = {}, {}, {}
    for year in ls_balance.index.year.unique():
        yd = ls_balance[ls_balance.index.year == year]
        if len(yd) > 0:
            annual_strat[year] = (yd['Balance_Total'].iloc[-1] - yd['Balance_Total'].iloc[0]) / yd['Balance_Total'].iloc[0] * 100
            annual_long[year]  = (yd['Balance_Long'].iloc[-1]  - yd['Balance_Long'].iloc[0])  / yd['Balance_Long'].iloc[0]  * 100
            annual_short[year] = (yd['Balance_Short'].iloc[-1] - yd['Balance_Short'].iloc[0]) / yd['Balance_Short'].iloc[0] * 100

    annual_bench = ls_benchmark_period["NDX Index"].groupby(ls_benchmark_period.index.year).apply(lambda x: (1+x).prod()-1) * 100
    annual_df = pd.DataFrame({
        'Year': list(annual_strat.keys()),
        'Strategy (%)': list(annual_strat.values()),
        'Benchmark (%)': annual_bench.reindex(list(annual_strat.keys())).values,
        'Long (%)': list(annual_long.values()),
        'Short (%)': list(annual_short.values()),
    })

    fig_ann = go.Figure()
    fig_ann.add_trace(go.Bar(
        x=annual_df['Year'], y=annual_df['Strategy (%)'], name='Strategy',
        marker_color='blue',
        hovertemplate='Year: %{x}<br>Performance: %{y:.2f}%<extra></extra>'
    ))
    fig_ann.add_trace(go.Bar(
        x=annual_df['Year'], y=annual_df['Benchmark (%)'], name='Benchmark',
        marker_color='orange',
        hovertemplate='Year: %{x}<br>Performance: %{y:.2f}%<extra></extra>'
    ))
    fig_ann.update_layout(
        paper_bgcolor="#12181f", plot_bgcolor="#0b0e13",
        font=dict(family=SYS_FONT, color="#4a6680", size=10),
        margin=dict(l=50, r=16, t=8, b=40), height=350, barmode='group',
        legend=dict(orientation="h", x=0.01, y=0.99, font=dict(size=9, color="#7a9ab8"), bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680")),
        yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"), ticksuffix="%"))
    st.plotly_chart(fig_ann, use_container_width=True, config={"displayModeBar": False})

    # Tableau annuel
    st.dataframe(
        annual_df.style.format({
            'Strategy (%)': '{:.2f}%',
            'Benchmark (%)': '{:.2f}%',
            'Long (%)': '{:.2f}%',
            'Short (%)': '{:.2f}%',
        }).background_gradient(cmap='RdYlGn', subset=['Strategy (%)', 'Long (%)', 'Short (%)'], vmin=-20, vmax=20),
        use_container_width=True, hide_index=True
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Monthly Performance Details ───────────────────────────────────────
    st.markdown('<div class="pnl"><div class="pnl-title">MONTHLY PERFORMANCE DETAILS</div>', unsafe_allow_html=True)

    monthly_perf = pd.DataFrame(index=ls_balance.index)
    monthly_perf['Total (%)']     = ls_balance['Return_Month'] * 100
    monthly_perf['Long (%)']      = (ls_long_ret.reindex(ls_balance.index) * 100).fillna(0)
    monthly_perf['Short (%)']     = (ls_short_ret.reindex(ls_balance.index) * 100).fillna(0)
    monthly_perf['Benchmark (%)'] = (ls_benchmark_period["NDX Index"].reindex(ls_balance.index) * 100).fillna(0)
    monthly_perf['Year']          = monthly_perf.index.year
    monthly_perf['Month']         = monthly_perf.index.strftime('%B')

    mf1, mf2 = st.columns(2)
    with mf1:
        sel_yr = st.selectbox("Filter by Year",  ['All'] + sorted(monthly_perf['Year'].unique().tolist()), key="ls_yr")
    with mf2:
        sel_mo = st.selectbox("Filter by Month", ['All','January','February','March','April','May','June',
                                                   'July','August','September','October','November','December'], key="ls_mo")

    filt = monthly_perf.copy()
    if sel_yr != 'All': filt = filt[filt['Year'] == sel_yr]
    if sel_mo != 'All': filt = filt[filt['Month'] == sel_mo]

    display_monthly = filt[['Total (%)', 'Long (%)', 'Short (%)', 'Benchmark (%)']].copy()
    display_monthly.index = filt.index.strftime('%Y-%m-%d')
    st.dataframe(
        display_monthly.style.format({
            'Total (%)': '{:.2f}%', 'Long (%)': '{:.2f}%',
            'Short (%)': '{:.2f}%', 'Benchmark (%)': '{:.2f}%'
        }).background_gradient(cmap='RdYlGn', subset=['Total (%)', 'Long (%)', 'Short (%)'], vmin=-10, vmax=10),
        use_container_width=True, height=400
    )

    if len(filt) > 0:
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Avg Monthly Return", f"{filt['Total (%)'].mean():.2f}%")
        s2.metric("Best Month",         f"{filt['Total (%)'].max():.2f}%")
        s3.metric("Worst Month",        f"{filt['Total (%)'].min():.2f}%")
        s4.metric("Win Rate",           f"{(filt['Total (%)'] > 0).sum() / len(filt) * 100:.1f}%")

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Asset List — Top 20 Long Positions ───────────────────────────────
    st.markdown('<div class="pnl"><div class="pnl-title">📋 GENERATED ASSET LIST — TOP 20 LONG POSITIONS (MARCH)</div>', unsafe_allow_html=True)

    asset_df = pd.read_csv(EXPORT_PATH, sep=";", decimal=".")
    asset_df.columns = [c.strip().lstrip('\ufeff') for c in asset_df.columns]
    asset_df = asset_df[asset_df["Score"] > 0].head(20).reset_index(drop=True)

    st.markdown("""
    <div style="display:grid;grid-template-columns:2fr 1.5fr 1.5fr 2fr;gap:8px;
                padding:4px 8px;border-bottom:1px solid #1c2633;margin-bottom:4px;">
      <span style="font-size:9px;color:#4a6680;text-transform:uppercase;letter-spacing:1px;">Stock</span>
      <span style="font-size:9px;color:#4a6680;text-transform:uppercase;letter-spacing:1px;">Score</span>
      <span style="font-size:9px;color:#4a6680;text-transform:uppercase;letter-spacing:1px;">Weight (%)</span>
      <span style="font-size:9px;color:#4a6680;text-transform:uppercase;letter-spacing:1px;">Capital (€)</span>
    </div>
    """, unsafe_allow_html=True)

    for _, row in asset_df.iterrows():
        score_color = "#4cff72" if row["Score"] > 0 else "#ff4444"
        st.markdown(f"""
        <div style="display:grid;grid-template-columns:2fr 1.5fr 1.5fr 2fr;gap:8px;
                    padding:5px 8px;border-bottom:1px solid #0d1420;align-items:center;">
          <span style="font-size:13px;font-weight:700;color:#e8f0f8;">{row['Stock']}</span>
          <span style="font-size:12px;font-weight:600;color:{score_color};">{row['Score']:.4f}</span>
          <span style="font-size:12px;color:#7a9ab8;">{row['Poids (%)']:.2f}%</span>
          <span style="font-size:12px;font-weight:600;color:#00cfff;">€{row['Capital (€)']:,.2f}</span>
        </div>
        """, unsafe_allow_html=True)

    total_capital_long = asset_df["Capital (€)"].sum()
    total_weight_long  = asset_df["Poids (%)"].sum()
    st.markdown(f"""
    <div style="display:grid;grid-template-columns:2fr 1.5fr 1.5fr 2fr;gap:8px;
                padding:6px 8px;border-top:1px solid #1c2633;margin-top:4px;">
      <span style="font-size:10px;font-weight:700;color:#ffaa00;">TOTAL TOP 20</span>
      <span style="font-size:10px;color:#4a6680;">—</span>
      <span style="font-size:10px;font-weight:700;color:#ffaa00;">{total_weight_long:.2f}%</span>
      <span style="font-size:10px;font-weight:700;color:#ffaa00;">€{total_capital_long:,.2f}</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)
    
    # ── Max Drawdown — Strategy vs Benchmark ──────────────────────────────
    st.markdown('<div class="pnl"><div class="pnl-title">MAX DRAWDOWN — STRATEGY VS BENCHMARK</div>', unsafe_allow_html=True)

    # Drawdown strategy
    cumul_strat = ls_balance['Balance_Total']
    peak_strat = cumul_strat.cummax()
    dd_strat = (cumul_strat - peak_strat) / peak_strat * 100

    # Drawdown benchmark
    cumul_bench = (1 + ls_benchmark_period["NDX Index"]).cumprod() * ls_total_capital
    peak_bench = cumul_bench.cummax()
    dd_bench = (cumul_bench - peak_bench) / peak_bench * 100

    fig_dd = go.Figure()
    fig_dd.add_trace(go.Scatter(x=dd_strat.index, y=dd_strat, mode="lines", fill="tozeroy",
        fillcolor="rgba(0,207,255,0.08)", line=dict(color="#00cfff", width=2), name=f"Strategy (max: {dd_strat.min():.2f}%)"))
    fig_dd.add_trace(go.Scatter(x=dd_bench.index, y=dd_bench, mode="lines", fill="tozeroy",
        fillcolor="rgba(255,170,0,0.08)", line=dict(color="#ffaa00", width=2), name=f"NDX Benchmark (max: {dd_bench.min():.2f}%)"))
    fig_dd.update_layout(
        paper_bgcolor="#12181f", plot_bgcolor="#0b0e13",
        font=dict(family=SYS_FONT, color="#4a6680", size=10),
        margin=dict(l=60, r=16, t=16, b=40), height=300,
        legend=dict(orientation="h", x=0.01, y=0.99, font=dict(size=10, color="#7a9ab8"), bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680")),
        yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"), ticksuffix="%"))
    st.plotly_chart(fig_dd, use_container_width=True, config={"displayModeBar": False})

    st.markdown(f"""
    <div style="display:flex;gap:40px;font-size:13px;margin-top:4px;">
      <span style="color:#7a9ab8;">Strategy Max DD: <b style="color:#00cfff;">{dd_strat.min():.2f}%</b></span>
      <span style="color:#7a9ab8;">Benchmark Max DD: <b style="color:#ffaa00;">{dd_bench.min():.2f}%</b></span>
    </div>
    """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 4 — BONDS STRATEGY
# ═══════════════════════════════════════════════════════════════════════════
with tabs[4]:

    # ── Données statiques de la Bonds Strategy Version 2 ──────────────────
    BONDS_PORTFOLIO = [
        # (Class, Sub-class, Asset, Ticker, Alloc_EUR, Currency, Coupon, Maturity, Last_Price, Qty, Amount_EUR, Short_Note, Sector)
        ("Long Duration Gov & IG", "Ultra-Long Sovereign",   "US Treasury 2048",   "B-T-3.125-15052048",     30000, "USD", 3.125, 2048, 754.66,  46, 29771.20, "Max duration & convexity play",          "Government"),
        ("Long Duration Gov & IG", "Ultra-Long Sovereign",   "German Bund 2048",   "DE-B-1.250-15112048",    15000, "EUR", 1.250, 2048, 656.80,  22, 14586.44, "Core EUR risk-free duration anchor",      "Government"),
        ("Long Duration Gov & IG", "Ultra-Long Sovereign",   "French OAT 2045",    "FR-O-3.250-25052045",    15000, "EUR", 3.250, 2045, 886.65,  16, 14291.52, "Yield pickup vs Bund, semi-core",        "Government"),
        ("Long Duration Gov & IG", "Long-Dated IG Tech",     "Microsoft 2039",     "B-MSFT-5.2-01062039",    21000, "USD", 5.200, 2039, 1044.08, 23, 20572.81, "IG credit + rate-cut convexity",         "Technology"),
        ("Long Duration Gov & IG", "Long-Dated IG Tech",     "Apple 2031",         "B-APPL-1.700-05082031",  21000, "USD", 1.700, 2031, 887.61,  27, 20349.63, "Ultra-high quality, high rate sensitivity","Technology"),
        ("Long Duration Gov & IG", "Long-Dated IG Tech",     "ASML 2032",          "NL-ASML-2.250-17042032", 18000, "EUR", 2.250, 2032, 966.33,  18, 17515.98, "Monopoly-like semiconductor credit",     "Technology"),
        ("High Yield Credit",      "Fallen Angels",          "Ford 2026",          "B-F-7.500-01082026",     18000, "USD", 7.500, 2026, 1015.50, 20, 17257.20, "Short-dated HY carry + pull-to-par",     "Industrials"),
        ("High Yield Credit",      "Fallen Angels",          "Kraft Heinz 2040",   "B-KRFT-6.5-09022040",    18000, "USD", 6.500, 2040, 1079.59, 19, 17932.58, "Defensive staples HY income",            "Consumer Staples"),
        ("High Yield Credit",      "Stable Cash-Flow HY",    "Altria 2038",        "B-MO-9.95-10112038",     15000, "USD", 9.950, 2038, 1366.10, 12, 14159.52, "~10% coupon, stable tobacco cash flows", "Consumer Staples"),
        ("High Yield Credit",      "Stable Cash-Flow HY",    "Philip Morris 2038", "B-PM-6.375-16052038",    15000, "USD", 6.375, 2038, 1112.83, 15, 14352.75, "Global tobacco, lower regulatory risk",  "Consumer Staples"),
        ("High Yield Credit",      "Event / Cyclical HY",    "Boeing 2043",        "B-BA-6.875-15102043",    15000, "USD", 6.875, 2043, 1080.00, 15, 14152.65, "Cyclical recovery + spread compression", "Industrials"),
        ("High Yield Credit",      "Event / Cyclical HY",    "US Steel 2029",      "B-X-6.875-01032029",      9000, "USD", 6.875, 2029, 1004.57, 10,  8780.00, "Short-dated cyclical carry play",        "Materials"),
        ("Relative Value",         "Sovereign Spread Trade", "Long BTP 2040",      "IT-B-3.100-03012040",    15000, "EUR", 3.100, 2040, 925.86,  15, 14092.65, "BTP-Bund spread compression trade",      "Government"),
        ("Relative Value",         "Sovereign Spread Trade", "Short Bund 2040",    "DE-B-4.750-04072040",    15000, "EUR", 4.750, 2040, 1184.56, 12, 14269.20, "Hedge leg — duration-neutral RV",        "Government"),
        ("Relative Value",         "Credit RV",              "JPM 2038",           "B-JPM-6.4-15052038",     10000, "USD", 6.400, 2038, 1152.58, 10,  9908.00, "Bank credit RV anchor",                  "Financials"),
        ("Relative Value",         "Credit RV",              "JPM 2040",           "B-JPM-5.5-15102040",      8000, "USD", 5.500, 2040, 1041.40,  8,  7189.28, "Duration extension within issuer",       "Financials"),
        ("Relative Value",         "Credit RV",              "Sanofi 2038",        "FR-SNY-1.875-21032038",   7000, "EUR", 1.875, 2038, 838.91,   8,  6798.72, "Non-financial IG pharma diversifier",    "Healthcare"),
        ("Relative Value",         "Credit RV",              "TotalEnergies 2039", "FR-TTE-1.535-31052039",   5000, "EUR", 1.535, 2039, 749.91,   6,  4534.68, "Energy IG, macro-sensitive leg",         "Energy"),
        ("Tactical Cash",          "Liquidity Buffer",       "Cash Reserve",       "CASH",                   30000, "EUR", 0.000, None, 1.00,     0, 30000.00, "Optionality — deploy on dislocation",    "Cash"),
    ]

    # ── Construire le DataFrame principal ─────────────────────────────────
    bp = pd.DataFrame(BONDS_PORTFOLIO, columns=[
        "Class", "Sub_class", "Asset", "Ticker", "Alloc_EUR", "Currency",
        "Coupon", "Maturity", "Last_Price", "Qty", "Amount_EUR", "Short_Note", "Sector"
    ])

    # Montant réel investi = somme des BUY bonds depuis le Trading Journal
    bonds_tj_buys = tj[(tj["Asset Class"] == "Bond") & (tj["Direction"] == "Buy")].copy()
    TOTAL_PORTFOLIO = bonds_tj_buys["Amount($)"].abs().sum()  # ~259k réel en $, affiché avec sigle €
    bp["Weight_pct"] = bp["Amount_EUR"] / TOTAL_PORTFOLIO * 100

    # P&L depuis le Dashboard
    bonds_db_pnl = {
        "BTP 3.100% 03-Jan-2040":                          -309.85,
        "JPMORGAN CHASE & CO - 5.5% - Oct 2040":            56.55,
        "JPMORGAN CHASE & CO - 6.4% - May 2038":            19.74,
        "Philip Morris - 6.375% - May 2038":               121.46,
        "Boeing":                                           590.31,
        "Ford Motor Co. - 7.500% - Aug 2026":              313.00,
        "Apple Inc - 1.700% - Aug 2031":                   270.69,
        "United States Steel Corp - 6.875% - Mar 2029":    162.71,
        "Kraft Heinz Foods - 6.500% - Feb 2040":           225.91,
        "Microsoft Corp - 5.2% - Jun 2039":                492.78,
        "Altria Group Inc - 9.95% - Nov 2038":             456.38,
        "T-BOND 3.125% 15/05/2048":                        711.64,
        "OAT 3.250% 25-May-2045":                         -158.40,
        "Sanofi S.A - 1.875% - Mar 2038":                 -182.16,
        "ASML Holding - 2.250% - May 2032":               -348.30,
        "BUND 1.250% 15-Nov-2048":                         -73.86,
        "TotalEnergies Capital Intl. - 1.535% - May 2039": -107.40,
        "BUND 4.750% 04-Jul-2040":                        -134.98,
    }

    # Mapping Asset → P&L
    pnl_map = {
        "US Treasury 2048":   711.64,
        "German Bund 2048":   -73.86,
        "French OAT 2045":   -158.40,
        "Microsoft 2039":     492.78,
        "Apple 2031":         270.69,
        "ASML 2032":         -348.30,
        "Ford 2026":          313.00,
        "Kraft Heinz 2040":   225.91,
        "Altria 2038":        456.38,
        "Philip Morris 2038": 121.46,
        "Boeing 2043":        590.31,
        "US Steel 2029":      162.71,
        "Long BTP 2040":     -309.85,
        "Short Bund 2040":   -134.98,
        "JPM 2038":            19.74,
        "JPM 2040":            56.55,
        "Sanofi 2038":       -182.16,
        "TotalEnergies 2039": -107.40,
        "Cash Reserve":         0.00,
    }
    bp["PnL_EUR"] = bp["Asset"].map(pnl_map).fillna(0)
    total_bond_pnl = bp["PnL_EUR"].sum()
    total_invested = TOTAL_PORTFOLIO  # basé sur Amount($) des BUY réels

    # ── Coupons reçus (Fixed Income dans Dashboard = entrées automatiques Stocktrak) ──
    db_raw = pd.read_excel(EXCEL_PATH, sheet_name="Dashboard", header=None)
    db_raw.columns = ["Full Trade cycle", "Date", "Instrument", "Asset class", "Trader",
                      "Net P&L", "Net P&L $", "Portfolio wallet currently", "Portfolio $", "Col10", "Col11"]
    db_raw["Net P&L"] = pd.to_numeric(db_raw["Net P&L"], errors="coerce")
    coupons_df = db_raw[db_raw["Asset class"].astype(str).str.strip() == "Fixed Income"].copy()
    coupons_df = coupons_df[coupons_df["Net P&L"].notna()].copy()
    coupons_df["Date"] = pd.to_datetime(coupons_df["Date"], errors="coerce")
    total_coupons = coupons_df["Net P&L"].sum()
    coupon_by_bond = coupons_df.groupby("Instrument")["Net P&L"].sum().reset_index()
    coupon_by_bond.columns = ["Instrument", "Coupon_recu"]
    total_return_all = total_bond_pnl + total_coupons
    total_return_pct = total_return_all / total_invested * 100

    # ── KPI globaux ────────────────────────────────────────────────────────
    n_bonds     = len(bp[bp["Asset"] != "Cash Reserve"])
    ig_alloc    = bp[bp["Class"] == "Long Duration Gov & IG"]["Amount_EUR"].sum()
    hy_alloc    = bp[bp["Class"] == "High Yield Credit"]["Amount_EUR"].sum()
    rv_alloc    = bp[bp["Class"] == "Relative Value"]["Amount_EUR"].sum()
    cash_alloc  = bp[bp["Class"] == "Tactical Cash"]["Amount_EUR"].sum()
    winners     = (bp["PnL_EUR"] > 0).sum()
    losers      = (bp["PnL_EUR"] < 0).sum()
    avg_coupon_rate = bp[bp["Asset"] != "Cash Reserve"]["Coupon"].mean()
    pnl_cls     = "c-green" if total_bond_pnl >= 0 else "c-red"
    pnl_pct     = total_bond_pnl / total_invested * 100
    ret_cls     = "c-green" if total_return_all >= 0 else "c-red"

    st.markdown('<div class="sec-title">BONDS STRATEGY — FIXED INCOME PORTFOLIO</div>', unsafe_allow_html=True)

    st.markdown(f"""
    <div class="kpi-bar">
      <div class="kpi-cell"><div class="kpi-lbl">Total Invested (€)</div><div class="kpi-val c-white">€{TOTAL_PORTFOLIO:,.0f}</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Net P&amp;L Capital</div><div class="kpi-val {pnl_cls}">€{total_bond_pnl:+,.2f}</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Coupons Received</div><div class="kpi-val c-green">€{total_coupons:+,.2f}</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Total Return</div><div class="kpi-val {ret_cls}">€{total_return_all:+,.2f}</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Total Return %</div><div class="kpi-val {ret_cls}">{total_return_pct:+.2f}%</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Bonds Held</div><div class="kpi-val c-cyan">{n_bonds}</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Avg Coupon Rate</div><div class="kpi-val c-orange">{avg_coupon_rate:.2f}%</div></div>
      <div class="kpi-cell"><div class="kpi-lbl">Winners / Losers</div><div class="kpi-val c-white"><span style="color:#4cff72">{winners}</span> / <span style="color:#ff4444">{losers}</span></div></div>
    </div>
    """, unsafe_allow_html=True)


    # ── Row 1 : Allocation pie + P&L by bond ──────────────────────────────
    row1a, row1b = st.columns([1, 1.6])

    with row1a:
        st.markdown('<div class="pnl"><div class="pnl-title">ALLOCATION BY CLASS</div>', unsafe_allow_html=True)
        alloc_class = bp.groupby("Class")["Amount_EUR"].sum().reset_index()
        colors_pie = ["#00cfff", "#ffaa00", "#4cff72", "#7a9ab8"]
        fig_alloc = go.Figure(go.Pie(
            labels=alloc_class["Class"],
            values=alloc_class["Amount_EUR"],
            hole=0.52,
            marker=dict(colors=colors_pie),
            textfont=dict(size=11, color="#e8f0f8"),
            hovertemplate="%{label}<br>€%{value:,.0f}<br>%{percent}<extra></extra>"
        ))
        fig_alloc.add_annotation(
            text=f"€{TOTAL_PORTFOLIO/1000:.0f}k", x=0.5, y=0.5,
            font=dict(size=16, color="#e8f0f8", family=SYS_FONT), showarrow=False
        )
        fig_alloc.update_layout(
            paper_bgcolor="#12181f", height=240,
            margin=dict(l=0, r=0, t=8, b=8),
            legend=dict(font=dict(size=10, color="#7a9ab8"), orientation="v", x=0.75, y=0.5)
        )
        st.plotly_chart(fig_alloc, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="pnl"><div class="pnl-title">ALLOCATION BY SECTOR</div>', unsafe_allow_html=True)
        alloc_sector = bp[bp["Asset"] != "Cash Reserve"].groupby("Sector")["Amount_EUR"].sum().sort_values(ascending=True)
        colors_sector = ["#ff4444" if v < 0 else "#00cfff" for v in alloc_sector.values]
        fig_sector = go.Figure(go.Bar(
            x=alloc_sector.values, y=alloc_sector.index,
            orientation='h',
            marker_color=colors_sector,
            hovertemplate="%{y}<br>€%{x:,.0f}<extra></extra>"
        ))
        fig_sector.update_layout(
            paper_bgcolor="#12181f", plot_bgcolor="#0b0e13",
            font=dict(family=SYS_FONT, color="#4a6680", size=10),
            margin=dict(l=10, r=16, t=8, b=8), height=220,
            xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"), tickformat="€,.0f"),
            yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=10, color="#7a9ab8"))
        )
        st.plotly_chart(fig_sector, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    with row1b:
        st.markdown('<div class="pnl"><div class="pnl-title">P&L BY BOND — CLOSED POSITIONS</div>', unsafe_allow_html=True)
        pnl_bonds = bp[bp["Asset"] != "Cash Reserve"].sort_values("PnL_EUR", ascending=True).copy()
        bar_clrs = ["#4cff72" if v >= 0 else "#ff4444" for v in pnl_bonds["PnL_EUR"]]
        fig_pnl = go.Figure(go.Bar(
            x=pnl_bonds["PnL_EUR"],
            y=pnl_bonds["Asset"],
            orientation='h',
            marker_color=bar_clrs,
            hovertemplate="%{y}<br>P&L: €%{x:+,.2f}<extra></extra>"
        ))
        fig_pnl.add_vline(x=0, line_color="#2a3a4a", line_width=1)
        fig_pnl.update_layout(
            paper_bgcolor="#12181f", plot_bgcolor="#0b0e13",
            font=dict(family=SYS_FONT, color="#4a6680", size=10),
            margin=dict(l=10, r=16, t=8, b=8), height=470,
            xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"),
                       tickformat="€,.0f", title=dict(text="Net P&L (€)", font=dict(size=9, color="#4a6680"))),
            yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=10, color="#7a9ab8"))
        )
        st.plotly_chart(fig_pnl, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Row 2 : Maturity ladder + Coupon scatter ───────────────────────────
    row2a, row2b = st.columns(2)

    with row2a:
        st.markdown('<div class="pnl"><div class="pnl-title">MATURITY LADDER — DURATION PROFILE</div>', unsafe_allow_html=True)
        mat_data = bp[bp["Maturity"].notna() & (bp["Asset"] != "Cash Reserve")].copy()
        mat_data["Maturity_yr"] = mat_data["Maturity"].astype(int)
        mat_grp = mat_data.groupby("Maturity_yr").agg(
            Amount_EUR=("Amount_EUR", "sum"),
            Assets=("Asset", lambda x: ", ".join(x))
        ).reset_index()
        bar_mat = ["#5e1ce6" if yr >= 2035 else "#ad1de5" if yr >= 2030 else "#e485e6" for yr in mat_grp["Maturity_yr"]]
        fig_mat = go.Figure(go.Bar(
            x=mat_grp["Maturity_yr"].astype(str),
            y=mat_grp["Amount_EUR"],
            marker_color=bar_mat,
            hovertemplate="Maturity: %{x}<br>€%{y:,.0f}<extra></extra>"
        ))
        fig_mat.update_layout(
            paper_bgcolor="#12181f", plot_bgcolor="#0b0e13",
            font=dict(family=SYS_FONT, color="#4a6680", size=10),
            margin=dict(l=50, r=16, t=8, b=40), height=220,
            xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"), tickangle=-30),
            yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"), tickformat="€,.0f")
        )
        st.plotly_chart(fig_mat, use_container_width=True, config={"displayModeBar": False})

        # Légende durée
        st.markdown("""
        <div style="display:flex;gap:16px;font-size:10px;margin-top:-8px;padding:4px 0;">
          <span><span style="color:#e485e6;font-weight:700;">■</span> Short (&lt;2030)</span>
          <span><span style="color:#ad1de5;font-weight:700;">■</span> Medium (2030–2034)</span>
          <span><span style="color:#5e1ce6;font-weight:700;">■</span> Long (2035+)</span>
        </div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with row2b:
        st.markdown('<div class="pnl"><div class="pnl-title">COUPON vs MATURITY — RISK/CARRY MAP</div>', unsafe_allow_html=True)
        scatter_data = bp[bp["Asset"] != "Cash Reserve"].copy()
        color_map = {
            "Long Duration Gov & IG": "#ffaa00",
            "High Yield Credit":      "#00cfff",
            "Relative Value":         "#4cff72",
        }
        size_map = scatter_data["Amount_EUR"] / 500

        fig_scatter = go.Figure()
        for cls, grp in scatter_data.groupby("Class"):
            fig_scatter.add_trace(go.Scatter(
                x=grp["Maturity"], y=grp["Coupon"],
                mode="markers+text",
                marker=dict(
                    size=(grp["Amount_EUR"] / 500).clip(8, 40),
                    color=color_map.get(cls, "#7a9ab8"),
                    opacity=0.85,
                    line=dict(width=1, color="#0b0e13")
                ),
                text=grp["Asset"].str.replace(r" \d{4}$", "", regex=True),  # supprime juste l'année à la fin
                textposition="top center",
                textfont=dict(size=8, color="#ffffff"),
                name=cls,
                hovertemplate="<b>%{text}</b><br>Maturity: %{x}<br>Coupon: %{y:.2f}%<extra></extra>"
            ))
        fig_scatter.update_layout(
            paper_bgcolor="#12181f", plot_bgcolor="#0b0e13",
            font=dict(family=SYS_FONT, color="#4a6680", size=10),
            margin=dict(l=50, r=16, t=8, b=40), height=260,
            showlegend=True,
            legend=dict(orientation="h", x=0, y=-0.25, font=dict(size=9, color="#7a9ab8"), bgcolor="rgba(0,0,0,0)"),
            xaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"),
                       title=dict(text="Maturity Year", font=dict(size=9, color="#4a6680"))),
            yaxis=dict(gridcolor="#1a2533", linecolor="#1c2633", tickfont=dict(size=9, color="#4a6680"),
                       ticksuffix="%", title=dict(text="Coupon (%)", font=dict(size=9, color="#4a6680")))
        )
        st.plotly_chart(fig_scatter, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Row 3 : Tableau détaillé par classe ────────────────────────────────
    st.markdown('<div class="sec-title">DETAILED HOLDINGS — BONDS STRATEGY V2</div>', unsafe_allow_html=True)

    CLASS_COLORS = {
        "Long Duration Gov & IG": "#ffaa00",
        "High Yield Credit":      "#00cfff",
        "Relative Value":         "#4cff72",
        "Tactical Cash":          "#7a9ab8",
    }
    CLASS_ALLOC = {
        "Long Duration Gov & IG": "40% = €120,000",
        "High Yield Credit":      "30% = €90,000",
        "Relative Value":         "20% = €60,000",
        "Tactical Cash":          "10% = €30,000",
    }

    for cls in ["Long Duration Gov & IG", "High Yield Credit", "Relative Value", "Tactical Cash"]:
        grp = bp[bp["Class"] == cls].copy()
        cls_total = grp["Amount_EUR"].sum()
        cls_pnl   = grp["PnL_EUR"].sum()
        cls_color = CLASS_COLORS[cls]
        pnl_c     = "#4cff72" if cls_pnl >= 0 else "#ff4444"

        st.markdown(f"""
        <div style="background:#0d1420;border:1px solid {cls_color}33;border-left:3px solid {cls_color};
                    border-radius:4px;padding:8px 14px;margin-bottom:6px;">
          <div style="display:flex;justify-content:space-between;align-items:center;">
            <div>
              <span style="font-size:12px;font-weight:700;color:{cls_color};letter-spacing:1px;">{cls.upper()}</span>
              <span style="font-size:10px;color:#4a6680;margin-left:12px;">{CLASS_ALLOC[cls]}</span>
            </div>
            <div style="display:flex;gap:24px;font-size:11px;">
              <span style="color:#7a9ab8;">Invested: <b style="color:#e8f0f8;">€{cls_total:,.0f}</b></span>
              <span style="color:#7a9ab8;">P&amp;L: <b style="color:{pnl_c};">€{cls_pnl:+,.2f}</b></span>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # En-tête colonnes
        st.markdown("""
        <div style="display:grid;grid-template-columns:2fr 1.2fr 0.7fr 0.7fr 0.8fr 0.8fr 0.8fr 2fr;
                    gap:6px;padding:4px 8px;border-bottom:1px solid #1c2633;margin-bottom:2px;">
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;letter-spacing:1px;">Asset</span>
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;letter-spacing:1px;">Sub-Class</span>
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;letter-spacing:1px;">Coupon</span>
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;letter-spacing:1px;">Maturity</span>
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;letter-spacing:1px;">Qty</span>
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;letter-spacing:1px;">Amount €</span>
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;letter-spacing:1px;">P&L €</span>
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;letter-spacing:1px;">Note</span>
        </div>
        """, unsafe_allow_html=True)

        for _, row in grp.iterrows():
            pnl_color  = "#4cff72" if row["PnL_EUR"] >= 0 else "#ff4444"
            mat_str    = str(int(row["Maturity"])) if pd.notna(row["Maturity"]) and row["Maturity"] else "—"
            coupon_str = f"{row['Coupon']:.3f}%" if row["Coupon"] > 0 else "—"
            qty_str    = str(int(row["Qty"])) if row["Qty"] > 0 else "—"
            curr_badge = f'<span style="font-size:9px;background:#1c2633;border-radius:2px;padding:1px 4px;color:#7a9ab8;margin-left:4px;">{row["Currency"]}</span>'

            st.markdown(f"""
            <div style="display:grid;grid-template-columns:2fr 1.2fr 0.7fr 0.7fr 0.8fr 0.8fr 0.8fr 2fr;
                        gap:6px;padding:5px 8px;border-bottom:1px solid #0d1420;align-items:center;">
              <div>
                <span style="font-size:12px;font-weight:700;color:#e8f0f8;">{row['Asset']}</span>
                {curr_badge}
              </div>
              <span style="font-size:10px;color:#7a9ab8;">{row['Sub_class']}</span>
              <span style="font-size:11px;font-weight:600;color:#ffaa00;">{coupon_str}</span>
              <span style="font-size:11px;color:#7a9ab8;">{mat_str}</span>
              <span style="font-size:11px;color:#e8f0f8;">{qty_str}</span>
              <span style="font-size:11px;font-weight:600;color:#00cfff;">€{row['Amount_EUR']:,.0f}</span>
              <span style="font-size:11px;font-weight:700;color:{pnl_color};">€{row['PnL_EUR']:+,.2f}</span>
              <span style="font-size:10px;color:#4a6680;font-style:italic;">{row['Short_Note']}</span>
            </div>
            """, unsafe_allow_html=True)

        # Total de la classe
        st.markdown(f"""
        <div style="display:grid;grid-template-columns:2fr 1.2fr 0.7fr 0.7fr 0.8fr 0.8fr 0.8fr 2fr;
                    gap:6px;padding:5px 8px;border-top:1px solid {cls_color}44;margin-bottom:12px;background:#0d1420;">
          <span style="font-size:10px;font-weight:700;color:{cls_color};">TOTAL {cls.upper()}</span>
          <span></span><span></span><span></span><span></span>
          <span style="font-size:10px;font-weight:700;color:{cls_color};">€{cls_total:,.0f}</span>
          <span style="font-size:10px;font-weight:700;color:{pnl_c};">€{cls_pnl:+,.2f}</span>
          <span></span>
        </div>
        """, unsafe_allow_html=True)

    # ── Row 4 : Stratégie originale (Version 1) ───────────────────────────
    st.markdown('<div class="sec-title">BONDS STRATEGY V1 — THEORETICAL ALLOCATION</div>', unsafe_allow_html=True)

    BONDS_V1 = [
        # IG Buckets
        ("IG-1 Gov Duration", "US",      "US Treasury 15-May-2040",          2040, 0.10, 200000, 30000, "Very long duration; maximizes rate sensitivity"),
        ("IG-1 Gov Duration", "US",      "US Treasury 15-Feb-2031",          2031, 0.08, 160000, 24000, "Medium duration benchmark, US curve reference"),
        ("IG-1 Gov Duration", "Germany", "German Bund 04-Jul-2040",          2040, 0.06, 120000, 18000, "Core EUR risk-free, USD vs EUR rate dynamics"),
        ("IG-1 Gov Duration", "France",  "French OAT 25-May-2043",           2043, 0.06, 120000, 18000, "Sovereign risk premium vs Bund"),
        ("IG-2 IG Corporate", "Tech",    "Apple Inc 3.25% Feb-2026",         2026, 0.06, 120000,     0, "Ultra-high quality IG benchmark"),
        ("IG-2 IG Corporate", "Tech",    "Microsoft Corp 5.20% Jun-2039",    2039, 0.06, 120000,     0, "Long-term IG tech, duration + fundamentals"),
        ("IG-2 IG Corporate", "Health",  "Johnson & Johnson 4.95% May-2033", 2033, 0.06, 120000,     0, "Defensive cash flows, low cyclicality"),
        ("IG-2 IG Corporate", "Finance", "JPMorgan Chase 5.50% Oct-2040",    2040, 0.06, 120000,     0, "Financial sector, rate/economic cycle"),
        ("IG-2 IG Corporate", "Utility", "Enel Finance 3.875% Mar-2029",     2029, 0.06, 120000,     0, "Stable regulated cash flows"),
        ("IG-3 Geography",   "France",  "AXA S.A. 3.25% May-2029",          2029, 0.05, 100000,     0, "European insurance IG in EUR"),
        ("IG-3 Geography",   "Germany", "Deutsche Telekom 1.75% Mar-2031",   2031, 0.05, 100000,     0, "Core European telecom, low-risk EUR"),
        ("IG-3 Geography",   "Europe",  "Nestlé/Danone equivalent",          2029, 0.05, 100000,     0, "Defensive consumer exposure"),
        # HY Buckets
        ("HY-1 Fallen Angels","Indust.", "Ford Motor 7.5% Aug-2026",         2026, 0.10, 200000,     0, "Classic fallen angel, spread recovery analysis"),
        ("HY-2 Cyclical HY",  "Energy",  "ENI S.p.A. 4.25% May-2033",        2033, 0.09, 180000,     0, "Commodity-linked, oil price sensitive"),
        ("HY-2 Cyclical HY",  "Indust.", "Boeing 6.875% Oct-2043",           2043, 0.08, 160000,     0, "High leverage, long maturity, credit cyclicality"),
        ("HY-3 Stable HY",   "Tobacco", "Altria Group 9.95% Nov-2038",       2038, 0.06, 120000,     0, "Very high coupon, stable tobacco cash flows"),
        ("HY-3 Stable HY",   "Tobacco", "Philip Morris 6.375% May-2038",     2038, 0.06, 120000,     0, "Defensive HY, strong global pricing power"),
    ]

    v1_df = pd.DataFrame(BONDS_V1, columns=["Bucket", "Country/Sector", "Bond", "Maturity", "Weight", "Amount", "Amount_Alloc", "Rationale"])

    v1_colors = {
        "IG-1 Gov Duration": "#00cfff",
        "IG-2 IG Corporate": "#4cff72",
        "IG-3 Geography":    "#88ccff",
        "HY-1 Fallen Angels":"#ffaa00",
        "HY-2 Cyclical HY":  "#ff8844",
        "HY-3 Stable HY":    "#ff4444",
    }

    for bucket in v1_df["Bucket"].unique():
        bgrp = v1_df[v1_df["Bucket"] == bucket]
        bc   = v1_colors.get(bucket, "#7a9ab8")
        total_b = bgrp["Amount"].sum()

        st.markdown(f"""
        <div style="background:#0d1420;border-left:3px solid {bc};border-radius:4px;
                    padding:6px 12px;margin-bottom:4px;">
          <span style="font-size:11px;font-weight:700;color:{bc};">{bucket}</span>
          <span style="font-size:9px;color:#4a6680;margin-left:10px;">Total: ${total_b:,.0f}</span>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div style="display:grid;grid-template-columns:2.5fr 1fr 0.7fr 0.8fr 0.8fr 2fr;
                    gap:6px;padding:3px 8px;border-bottom:1px solid #1c2633;margin-bottom:2px;">
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;">Bond</span>
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;">Country/Sector</span>
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;">Maturity</span>
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;">Weight</span>
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;">Amount $</span>
          <span style="font-size:8px;color:#4a6680;text-transform:uppercase;">Rationale</span>
        </div>
        """, unsafe_allow_html=True)

        for _, row in bgrp.iterrows():
            st.markdown(f"""
            <div style="display:grid;grid-template-columns:2.5fr 1fr 0.7fr 0.8fr 0.8fr 2fr;
                        gap:6px;padding:4px 8px;border-bottom:1px solid #0d1420;align-items:center;">
              <span style="font-size:11px;font-weight:600;color:#e8f0f8;">{row['Bond']}</span>
              <span style="font-size:10px;color:#7a9ab8;">{row['Country/Sector']}</span>
              <span style="font-size:11px;color:#ffaa00;">{row['Maturity']}</span>
              <span style="font-size:11px;color:#00cfff;">{row['Weight']*100:.0f}%</span>
              <span style="font-size:11px;color:#e8f0f8;">${row['Amount']:,.0f}</span>
              <span style="font-size:10px;color:#4a6680;font-style:italic;">{row['Rationale']}</span>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom:10px;'></div>", unsafe_allow_html=True)

    # ── Grand Total ────────────────────────────────────────────────────────
    st.markdown(f"""
    <div style="background:#0d1420;border:1px solid #00cfff44;border-radius:4px;
                padding:10px 16px;display:flex;justify-content:space-between;align-items:center;margin-top:8px;">
      <span style="font-size:12px;font-weight:700;color:#00cfff;letter-spacing:1px;">BONDS PORTFOLIO SUMMARY</span>
      <div style="display:flex;gap:32px;font-size:12px;">
        <span style="color:#7a9ab8;">Total Invested: <b style="color:#e8f0f8;">€{total_invested:,.0f}</b></span>
        <span style="color:#7a9ab8;">Net P&amp;L (Closed): <b style="color:{'#4cff72' if total_bond_pnl>=0 else '#ff4444'};">€{total_bond_pnl:+,.2f}</b></span>
        <span style="color:#7a9ab8;">Return: <b style="color:{'#4cff72' if pnl_pct>=0 else '#ff4444'};">{pnl_pct:+.2f}%</b></span>
        <span style="color:#7a9ab8;">Active Bonds: <b style="color:#00cfff;">{n_bonds}</b></span>
      </div>
    </div>
    """, unsafe_allow_html=True)

