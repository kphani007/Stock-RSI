"""
NIFTY Daily RSI Screener - Web App
----------------------------------
Same screener as nifty_rsi_screener.py, but as a Streamlit web app so it
lives at a public URL you can open every morning.

Deploy free on Streamlit Community Cloud:
  1. Put streamlit_app.py + requirements.txt in a GitHub repo.
  2. Go to https://share.streamlit.io -> "New app" -> pick the repo.
  3. Main file: streamlit_app.py -> Deploy.
  You get a URL like https://your-name.streamlit.app

Run locally instead:
  pip install -r requirements.txt
  streamlit run streamlit_app.py
"""

from __future__ import annotations

import datetime as dt
import io
import time

import pandas as pd
import streamlit as st
import yfinance as yf
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RSI_PERIOD = 14

NIFTY50 = [
    "ADANIENT.NS", "ADANIPORTS.NS", "APOLLOHOSP.NS", "ASIANPAINT.NS",
    "AXISBANK.NS", "BAJAJ-AUTO.NS", "BAJFINANCE.NS", "BAJAJFINSV.NS",
    "BEL.NS", "BHARTIARTL.NS", "CIPLA.NS", "COALINDIA.NS", "DRREDDY.NS",
    "EICHERMOT.NS", "GRASIM.NS", "HCLTECH.NS", "HDFCBANK.NS",
    "HDFCLIFE.NS", "HEROMOTOCO.NS", "HINDALCO.NS", "HINDUNILVR.NS",
    "ICICIBANK.NS", "INDUSINDBK.NS", "INFY.NS", "ITC.NS", "JSWSTEEL.NS",
    "KOTAKBANK.NS", "LT.NS", "M&M.NS", "MARUTI.NS", "NESTLEIND.NS",
    "NTPC.NS", "ONGC.NS", "POWERGRID.NS", "RELIANCE.NS", "SBILIFE.NS",
    "SBIN.NS", "SHRIRAMFIN.NS", "SUNPHARMA.NS", "TATACONSUM.NS",
    "TATAMOTORS.NS", "TATASTEEL.NS", "TCS.NS", "TECHM.NS", "TITAN.NS",
    "TRENT.NS", "ULTRACEMCO.NS", "WIPRO.NS",
]

HEADERS = [
    "Date", "Stock Name", "Current Price (₹)", "52 Week High (₹)", "RSI",
    "Recommendation to buy or sell", "1 Year Target (₹)",
]


def compute_rsi(close: pd.Series, period: int = RSI_PERIOD) -> float | None:
    close = close.dropna()
    if len(close) < period + 1:
        return None
    delta = close.diff()
    gain = delta.clip(lower=0.0)
    loss = -delta.clip(upper=0.0)
    avg_gain = gain.ewm(alpha=1.0 / period, min_periods=period, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1.0 / period, min_periods=period, adjust=False).mean()
    g, l = avg_gain.iloc[-1], avg_loss.iloc[-1]
    if pd.isna(g) or pd.isna(l):
        return None
    if l == 0:
        return 100.0
    return float(100.0 - (100.0 / (1.0 + (g / l))))


@st.cache_data(ttl=1800, show_spinner=False)  # cache 30 min so re-runs are fast
def scan(tickers: tuple[str, ...], threshold: float) -> pd.DataFrame:
    data = yf.download(list(tickers), period="1y", interval="1d",
                       group_by="ticker", auto_adjust=False, threads=True, progress=False)
    rows = []
    for symbol in tickers:
        try:
            df = data[symbol] if len(tickers) > 1 else data
            close = df["Close"].dropna()
            if close.empty:
                continue
            rsi = compute_rsi(close)
            if rsi is None or rsi < threshold:
                continue
            name, reco, target = symbol.replace(".NS", ""), "n/a", "n/a"
            try:
                info = yf.Ticker(symbol).info
                name = info.get("shortName") or name
                reco = (info.get("recommendationKey") or "n/a").replace("_", " ").title()
                tp = info.get("targetMeanPrice")
                target = round(float(tp), 2) if tp else "n/a"
                time.sleep(0.2)
            except Exception:
                pass
            rows.append({
                "Date": dt.date.today().isoformat(),
                "Stock Name": f'{name} ({symbol.replace(".NS", "")})',
                "Current Price (₹)": round(float(close.iloc[-1]), 2),
                "52 Week High (₹)": round(float(df["High"].dropna().max()), 2),
                "RSI": round(rsi, 1),
                "Recommendation to buy or sell": reco,
                "1 Year Target (₹)": target,
            })
        except Exception:
            continue
    out = pd.DataFrame(rows, columns=HEADERS)
    if not out.empty:
        out = out.sort_values("RSI", ascending=False).reset_index(drop=True)
    return out


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "RSI Screener"
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    body_font = Font(name="Arial", size=11)
    hot_fill = PatternFill("solid", fgColor="FCE4D6")

    ws.append(HEADERS)
    for cell in ws[1]:
        cell.fill, cell.font = header_fill, header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for _, r in df.iterrows():
        ws.append([r[h] for h in HEADERS])
        i = ws.max_row
        for cell in ws[i]:
            cell.font = body_font
        for col in (3, 4, 7):
            c = ws.cell(row=i, column=col)
            if isinstance(c.value, (int, float)):
                c.number_format = "#,##0.00"
        ws.cell(row=i, column=5).number_format = "0.0"
        if isinstance(r["RSI"], (int, float)) and r["RSI"] >= 70:
            for cell in ws[i]:
                cell.fill = hot_fill

    for i, w in enumerate([12, 34, 16, 16, 8, 26, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{max(ws.max_row, 2)}"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------- UI -------------------------------------

st.set_page_config(page_title="NIFTY RSI Screener", page_icon="📈", layout="wide")
st.title("📈 NIFTY Daily RSI Screener")
st.caption("Stocks in the NIFTY 50 with 14-day RSI at or above your threshold. "
           "Not financial advice — recommendations & targets are Yahoo Finance analyst consensus.")

col1, col2 = st.columns([1, 3])
with col1:
    threshold = st.slider("RSI threshold", min_value=50, max_value=90, value=65, step=1)
    run = st.button("Run screen", type="primary")

if run:
    with st.spinner("Fetching prices and computing RSI..."):
        df = scan(tuple(NIFTY50), float(threshold))
    if df.empty:
        st.info(f"No NIFTY 50 stocks with RSI ≥ {threshold} right now.")
    else:
        st.success(f"{len(df)} stocks flagged (RSI ≥ {threshold}).")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Download Excel report",
            data=to_excel_bytes(df),
            file_name=f"RSI_Screener_{dt.date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.write("Set your threshold and press **Run screen**.")
