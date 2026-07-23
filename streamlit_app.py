"""
NSE Daily RSI Screener - Web App
--------------------------------
Screens NSE stocks for 14-day RSI >= your threshold and shows:

  Date | Stock Name | Sector | Current Price (Rs) | 52 Week High (Rs) |
  RSI | Buy/Sell | 1 Year Target (Rs)

Universe options: NIFTY 50 (fast), NIFTY 500 (broad), the full NSE equity
list (slow), or your own pasted list of symbols.

Deploy free on Streamlit Community Cloud (share.streamlit.io). Locally:
  pip install -r requirements.txt
  streamlit run streamlit_app.py
"""

from __future__ import annotations

import datetime as dt
import io
import time

import pandas as pd
import requests
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
    "Date", "Stock Name", "Sector", "Current Price (Rs)", "52 Week High (Rs)",
    "RSI", "Buy/Sell", "1 Year Target (Rs)",
]

_EQUITY_LIST_URL = "https://nsearchives.nseindia.com/content/equities/EQUITY_L.csv"
_INDEX500_LIST_URL = "https://niftyindices.com/IndexConstituent/ind_nifty500list.csv"


# ------------------------------ indicators --------------------------------

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


# ------------------------------ universe ----------------------------------

def _fetch_symbol_csv(url: str) -> list[str]:
    """Download a CSV with a SYMBOL column and return .NS tickers. [] on failure."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "text/csv,*/*",
    }
    try:
        r = requests.get(url, headers=headers, timeout=20)
        r.raise_for_status()
        df = pd.read_csv(io.StringIO(r.text))
        col = next(c for c in df.columns if c.strip().upper() == "SYMBOL")
        return [f"{s.strip()}.NS" for s in df[col].dropna().astype(str) if s.strip()]
    except Exception:
        return []


@st.cache_data(ttl=86400, show_spinner=False)
def load_all_nse_tickers() -> list[str]:
    """Full NSE equity list (~2000 stocks). [] on failure."""
    return _fetch_symbol_csv(_EQUITY_LIST_URL)


@st.cache_data(ttl=86400, show_spinner=False)
def load_nifty500_tickers() -> list[str]:
    """NIFTY 500 constituent list. [] on failure."""
    return _fetch_symbol_csv(_INDEX500_LIST_URL)


def parse_uploaded_symbols(text: str) -> list[str]:
    out = []
    for line in text.replace(",", "\n").splitlines():
        s = line.strip().upper()
        if not s or s in ("SYMBOL",):
            continue
        out.append(s if s.endswith(".NS") else f"{s}.NS")
    return out


# ---------------------------- fundamentals --------------------------------

def fetch_fundamentals(symbol: str) -> dict:
    """
    Best-effort sector, consensus rating, 1-year mean target and display name.
    Returns 'n/a' where no data exists (common for smaller, thinly covered
    stocks) rather than failing silently.
    """
    name = symbol.replace(".NS", "")
    sector, reco, target = "n/a", "n/a", "n/a"
    tk = yf.Ticker(symbol)

    try:
        info = tk.info or {}
        name = info.get("shortName") or info.get("longName") or name
        sec = info.get("sector")
        if sec:
            sector = str(sec).strip()
        rk = info.get("recommendationKey")
        if rk and rk.lower() != "none":
            reco = rk.replace("_", " ").title()
        tp = info.get("targetMeanPrice")
        if tp:
            target = round(float(tp), 2)
    except Exception:
        pass

    if target == "n/a":  # fallback to the dedicated targets endpoint
        try:
            apt = tk.analyst_price_targets or {}
            m = apt.get("mean")
            if m:
                target = round(float(m), 2)
        except Exception:
            pass

    return {"name": name, "sector": sector, "reco": reco, "target": target}


@st.cache_data(ttl=1800, show_spinner=False)
def scan(tickers: tuple[str, ...], threshold: float, fetch_fund_limit: int) -> pd.DataFrame:
    """Download history in batches, screen RSI, then enrich the hits."""
    progress = st.progress(0.0, text="Downloading price history...")
    close_map, high_map = {}, {}
    batch = 200
    n = len(tickers)
    for i in range(0, n, batch):
        chunk = list(tickers[i:i + batch])
        data = yf.download(chunk, period="1y", interval="1d", group_by="ticker",
                           auto_adjust=False, threads=True, progress=False)
        for sym in chunk:
            try:
                df = data[sym] if len(chunk) > 1 else data
                c = df["Close"].dropna()
                if c.empty:
                    continue
                close_map[sym] = c
                high_map[sym] = float(df["High"].dropna().max())
            except Exception:
                continue
        progress.progress(min((i + batch) / n, 1.0),
                          text=f"Downloaded {min(i + batch, n)}/{n} stocks...")

    # RSI screen first (cheap), then enrich only the hits (expensive)
    hits = []
    for sym, c in close_map.items():
        rsi = compute_rsi(c)
        if rsi is not None and rsi >= threshold:
            hits.append((sym, rsi, float(c.iloc[-1]), high_map.get(sym)))
    hits.sort(key=lambda x: x[1], reverse=True)

    rows = []
    total = min(len(hits), fetch_fund_limit)
    progress.progress(0.0, text=f"Fetching details for {total} matched stocks...")
    for idx, (sym, rsi, price, high52) in enumerate(hits):
        if idx < fetch_fund_limit:
            f = fetch_fundamentals(sym)
            time.sleep(0.4)  # be polite; reduces rate-limiting
            progress.progress((idx + 1) / max(total, 1),
                              text=f"Fetched {idx + 1}/{total} details...")
        else:
            f = {"name": sym.replace(".NS", ""), "sector": "not fetched",
                 "reco": "not fetched", "target": "not fetched"}
        rows.append({
            "Date": dt.date.today().isoformat(),
            "Stock Name": f'{f["name"]} ({sym.replace(".NS", "")})',
            "Sector": f["sector"],
            "Current Price (Rs)": round(price, 2),
            "52 Week High (Rs)": round(high52, 2) if high52 else "n/a",
            "RSI": round(rsi, 1),
            "Buy/Sell": f["reco"],
            "1 Year Target (Rs)": f["target"],
        })
    progress.empty()
    return pd.DataFrame(rows, columns=HEADERS)


# -------------------------------- excel -----------------------------------

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

    num_cols = {4, 5, 8}  # price, 52w high, target
    for _, r in df.iterrows():
        ws.append([r[h] for h in HEADERS])
        i = ws.max_row
        for cell in ws[i]:
            cell.font = body_font
        for col in num_cols:
            c = ws.cell(row=i, column=col)
            if isinstance(c.value, (int, float)):
                c.number_format = "#,##0.00"
        ws.cell(row=i, column=6).number_format = "0.0"  # RSI
        if isinstance(r["RSI"], (int, float)) and r["RSI"] >= 70:
            for cell in ws[i]:
                cell.fill = hot_fill

    for i, w in enumerate([12, 34, 22, 16, 16, 8, 14, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{max(ws.max_row, 2)}"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------- UI ------------------------------------

st.set_page_config(page_title="NSE RSI Screener", page_icon="chart", layout="wide")
st.title("NSE Daily RSI Screener")
st.caption("Stocks with 14-day RSI at or above your threshold. For information only - "
           "not investment advice. Buy/Sell and 1-year target are third-party analyst "
           "consensus figures, not the view of this app, and are often blank for smaller stocks.")

with st.sidebar:
    st.header("Settings")
    threshold = st.slider("RSI threshold", 50, 90, 65, 1)
    universe_choice = st.radio(
        "Universe",
        ["NIFTY 50 (fast)", "NIFTY 500 (broad)", "All NSE (~2000, slow)", "Upload my own list"],
    )
    uploaded = None
    if universe_choice == "Upload my own list":
        uploaded = st.text_area("Paste symbols (one per line, e.g. RELIANCE or RELIANCE.NS)")
    fetch_limit = st.number_input(
        "Fetch sector/rating/target for top N matches", 10, 500, 100, 10,
        help="Each lookup is a separate request. Capping this keeps large scans "
             "from timing out.")
    run = st.button("Run screen", type="primary")

if universe_choice == "NIFTY 500 (broad)":
    st.info("NIFTY 500 covers ~92% of NSE market cap across large, mid and small caps - "
            "the recommended broad daily scan. Runs in about a minute.")
elif universe_choice == "All NSE (~2000, slow)":
    st.info("Full NSE scan pulls ~2000 stocks and can take several minutes. On the free "
            "tier it may occasionally time out - just press Run again if so. "
            "NIFTY 500 is the better broad option for daily use.")

if run:
    if universe_choice == "NIFTY 50 (fast)":
        tickers = NIFTY50
    elif universe_choice == "NIFTY 500 (broad)":
        tickers = load_nifty500_tickers()
        if not tickers:
            st.error("Couldn't load the NIFTY 500 list right now. Try again in a minute, "
                     "or use 'Upload my own list'.")
            st.stop()
    elif universe_choice == "All NSE (~2000, slow)":
        tickers = load_all_nse_tickers()
        if not tickers:
            st.error("Couldn't load the full NSE list right now. Try again in a minute, "
                     "or use 'Upload my own list'.")
            st.stop()
    else:
        tickers = parse_uploaded_symbols(uploaded or "")
        if not tickers:
            st.warning("Paste at least one symbol first.")
            st.stop()

    df = scan(tuple(tickers), float(threshold), int(fetch_limit))
    if df.empty:
        st.info(f"No stocks with RSI >= {threshold} right now.")
    else:
        st.success(f"{len(df)} stocks flagged (RSI >= {threshold}), scanned {len(tickers)} names.")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button(
            "Download Excel report",
            data=to_excel_bytes(df),
            file_name=f"RSI_Screener_{dt.date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        na = int((df["Buy/Sell"] == "n/a").sum())
        if na:
            st.caption(f"Note: {na} of {len(df)} stocks have no published analyst rating "
                       "(typical for smaller companies) and show 'n/a'.")
else:
    st.write("Pick your settings in the sidebar and press **Run screen**.")
