"""
NSE Daily RSI Screener - Web App
--------------------------------
Screens NSE stocks for 14-day RSI >= your threshold and shows:

  Date | Stock Name | Market Cap | Current Price | 52 Week High |
  RSI | Recommendation to buy or sell | 1 Year Target

Universe options: NIFTY 50 (fast) or the full NSE equity list (~2000 stocks,
slow). You can also upload your own list of symbols.

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

# Market-cap buckets in INR crore. SEBI's official definition is rank-based
# (top 100 = Large, 101-250 = Mid, 251+ = Small); these absolute thresholds
# are the commonly used retail approximation. Adjust if you prefer.
MCAP_LARGE_CR = 20000   # >= 20,000 cr  -> Large
MCAP_MID_CR = 5000      # 5,000-20,000  -> Mid
MCAP_SMALL_CR = 500     # 500-5,000     -> Small (below 500 -> Micro)

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
    "Date", "Stock Name", "Market Cap", "Current Price (Rs)", "52 Week High (Rs)",
    "RSI", "Recommendation to buy or sell", "1 Year Target (Rs)",
]

NSE_EQUITY_LIST_URL = "https://nsearchives.nseindia.com/content/equities/EQUITY_L.csv"
NIFTY500_LIST_URL = "https://niftyindices.com/IndexConstituent/ind_nifty500list.csv"


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


def classify_market_cap(mcap_inr: float | None) -> str:
    if not mcap_inr or mcap_inr <= 0:
        return "n/a"
    cr = mcap_inr / 1e7  # rupees -> crore
    if cr >= MCAP_LARGE_CR:
        return "Large"
    if cr >= MCAP_MID_CR:
        return "Mid"
    if cr >= MCAP_SMALL_CR:
        return "Small"
    return "Micro"


# ------------------------------ universe ----------------------------------

def _fetch_symbol_csv(url: str) -> list[str]:
    """Download a CSV with a SYMBOL column and return Yahoo (.NS) tickers. [] on failure."""
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
    """Fetch the official NSE equity master list (~2000 stocks). [] on failure."""
    return _fetch_symbol_csv(NSE_EQUITY_LIST_URL)


@st.cache_data(ttl=86400, show_spinner=False)
def load_nifty500_tickers() -> list[str]:
    """Fetch the official NIFTY 500 constituent list. [] on failure."""
    return _fetch_symbol_csv(NIFTY500_LIST_URL)


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
    Best-effort analyst reco, 1y mean target, market cap, and display name.
    Tries .info first (one request covers all), then analyst_price_targets as
    a target fallback. Returns 'n/a' where Yahoo has no data (common for
    small/micro caps with no analyst coverage) rather than failing silently.
    """
    name = symbol.replace(".NS", "")
    reco, target, mcap = "n/a", "n/a", None
    tk = yf.Ticker(symbol)

    try:
        info = tk.info or {}
        name = info.get("shortName") or info.get("longName") or name
        mcap = info.get("marketCap")
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

    if mcap is None:  # fast_info is a lighter path for market cap
        try:
            mcap = getattr(tk.fast_info, "market_cap", None)
        except Exception:
            pass

    return {"name": name, "reco": reco, "target": target, "mcap": mcap}


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
    progress.progress(0.0, text=f"Fetching fundamentals for {total} matched stocks...")
    for idx, (sym, rsi, price, high52) in enumerate(hits):
        if idx < fetch_fund_limit:
            f = fetch_fundamentals(sym)
            time.sleep(0.4)  # be polite; reduces rate-limiting
            progress.progress((idx + 1) / max(total, 1),
                              text=f"Fetched {idx + 1}/{total} fundamentals...")
        else:
            f = {"name": sym.replace(".NS", ""), "reco": "not fetched",
                 "target": "not fetched", "mcap": None}
        rows.append({
            "Date": dt.date.today().isoformat(),
            "Stock Name": f'{f["name"]} ({sym.replace(".NS", "")})',
            "Market Cap": classify_market_cap(f["mcap"]),
            "Current Price (Rs)": round(price, 2),
            "52 Week High (Rs)": round(high52, 2) if high52 else "n/a",
            "RSI": round(rsi, 1),
            "Recommendation to buy or sell": f["reco"],
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

    for i, w in enumerate([12, 34, 11, 16, 16, 8, 26, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{max(ws.max_row, 2)}"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------- UI ------------------------------------

st.set_page_config(page_title="NSE RSI Screener", page_icon="chart", layout="wide")
st.title("NSE Daily RSI Screener")
st.caption("Stocks with 14-day RSI at or above your threshold. Not financial advice - "
           "recommendation and 1-year target are Yahoo Finance analyst consensus and are "
           "often blank for smaller stocks with no analyst coverage.")

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
        "Fetch reco/target/mcap for top N matches", 10, 500, 100, 10,
        help="Each fundamentals lookup is a separate request. Capping this keeps "
             "large scans from timing out or getting rate-limited.")
    run = st.button("Run screen", type="primary")

if universe_choice == "NIFTY 500 (broad)":
    st.info("NIFTY 500 covers ~92% of NSE market cap across large, mid and small caps - "
            "the recommended broad daily scan. Runs in about a minute and most names have "
            "analyst coverage.")
elif universe_choice == "All NSE (~2000, slow)":
    st.info("Full NSE scan pulls ~2000 stocks and can take several minutes. On the free "
            "Streamlit tier it may occasionally time out or get rate-limited by Yahoo - "
            "just press Run again if so. NIFTY 500 is the better broad option for daily use.")

if run:
    if universe_choice == "NIFTY 50 (fast)":
        tickers = NIFTY50
    elif universe_choice == "NIFTY 500 (broad)":
        tickers = load_nifty500_tickers()
        if not tickers:
            st.error("Couldn't fetch the NIFTY 500 list (niftyindices.com sometimes blocks "
                     "cloud servers). Try again in a minute, or use 'Upload my own list'.")
            st.stop()
    elif universe_choice == "All NSE (~2000, slow)":
        tickers = load_all_nse_tickers()
        if not tickers:
            st.error("Couldn't fetch the NSE master list (NSE sometimes blocks cloud servers). "
                     "Use 'Upload my own list' instead - download EQUITY_L.csv from the NSE "
                     "website and paste the SYMBOL column.")
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
        na = int((df["Recommendation to buy or sell"] == "n/a").sum())
        if na:
            st.caption(f"Note: {na} of {len(df)} stocks have no analyst recommendation on "
                       "Yahoo Finance (typical for small/micro caps) and show 'n/a'.")
else:
    st.write("Pick your settings in the sidebar and press **Run screen**.")
