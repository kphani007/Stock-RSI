"""
NSE Daily RSI Screener - Web App
--------------------------------
Screens NSE stocks for 14-day RSI >= your threshold and shows:

  Date | Stock Symbol | Sector | Current Price (Rs) | 52 Week High (Rs) |
  RSI | Buy/Sell | 1 Year Target (Rs)

Click any row to open a detail panel with market cap, support/resistance
levels and basic financial quality checks.

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

# Market-cap buckets in INR crore (common retail approximation; SEBI's own
# definition is rank-based - top 100 Large, 101-250 Mid, 251+ Small).
MCAP_LARGE_CR, MCAP_MID_CR, MCAP_SMALL_CR = 20000, 5000, 500

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
    "Date", "Stock Symbol", "Sector", "Current Price (Rs)", "52 Week High (Rs)",
    "RSI", "Buy/Sell", "1 Year Target (Rs)",
]

_EQUITY_LIST_URL = "https://nsearchives.nseindia.com/content/equities/EQUITY_L.csv"
_INDEX500_LIST_URL = "https://niftyindices.com/IndexConstituent/ind_nifty500list.csv"

DISCLAIMER = (
    "**Disclaimer** — Everything shown here is for reference and general "
    "information purposes only. It is not investment advice, not a "
    "recommendation to buy or sell any security, and not a solicitation of any "
    "kind. The Buy/Sell rating and 1-year target are third-party analyst "
    "consensus figures reproduced as-is, not the view of this app. Technical "
    "levels and financial checks are mechanically computed and may contain "
    "errors or use stale data. Do your own research and consult a "
    "SEBI-registered adviser before making any investment decision. "
    "Securities investments are subject to market risk, including loss of capital."
)


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
    cr = mcap_inr / 1e7
    if cr >= MCAP_LARGE_CR:
        return "Large cap"
    if cr >= MCAP_MID_CR:
        return "Mid cap"
    if cr >= MCAP_SMALL_CR:
        return "Small cap"
    return "Micro cap"


def pivot_levels(h: float, l: float, c: float) -> dict:
    """Classic floor-trader pivot points from a completed period's H/L/C."""
    p = (h + l + c) / 3.0
    return {
        "R3": h + 2 * (p - l), "R2": p + (h - l), "R1": 2 * p - l,
        "Pivot": p,
        "S1": 2 * p - h, "S2": p - (h - l), "S3": l - 2 * (h - p),
    }


def swing_levels(close: pd.Series, price: float, window: int = 10):
    """Nearest swing low below and swing high above the current price."""
    s = close.dropna()
    highs = s[(s.shift(window) < s) & (s.shift(-window) < s)]
    lows = s[(s.shift(window) > s) & (s.shift(-window) > s)]
    sup = lows[lows < price]
    res = highs[highs > price]
    return (float(sup.max()) if len(sup) else None,
            float(res.min()) if len(res) else None)


# ---------------------------- financial checks ----------------------------

def _find_row(fin, *keys):
    if fin is None or getattr(fin, "empty", True):
        return None
    for k in keys:
        for idx in fin.index:
            if k.lower() in str(idx).lower():
                return fin.loc[idx]
    return None


def _latest(series, n=1):
    if series is None:
        return None
    s = pd.Series(series).dropna()
    if s.empty:
        return None
    s = s.sort_index(ascending=False)
    return float(s.iloc[0]) if n == 1 else s.head(n).astype(float)


def quality_checks(income, balance, cash) -> list[tuple[str, str, str]]:
    """Basic earnings-quality / leverage screens. Returns (metric, value, flag)."""
    out = []
    ni = _find_row(income, "Net Income")
    cfo = _find_row(cash, "Operating Cash Flow", "Total Cash From Operating")
    if ni is not None and cfo is not None:
        n3, c3 = _latest(ni, 3), _latest(cfo, 3)
        if n3 is not None and c3 is not None and len(n3) and n3.sum() != 0:
            ratio = c3.sum() / n3.sum()
            out.append(("Cash flow vs profit (3y CFO/PAT)", f"{ratio:.2f}x",
                        "OK" if ratio >= 0.8 else "Watch"))
    td = _find_row(balance, "Total Debt")
    eq = _find_row(balance, "Stockholders Equity", "Total Equity")
    d, e = _latest(td), _latest(eq)
    if d is not None and e:
        de = d / e
        out.append(("Debt to equity", f"{de:.2f}x", "OK" if de <= 1.5 else "Watch"))
    ebit = _find_row(income, "EBIT", "Operating Income")
    intr = _find_row(income, "Interest Expense")
    b, i = _latest(ebit), _latest(intr)
    if b is not None and i:
        cov = b / abs(i)
        out.append(("Interest coverage", f"{cov:.1f}x", "OK" if cov >= 3 else "Watch"))
    rev = _find_row(income, "Total Revenue")
    rec = _find_row(balance, "Receivables", "Accounts Receivable")
    r2, q2 = _latest(rev, 2), _latest(rec, 2)
    if (r2 is not None and q2 is not None and len(r2) == 2 and len(q2) == 2
            and r2.iloc[1] and q2.iloc[1]):
        rg = r2.iloc[0] / r2.iloc[1] - 1
        qg = q2.iloc[0] / q2.iloc[1] - 1
        out.append(("Receivables growth vs sales", f"{qg*100:.0f}% vs {rg*100:.0f}%",
                    "OK" if qg <= rg + 0.15 else "Watch"))
    return out


# ------------------------------ universe ----------------------------------

def _fetch_symbol_csv(url: str) -> list[str]:
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
               "Accept": "text/csv,*/*"}
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
    return _fetch_symbol_csv(_EQUITY_LIST_URL)


@st.cache_data(ttl=86400, show_spinner=False)
def load_nifty500_tickers() -> list[str]:
    return _fetch_symbol_csv(_INDEX500_LIST_URL)


def parse_uploaded_symbols(text: str) -> list[str]:
    out = []
    for line in text.replace(",", "\n").splitlines():
        s = line.strip().upper()
        if not s or s in ("SYMBOL",):
            continue
        out.append(s if s.endswith(".NS") else f"{s}.NS")
    return out


# ---------------------------- data fetching -------------------------------

def fetch_fundamentals(symbol: str) -> dict:
    """Best-effort sector, consensus rating, 1-year mean target, name."""
    name = symbol.replace(".NS", "")
    sector, reco, target = "n/a", "n/a", "n/a"
    tk = yf.Ticker(symbol)
    try:
        info = tk.info or {}
        name = info.get("shortName") or info.get("longName") or name
        if info.get("sector"):
            sector = str(info["sector"]).strip()
        rk = info.get("recommendationKey")
        if rk and rk.lower() != "none":
            reco = rk.replace("_", " ").title()
        tp = info.get("targetMeanPrice")
        if tp:
            target = round(float(tp), 2)
    except Exception:
        pass
    if target == "n/a":
        try:
            m = (tk.analyst_price_targets or {}).get("mean")
            if m:
                target = round(float(m), 2)
        except Exception:
            pass
    return {"name": name, "sector": sector, "reco": reco, "target": target}


@st.cache_data(ttl=1800, show_spinner=False)
def fetch_detail(symbol: str) -> dict:
    """Everything the detail panel needs for a single stock."""
    out = {"name": symbol.replace(".NS", ""), "mcap": None, "price": None,
           "high52": None, "low52": None, "rsi": None, "levels": {},
           "swing": (None, None), "checks": [], "error": None}
    try:
        tk = yf.Ticker(symbol)
        hist = tk.history(period="2y", interval="1d", auto_adjust=False)
        if hist is None or hist.empty:
            out["error"] = "No price history available."
            return out
        close = hist["Close"].dropna()
        out["price"] = float(close.iloc[-1])
        out["rsi"] = compute_rsi(close)
        last_year = hist.tail(252)
        out["high52"] = float(last_year["High"].max())
        out["low52"] = float(last_year["Low"].min())

        # pivots from the last completed calendar month
        monthly = hist.resample("ME").agg({"High": "max", "Low": "min", "Close": "last"})
        monthly = monthly.dropna()
        if len(monthly) >= 2:
            prev = monthly.iloc[-2]
            out["levels"] = pivot_levels(float(prev["High"]), float(prev["Low"]),
                                         float(prev["Close"]))
        out["swing"] = swing_levels(close.tail(180), out["price"])

        try:
            info = tk.info or {}
            out["name"] = info.get("shortName") or info.get("longName") or out["name"]
            out["mcap"] = info.get("marketCap")
        except Exception:
            pass
        try:
            out["checks"] = quality_checks(tk.income_stmt, tk.balance_sheet, tk.cashflow)
        except Exception:
            pass
    except Exception as exc:
        out["error"] = f"Could not load details ({type(exc).__name__})."
    return out


@st.cache_data(ttl=1800, show_spinner=False)
def scan(tickers: tuple[str, ...], threshold: float, fetch_fund_limit: int) -> pd.DataFrame:
    progress = st.progress(0.0, text="Downloading price history...")
    close_map, high_map = {}, {}
    batch, n = 200, len(tickers)
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
            time.sleep(0.4)
            progress.progress((idx + 1) / max(total, 1),
                              text=f"Fetched {idx + 1}/{total} details...")
        else:
            f = {"name": sym.replace(".NS", ""), "sector": "not fetched",
                 "reco": "not fetched", "target": "not fetched"}
        rows.append({
            "Date": dt.date.today().isoformat(),
            "Stock Symbol": sym.replace(".NS", ""),
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

    for _, r in df.iterrows():
        ws.append([r[h] for h in HEADERS])
        i = ws.max_row
        for cell in ws[i]:
            cell.font = body_font
        for col in (4, 5, 8):
            c = ws.cell(row=i, column=col)
            if isinstance(c.value, (int, float)):
                c.number_format = "#,##0.00"
        ws.cell(row=i, column=6).number_format = "0.0"
        if isinstance(r["RSI"], (int, float)) and r["RSI"] >= 70:
            for cell in ws[i]:
                cell.fill = hot_fill

    for i, w in enumerate([12, 16, 22, 16, 16, 8, 14, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{max(ws.max_row, 2)}"

    row = ws.max_row + 2
    ws.cell(row=row, column=1, value=DISCLAIMER.replace("**Disclaimer** — ", "Disclaimer: ")
            ).font = Font(name="Arial", italic=True, size=9)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ------------------------------ detail panel ------------------------------

def render_detail(symbol: str):
    d = fetch_detail(f"{symbol}.NS")
    st.subheader(f"{symbol} — {d['name']}")
    if d["error"]:
        st.warning(d["error"])
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Current price", f"Rs {d['price']:,.2f}" if d["price"] else "n/a")
    if d["mcap"]:
        c2.metric("Market cap", f"Rs {d['mcap']/1e7:,.0f} Cr", classify_market_cap(d["mcap"]))
    else:
        c2.metric("Market cap", "n/a")
    c3.metric("RSI (14d)", f"{d['rsi']:.1f}" if d["rsi"] else "n/a")
    c4.metric("52w range", f"{d['low52']:,.0f} - {d['high52']:,.0f}"
              if d["high52"] else "n/a")

    left, right = st.columns(2)

    with left:
        st.markdown("**Support & resistance**")
        if d["levels"]:
            lv = d["levels"]
            rows = [{"Level": k, "Price (Rs)": round(lv[k], 2)}
                    for k in ["R3", "R2", "R1", "Pivot", "S1", "S2", "S3"]]
            st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
            st.caption("Monthly floor-trader pivot points, from last completed month.")
        sup, res = d["swing"]
        st.write(f"Nearest swing support: **{f'Rs {sup:,.2f}' if sup else 'none found'}**")
        st.write(f"Nearest swing resistance: **{f'Rs {res:,.2f}' if res else 'none found'}**")
        st.caption("Swing levels from local highs/lows over the last ~6 months.")

    with right:
        st.markdown("**Financial quality checks**")
        if d["checks"]:
            chk = pd.DataFrame(d["checks"], columns=["Check", "Value", "Flag"])
            st.dataframe(chk, hide_index=True, use_container_width=True)
            st.caption("'Watch' flags a metric worth investigating - it is not a verdict "
                       "on the company. Ratios are unreliable for banks and NBFCs.")
        else:
            st.write("No financial statement data available for this stock.")
        st.caption("Not covered here: promoter pledging, auditor changes, related-party "
                   "transactions, contingent liabilities.")


# ---------------------------------- UI ------------------------------------

st.set_page_config(page_title="NSE RSI Screener", page_icon="chart", layout="wide")
st.title("NSE Daily RSI Screener")
st.caption("Stocks with 14-day RSI at or above your threshold. Click any row for details.")

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
            "tier it may occasionally time out - just press Run again if so.")

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
    st.session_state["results"] = scan(tuple(tickers), float(threshold), int(fetch_limit))
    st.session_state["scanned"] = len(tickers)
    st.session_state["threshold"] = threshold

results = st.session_state.get("results")

if results is None:
    st.write("Pick your settings in the sidebar and press **Run screen**.")
elif results.empty:
    st.info(f"No stocks with RSI >= {st.session_state.get('threshold', threshold)} right now.")
else:
    st.success(f"{len(results)} stocks flagged, scanned {st.session_state.get('scanned', 0)} names. "
               "Click a row to see details.")
    event = st.dataframe(
        results, use_container_width=True, hide_index=True,
        on_select="rerun", selection_mode="single-row", key="screener_table",
    )
    st.download_button(
        "Download Excel report",
        data=to_excel_bytes(results),
        file_name=f"RSI_Screener_{dt.date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    na = int((results["Buy/Sell"] == "n/a").sum())
    if na:
        st.caption(f"Note: {na} of {len(results)} stocks have no published analyst rating "
                   "(typical for smaller companies) and show 'n/a'.")

    picked = []
    try:
        picked = list(event.selection.rows)
    except Exception:
        pass
    st.divider()
    if picked:
        render_detail(str(results.iloc[picked[0]]["Stock Symbol"]))
    else:
        choice = st.selectbox("...or pick a symbol here",
                              ["-"] + results["Stock Symbol"].tolist())
        if choice != "-":
            render_detail(choice)

st.divider()
st.caption(DISCLAIMER)
