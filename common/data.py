"""Shared data download, MA250 calculation, and caching utilities."""

import os
import pickle
import yfinance as yf
import pandas as pd

TARGET_NAMES = ["沪深300 ETF", "中证500 ETF", "恒生指数 ETF", "纳指100 ETF"]

YFINANCE_TICKERS = {
    "沪深300 ETF": "510300.SS",
    "中证500 ETF": "510500.SS",
    "恒生指数 ETF": "159920.SZ",
    "纳指100 ETF": "513100.SS",
}

YFINANCE_INDEX_TICKERS = {
    "沪深300 ETF": "000300.SS",
    "中证500 ETF": "000905.SS",
    "恒生指数 ETF": "^HSI",
    "纳指100 ETF": "^NDX",
}

DATA_START = "2005-01-01"
DATA_END = "2026-04-23"

BACKTEST_PERIODS = [
    ("2年A(24-26)", "2024-04-01", "2026-04-23"),
    ("2年B(22-24)", "2022-04-01", "2024-04-01"),
    ("2年C(20-22)", "2020-04-01", "2022-04-01"),
    ("2年D(18-20)", "2018-04-01", "2020-04-01"),
    ("5年A(21-26)", "2021-04-01", "2026-04-23"),
    ("5年B(19-24)", "2019-04-01", "2024-04-01"),
    ("5年C(17-22)", "2017-04-01", "2022-04-01"),
    ("10年A(16-26)", "2016-04-01", "2026-04-23"),
    ("10年B(14-24)", "2014-04-01", "2024-04-01"),
    ("20年(06-26)", "2006-04-01", "2026-04-23"),
]

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_DIR = os.path.join(PROJECT_ROOT, "cache")


def download_price_data():
    """Download ETF + index data, merge, compute MA250. Returns dict of DataFrames.

    Strategy per ETF:
    - 沪深300 / 中证500: Use ETF market price directly (no QDII premium, index
      tickers 000300.SS / 000905.SS unavailable on yfinance).  Use scaled index
      only for dates before ETF listing.
    - 恒生指数: Use ETF market price directly (no significant QDII premium).
      Use scaled ^HSI for pre-listing dates.
    - 纳指100: Use ^NDX scaled to ETF price at first overlap, for the *entire*
      date range.  The ETF (513100.SS) exhibited large QDII premiums in 2020
      (up to 28% above NAV) which would distort the backtest.
    """
    price_data = {}
    for tname, ticker in YFINANCE_TICKERS.items():
        df_etf = yf.download(ticker, start=DATA_START, end=DATA_END, progress=False)
        if isinstance(df_etf.columns, pd.MultiIndex):
            df_etf.columns = df_etf.columns.get_level_values(0)
        df_etf = df_etf[["Close"]].copy()
        df_etf.columns = ["close"]

        idx_ticker = YFINANCE_INDEX_TICKERS[tname]
        df_idx = yf.download(idx_ticker, start=DATA_START, end=DATA_END, progress=False,
                             auto_adjust=True)
        if isinstance(df_idx.columns, pd.MultiIndex):
            df_idx.columns = df_idx.columns.get_level_values(0)
        df_idx = df_idx[["Close"]].copy()
        df_idx.columns = ["close"]

        overlap = df_etf.index.intersection(df_idx.index)

        if tname == "纳指100 ETF" and len(overlap) > 0:
            # Use scaled index for the full range to avoid QDII premium distortion.
            ratio = df_etf.loc[overlap[0], "close"] / df_idx.loc[overlap[0], "close"]
            df = df_idx.copy()
            df["close"] = df["close"] * ratio
        elif len(overlap) > 0:
            # Use ETF price where available; fill pre-listing dates with scaled index.
            ratio = df_etf.loc[overlap[0], "close"] / df_idx.loc[overlap[0], "close"]
            early = df_idx.loc[df_idx.index < df_etf.index[0]].copy()
            if len(early) > 0:
                early["close"] = early["close"] * ratio
                df = pd.concat([early, df_etf])
            else:
                df = df_etf
        else:
            df = df_etf

        df["ma250"] = df["close"].rolling(window=250).mean()
        df = df.dropna(subset=["ma250"])
        price_data[tname] = df
        print(f"  {tname}: {len(df)} rows with MA250, "
              f"from {df.index[0].date()} to {df.index[-1].date()}")

    return price_data


def save_price_cache(price_data):
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(os.path.join(CACHE_DIR, "price_data.pkl"), "wb") as f:
        pickle.dump(price_data, f)


def load_price_cache():
    path = os.path.join(CACHE_DIR, "price_data.pkl")
    with open(path, "rb") as f:
        return pickle.load(f)


def save_backtest_cache(all_results, tag=""):
    os.makedirs(CACHE_DIR, exist_ok=True)
    filename = f"backtest_results{tag}.pkl"
    with open(os.path.join(CACHE_DIR, filename), "wb") as f:
        pickle.dump(all_results, f)


def load_backtest_cache(tag=""):
    filename = f"backtest_results{tag}.pkl"
    with open(os.path.join(CACHE_DIR, filename), "rb") as f:
        return pickle.load(f)
