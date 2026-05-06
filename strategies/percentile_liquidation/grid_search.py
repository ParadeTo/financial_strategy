"""网格搜索最优 (清仓分位点, 再入分位点) 组合。

在 4 个关键回测周期上跑 25 组参数，按 20 年年化收益排序输出。
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from common.data import load_price_cache, BACKTEST_PERIODS, TARGET_NAMES
from strategies.percentile_liquidation.strategy import (
    compute_percentile_series, run_backtest, summarize,
)

LIQUIDATE_PCTS = [0.80, 0.85, 0.90, 0.92, 0.95]
RESUME_PCTS    = [0.30, 0.40, 0.50, 0.60, 0.70]

KEY_LABELS = ["2年A(24-26)", "5年A(21-26)", "10年A(16-26)", "20年(06-26)"]
OPTIMIZE_ON = "20年(06-26)"


def run_grid(price_data, pct_data):
    period_map = {label: (bt_start, bt_end)
                  for label, bt_start, bt_end in BACKTEST_PERIODS
                  if label in KEY_LABELS}

    combos = [(lp, rp) for lp in LIQUIDATE_PCTS for rp in RESUME_PCTS if rp < lp]
    records = []

    for liq_pct, res_pct in combos:
        row = {"liq_pct": liq_pct, "res_pct": res_pct}
        for label in KEY_LABELS:
            if label not in period_map:
                row[label] = None
                continue
            bt_start, bt_end = period_map[label]
            r = run_backtest(price_data, pct_data, bt_start, bt_end, liq_pct, res_pct)
            row[label] = summarize(r) if r is not None else None
        records.append(row)

    return records


def print_results(records):
    valid = [r for r in records if r.get(OPTIMIZE_ON) is not None]
    valid.sort(key=lambda r: r[OPTIMIZE_ON]["ann"], reverse=True)

    col_w = 14
    header = f"{'清仓%':>6} {'再入%':>6}"
    for lb in KEY_LABELS:
        header += f"  {lb[:6]:>{col_w}}"
    print(header)
    print("-" * (14 + col_w * len(KEY_LABELS) + 2 * len(KEY_LABELS)))

    for r in valid:
        line = f"{r['liq_pct']*100:>5.0f}% {r['res_pct']*100:>5.0f}%"
        for lb in KEY_LABELS:
            s = r.get(lb)
            if s is None:
                line += f"  {'N/A':>{col_w}}"
            else:
                cell = f"{s['ret']:+.1%}/{s['ann']:+.1%}({s['liquidates']}清)"
                line += f"  {cell:>{col_w}}"
        print(line)

    best = valid[0]
    print(f"\n最优组合（按{OPTIMIZE_ON}年化）："
          f"清仓={best['liq_pct']*100:.0f}%  再入={best['res_pct']*100:.0f}%  "
          f"年化={best[OPTIMIZE_ON]['ann']:+.2%}  "
          f"总收益={best[OPTIMIZE_ON]['ret']:+.2%}  "
          f"清仓{best[OPTIMIZE_ON]['liquidates']}次")


if __name__ == "__main__":
    print("Loading cache...")
    price_data = load_price_cache()
    print("Computing percentile series...")
    pct_data = compute_percentile_series(price_data)
    combos = [(lp, rp) for lp in LIQUIDATE_PCTS for rp in RESUME_PCTS if rp < lp]
    print(f"Running {len(combos)} combinations across {len(KEY_LABELS)} periods...\n")
    records = run_grid(price_data, pct_data)
    print_results(records)
