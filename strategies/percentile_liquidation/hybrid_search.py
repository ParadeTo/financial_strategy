"""混合策略网格搜索：A股用分位点清仓，纳指用偏离度阈值，恒生两种都试。

网格变量：
  - A股清仓分位点：90% / 92% / 95% / 97%
  - A股再入分位点：50% / 60% / 70%
  - 恒生：分位点（同A股参数） 或 偏离度（+35%，现行）
  - 纳指：固定用偏离度 +30%（现行）
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from common.data import load_price_cache, BACKTEST_PERIODS, TARGET_NAMES
from strategies.percentile_liquidation.strategy import (
    compute_percentile_series, TargetState,
    BASE_AMOUNTS, TOTAL_PER_PERIOD, PAUSE_TOTAL,
    INCREMENT_PER_PERIOD, RESERVE_INTEREST_ANNUAL,
)

A_LIQUIDATE_PCTS = [0.90, 0.92, 0.95, 0.97]
A_RESUME_PCTS    = [0.50, 0.60, 0.70]
HSI_MODES        = ["percentile", "deviation"]  # 恒生用哪种
NDX_DEV_THRESHOLD = 0.30   # 纳指偏离度阈值（现行）
HSI_DEV_THRESHOLD = 0.35   # 恒生偏离度阈值（现行）
NDX_COOLDOWN_RESUME_DEV = 0.03  # 偏离度模式下的解冻阈值

KEY_LABELS = ["2年A(24-26)", "5年A(21-26)", "10年A(16-26)", "20年(06-26)"]
OPTIMIZE_ON = "20年(06-26)"


def compute_period_hybrid(base, period, price, ma250, holding, state, global_cumulative,
                          frozen_target, price_pct, use_percentile,
                          liquidate_pct, resume_pct,
                          liquidate_dev, resume_dev):
    """
    use_percentile=True  → 用分位点判断清仓和解冻
    use_percentile=False → 用偏离度判断（liquidate_dev / resume_dev）
    """
    target_val = frozen_target if frozen_target is not None else base * period
    deviation = (price - ma250) / ma250 if ma250 != 0 else 0
    paused = global_cumulative >= PAUSE_TOTAL

    if use_percentile:
        should_liquidate = price_pct >= liquidate_pct
        still_cooling    = price_pct > resume_pct
    else:
        should_liquidate = deviation >= liquidate_dev
        still_cooling    = deviation >= resume_dev

    if state == "cooldown":
        if use_percentile:
            note_val = f"分位点{price_pct*100:.1f}%"
        else:
            note_val = f"偏离度{deviation*100:+.1f}%"
        return dict(
            target=target_val, deviation=deviation, price_pct=price_pct,
            regular=0, harvest=0, extra=0, actual=0,
            notes=f"冷却期中（{note_val}，等待回落）",
            state_out="cooldown" if still_cooling else "resume_next",
        )

    if should_liquidate and holding > 0:
        if use_percentile:
            note_val = f"分位点{price_pct*100:.1f}%"
        else:
            note_val = f"偏离度{deviation*100:+.1f}%"
        return dict(
            target=target_val, deviation=deviation, price_pct=price_pct,
            regular=0, harvest=0, extra=0, actual=-holding,
            notes=f"全仓清仓（{note_val}）",
            state_out="cooldown",
        )

    gap = max(target_val - holding, 0)
    excess = holding - target_val
    harvest = -excess if excess > 0 else 0.0
    regular = gap
    actual = regular + harvest

    notes_parts = []
    if paused:
        notes_parts.append("[增量阶段]")
    if state == "resume":
        notes_parts.append("冷却解除，恢复定投")
    elif period == 1 and not paused:
        notes_parts.append("初始买入")
    elif regular == 0 and harvest == 0:
        notes_parts.append("本期无操作")
    else:
        if regular > 0:
            notes_parts.append("跌了多投" if regular > base * 1.1
                               else "涨了少投" if regular < base * 0.9 else "正常投入")
        if harvest < 0:
            notes_parts.append("收割超额")

    return dict(
        target=target_val, deviation=deviation, price_pct=price_pct,
        regular=regular, harvest=harvest, extra=0, actual=actual,
        notes="；".join(notes_parts), state_out="normal",
    )


def run_hybrid_backtest(price_data, pct_data, bt_start, bt_end,
                        a_liq_pct, a_res_pct, hsi_mode):
    # 各标的配置
    cfg = {
        "沪深300 ETF":  dict(use_pct=True,  liq_pct=a_liq_pct, res_pct=a_res_pct, liq_dev=None, res_dev=None),
        "中证500 ETF":  dict(use_pct=True,  liq_pct=a_liq_pct, res_pct=a_res_pct, liq_dev=None, res_dev=None),
        "恒生指数 ETF": dict(use_pct=(hsi_mode=="percentile"),
                           liq_pct=a_liq_pct, res_pct=a_res_pct,
                           liq_dev=HSI_DEV_THRESHOLD, res_dev=NDX_COOLDOWN_RESUME_DEV),
        "纳指100 ETF":  dict(use_pct=False, liq_pct=None, res_pct=None,
                           liq_dev=NDX_DEV_THRESHOLD, res_dev=NDX_COOLDOWN_RESUME_DEV),
    }

    dfs_in_range = {}
    for tname in TARGET_NAMES:
        subset = price_data[tname].loc[bt_start:bt_end]
        if len(subset) == 0:
            return None
        dfs_in_range[tname] = subset

    common_start = max(df.index[0] for df in dfs_in_range.values())
    common_end   = min(df.index[-1] for df in dfs_in_range.values())
    ref_dates    = dfs_in_range[TARGET_NAMES[0]].loc[common_start:common_end].index
    sample_dates = [ref_dates[i] for i in range(0, len(ref_dates), 10)]
    if len(sample_dates) < 2:
        return None

    states = {t: TargetState(BASE_AMOUNTS[i]) for i, t in enumerate(TARGET_NAMES)}
    global_cumulative = 0.0
    reserve_pool = 0.0
    increment_cumulative = 0.0
    increment_deployed = 0.0
    backtest_rows = {t: [] for t in TARGET_NAMES}

    for date in sample_dates:
        paused = global_cumulative >= PAUSE_TOTAL
        if paused:
            reserve_pool += INCREMENT_PER_PERIOD
            increment_cumulative += INCREMENT_PER_PERIOD
            reserve_pool *= 1 + RESERVE_INTEREST_ANNUAL * 14 / 365

        plans = []
        for t_idx, tname in enumerate(TARGET_NAMES):
            ts = states[tname]
            df = price_data[tname]
            date_use = date if date in df.index else \
                df.index[df.index.get_indexer([date], method="nearest")[0]]

            price  = float(df.loc[date_use, "close"])
            ma250  = float(df.loc[date_use, "ma250"])
            holding = ts.shares * price
            ts.period += 1

            pct_series = pct_data[tname]
            price_pct  = float(
                pct_series.loc[date_use] if date_use in pct_series.index
                else pct_series.iloc[pct_series.index.get_indexer([date_use], method="nearest")[0]]
            )
            deviation = (price - ma250) / ma250

            c = cfg[tname]
            use_pct   = c["use_pct"]
            still_cooling = (price_pct > c["res_pct"]) if use_pct else (deviation >= c["res_dev"])

            if ts.state == "cooldown" and still_cooling:
                ts.period_offset = ts.period

            if paused and ts.state not in ("cooldown",):
                effective_prev = (ts.period - 1) - ts.period_offset
                if ts.frozen_target is None:
                    ts.frozen_target = ts.base * max(0, effective_prev)
                ts.frozen_target += ts.base * INCREMENT_PER_PERIOD / TOTAL_PER_PERIOD

            effective_period = ts.period - ts.period_offset
            effective_state  = ts.state
            if ts.state == "cooldown" and not still_cooling:
                effective_state = "resume"

            result = compute_period_hybrid(
                ts.base, effective_period, price, ma250, holding,
                effective_state, global_cumulative,
                frozen_target=ts.frozen_target,
                price_pct=price_pct,
                use_percentile=use_pct,
                liquidate_pct=c["liq_pct"],
                resume_pct=c["res_pct"],
                liquidate_dev=c["liq_dev"],
                resume_dev=c["res_dev"],
            )
            plans.append({"tname": tname, "ts": ts, "date_use": date_use,
                          "price": price, "ma250": ma250, "holding": holding,
                          "result": result})

        if paused:
            active = [(i, p) for i, p in enumerate(plans)
                      if p["result"].get("state_out") not in ("cooldown",)]
            total_reg = sum(p["result"]["regular"] for _, p in active if p["result"]["regular"] > 0)
            if total_reg > reserve_pool and total_reg > 0:
                scale = reserve_pool / total_reg
                for _, p in active:
                    if p["result"]["regular"] > 0:
                        p["result"]["regular"] *= scale
                        p["result"]["actual"] = p["result"]["regular"] + p["result"]["harvest"]

        for p in plans:
            tname   = p["tname"]
            ts      = p["ts"]
            result  = p["result"]
            price   = p["price"]
            holding = p["holding"]
            actual  = result["actual"]

            if result.get("state_out") == "cooldown":
                if holding > 0:
                    reserve_pool += holding
                    ts.shares = 0
                    ts.liquidate_count += 1
                    ts.period_offset = ts.period
                    ts.frozen_target = None
                ts.state = "cooldown"
            else:
                if actual > 0:
                    ts.shares += actual / price
                    ts.cumulative_invested += actual
                    if paused:
                        reserve_pool -= actual
                        increment_deployed += actual
                    else:
                        global_cumulative += actual
                elif actual < 0:
                    sell_amount = abs(actual)
                    ts.shares   = max(0, ts.shares - sell_amount / price)
                    reserve_pool += sell_amount
                ts.state = result.get("state_out", "normal")

            current_value = ts.shares * price
            ts.peak_value = max(ts.peak_value, current_value)
            if ts.peak_value > 0:
                ts.max_drawdown = max(ts.max_drawdown,
                                      (ts.peak_value - current_value) / ts.peak_value)
            ts.harvest_count += 1 if result["harvest"] < 0 else 0
            backtest_rows[tname].append(dict(actual=actual, notes=result["notes"], price=price))

    return dict(states=states, global_cumulative=global_cumulative,
                reserve_pool=reserve_pool,
                increment_cumulative=increment_cumulative,
                backtest_rows=backtest_rows)


def summarize(r):
    states = r["states"]
    rows   = r["backtest_rows"]
    n      = len(rows[TARGET_NAMES[0]])
    years  = n * 14 / 365
    total_holding = sum(states[t].shares * rows[t][-1]["price"] for t in TARGET_NAMES)
    reserve   = r["reserve_pool"]
    user_cash = r["global_cumulative"] + r.get("increment_cumulative", 0)
    net = total_holding + reserve
    ret = (net - user_cash) / user_cash if user_cash > 0 else 0
    ann = (1 + ret) ** (1 / years) - 1 if years > 0 and ret > -1 else 0
    liq = {t: states[t].liquidate_count for t in TARGET_NAMES}
    return dict(ret=ret, ann=ann, liq=liq, total_liq=sum(liq.values()))


if __name__ == "__main__":
    print("Loading cache...")
    price_data = load_price_cache()
    pct_data   = compute_percentile_series(price_data)

    period_map = {label: (s, e) for label, s, e in BACKTEST_PERIODS if label in KEY_LABELS}

    records = []
    for hsi_mode in HSI_MODES:
        for a_liq in A_LIQUIDATE_PCTS:
            for a_res in A_RESUME_PCTS:
                if a_res >= a_liq:
                    continue
                row = {"a_liq": a_liq, "a_res": a_res, "hsi": hsi_mode}
                for label in KEY_LABELS:
                    bt_start, bt_end = period_map[label]
                    r = run_hybrid_backtest(price_data, pct_data, bt_start, bt_end,
                                           a_liq, a_res, hsi_mode)
                    row[label] = summarize(r) if r else None
                records.append(row)

    # 按20年年化排序
    valid = [r for r in records if r.get(OPTIMIZE_ON)]
    valid.sort(key=lambda r: r[OPTIMIZE_ON]["ann"], reverse=True)

    print(f"\n混合策略网格搜索结果（按{OPTIMIZE_ON}年化排序）")
    print(f"纳指固定偏离度+{NDX_DEV_THRESHOLD*100:.0f}%，A股分位点，恒生可选\n")

    # 现行策略基准
    print("── 基准：现行策略（全偏离度 A50/港35/纳30）──")
    print(f"  2年: +28.2%/+13.8%  5年: +39.2%/+7.3%  10年: +51.2%/+4.5%  20年: +60.7%/+4.3%  清仓3次\n")

    print(f"{'A清仓':>7} {'A再入':>7} {'恒生':>8}"
          f"  {'2年总/年/清':>14}  {'5年总/年/清':>14}  {'10年总/年/清':>15}  {'20年总/年/清':>15}")
    print("-" * 100)

    for r in valid[:20]:
        s20 = r[OPTIMIZE_ON]
        line = (f"{r['a_liq']*100:>6.0f}%"
                f" {r['a_res']*100:>6.0f}%"
                f" {r['hsi']:>8}")
        for lb in KEY_LABELS:
            s = r.get(lb)
            if s:
                cell = f"{s['ret']:+.1%}/{s['ann']:+.1%}/{s['total_liq']}清"
                line += f"  {cell:>15}"
            else:
                line += f"  {'N/A':>15}"
        print(line)

    best = valid[0]
    b20  = best[OPTIMIZE_ON]
    print(f"\n最优：A股清仓={best['a_liq']*100:.0f}% 再入={best['a_res']*100:.0f}%"
          f" 恒生={best['hsi']}"
          f" → 20年年化={b20['ann']:+.2%} 总收益={b20['ret']:+.2%} 清仓{b20['total_liq']}次")
    print(f"  各标的清仓：{' '.join(f'{t.split()[0]}={v}' for t,v in b20['liq'].items())}")
