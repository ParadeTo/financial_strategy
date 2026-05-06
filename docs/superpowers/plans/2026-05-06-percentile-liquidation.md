# 分位点清仓策略 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 新建 `strategies/percentile_liquidation/` 策略变体，用价格扩展窗口分位点替换偏离度阈值，通过网格搜索找到最优 (清仓分位点, 再入分位点) 组合。

**Architecture:** 复制 `constant_value/strategy.py` 的核心引擎，增加 `compute_percentile_series()` 预计算各 ETF 的扩展窗口分位点序列，`compute_period()` 和 `run_backtest()` 接收 `price_pct` / `liquidate_pct` / `resume_pct` 替代原偏离度阈值参数。`grid_search.py` 遍历 25 组组合，输出按 20 年年化排序的结果表。

**Tech Stack:** Python 3, pandas (`expanding().rank(pct=True)`), 复用 `cache/price_data.pkl`

---

## 文件结构

| 文件 | 操作 | 职责 |
|---|---|---|
| `strategies/percentile_liquidation/__init__.py` | Create | 空文件，使目录成为 package |
| `strategies/percentile_liquidation/strategy.py` | Create | 恒定市值引擎 + 分位点触发逻辑 |
| `strategies/percentile_liquidation/grid_search.py` | Create | 网格搜索脚本，输出最优参数组合 |

---

## Task 1: 创建分支和目录

**Files:**
- Create: `strategies/percentile_liquidation/__init__.py`

- [ ] **Step 1: 创建分支**

```bash
git checkout -b feature/percentile-liquidation
```

- [ ] **Step 2: 创建目录和 __init__.py**

```bash
mkdir -p strategies/percentile_liquidation
touch strategies/percentile_liquidation/__init__.py
```

- [ ] **Step 3: Commit**

```bash
git add strategies/percentile_liquidation/__init__.py
git commit -m "feat: scaffold percentile-liquidation strategy directory"
```

---

## Task 2: 实现 strategy.py — 分位点计算 + 核心引擎

**Files:**
- Create: `strategies/percentile_liquidation/strategy.py`

- [ ] **Step 1: 创建 strategy.py，写入完整内容**

`strategies/percentile_liquidation/strategy.py` 完整内容如下：

```python
"""分位点清仓策略：用价格历史分位点替代偏离度阈值触发清仓/解冻。

策略要点：
- 恒定市值法买卖逻辑不变
- 清仓触发：当前价格的扩展窗口分位点 >= liquidate_pct
- 冷却解除：当前价格的扩展窗口分位点 <= resume_pct
- 分位点 = 数据起点到当日所有历史价格中，≤ 当前价格的比例
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

import pandas as pd
from common.data import (
    TARGET_NAMES, BACKTEST_PERIODS,
    load_price_cache, save_backtest_cache,
)

# ── Strategy Config（与 constant_value 保持一致）──────────────
BASE_AMOUNTS = [1500, 1500, 600, 2400]
TOTAL_PER_PERIOD = 6000
PAUSE_TOTAL = 150000
INCREMENT_PER_PERIOD = 2000
RESERVE_INTEREST_ANNUAL = 0.01
LARGE_INVEST_MULT = 2.5

# 默认参数（grid_search 会覆盖）
DEFAULT_LIQUIDATE_PCT = 0.90
DEFAULT_RESUME_PCT = 0.50


# ═══════════════════════════════════════════════════════════
# 分位点预计算
# ═══════════════════════════════════════════════════════════
def compute_percentile_series(price_data):
    """预计算各 ETF 的扩展窗口价格分位点序列。

    返回 dict: {tname: pd.Series}，index 与 price_data[tname] 相同，
    值为当日价格在 [数据起点, 当日] 所有收盘价中的百分位排名（0~1）。
    """
    pct_data = {}
    for tname, df in price_data.items():
        pct_data[tname] = df['close'].expanding().rank(pct=True)
    return pct_data


# ═══════════════════════════════════════════════════════════
# DCA Engine
# ═══════════════════════════════════════════════════════════
def compute_period(base, period, price, ma250, holding, state, global_cumulative,
                   frozen_target=None, price_pct=0.5,
                   liquidate_pct=DEFAULT_LIQUIDATE_PCT,
                   resume_pct=DEFAULT_RESUME_PCT):
    target_val = frozen_target if frozen_target is not None else base * period
    deviation = (price - ma250) / ma250 if ma250 != 0 else 0
    paused = global_cumulative >= PAUSE_TOTAL

    if state == "cooldown":
        return dict(
            target=target_val, deviation=deviation, price_pct=price_pct,
            regular=0, harvest=0, extra=0, actual=0,
            notes="冷却期中（分位点{:.1f}%，等待回落）".format(price_pct * 100),
            state_out="cooldown" if price_pct > resume_pct else "resume_next",
        )

    # 全仓清仓
    if price_pct >= liquidate_pct and holding > 0:
        return dict(
            target=target_val, deviation=deviation, price_pct=price_pct,
            regular=0, harvest=0, extra=0, actual=-holding,
            notes="全仓清仓（分位点{:.1f}%）".format(price_pct * 100),
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
        notes_parts.append("冷却解除（分位点{:.1f}%），恢复定投".format(price_pct * 100))
    elif period == 1 and not paused:
        notes_parts.append("初始买入")
    elif regular == 0 and harvest == 0:
        notes_parts.append("本期无操作")
    else:
        if regular > 0:
            if regular < base * 0.9:
                notes_parts.append("涨了少投")
            elif regular > base * 1.1:
                notes_parts.append("跌了多投")
            else:
                notes_parts.append("正常投入")
        if harvest < 0:
            notes_parts.append("收割超额")

    return dict(
        target=target_val, deviation=deviation, price_pct=price_pct,
        regular=regular, harvest=harvest, extra=0, actual=actual,
        notes="；".join(notes_parts), state_out="normal",
    )


class TargetState:
    def __init__(self, base):
        self.base = base
        self.shares = 0.0
        self.cumulative_invested = 0.0
        self.state = "normal"
        self.period = 0
        self.harvest_count = 0
        self.liquidate_count = 0
        self.peak_value = 0.0
        self.max_drawdown = 0.0
        self.frozen_target = None
        self.period_offset = 0


def run_backtest(price_data, pct_data, bt_start, bt_end,
                 liquidate_pct=DEFAULT_LIQUIDATE_PCT,
                 resume_pct=DEFAULT_RESUME_PCT):
    dfs_in_range = {}
    for tname in TARGET_NAMES:
        df = price_data[tname]
        subset = df.loc[bt_start:bt_end]
        if len(subset) == 0:
            return None
        dfs_in_range[tname] = subset

    common_start = max(df.index[0] for df in dfs_in_range.values())
    common_end = min(df.index[-1] for df in dfs_in_range.values())
    ref_dates = dfs_in_range[TARGET_NAMES[0]].loc[common_start:common_end].index
    sample_dates = [ref_dates[i] for i in range(0, len(ref_dates), 10)]

    if len(sample_dates) < 2:
        return None

    states = {tname: TargetState(BASE_AMOUNTS[i]) for i, tname in enumerate(TARGET_NAMES)}
    global_cumulative = 0.0
    reserve_pool = 0.0
    increment_cumulative = 0.0
    increment_deployed = 0.0
    backtest_rows = {tname: [] for tname in TARGET_NAMES}

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

            price = float(df.loc[date_use, "close"])
            ma250 = float(df.loc[date_use, "ma250"])
            holding = ts.shares * price
            ts.period += 1

            # 当日分位点（扩展窗口，已预计算）
            pct_series = pct_data[tname]
            price_pct = float(pct_series.loc[date_use] if date_use in pct_series.index
                              else pct_series.iloc[pct_series.index.get_indexer([date_use], method="nearest")[0]])

            # 冷却期目标冻结：分位点仍高于 resume_pct 时追 offset
            if ts.state == "cooldown" and price_pct > resume_pct:
                ts.period_offset = ts.period

            if paused and ts.state not in ("cooldown",):
                effective_prev = (ts.period - 1) - ts.period_offset
                if ts.frozen_target is None:
                    ts.frozen_target = ts.base * max(0, effective_prev)
                ts.frozen_target += ts.base * INCREMENT_PER_PERIOD / TOTAL_PER_PERIOD

            effective_period = ts.period - ts.period_offset
            effective_state = ts.state
            if ts.state == "cooldown" and price_pct <= resume_pct:
                effective_state = "resume"

            result = compute_period(
                ts.base, effective_period, price, ma250, holding,
                effective_state, global_cumulative,
                frozen_target=ts.frozen_target,
                price_pct=price_pct,
                liquidate_pct=liquidate_pct,
                resume_pct=resume_pct,
            )
            plans.append({
                "tname": tname, "ts": ts, "date_use": date_use,
                "price": price, "ma250": ma250, "holding": holding,
                "result": result, "price_pct": price_pct,
            })

        if paused:
            active_plans = [(i, p) for i, p in enumerate(plans)
                            if p["result"].get("state_out") not in ("cooldown",)]
            total_regular = sum(p["result"]["regular"] for _, p in active_plans
                                if p["result"]["regular"] > 0)
            if total_regular > reserve_pool and total_regular > 0:
                scale = reserve_pool / total_regular
                for _, p in active_plans:
                    if p["result"]["regular"] > 0:
                        p["result"]["regular"] *= scale
                        p["result"]["actual"] = p["result"]["regular"] + p["result"]["harvest"]

        for p in plans:
            tname = p["tname"]
            ts = p["ts"]
            result = p["result"]
            price = p["price"]
            holding = p["holding"]
            actual = result["actual"]

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
                    ts.shares = max(0, ts.shares - sell_amount / price)
                    reserve_pool += sell_amount
                ts.state = result.get("state_out", "normal")

            current_value = ts.shares * price
            ts.peak_value = max(ts.peak_value, current_value)
            if ts.peak_value > 0:
                ts.max_drawdown = max(ts.max_drawdown,
                                      (ts.peak_value - current_value) / ts.peak_value)
            ts.harvest_count += 1 if result["harvest"] < 0 else 0

            backtest_rows[tname].append(dict(
                price=price, ma250=p["ma250"], holding=holding,
                actual=actual, notes=result["notes"],
                deviation=result["deviation"], price_pct=result["price_pct"],
            ))

    return dict(
        states=states,
        global_cumulative=global_cumulative,
        reserve_pool=reserve_pool,
        increment_cumulative=increment_cumulative,
        increment_deployed=increment_deployed,
        backtest_rows=backtest_rows,
    )


def summarize(bt_result):
    states = bt_result["states"]
    rows = bt_result["backtest_rows"]
    n = len(rows[TARGET_NAMES[0]])
    years = n * 14 / 365
    total_holding = sum(states[t].shares * rows[t][-1]["price"] for t in TARGET_NAMES)
    reserve = bt_result["reserve_pool"]
    user_cash = bt_result["global_cumulative"] + bt_result.get("increment_cumulative", 0)
    net = total_holding + reserve
    ret = (net - user_cash) / user_cash if user_cash > 0 else 0
    ann = (1 + ret) ** (1 / years) - 1 if years > 0 and ret > -1 else 0
    liquidates = sum(states[t].liquidate_count for t in TARGET_NAMES)
    return dict(ret=ret, ann=ann, liquidates=liquidates, years=years)


if __name__ == "__main__":
    print("Loading cached price data...")
    price_data = load_price_cache()
    pct_data = compute_percentile_series(price_data)
    print("Percentile series computed.")

    key_labels = ["2年A(24-26)", "5年A(21-26)", "10年A(16-26)", "20年(06-26)"]
    print(f"\n{'回测周期':<16} {'总收益':>8} {'年化':>8} {'清仓次数':>8}")
    print("=" * 46)
    for label, bt_start, bt_end in BACKTEST_PERIODS:
        if label not in key_labels:
            continue
        r = run_backtest(price_data, pct_data, bt_start, bt_end)
        if r is None:
            continue
        s = summarize(r)
        print(f"{label:<16} {s['ret']:>+8.1%} {s['ann']:>+8.1%} {s['liquidates']:>8}")
```

- [ ] **Step 2: 快速验证语法和导入**

```bash
python3 -c "import strategies.percentile_liquidation.strategy as s; print('OK')"
```

期望输出：`OK`

- [ ] **Step 3: Commit**

```bash
git add strategies/percentile_liquidation/strategy.py
git commit -m "feat: percentile-liquidation core engine with expanding-window pct rank"
```

---

## Task 3: 实现 grid_search.py

**Files:**
- Create: `strategies/percentile_liquidation/grid_search.py`

- [ ] **Step 1: 创建 grid_search.py**

```python
"""网格搜索最优 (清仓分位点, 再入分位点) 组合。

在所有回测周期上跑 25 组参数，按 20 年年化收益排序输出。
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

KEY_PERIODS = {
    "2年A(24-26)": None,
    "5年A(21-26)": None,
    "10年A(16-26)": None,
    "20年(06-26)": None,
}


def run_grid(price_data, pct_data):
    period_map = {label: (bt_start, bt_end)
                  for label, bt_start, bt_end in BACKTEST_PERIODS
                  if label in KEY_PERIODS}

    records = []
    combos = [(lp, rp) for lp in LIQUIDATE_PCTS for rp in RESUME_PCTS if rp < lp]

    for liq_pct, res_pct in combos:
        row = {"liq_pct": liq_pct, "res_pct": res_pct}
        for label, (bt_start, bt_end) in period_map.items():
            r = run_backtest(price_data, pct_data, bt_start, bt_end, liq_pct, res_pct)
            if r is None:
                row[label] = None
                continue
            row[label] = summarize(r)
        records.append(row)

    return records


def print_results(records):
    target_label = "20年(06-26)"
    valid = [r for r in records if r.get(target_label) is not None]
    valid.sort(key=lambda r: r[target_label]["ann"], reverse=True)

    labels = ["2年A(24-26)", "5年A(21-26)", "10年A(16-26)", "20年(06-26)"]
    header = f"{'清仓%':>6} {'再入%':>6}" + "".join(
        f"  {lb[:5]:>5}总收益  {lb[:5]:>5}年化  清仓" for lb in labels
    )
    print(header)
    print("-" * len(header))

    for r in valid:
        line = f"{r['liq_pct']*100:>5.0f}% {r['res_pct']*100:>5.0f}%"
        for lb in labels:
            s = r.get(lb)
            if s is None:
                line += "        N/A        N/A   N/A"
            else:
                line += f"  {s['ret']:>+8.1%}  {s['ann']:>+7.1%}  {s['liquidates']:>3}"
        print(line)

    print(f"\n最优组合（按20年年化）：清仓={valid[0]['liq_pct']*100:.0f}%，"
          f"再入={valid[0]['res_pct']*100:.0f}%，"
          f"年化={valid[0][target_label]['ann']:+.2%}")


if __name__ == "__main__":
    print("Loading cache...")
    price_data = load_price_cache()
    print("Computing percentile series...")
    pct_data = compute_percentile_series(price_data)
    print(f"Running {len(LIQUIDATE_PCTS) * len(RESUME_PCTS)} combinations...\n")
    records = run_grid(price_data, pct_data)
    print_results(records)
```

- [ ] **Step 2: Commit**

```bash
git add strategies/percentile_liquidation/grid_search.py
git commit -m "feat: grid search script for percentile liquidation thresholds"
```

---

## Task 4: 运行并验证

- [ ] **Step 1: 运行默认参数回测，确认输出正常**

```bash
python3 strategies/percentile_liquidation/strategy.py
```

期望：打印 4 个关键周期的总收益 / 年化 / 清仓次数，无报错。

- [ ] **Step 2: 运行网格搜索**

```bash
python3 strategies/percentile_liquidation/grid_search.py
```

期望：打印 25 行结果表 + 最优组合一行摘要，运行时间 < 60s。

- [ ] **Step 3: 最终 commit**

```bash
git add -A
git commit -m "feat: percentile-liquidation strategy complete, grid search ready"
```
