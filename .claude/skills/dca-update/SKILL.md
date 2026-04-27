---
name: dca-update
description: 执行一期定投操作：读取Excel现状→计算建议→用户确认并执行→录入实际成交→更新Excel→汇报
---

# 定投追踪更新

执行一次完整的定投操作，分4个阶段。每个阶段完成后等待用户再继续。

---

## 策略说明（朴素恒定市值法）

- **目标市值** = 基准金额 × (期数 − 期数偏移)（增量阶段按 2000/6000 比例缓增）
- **每期操作** = 目标市值 − 持仓市值（正=买入缺口，负=卖出超额）
- **减半清仓**：250日均线偏离度 ≥ +35% 且持仓 > 目标市值×50%，卖至目标50%，**同时目标市值减半**（期数偏移更新），不进冷却
- **全仓清仓**：偏离度 ≥ +55%，全部清仓，进入冷却期，期数计数重置
- 无网格收割分档、无大跌加码分档

---

## 固定配置

```python
EXCEL_PATH   = "/Users/youxingzhi/ayou/financial_strategy/strategies/constant_value/定投追踪记录.xlsx"
ETF_SHEETS   = ["沪深300", "中证500", "恒生指数", "纳指100"]
BASE_AMOUNTS = {"沪深300": 1500, "中证500": 1500, "恒生指数": 600, "纳指100": 2400}

TICKERS = {
    "沪深300":  "510300.SS",
    "中证500":  "510500.SS",
    "恒生指数": "159920.SZ",
    "纳指100":  "513100.SS",
}

PARTIAL_LIQUIDATE = 0.35   # 减半清仓阈值
FULL_LIQUIDATE    = 0.55   # 全仓清仓阈值
PAUSE_TOTAL          = 150000
INCREMENT_PER_PERIOD = 2000
TOTAL_PER_PERIOD     = 6000

DATA_START_ROW = 5   # ETF sheet 数据从第5行开始
COL_B  = 2   # 期数      (手动)
COL_D  = 4   # 当前价格  (手动)
COL_E  = 5   # 250日均线 (手动)
COL_G  = 7   # 持仓份额  (手动，填上期操作后份额)
COL_L  = 12  # 实际操作金额 (手动，正=买入 负=卖出)
COL_M  = 13  # 买卖份额     (手动，正=买入 负=卖出)
COL_N  = 14  # 操作后份额   (公式，不手动写)
```

---

## 阶段一：读取当前状态 & 获取价格

写并运行以下 Python 代码（一次性脚本，直接在 Bash 中执行）：

```python
import openpyxl, yfinance as yf, math, datetime

EXCEL_PATH   = "/Users/youxingzhi/ayou/financial_strategy/strategies/constant_value/定投追踪记录.xlsx"
ETF_SHEETS   = ["沪深300", "中证500", "恒生指数", "纳指100"]
BASE_AMOUNTS = {"沪深300": 1500, "中证500": 1500, "恒生指数": 600, "纳指100": 2400}
TICKERS      = {"沪深300": "510300.SS", "中证500": "510500.SS",
                "恒生指数": "159920.SZ", "纳指100": "513100.SS"}

PARTIAL_LIQUIDATE = 0.35
FULL_LIQUIDATE    = 0.55
PAUSE_TOTAL          = 150000
INCREMENT_PER_PERIOD = 2000
TOTAL_PER_PERIOD     = 6000

# ── 读取 Excel 持仓 ──────────────────────────────────────────
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

state = {}
for sname in ETF_SHEETS:
    ws = wb[sname]
    last_period, last_shares = 0, 0.0
    for row in ws.iter_rows(min_row=5, values_only=True):
        b = row[1]   # B 列 期数
        n = row[13]  # N 列 操作后份额
        if b is not None:
            try:
                last_period = int(b)
                last_shares = float(n) if n is not None else 0.0
            except:
                pass
    # 读取期数偏移（减半清仓后由脚本自动更新，存于各 ETF sheet 的 S3 单元格）
    raw_offset = ws.cell(row=3, column=19).value
    period_offset = float(raw_offset) if raw_offset is not None else 0.0
    state[sname] = {"last_period": last_period, "shares": last_shares, "period_offset": period_offset}

# ── 读取每期汇总 ──────────────────────────────────────────────
ws_sum = wb["每期汇总"]
pool_balance, cum_invest, current_phase = 0.0, 0.0, "存量阶段"
for row in ws_sum.iter_rows(min_row=5, values_only=True):
    if row[0] is not None:
        if row[10] is not None: pool_balance  = float(row[10])
        if row[11] is not None: cum_invest    = float(row[11])
        if row[12] is not None: current_phase = str(row[12])
wb.close()

next_period = max(s["last_period"] for s in state.values()) + 1

# ── 获取价格 & MA250 ──────────────────────────────────────────
prices = {}
for sname, ticker in TICKERS.items():
    hist = yf.download(ticker, period="2y", progress=False, auto_adjust=True)
    if hist.empty:
        prices[sname] = None
        print(f"[警告] {sname} ({ticker}) 价格获取失败，需手动输入")
        continue
    close = hist["Close"].squeeze()
    price = float(close.iloc[-1])
    ma250 = float(close.rolling(250).mean().iloc[-1])
    prices[sname] = {"price": price, "ma250": ma250, "date": str(close.index[-1].date())}

# ── 计算建议 ──────────────────────────────────────────────────
paused = cum_invest >= PAUSE_TOTAL

print(f"\n══ 第 {next_period} 期定投建议 · {datetime.date.today()} ══")
print(f"当前阶段: {current_phase}  |  储备金池: ¥{pool_balance:,.2f}  |  累计净投: ¥{cum_invest:,.2f}\n")

HEADER = f"{'标的':<10} {'价格':>10} {'MA250':>10} {'偏离度':>8} {'持仓份额':>10} {'持仓市值':>10} {'目标市值':>10} {'操作建议':<10} {'建议金额':>12} {'建议份额':>10}"
print(HEADER)
print("─" * len(HEADER))

results = {}
for sname in ETF_SHEETS:
    if prices[sname] is None:
        print(f"{sname:<10}  [价格缺失]")
        continue
    p     = prices[sname]
    base  = BASE_AMOUNTS[sname]
    price = p["price"]
    ma250 = p["ma250"]
    dev   = (price - ma250) / ma250
    sh    = state[sname]["shares"]
    hold  = sh * price

    # 增量阶段目标市值按比例缓增
    period_offset = state[sname]["period_offset"]
    effective_period = next_period - period_offset
    tgt = BASE_AMOUNTS[sname] * effective_period
    if paused:
        # 增量阶段：每期只增长 base × (2000/6000)
        prev_effective = state[sname]["last_period"] - period_offset
        prev_tgt = base * max(0, prev_effective)
        tgt = prev_tgt + base * INCREMENT_PER_PERIOD / TOTAL_PER_PERIOD

    # 操作决策
    if dev >= FULL_LIQUIDATE and hold > 0:
        action = "全仓清仓"
        amt    = -hold
    elif dev >= PARTIAL_LIQUIDATE and hold > tgt * 0.5:
        action = "减半清仓"
        amt    = -(hold - tgt * 0.5)
    elif hold > tgt:
        action = "卖出超额"
        amt    = -(hold - tgt)
    else:
        action = "正常定投"
        amt    = tgt - hold

    # 买卖均须为 100 股整数倍
    if action == "全仓清仓":
        suggest_shares = -int(sh)   # 直接用持仓份额，避免浮点误差
    elif amt > 0:
        suggest_shares = math.floor(amt / price / 100) * 100
        if suggest_shares == 0:
            suggest_shares = 100    # 最少买 100 股
    else:
        suggest_shares = -(math.ceil(abs(amt) / price / 100) * 100)

    results[sname] = {
        "price": price, "ma250": ma250, "dev": dev,
        "shares": sh, "hold": hold, "target": tgt,
        "action": action, "suggest_amount": amt, "suggest_shares": suggest_shares,
    }
    print(f"{sname:<10} {price:>10.4f} {ma250:>10.4f} {dev:>+8.2%} {sh:>10.0f} "
          f"{hold:>10,.0f} {tgt:>10,.0f} {action:<10} {amt:>+12,.0f} {suggest_shares:>+10.0f}")

print()
print("请按以上建议执行买卖操作，完成后提供每个标的的实际成交数据：")
print("  实际操作金额（正=买入 负=卖出）")
print("  实际成交份额（正=买入 负=卖出）")
```

若任何 ticker 获取失败，向用户请求手动输入该标的的「当前价格」和「250日均线」，然后重新计算。

---

## 阶段二：等待用户执行

展示计算结果后，说：

> **以上是第 N 期建议操作，请按照计划去交易所买卖。完成后告诉我每个标的的实际成交金额和份额（格式随意，比如"沪深300买了200股，花了1320元"）。**

然后**等待用户回复**实际成交数据。

---

## 阶段三：录入实际成交 & 更新 Excel

收到用户实际成交数据后，解析成如下结构（0表示未操作）：

```python
actuals = {
    "沪深300":  {"amount": 1320.00,  "shares": 200},
    "中证500":  {"amount": 827.00,   "shares": 100},
    "恒生指数": {"amount": 590.00,   "shares": 100},
    "纳指100":  {"amount": 2100.00,  "shares": 200},
}
```

然后运行以下 Python 代码写入 Excel：

```python
import openpyxl, datetime

EXCEL_PATH = "/Users/youxingzhi/ayou/financial_strategy/strategies/constant_value/定投追踪记录.xlsx"
wb = openpyxl.load_workbook(EXCEL_PATH)
today = datetime.date.today().strftime("%Y-%m-%d")

for sname, actual in actuals.items():
    ws = wb[sname]
    # 找第一个 B 列为空的数据行
    write_row = None
    for row_idx in range(5, 205):
        if ws.cell(row=row_idx, column=2).value is None:
            write_row = row_idx
            break
    if write_row is None:
        print(f"[错误] {sname} 表格行数已满，请扩展 create_tracker.py 的 MAX_ROWS")
        continue

    ws.cell(row=write_row, column=1).value  = today                        # A 日期
    ws.cell(row=write_row, column=2).value  = next_period                  # B 期数
    ws.cell(row=write_row, column=4).value  = prices[sname]["price"]       # D 当前价格
    ws.cell(row=write_row, column=5).value  = prices[sname]["ma250"]       # E 250日均线
    ws.cell(row=write_row, column=7).value  = state[sname]["shares"]       # G 上期持仓份额
    ws.cell(row=write_row, column=12).value = actual["amount"]             # L 实际操作金额
    ws.cell(row=write_row, column=13).value = actual["shares"]             # M 买卖份额

# 更新每期汇总 sheet
ws_sum = wb["每期汇总"]
sum_row = None
for row_idx in range(5, 205):
    if ws_sum.cell(row=row_idx, column=1).value is None:
        sum_row = row_idx
        break

ws_sum.cell(row=sum_row, column=1).value = next_period   # A 期数
ws_sum.cell(row=sum_row, column=2).value = today          # B 日期
# H 本期增量注入（存量阶段=0，增量阶段=2000）
ws_sum.cell(row=sum_row, column=8).value = (
    INCREMENT_PER_PERIOD if current_phase == "增量阶段" else 0
)

wb.save(EXCEL_PATH)
print(f"[OK] Excel 已更新：{EXCEL_PATH}")

# 减半清仓后更新期数偏移（目标市值同步减半）
# 注意：此处的 results / state / next_period 在生成脚本时已由 Phase 1 的输出填入
partial_updated = False
for sname in ETF_SHEETS:
    if results[sname]["action"] == "减半清仓":
        ws = wb[sname]
        old_offset = state[sname]["period_offset"]
        ep = next_period - old_offset
        new_offset = next_period - ep / 2
        ws.cell(row=3, column=19).value = new_offset
        print(f"[INFO] {sname} 期数偏移 {old_offset} → {new_offset:.1f}（目标市值已减半）")
        partial_updated = True
if partial_updated:
    wb.save(EXCEL_PATH)
    print(f"[OK] 期数偏移已写入 Excel")
```

**注意**：C/F/H/I/J/K/N/O/P 列均为公式列，openpyxl 会保留模板里的公式，打开 Excel 时自动计算，无需手动写入。

---

## 阶段四：汇报

输出本期操作完成汇报，格式如下：

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  第 N 期定投完成 · YYYY-MM-DD
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【操作明细】
  沪深300   买入  200 份   花费 ¥1,320.00   （偏离度 +X.X%，正常定投）
  中证500   买入  100 份   花费 ¥  827.00   （偏离度 +X.X%，正常定投）
  恒生指数  买入  100 份   花费 ¥  590.00   （偏离度 -X.X%，正常定投）
  纳指100   买入  200 份   花费 ¥2,100.00   （偏离度 -X.X%，正常定投）

【本期自掏腰包】¥4,837.00
【当前阶段】存量阶段（距增量阶段还差 ¥XXX,XXX）
【储备金池】¥0.00

【持仓快照（操作前→操作后）】
  沪深300   0 → 200 份   市值 ¥X,XXX   目标 ¥X,XXX
  中证500   0 → 100 份   市值 ¥X,XXX   目标 ¥X,XXX
  恒生指数  0 → 100 份   市值 ¥X,XXX   目标 ¥X,XXX
  纳指100   0 → 200 份   市值 ¥X,XXX   目标 ¥X,XXX

Excel 已更新，用 Numbers/Excel 打开可查看公式计算结果。
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

汇报输出后，计算下期日期（今天 + 14 天），然后运行以下命令在 macOS 提醒事项中创建系统提醒：

```bash
osascript <<OSEOF
tell application "Reminders"
    tell list "提醒事项"
        make new reminder with properties {name:"第 {N+1} 期定投 — 执行 /dca-update", remind me date:date "{YYYY-MM-DD} 09:00:00", body:"沪深300 / 中证500 / 恒生指数 / 纳指100 双周定投"}
    end tell
end tell
OSEOF
```

将 `{N+1}` 替换为下期期数，`{YYYY-MM-DD}` 替换为下期日期（今天 + 14 天）。
创建成功后告知用户已在「提醒事项」中设好下期提醒。

---

## 边界情况处理

| 情况 | 处理方式 |
|------|---------|
| yfinance 某标的获取失败 | 向用户请求手动输入价格和MA120 |
| 用户说"按建议操作了" | 用建议金额和建议份额作为实际值 |
| 用户说某标的"没操作" | amount=0, shares=0，仍要写入行（记录本期持仓不变） |
| 全仓清仓后持仓为0 | G列填0，L列填负数（卖出金额），M列填负数（卖出份额） |
| 减半清仓 | L列填负数（卖出金额），M列填负数（卖出份额） |
| 增量阶段目标市值计算 | 如有期数重置（清仓后），以用户实际Excel数据为准，脚本计算仅供参考 |
| Excel行数已满（>100行）| 提示用户重新运行 create_tracker.py 扩大 MAX_ROWS |
