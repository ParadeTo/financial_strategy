#!/usr/bin/env python3
"""Plot price charts with DCA operation markers for each ETF."""

import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from strategies.constant_value.strategy import TargetState
from common.data import load_price_cache, load_backtest_cache
from common.plot import plt, font_prop, ETF_NAMES
from strategies.constant_value.strategy import BASE_AMOUNTS
import matplotlib.dates as mdates
import pandas as pd

price_data = load_price_cache()
all_results = load_backtest_cache()

STRATEGY_DIR = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(STRATEGY_DIR, 'charts')
os.makedirs(OUT, exist_ok=True)

for period_key, bt in all_results.items():
    backtest_rows = bt['backtest_rows']

    for etf_idx, tname in enumerate(ETF_NAMES):
        base = BASE_AMOUNTS[etf_idx]
        rows = backtest_rows[tname]
        df = price_data[tname].copy()
        bt_start = rows[0]['date']
        bt_end = rows[-1]['date']
        df = df[bt_start:bt_end]

        dates_all = df.index
        close_all = df['close'].values
        ma250_all = df['ma250'].values

        buy_dates, buy_prices, buy_sizes = [], [], []
        sell_dates, sell_prices, sell_sizes = [], [], []
        liquidate_dates, liquidate_prices = [], []

        for r in rows:
            d = pd.Timestamp(r['date'])
            p = r['price']

            if '全仓清仓' in r['notes'] and r['actual'] < 0:
                liquidate_dates.append(d)
                liquidate_prices.append(p)
            elif r['harvest'] < 0:
                sell_dates.append(d)
                sell_prices.append(p)
                sell_sizes.append(abs(r['harvest']))

            if r['actual'] > 0:
                regular = r['regular'] if r['regular'] > 0 else 0
                if regular > 0:
                    buy_dates.append(d)
                    buy_prices.append(p)
                    buy_sizes.append(regular)

        is_10y = period_key.startswith('10年')
        is_5y = period_key.startswith('5年')
        is_20y = period_key.startswith('20年')

        if is_20y:
            width, bar_w, interval = 30, 4, 6
        elif is_10y:
            width, bar_w, interval = 26, 6, 6
        elif is_5y:
            width, bar_w, interval = 22, 10, 3
        else:
            width, bar_w, interval = 16, 14, 2

        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(width, 12), height_ratios=[3, 1],
                                         sharex=True, gridspec_kw={'hspace': 0.06})

        ax1.plot(dates_all, close_all, color='#333333', linewidth=1.8, label='收盘价', zorder=2)
        ax1.plot(dates_all, ma250_all, color='#F4A261', linewidth=1.8, label='250日均线',
                 linestyle='--', alpha=0.85, zorder=2)

        s_base = 80
        if buy_dates:
            sizes = [max(s_base, s_base * s / 1500) for s in buy_sizes]
            ax1.scatter(buy_dates, buy_prices, c='#2A9D8F', s=sizes, marker='^',
                        alpha=0.85, zorder=3, label='常规买入', edgecolors='white', linewidths=0.8)
        if sell_dates:
            sizes = [max(s_base, s_base * s / 1500) for s in sell_sizes]
            ax1.scatter(sell_dates, sell_prices, c='#E63946', s=sizes, marker='v',
                        alpha=0.85, zorder=3, label='网格收割', edgecolors='white', linewidths=0.8)
        if liquidate_dates:
            ax1.scatter(liquidate_dates, liquidate_prices, c='#264653', s=350, marker='X',
                        alpha=1.0, zorder=6, label='极端清仓', edgecolors='white', linewidths=1.5)

        ax1.set_title(f'{tname} — {period_key} 股价走势与定投操作',
                      fontsize=22, fontweight='bold', fontproperties=font_prop, pad=15)
        ax1.set_ylabel('价格 (元)', fontsize=16, fontproperties=font_prop)
        ax1.legend(prop=font_prop, fontsize=14, loc='upper left', ncol=5,
                   framealpha=0.9, edgecolor='#cccccc')
        ax1.spines['top'].set_visible(False)
        ax1.spines['right'].set_visible(False)
        ax1.grid(axis='y', alpha=0.3)
        ax1.tick_params(axis='y', labelsize=13)
        for label in ax1.get_yticklabels():
            label.set_fontproperties(font_prop)

        bar_dates = [pd.Timestamp(r['date']) for r in rows]
        deviations = [r['deviation'] * 100 for r in rows]
        colors = ['#E63946' if d > 0 else '#2A9D8F' for d in deviations]
        ax2.bar(bar_dates, deviations, width=bar_w, color=colors, alpha=0.75)

        ax2.axhline(y=8, color='#E63946', linewidth=1, linestyle=':', alpha=0.7, label='+8% 收割线')
        ax2.axhline(y=-8, color='#2A9D8F', linewidth=1, linestyle=':', alpha=0.7, label='-8% 加码线')
        ax2.axhline(y=40, color='#264653', linewidth=1.2, linestyle='--', alpha=0.7, label='+40% 清仓线')
        ax2.axhline(y=0, color='gray', linewidth=0.6)

        ax2.set_ylabel('偏离度 (%)', fontsize=16, fontproperties=font_prop)
        ax2.set_xlabel('日期', fontsize=16, fontproperties=font_prop)
        ax2.spines['top'].set_visible(False)
        ax2.spines['right'].set_visible(False)
        ax2.grid(axis='y', alpha=0.3)
        ax2.legend(prop=font_prop, fontsize=12, loc='lower left', ncol=3,
                   framealpha=0.9, edgecolor='#cccccc')
        ax2.tick_params(axis='both', labelsize=13)

        ax2.xaxis.set_major_locator(mdates.MonthLocator(interval=interval))
        ax2.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        for label in ax2.get_xticklabels():
            label.set_fontproperties(font_prop)
            label.set_rotation(30)
            label.set_ha('right')
            label.set_fontsize(12)
        for label in ax2.get_yticklabels():
            label.set_fontproperties(font_prop)

        slug = tname.replace(' ', '_')
        period_slug = period_key.replace('(', '_').replace(')', '')
        path = f'{OUT}/price_{slug}_{period_slug}.png'
        fig.savefig(path, dpi=150, bbox_inches='tight')
        plt.close()
        print(f'Saved: {path}')

print('Done.')
