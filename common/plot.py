"""Shared plotting utilities: fonts, colors, helpers."""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.ticker import FuncFormatter

font_path = '/System/Library/Fonts/STHeiti Medium.ttc'
font_prop = fm.FontProperties(fname=font_path)
plt.rcParams['font.family'] = font_prop.get_name()
plt.rcParams['axes.unicode_minus'] = False

ETF_NAMES = ['沪深300 ETF', '上证50 ETF', '恒生指数 ETF', '标普500 ETF']

ETF_COLORS = {
    '沪深300 ETF': '#E63946',
    '上证50 ETF': '#F4A261',
    '恒生指数 ETF': '#2A9D8F',
    '标普500 ETF': '#264653',
}

PERIOD_COLORS = {
    '2年': '#457B9D',
    '5年': '#E76F51',
    '10': '#264653',
    '20': '#7B2D8E',
}


def short_label(p):
    """Extract the date range part like '24-26' from '2年A(24-26)'."""
    if '(' in p:
        return p[p.index('(') + 1:p.index(')')]
    return p


def wan_formatter(x, _):
    return f'{x / 10000:.0f}万'


wan_func_formatter = FuncFormatter(wan_formatter)
