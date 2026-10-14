#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学习机付费日报看板 · 固化更新脚本
用法:  python3 build.py <今年付费日报明细.xlsx> <明细表-全渠道.xlsx>
作用:  聚合 2026 年窗口(10/12–11/11)数据,与冻结的 2025 基线(data2025.json)合并,
       注入 dashboard_template.html,输出 ../index.html。
口径:  京东自营官旗+POP官旗 · 讯飞学习机 · 剔除空付费渠道行;品专(品牌合约)保留原始行,
       页面内展示时自动剔除。净GMV: 渠道=京东,同两店铺,剔合计行。
依赖:  /usr/bin/python3 (系统自带 pandas/openpyxl)。
"""
import pandas as pd, json, sys, os, datetime, warnings
warnings.filterwarnings('ignore')
HERE = os.path.dirname(os.path.abspath(__file__))
YR = 2026
LO, HI = f"{YR}-10-12", f"{YR}-11-11"

if len(sys.argv) != 3:
    sys.exit("用法: python3 build.py <付费日报明细.xlsx> <明细表-全渠道.xlsx>")
paid_path, net_path = sys.argv[1], sys.argv[2]

def shop_of(s):
    s = str(s)
    if '自营官旗' in s: return '自营'
    if 'POP官旗' in s: return 'POP'
    return None

# ---- 今年付费 ----
df = pd.ExcelFile(paid_path).parse(0)
df = df[df['产品系列'].astype(str) == '讯飞学习机'].copy()
df = df[df['付费渠道'].notna()]
df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
df = df[(df['日期'] >= LO) & (df['日期'] <= HI)]
df['shop'] = df['店铺'].astype(str).map(shop_of)
df = df[df['shop'].notna()]
df['d'] = df['日期'].dt.strftime('%m-%d')
df['简称'] = df['简称'].fillna('公共').astype(str)
for c in ['花费', '点击量', '总成交金额']:
    df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
g = df.groupby(['d', 'shop', '付费渠道', '简称']).agg(
    fee=('花费', 'sum'), clk=('点击量', 'sum'), gmv=('总成交金额', 'sum')).reset_index()
paid26 = [[YR, r['d'], r['shop'], str(r['付费渠道']), r['简称'],
           round(float(r['fee']), 2), int(r['clk']), round(float(r['gmv']), 2)]
          for _, r in g.iterrows()]
ex = df[df['付费渠道'].astype(str) != '品牌合约广告位']
print(f"今年付费: {df['d'].min()}–{df['d'].max()} 共{df['d'].nunique()}天, "
      f"剔品专花费 {ex['花费'].sum():,.0f}, ROI {ex['总成交金额'].sum()/ex['花费'].sum():.2f}")

# ---- 今年净GMV ----
df = pd.ExcelFile(net_path).parse(0)
df = df[(df['渠道'].astype(str) == '京东') & (df['产品系列'].astype(str) == '讯飞学习机')
        & (df['店铺'].astype(str).isin(['京东自营官旗', '京东POP官旗']))].copy()
df['时间'] = pd.to_datetime(df['时间'], errors='coerce')
df['净GMV'] = pd.to_numeric(df['净GMV'], errors='coerce').fillna(0)
w = df[(df['时间'] >= LO) & (df['时间'] <= HI)].copy()
w['d'] = w['时间'].dt.strftime('%m-%d')
w['shop'] = w['店铺'].astype(str).map(shop_of)
g = w.groupby(['d', 'shop'])['净GMV'].sum().reset_index()
net26 = [[YR, r['d'], r['shop'], round(float(r['净GMV']), 2)] for _, r in g.iterrows()]
print(f"今年净GMV: 至 {w['d'].max() if len(w) else '无'}, 合计 {w['净GMV'].sum():,.0f}")

# ---- 合并基线并注入 ----
base = json.load(open(os.path.join(HERE, 'data2025.json')))
built = datetime.date.today().isoformat()
data = json.dumps({'built': built, 'paid': base['paid'] + paid26, 'net': base['net'] + net26},
                  ensure_ascii=False, separators=(',', ':'))
html = open(os.path.join(HERE, 'dashboard_template.html')).read()
html = html.replace('__DATA__', data).replace('__BUILT__', built)
out = os.path.join(HERE, '..', 'index.html')
open(out, 'w').write(html)
print(f"已写入 {os.path.abspath(out)} ({len(html)} 字符)。发布: git add -A && git commit && git push")
