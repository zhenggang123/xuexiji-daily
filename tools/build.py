#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
学习机付费日报看板 · 固化更新脚本（增量合并版）
用法:  python3 build.py <底表1.xlsx> [<底表2.xlsx> ...]
       文件类型自动识别：含「付费渠道」列=付费日报明细；含「净GMV」列=明细表-全渠道。
       全量或单日增量导出均可：文件中出现的日期覆盖已存数据，未出现的日期保留
       （2026年数据累积在 tools/cur2026.json）。
作用:  聚合 2026 年窗口(10/12–11/11)数据，与冻结的 2025 基线(data2025.json)合并,
       注入 dashboard_template.html，输出 ../index.html。
口径:  京东自营官旗+POP官旗 · 讯飞学习机 · 剔除空付费渠道行;品专(品牌合约)保留原始行,
       页面内展示时自动剔除。净GMV: 渠道=京东,同两店铺,剔合计行。
依赖:  /usr/bin/python3 (系统自带 pandas/openpyxl)。
"""
import pandas as pd, json, sys, os, datetime, warnings
warnings.filterwarnings('ignore')
HERE = os.path.dirname(os.path.abspath(__file__))
YR = 2026
LO, HI = f"{YR}-10-12", f"{YR}-11-11"

if len(sys.argv) < 2:
    print("未传入底表：仅用已累积数据重新生成页面（模板改动后固化用）")

def shop_of(s):
    s = str(s)
    if '自营官旗' in s: return '自营'
    if 'POP官旗' in s: return 'POP'
    return None

def agg_paid(df):
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
    rows = [[YR, r['d'], r['shop'], str(r['付费渠道']), r['简称'],
             round(float(r['fee']), 2), int(r['clk']), round(float(r['gmv']), 2)]
            for _, r in g.iterrows()]
    return rows, sorted(df['d'].unique())

def agg_net(df):
    df = df[(df['渠道'].astype(str) == '京东') & (df['产品系列'].astype(str) == '讯飞学习机')
            & (df['店铺'].astype(str).isin(['京东自营官旗', '京东POP官旗']))].copy()
    df['时间'] = pd.to_datetime(df['时间'], errors='coerce')
    df['净GMV'] = pd.to_numeric(df['净GMV'], errors='coerce').fillna(0)
    w = df[(df['时间'] >= LO) & (df['时间'] <= HI)].copy()
    w['d'] = w['时间'].dt.strftime('%m-%d')
    w['shop'] = w['店铺'].astype(str).map(shop_of)
    g = w.groupby(['d', 'shop'])['净GMV'].sum().reset_index()
    rows = [[YR, r['d'], r['shop'], round(float(r['净GMV']), 2)] for _, r in g.iterrows()]
    return rows, sorted(w['d'].unique())

# ---- 读累积数据 ----
cur_path = os.path.join(HERE, 'cur2026.json')
cur = json.load(open(cur_path)) if os.path.exists(cur_path) else {'paid': [], 'net': []}

# ---- 逐文件识别类型并按日期覆盖合并 ----
for p in sys.argv[1:]:
    df = pd.ExcelFile(p).parse(0)
    cols = [str(c) for c in df.columns]
    name = os.path.basename(p)
    if '付费渠道' in cols:
        rows, days = agg_paid(df)
        if not days:
            print(f"[付费] {name}: 窗口内无数据，跳过"); continue
        cur['paid'] = [r for r in cur['paid'] if r[1] not in days] + rows
        print(f"[付费] {name}: 覆盖 {days[0]}–{days[-1]} 共{len(days)}天")
    elif '净GMV' in cols:
        rows, days = agg_net(df)
        if not days:
            print(f"[净GMV] {name}: 窗口内无数据，跳过"); continue
        cur['net'] = [r for r in cur['net'] if r[1] not in days] + rows
        print(f"[净GMV] {name}: 覆盖 {days[0]}–{days[-1]} 共{len(days)}天")
    else:
        print(f"[跳过] {name}: 无法识别（缺少 付费渠道/净GMV 列）")

cur['paid'].sort(key=lambda r: (r[1], r[2], r[3], r[4]))
cur['net'].sort(key=lambda r: (r[1], r[2]))
json.dump(cur, open(cur_path, 'w'), ensure_ascii=False, separators=(',', ':'))

# ---- 汇总校验输出 ----
ex = [r for r in cur['paid'] if r[3] != '品牌合约广告位']
fee = sum(r[5] for r in ex); gmv = sum(r[7] for r in ex)
pd_days = sorted({r[1] for r in cur['paid']}); nt_days = sorted({r[1] for r in cur['net']})
print(f"2026累计: 付费 {pd_days[0]}–{pd_days[-1]}（{len(pd_days)}天，剔品专花费 {fee:,.0f}，ROI {gmv/fee:.2f}）"
      f"；净GMV {nt_days[0]}–{nt_days[-1]}（{len(nt_days)}天，合计 {sum(r[3] for r in cur['net']):,.0f}）")

# ---- 合并基线并注入 ----
base = json.load(open(os.path.join(HERE, 'data2025.json')))
built = datetime.date.today().isoformat()
data = json.dumps({'built': built, 'paid': base['paid'] + cur['paid'], 'net': base['net'] + cur['net']},
                  ensure_ascii=False, separators=(',', ':'))
html = open(os.path.join(HERE, 'dashboard_template.html')).read()
html = html.replace('__DATA__', data).replace('__BUILT__', built)
out = os.path.join(HERE, '..', 'index.html')
open(out, 'w').write(html)
print(f"已写入 {os.path.abspath(out)} ({len(html)} 字符)。发布: git add -A && git commit && git push")
