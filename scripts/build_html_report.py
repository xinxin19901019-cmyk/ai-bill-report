#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从整理好的账单JSON生成一份好看的HTML消费报告(饼图+可优化项+待办)。
用法: python scripts/build_html_report.py input.json 消费报告.html
输入JSON字段同 build_report.py(见 references/input_schema.md)。只依赖标准库。"""
import sys, json, math, re

def num(s):
    m=re.search(r'-?\d[\d,]*\.?\d*', str(s).replace(',',''))
    return float(m.group()) if m else 0.0

def main():
    d=json.load(open(sys.argv[1],encoding='utf-8'))
    out=sys.argv[2] if len(sys.argv)>2 else '消费报告.html'
    month=d.get('month','')
    detail=d.get('detail',[])
    # 六类占比
    from collections import defaultdict
    cat=defaultdict(float)
    for r in detail: cat[r[3]]+=num(r[4])
    net=sum(cat.values())
    cats=sorted(cat.items(),key=lambda x:-x[1]) or [('无',0)]
    # 剔除总额
    exc=sum(abs(num(x[1])) for x in d.get('excluded',[]))
    rec=sum(num(x[1]) for x in d.get('reconciliation',[]))
    recv=sum(num(x[1]) for x in d.get('receivables',[]))
    剔除=exc+rec+recv
    # 可优化(年省)
    opt_a=d.get('optimize_actionable',[])
    年省=sum(num(x[3]) for x in opt_a)
    # 饼图
    colors=["#f97316","#fb923c","#fdba74","#fed7aa","#fde9d8","#e9d5c4","#f3ddc9"]
    total=sum(v for _,v in cats) or 1; ang=-90; paths=""; legend=""
    cx=cy=90;r=78;ir=46
    for i,(k,v) in enumerate(cats):
        frac=v/total;a2=ang+frac*360
        x1=cx+r*math.cos(math.radians(ang));y1=cy+r*math.sin(math.radians(ang))
        x2=cx+r*math.cos(math.radians(a2));y2=cy+r*math.sin(math.radians(a2))
        lg=1 if frac>0.5 else 0
        paths+=f'<path d="M {cx} {cy} L {x1:.1f} {y1:.1f} A {r} {r} 0 {lg} 1 {x2:.1f} {y2:.1f} Z" fill="{colors[i%len(colors)]}"/>'
        legend+=f'<div class="lg"><span class="dot" style="background:{colors[i%len(colors)]}"></span>{k} <b>{frac*100:.0f}%</b> ¥{v:,.0f}</div>'
        ang=a2
    donut=f'<svg viewBox="0 0 180 180" width="170" height="170">{paths}<circle cx="{cx}" cy="{cy}" r="{ir}" fill="#fff"/><text x="{cx}" y="{cy-4}" text-anchor="middle" font-size="12" fill="#9a8a80">净支出</text><text x="{cx}" y="{cy+14}" text-anchor="middle" font-size="15" font-weight="800" fill="#c2410c">¥{net:,.0f}</text></svg>'
    optrows="".join(f'<div class="opt"><div><b>{x[0]}</b> {x[1]}</div><div class="save">↓{x[3]}</div></div>' for x in opt_a) or '<div style="color:#9a8a80;font-size:13px">本月无明确可省项</div>'
    todorows="".join(f'▫️ {x[0]} ¥{num(x[1]):,.0f} — {x[2]}<br>' for x in d.get('receivables',[])) or '本月无待办'

    html=f'''<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>
*{{margin:0;padding:0;box-sizing:border-box;font-family:-apple-system,"PingFang SC",sans-serif}}
body{{background:linear-gradient(160deg,#fef6f0,#fdeee8);padding:26px 20px;color:#2b2b2b;max-width:520px;margin:0 auto}}
.title{{font-size:24px;font-weight:800;color:#c2410c}}.sub{{color:#9a8a80;font-size:13px;margin:4px 0 20px}}
.kpis{{display:flex;gap:10px;margin-bottom:18px}}.kpi{{flex:1;background:#fff;border-radius:15px;padding:14px 12px;box-shadow:0 4px 16px rgba(194,65,12,.08)}}
.kpi .n{{font-size:20px;font-weight:800;color:#c2410c}}.kpi.g .n{{color:#16a34a}}.kpi .t{{font-size:11px;color:#9a8a80;margin-top:3px}}
.card{{background:#fff;border-radius:17px;padding:18px;margin-bottom:14px;box-shadow:0 4px 16px rgba(194,65,12,.06)}}
.h{{font-size:16px;font-weight:700;margin-bottom:14px}}
.pie{{display:flex;gap:14px;align-items:center}}.lgs{{flex:1}}.lg{{font-size:13px;color:#555;margin:7px 0}}.lg b{{color:#c2410c}}
.dot{{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:6px;vertical-align:middle}}
.opt{{display:flex;justify-content:space-between;align-items:center;background:#fff7ed;border:1.5px solid #fed7aa;border-radius:12px;padding:12px 14px;margin:8px 0;font-size:13.5px;gap:10px}}
.opt b{{color:#c2410c}}.save{{color:#16a34a;font-weight:700;white-space:nowrap}}.todo{{font-size:14px;line-height:2;color:#444}}
</style></head><body>
<div class="title">🧾 消费报告</div><div class="sub">{month} · AI 自动生成</div>
<div class="kpis">
<div class="kpi"><div class="n">¥{net:,.0f}</div><div class="t">本月净支出</div></div>
<div class="kpi"><div class="n">¥{剔除:,.0f}</div><div class="t">已剔除·储蓄/还款/代垫/重复</div></div>
<div class="kpi g"><div class="n">¥{年省:,.0f}</div><div class="t">一年可省(可优化)</div></div></div>
<div class="card"><div class="h">消费去向</div><div class="pie">{donut}<div class="lgs">{legend}</div></div></div>
<div class="card"><div class="h">💡 可优化消费项</div>{optrows}</div>
<div class="card"><div class="h">✅ 待办</div><div class="todo">{todorows}</div></div>
</body></html>'''
    open(out,'w',encoding='utf-8').write(html)
    print(f"saved -> {out}")

if __name__=='__main__': main()
