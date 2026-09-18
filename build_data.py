# -*- coding: utf-8 -*-
"""
退费数据看板 · 本地数据更新脚本
用法：双击同目录下的 run_update.bat 即可（或 python build_data.py）
作用：
  1) 读取最新 Excel，遍历所有 sheet（每个 sheet = 一个渠道）
  2) 生成 channels.js（外部数据文件，供高级用途）
  3) 生成 dashboard.html（单文件看板：CSS + Chart.js + 数据 + 应用 JS 全部内嵌，
     不再依赖任何外部 .js，彻底规避 file:// 加载脚本失败的问题）
更新 Excel 后重跑脚本即可刷新看板。双击 dashboard.html 即可打开。
"""
import os, json, time, openpyxl, re
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(HERE, "退费渠道订单数据.xlsx")
CHANNELS_JS = os.path.join(HERE, "channels.js")
DASHBOARD_HTML = os.path.join(HERE, "dashboard.html")
CHART_JS_PATH = os.path.join(HERE, "chart.umd.min.js")

# === 按表头名称定位列（先精确匹配，再子串匹配，避免「创建日期」误匹配「日」） ===
def find_col(header, *keys):
    for k in keys:
        for i, h in enumerate(header):
            if h is not None and str(h) == k:
                return i
    for k in keys:
        for i, h in enumerate(header):
            if h is not None and k in str(h):
                return i
    return -1

def parse_sheet(matrix):
    if not matrix or len(matrix) < 2:
        return []
    header = [(c if c is not None else "") for c in matrix[0]]
    iDate = find_col(header, "创建日期")
    iY, iM, iD = find_col(header, "年"), find_col(header, "月"), find_col(header, "日")
    iPrice = find_col(header, "退费单价") if find_col(header, "退费单价") >= 0 else find_col(header, "单价")
    iServ, iChan = find_col(header, "服务商"), find_col(header, "渠道")
    iProd, iProv, iStatus = find_col(header, "产品名称"), find_col(header, "省份"), find_col(header, "订单状态")
    rows = []
    for r in matrix[1:]:
        if not r or all(c is None or c == "" for c in r):
            continue
        y = m = day = None
        date = ""
        if iDate >= 0 and r[iDate] is not None:
            date = str(r[iDate])[:10]
        # 优先从创建日期字符串里提取 y/m/day（最可靠，避免日期类型/字符串差异）
        if date:
            mt = re.search(r'(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})', date)
            if mt:
                try:
                    y, m, day = int(mt.group(1)), int(mt.group(2)), int(mt.group(3))
                except Exception:
                    pass
        # 退路：date 拿不到或未提取出时，用独立的 年/月/日 列（三列都必须存在且有数）
        if y is None and iY >= 0 and iM >= 0 and iD >= 0 \
           and r[iY] is not None and r[iM] is not None and r[iD] is not None:
            try:
                y, m, day = int(r[iY]), int(r[iM]), int(r[iD])
            except Exception:
                pass
        if not date and y and m:
            date = f"{y}-{m:02d}-{day if day else 1:02d}"
        # 必须有年月才入库；否则按月统计会出现 null-null / NaN 月
        if not y or not m:
            continue
        try:
            pr = float(r[iPrice]) if iPrice >= 0 else 0
        except Exception:
            pr = 0
        rows.append({
            "date": date,
            "y": y, "m": m, "day": day,
            "price": pr,
            "serv": r[iServ] if iServ >= 0 else None,
            "chan": r[iChan] if iChan >= 0 else None,
            "prod": r[iProd] if iProd >= 0 else None,
            "prov": r[iProv] if iProv >= 0 else None,
            "status": r[iStatus] if iStatus >= 0 else None,
        })
    return rows

# === 单文件看板：所有 CSS / Chart.js / 数据 / 应用 JS 都内嵌在 HTML 里 ===
# 冷色调商务风：深海军蓝 hero、KPI 用蓝/青/青绿/石板（无紫粉）

DASHBOARD_HEAD = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>退费渠道订单数据看板</title>
<style>
  :root {
    --bg: #f1f5f9;
    --card: #ffffff;
    --ink: #0f172a;
    --sub: #64748b;
    --line: #e2e8f0;
    --line-strong: #cbd5e1;
    --primary: #1e40af;
    --primary-hover: #1e3a8a;
    --cyan: #0e7490;
    --teal: #0f766e;
    --slate: #475569;
    --red: #dc2626;
    --green: #059669;
    --shadow: 0 1px 2px rgba(15,23,42,.04), 0 1px 4px rgba(15,23,42,.06);
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 0;
    background: var(--bg);
    color: var(--ink);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
    font-size: 14px;
    line-height: 1.5;
  }

  /* === Hero（深海军蓝商务风） === */
  .hero {
    background: linear-gradient(120deg, #0f172a 0%, #1e293b 50%, #1e3a8a 100%);
    color: #fff;
    border-radius: 0 0 14px 14px;
    padding: 24px 36px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 14px;
    box-shadow: 0 4px 16px rgba(15,23,42,.18);
  }
  .hero-left { display: flex; align-items: center; }
  .hero h1 {
    font-size: 26px;
    font-weight: 700;
    margin: 0;
    letter-spacing: 0.5px;
    line-height: 1.2;
  }
  .hero h1 small {
    display: block;
    font-size: 12px;
    font-weight: 400;
    opacity: 0.72;
    margin-top: 4px;
    letter-spacing: 0.3px;
  }
  .hero-meta { display: flex; flex-direction: column; align-items: flex-end; gap: 10px; flex-wrap: wrap; }
  .hero-pills { display: flex; gap: 10px; flex-wrap: wrap; justify-content: flex-end; }
  .hero-pill {
    display: inline-flex; align-items: center; gap: 6px;
    background: rgba(255,255,255,.08);
    border: 1px solid rgba(255,255,255,.14);
    padding: 7px 14px;
    border-radius: 18px;
    font-size: 12.5px;
    color: rgba(255,255,255,.92);
  }
  .hero-pill .ic { opacity: 0.85; }
  .hero-link {
    display: inline-flex; align-items: center; gap: 6px;
    background: rgba(255,255,255,.14);
    border: 1px solid rgba(255,255,255,.28);
    color: #fff; text-decoration: none;
    padding: 7px 16px; border-radius: 18px;
    font-size: 13px; font-weight: 600;
    cursor: pointer; transition: .12s;
    box-shadow: 0 1px 6px rgba(0,0,0,.12);
  }
  .hero-link:hover { background: #ffffff; color: var(--primary); border-color: #fff; }

  /* === 主区 === */
  .wrap { max-width: 1280px; margin: 0 auto; padding: 20px 24px 60px; }

  /* === 头部渠道按钮组 === */
  .hero-channels { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; justify-content: flex-end; max-width: 600px; }
  .hero-channels .lbl { font-size: 12.5px; color: rgba(255,255,255,.7); font-weight: 600; }
  .hero-chips { display: flex; gap: 6px; flex-wrap: wrap; justify-content: flex-end; }
  .hchip {
    display: inline-flex; align-items: center; gap: 6px;
    background: rgba(255,255,255,.10);
    border: 1px solid rgba(255,255,255,.18);
    border-radius: 16px;
    padding: 6px 13px;
    font-size: 13px;
    color: rgba(255,255,255,.92);
    cursor: pointer; user-select: none;
    transition: .12s;
  }
  .hchip:hover { background: rgba(255,255,255,.18); }
  .hchip.active {
    background: #ffffff;
    border-color: #ffffff;
    color: var(--primary);
    font-weight: 700;
    box-shadow: 0 2px 12px rgba(255,255,255,.35);
  }

  /* === 分组小标题 === */
  .section-title {
    display: flex; align-items: center;
    margin: 22px 2px 14px;
  }
  .section-title-bar {
    width: 3px; height: 16px;
    background: var(--primary);
    border-radius: 2px;
    margin-right: 10px;
  }
  .section-title-text { font-size: 14px; font-weight: 700; color: var(--ink); letter-spacing: 0.3px; }
  .section-title-sub { font-size: 12px; color: var(--sub); font-weight: 400; margin-left: 8px; }

  /* === KPI 卡片 === */
  .kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 16px; }
  @media (max-width: 880px) { .kpi-grid { grid-template-columns: repeat(2, 1fr); } }
  .kpi {
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 8px;
    padding: 18px 20px;
    position: relative;
    overflow: hidden;
    box-shadow: var(--shadow);
    transition: 0.15s;
  }
  .kpi:hover { box-shadow: 0 4px 16px rgba(15,23,42,.10); transform: translateY(-1px); }
  .kpi::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    background: var(--accent, var(--primary));
  }
  .kpi-label {
    display: flex; align-items: center; justify-content: space-between;
    font-size: 12.5px;
    color: var(--sub);
    margin-bottom: 10px;
    font-weight: 600;
  }
  .kpi-label .badge {
    background: var(--accent-bg);
    color: var(--accent);
    font-size: 10.5px;
    padding: 2px 8px;
    border-radius: 10px;
    font-weight: 700;
    letter-spacing: 0.5px;
  }
  .kpi-value {
    font-size: 30px;
    font-weight: 800;
    color: var(--accent, var(--primary));
    letter-spacing: -0.3px;
    line-height: 1.1;
    font-variant-numeric: tabular-nums;
    margin: 6px 0 12px;
  }
  .kpi-value small {
    font-size: 14px;
    font-weight: 600;
    color: var(--sub);
    margin-left: 3px;
  }
  .kpi-foot {
    font-size: 12px;
    color: var(--sub);
    display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
  }
  .kpi-foot .sep { color: var(--line-strong); }
  .kpi-foot .cmp {
    font-weight: 700;
    display: inline-flex; align-items: center; gap: 2px;
  }
  .kpi-foot .cmp.inc { color: var(--red); }
  .kpi-foot .cmp.dec { color: var(--green); }
  .kpi-foot .cmp.neu { color: var(--sub); }

  /* 冷色 accent 变体 */
  .kpi.c-blue { --accent: #1e40af; --accent-bg: #dbeafe; }
  .kpi.c-cyan { --accent: #0e7490; --accent-bg: #cffafe; }
  .kpi.c-teal { --accent: #0f766e; --accent-bg: #ccfbf1; }
  .kpi.c-slate { --accent: #475569; --accent-bg: #f1f5f9; }

  /* === 模块卡片 === */
  .card {
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 8px;
    padding: 20px 24px 22px;
    margin-bottom: 16px;
    box-shadow: var(--shadow);
  }
  .card-head {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 14px;
    gap: 10px; flex-wrap: wrap;
    padding-bottom: 12px;
    border-bottom: 1px solid var(--line);
  }
  .card-head h2 {
    font-size: 15px;
    font-weight: 700;
    margin: 0;
    display: inline-flex;
    align-items: center;
    gap: 10px;
    color: var(--ink);
    letter-spacing: 0.3px;
  }
  .card-head h2 .num {
    display: inline-flex; align-items: center; justify-content: center;
    width: 22px; height: 22px;
    background: var(--primary);
    color: #fff;
    border-radius: 5px;
    font-size: 12px;
    font-weight: 700;
  }
  .card-head .meta { font-size: 12px; color: var(--sub); }

  .grid2 { display: grid; grid-template-columns: 6fr 4fr; gap: 18px; }
  .grid2 .chartbox, .grid2 .tblbox { height: 240px; }
  .grid2 .chartbox { position: relative; }
  .grid2 canvas { width: 100% !important; height: 100% !important; }
  .grid2 .tblbox { overflow: auto; }
  @media (max-width: 880px) { .grid2 { grid-template-columns: 1fr; } .grid2 .chartbox, .grid2 .tblbox { height: 240px; } }

  table { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 4px; }
  th, td { padding: 10px 12px; text-align: right; border-bottom: 1px solid var(--line); }
  th:first-child, td:first-child { text-align: left; }
  thead th {
    background: #f8fafc; color: var(--sub);
    font-weight: 600; font-size: 12px; position: sticky; top: 0;
  }
  tbody tr:hover { background: #f8fafc; }
  td.num { font-variant-numeric: tabular-nums; font-weight: 600; }
  tbody tr.total {
    background: #eef2ff; font-weight: 700;
    border-top: 2px solid var(--primary);
  }
  tbody tr.total td { color: var(--ink); }
  tbody tr.total td.num { font-weight: 700; }
  td.rate { color: var(--primary); }
  td.rate.low { color: #dc2626; }

  .controls { display: flex; flex-wrap: wrap; gap: 14px 22px; align-items: flex-end; margin-bottom: 10px; }
  .ctl { display: flex; flex-direction: column; gap: 5px; }
  .ctl label { font-size: 12px; color: var(--sub); font-weight: 600; }
  input[type="date"] {
    padding: 8px 10px; border: 1px solid var(--line); border-radius: 6px;
    font-size: 13px; color: var(--ink); background: #fff;
    outline: none; transition: 0.12s;
  }
  input[type="date"]:focus { border-color: var(--primary); box-shadow: 0 0 0 3px rgba(30,64,175,.1); }

  .btn {
    background: var(--primary); color: #fff; border: none;
    border-radius: 6px; padding: 9px 18px; font-size: 13px;
    cursor: pointer; font-weight: 600;
    transition: 0.12s;
  }
  .btn:hover { background: var(--primary-hover); transform: translateY(-1px); box-shadow: 0 4px 12px rgba(30,64,175,.3); }

  .monthpick { display: flex; flex-wrap: wrap; gap: 8px; }
  .mchip {
    cursor: pointer; border: 1px solid var(--line); border-radius: 6px;
    padding: 7px 12px; font-size: 13px; background: #fff;
    transition: 0.12s; user-select: none;
  }
  .mchip:hover { border-color: var(--primary); color: var(--primary); }
  .mchip.active { background: var(--primary); color: #fff; border-color: var(--primary); }
  .mchip:disabled { opacity: 0.4; cursor: not-allowed; }

  .tag { font-size: 12px; color: var(--sub); font-weight: 500; }
  canvas { max-width: 100%; }

  .note {
    font-size: 12px; color: var(--sub);
    margin-top: 12px; padding: 8px 12px;
    background: #f8fafc; border-radius: 6px;
    border-left: 3px solid var(--primary);
  }

  .footer {
    text-align: center; color: var(--sub);
    font-size: 12px; margin-top: 24px; line-height: 1.8;
  }
  .footer code {
    background: #e2e8f0; padding: 2px 6px; border-radius: 4px;
    font-family: ui-monospace, Menlo, monospace; font-size: 11.5px;
    color: var(--ink);
  }

  /* === 加载/错误兜底 === */
  .bd-error {
    margin: 60px auto; max-width: 720px;
    background: #fff; border: 1px solid var(--line); border-left: 4px solid var(--red);
    border-radius: 6px; padding: 16px 20px;
    color: #7f1d1d; font-size: 13px;
    display: none;
  }
  .bd-error.show { display: block; }
  .bd-error b { color: var(--red); }
</style>
</head>
<body>
<div class="bd-error" id="bdError"></div>
<div class="hero">
  <div class="hero-left">
    <h1>
      退费渠道订单数据看板
      <small>按渠道查看退费笔数、金额、面额分布与日趋势</small>
    </h1>
  </div>
  <div class="hero-meta">
    <div class="hero-channels">
      <span class="lbl">渠道</span>
      <div class="hero-chips" id="channelChips"></div>
    </div>
    <div class="hero-pills">
      <a class="hero-link" href="compare.html">📊 渠道对比</a>
      <div class="hero-pill"><span class="ic">📅</span><span id="dataRange">—</span></div>
      <div class="hero-pill"><span class="ic">🕒</span>数据更新于 <span id="updatedAt" style="margin-left:4px">—</span></div>
    </div>
  </div>
</div>

<div class="wrap">
  <div class="section-title">
    <span class="section-title-bar"></span>
    <span class="section-title-text">退费概览</span>
    <span class="section-title-sub">（年累计 / 当月，单选渠道）</span>
  </div>
  <div class="kpi-grid" id="kpis"></div>

  <div class="card">
    <div class="card-head">
      <h2><span class="num">1</span>全年 / 每月 退费统计</h2>
      <div class="meta">按月汇总退费笔数、金额、日均与成功率（成功率 = 成功笔数 ÷ (成功+失败) 总笔数）</div>
    </div>
    <div class="grid2">
      <div class="chartbox"><canvas id="chartMonth"></canvas></div>
      <div class="tblbox">
        <table id="tblMonth">
          <thead><tr><th>月份</th><th>笔数</th><th>金额(元)</th><th>日均金额(元)</th><th>退费成功率</th></tr></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
    <div class="note">日均退费金额 = 当月退费总金额 ÷ 该月自然天数；当前进行中的月份按已过的天数（数据实际覆盖的最大日）计算。"全年"按实际数据区间展示。</div>
  </div>

  <div class="card">
    <div class="card-head">
      <h2><span class="num">2</span>退费日趋势</h2>
      <div class="meta">勾选 1–3 个月份；左图为退费金额、右图为退费笔数，横轴为 1–31 日，每月一条趋势线</div>
    </div>
    <div class="controls">
      <div class="ctl" style="flex:1;min-width:240px">
        <label>选择月份（最多 3 个）</label>
        <div class="monthpick" id="monthPick"></div>
      </div>
    </div>
    <div class="grid2" style="grid-template-columns:1fr 1fr; gap:18px; margin-top:14px">
      <div class="chartbox" style="height:240px"><canvas id="chartTrendAmt"></canvas></div>
      <div class="chartbox" style="height:240px"><canvas id="chartTrendCount"></canvas></div>
    </div>
  </div>

  <div class="card">
    <div class="card-head">
      <h2><span class="num">3</span>自定义时间段 · 各面额退费统计</h2>
      <div class="meta">按起止日期筛选，统计每种退费面额的笔数、金额、占比与日均</div>
    </div>
    <div class="controls">
      <div class="ctl"><label>开始日期</label><input type="date" id="dateFrom"></div>
      <div class="ctl"><label>结束日期</label><input type="date" id="dateTo"></div>
      <div class="ctl"><label>&nbsp;</label><button class="btn" id="btnRange">应用筛选</button></div>
      <div class="ctl" style="margin-left:auto"><label>&nbsp;</label><span class="tag" id="rangeInfo"></span></div>
    </div>
    <div style="margin-top:14px">
      <div style="overflow:auto;max-height:340px">
        <table id="tblDenom">
          <thead><tr><th>退费面额</th><th>笔数</th><th>金额(元)</th><th>金额占比</th><th>日均金额(元)</th></tr></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
    <div class="note">日均退费金额 = 该面额退费总金额 ÷ 时间范围自然天数（含首尾）。</div>
  </div>

  <div class="footer" id="footer"></div>
</div>
'''

DASHBOARD_APP_JS = r'''
let CHANNELS = [];
let selectedChannel = null;
let charts = {};

const fmt = n => (n==null?'0':Number(n).toLocaleString('zh-CN'));
const fmtMoney = n => (n==null?'0':Number(n).toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2}));
function daysBetween(a,b){
  const d1=new Date(a+'T00:00:00'), d2=new Date(b+'T00:00:00');
  return Math.round((d2-d1)/86400000)+1;
}
// 现有统计（笔数/金额/日均/面额/趋势）只统计「成功」订单，与历史口径一致；
// status 为空（旧数据无此列）也按成功处理，避免历史数据被误删
function curRows(){ const c = CHANNELS.find(c=>c.name===selectedChannel); return c ? c.rows.filter(r => r.status===null || r.status==='成功') : []; }
// 保留全部记录（含失败），供后续「成功率」对比使用
function curRowsAll(){ const c = CHANNELS.find(c=>c.name===selectedChannel); return c ? c.rows : []; }
function allDates(onlySucc=true){ const src = onlySucc ? curRows() : curRowsAll(); const s=new Set(); src.forEach(r=>s.add(r.date)); return [...s].sort(); }

function showError(msg){
  const el = document.getElementById('bdError');
  if (!el) return;
  el.innerHTML = '<b>⚠ 渲染异常</b><br>' + msg;
  el.classList.add('show');
}

window.addEventListener('error', (e) => {
  showError((e.message || '未知错误') + (e.filename ? ' (' + e.filename + ':' + e.lineno + ')' : ''));
});

function renderChannelChips(){
  const box = document.getElementById('channelChips'); if(!box) return; box.innerHTML='';
  CHANNELS.forEach(c=>{
    const el = document.createElement('div');
    el.className = 'hchip' + (selectedChannel===c.name?' active':'');
    const label = c.displayName || c.name;
    el.title = c.name;
    el.innerHTML = `${label}`;
    el.onclick = () => { selectedChannel = c.name; renderAll(); };
    box.appendChild(el);
  });
}

function renderKPIs(){
  const rows = curRows();
  const yAmt = rows.reduce((s,r) => s + (+r.price || 0), 0);
  const yCnt = rows.length;
  let maxY = 0, maxM = 0;
  rows.forEach(r => { if (r.y && (r.y>maxY || (r.y===maxY && r.m>maxM))) { maxY=r.y; maxM=r.m; } });
  const mRows = rows.filter(r => r.y===maxY && r.m===maxM);
  const mAmt = mRows.reduce((s,r) => s + (+r.price || 0), 0);
  const mCnt = mRows.length;

  let prevY = null, prevM = null;
  if (maxM === 1) { prevY = maxY - 1; prevM = 12; }
  else { prevY = maxY; prevM = maxM - 1; }
  // 同期环比：上月仅取到与当前月相同的最新日（如当前月覆盖到 19 号，上月也只比到 19 号）
  const curMaxDay = mRows.reduce((d, r) => Math.max(d, r.day || 0), 0);
  const pRows = rows.filter(r => r.y===prevY && r.m===prevM && (r.day || 0) <= curMaxDay);
  const pAmt = pRows.reduce((s,r) => s + (+r.price || 0), 0);
  const pCnt = pRows.length;

  function pctStr(cur, prev){
    if (!prev) return {text:'—', dir:'neu', has:false};
    const pct = (cur - prev) / prev * 100;
    const sign = pct > 0 ? '+' : '';
    return {
      text: `${sign}${pct.toFixed(1)}%`,
      dir: pct > 0 ? 'inc' : pct < 0 ? 'dec' : 'neu',
      has: true
    };
  }
  const cmpAmt = pctStr(mAmt, pAmt);
  const cmpCnt = pctStr(mCnt, pCnt);

  const ym = maxY ? `${maxY}-${String(maxM).padStart(2,'0')}` : '—';
  // 数据最新日期（精确到日），供年累计卡显示
  const maxDate = rows.length ? [...rows].map(r => r.date).filter(Boolean).sort().pop() : '';
  const ytdLabel = maxDate || ym;
  const cmpA = cmpAmt.has ? `<span class="sep">|</span><span>同期环比 <span class="cmp ${cmpAmt.dir}">${cmpAmt.text}</span></span>` : '';
  const cmpC = cmpCnt.has ? `<span class="sep">|</span><span>同期环比 <span class="cmp ${cmpCnt.dir}">${cmpCnt.text}</span></span>` : '';

  const box = document.getElementById('kpis'); if (!box) return;
  box.innerHTML = `
    <div class="kpi c-blue">
      <div class="kpi-label"><span>年累计退费金额</span><span class="badge">YTD</span></div>
      <div class="kpi-value">¥${fmtMoney(yAmt)}</div>
      <div class="kpi-foot"><span>累计至 ${ytdLabel}</span></div>
    </div>
    <div class="kpi c-cyan">
      <div class="kpi-label"><span>年累计退费笔数</span><span class="badge">YTD</span></div>
      <div class="kpi-value">${fmt(yCnt)} <small>笔</small></div>
      <div class="kpi-foot"><span>累计至 ${ytdLabel}</span></div>
    </div>
    <div class="kpi c-teal">
      <div class="kpi-label"><span>当月退费金额</span><span class="badge">M</span></div>
      <div class="kpi-value">¥${fmtMoney(mAmt)}</div>
      <div class="kpi-foot"><span>当月 ${ym}</span>${cmpA}</div>
    </div>
    <div class="kpi c-slate">
      <div class="kpi-label"><span>当月退费笔数</span><span class="badge">M</span></div>
      <div class="kpi-value">${fmt(mCnt)} <small>笔</small></div>
      <div class="kpi-foot"><span>当月 ${ym}</span>${cmpC}</div>
    </div>
  `;
}

function renderMonth(){
  const rows = curRows();          // 仅成功
  const all = curRowsAll();         // 成功 + 失败（成功率分母用）
  const map = {};
  rows.forEach(r => { const k = r.y + '-' + String(r.m).padStart(2,'0');
    if (!map[k]) map[k] = {count:0, amt:0, maxDay:0};
    map[k].count++; map[k].amt += (+r.price || 0);
    if (r.day && r.day > map[k].maxDay) map[k].maxDay = r.day; });
  const allMap = {};
  all.forEach(r => { const k = r.y + '-' + String(r.m).padStart(2,'0');
    allMap[k] = allMap[k] || {total:0, succ:0};
    allMap[k].total++; if ((r.status||'') === '成功') allMap[k].succ++; });
  const keys = Object.keys(map).sort();
  const tb = document.querySelector('#tblMonth tbody'); if (tb) tb.innerHTML = '';
  const labels = [], cnts = [], amts = [], rates = [];
  const now = new Date();
  const curY = now.getFullYear(), curM = now.getMonth() + 1;
  keys.forEach(k => {
    const [y, m] = k.split('-');
    const dInM = new Date(+y, +m, 0).getDate();
    // 当月（进行中）按「数据实际覆盖到的最大日」算日均（≤今日）；
    // 历史月仍用整月天数，保证口径一致
    const isCurMonth = +y === curY && +m === curM;
    const denom = (isCurMonth && map[k].maxDay > 0) ? map[k].maxDay : dInM;
    const daily = map[k].amt / denom;
    const ar = allMap[k] || {total:0, succ:0};
    const rate = ar.total ? ar.succ / ar.total * 100 : 0;
    if (tb) tb.innerHTML += `<tr><td>${k}</td><td class="num">${fmt(map[k].count)}</td>
      <td class="num">${fmtMoney(map[k].amt)}</td>
      <td class="num">${fmtMoney(daily)}</td>
      <td class="num rate${rate < 90 ? ' low' : ''}">${rate.toFixed(1)}%</td></tr>`;
    labels.push(k); cnts.push(map[k].count); amts.push(map[k].amt); rates.push(+rate.toFixed(1));
  });
  drawCombo('chartMonth', labels, cnts, amts, rates, '笔数', '金额(元)', '成功率(%)');
}

function drawCombo(id, labels, cnts, amts, rates, lab1, lab2, lab3){
  const canvas = document.getElementById(id); if (!canvas) return;
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(canvas, {
    data: { labels, datasets: [
      { type:'bar', label: lab1, data: cnts, backgroundColor:'#93c5fd', borderRadius:6, yAxisID:'y', order:2 },
      { type:'line', label: lab2, data: amts, borderColor:'#1e40af', backgroundColor:'#1e40af',
        tension:.35, pointRadius:4, borderWidth:2.5, yAxisID:'y1', order:1 },
      { type:'line', label: lab3, data: rates, borderColor:'#f59e0b', backgroundColor:'#f59e0b',
        tension:.35, pointRadius:4, borderWidth:2.5, yAxisID:'y2', order:0,
        tooltip:{ callbacks:{ label: c => `${c.dataset.label}: ${(c.parsed.y).toFixed(1)}%` } } },
    ]},
    options: { responsive:true, maintainAspectRatio:false, interaction:{mode:'index', intersect:false},
      plugins: { legend:{position:'top', labels:{usePointStyle:true, padding:15}} },
      scales: {
        y: { position:'left', title:{display:true, text:lab1}, grid:{color:'#e2e8f0'}, beginAtZero:true },
        y1: { position:'right', title:{display:true, text:lab2}, grid:{drawOnChartArea:false}, beginAtZero:true },
        y2: { position:'right', title:{display:true, text:lab3}, grid:{drawOnChartArea:false},
          beginAtZero:true, max:100, suggestedMin:0,
          ticks:{ callback: v => v + '%' } }
      }}
  });
}

function renderMonthPick(){
  const rows = curRows();
  const months = [...new Set(rows.map(r => r.y + '-' + String(r.m).padStart(2,'0')))].sort();
  const box = document.getElementById('monthPick'); if (!box) return; box.innerHTML = '';
  if (!window._selMonths) window._selMonths = months.slice(-1);  // 默认选中最新月份（当月）
  window._selMonths = window._selMonths.filter(m => months.includes(m));
  if (window._selMonths.length === 0 && months.length) window._selMonths = [months[months.length - 1]];
  months.forEach(m => {
    const el = document.createElement('div');
    const dis = (!window._selMonths.includes(m)) && window._selMonths.length >= 3;
    el.className = 'mchip' + (window._selMonths.includes(m) ? ' active' : '');
    el.textContent = m;
    el.disabled = dis;
    el.onclick = () => {
      const i = window._selMonths.indexOf(m);
      if (i >= 0) { if (window._selMonths.length > 1) window._selMonths.splice(i, 1); }
      else { if (window._selMonths.length < 3) window._selMonths.push(m); }
      renderMonthPick(); renderTrend();
    };
    box.appendChild(el);
  });
}

function renderTrend(){
  const rows = curRows(); const sel = window._selMonths || [];
  const labels = Array.from({length:31}, (_, i) => (i+1) + '日');
  // 高对比分类色：蓝 / 橙 / 绿 / 青绿，相邻月份一目了然（避开蓝紫相近）
  const palette = ['#2563eb', '#f59e0b', '#16a34a', '#14b8a6'];
  // 每个所选月份一条线：按 day-of-month(1-31) 聚合
  const series = sel.map(mk => {
    const [y, m] = mk.split('-').map(Number);
    const amt = new Array(31).fill(0), cnt = new Array(31).fill(0);
    rows.forEach(r => {
      if (r.y === y && r.m === m) { const d = (r.day || 0); if (d >= 1 && d <= 31) { amt[d-1] += (+r.price || 0); cnt[d-1]++; } }
    });
    return { label: m + '月', amt, cnt };
  });
  const dsFor = (key) => series.map((s, i) => ({
    label: s.label, data: s[key], borderColor: palette[i % palette.length],
    backgroundColor: palette[i % palette.length], fill:false, tension:.35, pointRadius:3.5, borderWidth:3
  }));
  const mkOpts = (yTitle) => ({
    responsive:true, maintainAspectRatio:false,
    interaction:{ mode:'index', intersect:false },
    plugins:{ legend:{ display:true, position:'top', labels:{ usePointStyle:true, padding:12 } } },
    scales:{
      y:{ beginAtZero:true, grid:{color:'#e2e8f0'}, title:{display:true, text:yTitle} },
      x:{ grid:{display:false}, title:{display:true, text:'日'} }
    }
  });
  // 左图：退费金额
  const am = document.getElementById('chartTrendAmt');
  if (am) {
    if (charts.chartTrendAmt) charts.chartTrendAmt.destroy();
    charts.chartTrendAmt = new Chart(am, {
      type:'line', data:{ labels, datasets: dsFor('amt') }, options: mkOpts('退费金额(元)')
    });
  }
  // 右图：退费笔数
  const cn = document.getElementById('chartTrendCount');
  if (cn) {
    if (charts.chartTrendCount) charts.chartTrendCount.destroy();
    charts.chartTrendCount = new Chart(cn, {
      type:'line', data:{ labels, datasets: dsFor('cnt') }, options: mkOpts('退费笔数')
    });
  }
}

function renderDenom(){
  const fromEl = document.getElementById('dateFrom');
  const toEl = document.getElementById('dateTo');
  const from = fromEl ? fromEl.value : '';
  const to = toEl ? toEl.value : '';
  const rows = curRows().filter(r => (!from || r.date >= from) && (!to || r.date <= to));
  const info = document.getElementById('rangeInfo');
  if (info) {
    if (from && to && from > to) info.textContent = '⚠ 开始日期晚于结束日期';
    else info.textContent = `区间 ${from||'起'} ~ ${to||'止'} ｜ 共 ${fmt(rows.length)} 笔`;
  }
  const dInRange = (from && to) ? daysBetween(from, to) : (new Set(rows.map(r => r.date)).size || 1);
  const map = {};
  rows.forEach(r => { const p = (+r.price || 0); if (!map[p]) map[p] = {count:0, amt:0}; map[p].count++; map[p].amt += p; });
  const keys = Object.keys(map).map(Number).sort((a,b) => a-b);
  const totalAmt = rows.reduce((s,r) => s + (+r.price || 0), 0);
  const tb = document.querySelector('#tblDenom tbody'); if (tb) tb.innerHTML = '';
  keys.forEach((p) => {
    const daily = map[p].amt / dInRange;
    if (tb) tb.innerHTML += `<tr><td>${p} 元</td><td class="num">${fmt(map[p].count)}</td>
      <td class="num">${fmtMoney(map[p].amt)}</td>
      <td class="num">${totalAmt ? (map[p].amt/totalAmt*100).toFixed(1) : 0}%</td>
      <td class="num">${fmtMoney(daily)}</td></tr>`;
  });
  const totalCount = rows.length;
  const totalDaily = totalAmt / dInRange;
  if (tb) tb.innerHTML += `<tr class="total"><td>合计</td><td class="num">${fmt(totalCount)}</td>`
    + `<td class="num">${fmtMoney(totalAmt)}</td>`
    + `<td class="num">100%</td>`
    + `<td class="num">${fmtMoney(totalDaily)}</td></tr>`;
}

function updateHeroMeta(){
  // 头部数据范围按「该渠道全部记录」（含失败/其他）的创建日期，
  // 才能看出各渠道数据实际更新到哪一天（成功订单可能滞后于数据本身）
  const dates = allDates(false);
  const range = document.getElementById('dataRange');
  if (range) range.textContent = dates.length ? `${dates[0]} ~ ${dates[dates.length-1]}` : '—';
  const upd = document.getElementById('updatedAt');
  if (upd) {
    if (window.CHANNEL_META && window.CHANNEL_META.generatedAt) {
      upd.textContent = window.CHANNEL_META.generatedAt;
    } else {
      const d = new Date();
      const pad = n => String(n).padStart(2,'0');
      upd.textContent = `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
    }
  }
}

function renderAll(){
  renderChannelChips();
  renderKPIs();
  renderMonth();
  renderMonthPick();
  renderTrend();
  renderDenom();
  updateHeroMeta();
}

function initData(){
  try {
    CHANNELS = (window.CHANNELS && window.CHANNELS.length) ? window.CHANNELS : [];
    selectedChannel = CHANNELS.length ? CHANNELS[0].name : null;
    const dates = allDates();
    const df = document.getElementById('dateFrom'), dt = document.getElementById('dateTo');
    if (df && dt && dates.length) {
      // 开始日期默认取数据最新月份的 1 日（当月1日），结束日期取最新日期
      const months = [...new Set(dates.map(d => d.slice(0,7)))].sort();
      df.value = months[months.length - 1] + '-01';
      dt.value = dates[dates.length - 1];
    }
    window._selMonths = null;
    renderAll();
    const footer = document.getElementById('footer');
    if (footer) footer.innerHTML = `看板由 WorkBuddy 生成 ｜ 更新数据：编辑 Excel 后双击运行同目录 <code>run_update.bat</code>，再刷新本页`;
  } catch (e) {
    showError('初始化失败：' + (e.message || e));
  }
}

const btnR = document.getElementById('btnRange');
if (btnR) btnR.onclick = renderDenom;
initData();
'''

def build_dashboard(channels, chart_js_content, generated_at):
    """把 HTML 头 + Chart.js + 数据 + 应用 JS 全部内嵌到一个文件"""
    data_json = json.dumps(channels, ensure_ascii=False)
    meta_json = json.dumps({"generatedAt": generated_at})
    return (
        DASHBOARD_HEAD +
        '\n<script>\n/* === Chart.js (内嵌，避免 file:// 加载失败) === */\n'
        + chart_js_content +
        '\n</script>\n'
        '<script>\n'
        '/* === 自动注入的数据 === */\n'
        f'window.CHANNEL_META = {meta_json};\n'
        f'window.CHANNELS = {data_json};\n'
        '</script>\n'
        '<script>\n/* === 应用逻辑 === */\n'
        + DASHBOARD_APP_JS +
        '\n</script>\n</body>\n</html>\n'
    )


# ===================== 多渠道对比页（compare.html） =====================
COMPARE_BODY = r'''
<body>
<div class="bd-error" id="bdError"></div>
<div class="hero">
  <div class="hero-left">
    <h1>退费渠道数据对比<small>多渠道横向对比 · 年累计 / 当月 / 月度明细 / 日趋势</small></h1>
  </div>
  <div class="hero-pills">
    <a class="hero-link" href="dashboard.html">📄 单渠道看板</a>
    <div class="hero-pill"><span class="ic">🕒</span>数据更新于 <span id="updatedAt" style="margin-left:4px">—</span></div>
  </div>
</div>

<div class="wrap">
  <div class="card">
    <div class="card-head">
      <h2><span class="num">1</span>渠道总览对比</h2>
      <div class="meta" id="ovMeta">各渠道 年累计 / 当月累计 退费笔数与金额</div>
    </div>
    <div class="grid2" style="grid-template-columns:1fr 320px; gap:18px; margin-top:14px; align-items:start">
      <div style="overflow:auto; max-height:340px">
        <table id="tblOverview">
          <thead><tr><th>渠道</th><th>年累计笔数</th><th>年累计金额(元)</th><th>当月累计笔数</th><th>当月累计金额(元)</th></tr></thead>
          <tbody></tbody>
        </table>
      </div>
      <div style="height:320px; min-width:0; display:flex; flex-direction:column">
        <div class="meta" style="text-align:center; margin-bottom:6px">年累计退费金额占比</div>
        <div style="flex:1; min-height:0; position:relative"><canvas id="chartOvShare"></canvas></div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-head">
      <h2><span class="num">2</span>各渠道月度明细</h2>
      <div class="meta">每月退费笔数 / 金额 / 成功率(成功÷总笔数) / 金额占比(占当月各渠道总额)</div>
    </div>
    <div class="monthpick" id="cmpMetricPick" style="margin:12px 0 4px"></div>
    <div class="chartbox" style="height:300px"><canvas id="chartMonthCmp"></canvas></div>
    <div style="overflow:auto;max-height:440px; margin-top:14px">
      <table id="tblMonthDetail">
        <thead><tr><th>渠道</th><th>月份</th><th>退费笔数</th><th>退费金额(元)</th><th>退费成功率</th><th>退费金额占比</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>

  <div class="card">
    <div class="card-head">
      <h2><span class="num">3</span>各渠道日趋势</h2>
      <div class="meta">左：退费金额 右：退费笔数，每渠道一条线（仅成功记录，缺失日期不连线）</div>
    </div>
    <div class="monthpick" id="cmpMonthPick" style="margin:12px 0 4px"></div>
    <div class="grid2" style="grid-template-columns:1fr 1fr; gap:18px; margin-top:10px">
      <div class="chartbox" style="height:240px"><canvas id="chartTrendAmt"></canvas></div>
      <div class="chartbox" style="height:240px"><canvas id="chartTrendCnt"></canvas></div>
    </div>
  </div>

  <div class="footer" id="footer"></div>
</div>
'''

COMPARE_APP_JS = r'''
const DATA = window.COMPARE || {m1:[], m2:[], m3:{dates:[], series:[]}};
const PALETTE = ['#2563eb', '#f59e0b', '#16a34a', '#dc2626'];
// 渠道固定配色：文航红 / 博迈橙 / 自由行蓝 / 瑞牛绿（模块②③统一，按渠道名取色，不随顺序变化）
const CHANNEL_COLORS = { '文航':'#dc2626', '博迈':'#f59e0b', '自由行':'#2563eb', '瑞牛':'#16a34a' };
const colorOf = name => CHANNEL_COLORS[name] || PALETTE[0];
const charts = {};
const fmt = n => (n==null?'0':Number(n).toLocaleString('zh-CN'));
const fmtMoney = n => Number(n||0).toLocaleString('zh-CN', {minimumFractionDigits:0, maximumFractionDigits:2});

function renderOverview(){
  const m1 = DATA.m1;
  const curLabel = (m1[0] && m1[0].mLabel) || '—';
  const metaEl = document.getElementById('ovMeta');
  if (metaEl) metaEl.textContent = `当前月：${curLabel} · 各渠道年累计 / 当月累计 退费笔数与金额`;
  const tb = document.querySelector('#tblOverview tbody'); if (tb) tb.innerHTML='';
  let tYc=0, tYa=0, tMc=0, tMa=0;
  m1.forEach(r=>{
    if (tb) tb.innerHTML += `<tr><td>${r.name}</td><td class="num">${fmt(r.yCnt)}</td>`
      + `<td class="num">${fmtMoney(r.yAmt)}</td><td class="num">${fmt(r.mCnt)}</td>`
      + `<td class="num">${fmtMoney(r.mAmt)}</td></tr>`;
    tYc += r.yCnt; tYa += r.yAmt; tMc += r.mCnt; tMa += r.mAmt;
  });
  if (tb) tb.innerHTML += `<tr class="total"><td>合计</td><td class="num">${fmt(tYc)}</td>`
    + `<td class="num">${fmtMoney(tYa)}</td><td class="num">${fmt(tMc)}</td>`
    + `<td class="num">${fmtMoney(tMa)}</td></tr>`;
  // 年累计退费金额占比环形图（渠道固定配色，与模块②③一致）
  const labels = m1.map(r=>r.name);
  const data = m1.map(r=>+r.yAmt);
  const colors = m1.map(r=>colorOf(r.name));
  drawDoughnut('chartOvShare', labels, data, colors);
}

function drawDoughnut(id, labels, data, colors){
  const canvas = document.getElementById(id); if(!canvas) return;
  if (charts[id]) charts[id].destroy();
  const total = data.reduce((a,b)=>a+b, 0) || 1;
  charts[id] = new Chart(canvas, {
    type:'doughnut',
    data:{ labels, datasets:[{ data, backgroundColor:colors, borderWidth:2, borderColor:'#fff' }] },
    options:{ responsive:true, maintainAspectRatio:false, cutout:'58%',
      layout:{ padding:{ top:4, bottom:4, left:4, right:4 } },
      plugins:{ legend:{position:'bottom', labels:{usePointStyle:true, padding:8, boxWidth:10, font:{size:11}}},
        tooltip:{ callbacks:{ label: c => {
          const v = c.parsed; const pct = (v/total*100).toFixed(1);
          return `${c.label}: ¥${fmtMoney(v)} (${pct}%)`;
        } } } } }
  });
}

function drawGroupedBar(id, labels, datasets, yTitle, yMax, fmtFn){
  const canvas = document.getElementById(id); if(!canvas) return;
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(canvas, {
    type:'bar',
    data:{ labels, datasets },
    options:{ responsive:true, maintainAspectRatio:false, interaction:{mode:'index', intersect:false},
      plugins:{ legend:{position:'top', labels:{usePointStyle:true, padding:12}},
        tooltip:{ callbacks:{ label: c => `${c.dataset.label}: ${fmtFn ? fmtFn(c.parsed.y) : fmt(c.parsed.y)}` } } },
      scales:{ x:{ grid:{display:true, color:'#e2e8f0', drawTicks:true, lineWidth:1}, ticks:{autoSkip:false, maxRotation:0, minRotation:0} },
        y:{ beginAtZero:true, max: yMax, grid:{color:'#e2e8f0'}, title:{display:true, text:yTitle} } } }
  });
}

function renderMonthDetail(){
  const m2 = DATA.m2;
  const tb = document.querySelector('#tblMonthDetail tbody'); if (tb) tb.innerHTML='';
  let lastCh = null;
  m2.forEach(r=>{
    const sep = (r.name!==lastCh) ? ' class="ch-sep"' : '';
    lastCh = r.name;
    if (tb) tb.innerHTML += `<tr${sep}><td>${r.name}</td><td>${r.month}</td>`
      + `<td class="num">${fmt(r.cnt)}</td><td class="num">${fmtMoney(r.amt)}</td>`
      + `<td class="num rate${r.rate<90?' low':''}">${r.rate}%</td>`
      + `<td class="num">${r.share}%</td></tr>`;
  });
}

const METRICS = [
  {key:'cnt',   label:'退费笔数',     unit:'笔',  ymax:null, fmt:n=>fmt(n)+'笔'},
  {key:'amt',   label:'退费金额',     unit:'元',  ymax:null, fmt:n=>'¥'+fmtMoney(n)},
  {key:'rate',  label:'退费成功率',   unit:'%',   ymax:100,  fmt:n=>n+'%'},
  {key:'share', label:'退费金额占比', unit:'%',   ymax:100,  fmt:n=>n+'%'},
];
let cmpMetric = 'amt';

function renderMetricSwitch(){
  const box = document.getElementById('cmpMetricPick'); if (!box) return;
  box.innerHTML = '';
  METRICS.forEach(m=>{
    const b = document.createElement('button');
    b.className = 'mchip' + (m.key===cmpMetric ? ' active':'');
    b.textContent = m.label;
    b.onclick = ()=>{ cmpMetric = m.key; renderMetricSwitch(); renderMonthCmp(); };
    box.appendChild(b);
  });
}

function renderMonthCmp(){
  const m2 = DATA.m2;
  const months = [...new Set(m2.map(r=>r.month))].sort();
  const chans = [...new Set(m2.map(r=>r.name))];
  const lookup = {};
  m2.forEach(r=>{ lookup[r.name+'|'+r.month] = r; });
  const m = METRICS.find(x=>x.key===cmpMetric) || METRICS[0];
  const datasets = chans.map((name,i)=>({
    label: name,
    data: months.map(mk => { const r = lookup[name+'|'+mk]; return r ? r[m.key] : null; }),
    backgroundColor: colorOf(name),
    borderRadius: 4,
  }));
  drawGroupedBar('chartMonthCmp', months, datasets, m.label + (m.unit==='%'?' (%)':''), m.ymax, m.fmt);
}

function renderTrend(){
  const m3 = DATA.m3;
  const mk = cmpSelMonth || (m3.months && m3.months[m3.months.length-1]);
  const block = (mk && m3.byMonth && m3.byMonth[mk]) || {dates:[], series:[]};
  const labels = block.dates;
  const mk2 = (key) => block.series.map((s,i)=>({
    label:s.name, data:s[key], borderColor:colorOf(s.name),
    backgroundColor:colorOf(s.name), tension:.3, borderWidth:2,
    // 有效点很少时（如月初仅 1-2 天有数据）显示点标记，否则单点连不成线会完全看不见
    pointRadius: ctx => {
      const arr = ((ctx.dataset && ctx.dataset.data) || []).filter(v => v !== null && v !== undefined);
      return arr.length <= 2 ? 4 : 0;
    },
    pointHoverRadius: 4
  }));
  drawLine('chartTrendAmt', labels, mk2('amt'), '退费金额(元)');
  drawLine('chartTrendCnt', labels, mk2('cnt'), '退费笔数');
}

let cmpSelMonth = null;
function renderMonthSwitch(){
  const box = document.getElementById('cmpMonthPick'); if (!box) return;
  const months = DATA.m3.months || [];
  box.innerHTML = '';
  if (!months.length) return;
  if (!cmpSelMonth || !months.includes(cmpSelMonth)) cmpSelMonth = months[months.length-1];
  months.forEach(mk=>{
    const b = document.createElement('button');
    b.className = 'mchip' + (mk===cmpSelMonth ? ' active':'');
    b.textContent = mk + '月';
    b.onclick = ()=>{ cmpSelMonth = mk; renderMonthSwitch(); renderTrend(); };
    box.appendChild(b);
  });
}

function drawLine(id, labels, datasets, yTitle){
  const canvas = document.getElementById(id); if(!canvas) return;
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(canvas, {
    type:'line',
    data:{ labels, datasets },
    options:{ responsive:true, maintainAspectRatio:false, interaction:{mode:'index', intersect:false},
      plugins:{ legend:{position:'top', labels:{usePointStyle:true, padding:10}},
        tooltip:{ callbacks:{ label: c => `${c.dataset.label}: ${fmt(c.parsed.y)}` } } },
      scales:{ x:{ ticks:{ autoSkip:true, maxTicksLimit:12, maxRotation:0 } },
        y:{ beginAtZero:true, grid:{color:'#e2e8f0'}, title:{display:true, text:yTitle} } } }
  });
}

function updateMeta(){
  const upd = document.getElementById('updatedAt');
  if (upd && window.CHANNEL_META && window.CHANNEL_META.generatedAt) upd.textContent = window.CHANNEL_META.generatedAt;
  const footer = document.getElementById('footer');
  if (footer) footer.innerHTML = '看板由 WorkBuddy 生成 ｜ 更新数据：编辑 Excel 后双击运行同目录 run_update.bat，再刷新本页';
}
try { renderOverview(); renderMonthDetail(); renderMetricSwitch(); renderMonthCmp(); renderMonthSwitch(); renderTrend(); updateMeta(); }
catch(e){ console.error(e); }
'''

def build_compare_data(channels):
    """从各渠道全量记录（含成功/失败）预聚合出对比页所需数据"""
    # 模块1：年累计 / 当月累计 笔数 & 金额（仅成功计入退费口径）
    # "当月" = 全局最新数据月（所有渠道成功记录中的最大年-月）；某渠道该月无成功记录则记 0
    all_ym = set()
    for c in channels:
        for r in c["rows"]:
            if r.get("status") in (None, "成功") and r.get("y") and r.get("m"):
                all_ym.add((r["y"], r["m"]))
    cur_ym = max(all_ym) if all_ym else None
    cur_label = f"{cur_ym[0]}-{cur_ym[1]:02d}" if cur_ym else "—"
    m1 = []
    for c in channels:
        succ = [r for r in c["rows"] if r.get("status") in (None, "成功")]
        y_cnt = len(succ)
        y_amt = sum(r["price"] for r in succ)
        if cur_ym:
            mrows = [r for r in succ if r["y"] == cur_ym[0] and r["m"] == cur_ym[1]]
            m_cnt = len(mrows); m_amt = sum(r["price"] for r in mrows)
        else:
            m_cnt = m_amt = 0
        m1.append({"name": c["displayName"], "yCnt": y_cnt, "yAmt": round(y_amt, 2),
                   "mCnt": m_cnt, "mAmt": round(m_amt, 2), "mLabel": cur_label})

    # 模块2：各渠道各月 笔数/金额/成功率/金额占比（占比=该月该渠道成功金额÷当月各渠道成功总额）
    month_total = {}
    for c in channels:
        for r in c["rows"]:
            if r.get("status") in (None, "成功") and r.get("y") and r.get("m"):
                k = (r["y"], r["m"])
                month_total[k] = month_total.get(k, 0) + r["price"]
    m2 = []
    for c in channels:
        mp = {}
        for r in c["rows"]:
            if not (r.get("y") and r.get("m")):
                continue
            k = (r["y"], r["m"])
            mp.setdefault(k, {"cnt": 0, "amt": 0.0, "total": 0})
            mp[k]["total"] += 1
            if r.get("status") in (None, "成功"):
                mp[k]["cnt"] += 1
                mp[k]["amt"] += r["price"]
        for k in sorted(mp.keys()):
            v = mp[k]
            rate = v["total"] and v["cnt"] / v["total"] * 100 or 0
            share = month_total.get(k, 0) and v["amt"] / month_total[k] * 100 or 0
            m2.append({"name": c["displayName"], "month": f"{k[0]}-{k[1]:02d}",
                       "cnt": v["cnt"], "amt": round(v["amt"], 2),
                       "rate": round(rate, 1), "share": round(share, 1)})

    # 模块3：各渠道 退费金额/笔数 日趋势（按自然月分组，支持按月切换；仅成功；缺失日期 null 不连线）
    month_set = set()
    per_chan_month = {}
    for c in channels:
        name = c["displayName"]
        per_chan_month[name] = {}
        for r in c["rows"]:
            if r.get("status") in (None, "成功") and r.get("date") and r.get("y") and r.get("m"):
                mk = f"{r['y']}-{r['m']:02d}"
                month_set.add(mk)
                per_chan_month[name].setdefault(mk, {"amt": {}, "cnt": {}})
                d = per_chan_month[name][mk]
                d["amt"][r["date"]] = d["amt"].get(r["date"], 0) + r["price"]
                d["cnt"][r["date"]] = d["cnt"].get(r["date"], 0) + 1
    months = sorted(month_set)
    by_month = {}
    for mk in months:
        md = set()
        for name in per_chan_month:
            if mk in per_chan_month[name]:
                md |= set(per_chan_month[name][mk]["amt"].keys())
        dates = sorted(md)
        series = []
        for c in channels:
            name = c["displayName"]
            d = per_chan_month[name].get(mk, {"amt": {}, "cnt": {}})
            series.append({
                "name": name,
                "amt": [round(d["amt"][x], 2) if x in d["amt"] else None for x in dates],
                "cnt": [d["cnt"][x] if x in d["cnt"] else None for x in dates],
            })
        by_month[mk] = {"dates": dates, "series": series}
    m3 = {"months": months, "byMonth": by_month}
    return {"m1": m1, "m2": m2, "m3": m3}

def build_compare_html(compare_data, chart_js_content, generated_at):
    """复用 dashboard 的主题 CSS，生成独立的多渠道对比单文件页面"""
    css_match = re.search(r"<style>(.*?)</style>", DASHBOARD_HEAD, re.S)
    base_css = css_match.group(1) if css_match else ""
    extra = "\n  tbody tr.ch-sep td { border-top: 2px solid var(--primary); }\n"
    extra += "  tbody tr.total td { background:#eef2ff; font-weight:700; border-top:2px solid var(--primary); }\n"
    head = ('<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n<meta charset="UTF-8">\n'
            '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
            '<title>退费渠道数据对比</title>\n<style>' + base_css + extra + '</style>\n</head>\n')
    data_json = json.dumps(compare_data, ensure_ascii=False)
    meta_json = json.dumps({"generatedAt": generated_at})
    return (head + COMPARE_BODY +
            '\n<script>\n/* === Chart.js (内嵌) === */\n' + chart_js_content + '\n</script>\n'
            '<script>\nwindow.CHANNEL_META = ' + meta_json + ';\nwindow.COMPARE = ' + data_json + ';\n</script>\n'
            '<script>\n/* === 对比页逻辑 === */\n' + COMPARE_APP_JS + '\n</script>\n</body>\n</html>\n')


def main():
    if not os.path.exists(EXCEL):
        print("未找到 Excel：", EXCEL)
        return
    wb = openpyxl.load_workbook(EXCEL, data_only=True)
    channels = []
    # 渠道按钮文案覆盖（服务商列实际值 → 想显示的简称）；未列出的仍按服务商列众数
    DISPLAY_OVERRIDE = {"博迈通达充值": "博迈", "杭州自由行": "自由行"}
    for name in wb.sheetnames:
        ws = wb[name]
        matrix = [list(r) for r in ws.iter_rows(values_only=True)]
        rows = parse_sheet(matrix)
        if rows:
            # 渠道按钮文案：优先取「服务商」列的众数，回退到 sheet 名
            servs = [r.get("serv") for r in rows if r.get("serv")]
            raw = Counter(servs).most_common(1)[0][0] if servs else name
            display_name = DISPLAY_OVERRIDE.get(raw, raw)
            # 精简：看板渲染只用 date/y/m/day/price/status，移除其余字段以控制单文件体积
            for r in rows:
                for k in ("serv", "chan", "prod", "prov"):
                    r.pop(k, None)
            channels.append({"name": name, "displayName": display_name, "rows": rows})

    # 渠道展示顺序（用户指定）：瑞牛 > 文航 > 博迈 > 自由行
    # 匹配显示名或 sheet 名；未列出的渠道保持原有相对顺序并排在其后（sort 稳定）
    CHANNEL_ORDER = ["瑞牛", "文航", "博迈", "自由行"]
    def _chan_order_key(c):
        dn, nm = c.get("displayName") or "", c.get("name") or ""
        for i, key in enumerate(CHANNEL_ORDER):
            if key in dn or key in nm:
                return i
        return len(CHANNEL_ORDER)
    channels.sort(key=_chan_order_key)

    total = sum(len(c["rows"]) for c in channels)
    ok = sum(1 for c in channels for r in c["rows"] if r.get("status") in (None, "成功"))
    generated_at = time.strftime("%Y-%m-%d %H:%M:%S")
    print(f"解析: {len(channels)} 个渠道 / 共 {total} 条记录（成功 {ok} 条，其余为失败/其他，不计入退费统计）")
    for c in channels:
        cok = sum(1 for r in c["rows"] if r.get("status") in (None, "成功"))
        print(f"  - {c['name']}：{len(c['rows'])} 条（成功 {cok} 条）")

    # 1) 写 channels.js（外部数据文件，供高级用途，比如想拆开部署）
    with open(CHANNELS_JS, "w", encoding="utf-8") as f:
        f.write("// 由 build_data.py 自动生成。\n")
        f.write("window.CHANNEL_META = ")
        json.dump({"generatedAt": generated_at}, f, ensure_ascii=False)
        f.write(";\nwindow.CHANNELS = ")
        json.dump(channels, f, ensure_ascii=False)
        f.write(";\n")

    # 2) 读 Chart.js（外部依赖，本工程内嵌到 dashboard.html 用）
    chart_js_content = ""
    if os.path.exists(CHART_JS_PATH):
        with open(CHART_JS_PATH, "r", encoding="utf-8") as f:
            chart_js_content = f.read()
        print(f"已读取 Chart.js: {len(chart_js_content)/1024:.1f} KB")
    else:
        print(f"⚠ 未找到 {CHART_JS_PATH}，dashboard.html 将无法渲染图表")

    # 3) 生成单文件 dashboard.html（推荐使用此文件打开）
    html = build_dashboard(channels, chart_js_content, generated_at)
    with open(DASHBOARD_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"已生成 dashboard.html：{len(html)/1024:.1f} KB（单文件，自带 Chart.js 与数据）")

    # 4) 生成多渠道对比页 compare.html（独立页面，复用同一主题与数据）
    compare_data = build_compare_data(channels)
    chtml = build_compare_html(compare_data, chart_js_content, generated_at)
    compare_path = os.path.join(HERE, "compare.html")
    with open(compare_path, "w", encoding="utf-8") as f:
        f.write(chtml)
    print(f"已生成 compare.html：{len(chtml)/1024:.1f} KB（多渠道对比页）")

    print(f"  生成时间：{generated_at}")
    print("打开方式：双击 dashboard.html 或 compare.html 即可，无需任何额外依赖。")

if __name__ == "__main__":
    main()
