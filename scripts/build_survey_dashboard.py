#!/usr/bin/env python3
import argparse
import csv
import html
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path


RATING_ORDER = ["非常满意", "满意", "一般", "不满意", "非常不满意"]
RATING_SCORE = {"非常满意": 5, "满意": 4, "一般": 3, "不满意": 2, "非常不满意": 1}
POSITIVE_VALUES = {"非常满意", "满意", "5", "4", "很好", "好", "满意"}
NO_SUGGESTION_VALUES = {"", "无", "暂无", "没有", "还没想好", "无建议", "没意见", "nil", "none", "n/a", "na"}

GAIN_THEMES = [
    ("实际业务落地", ["业务", "工作", "岗位", "场景", "落地", "应用", "日常", "实际", "客户", "市场", "制造"]),
    ("工具与流程能力", ["工具", "系统", "平台", "软件", "模型", "ai", "流程", "搭建", "开发", "代码", "程序"]),
    ("自动化与效率提升", ["自动化", "自动", "提效", "效率", "节省", "一键", "批量", "减少", "替"]),
    ("知识技能获得", ["学会", "学习", "了解", "掌握", "入门", "知道", "理解", "熟悉", "方法"]),
    ("信心与认知变化", ["信心", "自信", "小白", "零基础", "成就", "认知", "敢", "焦虑", "心理"]),
]

SUGGESTION_THEMES = [
    ("课程节奏与消化时间", ["节奏", "速度", "时间", "紧张", "消化", "讲透", "吃力", "太快"]),
    ("案例深度与实操", ["案例", "场景", "示范", "实操", "实践", "练习", "演示", "针对性"]),
    ("工具选择与安全边界", ["工具", "选择", "安全", "风险", "泄密", "边界", "提示"]),
    ("内容结构与重点", ["内容", "太多", "深度", "系统", "重点", "乱", "清晰"]),
    ("讲解支持与技巧", ["老师", "讲解", "答疑", "技巧", "指导", "帮助"]),
]


def clean(value):
    return re.sub(r"\s+", " ", str(value or "").strip())


def pct(n, d):
    return 0 if d == 0 else round(n * 100 / d, 1)


def esc(value):
    return html.escape(str(value), quote=True)


def contains_any(text, needles):
    lower = text.lower()
    return any(needle.lower() in lower for needle in needles)


def read_csv_like(path, delimiter=None):
    encodings = ["utf-8-sig", "utf-8", "gb18030"]
    last_error = None
    for encoding in encodings:
        try:
            with path.open(newline="", encoding=encoding) as f:
                sample = f.read(4096)
                f.seek(0)
                dialect = csv.Sniffer().sniff(sample) if delimiter is None else csv.excel
                if delimiter is not None:
                    dialect.delimiter = delimiter
                return list(csv.DictReader(f, dialect=dialect))
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"Unable to parse {path}: {last_error}")


def read_xlsx(path):
    try:
        import openpyxl
    except ImportError as exc:
        raise RuntimeError("XLSX input requires openpyxl. Use the Codex workspace Python runtime or convert to CSV.") from exc

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [clean(cell) or f"列{i + 1}" for i, cell in enumerate(rows[0])]
    records = []
    for values in rows[1:]:
        if not any(clean(v) for v in values):
            continue
        records.append({headers[i]: clean(values[i]) if i < len(values) else "" for i in range(len(headers))})
    return records


def read_records(path):
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return read_csv_like(path)
    if suffix == ".tsv":
        return read_csv_like(path, delimiter="\t")
    if suffix == ".xlsx":
        return read_xlsx(path)
    raise RuntimeError(f"Unsupported file type: {suffix}. Use CSV, TSV, or XLSX.")


def infer_columns(rows):
    if not rows:
        return {"headers": [], "date": None, "ratings": [], "suggestions": [], "open_text": []}
    headers = list(rows[0].keys())
    result = {"headers": headers, "date": None, "ratings": [], "suggestions": [], "open_text": []}

    for header in headers:
        h = header.lower()
        values = [clean(row.get(header, "")) for row in rows]
        nonblank = [v for v in values if v]
        unique = set(nonblank)
        avg_len = sum(len(v) for v in nonblank) / len(nonblank) if nonblank else 0

        if result["date"] is None and any(k in h for k in ["时间", "日期", "date", "time"]):
            result["date"] = header

        rating_like_values = sum(1 for v in nonblank if v in RATING_SCORE or v in {"5", "4", "3", "2", "1"})
        if nonblank and (rating_like_values / len(nonblank) >= 0.7 or any(k in h for k in ["满意", "评分", "评价", "rating", "score"])):
            if len(unique) <= 8 and avg_len <= 12:
                result["ratings"].append(header)
                continue

        if any(k in h for k in ["建议", "意见", "改进", "不足", "comment", "suggest"]):
            result["suggestions"].append(header)
            result["open_text"].append(header)
        elif avg_len >= 12 or any(k in h for k in ["收获", "反馈", "感受", "备注", "说明", "gain", "feedback"]):
            if not any(k in h for k in ["姓名", "name", "提交人", "编号", "id"]):
                result["open_text"].append(header)

    return result


def count_rating(rows, column):
    counter = Counter(clean(row.get(column, "")) for row in rows)
    labels = [label for label in RATING_ORDER if counter.get(label)]
    labels.extend(sorted(k for k in counter if k and k not in labels))
    return [{"label": label, "count": counter[label], "pct": pct(counter[label], len(rows))} for label in labels]


def avg_rating(rows, column):
    values = []
    for row in rows:
        value = clean(row.get(column, ""))
        if value in RATING_SCORE:
            values.append(RATING_SCORE[value])
        elif value in {"5", "4", "3", "2", "1"}:
            values.append(int(value))
    return round(sum(values) / len(values), 2) if values else None


def count_themes(rows, columns, rules):
    combined = []
    for row in rows:
        text = " ".join(clean(row.get(column, "")) for column in columns)
        combined.append(text)
    items = []
    for name, needles in rules:
        count = sum(1 for text in combined if contains_any(text, needles))
        if count:
            items.append({"name": name, "count": count, "pct": pct(count, len(rows))})
    return sorted(items, key=lambda item: item["count"], reverse=True)


def suggestion_text(row, columns):
    return " ".join(clean(row.get(column, "")) for column in columns).strip()


def primary_rating_column(columns):
    for column in columns:
        if "整体" in column or "总体" in column:
            return column
    return columns[-1] if columns else None


def build_html(rows, columns, title, source_name):
    total = len(rows)
    date_col = columns["date"]
    rating_cols = columns["ratings"]
    suggestion_cols = columns["suggestions"]
    open_cols = columns["open_text"]
    primary_rating = primary_rating_column(rating_cols)

    date_items = []
    if date_col:
        date_counts = Counter(clean(row.get(date_col, "")) for row in rows if clean(row.get(date_col, "")))
        date_items = [{"name": date, "count": count, "pct": pct(count, total)} for date, count in sorted(date_counts.items())]

    ratings = []
    for column in rating_cols:
        counts = count_rating(rows, column)
        very = next((item for item in counts if item["label"] in {"非常满意", "5"}), {"count": 0, "pct": 0})
        positive = sum(item["count"] for item in counts if item["label"] in POSITIVE_VALUES)
        ratings.append({
            "name": column,
            "short": shorten(column, 16),
            "avg": avg_rating(rows, column),
            "counts": counts,
            "veryPct": very["pct"],
            "positivePct": pct(positive, total),
        })

    suggestion_count = 0
    if suggestion_cols:
        suggestion_count = sum(1 for row in rows if clean(suggestion_text(row, suggestion_cols)).lower() not in NO_SUGGESTION_VALUES)

    primary_positive = 0
    primary_very = 0
    if primary_rating:
        primary_positive = sum(1 for row in rows if clean(row.get(primary_rating, "")) in POSITIVE_VALUES)
        primary_very = sum(1 for row in rows if clean(row.get(primary_rating, "")) in {"非常满意", "5"})

    gain_cols = [c for c in open_cols if c not in suggestion_cols] or open_cols
    gain_themes = count_themes(rows, gain_cols, GAIN_THEMES) if gain_cols else []
    suggestion_themes = count_themes(rows, suggestion_cols, SUGGESTION_THEMES) if suggestion_cols else []

    detail_rows = []
    for index, row in enumerate(rows, 1):
        detail_rows.append({
            "id": clean(row.get("编号", "")) or str(index),
            "date": clean(row.get(date_col, "")) if date_col else "",
            "ratings": [clean(row.get(column, "")) for column in rating_cols],
            "feedback": "；".join(clean(row.get(column, "")) for column in gain_cols if clean(row.get(column, ""))) or "未填写",
            "suggestion": suggestion_text(row, suggestion_cols) or "未填写",
        })

    data = json.dumps({
        "dateItems": date_items,
        "ratings": ratings,
        "ratingCols": [shorten(c, 14) for c in rating_cols],
        "gainThemes": gain_themes,
        "suggestionThemes": suggestion_themes,
        "detailRows": detail_rows,
    }, ensure_ascii=False).replace("</", "<\\/")

    date_span = " / ".join([date_items[0]["name"], date_items[-1]["name"]]) if date_items else "未识别"
    positive_pct = pct(primary_positive, total) if primary_rating else 0
    very_pct = pct(primary_very, total) if primary_rating else 0

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{esc(title)}</title>
  <style>
    :root {{
      --bg:#f7f8f5; --surface:#fff; --surface-2:#eef3ee; --ink:#1e2528; --muted:#607071;
      --line:#dbe2dc; --primary:#0f766e; --primary-2:#115e59; --accent:#c2410c;
      --indigo:#4f46e5; --gold:#b7791f; --rose:#be123c; --shadow:0 16px 36px rgba(28,38,35,.08); --radius:8px;
    }}
    *{{box-sizing:border-box}} body{{margin:0;background:var(--bg);color:var(--ink);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;font-size:16px;line-height:1.55}}
    .page{{width:min(1180px,calc(100% - 32px));margin:0 auto;padding:28px 0 48px}}
    header{{display:grid;grid-template-columns:1fr auto;gap:24px;align-items:end;padding:22px 0 20px;border-bottom:1px solid var(--line)}}
    .eyebrow{{margin:0 0 8px;color:var(--primary-2);font-size:13px;font-weight:800}} h1{{margin:0;font-size:clamp(28px,4vw,44px);line-height:1.12;letter-spacing:0}}
    .header-copy{{max-width:760px;margin:12px 0 0;color:var(--muted)}} .meta{{min-width:220px;padding:14px 16px;border:1px solid var(--line);border-radius:var(--radius);background:var(--surface);box-shadow:var(--shadow)}}
    .meta span{{display:block;color:var(--muted);font-size:12px;font-weight:800}} .meta strong{{display:block;margin-top:2px;font-size:18px;font-variant-numeric:tabular-nums}}
    nav{{position:sticky;top:0;z-index:5;display:flex;gap:8px;align-items:center;margin:18px 0 20px;padding:8px;overflow-x:auto;border:1px solid var(--line);border-radius:var(--radius);background:rgba(247,248,245,.92);backdrop-filter:blur(10px)}}
    button{{font:inherit}} .tab,.filter{{min-height:44px;cursor:pointer;font-weight:800}} .tab{{flex:0 0 auto;padding:10px 14px;border:0;border-radius:6px;background:transparent;color:var(--muted)}} .tab[aria-selected=true]{{background:var(--ink);color:#fff}}
    .tab:focus-visible,.filter:focus-visible{{outline:3px solid rgba(15,118,110,.28);outline-offset:2px}} .section{{display:none}} .section.active{{display:block}}
    .grid{{display:grid;gap:16px}} .kpi-grid{{grid-template-columns:repeat(4,minmax(0,1fr));margin-bottom:18px}} .wide-grid{{grid-template-columns:minmax(0,1.2fr) minmax(300px,.8fr)}} .score-grid{{grid-template-columns:repeat(3,minmax(0,1fr))}}
    .card{{border:1px solid var(--line);border-radius:var(--radius);background:var(--surface);box-shadow:var(--shadow)}} .panel,.kpi,.score-card{{padding:18px}} .kpi{{min-height:138px}}
    .label{{margin:0 0 10px;color:var(--muted);font-size:13px;font-weight:800}} .value{{margin:0;font-size:34px;font-weight:850;line-height:1.08;font-variant-numeric:tabular-nums}} .note,.muted{{color:var(--muted)}} .note{{margin:10px 0 0;font-size:13px}}
    h2{{margin:0 0 14px;font-size:22px;line-height:1.25;letter-spacing:0}} h3{{margin:0 0 10px;font-size:17px}}
    .takeaways{{display:grid;gap:10px;margin:0;padding:0;list-style:none}} .takeaways li{{display:grid;grid-template-columns:28px 1fr;gap:10px;color:var(--muted)}} .takeaways b{{display:inline-grid;place-items:center;width:28px;height:28px;border-radius:999px;background:var(--surface-2);color:var(--primary-2);font-size:13px}}
    .chart-stack{{display:grid;gap:14px}} .bar-row{{display:grid;grid-template-columns:150px minmax(0,1fr) 76px;gap:12px;align-items:center;min-height:44px}} .bar-label{{font-weight:800}} .track{{height:14px;overflow:hidden;border-radius:999px;background:#e7ebe6}} .fill{{height:100%;min-width:2px;border-radius:999px;background:var(--primary)}} .fill.indigo{{background:var(--indigo)}} .fill.accent{{background:var(--accent)}} .fill.gold{{background:var(--gold)}} .fill.rose{{background:var(--rose)}} .bar-value{{color:var(--muted);text-align:right;font-variant-numeric:tabular-nums;font-weight:800}}
    .score-head{{display:flex;justify-content:space-between;gap:12px;align-items:start;margin-bottom:14px}} .score-number{{color:var(--primary-2);font-size:30px;font-weight:850;line-height:1;font-variant-numeric:tabular-nums}} .stacked{{display:flex;height:18px;overflow:hidden;border-radius:999px;background:#e8ece7}} .segment-very{{background:var(--primary)}} .segment-ok{{background:var(--gold)}} .legend{{display:flex;flex-wrap:wrap;gap:10px 16px;margin-top:12px;color:var(--muted);font-size:13px}} .legend span{{display:inline-flex;gap:6px;align-items:center}} .swatch{{width:12px;height:12px;border-radius:3px;background:var(--primary)}} .swatch.ok{{background:var(--gold)}}
    .filters{{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 14px}} .filter{{padding:9px 13px;border:1px solid var(--line);border-radius:999px;background:var(--surface);color:var(--muted)}} .filter.active{{border-color:var(--primary);background:#e1f1ee;color:var(--primary-2)}}
    .table-wrap{{overflow-x:auto;border:1px solid var(--line);border-radius:var(--radius);background:var(--surface)}} table{{width:100%;border-collapse:collapse;min-width:880px}} th,td{{padding:12px 14px;border-bottom:1px solid var(--line);vertical-align:top;text-align:left}} th{{background:#eef3ee;color:#304042;font-size:13px;font-weight:850}} td{{color:var(--muted);font-size:14px}} .pill{{display:inline-flex;min-height:28px;align-items:center;padding:3px 9px;border-radius:999px;background:#e1f1ee;color:var(--primary-2);font-size:12px;font-weight:850;white-space:nowrap}} .pill.ok{{background:#fff1d7;color:#8a5a12}}
    @media(max-width:920px){{header,.wide-grid,.score-grid,.kpi-grid{{grid-template-columns:1fr}}.meta{{min-width:0}}.bar-row{{grid-template-columns:108px minmax(0,1fr) 64px}}}} @media(max-width:560px){{.page{{width:min(100% - 20px,1180px);padding-top:12px}}.panel,.kpi,.score-card{{padding:14px}}.bar-row{{grid-template-columns:1fr;gap:6px}}.bar-value{{text-align:left}}}}
  </style>
</head>
<body>
<main class="page">
  <header>
    <div><p class="eyebrow">问卷反馈分析 / {esc(date_span)}</p><h1>{esc(title)}</h1><p class="header-copy">基于 {total} 份有效答卷，自动汇总评分分布、开放反馈主题、改进建议与明细记录，用于复盘和外发汇报。</p></div>
    <aside class="meta"><span>数据来源</span><strong>{esc(source_name)}</strong><span style="margin-top:10px">统计口径</span><strong>有效答卷 {total} 份</strong></aside>
  </header>
  <nav role="tablist">
    <button class="tab" role="tab" aria-selected="true" data-target="overview">总览</button>
    <button class="tab" role="tab" aria-selected="false" data-target="ratings">评分分布</button>
    <button class="tab" role="tab" aria-selected="false" data-target="themes">反馈主题</button>
    <button class="tab" role="tab" aria-selected="false" data-target="details">原始反馈</button>
  </nav>
  <section id="overview" class="section active" role="tabpanel">
    <div class="grid kpi-grid">
      <article class="card kpi"><p class="label">有效答卷</p><p class="value">{total}</p><p class="note">来自源表解析后的记录数</p></article>
      <article class="card kpi"><p class="label">主要评分正向率</p><p class="value">{positive_pct}%</p><p class="note">{esc(shorten(primary_rating or "未识别评分列", 18))}</p></article>
      <article class="card kpi"><p class="label">非常满意 / 最高分</p><p class="value">{very_pct}%</p><p class="note">基于主要评分列</p></article>
      <article class="card kpi"><p class="label">可行动建议</p><p class="value">{suggestion_count}</p><p class="note">已排除空白和“无”类回答</p></article>
    </div>
    <div class="grid wide-grid">
      <article class="card panel"><h2>关键结论</h2><ul class="takeaways"><li><b>1</b><span>整体反馈以正向评价为主，可优先关注高频收获主题与可行动建议。</span></li><li><b>2</b><span>开放反馈中的主题可作为课程、产品或服务下一步优化的优先级线索。</span></li><li><b>3</b><span>原始反馈表保留了逐条记录，方便回看证据和进一步人工归纳。</span></li></ul></article>
      <article class="card panel"><h2>提交时间分布</h2><div id="dateChart" class="chart-stack"></div></article>
    </div>
  </section>
  <section id="ratings" class="section" role="tabpanel"><div class="grid score-grid" id="scoreCards"></div><article class="card panel" style="margin-top:16px"><h2>评分对比</h2><div id="ratingCompare" class="chart-stack"></div></article></section>
  <section id="themes" class="section" role="tabpanel"><div class="grid wide-grid"><article class="card panel"><h2>开放反馈主题</h2><div id="gainThemes" class="chart-stack"></div></article><article class="card panel"><h2>建议主题</h2><div id="suggestionThemes" class="chart-stack"></div></article></div></section>
  <section id="details" class="section" role="tabpanel"><article class="card panel"><h2>原始反馈浏览</h2><div class="filters"><button class="filter active" data-filter="all">全部</button><button class="filter" data-filter="actionable">仅看有建议</button><button class="filter" data-filter="very">主要评分最高</button></div><div class="table-wrap"><table><thead><tr><th>编号</th><th>日期</th><th>评分</th><th>开放反馈</th><th>意见与建议</th></tr></thead><tbody id="detailBody"></tbody></table></div></article></section>
</main>
<script>
const DATA = {data};
const noSuggestionValues = new Set(["", "无", "暂无", "没有", "还没想好", "未填写", "无建议", "没意见"]);
const colorClass = ["", "indigo", "accent", "gold", "rose"];
function escapeHtml(value) {{ return String(value ?? "").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;"); }}
function renderBars(targetId, items) {{
  const target = document.getElementById(targetId);
  if (!items.length) {{ target.innerHTML = '<p class="muted">未识别到可绘制数据。</p>'; return; }}
  const max = Math.max(...items.map(item => item.count), 1);
  target.innerHTML = items.map((item, index) => `<div class="bar-row"><div class="bar-label">${{escapeHtml(item.name || item.short || item.label)}}</div><div class="track" aria-hidden="true"><div class="fill ${{colorClass[index % colorClass.length]}}" style="width:${{Math.max(2, item.count / max * 100)}}%"></div></div><div class="bar-value">${{item.count}} · ${{item.pct}}%</div></div>`).join("");
}}
function renderScoreCards() {{
  const target = document.getElementById("scoreCards");
  if (!DATA.ratings.length) {{ target.innerHTML = '<article class="card panel"><h2>未识别评分列</h2><p class="muted">请检查表头是否包含满意度、评价或评分字段。</p></article>'; return; }}
  target.innerHTML = DATA.ratings.map(info => {{
    const very = info.counts.find(item => ["非常满意", "5"].includes(item.label)) || {{count:0,pct:0}};
    const ok = info.counts.find(item => item.label === "满意" || item.label === "4") || {{count:0,pct:0}};
    return `<article class="card score-card"><div class="score-head"><div><h2>${{escapeHtml(info.short)}}</h2><p class="muted" style="margin:0">${{escapeHtml(info.name)}}</p></div><div class="score-number">${{info.avg ?? "-"}}</div></div><div class="stacked"><div class="segment-very" style="width:${{very.pct}}%"></div><div class="segment-ok" style="width:${{ok.pct}}%"></div></div><div class="legend"><span><i class="swatch"></i>最高分 ${{very.count}} · ${{very.pct}}%</span><span><i class="swatch ok"></i>次高分 ${{ok.count}} · ${{ok.pct}}%</span></div></article>`;
  }}).join("");
}}
function renderRatingCompare() {{ renderBars("ratingCompare", DATA.ratings.map(item => ({{name:item.short,count:Math.round(item.veryPct),pct:item.veryPct}}))); }}
function renderDetails(filter = "all") {{
  const rows = DATA.detailRows.filter(row => {{
    if (filter === "actionable") return !noSuggestionValues.has(String(row.suggestion).trim().toLowerCase());
    if (filter === "very") return row.ratings.some(value => ["非常满意", "5"].includes(value));
    return true;
  }});
  document.getElementById("detailBody").innerHTML = rows.map(row => `<tr><td>${{escapeHtml(row.id)}}</td><td>${{escapeHtml(row.date)}}</td><td>${{row.ratings.map(value => `<span class="pill ${{["非常满意","5"].includes(value) ? "" : "ok"}}">${{escapeHtml(value || "-")}}</span>`).join(" ")}}</td><td>${{escapeHtml(row.feedback)}}</td><td>${{escapeHtml(row.suggestion)}}</td></tr>`).join("");
}}
document.querySelectorAll(".tab").forEach(tab => tab.addEventListener("click", () => {{ document.querySelectorAll(".tab").forEach(item => item.setAttribute("aria-selected","false")); document.querySelectorAll(".section").forEach(section => section.classList.remove("active")); tab.setAttribute("aria-selected","true"); document.getElementById(tab.dataset.target).classList.add("active"); }}));
document.querySelectorAll(".filter").forEach(button => button.addEventListener("click", () => {{ document.querySelectorAll(".filter").forEach(item => item.classList.remove("active")); button.classList.add("active"); renderDetails(button.dataset.filter); }}));
renderBars("dateChart", DATA.dateItems); renderScoreCards(); renderRatingCompare(); renderBars("gainThemes", DATA.gainThemes); renderBars("suggestionThemes", DATA.suggestionThemes); renderDetails();
</script>
</body>
</html>"""


def shorten(text, max_len):
    text = clean(text)
    return text if len(text) <= max_len else text[: max_len - 1] + "…"


def main():
    parser = argparse.ArgumentParser(description="Build a self-contained survey HTML dashboard from CSV/TSV/XLSX.")
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--title", default="问卷数据看板")
    parser.add_argument("--zip", action="store_true", help="Also create a zip file next to the HTML.")
    args = parser.parse_args()

    rows = read_records(args.input)
    if not rows:
        raise SystemExit("No records found in input file.")
    columns = infer_columns(rows)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(build_html(rows, columns, args.title, args.input.name), encoding="utf-8")

    print(args.output)
    if args.zip:
        zip_path = args.output.with_suffix(".zip")
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.write(args.output, arcname=args.output.name)
        print(zip_path)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise
