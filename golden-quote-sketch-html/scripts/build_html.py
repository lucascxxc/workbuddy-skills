#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json, base64, os

WORKSPACE = r"C:/Users/EngineeringCenter/WorkBuddy/2026-08-28-14-51-48"
QUOTES_PATH = os.path.join(WORKSPACE, "quotes.json")
IMG_DIR = os.path.join(WORKSPACE, "art", "web")
OUT_PATH = os.path.join(WORKSPACE, "金句手绘解读_WAIC2026.html")

with open(QUOTES_PATH, "r", encoding="utf-8") as f:
    quotes = json.load(f)

def b64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")

html_parts = []
html_parts.append("""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>金句 · 手绘 · 解读｜WAIC 2026 智能财务论坛</title>
<style>
  :root {
    --bg: #faf9f6;
    --text: #1f1f1f;
    --muted: #5c5c5c;
    --line: #d8d4cc;
    --accent: #2b3a4a;
    --card: #ffffff;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    background: var(--bg);
    color: var(--text);
    font-family: "Noto Serif SC", "Source Han Serif SC", "Songti SC", "SimSun", "STSong", serif;
    line-height: 1.85;
    -webkit-font-smoothing: antialiased;
  }
  .wrap {
    max-width: 760px;
    margin: 0 auto;
    padding: 0 24px 160px;
  }
  header {
    padding: 140px 0 120px;
    text-align: center;
    border-bottom: 1px solid var(--line);
  }
  h1 {
    font-size: clamp(42px, 9vw, 72px);
    font-weight: 700;
    letter-spacing: 0.08em;
    line-height: 1.15;
    color: var(--text);
  }
  .subtitle {
    margin-top: 32px;
    font-size: 20px;
    color: var(--muted);
    letter-spacing: 0.05em;
  }
  .meta {
    margin-top: 60px;
    font-size: 14px;
    color: var(--muted);
    line-height: 2;
  }
  .meta span { display: inline-block; margin: 0 10px; }
  .card {
    background: var(--card);
    margin-top: 120px;
    padding: 72px 48px 80px;
    border-radius: 4px;
    box-shadow: 0 12px 40px rgba(0,0,0,0.04);
  }
  .card:first-of-type { margin-top: 100px; }
  .visual {
    text-align: center;
    margin-bottom: 48px;
  }
  .visual img {
    max-width: 100%;
    width: 560px;
    height: auto;
    display: inline-block;
    border-radius: 2px;
  }
  .label {
    display: flex;
    align-items: baseline;
    gap: 14px;
    margin-bottom: 28px;
    flex-wrap: wrap;
  }
  .num {
    font-family: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace;
    font-size: 14px;
    color: var(--accent);
    letter-spacing: 0.1em;
  }
  .theme {
    font-size: 15px;
    color: var(--accent);
    border: 1px solid var(--line);
    padding: 4px 12px;
    border-radius: 999px;
    letter-spacing: 0.05em;
  }
  .quote {
    font-size: clamp(26px, 5vw, 36px);
    line-height: 1.55;
    font-weight: 600;
    color: var(--text);
    margin-bottom: 24px;
  }
  .speaker {
    font-size: 15px;
    color: var(--muted);
    margin-bottom: 36px;
    font-style: italic;
  }
  .reading {
    font-size: 17px;
    color: var(--muted);
    line-height: 2;
    text-align: justify;
  }
  footer {
    margin-top: 140px;
    text-align: center;
    font-size: 13px;
    color: var(--muted);
    letter-spacing: 0.05em;
  }
  @media (max-width: 600px) {
    .wrap { padding: 0 18px 100px; }
    header { padding: 90px 0 70px; }
    .card { padding: 44px 24px 56px; margin-top: 80px; }
    .quote { font-size: 24px; }
  }
</style>
</head>
<body>
<div class="wrap">
<header>
  <h1>金句 · 手绘 · 解读</h1>
  <p class="subtitle">WAIC 2026 智能财务论坛 十六句</p>
  <p class="meta">
    来源：2026 WAIC 智能财务论坛会议精神与行业观察总结报告<br>
    <span>上海国家会计学院 × 汇付天下 联合主办</span>
  </p>
</header>
""")

for q in quotes:
    img_path = os.path.join(IMG_DIR, f"q{q['id']}.jpg")
    b64img = b64(img_path)
    html_parts.append(f"""
<article class="card" id="q{q['id']}">
  <div class="visual">
    <img src="data:image/jpeg;base64,{b64img}" alt="配图">
  </div>
  <div class="label">
    <span class="num">{q['id']}</span>
    <span class="theme">{q['theme']}</span>
  </div>
  <p class="quote">「{q['quote']}」</p>
  <p class="speaker">{q['speaker']}</p>
  <p class="reading">{q['reading']}</p>
</article>
""")

html_parts.append("""
<footer>
  单文件 HTML · 图片 base64 内联 · 无需联网
</footer>
</div>
</body>
</html>
""")

with open(OUT_PATH, "w", encoding="utf-8") as f:
    f.write("".join(html_parts))

print(f"HTML written: {OUT_PATH}")
print(f"Size: {os.path.getsize(OUT_PATH) / 1024 / 1024:.2f} MB")
