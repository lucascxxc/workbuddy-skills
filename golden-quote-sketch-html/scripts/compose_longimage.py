# -*- coding: utf-8 -*-
"""Compose the 16 cards + cover into a single vertical long-image PNG (grayscale)."""
import json, os
from PIL import Image, ImageDraw, ImageFont

BASE = r"C:\Users\EngineeringCenter\WorkBuddy\2026-08-28-14-51-48"
WEB = os.path.join(BASE, "art", "web")
FONT_DIR = r"C:\Windows\Fonts"
OUT = r"C:\Users\EngineeringCenter\Desktop\金句手绘解读_WAIC2026_长图.png"

# ---- fonts ----
def load_font(name, size):
    candidates = [os.path.join(FONT_DIR, name)]
    for p in candidates:
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            pass
    return ImageFont.load_default()

SERIF = os.path.join(FONT_DIR, "STSONG.TTF")      # 华文宋体 (serif)
SANS  = os.path.join(FONT_DIR, "simhei.ttf")      # 黑体 (sans, labels)
try:
    f_title  = ImageFont.truetype(SERIF, 96)
    f_num    = ImageFont.truetype(SANS, 64)
    f_theme  = ImageFont.truetype(SANS, 30)
    f_quote  = ImageFont.truetype(SERIF, 46)
    f_speak  = ImageFont.truetype(SANS, 26)
    f_read   = ImageFont.truetype(SERIF, 31)
    f_cover_s= ImageFont.truetype(SERIF, 40)
    f_cover_x= ImageFont.truetype(SANS, 26)
    f_foot   = ImageFont.truetype(SANS, 22)
except Exception as e:
    print("font load error", e)
    raise

WHITE = 255
BLACK = 0
GRAY_L = 200   # light divider
GRAY_M = 150   # secondary text
GRAY_D = 70    # reading text

W = 1080
MX = 130                      # side margin
CW = W - 2 * MX               # content width 820
IMG_W = 720                   # sketch display width
IMG_X = (W - IMG_W) // 2

def wrap(draw, text, font, max_w):
    """Wrap CJK + latin text by character when needed."""
    lines = []
    for para in text.split("\n"):
        if para == "":
            lines.append("")
            continue
        cur = ""
        for ch in para:
            test = cur + ch
            if draw.textlength(test, font=font) <= max_w:
                cur = test
            else:
                if cur:
                    lines.append(cur)
                cur = ch
        if cur:
            lines.append(cur)
    return lines

def new_section(h):
    return Image.new("L", (W, h), WHITE), ImageDraw.Draw(None)

# ---- build section images ----
sections = []

def render_cover():
    h = 1500
    img = Image.new("L", (W, h), WHITE)
    d = ImageDraw.Draw(img)
    # title
    title = "金句 · 手绘 · 解读"
    d.text((W/2, 360), title, font=f_title, fill=BLACK, anchor="mm")
    # subtitle
    sub = "WAIC 2026 智能财务论坛"
    d.text((W/2, 520), sub, font=f_cover_s, fill=BLACK, anchor="mm")
    sub2 = "会议精神与行业观察 · 16 句代表性金句"
    d.text((W/2, 590), sub2, font=f_cover_x, fill=GRAY_M, anchor="mm")
    # divider
    d.line([(W/2-60, 700), (W/2+60, 700)], fill=GRAY_L, width=2)
    # source
    src = "素材来源：《2026 WAIC 智能财务论坛\n会议精神与行业观察总结报告》"
    y = 770
    for ln in src.split("\n"):
        d.text((W/2, y), ln, font=f_cover_x, fill=GRAY_M, anchor="mm")
        y += 40
    note = "铅笔素描 · 黑白灰 · 单一诗意意象 · 留白极简"
    d.text((W/2, 900), note, font=f_cover_x, fill=GRAY_M, anchor="mm")
    return img

sections.append(render_cover())

with open(os.path.join(BASE, "quotes.json"), encoding="utf-8") as f:
    quotes = json.load(f)

for q in quotes:
    # measure heights first
    tmp = Image.new("L", (10, 10), WHITE)
    td = ImageDraw.Draw(tmp)
    quote_lines = wrap(td, q["quote"], f_quote, CW)
    read_lines = wrap(td, q["reading"], f_read, CW)

    img_h = 0
    img_h += 70   # top gap
    img_h += 64   # number
    img_h += 14
    img_h += 40   # theme
    img_h += 56   # gap before image
    img_h += int(IMG_W * 1536 / 1024)   # sketch height (1024x1536)
    img_h += 52   # gap after image
    img_h += len(quote_lines) * 64      # quote lines (46px, lh 1.4)
    img_h += 22   # gap
    img_h += 36   # speaker
    img_h += 40   # gap
    img_h += len(read_lines) * 53       # reading lines (31px, lh ~1.7)
    img_h += 100  # bottom gap

    img = Image.new("L", (W, img_h), WHITE)
    d = ImageDraw.Draw(img)
    y = 70
    # number (light, large) + theme tag
    d.text((MX, y), q["id"], font=f_num, fill=GRAY_L, anchor="lm")
    # theme tag with small left bar
    theme = q["theme"]
    bar_x = MX + 130
    d.line([(bar_x, y+8), (bar_x, y+34)], fill=GRAY_M, width=3)
    d.text((bar_x+16, y+10), theme, font=f_theme, fill=GRAY_M, anchor="lm")
    y += 64 + 14 + 40 + 56
    # sketch
    sk = Image.open(os.path.join(WEB, f"q{q['id']}.jpg")).convert("L")
    sk = sk.resize((IMG_W, int(IMG_W*1536/1024)), Image.LANCZOS)
    img.paste(sk, (IMG_X, y))
    y += sk.height + 52
    # quote (centered)
    for ln in quote_lines:
        d.text((W/2, y), ln, font=f_quote, fill=BLACK, anchor="mm")
        y += 64
    y += 22
    # speaker (centered, gray)
    d.text((W/2, y), "—— " + q["speaker"], font=f_speak, fill=GRAY_M, anchor="mm")
    y += 36 + 40
    # reading (left aligned)
    for ln in read_lines:
        d.text((MX, y), ln, font=f_read, fill=GRAY_D, anchor="lm")
        y += 53
    # divider at bottom gap
    d.line([(W/2-40, img_h-48), (W/2+40, img_h-48)], fill=GRAY_L, width=1)
    sections.append(img)

# footer
fh = 120
foot = Image.new("L", (W, fh), WHITE)
fd = ImageDraw.Draw(foot)
fd.text((W/2, 50), "金句 · 手绘 · 解读  |  WAIC 2026 智能财务论坛", font=f_foot, fill=GRAY_M, anchor="mm")
fd.text((W/2, 84), "单文件长图 · 离线可阅", font=f_foot, fill=GRAY_L, anchor="mm")
sections.append(foot)

# ---- stack ----
total_h = sum(s.height for s in sections)
canvas = Image.new("L", (W, total_h), WHITE)
y = 0
for s in sections:
    canvas.paste(s, (0, y))
    y += s.height

canvas.save(OUT, "JPEG", quality=88, optimize=True, progressive=True)
print("saved:", OUT)
print("size:", canvas.size, "bytes:", os.path.getsize(OUT))
