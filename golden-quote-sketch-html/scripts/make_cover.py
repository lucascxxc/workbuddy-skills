# -*- coding: utf-8 -*-
"""Generate a 1:1 square article cover (grayscale, minimalist, no AI/process text)."""
import os
from PIL import Image, ImageDraw, ImageFont

BASE = r"C:\Users\EngineeringCenter\WorkBuddy\2026-08-28-14-51-48"
WEB = os.path.join(BASE, "art", "web")
FONT_DIR = r"C:\Windows\Fonts"
OUT = r"C:\Users\EngineeringCenter\Desktop\WAIC2026_智能财务论坛_封面_1x1.png"

SERIF = os.path.join(FONT_DIR, "STSONG.TTF")   # 华文宋体 serif
SANS  = os.path.join(FONT_DIR, "simhei.ttf")   # 黑体 sans

WHITE = 255
BLACK = 0
GRAY_M = 150
GRAY_L = 200

W = 1080
H = 1080

f_big1 = ImageFont.truetype(SERIF, 80)   # WAIC 2026
f_big2 = ImageFont.truetype(SERIF, 56)   # 智能财务论坛
f_sub  = ImageFont.truetype(SANS, 30)    # 行业观察与观点摘录
f_cap  = ImageFont.truetype(SANS, 24)    # 2026 · 上海

# hero sketch -> square center crop -> resize to fit
sk_src = Image.open(os.path.join(WEB, "q13.jpg")).convert("L")
sw, sh = sk_src.size                       # 1024 x 1536
sq = min(sw, sh)
left = (sw - sq) // 2
top = (sh - sq) // 2
sk_sq = sk_src.crop((left, top, left + sq, top + sq)).resize((560, 560), Image.LANCZOS)
SK = 560
SK_H = 560

# layout (square)
title1_y = 100
title2_y = 178
sub_y    = 244
div_y    = 292
img_y    = 330
cap_y    = 1062

img = Image.new("L", (W, H), WHITE)
d = ImageDraw.Draw(img)

d.text((W/2, title1_y), "WAIC 2026", font=f_big1, fill=BLACK, anchor="mm")
d.text((W/2, title2_y), "智能财务论坛", font=f_big2, fill=BLACK, anchor="mm")
d.text((W/2, sub_y), "行业观察与观点摘录", font=f_sub, fill=GRAY_M, anchor="mm")
d.line([(W/2-70, div_y), (W/2+70, div_y)], fill=GRAY_L, width=2)
img.paste(sk_sq, ((W-SK)//2, img_y))
d.text((W/2, cap_y), "2026 · 上海", font=f_cap, fill=GRAY_M, anchor="mm")

img.save(OUT, "JPEG", quality=90, optimize=True, progressive=True)
print("saved:", OUT, "size:", img.size, "bytes:", os.path.getsize(OUT))
