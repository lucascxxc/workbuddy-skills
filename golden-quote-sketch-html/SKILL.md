---
name: golden-quote-sketch-html
description: 当用户要求把长文、访谈、报告或播客文字稿做成「金句·手绘·解读」风格的单文件 HTML 长图，或进一步导出为单张竖向长图图片时使用。技能覆盖：提炼金句、撰写语境解读、生成铅笔素描配图、压缩图片、组装成 base64 内联的离线 HTML，以及可选导出 PNG/JPEG 长图。
agent_created: true
---

# 金句 · 手绘 · 解读 HTML 长图

## 概述

把一篇长文本（访谈、播客、会议报告、演讲稿等）转换成「金句 + 手绘 + 解读」的极简长图。可输出：
1. 单文件 HTML（零外部资源、base64 内联图片、可离线打开）；
2. 单张竖向长图图片 PNG/JPEG（1080 px 宽，适合分享）。

## 触发场景

用户请求包含以下任一关键词或意图时触发：

- 「金句 · 手绘 · 解读」
- 把访谈/播客/报告做成金句长图
- 金句海报、金句插画、金句可视化
- 单文件 HTML 长图、可离线打开、base64 内联图片

## 工作流程

### Phase 1：通读与提炼

1. 读取用户提供的长文本（docx、pdf、txt、markdown 等）。
2. 提炼 15–18 句最有代表性的金句。
3. 为每句标注：
   - 编号（01、02…）
   - 主题标签（如「关于初衷」「关于克制」）
   - 原话（尽量使用原文）
   - 说话人/来源
   - 回到语境的解读（背景 + 为什么重要）
   - 单一诗意视觉隐喻（用于配图，例如「一扇将启的门」「沙漠中的海市蜃楼」）
4. 将结果写入 `quotes.json`。

### Phase 2：生成手绘配图

1. 使用 ImageGen 为每句金句生成一幅配图。
2. 强制参数：
   - `size`: `"1024x1536"`（竖版）
   - `quality`: `"high"`
3. 固定提示词模板：
   ```
   Hand-drawn pencil sketch illustration on pure white paper background, graphite drawing, minimalist with generous whitespace, single poetic central image, black white and gray tones only, absolutely no text, no letters, no words, no numbers, no symbols, no color, no watermark. [视觉隐喻描述]. Soft graphite shading, loose sketchy pencil strokes, delicate elegant fine-art style. Vertical portrait composition.
   ```
4. ⚠️ **不要并行调用 ImageGen**：并行会导致时间戳文件名冲突、互相覆盖。必须串行逐张生成，生成一张后立即重命名为 `q01.png`、`q02.png` 等固定名称，再继续下一张。

### Phase 3：图片后处理

1. 用 Pillow 打开每张 PNG，转换为灰度（`convert('L')`）。
2. 以 JPEG 格式保存，质量建议 80–85，并启用 `optimize=True`。
3. 输出到 `art/web/q01.jpg` 等路径，目标是把 16 张图控制在 2 MB 以内。

### Phase 4：组装 HTML

1. 读取 `quotes.json` 和 `art/web/*.jpg`。
2. 将 JPEG 转为 base64 数据 URI（`data:image/jpeg;base64,...`）。
3. 生成单文件 HTML，结构：
   - 封面区：大标题「金句 · 手绘 · 解读」、副标题、来源信息
   - 16 张卡片：配图 + 编号/主题 + 大字金句 + 说话人 + 解读
4. 样式要求：
   - 浅色主题、大量留白
   - 衬线大字（`font-family: "Noto Serif SC", "Source Han Serif SC", "Songti SC", serif`）
   - 卡片式分区、阴影轻、圆角小
   - 响应式、移动端可用
5. 写入最终 HTML 文件。

### Phase 5（可选）：导出单张长图 PNG/JPEG

1. 如果用户要求「长图形式」「长图图片」「发朋友圈/小红书」，用 Pillow 将已有 `quotes.json` + `art/web/*.jpg` 合成为单张竖向长图。
2. 推荐规格：
   - 宽度 `1080 px`（移动端分享标准宽度）
   - 灰度模式（`L`），背景纯白
   - 封面 + 16 区 + 页脚，上下连续，各区以细线和留白分隔
   - 金句用衬线大字居中，解读用衬线小字左对齐
   - 说话人、主题标签用无衬线小字
3. 输出为 JPEG（质量 85–90， progressive）以控制体积，约 2–4 MB。
4. 示例实现：`scripts/compose_longimage.py`。
5. **发布版注意事项**：若用户要把长图发布到网络平台并弱化 AI/制作过程痕迹，应删除或替换以下内容：
   - 封面标题中的「金句 · 手绘 · 解读」等过程性标题
   - 「铅笔素描」「黑白灰」「单一诗意意象」「留白极简」等风格说明
   - 「素材来源：……」等来源说明
   - 页脚中的「单文件长图」「离线可阅」等制作说明
   - 可替换为中性封面标题（如「WAIC 2026 智能财务论坛」「行业观察与观点摘录」）和简洁结尾（如「— 完 —」）

### Phase 6（可选）：独立文章封面

1. 若用户需要一张单独的「文章封面」（用于公众号/小红书头图或长图首屏），用 Pillow 合成一张竖版封面。
2. 推荐规格：1080 × 约 2050 px，灰度 L 模式，白色背景。
3. 结构：顶部衬线大标题（如「WAIC 2026」「智能财务论坛」）+ 无衬线副标题 + 细线分隔；中部一张主视觉素描（复用 `art/web/q13.jpg` 等现有素材，避免重新生成）；底部中性落款（如「2026 · 上海」）。
4. 不出现任何 AI/制作过程文字。
5. 示例实现：`scripts/make_cover.py`。

## 脚本

- `scripts/build_html.py`：读取 `quotes.json` 与 `art/web/*.jpg`，输出单文件 HTML。可直接复用。
- `scripts/compose_longimage.py`：读取 `quotes.json` 与 `art/web/*.jpg`，输出单张竖向长图 PNG/JPEG。可直接复用。
- `scripts/make_cover.py`：读取 `art/web/q13.jpg`，输出独立文章封面 PNG/JPEG。可直接复用。

## 交付与验证

- HTML 产物：单个 `.html` 文件，无任何外部 URL，可离线打开，用 `present_files` 呈现。
- 长图产物：单个 `.png`/`.jpg` 文件，宽高适中，适合社交媒体分享，同样用 `present_files` 呈现。
