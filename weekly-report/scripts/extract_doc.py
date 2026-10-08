# -*- coding: utf-8 -*-
"""从 Word 97-2003 (.doc) 老格式文件中提取文本、方向表和图片。

用法:
    python extract_doc.py <file.doc>                     # 打印文本到 stdout
    python extract_doc.py <file.doc> --out <dir>         # 文本打印 + 图片导出到 dir
    python extract_doc.py <file.doc> --out <dir> --text <txtfile>   # 文本写文件 + 图片导出
    python extract_doc.py <file.doc> --direction <json>  # 解析"工作内容、进度情况及推进方向"表格 → JSON 数组

依赖: olefile (pip install olefile)
"""
import sys
import re
import struct
import os
import json
import olefile

def extract_text_from_doc(path, keep_cell_sep=False):
    """提取 .doc 的纯文本（WordDocument 流的 UTF-16LE 文本）。

    keep_cell_sep=True 时保留 \\x07（表格单元格分隔符），供方向表解析使用。
    """
    ole = olefile.OleFileIO(path)
    try:
        if not ole.exists("WordDocument"):
            return ""
        data = ole.openstream("WordDocument").read()
        # 老 doc 文本通常以 UTF-16LE 存储，尝试两种解码
        # 方式1：整流 UTF-16LE 解码，过滤控制字符
        try:
            text = data.decode("utf-16-le", errors="ignore")
            if keep_cell_sep:
                # 保留 \x07（单元格分隔），只过滤其他控制字符
                text = re.sub(r"[\x00-\x06\x08\x0b\x0c\x0e-\x1f]", "", text)
            else:
                text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
            text = re.sub(r"[ \t]+\n", "\n", text)
            text = re.sub(r"\n{3,}", "\n\n", text)
            return text.strip()
        except Exception:
            return ""
    finally:
        ole.close()

def _is_readable_work(s):
    """工作内容标题应为正常中文/数字/标点，过滤二进制流乱码行。
    CJK 字符占比过低（<50%）视为乱码。"""
    if not s:
        return False
    cjk = sum(1 for ch in s if '\u4e00' <= ch <= '\u9fff')
    return cjk / max(1, len(s)) >= 0.5

def parse_direction_table(path):
    """解析科研组周报 .doc 中的"工作内容、进度情况及推进方向"表格。

    表格结构（老 .doc 单元格分隔符 \\x07，行间 \\x07\\x07）：
        工作内容 | 进度情况 | 推进方向/亟待解决
        标题1 | 进度1 | 推进方向：方向1
        ...
    返回: [{'work': ..., 'progress': ..., 'direction': ...}, ...]
    """
    text = extract_text_from_doc(path, keep_cell_sep=True)
    if not text:
        return []
    # 定位表格区域（跳过文档标题行）
    j = text.find("工作内容、进度情况及推进方向")
    if j < 0:
        j = text.find("工作内容")
    if j < 0:
        return []
    seg = text[j:]
    rows = []
    for block in seg.split("\x07\x07"):
        cells = [c.strip() for c in block.split("\x07") if c.strip()]
        if not cells:
            continue
        # 跳过表头行（"工作内容..." 或 "进度情况"）
        if cells[0].startswith("工作内容") or "进度情况" in cells[0]:
            continue
        work = " ".join(cells[0].split())  # 合并单元格内换行（如"慧眼行动\n（涉军涉密）"）
        if not _is_readable_work(work):
            continue  # 过滤二进制流乱码行
        if len(cells) >= 3:
            rows.append({
                "work": work,
                "progress": cells[1],
                "direction": cells[2],
            })
        elif len(cells) == 2:
            # 容错：方向缺失时补空
            rows.append({"work": work, "progress": cells[1], "direction": ""})
    return rows

def extract_images_from_doc(path, outdir):
    """提取 .doc 中嵌入的图片（Data 流 / 内嵌 OLE），按出现顺序导出 PNG/JPG。"""
    os.makedirs(outdir, exist_ok=True)
    ole = olefile.OleFileIO(path)
    saved = []
    try:
        streams = ole.listdir()
        for parts in streams:
            name = "/".join(parts)
            if "Data" in parts or (len(parts) == 1 and parts[0].startswith("Data")):
                data = ole.openstream(parts).read()
                # 在 Data 流中找 PNG/JPEG 魔数
                for ext, magic in (("png", b"\x89PNG\r\n\x1a\n"), ("jpg", b"\xff\xd8\xff")):
                    idx = 0
                    n = 0
                    while True:
                        idx = data.find(magic, idx)
                        if idx < 0:
                            break
                        # 尽力找结束
                        if ext == "png":
                            end = data.find(b"IEND\xaeB`\x82", idx)
                            end = end + 8 if end > 0 else len(data)
                        else:
                            end = data.find(b"\xff\xd9", idx + 2)
                            end = end + 2 if end > 0 else len(data)
                        blob = data[idx:end]
                        if len(blob) > 1000:
                            fn = os.path.join(outdir, f"img_{n+1}.{ext}")
                            with open(fn, "wb") as f:
                                f.write(blob)
                            saved.append(fn)
                            n += 1
                        idx = idx + len(magic)
    finally:
        ole.close()
    return saved

if __name__ == "__main__":
    path = sys.argv[1]

    # 方向表解析模式
    if "--direction" in sys.argv:
        rows = parse_direction_table(path)
        dpath = sys.argv[sys.argv.index("--direction") + 1]
        with open(dpath, "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=1)
        print(f"[方向表 {len(rows)} 条 -> {dpath}]")
        for r in rows:
            print(f"  - {r['work']} | 方向: {r['direction'][:30]}")
        sys.exit(0)

    text = extract_text_from_doc(path)
    print(f"===== {path} =====")
    print(text)
    if "--out" in sys.argv:
        outdir = sys.argv[sys.argv.index("--out") + 1]
        imgs = extract_images_from_doc(path, outdir)
        if imgs:
            print(f"\n[导出图片 {len(imgs)} 张]")
            for i in imgs:
                print(f"  {i}")
        else:
            print("\n[未发现内嵌图片]")
    if "--text" in sys.argv:
        tpath = sys.argv[sys.argv.index("--text") + 1]
        with open(tpath, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"\n[文本已写入 {tpath}]")
