# -*- coding: utf-8 -*-
"""解析周报数据源 Excel，输出结构化 JSON。

支持三类输入：
1. 《WK{编号}周报填写.xlsx》 → 项目进度成本表 / 资金计划表 / 物资计划表
2. 《每周更新.xlsx》 → 科研组每周更新表
3. 《政策性课题、项目申报台账.xlsx》 → 科研组工作内容表（含汇总统计、子行）

用法:
    python parse_excel.py <excel路径> [--type 项目|资金|物资|每周更新|台账] [--out out.json]
    不指定 --type 时自动识别 sheet 名。

输出 JSON 结构（项目）:
    {"projects": [{"xuhao":..,"center":..,"code":..,"name":..,"leader":..,"stage":..,
                   "hours":..,"reported":..,"progress":..,"node":..,"light":..,"months":..,
                   "ziyan":true/false}], "stats": {"红":N,"黄":N,"绿":N,"总":N}}

资金/物资: {"rows": [[...], ...], "headers": [...]}
每周更新: {"rows": [...], "headers": [...]}
台账: {"summary": "...", "rows": [[...], ...], "headers": [...]}
"""
import sys
import json
import datetime
import openpyxl

def _val(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y/%m/%d")
    if isinstance(v, datetime.date):
        return v.strftime("%Y/%m/%d")
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()

def parse_projects(wb):
    """项目进度成本表 → 项目列表 + 统计。自动定位含'项目进度成本表'的 sheet。"""
    name = None
    for sn in wb.sheetnames:
        if "项目进度成本表" in sn:
            name = sn
            break
    if not name:
        # 兜底：第一个非空 sheet
        name = wb.sheetnames[0]
    ws = wb[name]
    projects = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        # 需要核心列非空才视为数据行
        if row[2] is None and row[3] is None:
            continue
        p = {
            "xuhao": _val(row[0]),
            "center": _val(row[1]),
            "code": _val(row[2]),
            "name": _val(row[3]),
            "leader": _val(row[4]),
            "stage": _val(row[5]),
            "hours": _val(row[6]),
            "reported": _val(row[7]),
            "progress": _val(row[8]),
            "node": _val(row[9]),
            "light": _val(row[10]),
            "months": _val(row[11]),
            "ziyan": _val(row[22]) == "是",   # W列(23) 自研
            "paused": _val(row[27]) if len(row) > 27 else "",  # AB列(28) 暂停
        }
        projects.append(p)
    # 统计
    stats = {"红": 0, "黄": 0, "绿": 0, "总": len(projects)}
    for p in projects:
        if p["light"] == "红":
            stats["红"] += 1
        elif p["light"] == "黄":
            stats["黄"] += 1
        elif p["light"] == "绿":
            stats["绿"] += 1
    return {"projects": projects, "stats": stats}

def parse_table_sheet(ws, skip_rows=2, has_header=True):
    """通用表格解析：跳过 skip_rows 行，返回 {headers, rows}。空行自动跳过。"""
    rows = []
    headers = []
    for ri, row in enumerate(ws.iter_rows(values_only=True), 1):
        if ri <= skip_rows:
            continue
        vals = [_val(v) for v in row]
        if ri == skip_rows + 1 and has_header:
            headers = vals
            continue
        if any(vals):
            rows.append(vals)
    return {"headers": headers, "rows": rows}

def parse_fund(wb):
    """资金计划表：R1 标题、R2 表头、R3 起数据；剔除"合计"残留行（PPT 侧重算）。"""
    name = None
    for sn in wb.sheetnames:
        if "资金计划表" in sn:
            name = sn
            break
    if not name:
        return {"headers": [], "rows": [], "found": False}
    data = parse_table_sheet(wb[name], skip_rows=1)
    rows = [r for r in data["rows"]
            if str(r[0]).strip() not in ("合计", "总计", "") and any(str(v).strip() for v in r[1:])]
    return {"headers": data["headers"], "rows": rows, "found": True}

def parse_material(wb):
    """物资计划表：R1 标题、R2 表头、R3 起数据；剔除仅有序号的空模板行。"""
    name = None
    for sn in wb.sheetnames:
        if "物资计划表" in sn:
            name = sn
            break
    if not name:
        return {"headers": [], "rows": [], "found": False}
    data = parse_table_sheet(wb[name], skip_rows=1)
    rows = [r for r in data["rows"]
            if str(r[0]).strip() not in ("合计", "总计", "") and any(str(v).strip() for v in r[1:])]
    return {"headers": data["headers"], "rows": rows, "found": True}

def parse_weekly(wb):
    """科研组每周更新.xlsx：Sheet1，R1 标题、R2 统计、R3 表头、R4 起数据。"""
    ws = wb[wb.sheetnames[0]]
    return parse_table_sheet(ws, skip_rows=3)

def parse_ledger(wb):
    """科研组台账.xlsx：Sheet1，R1 标题、R2 汇总、R3 表头、R4 起数据（含子行）。"""
    ws = wb[wb.sheetnames[0]]
    summary = ""
    r2 = [c.value for c in ws[2] if c.value is not None]
    if r2:
        # 汇总行可能拆在多个单元格（如 A2 标题 + B2 分类明细），全部拼接提取数字
        summary = " ".join(str(v).strip() for v in r2)
    # skip_rows=2：R1 标题、R2 汇总跳过；R3 表头 → headers，R4 起 → 数据（含序号 1）
    return {"summary": summary, **parse_table_sheet(ws, skip_rows=2)}

def main():
    path = sys.argv[1]
    t = None
    if "--type" in sys.argv:
        t = sys.argv[sys.argv.index("--type") + 1]
    out = None
    if "--out" in sys.argv:
        out = sys.argv[sys.argv.index("--out") + 1]

    wb = openpyxl.load_workbook(path, data_only=True)
    result = {}
    if t is None:
        # 自动识别
        for sn in wb.sheetnames:
            if "项目进度成本表" in sn:
                t = "项目"
                break
        else:
            if "资金" in "".join(wb.sheetnames):
                t = "资金"
            elif "物资" in "".join(wb.sheetnames):
                t = "物资"
            else:
                t = "台账"

    if t == "项目":
        result = parse_projects(wb)
    elif t == "资金":
        result = parse_fund(wb)
    elif t == "物资":
        result = parse_material(wb)
    elif t == "每周更新":
        result = parse_weekly(wb)
    elif t == "台账":
        result = parse_ledger(wb)
    else:
        result = parse_projects(wb)

    js = json.dumps(result, ensure_ascii=False, indent=1)
    if out:
        with open(out, "w", encoding="utf-8") as f:
            f.write(js)
        print(f"written -> {out}")
        print(js[:2000])
    else:
        print(js)

if __name__ == "__main__":
    main()
