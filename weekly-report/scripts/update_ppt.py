# -*- coding: utf-8 -*-
"""周报 PPT 更新引擎 —— 在"上周 PPT"基础上产出"本周 PPT"。

设计原则：
- 所有修改保留原版式/字体/图表样式，只换内容
- 表格文本替换保留原 run 格式（不直接 cell.text=）
- 项目表更新三规则：同编号更新 / 活跃新增 / 未提及保留

用法:
    python update_ppt.py --ppt <上周PPT> --out <新PPT> \
        --wk 35 --dates "8/21-8/28" \
        [--projects projects.json]      # parse_excel.py --type 项目 输出
        [--fund fund.json]              # --type 资金
        [--material material.json]      # --type 物资
        [--weekly weekly.json]          # --type 每周更新
        [--ledger ledger.json]          # --type 台账
        [--direction direction.json]    # [{"work":..,"progress":..,"direction":..}]
        [--quality quality.json]        # {"paragraphs":[{"title":..,"body":..}], "table":[[..]]}
        [--ops ops.json]                # {"items":["1.xxx",...]}
        [--cover <图片文件>]            # 课题申报成果图1
        [--cover2 <图片文件>]           # 课题申报成果图2
"""
import sys, os, json, re, copy, argparse
from pptx import Presentation
from pptx.util import Emu
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

# ============ 通用工具 ============

def iter_shapes(shapes):
    """递归遍历所有形状（含组合内）。"""
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:  # GROUP
            yield from iter_shapes(sh.shapes)

def find_first(iterable, pred):
    for x in iterable:
        if pred(x):
            return x
    return None

def get_shape_text(sh):
    if not sh.has_text_frame:
        return ""
    return sh.text_frame.text

def find_text_shape(slide, text_pred):
    """在页内找第一个文本满足条件的形状（含组合内）。"""
    return find_first(iter_shapes(slide.shapes), lambda s: s.has_text_frame and text_pred(s.text_frame.text))

def find_table(slide, ncols=None, nrows=None):
    for sh in slide.shapes:
        if sh.shape_type == 19:
            t = sh.table
            if ncols is not None and len(t.columns) != ncols:
                continue
            if nrows is not None and len(t.rows) != nrows:
                continue
            return t
    return None

def find_chart(slide):
    for sh in slide.shapes:
        if sh.shape_type == 3:
            return sh.chart
    return None

# ============ 文本替换（保格式） ============

def set_para_text(p, text):
    """替换段落文本，保留第一个 run 的字体格式；\n 转为 <a:br> 换行。"""
    runs = p.runs
    rPr = None
    if runs:
        rPr_el = runs[0]._r.find(qn('a:rPr'))
        if rPr_el is not None:
            rPr = copy.deepcopy(rPr_el)
    for r in list(runs):
        r._r.getparent().remove(r._r)
    lines = str(text).split('\n')
    for i, line in enumerate(lines):
        if i > 0:
            br = p._p.makeelement(qn('a:br'), {})
            p._p.append(br)
        if line or i == 0:
            r_el = p._p.makeelement(qn('a:r'), {})
            if rPr is not None:
                r_el.append(copy.deepcopy(rPr))
            t_el = p._p.makeelement(qn('a:t'), {})
            t_el.text = line
            t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r_el.append(t_el)
            p._p.append(r_el)

def set_cell_text(cell, text):
    """替换表格单元格文本（保留格式），多余段落删除。"""
    tf = cell.text_frame
    set_para_text(tf.paragraphs[0], text)
    for para in tf.paragraphs[1:]:
        para._p.getparent().remove(para._p)

def set_shape_text(shape, text):
    """替换形状文本框全部内容（单段落）。"""
    tf = shape.text_frame
    set_para_text(tf.paragraphs[0], text)
    for para in tf.paragraphs[1:]:
        para._p.getparent().remove(para._p)

def update_digit_runs(text_shape, numbers):
    """按顺序更新文本框中纯数字 run 的文本（用于统计行/椭圆数字）。"""
    idx = 0
    for p in text_shape.text_frame.paragraphs:
        for r in p.runs:
            if r.text.strip().isdigit() and idx < len(numbers):
                r.text = str(numbers[idx])
                idx += 1
    return idx

# ============ 表格行增删 ============

def table_add_row(table, template_idx=-1):
    """复制模板行（默认最后一行）追加到表尾，返回新行。"""
    tbl = table._tbl
    trs = tbl.findall(qn('a:tr'))
    tpl = trs[template_idx]
    new_tr = copy.deepcopy(tpl)
    tbl.append(new_tr)
    return list(table.rows)[-1]

def table_del_row(table, idx):
    tbl = table._tbl
    trs = tbl.findall(qn('a:tr'))
    trs[idx].getparent().remove(trs[idx])

def table_set_data_rows(table, n_data, header_rows=1):
    """把表格数据行数调整为 n_data（表头行保留）。返回是否发生增删。"""
    total = len(table.rows)
    cur_data = total - header_rows
    if cur_data < n_data:
        for _ in range(n_data - cur_data):
            table_add_row(table, -1)
        return True
    if cur_data > n_data:
        for _ in range(cur_data - n_data):
            table_del_row(table, total - 1)
            total -= 1
        return True
    return False

def auto_row_heights(table, header_h=30.0, max_h=345.0):
    """根据数据行数自动分配行高（单位 pt），保持表格整体不超页面。"""
    n = len(table.rows)
    if n <= 1:
        return
    data_n = n - 1
    data_h = min(55.0, max(18.0, (max_h - header_h) / data_n))
    for ri, row in enumerate(table.rows):
        h = header_h if ri == 0 else data_h
        try:
            row.height = Emu(int(h * 12700))
        except Exception:
            pass

# ============ 幻灯片复制与移动 ============

def copy_slide(prs, src):
    """复制幻灯片（含表格/文本框/图表/图片），返回新 slide。"""
    layout = src.slide_layout
    new = prs.slides.add_slide(layout)
    for shp in list(new.shapes):
        shp._element.getparent().remove(shp._element)
    spTree = src.shapes._spTree
    for child in list(spTree):
        if child.tag == qn('p:spTree'):
            continue
        new.shapes._spTree.append(copy.deepcopy(child))
    # 重映射图片关系
    src_part = src.part
    for shp in new.shapes:
        for blip in shp._element.iter(qn('a:blip')):
            rId = blip.get(qn('r:embed'))
            if rId:
                try:
                    img_part = src_part.related_part(rId)
                    new_rId = new.part.relate_to(img_part, RT.IMAGE)
                    blip.set(qn('r:embed'), new_rId)
                except KeyError:
                    continue
    return new

def move_slide_after(prs, slide, after_slide):
    """把 slide 移动到 after_slide 之后。"""
    sldIdLst = prs.slides._sldIdLst
    src = ref = None
    for sid in list(sldIdLst):
        if sid.get('id') == str(slide.slide_id):
            src = sid
        if sid.get('id') == str(after_slide.slide_id):
            ref = sid
    if src is not None and ref is not None:
        ref.addnext(src)

def insert_slide_after(prs, src_slide, after_slide):
    """复制 src_slide 并插入到 after_slide 之后，返回新页。"""
    new = copy_slide(prs, src_slide)
    move_slide_after(prs, new, after_slide)
    return new

def remove_slide(prs, slide):
    """删除一页：移除 _sldIdLst 中对应的 sldId 条目（part 不再被引用）。"""
    id_lst = prs.slides._sldIdLst
    for sldId in list(id_lst):
        if prs.part.related_part(sldId.get(qn('r:id'))) is slide.part:
            id_lst.remove(sldId)
            return True
    return False

# ============ 页面定位 ============

def get_slide_title(slide):
    """提取页面顶部标题（'计划组 — 项目仪表盘-WK35' 之类）。"""
    sh = find_text_shape(slide, lambda t: bool(re.match(r'^(计划组|科研组|质量组|运维组)', t)))
    if sh is None:
        return ""
    return sh.text_frame.text.strip()

def title_key(title):
    """'计划组 — 立方星项目办项目进度WK35' -> '立方星项目办项目进度'"""
    t = re.sub(r'^.*?—\s*', '', title)
    t = re.sub(r'WK\d+.*$', '', t)
    return t.strip()

def find_slide_by_title(prs, key_pred):
    """按标题 key 找所有页。"""
    res = []
    for slide in prs.slides:
        t = get_slide_title(slide)
        if t and key_pred(title_key(t)):
            res.append(slide)
    return res

# ============ 步骤 1/2: 封面日期 + WK 编号 ============

def update_cover_date(prs, new_date):
    """更新封面日期（形状文本形如 '8/14-8/21'）。"""
    for slide in prs.slides:
        sh = find_text_shape(slide, lambda t: bool(re.match(r'^\d{1,2}/\d{1,2}-\d{1,2}/\d{1,2}$', t.strip())))
        if sh is not None:
            set_shape_text(sh, new_date)
            return True
    return False

def update_wk_number(prs, old_wk, new_wk):
    """更新所有页面标题中的 WK 编号（只动标题，不动正文）。"""
    cnt = 0
    pat = re.compile(r'WK\s*' + str(old_wk))
    for slide in prs.slides:
        sh = find_text_shape(slide, lambda t: bool(re.match(r'^(计划组|科研组|质量组|运维组)', t)))
        if sh is None:
            continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if pat.search(r.text):
                    r.text = pat.sub(f'WK{new_wk}', r.text)
                    cnt += 1
    return cnt

def find_project_table(slide):
    """查找项目进度表，支持 10 列或 12 列底稿，返回 (table, ncols)。"""
    tbl = find_table(slide, ncols=12)
    if tbl is not None:
        return tbl, 12
    tbl = find_table(slide, ncols=10)
    if tbl is not None:
        return tbl, 10
    return None, 0

def is_project_header(tbl):
    """项目表表头校验：R0 含'序号'且含'警示灯'（区分其他 10 列表）。"""
    if len(tbl.rows) < 1:
        return False
    head = [c.text.strip() for c in tbl.rows[0].cells]
    return '序号' in head and any('警示灯' in h for h in head)

def collect_project_pages_by_content(prs):
    """无标题底稿（展览版）回退：按表格内容分组收集项目表页。
    数据行单一中心的页 → 该中心组；混合中心页 → 自研组（活跃汇总页）。"""
    pages = {}
    for slide in prs.slides:
        tbl, ncols = find_project_table(slide)
        if tbl is None or not is_project_header(tbl):
            continue
        centers = set()
        for row in list(tbl.rows)[1:]:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) > 2 and cells[2]:
                centers.add(cells[1])
        if not centers:
            continue
        if len(centers) == 1:
            key = CENTER_KEYS.get(next(iter(centers)))
            if key is None:
                continue
        else:
            key = '自研项目进度'
        pages.setdefault(key, []).append((slide, tbl, ncols))
    return pages

# ============ 步骤 3: 项目表更新 ============

CENTER_KEYS = {
    '立方星项目办': '立方星项目办项目进度',
    '通信星项目办': '通信星项目办项目进度',
    '灵鹊星通': '灵鹊星通项目进度',
    '卫星工程中心': '工程中心项目进度',
    '卫星数据中心': '数据中心项目进度',
    '卫星研发中心': '研发中心项目进度',
    '卫星应用中心': '应用中心项目进度',
    '遥感星项目办': '遥感星项目办项目进度',
}

def read_project_table(tbl):
    """读 12 列项目表 → [{'cells':[12个文本], 'code':...}]（跳过表头与空行）。"""
    rows = []
    for row in list(tbl.rows)[1:]:
        cells = [c.text.replace('\n', '/').strip() for c in row.cells]
        code = cells[2] if len(cells) > 2 else ''
        if code or any(cells):
            rows.append({'cells': cells, 'code': code})
    return rows

def write_project_row(row_obj, proj, xh):
    """把 Excel 项目数据写入一行（12 列）。"""
    cells = row_obj['cells']
    cols = [
        str(xh), proj.get('center', ''), proj.get('code', ''), proj.get('name', ''),
        proj.get('leader', ''), proj.get('stage', ''), proj.get('hours', ''),
        proj.get('reported', ''), proj.get('progress', ''), proj.get('node', ''),
        proj.get('light', ''), proj.get('months', ''),
    ]
    for i, c in enumerate(cols[:12]):
        if c != cells[i] or True:
            pass
    return cols

def is_active(proj):
    """活跃判定：进展非空且非'无'（以及'/'等占位）。"""
    pg = (proj.get('progress') or '').strip()
    pg2 = re.sub(r'[/\s]', '', pg)
    return pg2 != '' and pg2 != '无'

def update_project_tables(prs, projects):
    """核心：按三规则更新所有项目表（10/12 列自适应）。返回 (统计, 未加入清单)。"""
    pages = {}
    for slide in prs.slides:
        t = get_slide_title(slide)
        if not t:
            continue
        key = title_key(t)
        if '项目进度' not in key:
            continue
        tbl, ncols = find_project_table(slide)
        if tbl is None:
            continue
        pages.setdefault(key, []).append((slide, tbl, ncols))
    if not pages:
        # 无标题底稿（展览版）：按表格内容分组
        pages = collect_project_pages_by_content(prs)

    # Excel 数据分组
    ziyan = [p for p in projects if p.get('ziyan')]
    others = [p for p in projects if not p.get('ziyan')]
    by_center = {}
    for p in others:
        by_center.setdefault(p['center'], []).append(p)

    groups = [('自研项目进度', ziyan)]
    for center, ps in by_center.items():
        key = CENTER_KEYS.get(center)
        if key:
            groups.append((key, ps))

    not_added = []  # 低活跃未加入清单
    # 旧 PPT 全量编号（用于判定"新增"）
    old_codes = set()
    for key, sl_tbls in pages.items():
        for _, tbl, _ in sl_tbls:
            for r in read_project_table(tbl):
                if r['code']:
                    old_codes.add(r['code'])

    # 处理每个分组
    for key, new_ps in groups:
        # 新项目（Excel 有、旧 PPT 无）：按活跃度筛选
        active_new = []
        for p in new_ps:
            if p['code'] in old_codes:
                continue
            if is_active(p):
                active_new.append(p)
            else:
                not_added.append(p)
        # 计算最终行：先收集旧行
        if key not in pages:
            continue  # 该组无旧页（如遥感星），暂不支持新建页，记录
        sl_tbls = pages[key]
        # 逐页读出旧行
        old_rows = []
        for _, tbl, _ in sl_tbls:
            rs = read_project_table(tbl)
            old_rows.extend(rs)
        # 构建最终行内容（保持旧行顺序；更新匹配；追加新增）
        new_by_code = {p['code']: p for p in new_ps}
        final_rows = []
        seen = set()
        for r in old_rows:
            code = r['code']
            if code and code in new_by_code:
                final_rows.append({'src': 'new', 'proj': new_by_code[code]})
                seen.add(code)
            elif code:
                final_rows.append({'src': 'keep', 'cells': r['cells'], 'code': code})
                seen.add(code)
            else:
                final_rows.append({'src': 'drop'})
        for p in active_new:
            if p['code'] not in seen:
                final_rows.append({'src': 'new', 'proj': p})
        # 序号重排
        xh = 1
        for r in final_rows:
            if r['src'] in ('new', 'keep'):
                r['xh'] = xh
                xh += 1
        # 写回：需要多少页？（一页最多容纳 12 行，新增在页内消化；仅当行数真正超容时才拆页，
        # 避免小量"活跃新增"导致页数膨胀、整份 PPT 页面后移）
        per_page = 12
        n_total = sum(1 for r in final_rows if r['src'] in ('new', 'keep'))
        needed_pages = max(1, (n_total + per_page - 1) // per_page)
        if needed_pages > len(sl_tbls):
            tpl_slide, tpl_tbl, _ = sl_tbls[-1]
            for _ in range(needed_pages - len(sl_tbls)):
                new_slide = insert_slide_after(prs, tpl_slide, sl_tbls[-1][0])
                new_tbl, new_ncols = find_project_table(new_slide)
                # 清空数据行（保留表头）
                for ri in range(len(new_tbl.rows) - 1, 0, -1):
                    table_del_row(new_tbl, ri)
                table_add_row(new_tbl, 0)
                sl_tbls.append((new_slide, new_tbl, new_ncols))
        elif needed_pages < len(sl_tbls):
            # 删除多余页（保留前 needed_pages 页）
            for slide, _, _ in sl_tbls[needed_pages:]:
                remove_slide(prs, slide)
            sl_tbls = sl_tbls[:needed_pages]
        # 填充页面
        ri = 0
        for pi, (slide, tbl, nc) in enumerate(sl_tbls[:needed_pages]):
            page_rows = [r for r in final_rows if r['src'] in ('new', 'keep')][pi * per_page:(pi + 1) * per_page]
            table_set_data_rows(tbl, len(page_rows))
            for ci, r in enumerate(page_rows):
                row_el = tbl.rows[ci + 1]
                if r['src'] == 'new':
                    p = r['proj']
                    if nc == 12:
                        cols = [str(r['xh']), p.get('center', ''), p.get('code', ''), p.get('name', ''),
                                p.get('leader', ''), p.get('stage', ''), p.get('hours', ''),
                                p.get('reported', ''), p.get('progress', ''), p.get('node', ''),
                                p.get('light', ''), p.get('months', '')]
                    else:  # 10 列底稿（无"项目名称/本周进展"列）
                        cols = [str(r['xh']), p.get('center', ''), p.get('code', ''),
                                p.get('leader', ''), p.get('stage', ''), p.get('hours', ''),
                                p.get('reported', ''), p.get('node', ''),
                                p.get('light', ''), p.get('months', '')]
                else:
                    cols = list(r['cells'])
                    cols[0] = str(r['xh'])
                for cci, val in enumerate(cols[:nc]):
                    set_cell_text(row_el.cells[cci], val)
            auto_row_heights(tbl)
            # 分页提示
            has_more = (pi + 1 < needed_pages)
            set_page_continue_hint(slide, has_more)
    return not_added

def set_page_continue_hint(slide, has_more):
    """控制右下角'接下一页'提示（有则显示/隐藏）。"""
    for sh in slide.shapes:
        if sh.has_text_frame and '接下一页' in sh.text_frame.text:
            if has_more:
                pass  # 保持可见
            else:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        r.text = ''
            return
    # 无提示形状且需要时：跳过（不影响主流程）

# ============ 步骤 4: 仪表盘 ============

def stats_from_tables(prs):
    """重新读所有项目进度表（10/12 列）→ 统计红黄绿。"""
    centers = ['立方星项目办', '灵鹊星通', '通信星项目办', '卫星工程中心',
               '卫星数据中心', '卫星研发中心', '卫星应用中心', '遥感星项目办']
    data = {c: {'红': 0, '黄': 0, '绿': 0, '总': 0} for c in centers}
    total = 0
    # 标题模式：底稿含"项目进度"标题页；否则按表头校验统计所有项目表（展览版）
    has_titled = False
    for slide in prs.slides:
        t = get_slide_title(slide)
        if t and '项目进度' in title_key(t):
            has_titled = True
            break
    for slide in prs.slides:
        t = get_slide_title(slide)
        tbl, ncols = find_project_table(slide)
        if tbl is None:
            continue
        if has_titled:
            if not t or '项目进度' not in title_key(t):
                continue
        else:
            if not is_project_header(tbl):
                continue
        light_idx = 10 if ncols == 12 else 8
        for row in list(tbl.rows)[1:]:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) < light_idx + 1 or not cells[2]:
                continue
            center, light = cells[1], cells[light_idx]
            total += 1
            if center in data:
                data[center]['总'] += 1
                if light == '红':
                    data[center]['红'] += 1
                elif light == '黄':
                    data[center]['黄'] += 1
                elif light == '绿':
                    data[center]['绿'] += 1
    red = sum(v['红'] for v in data.values())
    yellow = sum(v['黄'] for v in data.values())
    green = sum(v['绿'] for v in data.values())
    return data, red, yellow, green, total

def _set_cache_values(cache_el, values, numeric=True):
    for pt in cache_el.findall(qn('c:pt')):
        cache_el.remove(pt)
    ptCount = cache_el.find(qn('c:ptCount'))
    if ptCount is not None:
        ptCount.set('val', str(len(values)))
    for i, v in enumerate(values):
        pt = cache_el.makeelement(qn('c:pt'), {'idx': str(i)})
        v_el = cache_el.makeelement(qn('c:v'), {})
        v_el.text = str(v)
        pt.append(v_el)
        cache_el.append(pt)

def update_chart_data(chart, categories, series_data, pct_data):
    """更新图表 categories 与 series 数值（保留图表类型/样式）。"""
    try:
        plot0 = chart.plots[0]
        # categories
        cat_el = plot0._element.find(qn('c:cat'))
        if cat_el is not None:
            for tag in ('c:strCache', 'c:numCache'):
                cache = cat_el.find(qn(tag))
                if cache is not None:
                    _set_cache_values(cache, categories, numeric=False)
                    break
        # series
        for s, vals in zip(plot0.series, series_data):
            ser = s._element
            val_el = ser.find(qn('c:val'))
            if val_el is None:
                continue
            for tag in ('c:numCache',):
                numRef = val_el.find(qn('c:numRef'))
                cache = numRef.find(qn(tag)) if numRef is not None else None
                if cache is not None:
                    _set_cache_values(cache, vals)
                    break
        if len(chart.plots) > 1:
            plot1 = chart.plots[1]
            cat1 = plot1._element.find(qn('c:cat'))
            if cat1 is not None:
                for tag in ('c:strCache', 'c:numCache'):
                    cache = cat1.find(qn(tag))
                    if cache is not None:
                        _set_cache_values(cache, categories, numeric=False)
                        break
            if plot1.series:
                ser = plot1.series[0]._element
                val_el = ser.find(qn('c:val'))
                if val_el is not None:
                    numRef = val_el.find(qn('c:numRef'))
                    cache = numRef.find(qn('c:numCache')) if numRef is not None else None
                    if cache is not None:
                        _set_cache_values(cache, pct_data)
    except Exception as e:
        print(f"[warn] 图表更新失败: {e}")

def update_dashboard(prs):
    """根据项目表现状更新 P2 数字/图表、P3 统计表。"""
    data, red, yellow, green, total = stats_from_tables(prs)
    centers = ['立方星项目办', '灵鹊星通', '通信星项目办', '卫星工程中心',
               '卫星数据中心', '卫星研发中心', '卫星应用中心', '遥感星项目办']
    # 找仪表盘页（标题含'项目仪表盘'；无标题底稿回退：含"本周活跃项目总数"文本的页）
    dash_pages = find_slide_by_title(prs, lambda k: '项目仪表盘' in k)
    if not dash_pages:
        for slide in prs.slides:
            if find_text_shape(slide, lambda t: '本周活跃项目总数' in t) is not None:
                dash_pages = [slide]
                break
    if not dash_pages:
        return
    dash = dash_pages[0]
    dash2 = dash_pages[1] if len(dash_pages) > 1 else dash
    # 顶部总数
    sh = find_text_shape(dash, lambda t: '本周活跃项目总数' in t)
    if sh is not None:
        update_digit_runs(sh, [total])
    # 椭圆数字：绿/黄/红（顺序按页内出现）
    ellipses = []
    for sh in dash.shapes:
        if sh.shape_type == 1 and sh.has_text_frame:
            txt = sh.text_frame.text.strip()
            if txt.isdigit():
                ellipses.append(sh)
    # 椭圆顺序：绿(25) 黄(4) 红(1) —— 按位置从左到右
    ellipses.sort(key=lambda s: (s.left or 0))
    nums = [green, yellow, red]
    for i, sh in enumerate(ellipses[:3]):
        set_shape_text(sh, str(nums[i]))
    # 图表
    chart = find_chart(dash)
    if chart is not None:
        cats = centers
        sdata = [[data[c]['红'] for c in centers],
                 [data[c]['黄'] for c in centers],
                 [data[c]['绿'] for c in centers],
                 [data[c]['总'] for c in centers]]
        pct = [round(data[c]['红'] / data[c]['总'], 3) if data[c]['总'] else 0 for c in centers]
        update_chart_data(chart, cats, sdata, pct)
    # P3 统计表
    tbl = find_table(dash2, ncols=6)
    if tbl is not None:
        # R1-R8 中心，R9 合计
        for i, c in enumerate(centers):
            if i + 1 < len(tbl.rows):
                d = data[c]
                row = tbl.rows[i + 1]
                pct_txt = f"{round(d['红'] / d['总'] * 100)}%" if d['总'] else "0%"
                vals = [c, str(d['红']) if d['红'] else '', str(d['黄']) if d['黄'] else '',
                        str(d['绿']) if d['绿'] else '', str(d['总']), pct_txt]
                for ci, v in enumerate(vals):
                    set_cell_text(row.cells[ci], v)
        if len(tbl.rows) >= 10:
            row = tbl.rows[9]
            vals = ['合计', str(red) if red else '', str(yellow) if yellow else '',
                    str(green) if green else '', str(total), '']
            for ci, v in enumerate(vals):
                set_cell_text(row.cells[ci], v)

# ============ 步骤 5: 资金/物资表 ============

FUND_MAP = [0, 1, 2, 3, 4, 5, 9, 8, 10, 11]  # 老版 12 列 Excel 布局兜底（含 发票/到票金额）

PPT_FUND_COLS = ['序号', '反馈日期', '所属中心', '项目名称', '收款方',
                 '金额/元', '优先级', '建议支付节点', '状态', '备注']

def fund_map_for(headers):
    """按 Excel 表头名生成映射（返回按 PPT 列序的 Excel 索引列表，缺列为 None）。

    兼容两种资金表布局：
    - 新版 10 列（与 PPT 同构，优先级在支付节点前）
    - 老版 12 列（多 发票/到票金额，支付节点在优先级前）
    """
    if not headers:
        return None
    h = [str(x).strip() for x in headers]

    def find(name):
        for i, hn in enumerate(h):
            if not hn:
                continue
            if hn == name or (len(name) >= 2 and name in hn) or (len(hn) >= 2 and hn in name):
                return i
        return None

    m = [find(c) for c in PPT_FUND_COLS]
    if sum(1 for x in m if x is not None) >= 8:
        return m
    return None

def update_fund_table(prs, rows, headers=None):
    """资金计划表：Excel 有数据才更新（空则保留）。表头自适应列映射。"""
    # 过滤无效行：全空 或 仅数字 0（公式残留值），避免清空上周资金表
    rows = [r for r in rows if any(str(v).strip() not in ('', '0', '0.0') for v in r)]
    if not rows:
        return
    pages = find_slide_by_title(prs, lambda k: k == '资金计划表')
    if not pages:
        return
    tbl = find_table(pages[0], ncols=10)
    if tbl is None:
        return
    # 列映射：优先按表头名对位，失败则用老版 12 列布局兜底
    mapping = fund_map_for(headers) or FUND_MAP
    # 映射 Excel 行 → PPT 行（安全索引）
    mapped = []
    for r in rows:
        line = []
        for eidx in mapping:
            line.append(r[eidx] if eidx is not None and eidx < len(r) else '')
        mapped.append(line)
    total = sum(int(float(r[5])) for r in rows if len(r) > 5 and str(r[5]).replace('.', '').isdigit())
    n_data = len(mapped)
    table_set_data_rows(tbl, n_data + 1)  # 数据行 + 合计行
    for i, line in enumerate(mapped):
        row = tbl.rows[i + 1]
        for ci, v in enumerate(line):
            set_cell_text(row.cells[ci], v)
    # 合计行
    total_row = tbl.rows[n_data + 1]
    for ci in range(10):
        if ci == 0:
            set_cell_text(total_row.cells[ci], '合计')
        elif ci == 5:
            set_cell_text(total_row.cells[ci], str(total))
        else:
            set_cell_text(total_row.cells[ci], '')
    auto_row_heights(tbl, header_h=28.0, max_h=300.0)

def update_material_table(prs, rows):
    if not rows:
        return
    pages = find_slide_by_title(prs, lambda k: k == '物资计划表')
    if not pages:
        return
    tbl = find_table(pages[0], ncols=10)
    if tbl is None:
        return
    table_set_data_rows(tbl, len(rows))
    for i, line in enumerate(rows):
        row = tbl.rows[i + 1]
        for ci in range(10):
            v = line[ci] if ci < len(line) else ''
            set_cell_text(row.cells[ci], v)
    auto_row_heights(tbl, header_h=28.0, max_h=280.0)

# ============ 步骤 6: 科研组 ============

def update_weekly_table(prs, rows):
    """每周更新表 P18（9 列）。"""
    if rows is None:
        return
    pages = find_slide_by_title(prs, lambda k: k == '每周更新')
    if not pages:
        return
    tbl = find_table(pages[0], ncols=9)
    if tbl is None:
        return
    table_set_data_rows(tbl, len(rows))
    for i, line in enumerate(rows):
        row = tbl.rows[i + 1]
        for ci in range(9):
            v = line[ci] if ci < len(line) else ''
            set_cell_text(row.cells[ci], v)
    auto_row_heights(tbl, header_h=30.0, max_h=300.0)

def update_ledger_tables(prs, rows, summary):
    """科研组工作内容/申报台账页：统计行 + 表格更新（9 列全量表 / 10 列精选表自适应）。"""
    pages = find_slide_by_title(prs, lambda k: k == '工作内容')
    if not pages:
        return
    # 统计行（第一页）
    if summary:
        nums = re.findall(r'(\d+)', summary)
        total = nums[0] if nums else ''
        counts = nums[1:8]
        while len(counts) < 7:
            counts.append('')
        stat_sh = find_text_shape(pages[0], lambda t: '项机会' in t)
        if stat_sh is not None:
            update_digit_runs(stat_sh, [total] + [c for c in counts[:7]])
    # 判定表类型：9 列（30 页式全量）或 10 列（15 页式精选）
    tbl0 = find_table(pages[0], ncols=9)
    ncols = 9
    if tbl0 is None:
        tbl0 = find_table(pages[0], ncols=10)
        ncols = 10
    if tbl0 is None:
        return
    if ncols == 9:
        # —— 9 列全量表：按序号全量替换 + 自动分页 ——
        per_page = 8
        n_total = len(rows)
        needed = max(1, (n_total + per_page - 1) // per_page)
        if needed > len(pages):
            tpl = pages[-1]
            for _ in range(needed - len(pages)):
                new_slide = insert_slide_after(prs, tpl, pages[-1])
                new_tbl = find_table(new_slide, ncols=9)
                if new_tbl is not None:
                    for ri in range(len(new_tbl.rows) - 1, 0, -1):
                        table_del_row(new_tbl, ri)
                    table_add_row(new_tbl, 0)
                pages.append(new_slide)
        elif needed < len(pages):
            for slide in pages[needed:]:
                prs.slides._sldIdLst.remove(slide._element.getparent())
        for pi, slide in enumerate(pages[:needed]):
            tbl = find_table(slide, ncols=9)
            if tbl is None:
                continue
            chunk = rows[pi * per_page:(pi + 1) * per_page]
            table_set_data_rows(tbl, len(chunk))
            for ci, line in enumerate(chunk):
                row = tbl.rows[ci + 1]
                for cci in range(9):
                    v = line[cci] if cci < len(line) else ''
                    set_cell_text(row.cells[cci], v)
            auto_row_heights(tbl, header_h=30.0, max_h=340.0)
    else:
        # —— 10 列精选表：按申报题目同题更新，不自动增行（新增申报只体现在统计行）——
        # Excel 列 → PPT 列：申报题目(1)→1、通知文件(2)→2、送交单位(3)→3、
        # 技术负责人(5)→4、报送日期(7)→8、进展反馈(8)→9；申报主体/预算编制无对应列，跳过保留原值
        E_MAP = [(1, 1), (2, 2), (3, 3), (5, 4), (7, 8), (8, 9)]
        for line in rows:
            if not line or not str(line[1] or '').strip():
                continue  # 子行（序号空）不单独成行，跳过
            for slide in pages:
                tbl = find_table(slide, ncols=10)
                if tbl is None:
                    continue
                for row in list(tbl.rows)[1:]:
                    cells = [c.text.strip() for c in row.cells]
                    if cells[1] and match_work(line[1], cells[1]):
                        for e_i, p_i in E_MAP:
                            if e_i < len(line):
                                set_cell_text(row.cells[p_i], str(line[e_i]))
                        break

def norm_work(s):
    return re.sub(r'[\s\d年月]', '', str(s))

def match_work(a, b):
    """模糊匹配工作/题目：归一化后一方包含另一方，或有 ≥4 字公共子串（如'频轨资源申报' vs '频轨资源相关工作'）。"""
    na, nb = norm_work(a), norm_work(b)
    if not na or not nb:
        return False
    if na in nb or nb in na:
        return True
    g = {na[i:i + 4] for i in range(max(0, len(na) - 3))} & \
        {nb[i:i + 4] for i in range(max(0, len(nb) - 3))}
    return bool(g)

def _dir_text(d):
    """规范化推进方向文本：去掉数据自带'推进方向：'前缀，返回纯内容。"""
    return re.sub(r'^推进方向[：:]\s*', '', str(d or '')).strip()

def update_direction_tables(prs, items):
    """推进方向表 P26-27（4 列）。items: [{'work':..,'progress':..,'direction':..}]"""
    if not items:
        return
    pages = find_slide_by_title(prs, lambda k: '工作内容、进度' in k)
    if not pages:
        return
    # 收集旧行
    all_tbls = []
    for slide in pages:
        tbl = find_table(slide, ncols=4)
        if tbl:
            all_tbls.append((slide, tbl))
    old_rows = []
    for _, tbl in all_tbls:
        for row in list(tbl.rows)[1:]:
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                old_rows.append({'cells': cells, 'slide_tbl': (None, tbl)})
    # 按 work 匹配更新
    matched = set()
    for item in items:
        target = None
        for r in old_rows:
            if r['cells'][1] and match_work(item['work'], r['cells'][1]):
                target = r
                break
        if target:
            target['cells'][2] = item.get('progress', '')
            d = _dir_text(item.get('direction'))
            if d:
                target['cells'][3] = '推进方向：' + d
            matched.add(id(target))
    # 新增条目（未匹配）→ 追加到最后一页
    new_items = [it for it in items if not any(
        match_work(it['work'], r['cells'][1]) for r in old_rows if r['cells'][1])]
    if new_items:
        last_slide, last_tbl = all_tbls[-1]
        xh = len([r for r in old_rows if r['cells'][0].isdigit()])
        for item in new_items:
            xh += 1
            table_add_row(last_tbl, -1)
            row = list(last_tbl.rows)[-1]
            set_cell_text(row.cells[0], str(xh))
            set_cell_text(row.cells[1], item.get('work', ''))
            set_cell_text(row.cells[2], item.get('progress', ''))
            set_cell_text(row.cells[3], '推进方向：' + _dir_text(item.get('direction')))
        auto_row_heights(last_tbl, header_h=30.0, max_h=340.0)
    # 写回匹配行
    for r in old_rows:
        if id(r) in matched:
            slide_tbl = r['slide_tbl'][1]
            for row in list(slide_tbl.rows)[1:]:
                cells = [c.text.strip() for c in row.cells]
                if cells[1] and norm_work(cells[1]) and norm_work(cells[1]) == norm_work(r['cells'][1]):
                    set_cell_text(row.cells[2], r['cells'][2])
                    set_cell_text(row.cells[3], r['cells'][3])

def update_cover_images(prs, img1=None, img2=None):
    """替换课题申报成果页的两张证书图（跳过页眉 logo，按位置取下方大图）。"""
    pages = find_slide_by_title(prs, lambda k: '课题申报成果' in k)
    if not pages:
        return
    slide = pages[0]
    page_h = prs.slide_height
    # 页眉 logo 位于页面顶部；证书图为页面下方的大图
    pics = [sh for sh in slide.shapes
            if sh.shape_type == 13 and (sh.top or 0) > int(page_h * 0.2)]
    pics.sort(key=lambda p: (p.top or 0))
    if not pics:
        return
    imgs = [i for i in (img1, img2) if i]
    for pic, img in zip(pics[:len(imgs)], imgs):
        try:
            pic._element.getparent().remove(pic._element)
            slide.shapes.add_picture(img, Emu(int(pic.left or 0)), Emu(int(pic.top or 0)),
                                     width=Emu(int(pic.width or 0)), height=Emu(int(pic.height or 0)))
        except Exception as e:
            print(f"[warn] 图片替换失败: {e}")

# ============ 步骤 7: 质量/运维 ============

def update_textbox_lines(slide, lines, match_pred):
    """更新文本框段落（保留每段格式）。lines 按顺序对应段落。"""
    sh = find_text_shape(slide, match_pred)
    if sh is None:
        return
    paras = sh.text_frame.paragraphs
    n = min(len(paras), len(lines))
    for i in range(n):
        set_para_text(paras[i], lines[i])
    # 段落数不足则补
    while len(paras) < len(lines):
        p = sh.text_frame.add_paragraph()
        set_para_text(p, lines[len(paras) - 1])
        paras = sh.text_frame.paragraphs

def update_quality(prs, quality):
    """质量组：paragraphs 更新正文，table 更新跟踪表。"""
    if not quality:
        return
    pages = find_slide_by_title(prs, lambda k: k == '工作内容' and False)  # 占位
    # 质量组页 = 标题'质量组 — 工作内容'
    for slide in prs.slides:
        t = get_slide_title(slide)
        if t and t.startswith('质量组'):
            if quality.get('paragraphs'):
                lines = []
                for p in quality['paragraphs']:
                    lines.append(p.get('title', ''))
                    body = p.get('body', '')
                    if body:
                        lines.append('    ' + body)
                update_textbox_lines(slide, lines, lambda txt: bool(re.search(r'[1-9]\.', txt) or '：' in txt))
            if quality.get('table'):
                tbl = find_table(slide, ncols=5)
                if tbl is not None:
                    table_set_data_rows(tbl, len(quality['table']))
                    for i, line in enumerate(quality['table']):
                        row = tbl.rows[i + 1]
                        for ci in range(5):
                            v = line[ci] if ci < len(line) else ''
                            set_cell_text(row.cells[ci], v)
            return

def update_ops(prs, ops):
    if not ops or not ops.get('items'):
        return
    for slide in prs.slides:
        t = get_slide_title(slide)
        if t and t.startswith('运维组'):
            update_textbox_lines(slide, ops['items'], lambda txt: bool(re.match(r'^\s*\d+\.', txt)))
            return

# ============ 主流程 ============

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--ppt', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--wk', type=int, required=True)
    ap.add_argument('--dates', default='')
    ap.add_argument('--projects')
    ap.add_argument('--fund')
    ap.add_argument('--material')
    ap.add_argument('--weekly')
    ap.add_argument('--ledger')
    ap.add_argument('--direction')
    ap.add_argument('--quality')
    ap.add_argument('--ops')
    ap.add_argument('--cover')
    ap.add_argument('--cover2')
    args = ap.parse_args()

    def load(p):
        if not p or not os.path.exists(p):
            return None
        with open(p, encoding='utf-8') as f:
            return json.load(f)

    prs = Presentation(args.ppt)
    old_wk = args.wk - 1

    print(f"[1/8] 复制底稿 {args.ppt}")
    print(f"[2/8] 更新 WK 编号 {old_wk}->{args.wk} ...")
    n = update_wk_number(prs, old_wk, args.wk)
    print(f"       更新标题 {n} 处")
    if args.dates:
        if update_cover_date(prs, args.dates):
            print(f"       封面日期 -> {args.dates}")

    projects = load(args.projects)
    not_added = []
    if projects is not None:
        print(f"[3/8] 更新计划组项目表 ({len(projects.get('projects', []))} 个) ...")
        not_added = update_project_tables(prs, projects.get('projects', []))
    print(f"[4/8] 更新项目仪表盘 ...")
    update_dashboard(prs)

    fund = load(args.fund)
    if fund is not None and fund.get('rows'):
        print(f"[5/8] 更新资金计划表 ({len(fund['rows'])} 行)")
        update_fund_table(prs, fund['rows'], fund.get('headers'))
    material = load(args.material)
    if material is not None and material.get('rows'):
        print(f"[5/8] 更新物资计划表 ({len(material['rows'])} 行)")
        update_material_table(prs, material['rows'])

    weekly = load(args.weekly)
    if weekly is not None:
        print(f"[6/8] 更新科研组每周更新 ({len(weekly.get('rows', []))} 行)")
        update_weekly_table(prs, weekly.get('rows', []))
    ledger = load(args.ledger)
    if ledger is not None:
        print(f"[6/8] 更新科研组工作内容 ({len(ledger.get('rows', []))} 行)")
        update_ledger_tables(prs, ledger.get('rows', []), ledger.get('summary', ''))
    direction = load(args.direction)
    if direction:
        print(f"[6/8] 更新科研组推进方向 ({len(direction)} 条)")
        update_direction_tables(prs, direction)
    if args.cover or args.cover2:
        print(f"[6/8] 更新课题申报成果图片")
        update_cover_images(prs, args.cover, args.cover2)

    quality = load(args.quality)
    if quality:
        print(f"[7/8] 更新质量组")
        update_quality(prs, quality)
    ops = load(args.ops)
    if ops:
        print(f"[7/8] 更新运维组")
        update_ops(prs, ops)

    prs.save(args.out)
    print(f"[8/8] 已保存 -> {args.out}")
    if not_added:
        print(f"\n⚠ 低活跃未加入清单（{len(not_added)} 个）:")
        for p in not_added:
            print(f"  - {p.get('code')} {p.get('name')} (进展: {p.get('progress','')[:30]})")
    print("完成。")

if __name__ == '__main__':
    main()
