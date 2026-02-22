# -*- coding: utf-8 -*-
"""
全页配置驱动版生成器 v3 —— XML 级精确替换

核心思路：
1) 用 generate_report.py 从 Excel 生成"数据稿"PPT（含各页数值、图表、表格）
2) **直接复制 demo** 作为输出文件（保留所有形状、格式、样式、换行 <a:br>）
3) 逐页把数据稿内容**在 XML 级别**回填到 demo 副本的对应控件
   - 文本框：仅替换 <a:t> 文本值，保留 <a:rPr>（字体/字号/颜色/粗细）
   - 图表：用 replace_data 替换数据
   - 表格：逐单元格 XML 级替换文本
4) 第 1 页封面做特殊处理：精确替换日期 run，保留 <a:br> 换行和标题

确保输出 PPT 与 demo 格式完全一致（字号、字体、颜色、换行、空行、标点）。
"""

import argparse
import os
import re
import shutil
import tempfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData

import generate_report as base

# ─── XML 命名空间 ───
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NSMAP = {'a': A_NS, 'p': P_NS}

# 控制字符清洗
_VTAB_PATTERN = re.compile(r'_x000[0-9A-Fa-f]_')


def _sanitize(text):
    """清除 _x000B_ 等控制字符转义"""
    if text is None:
        return ''
    s = str(text)
    s = _VTAB_PATTERN.sub('', s)
    s = s.replace('\x0b', '').replace('\x0c', '')
    return s


# ═══════════════════════════════════════════════════════════════
# XML 级文本操作
# ═══════════════════════════════════════════════════════════════

def _get_para_text(p_elem):
    """从 <a:p> 元素提取纯文本（<a:br> → \\n）"""
    parts = []
    for child in p_elem:
        tag = etree.QName(child.tag).localname
        if tag == 'r':
            t = child.find('a:t', NSMAP)
            if t is not None and t.text:
                parts.append(t.text)
        elif tag == 'br':
            parts.append('\n')
    return ''.join(parts)


def _get_runs(p_elem):
    """获取段落内所有 <a:r> 元素"""
    return p_elem.findall('a:r', NSMAP)


def _get_breaks(p_elem):
    """获取段落内所有 <a:br> 元素"""
    return p_elem.findall('a:br', NSMAP)


def _set_run_text(run_elem, text):
    """设置 run 的 <a:t> 文本值，保留 <a:rPr>"""
    t = run_elem.find('a:t', NSMAP)
    if t is not None:
        t.text = text
    else:
        t = etree.SubElement(run_elem, '{http://schemas.openxmlformats.org/drawingml/2006/main}t')
        t.text = text


def _replace_para_text(target_p, new_text):
    """
    替换段落文本，保留所有格式。

    策略：
    - 如果段落有 <a:br>：按 \\n 拆分文本，分配到 br 前后的 run 组
    - 否则：将全部文本写入第一个 run，清空其余 run
    - 始终保留每个 run 的 <a:rPr>（字体/字号/颜色/粗细）
    """
    runs = _get_runs(target_p)
    breaks = _get_breaks(target_p)

    if not runs:
        return

    if breaks and '\n' in new_text:
        # 有 <a:br>：按换行分段分配到 run 组
        segments = new_text.split('\n')
        # 按 XML 子元素顺序将 run 分组（以 br 为分隔）
        run_groups = []
        current_group = []
        for child in target_p:
            tag = etree.QName(child.tag).localname
            if tag == 'r':
                current_group.append(child)
            elif tag == 'br':
                run_groups.append(current_group)
                current_group = []
        run_groups.append(current_group)

        for gi, group in enumerate(run_groups):
            seg_text = segments[gi] if gi < len(segments) else ''
            if group:
                _set_run_text(group[0], seg_text)
                for r in group[1:]:
                    _set_run_text(r, '')
    else:
        # 无 br：全部文本写入第一个 run，清空其余
        # 若源文本含 \n（来自 <a:br>）但目标无 <a:br>，只取第一行，
        # 避免将后续行重复写入（其余目标段落保持原样）
        text_to_write = new_text.split('\n')[0] if '\n' in new_text else new_text
        _set_run_text(runs[0], text_to_write)
        for r in runs[1:]:
            _set_run_text(r, '')


def _clear_para_text(target_p):
    """清空段落所有 run 的文本"""
    for r in _get_runs(target_p):
        _set_run_text(r, '')


# ═══════════════════════════════════════════════════════════════
# Shape 级操作
# ═══════════════════════════════════════════════════════════════

def _shape_type(shape):
    if getattr(shape, 'has_chart', False):
        return 'chart'
    if getattr(shape, 'has_table', False):
        return 'table'
    if getattr(shape, 'has_text_frame', False):
        return 'text'
    return 'other'


def _shape_center(shape):
    return (int(shape.left + shape.width / 2), int(shape.top + shape.height / 2))


def _dist2(a, b):
    return (a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2


def _iter_shapes_by_type(slide, type_name):
    out = []
    for idx, shp in enumerate(slide.shapes):
        if _shape_type(shp) == type_name:
            out.append((idx, shp))
    return out


def _match_nearest(target_items, src_items):
    """
    两轮匹配：
    1. 短文本先做内容匹配（防止静态标签被交叉覆盖）
    2. 剩余的做最近邻位置匹配
    """
    mapping = []
    used_src = set()
    used_tgt = set()

    # 第一轮：短文本内容匹配
    for ti, tshp in target_items:
        if not getattr(tshp, 'has_text_frame', False):
            continue
        ttext = tshp.text.strip().replace('\x0b', '\n')
        if not ttext or len(ttext) > 30:
            continue
        for si, sshp in src_items:
            if si in used_src:
                continue
            if not getattr(sshp, 'has_text_frame', False):
                continue
            stext = sshp.text.strip().replace('\x0b', '\n')
            if stext == ttext:
                mapping.append((ti, si))
                used_src.add(si)
                used_tgt.add(ti)
                break

    # 第二轮：剩余的做最近邻位置匹配
    for ti, tshp in target_items:
        if ti in used_tgt:
            continue
        tc = _shape_center(tshp)
        best = None
        best_d = None
        for si, sshp in src_items:
            if si in used_src:
                continue
            d = _dist2(tc, _shape_center(sshp))
            if best is None or best_d is None or d < best_d:
                best = si
                best_d = d
        if best is not None:
            used_src.add(best)
            mapping.append((ti, best))
    return mapping


# ═══════════════════════════════════════════════════════════════
# 文本同步（XML 级）
# ═══════════════════════════════════════════════════════════════

def _sync_text_shape(target_shape, src_shape):
    """
    将 src_shape 的文本内容写入 target_shape，
    但完全保留 target_shape 的 XML 格式（rPr / br / pPr）。

    使用"非空段落对齐"策略：
    - 提取两边的非空段落索引
    - 按顺序一一配对
    - 空段落保持不变（保留空行分隔结构）
    """
    src_text = src_shape.text if hasattr(src_shape, 'text') else ''
    if not src_text or not src_text.strip():
        return

    # 内容相同则跳过（防止静态标签被错误交叉覆盖）
    # 统一 \x0b / \n 后比较
    target_text = target_shape.text if hasattr(target_shape, 'text') else ''
    _norm = lambda t: _sanitize(t).replace('\x0b', '').replace('\n', '').strip()
    if _norm(src_text) == _norm(target_text):
        return

    # 获取 <p:txBody> → <a:p> 列表
    t_body = target_shape._element.find(f'.//{{{P_NS}}}txBody')
    s_body = src_shape._element.find(f'.//{{{P_NS}}}txBody')
    if t_body is None:
        t_body = target_shape._element.find(f'.//{{{A_NS}}}txBody')
    if s_body is None:
        s_body = src_shape._element.find(f'.//{{{A_NS}}}txBody')
    if t_body is None or s_body is None:
        return

    t_paras = t_body.findall('a:p', NSMAP)
    s_paras = s_body.findall('a:p', NSMAP)

    # 提取非空段落索引
    t_content = [i for i, p in enumerate(t_paras) if _get_para_text(p).strip()]
    s_content = [i for i, p in enumerate(s_paras) if _get_para_text(p).strip()]

    # 按非空段落一一配对
    n = min(len(t_content), len(s_content))
    for j in range(n):
        ti = t_content[j]
        si = s_content[j]
        src_text_j = _sanitize(_get_para_text(s_paras[si]))
        _replace_para_text(t_paras[ti], src_text_j)

    # 如果源有更多非空段落，追加到最后一个已配对的目标段落
    if len(s_content) > len(t_content) and t_content:
        extra_texts = []
        for j in range(len(t_content), len(s_content)):
            txt = _sanitize(_get_para_text(s_paras[s_content[j]]))
            if txt:
                extra_texts.append(txt)
        if extra_texts:
            last_tp = t_paras[t_content[-1]]
            runs = _get_runs(last_tp)
            if runs:
                t_elem = runs[-1].find('a:t', NSMAP)
                if t_elem is not None:
                    t_elem.text = (t_elem.text or '') + '\n'.join([''] + extra_texts)


# ═══════════════════════════════════════════════════════════════
# 图表数据同步
# ═══════════════════════════════════════════════════════════════

def _extract_chart_data(src_chart):
    chart_data = CategoryChartData()
    categories = []
    try:
        categories = [_sanitize(str(c.label)) if c.label is not None else ''
                      for c in src_chart.plots[0].categories]
    except (AttributeError, TypeError, IndexError):
        pass
    chart_data.categories = categories

    for ser in src_chart.series:
        values = []
        try:
            values = list(ser.values)
        except (AttributeError, TypeError):
            pass
        name = _sanitize(ser.name) if ser.name else 'Series'
        chart_data.add_series(name, values)
    return chart_data


def _sync_chart(target_shape, src_shape):
    try:
        target_shape.chart.replace_data(_extract_chart_data(src_shape.chart))
    except (AttributeError, TypeError):
        pass


# ═══════════════════════════════════════════════════════════════
# 表格数据同步
# ═══════════════════════════════════════════════════════════════

def _sync_table(target_shape, src_shape):
    """替换表格单元格文本，保留目标格式（XML 级）"""
    tt = target_shape.table
    st = src_shape.table
    rows = min(len(tt.rows), len(st.rows))
    cols = min(len(tt.columns), len(st.columns))

    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    for r in range(rows):
        for c in range(cols):
            new_text = _sanitize(st.cell(r, c).text)
            tc_elem = tt.cell(r, c)._tc
            tc_body = tc_elem.find(f'.//{{{a_ns}}}txBody')
            if tc_body is not None:
                paras = tc_body.findall('a:p', NSMAP)
                if paras:
                    _replace_para_text(paras[0], new_text)
                    for p in paras[1:]:
                        _clear_para_text(p)
            else:
                tt.cell(r, c).text = new_text


# ═══════════════════════════════════════════════════════════════
# 标题修复（简版模式下 data draft 无标题，需内联构建）
# ═══════════════════════════════════════════════════════════════

def _change_desc(pct, up_word='提升', down_word='下降', flat_word='基本持平'):
    if abs(pct) <= 5:
        return flat_word
    return f'{up_word}{abs(pct):.0f}%' if pct > 0 else f'{down_word}{abs(pct):.0f}%'


def _rank_change_desc(change):
    if change > 0:
        return f'上升{change}位'
    elif change < 0:
        return f'下滑{abs(change)}位'
    return '保持不变'


def _build_slide_titles(data):
    """
    根据 ReportData 构建各页标题（与 generate_report.py 逻辑一致）。
    返回 dict: {slide_index: title_text}，仅含数据相关标题。
    """
    ms = data.market_share
    titles = {}

    # Page 3 (idx=2): 市场份额概览
    share_chg = _change_desc(ms.share_change)
    rating_chg = _change_desc(ms.rating_change)
    titles[2] = f'{data.report_date_short}，市场份额{share_chg}  收视率{rating_chg}'

    # Page 4 (idx=3): 台组内排名
    org_chg = _rank_change_desc(data.org_rank_change)
    titles[3] = f'央视台组排名第{data.org_rank}位，较前一日{org_chg}'

    # Page 5 (idx=4): 上星频道排名
    ch_chg = _rank_change_desc(data.channel_rank_change)
    titles[4] = f'上星频道排名第{data.channel_rank}位，较前一日{ch_chg}'

    # Page 6 (idx=5): 串单市场份额
    titles[5] = (f'{data.report_date_short}，频道市场份额{ms.cctv17_current_share:.3f}%，'
                 f'收视率{ms.cctv17_current_rating:.3f}%')

    # Page 7 (idx=6): 栏目收视率排名
    titles[6] = f'{data.report_date_short}栏目收视率排名'

    # Page 8 (idx=7): 栏目收视份额排名
    titles[7] = f'{data.report_date_short}栏目收视份额排名'

    # Page 9-13: 静态标题，无需更新
    return titles


def _find_title_shape(slide):
    """查找幻灯片中的标题 shape（按名称识别）"""
    for shp in slide.shapes:
        if not getattr(shp, 'has_text_frame', False):
            continue
        name = shp.name or ''
        if '标题' in name:
            return shp
    return None


def _fix_titles(target_prs, data):
    """修复简版模板中未被数据稿同步的标题文本"""
    titles = _build_slide_titles(data)
    for slide_idx, title_text in titles.items():
        if slide_idx >= len(target_prs.slides):
            continue
        shp = _find_title_shape(target_prs.slides[slide_idx])
        if shp is None:
            continue
        # 找到标题段落并用 XML 级替换
        body = shp._element.find(f'.//{{{P_NS}}}txBody')
        if body is None:
            body = shp._element.find(f'.//{{{A_NS}}}txBody')
        if body is None:
            continue
        paras = body.findall('a:p', NSMAP)
        if paras:
            # 标题通常只有一个段落，有前导空格的 run（如"  收视速报"）需保留空格
            # 但 data-dependent titles 不含前导空格
            _replace_para_text(paras[0], title_text)
            for p in paras[1:]:
                _clear_para_text(p)
        print(f'    ✧ 标题已更新: 第{slide_idx + 1}页 → {title_text[:40]}...')


# ═══════════════════════════════════════════════════════════════
# 第 1 页封面特殊处理
# ═══════════════════════════════════════════════════════════════

def _update_cover(target_slide, src_slide):
    """
    封面页精确更新：
    - 标题 shape：只替换日期相关 run 的 <a:t>，保留 <a:br> 换行和标题文字
    - 署名 shape：只替换年份和期号 run
    """
    # ── 从源幻灯片提取日期和期号 ──
    src_title_text = ''
    src_subtitle_text = ''
    for shp in src_slide.shapes:
        if not getattr(shp, 'has_text_frame', False):
            continue
        text = shp.text
        if '频道收视日报' in text or '农业' in text:
            src_title_text = text
        elif '统筹策划部' in text or '策划部' in text:
            src_subtitle_text = text

    # 提取年月日
    date_match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', src_title_text)
    new_year = date_match.group(1) if date_match else None
    new_month = date_match.group(2) if date_match else None
    new_day = date_match.group(3) if date_match else None

    # 提取期号
    period_match = re.search(r'第(\d+)期', src_subtitle_text)
    new_period = period_match.group(1) if period_match else None

    # ── 更新目标标题 shape 的日期 run ──
    for shp in target_slide.shapes:
        if not getattr(shp, 'has_text_frame', False):
            continue
        full_text = shp.text

        if '频道收视日报' in full_text or '农业' in full_text:
            # 标题 shape：逐 run 精确替换日期数字
            for p in shp.text_frame._element.findall('a:p', NSMAP):
                runs = p.findall('a:r', NSMAP)
                for i, r in enumerate(runs):
                    t = r.find('a:t', NSMAP)
                    if t is None or t.text is None:
                        continue
                    val = t.text.strip()

                    # 年份 run：4 位数字
                    if val.isdigit() and len(val) == 4 and new_year:
                        t.text = new_year
                        continue

                    # 月份 / 日期 run：1-2 位数字，根据相邻 run 判断
                    if val.isdigit() and len(val) <= 2:
                        # 看后一个 run
                        if i + 1 < len(runs):
                            next_t = runs[i + 1].find('a:t', NSMAP)
                            if next_t is not None and next_t.text:
                                if '月' in next_t.text and new_month:
                                    t.text = new_month
                                    continue
                        # 看前一个 run
                        if i > 0:
                            prev_t = runs[i - 1].find('a:t', NSMAP)
                            if prev_t is not None and prev_t.text:
                                if '月' in prev_t.text and new_day:
                                    t.text = new_day
                                    continue

        elif '统筹策划部' in full_text or '策划部' in full_text:
            # 署名 shape：更新年份和期号
            for p in shp.text_frame._element.findall('a:p', NSMAP):
                runs = p.findall('a:r', NSMAP)
                for i, r in enumerate(runs):
                    t = r.find('a:t', NSMAP)
                    if t is None or t.text is None:
                        continue

                    # 年份 run：含空格+4位数字如 " 2026"
                    year_m = re.match(r'^(\s*)(\d{4})$', t.text)
                    if year_m and new_year:
                        t.text = year_m.group(1) + new_year
                        continue

                    # 期号 run：纯数字且前面 run 含"第"
                    val = t.text.strip()
                    if val.isdigit() and new_period and i > 0:
                        prev_t = runs[i - 1].find('a:t', NSMAP)
                        if prev_t is not None and prev_t.text and '第' in prev_t.text:
                            t.text = new_period


# ═══════════════════════════════════════════════════════════════
# 逐页同步
# ═══════════════════════════════════════════════════════════════

def _is_change_label(text):
    """判断是否为变化百分比标签（如 ↑14%、-15%、↓8%）"""
    t = text.strip()
    if not t:
        return False
    return bool(re.match(r'^[↑↓▲▼+-]?\d+\.?\d*%$', t))


def _sync_audience_labels(target_slide, src_slide):
    """
    页面13（频道分类观众规模）专用：按 X 坐标排序匹配变化标签。
    demo 的标签散布在图表区各位置，draft 的标签排成一行。
    按 X 排序后一一对应。
    """
    # 收集 demo/draft 中的变化标签
    t_labels = []
    s_labels = []
    for idx, shp in enumerate(target_slide.shapes):
        if getattr(shp, 'has_text_frame', False) and _is_change_label(shp.text):
            t_labels.append((shp.left, idx, shp))
    for idx, shp in enumerate(src_slide.shapes):
        if getattr(shp, 'has_text_frame', False) and _is_change_label(shp.text):
            s_labels.append((shp.left, idx, shp))

    # 按 X 排序
    t_labels.sort(key=lambda x: x[0])
    s_labels.sort(key=lambda x: x[0])

    n = min(len(t_labels), len(s_labels))
    for i in range(n):
        t_shp = t_labels[i][2]
        s_shp = s_labels[i][2]
        _sync_text_shape(t_shp, s_shp)


def _compute_nice_axis(max_val):
    """计算 Y 轴最大值（模拟 PowerPoint 自动缩放）。"""
    import math
    if max_val <= 0:
        return 100
    raw_interval = max_val / 6
    magnitude = 10 ** math.floor(math.log10(raw_interval))
    candidates = [magnitude, 2 * magnitude, 5 * magnitude, 10 * magnitude]
    best = None
    for interval in candidates:
        n = math.ceil(max_val / interval) + 1          # +1 留余量
        axis_max = n * interval
        n_ticks = n + 1
        if 5 <= n_ticks <= 12:
            if best is None or axis_max < best:
                best = axis_max
    if best is None:
        interval = 2 * magnitude
        best = (math.ceil(max_val / interval) + 1) * interval
    return best


def _sync_and_position_audience(target_slide, src_slide):
    """
    Page 13（频道分类观众规模）：同步图表数据 + 动态定位变化标签。

    1. 同步图表数据（replace_data）
    2. 根据图表数据 **计算** 变化百分比文本
    3. 将标签 X 居中于对应类别柱组，Y 紧贴最高柱上方
    4. 图例区总体变化标签单独处理（仅替换文字）
    """
    C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

    # ── 1. 图表数据同步 ──
    target_charts = _iter_shapes_by_type(target_slide, 'chart')
    src_charts = _iter_shapes_by_type(src_slide, 'chart')
    src_chart_map = {idx: shp for idx, shp in src_charts}
    chart_shape = None
    for ti, si in _match_nearest(target_charts, src_charts):
        t_shp = dict(target_charts)[ti]
        s_shp = src_chart_map[si]
        _sync_chart(t_shp, s_shp)
        chart_shape = t_shp

    if chart_shape is None:
        return

    # ── 1b. 清除 replace_data 产生的 "None" 间隔类别 ──
    cs_tmp = chart_shape.chart._chartSpace
    for str_cache in cs_tmp.findall(f'.//{{{C_NS}}}strCache'):
        for pt in list(str_cache.findall(f'{{{C_NS}}}pt')):
            v = pt.find(f'{{{C_NS}}}v')
            if v is not None and v.text in ('None', ''):
                str_cache.remove(pt)

    # ── 2. 图表几何计算 ──
    chart = chart_shape.chart
    cs = chart._chartSpace
    pa = cs.find(f'.//{{{C_NS}}}plotArea')
    man = pa.find(f'{{{C_NS}}}layout/{{{C_NS}}}manualLayout')

    if man is not None:
        px = float(man.find(f'{{{C_NS}}}x').get('val'))
        py = float(man.find(f'{{{C_NS}}}y').get('val'))
        pw = float(man.find(f'{{{C_NS}}}w').get('val'))
        ph = float(man.find(f'{{{C_NS}}}h').get('val'))
    else:
        px, py, pw, ph = 0.05, 0.15, 0.9, 0.6

    plot_left = chart_shape.left + int(px * chart_shape.width)
    plot_top = chart_shape.top + int(py * chart_shape.height)
    plot_width = int(pw * chart_shape.width)
    plot_height = int(ph * chart_shape.height)
    plot_bottom = plot_top + plot_height

    # ── 3. 读取图表数据 ──
    s0 = list(chart.series[0].values)
    s1 = list(chart.series[1].values)
    num_cats = len(s0)

    # 每个类别的 max 值与变化百分比
    cat_data = []          # (index, max_value, change_text)
    for i in range(num_cats):
        v0 = s0[i] if s0[i] is not None else 0
        v1 = s1[i] if s1[i] is not None else 0
        max_v = max(v0, v1)
        if max_v <= 0:
            continue          # 间隔类别
        change_pct = ((v1 - v0) / v0 * 100) if v0 > 0 else 0
        # 格式：|val|<1% 保留 1 位小数，否则取整；负值用 - 前缀，正值无前缀
        abs_chg = abs(change_pct)
        if abs_chg < 1 and abs_chg > 0:
            val_str = f'{abs_chg:.1f}'
        else:
            val_str = str(round(abs_chg))
        if change_pct < 0:
            text = f'-{val_str}%'
        elif change_pct > 0:
            text = f'{val_str}%'
        else:
            text = '0%'
        cat_data.append((i, max_v, text))

    if not cat_data:
        return

    # ── 4. 计算坐标轴范围 & 设置显式最大值 ──
    overall_max = max(d[1] for d in cat_data)
    axis_max = _compute_nice_axis(overall_max)

    val_ax = pa.find(f'{{{C_NS}}}valAx')
    if val_ax is not None:
        scaling = val_ax.find(f'{{{C_NS}}}scaling')
        if scaling is not None:
            max_elem = scaling.find(f'{{{C_NS}}}max')
            if max_elem is None:
                max_elem = etree.SubElement(scaling, f'{{{C_NS}}}max')
            max_elem.set('val', str(axis_max))

    # ── 5. 分离柱上标签 vs 图例标签 ──
    # 图例标签用 ↑↓▲▼ 箭头前缀；柱上标签用 -/数字前缀
    bar_labels = []
    legend_label = None
    for sh in target_slide.shapes:
        if getattr(sh, 'has_text_frame', False) and _is_change_label(sh.text):
            t = sh.text.strip()
            if t and t[0] in '↑↓▲▼':
                legend_label = sh          # 图例区（如 ↓8%）
            else:
                bar_labels.append(sh)
    bar_labels.sort(key=lambda s: s.left)

    # ── 6. 定位每个柱上标签 ──
    GAP_EMU = 55000          # 标签底边与柱顶的间距（~0.15 cm）
    n = min(len(bar_labels), len(cat_data))
    for i in range(n):
        lbl = bar_labels[i]
        cat_idx, max_v, new_text = cat_data[i]

        # X: 居中于类别柱组
        cat_center_x = plot_left + int(plot_width * (cat_idx + 0.5) / num_cats)
        lbl.left = cat_center_x - lbl.width // 2

        # Y: 紧贴最高柱上方
        bar_top = plot_bottom - int((max_v / axis_max) * plot_height)
        lbl.top = bar_top - lbl.height - GAP_EMU

        # 文字：XML 级替换（保留字体/颜色/粗细格式）
        p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        txBody = lbl._element.find(f'.//{{{A_NS}}}txBody')
        if txBody is None:
            txBody = lbl._element.find(f'.//{{{p_ns}}}txBody')
        if txBody is not None:
            paras = txBody.findall(f'{{{A_NS}}}p')
            if paras:
                _replace_para_text(paras[0], new_text)

    # ── 7. 图例变化标签（仅替换文字） ──
    if legend_label:
        total_s0 = sum(v for v in s0 if v is not None)
        total_s1 = sum(v for v in s1 if v is not None)
        total_chg = round((total_s1 - total_s0) / total_s0 * 100) if total_s0 > 0 else 0
        if total_chg > 0:
            leg_text = f'↑{total_chg}%'
        elif total_chg < 0:
            leg_text = f'↓{abs(total_chg)}%'
        else:
            leg_text = '0%'
        txBody = legend_label._element.find(f'.//{{{A_NS}}}txBody')
        if txBody is None:
            p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
            txBody = legend_label._element.find(f'.//{{{p_ns}}}txBody')
        if txBody is not None:
            paras = txBody.findall(f'{{{A_NS}}}p')
            if paras:
                _replace_para_text(paras[0], leg_text)


def _sync_slide(target_slide, src_slide, slide_index, total_slides):
    """同步单页内容"""
    # 第 1 页封面：精确替换日期和期号
    if slide_index == 0:
        _update_cover(target_slide, src_slide)
        return

    # 最后一页感谢观看：不做任何替换
    if slide_index == total_slides - 1:
        return

    # 第 13 页（idx=12）频道分类观众规模：图表同步 + 变化标签动态定位
    if slide_index == 12:
        _sync_and_position_audience(target_slide, src_slide)
        return

    # 其他页：按类型匹配并同步
    target_texts = _iter_shapes_by_type(target_slide, 'text')
    src_texts = _iter_shapes_by_type(src_slide, 'text')

    target_charts = _iter_shapes_by_type(target_slide, 'chart')
    src_charts = _iter_shapes_by_type(src_slide, 'chart')

    target_tables = _iter_shapes_by_type(target_slide, 'table')
    src_tables = _iter_shapes_by_type(src_slide, 'table')

    # 文本框：最近邻匹配 + XML 级文本替换
    src_text_map = {idx: shp for idx, shp in src_texts}
    for ti, si in _match_nearest(target_texts, src_texts):
        t_shp = dict(target_texts)[ti]
        s_shp = src_text_map[si]
        _sync_text_shape(t_shp, s_shp)

    # 图表：最近邻匹配 + 数据替换
    src_chart_map = {idx: shp for idx, shp in src_charts}
    for ti, si in _match_nearest(target_charts, src_charts):
        t_shp = dict(target_charts)[ti]
        s_shp = src_chart_map[si]
        _sync_chart(t_shp, s_shp)

    # 表格：最近邻匹配 + 单元格文本替换
    src_table_map = {idx: shp for idx, shp in src_tables}
    for ti, si in _match_nearest(target_tables, src_tables):
        t_shp = dict(target_tables)[ti]
        s_shp = src_table_map[si]
        _sync_table(t_shp, s_shp)


# ═══════════════════════════════════════════════════════════════
# 主入口
# ═══════════════════════════════════════════════════════════════

def _make_data_draft(excel_path, template_path):
    """用现有生成器产出数据稿（临时文件）"""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
    tmp.close()
    base.generate_report(excel_path, template_path, tmp.name)
    return tmp.name


def generate_report_config_driven(excel_path, template_path, output_path):
    """
    配置驱动生成：

    1. 生成数据稿（临时）
    2. 复制 demo → 输出文件
    3. XML 级逐页回填数据
    """
    if not template_path or not Path(template_path).exists():
        raise ValueError(f'模板不存在: {template_path}')

    print(f'[1/4] 生成数据稿: {excel_path}')
    draft_path = _make_data_draft(excel_path, template_path)

    print(f'[2/4] 复制模板: {template_path} → {output_path}')
    shutil.copy2(template_path, output_path)

    print('[3/5] XML 级精确回填...')
    target = Presentation(output_path)
    source = Presentation(draft_path)

    n = min(len(target.slides), len(source.slides))
    for i in range(n):
        print(f'  - 第 {i + 1}/{n} 页')
        _sync_slide(target.slides[i], source.slides[i], i, n)

    # 修复标题（简版模式下 data draft 不生成标题）
    print('[4/5] 修复标题...')
    from data_reader import read_excel_data
    data = read_excel_data(excel_path)
    # 确保日期正确（与 generate_report.py 同逻辑）
    ymd_match = re.search(r'(\d{4})(\d{2})(\d{2})', os.path.basename(excel_path))
    if ymd_match:
        year, month, day = int(ymd_match.group(1)), int(ymd_match.group(2)), int(ymd_match.group(3))
        data.report_date_short = f'{month}月{day}日'
        data.report_date = f'{year}年{month}月{day}日'
    _fix_titles(target, data)

    print(f'[5/5] 保存: {output_path}')
    target.save(output_path)
    print(f'✅ 完成！共 {len(target.slides)} 张幻灯片')


def main():
    parser = argparse.ArgumentParser(description='全页配置驱动 PPT 生成器 v3')
    parser.add_argument('excel', help='Excel 数据文件路径')
    parser.add_argument('--template', '-t', required=True, help='demo 模板 pptx 路径')
    parser.add_argument('--output', '-o', required=True, help='输出 PPT 路径')
    args = parser.parse_args()
    generate_report_config_driven(args.excel, args.template, args.output)


if __name__ == '__main__':
    main()
