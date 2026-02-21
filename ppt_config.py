# -*- coding: utf-8 -*-
"""
PPT报告生成器 - 主题配置
定义颜色方案、字体、布局参数等
"""
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


# ──────────── 幻灯片尺寸 ────────────
SLIDE_WIDTH = Emu(12192000)   # 13.33 inches
SLIDE_HEIGHT = Emu(6858000)   # 7.50 inches

# ──────────── 颜色方案 ────────────
class Colors:
    # 主色调
    PRIMARY = RGBColor(0x1A, 0x3C, 0x5E)      # 深蓝
    SECONDARY = RGBColor(0x42, 0x24, 0x19)     # 深棕(origin模板用色)
    ACCENT_BLUE = RGBColor(0x2B, 0x7A, 0xB8)   # 蓝色强调
    ACCENT_RED = RGBColor(0xE7, 0x4C, 0x3C)    # 红色(下降)
    ACCENT_GREEN = RGBColor(0x27, 0xAE, 0x60)  # 绿色(上升)
    ACCENT_ORANGE = RGBColor(0xF3, 0x9C, 0x12) # 橙色
    ACCENT_PURPLE = RGBColor(0x8E, 0x44, 0xAD) # 紫色

    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    BLACK = RGBColor(0x00, 0x00, 0x00)
    DARK_GRAY = RGBColor(0x33, 0x33, 0x33)
    MEDIUM_GRAY = RGBColor(0x66, 0x66, 0x66)
    LIGHT_GRAY = RGBColor(0xCC, 0xCC, 0xCC)
    BG_LIGHT = RGBColor(0xF5, 0xF5, 0xF5)

    # 图表系列配色
    CHART_SERIES = [
        RGBColor(0x2B, 0x7A, 0xB8),  # 蓝
        RGBColor(0xE7, 0x4C, 0x3C),  # 红
        RGBColor(0x27, 0xAE, 0x60),  # 绿
        RGBColor(0xF3, 0x9C, 0x12),  # 橙
        RGBColor(0x8E, 0x44, 0xAD),  # 紫
        RGBColor(0x1A, 0xBC, 0x9C),  # 青
        RGBColor(0xE6, 0x7E, 0x22),  # 深橙
        RGBColor(0x34, 0x98, 0xDB),  # 亮蓝
    ]

    # 表格配色
    TABLE_HEADER_BG = RGBColor(0x2B, 0x7A, 0xB8)
    TABLE_HEADER_FG = RGBColor(0xFF, 0xFF, 0xFF)
    TABLE_ROW_EVEN = RGBColor(0xEA, 0xF2, 0xFA)
    TABLE_ROW_ODD  = RGBColor(0xFF, 0xFF, 0xFF)
    TABLE_HIGHLIGHT = RGBColor(0xFF, 0xF3, 0xCD)  # 高亮行(CCTV-17)

    # 封面配色
    COVER_BG = RGBColor(0x1A, 0x3C, 0x5E)
    COVER_ACCENT = RGBColor(0x2B, 0x7A, 0xB8)
    COVER_TEXT = RGBColor(0xFF, 0xFF, 0xFF)

# ──────────── 字体设置 ────────────
class Fonts:
    MAIN = '微软雅黑'
    TITLE_SIZE = Pt(28)
    SUBTITLE_SIZE = Pt(18)
    BODY_SIZE = Pt(12)
    SMALL_SIZE = Pt(10)
    TINY_SIZE = Pt(8)
    TABLE_HEADER_SIZE = Pt(11)
    TABLE_BODY_SIZE = Pt(10)
    CHART_LABEL_SIZE = Pt(9)
    COVER_TITLE_SIZE = Pt(48)
    COVER_SUBTITLE_SIZE = Pt(24)
    COVER_DATE_SIZE = Pt(18)
    PAGE_NUM_SIZE = Pt(14)

# ──────────── 布局参数 ────────────
class Layout:
    # 页面边距
    MARGIN_LEFT = Cm(1.5)
    MARGIN_RIGHT = Cm(1.5)
    MARGIN_TOP = Cm(1.5)
    MARGIN_BOTTOM = Cm(1.0)

    # 标题区域
    TITLE_LEFT = Cm(1.5)
    TITLE_TOP = Cm(0.3)
    TITLE_WIDTH = Cm(30)
    TITLE_HEIGHT = Cm(1.8)

    # 副标题/描述区域
    DESC_LEFT = Cm(1.5)
    DESC_TOP = Cm(1.8)
    DESC_WIDTH = Cm(30)
    DESC_HEIGHT = Cm(1.2)

    # 内容区域 (标题下方)
    CONTENT_LEFT = Cm(1.5)
    CONTENT_TOP = Cm(3.0)
    CONTENT_WIDTH = Cm(30.8)
    CONTENT_HEIGHT = Cm(15.0)

    # 页码位置
    PAGE_NUM_LEFT = Cm(31.5)
    PAGE_NUM_TOP = Cm(17.5)
    PAGE_NUM_WIDTH = Cm(1.5)
    PAGE_NUM_HEIGHT = Cm(0.8)

    # 注释文本位置
    NOTE_LEFT = Cm(1.5)
    NOTE_BOTTOM = Cm(17.8)
    NOTE_WIDTH = Cm(25)
    NOTE_HEIGHT = Cm(0.6)


# ──────────── CCTV-17标识 ────────────
CHANNEL_NAME = 'CCTV-17农业农村'
CHANNEL_FULL_NAME = '中央电视台农业农村频道'
DEPARTMENT = '农业农村节目中心\n统筹策划部'
