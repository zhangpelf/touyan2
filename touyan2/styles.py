"""
touyan2.styles
==============
Word 文档公文级排版与样式工具函数。
提供政企商务蓝标准色彩体系、字体中英文双轨设置、单元格底纹、自适应边框等。
"""

from typing import List, Optional
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from datetime import datetime


# ═══════════════════════════════════════════════════════════════
#  公文与商务投资级标准色盘
# ═══════════════════════════════════════════════════════════════

BLUE_HEADING = RGBColor(0x1A, 0x3C, 0x6E)       # 一/二级标题：商务深青蓝
BLUE_TABLE_HEADER = "1A3C6E"                    # 表头底纹 HEX
BLUE_ACCENT = RGBColor(0x2B, 0x5B, 0x99)        # 强调文字/三级标题蓝
DARK_TEXT = RGBColor(0x33, 0x33, 0x33)          # 正文主色：深灰黑（比纯黑更护眼）
MUTED_TEXT = RGBColor(0x66, 0x66, 0x66)         # 辅助文本/注释/免责声明
ALERT_RED = RGBColor(0xC0, 0x00, 0x00)          # 高风险/红旗警示色
SUCCESS_GREEN = RGBColor(0x1E, 0x7E, 0x34)      # 优势/合规健康色
LIGHT_ROW_SHD = "F7F9FC"                        # 表格交替行浅色底纹


def set_cell_shading(cell, color_hex: str):
    """设置表格单元格背景色（hex, 不带 #）"""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_table_borders(table, border_color: str = "D3D8E0"):
    """为表格添加专业规整细边框"""
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
        table._tbl.append(tblPr)
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="{border_color}"/>'
        f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>'
        f'  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def set_run_font(
    run,
    name_cn: str = "宋体",
    name_en: str = "Times New Roman",
    size: float = 11,
    bold: bool = False,
    color: Optional[RGBColor] = None,
    italic: bool = False,
):
    """统一设置 run 的字体属性（中西文双轨）"""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    run.font.name = name_en
    r = run._element
    r.rPr.rFonts.set(qn("w:eastAsia"), name_cn)


def add_heading_styled(doc, text: str, level: int = 1):
    """添加带公文样式的标题（一级深蓝16pt加粗，二级深蓝14pt加粗，三级12pt加粗）"""
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        if level == 1:
            set_run_font(run, name_cn="黑体", size=16, bold=True, color=BLUE_HEADING)
        elif level == 2:
            set_run_font(run, name_cn="黑体", size=14, bold=True, color=BLUE_HEADING)
        elif level == 3:
            set_run_font(run, name_cn="黑体", size=12, bold=True, color=BLUE_ACCENT)
    return h


def add_para(
    doc,
    text: str,
    size: float = 11,
    bold: bool = False,
    color: Optional[RGBColor] = None,
    align: Optional[str] = None,
    space_after: float = 6,
    space_before: float = 0,
    italic: bool = False,
    first_line_indent: Optional[float] = None,
    line_spacing: float = 1.5,
):
    """添加排版段落"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color or DARK_TEXT, italic=italic)
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = line_spacing
    if align == "center":
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "justify":
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if first_line_indent:
        pf.first_line_indent = Cm(first_line_indent)
    return p


def add_styled_table(
    doc,
    headers: List[str],
    rows: List[List[str]],
    col_widths: Optional[List[float]] = None,
    alternate_row_shd: bool = True,
) -> None:
    """添加带样式的三线/商务表格（表头深蓝底白字、交替浅色底纹）"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)

    # 表头处理
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(header)
        set_run_font(run, size=10, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, BLUE_TABLE_HEADER)

    # 数据行处理
    for r_idx, row in enumerate(rows):
        is_even = (r_idx % 2 == 1)
        for c_idx, val in enumerate(row):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            val_str = str(val)
            run = p.add_run(val_str)
            set_run_font(run, size=9.5, color=DARK_TEXT)
            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            if alternate_row_shd and is_even:
                set_cell_shading(cell, LIGHT_ROW_SHD)

    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                if i < len(row.cells):
                    row.cells[i].width = Cm(w)

    doc.add_paragraph()


def add_cover_page(
    doc,
    title: str = "投研分析报告",
    subtitle: str = "——产线落地项目投资价值评估——",
    company: str = "目标企业",
    date_str: Optional[str] = None,
):
    """添加标准报告封面"""
    if date_str is None:
        date_str = datetime.now().strftime("%Y年%m月%d日")

    for _ in range(5):
        doc.add_paragraph()

    add_para(doc, title, size=24, bold=True, color=BLUE_HEADING, align="center")
    doc.add_paragraph()
    add_para(doc, subtitle, size=14, color=MUTED_TEXT, align="center")

    for _ in range(6):
        doc.add_paragraph()

    add_para(doc, f"目标企业：{company}", size=12, bold=True, color=DARK_TEXT, align="center")
    add_para(doc, f"评估机构：产业投研与国资招商评估小组", size=11, color=MUTED_TEXT, align="center")
    add_para(doc, f"报告日期：{date_str}", size=11, color=MUTED_TEXT, align="center")
    doc.add_page_break()


def add_toc_page(doc, chapters: List[str]):
    """添加目录页"""
    add_para(doc, "目  录", size=16, bold=True, color=BLUE_HEADING, align="center", space_after=14)
    for ch in chapters:
        add_para(doc, f"{'　' * 2}{ch}", size=11.5, color=DARK_TEXT, space_after=6)
    doc.add_page_break()


def add_disclaimer(doc, company: str):
    """添加免责声明页"""
    add_heading_styled(doc, "免责声明与评估说明", level=1)
    paras = [
        "一、本报告由评估团队基于项目方提供的商业计划书、交流材料以及相关公开数据编制，旨在为地方政府招商引资决策、国资产业基金投资研判提供独立第三方框架参考，不构成法定审计结论或投资担保。",
        f"二、报告中关于【{company}】的经营历史、产能规划、财务预测及客户承诺等内容，主要源自项目方单方陈述。在正式签署招商落地协议或出资协议前，务必以法律、财务、技术尽职调查底稿及第三方权威复核结论为准。",
        "三、地方招商涉及财政资金出资、用地指标划拨、厂房代建及税收奖补等公共资源配置，建议严格遵守“分期投入、里程碑考核、对赌回收”原则，动态监控履约风险。"
    ]
    for p_block in paras:
        add_para(doc, p_block, size=10.5, color=MUTED_TEXT, space_after=8, first_line_indent=0.74)
    doc.add_page_break()
