#!/usr/bin/env python3
"""
投研分析报告生成器 — touyan2
================================
产线落地/招商项目投研分析报告 Word 文档生成脚本。

用法：
    python3 gen_report.py \\
        --company "公司名称" \\
        --output "/path/to/投研分析报告_公司名称.docx"

本脚本为模板脚本。每次使用时：
  1. 确认分析内容已在对话中完成
  2. 复制本脚本到临时文件
  3. 修改 main() 中 assign_content() 里的内容
  4. 执行生成
"""

import argparse
import os
import sys
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from typing import List, Tuple, Optional


# ═══════════════════════════════════════════════════════════════
#  样式工具函数
# ═══════════════════════════════════════════════════════════════

BLUE_HEADING = RGBColor(0x1A, 0x3C, 0x6E)
BLUE_TABLE_HEADER = "1A3C6E"
DARK_TEXT = RGBColor(0x33, 0x33, 0x33)


def set_cell_shading(cell, color: str):
    """设置表格单元格背景色（hex, 不带 #）"""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_run_font(run, name_cn: str = "宋体", name_en: str = "Times New Roman",
                 size: float = 11, bold: bool = False, color: RGBColor = None,
                 italic: bool = False):
    """统一设置 run 的字体属性"""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    run.font.name = name_en
    r = run._element
    r.rPr.rFonts.set(qn('w:eastAsia'), name_cn)


def add_heading_styled(doc, text: str, level: int = 1):
    """添加带样式的标题（一级蓝色16pt加粗，二级蓝色14pt加粗）"""
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        if level == 1:
            set_run_font(run, size=16, bold=True, color=BLUE_HEADING)
        elif level == 2:
            set_run_font(run, size=14, bold=True, color=BLUE_HEADING)
        elif level == 3:
            set_run_font(run, size=12, bold=True, color=BLUE_HEADING)
    return h


def add_para(doc, text: str, size: float = 11, bold: bool = False,
             color: RGBColor = None, align: str = None,
             space_after: float = 6, italic: bool = False,
             first_line_indent: float = None):
    """添加段落"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color, italic=italic)
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.5
    if align == "center":
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if first_line_indent:
        pf.first_line_indent = Cm(first_line_indent)
    return p


def add_styled_table(doc, headers: List[str], rows: List[List[str]],
                     col_widths: List[float] = None) -> None:
    """添加带样式的表格（表头深蓝底白字）"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Light Grid Accent 1"

    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(header)
        set_run_font(run, size=10, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, BLUE_TABLE_HEADER)

    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(val)
            set_run_font(run, size=10)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)

    doc.add_paragraph()


def add_cover_page(doc, title: str, subtitle: str, company: str, date_str: str = None):
    """添加封面页"""
    if date_str is None:
        date_str = datetime.now().strftime("%Y年%m月%d日")

    for _ in range(6):
        doc.add_paragraph()

    add_para(doc, title, size=22, bold=True, color=BLUE_HEADING, align="center")
    doc.add_paragraph()

    add_para(doc, subtitle, size=14, color=RGBColor(0x66, 0x66, 0x66), align="center")

    for _ in range(6):
        doc.add_paragraph()

    add_para(doc, f"目标企业：{company}", size=12, color=DARK_TEXT, align="center")
    add_para(doc, date_str, size=12, color=DARK_TEXT, align="center")

    doc.add_page_break()


def add_toc_page(doc, chapters: List[str]):
    """添加目录页"""
    add_para(doc, "目  录", size=16, bold=True, color=BLUE_HEADING, align="center", space_after=12)
    for i, ch in enumerate(chapters, 1):
        add_para(doc, f"{'　' * 2}{ch}", size=12, color=DARK_TEXT)
    doc.add_page_break()


def add_disclaimer(doc, company: str):
    """添加免责声明页"""
    add_heading_styled(doc, "免责声明", level=1)
    disclaimer = (
        "本报告由分析团队基于项目方提供的《商业计划书》及相关公开信息编制，仅供内部决策参考，"
        "不构成投资建议。报告中的分析和判断基于项目方提供的数据和行业公开信息，分析团队不对"
        f"项目方提供数据的真实性、准确性、完整性做任何明示或暗示的保证。{company}及其关联方"
        "不对因使用本报告而导致的任何损失承担任何责任。"
    )
    add_para(doc, disclaimer, size=10.5, color=RGBColor(0x66, 0x66, 0x66))
    doc.add_page_break()


# ═══════════════════════════════════════════════════════════════
#  内容定义
# ═══════════════════════════════════════════════════════════════

class ReportContent:
    """报告内容容器。每次使用时按需填充"""

    def __init__(self, company: str):
        self.company = company

        # ---- 章节内容 ----
        self.chapter1_title = "项目概览"
        self.ch1_basic_info_headers = ["维度", "内容"]
        self.ch1_basic_info_rows: List[List[str]] = []
        self.ch1_investment_headers = ["阶段", "时间", "投资金额", "主要用途"]
        self.ch1_investment_rows: List[List[str]] = []
        self.ch1_usage_headers = ["用途类别", "金额", "占比"]
        self.ch1_usage_rows: List[List[str]] = []
        self.ch1_paragraphs: List[str] = []

        self.chapter2_title = "行业与市场分析"
        self.ch2_paragraphs: List[str] = []
        self.ch2_market_headers = ["指标", "当前数据", "预测数据"]
        self.ch2_market_rows: List[List[str]] = []
        self.ch2_comp_headers = ["企业", "市场份额", "定位", "特点"]
        self.ch2_comp_rows: List[List[str]] = []

        self.chapter3_title = "项目竞争力评估"
        self.ch3_swot_headers = ["评估维度", "评估结果"]
        self.ch3_swot_rows: List[List[str]] = []
        self.ch3_paragraphs: List[str] = []

        self.chapter4_title = "财务效益分析"
        self.ch4_return_headers = ["指标", "一期", "二期", "三期", "整体"]
        self.ch4_return_rows: List[List[str]] = []
        self.ch4_paragraphs: List[str] = []

        self.chapter5_title = "地方落地价值评估"
        self.ch5_value_headers = ["评估维度", "评估意见"]
        self.ch5_value_rows: List[List[str]] = []
        self.ch5_paragraphs: List[str] = []

        self.chapter6_title = "风险分析与缓释"
        self.ch6_risk_headers = ["风险类别", "风险描述", "等级", "缓释措施"]
        self.ch6_risk_rows: List[List[str]] = []
        self.ch6_paragraphs: List[str] = []

        self.chapter7_title = "综合评估与建议"
        self.ch7_score_headers = ["评估维度", "评分", "评语"]
        self.ch7_score_rows: List[List[str]] = []
        self.ch7_overall_verdict = ""
        self.ch7_stage_headers = ["阶段", "建议", "前提条件"]
        self.ch7_stage_rows: List[List[str]] = []
        self.ch7_dd_headers = ["尽调重点", "说明"]
        self.ch7_dd_rows: List[List[str]] = []
        self.ch7_paragraphs: List[str] = []

        self.date_str = datetime.now().strftime("%Y年%m月%d日")


# ═══════════════════════════════════════════════════════════════
#  文档生成
# ═══════════════════════════════════════════════════════════════

def build_chapter(doc, content: ReportContent,
                  chapter_num: int, title: str,
                  paragraphs: List[str],
                  table1: Tuple[str, List[str], List[List[str]]] = None,
                  table2: Tuple[str, List[str], List[List[str]]] = None,
                  table3: Tuple[str, List[str], List[List[str]]] = None):
    """
    通用章节构建函数
    paragraphs: 正文段落列表（每段 str）
    tableX: (table_title, [headers], [rows]) 可选
    """
    add_heading_styled(doc, f"第{chapter_num}章 {title}", level=1)

    for para_text in paragraphs:
        add_para(doc, para_text, first_line_indent=0.74)

    for tbl in [table1, table2, table3]:
        if tbl:
            tbl_title, headers, rows = tbl
            if tbl_title:
                add_para(doc, tbl_title, size=11, bold=True, space_after=4)
            if rows:
                add_styled_table(doc, headers, rows)


def generate_report(content: ReportContent, output_path: str):
    """生成完整投研报告"""
    doc = Document()

    # 设置默认字体
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Times New Roman"
    style.element.rPr.rFonts.set(qn('w:eastAsia'), "宋体")
    style.font.size = Pt(11)
    style.paragraph_format.line_spacing = 1.5

    # ── 封面 ──
    add_cover_page(
        doc,
        title="投研分析报告",
        subtitle="——产线落地项目投资价值评估——",
        company=content.company,
        date_str=content.date_str,
    )

    # ── 目录 ──
    toc = [
        f"一、{content.chapter1_title}",
        f"二、{content.chapter2_title}",
        f"三、{content.chapter3_title}",
        f"四、{content.chapter4_title}",
        f"五、{content.chapter5_title}",
        f"六、{content.chapter6_title}",
        f"七、{content.chapter7_title}",
    ]
    add_toc_page(doc, toc)

    # ── 免责声明 ──
    add_disclaimer(doc, content.company)

    # ── 第一章 项目概览 ──
    add_heading_styled(doc, f"一、{content.chapter1_title}", level=1)

    for para_text in content.ch1_paragraphs:
        add_para(doc, para_text, first_line_indent=0.74)

    if content.ch1_basic_info_rows:
        add_para(doc, "▎基本信息", size=11, bold=True, space_after=4)
        add_styled_table(doc, content.ch1_basic_info_headers, content.ch1_basic_info_rows)

    if content.ch1_investment_rows:
        add_para(doc, "▎投资节奏与结构", size=11, bold=True, space_after=4)
        add_styled_table(doc, content.ch1_investment_headers, content.ch1_investment_rows)

    if content.ch1_usage_rows:
        add_para(doc, "▎资金用途明细", size=11, bold=True, space_after=4)
        add_styled_table(doc, content.ch1_usage_headers, content.ch1_usage_rows)

    # ── 第二章 行业与市场分析 ──
    add_heading_styled(doc, f"二、{content.chapter2_title}", level=1)
    for para_text in content.ch2_paragraphs:
        add_para(doc, para_text, first_line_indent=0.74)
    if content.ch2_market_rows:
        add_para(doc, "▎市场规模", size=11, bold=True, space_after=4)
        add_styled_table(doc, content.ch2_market_headers, content.ch2_market_rows)
    if content.ch2_comp_rows:
        add_para(doc, "▎竞争格局", size=11, bold=True, space_after=4)
        add_styled_table(doc, content.ch2_comp_headers, content.ch2_comp_rows)

    # ── 第三章 项目竞争力评估 ──
    add_heading_styled(doc, f"三、{content.chapter3_title}", level=1)
    if content.ch3_swot_rows:
        add_para(doc, "▎优势/劣势矩阵", size=11, bold=True, space_after=4)
        add_styled_table(doc, content.ch3_swot_headers, content.ch3_swot_rows)
    for para_text in content.ch3_paragraphs:
        add_para(doc, para_text, first_line_indent=0.74)

    # ── 第四章 财务效益分析 ──
    add_heading_styled(doc, f"四、{content.chapter4_title}", level=1)
    if content.ch4_return_rows:
        add_para(doc, "▎投资回收分析", size=11, bold=True, space_after=4)
        add_styled_table(doc, content.ch4_return_headers, content.ch4_return_rows)
    for para_text in content.ch4_paragraphs:
        add_para(doc, para_text, first_line_indent=0.74)

    # ── 第五章 地方落地价值评估 ──
    add_heading_styled(doc, f"五、{content.chapter5_title}", level=1)
    for para_text in content.ch5_paragraphs:
        add_para(doc, para_text, first_line_indent=0.74)
    if content.ch5_value_rows:
        add_para(doc, "▎落地价值评估", size=11, bold=True, space_after=4)
        add_styled_table(doc, content.ch5_value_headers, content.ch5_value_rows)

    # ── 第六章 风险分析与缓释 ──
    add_heading_styled(doc, f"六、{content.chapter6_title}", level=1)
    for para_text in content.ch6_paragraphs:
        add_para(doc, para_text, first_line_indent=0.74)
    if content.ch6_risk_rows:
        add_para(doc, "▎风险矩阵", size=11, bold=True, space_after=4)
        add_styled_table(doc, content.ch6_risk_headers, content.ch6_risk_rows)

    # ── 第七章 综合评估与建议 ──
    add_heading_styled(doc, f"七、{content.chapter7_title}", level=1)
    if content.ch7_score_rows:
        add_para(doc, "▎综合评价", size=11, bold=True, space_after=4)
        add_styled_table(doc, content.ch7_score_headers, content.ch7_score_rows)

    for para_text in content.ch7_paragraphs:
        add_para(doc, para_text, first_line_indent=0.74)

    if content.ch7_stage_rows:
        add_para(doc, "▎分期投资建议", size=11, bold=True, space_after=4)
        add_styled_table(doc, content.ch7_stage_headers, content.ch7_stage_rows)

    if content.ch7_dd_rows:
        add_para(doc, "▎尽调重点", size=11, bold=True, space_after=4)
        add_styled_table(doc, content.ch7_dd_headers, content.ch7_dd_rows)

    # ── 保存 ──
    doc.save(output_path)
    return output_path


# ═══════════════════════════════════════════════════════════════
#  主入口 — 每次使用时修改此处填充内容
# ═══════════════════════════════════════════════════════════════

def assign_content(content: ReportContent) -> None:
    """
    填充报告内容。
    每次使用本脚本时，将对话中完成的分析内容填写到此处。
    ====== 使用范例（伪代码）======
    以下为完整填充示例。每次生成新报告时替换为实际分析结果。
    保持结构不变，替换内容即可。
    """

    # ════════════════════ 第一章 ════════════════════

    content.ch1_paragraphs = [
        # 第一段：总体描述
        f"{content.company}本次拟在[地点]投资建设[产品]产线，项目总投资额约[X]亿元，分[X]期实施。"
        "本报告基于项目方提供的商业计划书及相关资料，从项目基本面、行业前景、竞争力、"
        "财务效益、地方落地价值及风险等维度进行综合评估。",

        # 第二段：建设内容概述
        "项目主要建设内容包括[具体建设内容]，规划年产能[产能]亿元。项目达产后预计年产值[X]亿元，"
        "年税收贡献约[X]亿元，带动就业约[X]人。",
    ]

    # 基本信息表
    content.ch1_basic_info_rows = [
        ["项目名称", f"{content.company}产线落地项目"],
        ["实施主体", content.company],
        ["投资总额", "[X]亿元"],
        ["建设内容", "[产品方向]产线建设及配套设施"],
        ["厂房需求", "[面积]平方米，[洁净度等级等要求]"],
        ["建设周期", "一期[X]个月，总周期[X]年"],
        ["带动就业", "[X]人（含直接+间接）"],
    ]

    # 投资节奏
    content.ch1_investment_rows = [
        ["一期", "第1年", "[X]亿元", "厂房建设、设备采购、流资"],
        ["二期", "第2年", "[X]亿元", "产能扩建、技术升级"],
        ["三期", "第3年", "[X]亿元", "全面达产、市场拓展"],
    ]

    # 资金用途
    content.ch1_usage_rows = [
        ["设备购置", "[X]亿元", "[X]%"],
        ["厂房建设", "[X]亿元", "[X]%"],
        ["流动资金", "[X]亿元", "[X]%"],
        ["其他", "[X]亿元", "[X]%"],
    ]

    # ════════════════════ 第二章 ════════════════════

    content.ch2_paragraphs = [
        # 宏观背景
        "[描述行业宏观背景、与政策热点的关联。1-2段]",

        # 市场分析
        f"根据第三方数据，{content.company}所在的目标市场当前规模约[X]亿元，"
        "预计未来3年保持[X]%的年复合增长。主要增长驱动包括：...",

        # 竞争格局
        "目前国内主要参与者包括[X]、[Y]、[Z]等。市场集中度[高/中/低]，CR5约[X]%。"
        f"{content.company}定位于[具体市场定位]，差异化主要体现在[差异化特点]。",
    ]

    content.ch2_market_rows = [
        ["全球市场规模", "[X]亿元", "[X]亿元（[X]年E）"],
        ["国内市场规模", "[X]亿元", "[X]亿元（[X]年E）"],
        ["年复合增速", "—", "[X]%"],
    ]

    content.ch2_comp_rows = [
        ["[头部企业A]", "[X]%", "高端市场", "[技术/客户/渠道特点]"],
        ["[头部企业B]", "[X]%", "中高端市场", "[特点]"],
        ["[头部企业C]", "[X]%", "中低端市场", "[特点]"],
        [content.company, "[X]%", "[定位]", "[差异化]"],
    ]

    # ════════════════════ 第三章 ════════════════════

    content.ch3_swot_rows = [
        ["团队背景", "[核心技术团队履历摘要]"],
        ["技术能力", "[自研/代工、核心专利、良率水平]"],
        ["产品力", "[产品性能指标、认证进展]"],
        ["客户基础", "[已有客户/在手订单/意向客户]"],
        ["产能规模", "[规划产能 vs 行业比较]"],
        ["成本优势", "[原材料/制造成本/人工方面]"],
        ["品牌/渠道", "[行业知名度、销售渠道建设]"],
    ]

    content.ch3_paragraphs = [
        "综合分析，该项目[核心竞争力所在的1-2个维度]表现较强，但在[薄弱维度]方面需进一步验证。"
        "最大变量在于[关键不确定性因素]。",

        "差异化策略方面，项目方计划[差异化路径描述]，若[头部厂商/竞争对手]跟进该策略，"
        "差异化效果可能有所削弱。",

        f"目前{content.company}已取得[认证/客户/专利等具体进展]，但整体仍处于[阶段]，"
        "距离大规模量产仍需[时间/资金/验证]。",
    ]

    # ════════════════════ 第四章 ════════════════════

    content.ch4_return_rows = [
        ["投资额", "[X]亿元", "[X]亿元", "[X]亿元", "[X]亿元"],
        ["达产年产值", "[X]亿元", "[X]亿元", "[X]亿元", "[X]亿元"],
        ["达产年净利润", "[X]亿元", "[X]亿元", "[X]亿元", "[X]亿元"],
        ["投资回报率", "[X]%", "[X]%", "[X]%", "[X]%"],
        ["静态回收期", "[X]年", "[X]年", "[X]年", "[X]年"],
    ]

    content.ch4_paragraphs = [
        "上述测算基于项目方提供的财务模型，核心假设包括：",

        "① 毛利率假设：达产年毛利率[X]%，需与行业平均[X]%对比确认合理性。"
        "② 良率爬坡：第一年良率[X]% → 第二年[X]% → 稳态[X]%（行业正常爬坡速度为...）。"
        "③ 费用率假设：销售费用率[X]%、管理费用率[X]%、研发费用率[X]%。"
        "④ 产能利用率：第一年[X]% → 第二年[X]% → 第三年起[X]%。",

        "敏感性分析：销售收入下降10%时，整体IRR从[X]%降至[X]%，回收期延长至[X]年；"
        "毛利率下降5个百分点时，整体IRR降至[X]%。整体来看，[财务评估结论]。",
    ]

    # ════════════════════ 第五章 ════════════════════

    content.ch5_paragraphs = [
        "从地方政府视角来看，本项目对当地经济的综合带动效益需要从多个维度评估。"
        f"项目方计划落地位于[地点]，需要政府提供的支持包括[主要需求清单]。",

        f"这些需求在[地方]现有的招商政策框架内[部分/全部]可以满足，"
        "超出常规的部分包括[特殊条件清单]，需在决策前明确政策依据。",

        "政府端风险管理的核心在于[出资/监管/回购等关键安排]。",
    ]

    content.ch5_value_rows = [
        ["产业契合度", "[是否匹配地方产业规划，一句话]"],
        ["投资密度", "[亩均投资强度，对比地方标准]"],
        ["产值贡献", "达产后年产值[X]亿元，年税收[X]亿元"],
        ["就业带动", "直接就业[X]人，间接就业[X]人"],
        ["产业带动", "[能否吸引上下游形成集群]"],
        ["招商条件", "[反投比例/总部搬迁等条件]"],
        ["品牌效应", "[对地方产业名片的价值]"],
    ]

    # ════════════════════ 第六章 ════════════════════

    content.ch6_paragraphs = [
        "经综合评估，本项目面临的主要风险集中在以下方面，各风险的等级判断及对应缓释措施如下：",
    ]

    content.ch6_risk_rows = [
        ["技术风险", "[具体风险描述]", "高/中/低", "[缓释措施]"],
        ["市场风险", "[具体风险描述]", "高/中/低", "[缓释措施]"],
        ["资金风险", "[具体风险描述]", "高/中/低", "[缓释措施]"],
        ["团队风险", "[具体风险描述]", "高/中/低", "[缓释措施]"],
        ["建设风险", "[具体风险描述]", "高/中/低", "[缓释措施]"],
        ["政策风险", "[具体风险描述]", "高/中/低", "[缓释措施]"],
        ["竞争风险", "[具体风险描述]", "高/中/低", "[缓释措施]"],
    ]

    # ════════════════════ 第七章 ════════════════════

    content.ch7_score_rows = [
        ["赛道前景", "★★★★☆", "[评语]"],
        ["项目本身质量", "★★★★☆", "[评语]"],
        ["投资条件", "★★★☆☆", "[评语]"],
        ["产值/税收贡献", "★★★☆☆", "[评语]"],
        ["风险可控度", "★★★☆☆", "[评语]"],
    ]

    content.ch7_paragraphs = [
        "总体判断：[推荐落地 / 审慎推进 / 不推荐]。",

        f"综合来看，{content.company}在[核心优势维度]方面具备一定竞争力，"
        '但[主要短板/风险]需要在投资决策前进一步核实。建议采取「分期投资、条件触发」的方式推进。',
    ]

    content.ch7_stage_rows = [
        ["一期（X万元）", "建议投资 / 审慎 / 暂缓", "① [条件1] ② [条件2]"],
        ["二期（X万元）", "条件触发 / 暂不考虑", "① 一期达到预期指标 ② [其他条件]"],
        ["三期（X万元）", "机会导向 / 暂不考虑", "① 市场验证通过 ② [其他条件]"],
    ]

    content.ch7_dd_rows = [
        ["[尽调重点1]", "[说明]"],
        ["[尽调重点2]", "[说明]"],
        ["[尽调重点3]", "[说明]"],
    ]


# ═══════════════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════════════

def _parse_args():
    parser = argparse.ArgumentParser(
        description="投研分析报告生成器（touyan2）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""使用步骤：
  1. 确认分析内容已在对话中完成
  2. 修改 assign_content() 中的内容
  3. 运行 python3 gen_report.py --company "公司名" --output /path/to/result.docx
"""
    )
    parser.add_argument("--company", required=True, help="目标公司名称")
    parser.add_argument("--output", required=True, help="输出 .docx 文件路径")
    return parser.parse_args()


def main():
    args = _parse_args()

    print(f"=== touyan2 投研分析报告生成器 ===")
    print(f"目标公司: {args.company}")
    print(f"输出路径: {args.output}")
    print()

    content = ReportContent(args.company)
    assign_content(content)

    out = generate_report(content, args.output)
    print(f"✅ 报告已生成: {out}")
    print(f"   文件大小: {os.path.getsize(out) / 1024:.1f} KB")


if __name__ == "__main__":
    main()
