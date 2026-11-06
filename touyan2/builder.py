"""
touyan2.builder
===============
投研分析报告 Word 文档装配引擎。
负责从 ReportContent 数据模型组装高质量排版且规范严谨的 .docx 文档。
"""

import os
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn

from .models import ReportContent
from .styles import (
    add_cover_page,
    add_toc_page,
    add_disclaimer,
    add_heading_styled,
    add_para,
    add_styled_table,
    BLUE_HEADING,
    BLUE_ACCENT,
    DARK_TEXT,
    MUTED_TEXT,
)


def build_chapter(
    doc,
    chapter_num_str: str,
    title: str,
    paragraphs: list,
    tables: list = None,
):
    """
    通用章节构建器
    chapter_num_str: '一'、'二' 等
    tables: 列表，每个元素为 (table_title, headers, rows)
    """
    add_heading_styled(doc, f"{chapter_num_str}、{title}", level=1)

    for p in paragraphs:
        if isinstance(p, str) and p.strip():
            add_para(doc, p.strip(), first_line_indent=0.74)

    if tables:
        for tbl in tables:
            if not tbl:
                continue
            tbl_title, headers, rows = tbl
            if tbl_title:
                add_para(doc, f"▎{tbl_title}", size=11, bold=True, color=BLUE_ACCENT, space_after=4, space_before=6)
            if headers and rows:
                add_styled_table(doc, headers, rows)


class ReportGenerator:
    """投研分析报告文档生成器"""

    def __init__(self, content: ReportContent):
        self.content = content

    def generate(self, output_path: str) -> str:
        """生成完整报告 .docx 文件"""
        doc = Document()

        # 页面边距与基础段落样式
        for section in doc.sections:
            section.top_margin = Inches(1.0)
            section.bottom_margin = Inches(1.0)
            section.left_margin = Inches(1.0)
            section.right_margin = Inches(1.0)

        normal_style = doc.styles["Normal"]
        normal_style.font.name = "Times New Roman"
        normal_style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        normal_style.font.size = Pt(11)
        normal_style.paragraph_format.line_spacing = 1.5

        # 1. 封面
        add_cover_page(
            doc,
            title=self.content.report_title,
            subtitle=self.content.subtitle,
            company=self.content.company,
            date_str=self.content.date_str,
        )

        # 2. 投资亮点前置页（若有定义）
        if getattr(self.content, "ch1_highlights", None) and len(self.content.ch1_highlights) > 0:
            add_heading_styled(doc, "【核心投资亮点前置】", level=1)
            add_para(
                doc,
                "以下核心投资亮点与业务锚点提炼自项目方资料及产业独立交叉比对，作为决策层快速研判之依据：",
                size=11,
                color=MUTED_TEXT,
                space_after=10,
            )
            for hl in self.content.ch1_highlights:
                add_para(doc, f"●  {hl}", size=11, color=DARK_TEXT, space_after=6, space_before=2)
            doc.add_page_break()

        # 3. 目录页
        toc = [
            f"一、{self.content.chapter1_title}",
            f"二、{self.content.chapter2_title}",
            f"三、{self.content.chapter3_title}",
            f"四、{self.content.chapter4_title}",
            f"五、{self.content.chapter5_title}",
            f"六、{self.content.chapter6_title}",
            f"七、{self.content.chapter7_title}",
        ]
        add_toc_page(doc, toc)

        # 4. 免责声明页
        add_disclaimer(doc, self.content.company)

        # 5. 第一章 项目概览
        ch1_tables = []
        if self.content.ch1_basic_info_rows:
            ch1_tables.append(("基本信息表", self.content.ch1_basic_info_headers, self.content.ch1_basic_info_rows))
        if self.content.ch1_investment_rows:
            ch1_tables.append(("投资节奏与规划结构", self.content.ch1_investment_headers, self.content.ch1_investment_rows))
        if self.content.ch1_usage_rows:
            ch1_tables.append(("募集/落地资金用途明细", self.content.ch1_usage_headers, self.content.ch1_usage_rows))
        build_chapter(doc, "一", self.content.chapter1_title, self.content.ch1_paragraphs, ch1_tables)

        # 6. 第二章 行业与市场分析
        ch2_tables = []
        if self.content.ch2_market_rows:
            ch2_tables.append(("赛道市场空间与增长预测", self.content.ch2_market_headers, self.content.ch2_market_rows))
        if self.content.ch2_comp_rows:
            ch2_tables.append(("产业链主要竞品与市场竞争格局", self.content.ch2_comp_headers, self.content.ch2_comp_rows))
        build_chapter(doc, "二", self.content.chapter2_title, self.content.ch2_paragraphs, ch2_tables)

        # 7. 第三章 项目竞争力评估
        ch3_tables = []
        if self.content.ch3_swot_rows:
            ch3_tables.append(("项目核心优劣势对比矩阵", self.content.ch3_swot_headers, self.content.ch3_swot_rows))
        if getattr(self.content, "ch3_product_rows", None) and self.content.ch3_product_rows:
            ch3_tables.append(("核心技术与产品关键参数对标", self.content.ch3_product_headers, self.content.ch3_product_rows))
        build_chapter(doc, "三", self.content.chapter3_title, self.content.ch3_paragraphs, ch3_tables)

        # 8. 第四章 财务效益分析
        ch4_tables = []
        if self.content.ch4_return_rows:
            ch4_tables.append(("分期产值、效益与投资回报测算", self.content.ch4_return_headers, self.content.ch4_return_rows))
        if getattr(self.content, "ch4_sens_rows", None) and self.content.ch4_sens_rows:
            ch4_tables.append(("核心变量敏感性分析与抗风险测算", self.content.ch4_sens_headers, self.content.ch4_sens_rows))
        build_chapter(doc, "四", self.content.chapter4_title, self.content.ch4_paragraphs, ch4_tables)

        # 9. 第五章 地方落地价值评估
        ch5_tables = []
        if self.content.ch5_value_rows:
            ch5_tables.append(("地方产业政策匹配度与经济贡献考核", self.content.ch5_value_headers, self.content.ch5_value_rows))
        if getattr(self.content, "ch5_policy_rows", None) and self.content.ch5_policy_rows:
            ch5_tables.append(("项目方政策要素诉求与谈判应对方案", self.content.ch5_policy_headers, self.content.ch5_policy_rows))
        build_chapter(doc, "五", self.content.chapter5_title, self.content.ch5_paragraphs, ch5_tables)

        # 10. 第六章 风险分析与缓释
        ch6_tables = []
        if self.content.ch6_risk_rows:
            ch6_tables.append(("全生命周期关键风险清单与缓释措施", self.content.ch6_risk_headers, self.content.ch6_risk_rows))
        build_chapter(doc, "六", self.content.chapter6_title, self.content.ch6_paragraphs, ch6_tables)

        # 11. 第七章 综合评估与建议
        ch7_tables = []
        if self.content.ch7_score_rows:
            ch7_tables.append(("招商落地项目综合研判评分卡", self.content.ch7_score_headers, self.content.ch7_score_rows))
        if self.content.ch7_stage_rows:
            ch7_tables.append(("分期投资决策与里程碑对赌考核条件", self.content.ch7_stage_headers, self.content.ch7_stage_rows))
        if self.content.ch7_dd_rows:
            ch7_tables.append(("投前尽调核心排查清单（Red Flags）", self.content.ch7_dd_headers, self.content.ch7_dd_rows))
        build_chapter(doc, "七", self.content.chapter7_title, self.content.ch7_paragraphs, ch7_tables)

        # 确保输出目录存在
        out_dir = os.path.dirname(os.path.abspath(output_path))
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        doc.save(output_path)
        return output_path


def generate_report(content: ReportContent, output_path: str) -> str:
    """顶层便利函数：生成投研分析报告 .docx"""
    generator = ReportGenerator(content)
    return generator.generate(output_path)
