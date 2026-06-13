#!/usr/bin/env python3
"""
reformat_from_source.py — 已有投研报告 → touyan2 格式重构辅助工具
================================================================
读取已有的投研报告（.docx）,提取所有段落和表格内容,
输出为结构化的参考材料 + assign_content() 模板框架。

不生成完整报告。只做提取 + 模板输出——内容映射需要 AI 判断。

用法：
    python3 reformat_from_source.py \\
        --source "/path/to/源报告.docx" \\
        --company "公司名称" \\
        --output "/path/to/内容提取.txt"

输出格式：
    1. 源报告概览（段落数、表格数、总字数）
    2. 全部段落（带索引 + 样式标注）
    3. 全部表格（带索引 + 行列展开）
    4. assign_content() 模板（空壳,待填充）
"""

import argparse
import os
import re
import sys
from datetime import datetime
from docx import Document


def extract_content(input_path: str):
    """读取 .docx,返回 (paragraphs, tables)。
    
    paragraphs: list of (index[int], text[str], style[str])
    tables:     list of (index[int], rows[list[list[str]]])
    """
    doc = Document(input_path)

    paragraphs = []
    for i, p in enumerate(doc.paragraphs, 1):
        text = p.text.strip()
        if not text:
            continue
        style_name = p.style.name if p.style else "Normal"
        paragraphs.append((i, text, style_name))

    tables = []
    for t_idx, table in enumerate(doc.tables, 1):
        rows = []
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            rows.append(cells)
        tables.append((t_idx, rows))

    return paragraphs, tables


# ═══════════════════════════════════════════════════════════════
#  输出构建
# ═══════════════════════════════════════════════════════════════

def build_overview(paragraphs, tables) -> list:
    lines = []
    total_words = sum(len(t) for _, t, _ in paragraphs)
    heading_count = sum(1 for _, _, s in paragraphs if s.startswith("Heading"))
    table_row_count = sum(len(rows) for _, rows in tables)

    lines.append("=" * 66)
    lines.append("  源报告内容提取")
    lines.append("=" * 66)
    lines.append("")
    lines.append(f"  段落数：    {len(paragraphs)}（其中标题 {heading_count} 个）")
    lines.append(f"  表格数：    {len(tables)}（总行数 {table_row_count} 行）")
    lines.append(f"  总字数：    {total_words} 字")
    lines.append("")
    return lines


def build_paragraphs_section(paragraphs) -> list:
    lines = []
    lines.append("─" * 66)
    lines.append("  一、段落内容")
    lines.append("─" * 66)
    lines.append("")

    prev_style = None
    for idx, text, style in paragraphs:
        if style.startswith("Heading"):
            if prev_style and not prev_style.startswith("Heading"):
                lines.append("")
            lines.append(f"  ■ [{idx}] [{style}] {text}")
            lines.append("")
        else:
            lines.append(f"  [{idx}] {text}")
            lines.append("")
        prev_style = style

    return lines


def build_tables_section(tables) -> list:
    lines = []
    if not tables:
        lines.append("（无表格）")
        lines.append("")
        return lines

    lines.append("─" * 66)
    lines.append("  二、表格内容")
    lines.append("─" * 66)
    lines.append("")

    for t_idx, rows in tables:
        lines.append(f"  【表格 {t_idx}】（{len(rows)} 行 x {len(rows[0]) if rows else 0} 列）")
        lines.append("")
        for r_idx, row in enumerate(rows):
            prefix = "    ┌ " if r_idx == 0 else "    ├ "
            display = [c[:40] for c in row[:8]]
            lines.append(f"{prefix}{' │ '.join(display)}")
        lines.append("")
        lines.append("")

    return lines


def build_assign_content_template(company, paragraphs, tables) -> list:
    """输出空壳 assign_content(),供用户/AI 填充"""
    lines = []
    lines.append("─" * 66)
    lines.append("  三、assign_content() 模板（待填充）")
    lines.append("─" * 66)
    lines.append("")
    lines.append("# 以下为 gen_touyan2_report.py 中 assign_content() 的填充模板。")
    lines.append("# 将上方提取的源报告内容按 7 章映射后填入此处。")
    lines.append("# 使用方法：复制到 gen_touyan2_report.py 中替换 assign_content()。")
    lines.append("")

    # 第一章
    lines.append("    # ══════════ 第一章 项目概览 ══════════")
    lines.append("    content.ch1_paragraphs = [")
    lines.append(f'        "【TODO】{company}本次拟在...建设...产线,项目总投资额约...亿元。",')
    lines.append(f'        "【TODO】项目主要建设内容包括...,规划年产能...亿元。",')
    lines.append("    ]")
    lines.append("")
    lines.append("    content.ch1_basic_info_rows = [")
    lines.append('        ["项目名称", f"{content.company}产线落地项目"],')
    lines.append('        ["实施主体", content.company],')
    lines.append('        ["投资总额", "【TODO】"],')
    lines.append('        ["建设内容", "【TODO】"],')
    lines.append('        ["厂房需求", "【TODO】"],')
    lines.append('        ["建设周期", "【TODO】"],')
    lines.append('        ["带动就业", "【TODO】"],')
    lines.append("    ]")
    lines.append("")
    lines.append("    content.ch1_investment_rows = [")
    lines.append('        ["一期", "【TODO】", "【TODO】亿元", "【TODO】"],')
    lines.append('        ["二期", "【TODO】", "【TODO】亿元", "【TODO】"],')
    lines.append('        ["三期", "【TODO】", "【TODO】亿元", "【TODO】"],')
    lines.append("    ]")
    lines.append("")
    lines.append("    content.ch1_usage_rows = [")
    lines.append('        ["设备购置", "【TODO】亿元", "【TODO】%"],')
    lines.append('        ["厂房建设", "【TODO】亿元", "【TODO】%"],')
    lines.append('        ["流动资金", "【TODO】亿元", "【TODO】%"],')
    lines.append('        ["其他", "【TODO】亿元", "【TODO】%"],')
    lines.append("    ]")
    lines.append("")

    # 第二章
    lines.append("    # ══════════ 第二章 行业与市场分析 ══════════")
    lines.append("    content.ch2_paragraphs = [")
    lines.append('        "【TODO】行业宏观背景...",')
    lines.append('        "【TODO】市场规模及增长...",')
    lines.append('        "【TODO】竞争格局...",')
    lines.append("    ]")
    lines.append("")
    lines.append("    content.ch2_market_rows = [")
    lines.append('        ["全球市场规模", "【TODO】", "【TODO】"],')
    lines.append('        ["国内市场规模", "【TODO】", "【TODO】"],')
    lines.append('        ["年复合增速", "—", "【TODO】%"],')
    lines.append("    ]")
    lines.append("")
    lines.append("    content.ch2_comp_rows = [")
    lines.append('        ["【头部企业A】", "【TODO】%", "【TODO】", "【TODO】"],')
    lines.append('        ["【头部企业B】", "【TODO】%", "【TODO】", "【TODO】"],')
    lines.append('        [content.company, "【TODO】%", "【TODO】", "【TODO】"],')
    lines.append("    ]")
    lines.append("")

    # 第三章
    lines.append("    # ══════════ 第三章 项目竞争力评估 ══════════")
    lines.append("    content.ch3_swot_rows = [")
    lines.append('        ["团队背景", "【从源报告提取】"],')
    lines.append('        ["技术能力", "【从源报告提取】"],')
    lines.append('        ["产品力", "【从源报告提取】"],')
    lines.append('        ["客户基础", "【从源报告提取】"],')
    lines.append('        ["产能规模", "【从源报告提取】"],')
    lines.append('        ["成本优势", "【从源报告提取】"],')
    lines.append('        ["品牌/渠道", "【从源报告提取】"],')
    lines.append("    ]")
    lines.append("")
    lines.append("    content.ch3_paragraphs = [")
    lines.append('        "【TODO】综合竞争力判断...",')
    lines.append('        "【TODO】差异化策略评估...",')
    lines.append('        "【TODO】关键不确定性...",')
    lines.append("    ]")
    lines.append("")

    # 第四章
    lines.append("    # ══════════ 第四章 财务效益分析 ══════════")
    lines.append("    content.ch4_return_rows = [")
    lines.append('        ["投资额", "【TODO】", "【TODO】", "【TODO】", "【TODO】"],')
    lines.append('        ["达产年产值", "【TODO】", "【TODO】", "【TODO】", "【TODO】"],')
    lines.append('        ["达产年净利润", "【TODO】", "【TODO】", "【TODO】", "【TODO】"],')
    lines.append('        ["投资回报率", "【TODO】%", "【TODO】%", "【TODO】%", "【TODO】%"],')
    lines.append('        ["静态回收期", "【TODO】年", "【TODO】年", "【TODO】年", "【TODO】年"],')
    lines.append("    ]")
    lines.append("")
    lines.append("    content.ch4_paragraphs = [")
    lines.append('        "【TODO】核心假设说明...",')
    lines.append('        "【TODO】敏感性分析...",')
    lines.append("    ]")
    lines.append("")

    # 第五章
    lines.append("    # ══════════ 第五章 地方落地价值评估 ══════════")
    lines.append("    content.ch5_paragraphs = [")
    lines.append('        "【TODO】对地方经济的综合带动效益评估...",')
    lines.append('        "【TODO】政策需求清单及可行性评估...",')
    lines.append("    ]")
    lines.append("")
    lines.append("    content.ch5_value_rows = [")
    lines.append('        ["产业契合度", "【TODO】"],')
    lines.append('        ["投资密度", "【TODO】"],')
    lines.append('        ["产值贡献", "【TODO】"],')
    lines.append('        ["就业带动", "【TODO】"],')
    lines.append('        ["产业带动", "【TODO】"],')
    lines.append('        ["招商条件", "【TODO】"],')
    lines.append('        ["品牌效应", "【TODO】"],')
    lines.append("    ]")
    lines.append("")

    # 第六章
    lines.append("    # ══════════ 第六章 风险分析与缓释 ══════════")
    lines.append("    content.ch6_paragraphs = [")
    lines.append('        "经综合评估,本项目面临的主要风险集中在以下方面：",')
    lines.append("    ]")
    lines.append("")
    lines.append("    content.ch6_risk_rows = [")
    lines.append('        ["技术风险", "【TODO】", "高/中/低", "【TODO】"],')
    lines.append('        ["市场风险", "【TODO】", "高/中/低", "【TODO】"],')
    lines.append('        ["资金风险", "【TODO】", "高/中/低", "【TODO】"],')
    lines.append('        ["团队风险", "【TODO】", "高/中/低", "【TODO】"],')
    lines.append('        ["建设风险", "【TODO】", "高/中/低", "【TODO】"],')
    lines.append('        ["政策风险", "【TODO】", "高/中/低", "【TODO】"],')
    lines.append('        ["竞争风险", "【TODO】", "高/中/低", "【TODO】"],')
    lines.append("    ]")
    lines.append("")

    # 第七章
    lines.append("    # ══════════ 第七章 综合评估与建议 ══════════")
    lines.append("    content.ch7_score_rows = [")
    lines.append('        ["赛道前景", "★★★☆☆", "【TODO】"],')
    lines.append('        ["项目本身质量", "★★★☆☆", "【TODO】"],')
    lines.append('        ["投资条件", "★★★☆☆", "【TODO】"],')
    lines.append('        ["产值/税收贡献", "★★★☆☆", "【TODO】"],')
    lines.append('        ["风险可控度", "★★★☆☆", "【TODO】"],')
    lines.append("    ]")
    lines.append("")
    lines.append("    content.ch7_paragraphs = [")
    lines.append('        "总体判断：【推荐落地 / 审慎推进 / 不推荐】。",')
    lines.append("    ]")
    lines.append("")
    lines.append("    content.ch7_stage_rows = [")
    lines.append('        ["一期（【TODO】万元）", "建议投资 / 审慎 / 暂缓", "① 【TODO】 ② 【TODO】"],')
    lines.append('        ["二期（【TODO】万元）", "条件触发 / 暂不考虑", "① 【TODO】 ② 【TODO】"],')
    lines.append('        ["三期（【TODO】万元）", "机会导向 / 暂不考虑", "① 【TODO】 ② 【TODO】"],')
    lines.append("    ]")
    lines.append("")
    lines.append("    content.ch7_dd_rows = [")
    lines.append('        ["【尽调重点1】", "【TODO】"],')
    lines.append('        ["【尽调重点2】", "【TODO】"],')
    lines.append('        ["【尽调重点3】", "【TODO】"],')
    lines.append("    ]")
    lines.append("")

    return lines


# ═══════════════════════════════════════════════════════════════
#  主入口
# ═══════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="已有投研报告 → touyan2 格式重构辅助工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""使用步骤：
  1. 运行脚本提取源报告内容
  2. 阅读提取结果,理解源报告结构
  3. 按 touyan2 7 章模板在对话中重构内容
  4. 使用 gen_touyan2_report.py 生成最终报告
""",
    )
    parser.add_argument("--source", required=True, help="源报告 .docx 路径")
    parser.add_argument("--company", required=True, help="目标公司名称")
    parser.add_argument("--output", required=True, help="输出文本文件路径")
    args = parser.parse_args()

    if not os.path.exists(args.source):
        print(f"❌ 源文件不存在: {args.source}")
        sys.exit(1)

    print(f"=== touyan2 格式重构辅助工具 ===")
    print(f"源报告:   {args.source}")
    print(f"公司:     {args.company}")
    print(f"输出:     {args.output}")
    print()

    paragraphs, tables = extract_content(args.source)
    print(f"提取完成：{len(paragraphs)} 段,{len(tables)} 表")

    out_lines = []
    out_lines.append(f"# touyan2 格式重构 — 内容提取")
    out_lines.append(f"# 源报告：{args.source}")
    out_lines.append(f"# 公司：  {args.company}")
    out_lines.append(f"# 时间：  {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    out_lines.append("")

    out_lines.extend(build_overview(paragraphs, tables))
    out_lines.extend(build_paragraphs_section(paragraphs))
    out_lines.extend(build_tables_section(tables))
    out_lines.extend(build_assign_content_template(args.company, paragraphs, tables))

    output_text = "\n".join(out_lines)
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(output_text)

    file_size = os.path.getsize(args.output)
    print(f"✅ 已输出: {args.output}")
    print(f"   文件大小: {file_size / 1024:.1f} KB")
    print(f"   总行数: {output_text.count(chr(10))}")


if __name__ == "__main__":
    main()
