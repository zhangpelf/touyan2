#!/usr/bin/env python3
"""
examples/generate_from_json.py
==============================
演示如何直接从结构化 JSON 文件一键生成高规格投研分析报告 (.docx)。
特别适用于配合大语言模型 (LLM) 提取结构化数据后的自动化文档渲染。

用法：
    python3 examples/generate_from_json.py         --json examples/sample_project_metadata.json         --output ./output/艾沃模智能装备_产线落地投研分析报告.docx
"""

import argparse
import sys
import os

# 将项目根目录加入 sys.path 以支持直接本地导入
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from touyan2.models import ReportContent
from touyan2.builder import generate_report


def main():
    parser = argparse.ArgumentParser(description="从 JSON 结构化数据一键渲染 Word 投研报告")
    parser.add_argument("-j", "--json", default="examples/sample_project_metadata.json", help="JSON 数据文件路径")
    parser.add_argument("-o", "--output", default="./output/艾沃模智能装备_产线落地投研分析报告.docx", help="输出 .docx 路径")
    args = parser.parse_args()

    if not os.path.exists(args.json):
        print(f"❌ 找不到输入 JSON 文件: {args.json}", file=sys.stderr)
        sys.exit(1)

    print(f"[*] 正在读取项目元数据: {args.json}")
    content = ReportContent.load_json(args.json)
    print(f"[*] 目标企业: {content.company}")
    print(f"[*] 报告主标题: {content.report_title}")

    out_file = generate_report(content, args.output)
    size_kb = os.path.getsize(out_file) / 1024
    print(f"[+] 报告生成成功: {out_file} (文件大小: {size_kb:.1f} KB)")


if __name__ == "__main__":
    main()
