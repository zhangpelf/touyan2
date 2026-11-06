"""
touyan2.cli
===========
touyan2 命令行入口程序。
提供报告生成、空白模板导出、文档提取审查等功能。
"""

import argparse
import sys
import os
from .models import ReportContent
from .builder import generate_report


def cmd_generate(args):
    print("=== touyan2 投研分析报告生成器 ===")
    if args.data:
        if not os.path.exists(args.data):
            print(f"❌ 错误：指定的数据文件不存在：{args.data}", file=sys.stderr)
            sys.exit(1)
        print(f"从数据文件载入: {args.data}")
        content = ReportContent.load_json(args.data)
        if args.company:
            content.company = args.company
    else:
        content = ReportContent(args.company or "示例企业")

    if args.title:
        content.report_title = args.title
    if args.subtitle:
        content.subtitle = args.subtitle

    output_path = args.output
    out = generate_report(content, output_path)
    file_size_kb = os.path.getsize(out) / 1024
    print(f"✅ 报告已成功生成: {out}")
    print(f"   文件大小: {file_size_kb:.1f} KB")


def cmd_template(args):
    out_path = args.output or "touyan2_template.json"
    content = ReportContent("示例科技股份公司")
    content.ch1_highlights = [
        "**核心技术行业领先**：自主研发第三代高精度装备，关键指标对标国际顶尖水准",
        "**在手订单充足**：已签署年度框架供货协议超 1.2 亿元，客户覆盖航天科工及头部整车厂",
    ]
    content.ch1_basic_info_rows = [
        ["项目全称", "高端智能装备制造及关键零部件产线落地项目"],
        ["实施主体", "某某智能制造科技有限公司"],
        ["总投资额", "50,000 万元（固定资产投资 38,000 万元）"],
        ["厂房需求", "总面积约 25,000 平方米，需满足重型地坪承重"],
    ]
    content.save_json(out_path)
    print(f"✅ 模板已导出至: {out_path}")


def cmd_version(args):
    from . import __version__
    print(f"touyan2 v{__version__} — 产线落地投研分析引擎")


def build_parser():
    parser = argparse.ArgumentParser(
        prog="touyan2",
        description="touyan2 — 面向地方政府/国资平台的产线落地投研分析报告生成引擎",
    )
    subparsers = parser.add_subparsers(dest="subcommand", help="子命令")

    # generate 子命令
    p_gen = subparsers.add_parser("generate", help="生成投研分析报告 (.docx)")
    p_gen.add_argument("-c", "--company", help="目标企业全称")
    p_gen.add_argument("-o", "--output", required=True, help="输出 .docx 文件路径")
    p_gen.add_argument("-d", "--data", help="结构化 JSON 数据文件路径")
    p_gen.add_argument("--title", help="自定义报告主标题")
    p_gen.add_argument("--subtitle", help="自定义报告副标题")

    # template 子命令
    p_tmpl = subparsers.add_parser("template", help="导出结构化 JSON 数据模板")
    p_tmpl.add_argument("-o", "--output", default="template.json", help="导出的 JSON 文件路径")

    # version 子命令
    subparsers.add_parser("version", help="显示版本号")

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    if not args.subcommand:
        parser.print_help()
        sys.exit(0)

    if args.subcommand == "generate":
        cmd_generate(args)
    elif args.subcommand == "template":
        cmd_template(args)
    elif args.subcommand == "version":
        cmd_version(args)


if __name__ == "__main__":
    main()
