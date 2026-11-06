"""
touyan2 — 产线落地投研分析引擎
==============================
面向地方政府招商部门、国资产业引导基金的结构化投研报告生成框架。
"""

__version__ = "1.1.0"
__author__ = "zhangpeifu"

from .models import ReportContent
from .builder import ReportGenerator, generate_report, build_chapter
from .styles import (
    add_heading_styled,
    add_para,
    add_styled_table,
    add_cover_page,
    add_toc_page,
    add_disclaimer,
)

__all__ = [
    "ReportContent",
    "ReportGenerator",
    "generate_report",
    "build_chapter",
    "add_heading_styled",
    "add_para",
    "add_styled_table",
    "add_cover_page",
    "add_toc_page",
    "add_disclaimer",
    "__version__",
]
