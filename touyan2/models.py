"""
touyan2.models
==============
投研分析报告的核心数据模型与容器。
定义 7 章评估框架的字段结构，支持字典序列化及 JSON 读写。
"""

import json
from datetime import datetime
from typing import List, Dict, Any, Optional


class ReportContent:
    """
    投研分析报告内容容器。
    覆盖面向地方政府/国资招商项目的 7 维全景评估架构：
      1. 项目概览（含投资亮点前置、基本信息、投资节奏、资金明细）
      2. 行业与市场分析（赛道规模、宏观政策、竞争格局矩阵）
      3. 项目竞争力评估（技术、团队、产品力、在手订单）
      4. 财务效益分析（投资回收期、敏感性测算、关键假设）
      5. 地方落地价值评估（税收、就业、强链补链、要素诉求）
      6. 风险分析与缓释（技术/市场/履约/政策风险与管控措施）
      7. 综合评估与建议（综合评分卡、分期投资建议、尽调排雷清单）
    """

    def __init__(self, company: str, report_title: str = "投研分析报告", subtitle: str = "——产线落地项目投资价值评估——"):
        self.company = company
        self.report_title = report_title
        self.subtitle = subtitle
        self.date_str = datetime.now().strftime("%Y年%m月%d日")

        # ---- 第一章：项目概览 ----
        self.chapter1_title = "项目概览"
        self.ch1_highlights: List[str] = []
        self.ch1_paragraphs: List[str] = []
        self.ch1_basic_info_headers = ["维度", "内容"]
        self.ch1_basic_info_rows: List[List[str]] = []
        self.ch1_investment_headers = ["阶段", "时间", "投资金额", "主要用途"]
        self.ch1_investment_rows: List[List[str]] = []
        self.ch1_usage_headers = ["用途类别", "金额（万元）", "占比", "说明"]
        self.ch1_usage_rows: List[List[str]] = []

        # ---- 第二章：行业与市场分析 ----
        self.chapter2_title = "行业与市场分析"
        self.ch2_paragraphs: List[str] = []
        self.ch2_market_headers = ["指标", "当前基期", "未来预测", "复合增速 (CAGR)", "数据来源"]
        self.ch2_market_rows: List[List[str]] = []
        self.ch2_comp_headers = ["企业名称", "市场份额", "主要定位", "核心优势 / 短板"]
        self.ch2_comp_rows: List[List[str]] = []

        # ---- 第三章：项目竞争力评估 ----
        self.chapter3_title = "项目竞争力评估"
        self.ch3_paragraphs: List[str] = []
        self.ch3_swot_headers = ["评估维度", "企业现状与优势", "潜在短板与差距"]
        self.ch3_swot_rows: List[List[str]] = []
        self.ch3_product_headers = ["核心产品 / 技术", "性能指标", "对标行业水平", "自主可控度"]
        self.ch3_product_rows: List[List[str]] = []

        # ---- 第四章：财务效益分析 ----
        self.chapter4_title = "财务效益分析"
        self.ch4_paragraphs: List[str] = []
        self.ch4_return_headers = ["财务指标", "一期投产", "二期达产", "三期满负荷", "整体累计"]
        self.ch4_return_rows: List[List[str]] = []
        self.ch4_sens_headers = ["敏感性变量", "变动幅度", "净利润变动", "投资回收期影响", "抗风险承受力"]
        self.ch4_sens_rows: List[List[str]] = []

        # ---- 第五章：地方落地价值评估 ----
        self.chapter5_title = "地方落地价值评估"
        self.ch5_paragraphs: List[str] = []
        self.ch5_value_headers = ["评估维度", "地方政策要求 / 考核基准", "项目方承诺指标", "契合度评价"]
        self.ch5_value_rows: List[List[str]] = []
        self.ch5_policy_headers = ["资源诉求类别", "项目方申请标准", "地方常规支持上限", "建议谈判策略"]
        self.ch5_policy_rows: List[List[str]] = []

        # ---- 第六章：风险分析与缓释 ----
        self.chapter6_title = "风险分析与缓释"
        self.ch6_paragraphs: List[str] = []
        self.ch6_risk_headers = ["风险类别", "主要风险表现", "风险等级", "防范与缓释对策"]
        self.ch6_risk_rows: List[List[str]] = []

        # ---- 第七章：综合评估与建议 ----
        self.chapter7_title = "综合评估与建议"
        self.ch7_paragraphs: List[str] = []
        self.ch7_overall_verdict: str = ""
        self.ch7_score_headers = ["评估维度", "权重", "得分（百分制）", "核心研判意见"]
        self.ch7_score_rows: List[List[str]] = []
        self.ch7_stage_headers = ["推进阶段", "拟投入额度", "建议决策", "关键里程碑触发条件"]
        self.ch7_stage_rows: List[List[str]] = []
        self.ch7_dd_headers = ["重点尽调领域", "排查事项（Red Flag 核查）", "验证方式与证据底稿"]
        self.ch7_dd_rows: List[List[str]] = []

    def to_dict(self) -> Dict[str, Any]:
        """序列化为纯字典"""
        return {k: v for k, v in self.__dict__.items()}

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "ReportContent":
        """从字典重建 ReportContent 实例"""
        company = data.get("company", "目标企业")
        report_title = data.get("report_title", "投研分析报告")
        subtitle = data.get("subtitle", "——产线落地项目投资价值评估——")
        instance = cls(company=company, report_title=report_title, subtitle=subtitle)
        for k, v in data.items():
            setattr(instance, k, v)
        return instance

    def save_json(self, json_path: str, indent: int = 2):
        """将模型保存为 JSON 结构化数据文件"""
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=indent)

    @classmethod
    def load_json(cls, json_path: str) -> "ReportContent":
        """从 JSON 文件读取并生成 ReportContent"""
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls.from_dict(data)
