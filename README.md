# touyan2 — 产线落地投研分析引擎

<p align="center">
  <strong>面向地方政府招商引资、经开区管委会与国资产业引导基金的结构化投研评估引擎</strong>
</p>

<p align="center">
  <a href="https://github.com/zhangpelf/touyan2"><img src="https://img.shields.io/badge/python-3.8+-blue.svg" alt="Python 3.8+"></a>
  <a href="https://github.com/zhangpelf/touyan2/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-MIT-green.svg" alt="License MIT"></a>
  <a href="https://github.com/zhangpelf/touyan2/actions"><img src="https://img.shields.io/badge/CI-passing-brightgreen.svg" alt="CI Status"></a>
  <a href="https://github.com/zhangpelf/touyan2/pulls"><img src="https://img.shields.io/badge/PRs-welcome-orange.svg" alt="PRs Welcome"></a>
  <a href="https://github.com/zhangpelf/touyan2"><img src="https://img.shields.io/badge/Word-docx%20generator-1A3C6E.svg" alt="Word docx generator"></a>
</p>

---

## 📖 项目定位：为什么需要 touyan2？

当招商部门或国资平台拿到一份企业的《商业计划书》（BP）或《产线投资方案》时，常常面临以下核心痛点：

1. **企业自述水分大**：BP 中充斥着“国际领先”、“千亿蓝海”、“行业第一”等修饰词，缺少真实的固定资产明细、工艺良率与在手订单证据；
2. **诉求失衡与对赌风险**：项目方索要大额厂房免租、设备补贴及政府产业基金直投，但承诺的税收产值缺乏科学测算，地方政府面临政策被套取甚至国有资产流失风险；
3. **缺少标准化的研判抓手**：传统财务 VC 关注 IRR 与资本市场估值退出，而国资招商看重**强链补链、固投真实性、能耗环境指标、亩均税收与对赌回收保障**；
4. **材料杂乱无法直接呈报**：从分散的沟通纪要到形成一份符合体制内/国资审阅规范的 Word 报告，需要耗费大量排版与数据核验工时。

**`touyan2`（投研2）** 应运而生。它是一个**产线落地投研分析与自动化文档生成框架**。通过**7 维全景评估模型**，将杂乱的企业申报材料剥离重组，快速生成一份**结论先行、数据驱动、风险显式、符合国资政企公文级排版标准**的完整投研分析报告（`.docx`）。

---

## 🚀 核心特性

- 🏛️ **专属国资招商视角**：严格区分“财务 VC 投估值”与“地方国资投产业落地”的本质差异，聚焦亩均税收、固投强度与产业生态协同。
- 📊 **7 维全景评估架构**：覆盖投资概览、行业空间、竞争力对标、财务效益压力测试、地方落地价值、风险缓释以及综合评分卡。
- ⚖️ **事实与宣称严格分离**：确立“未核实企业自述一律标注待尽调确认”原则，自动输出投前排雷（Red Flags）核查清单。
- 📑 **公文级 Word 排版标准**：内置政企商务深青蓝体系（`#1A3C6E`）、中西文双轨字体（黑体/宋体 + Times New Roman）、三线商务表格、交替行底纹与严苛的防孤行分页控制。
- 💻 **双模调用支持**：既支持标准 Python 库导入（`from touyan2 import ...`），也提供现代化的命令行 CLI 工具与 JSON 自动化流水线。
- 🛡️ **30+ 项投前尽调核验清单**：内附涵盖知识产权权属、关联交易、设备真实估值、多地套补防范等全生命周期尽调指引。

---

## 🔄 端到端评估工作流

```mermaid
flowchart TD
    A[📄 原始商业计划书 / 产线计划 (.docx/.pdf/交流纪要)] --> B[🔍 Step 1: 资料提取与事实边界界定]
    B --> C{是否具备落地诉求?}
    C -->|是: 固投/厂房/补贴/基金| D[📋 走完整 7 章国资投研框架]
    C -->|否: 纯轻资产模式| E[📑 走简化版产品与市场研判]
    D --> F[⭐ Step 1.5: 提炼 4-6 条核心投资亮点 (数据锚点前置)]
    F --> G[🔬 Step 2: 7 维全景结构化分析与敏感性测算]
    G --> H[🛠️ Step 3: 渲染公文级排版 Word 报告 (touyan2.builder)]
    H --> I[🎯 Step 4: 输出分期决策建议、投资对赌协议与尽调 Red Flags]
```

---

## 📚 7 维全景评估章节体系

| 章节 | 核心研判任务 | 必备核心表格与量化输出 | 核心风控关注点 (Red Flags) |
| :--- | :--- | :--- | :--- |
| **第一章 项目概览** | 厘清谁在投资、投向何处、自筹比例 | 投资亮点前置、基本信息表、投资分期表、资金用途拆解表 | ⚠️ 总投虚高、自有出资低于 30%、重度依赖政府垫资 |
| **第二章 行业与市场分析** | 检验赛道天花板与真实市场需求 | 市场规模预测表 (CAGR)、产业链竞争格局矩阵表 | ⚠️ 产能过剩赛道新扩产能、引用不可信数据虚夸占有率 |
| **第三章 项目竞争力评估** | 考察核心技术护城河与供应链掌控 | SWOT 优劣势矩阵、核心产品技术参数横向对标表 | ⚠️ 核心专利职务发明争议、关键元器件进口断供、代工改标 |
| **第四章 财务效益分析** | 审慎测算投资回收期与抗风险韧性 | 分期产值与效益测算表、核心变量敏感性分析表 | ⚠️ 忽视爬坡期亏损、盲目乐观毛利率、缺乏压力测试 |
| **第五章 地方落地价值** | 锚定税收产出、拉动就业与强链补链 | 亩均税收/产值考核表、政策诉求与谈判应对表 | ⚠️ 亩均税收不达园区门槛、高能耗/危废无排放指标 |
| **第六章 风险分析与缓释** | 显式列出全流程风险并配备防御抓手 | 关键风险清单与缓释措施表（技术/市场/履约/政策） | ⚠️ 拿完补贴撤退迁址、大客户回款账期过长拖垮现金流 |
| **第七章 综合评估与建议** | 明确决策结论（推进/审慎/否决） | 综合研判评分卡、分期投资里程碑表、尽调核心排查清单 | ⚠️ 缺乏对赌回收机制、未约定固定资产投资专项审计 |

---

## ⚡ 快速开始

### 1. 安装

```bash
# 从 GitHub 克隆并安装
git clone https://github.com/zhangpelf/touyan2.git
cd touyan2
pip install -e .
```

### 2. 方式 A：使用 CLI 命令行工具

```bash
# 1. 导出空白或示例 JSON 数据模板
touyan2 template -o my_project.json

# 2. 从 JSON 数据一键生成公文级 Word 报告
touyan2 generate -d my_project.json -o ./output/产线落地投研分析报告.docx

# 3. 命令行极速生成模板报告
touyan2 generate -c "某某高新制造股份公司" -o ./output/初审报告.docx
```

### 3. 方式 B：通过 Python API 进行编程构建

```python
from touyan2 import ReportContent, generate_report

# 1. 实例化报告模型
content = ReportContent(
    company="苏州艾沃模智能装备科技股份有限公司",
    report_title="高端精密装备及航空构件产线落地投研报告",
    subtitle="——先进制造业落地项目投资研判与地方国资招商决策参考——"
)

# 2. 填充投资亮点前置
content.ch1_highlights = [
    "**技术指标行业领先**：关键尺寸公差控制在 ±0.015mm，粗糙度达 Ra 1.6μm",
    "**在手订单饱满**：已与两大航空主机厂签署 1.6 亿元供货框架协议",
    "**地方经济效益显著**：规划总投资 4.8 亿元，达产后亩均税收超 65 万元/年"
]

# 3. 填充章节段落与数据表格
content.ch1_paragraphs = ["本项目拟在地方高新技术产业园区建设 16 条高精密增材产线..."]
content.ch1_basic_info_rows = [
    ["项目全称", "高端精密增材制造装备产线项目"],
    ["规划总投资", "48,000 万元（一期固投 26,000 万元）"],
    ["规划用地", "工业用地 45 亩（标准厂房 22,000 ㎡）"]
]

# 4. 一键生成高质量 .docx
generate_report(content, "./output/投研报告_艾沃模.docx")
print("✅ 报告生成成功！")
```

---

## 🧪 实战范例库 (Examples)

仓库在 `examples/` 目录下提供了多个工业级实战脚本与标准数据，可直接运行验证：

```bash
# 运行先进制造业（金属 3D 打印与航空零部件）完整投研报告生成
python3 examples/generate_advanced_manufacturing_report.py

# 运行生物医药与合成生物制造（靶向噬菌体制剂）完整投研报告生成
python3 examples/generate_biotech_report.py

# 从标准 JSON 文件读取并生成报告
python3 examples/generate_from_json.py --json examples/sample_project_metadata.json
```

---

## 🎨 排版与公文级视觉规范

生成的 `.docx` 文档严格符合政企与金融投资机构报告的视觉标准：

```
+---------------------------------------------------------------+
|                       投 研 分 析 报 告                        |
|              ——产线落地项目投资价值评估与招商决策参考——              |
|                                                               |
|                 目标企业：某某智能科技有限公司                     |
|                 评估机构：产业投研与招商评估小组                   |
|                 报告日期：2026年11月07日                       |
+---------------------------------------------------------------+
                             [分页符]
+---------------------------------------------------------------+
| 【核心投资亮点前置】                                            |
| ● 核心技术国际对标领先：独创激光复合工艺，关键公差控制在 ±0.015mm... |
| ● 在手订单饱满度高：已锁定国内两大主机厂近 1.6 亿元长期订单...       |
+---------------------------------------------------------------+
                             [分页符]
+---------------------------------------------------------------+
| 一、项目概览                                                  |
|   本项目针对目标企业拟在开发区建设的高端制造产线进行系统评估...      |
|                                                               |
| ▎基本信息表                                                   |
| +---------------------+-------------------------------------+ |
| | 维度 (深蓝底白字)     | 内容 (9.5pt，交替浅灰底纹)           | |
| +---------------------+-------------------------------------+ |
| | 实施主体            | 某某高新技术产业发展有限公司          | |
| | 总投资额            | 48,000 万元 (固定资产投资 38,000 万元) | |
+---------------------------------------------------------------+
```

---

## 📁 仓库文件结构

```
touyan2/
├── touyan2/                       # 核心 Python 库
│   ├── __init__.py                # 顶层导出 (ReportContent, generate_report, 等)
│   ├── models.py                  # 7 章数据容器与 JSON 读写支持
│   ├── builder.py                 # Word 文档组装与章节渲染引擎
│   ├── styles.py                  # 公文级排版美学与 XML 样式工具
│   └── cli.py                     # touyan2 命令行入口程序
├── docs/                          # 深度方法论与风控指引
│   ├── FRAMEWORK_7_CHAPTERS.md    # 7 维评估框架深度实战指南
│   ├── DUE_DILIGENCE_CHECKLIST.md # 投前 30 项尽调排雷 Checklist (Red Flags)
│   ├── GOVERNMENT_LANDING_POLICY_GUIDE.md # 地方招商要素保障与财税测算指引
│   └── WORD_TYPOGRAPHY_SPEC.md    # 政企投资报告 Word 排版美学标准
├── examples/                      # 实战案例库
│   ├── generate_advanced_manufacturing_report.py # 先进制造业落地案例脚本
│   ├── generate_biotech_report.py # 生物医药制药落地案例脚本
│   ├── generate_from_json.py      # JSON 数据渲染脚本
│   └── sample_project_metadata.json # 真实项目级数据契约范本
├── scripts/                       # 辅助工具脚本
│   ├── gen_report.py              # 基础模板脚本 (兼容历史用法)
│   └── reformat_from_source.py    # 源材料提取与映射结构工具
├── references/                    # 规范参考
│   └── template.md                # 7 章内容规范要点清单
├── tests/                         # 单元测试套件
│   ├── test_models.py             # 数据模型与序列化测试
│   └── test_builder.py            # Word 生成与排版测试
├── .github/                       # GitHub 协作规范
│   ├── workflows/ci.yml           # 自动化 CI 测试流水线
│   ├── ISSUE_TEMPLATE/            # Bug 与 Feature 模板
│   └── pull_request_template.md   # PR 审核清单模板
├── SKILL.md                       # AI Agent 主技能设定与提示词指引
├── pyproject.toml                 # 现代构建标准配置文件
├── setup.py                       # 传统安装支持
├── requirements.txt               # 运行依赖
├── requirements-dev.txt           # 开发依赖
├── CHANGELOG.md                   # 版本演进日志
├── CONTRIBUTING.md                # 贡献者指南
├── LICENSE                        # MIT 开源许可证
└── README.md                      # 项目说明主文档
```

---

## 📖 深度文档索引

- [7 维全景评估框架深度指南 (docs/FRAMEWORK_7_CHAPTERS.md)](docs/FRAMEWORK_7_CHAPTERS.md)
- [投前尽职调查全量排雷清单 (docs/DUE_DILIGENCE_CHECKLIST.md)](docs/DUE_DILIGENCE_CHECKLIST.md)
- [地方招商落地政策要素保障与财税测算指引 (docs/GOVERNMENT_LANDING_POLICY_GUIDE.md)](docs/GOVERNMENT_LANDING_POLICY_GUIDE.md)
- [投资评估报告公文级排版规范指南 (docs/WORD_TYPOGRAPHY_SPEC.md)](docs/WORD_TYPOGRAPHY_SPEC.md)
- [AI Agent 技能规范说明书 (SKILL.md)](SKILL.md)

---

## 🤝 参与贡献

欢迎提交 Issue、提出新评估维度的建议或贡献代码！请在发起 PR 之前仔细阅读 [贡献指南 (CONTRIBUTING.md)](CONTRIBUTING.md)。

---

## 📄 License

本项目采用 [MIT 许可证](LICENSE)。自由用于地方政府招商、国资投研、产业咨询及相关学术研究。
