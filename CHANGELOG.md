# 更新日志 (Changelog)

所有关键架构变动与功能发布均记录于此。

## [v1.1.0] - 2026-11-07

### 核心升级
- **模块化架构重构**：新增 `touyan2` 原生 Python 核心库，支持 `from touyan2 import ReportContent, generate_report`，提供完整的面向对象模型与样式引擎。
- **现代化 CLI 工具**：新增 `touyan2` 命令行工具（`generate`、`template`、`version`），支持直接从 JSON 结构化数据批量渲染 Word 报告。
- **工业级实战范例**：
  - 新增高端精密增材制造装备范例（`examples/generate_advanced_manufacturing_report.py`）；
  - 新增合成生物学高活性制剂范例（`examples/generate_biotech_report.py`）；
  - 新增 JSON 自动化渲染脚本（`examples/generate_from_json.py`）与标准数据契约（`examples/sample_project_metadata.json`）。
- **专业级投研文档沉淀**：
  - 新增 `docs/FRAMEWORK_7_CHAPTERS.md`（7 维全景评估框架方法论）；
  - 新增 `docs/DUE_DILIGENCE_CHECKLIST.md`（投前 30 项尽调排雷清单与 Red Flags）；
  - 新增 `docs/GOVERNMENT_LANDING_POLICY_GUIDE.md`（地方招商政策要素与亩均财税测算指引）；
  - 新增 `docs/WORD_TYPOGRAPHY_SPEC.md`（公文级 Word 排版标准与色彩规范）。
- **CI / 质量工程**：引入 GitHub Actions 自动化 CI 流程，覆盖代码语法风格检验与多场景报告生成单元测试。

## [v1.0.0] - 初始版本
- 提供初步 7 章分析规范 (`SKILL.md`, `references/template.md`)；
- 提供基础报告生成脚本 (`scripts/gen_report.py`) 与内容提取重构工具 (`scripts/reformat_from_source.py`)。
