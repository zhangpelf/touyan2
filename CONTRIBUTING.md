# 贡献指南 (CONTRIBUTING)

感谢对 `touyan2`（产线落地投研分析引擎）的关注与支持！

## 一、 核心设计哲学
在提交代码或编写新评估模板时，请务必遵守以下三项铁律：

1. **结论先行、数据锚点**：每个章节的第一句话必须是定性评估结论；所有论点必须有加粗的数据指标或第三方独立出处作支撑。
2. **事实与宣称严格分离**：对未经审计或独立核实的企业单方陈述，必须显式标注「*项目方声称，待尽调核实*」，严禁将未经证实的宣称直接当作既定事实。
3. **国资风控与排雷导向**：重点关注固定资产真实性、知识产权权属争议、亩均税收对赌可行性与清算回收抓手。

## 二、 本地开发与测试

1. 克隆仓库并安装开发依赖：
   ```bash
   git clone https://github.com/zhangpelf/touyan2.git
   cd touyan2
   pip install -e .
   pip install -r requirements-dev.txt
   ```

2. 运行自动化单元测试：
   ```bash
   python3 -m unittest discover tests
   ```

3. 运行实战案例验证文档生成：
   ```bash
   python3 examples/generate_advanced_manufacturing_report.py
   python3 examples/generate_biotech_report.py
   python3 examples/generate_from_json.py
   ```

## 三、 Pull Request 流程
1. Fork 本仓库并基于 `main` 分支创建特性分支（如 `feature/financial-stress-test`）。
2. 保持代码简洁易读，确保中西文排版格式统一。
3. 提交 PR 并详细描述改动点与业务价值。
