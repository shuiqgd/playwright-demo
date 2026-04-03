---
description: 执行从 PRD 解析到自动化测试并导出报告的全流程 (PRD -> Case -> Auto-Test -> Reports)
---

# 终极测试提效全流程 (Ultimate Test Automation Flow)

此工作流基于“AI 提效终极工具链”，整合了需求分析、多维设计、脚本生成及自动执行。

## 准备阶段
1. 准备好待解析的 PRD 文本、截图或接口文档。
2. 配置好 `input/config.json` 环境参数。

## 执行步骤

### 1. 深度测试设计 (prd_analyzer)
使用 `prd_analyzer` 进行全量设计：
> 指令: `use skill prd_analyzer 解析 [需求/截图]`
> **要求**: 生成包含测试点、用例、接口、数据及自动化建议的 JSON。
> **存储**: 保存至 `output/temp/prd_analyzer/output.json`。

### 2. 生成多维测试报告 (test_case_template_generator)
使用 `test_case_template_generator` 汇总设计资产：
> 指令: `use skill test_case_template_generator 处理 output/temp/prd_analyzer/output.json`
> **存储**: 将整合后的 JSON 保存至 `output/temp/test_case_template_generator/cases.json`。

// turbo
### 3. 物理导出 Excel
执行脚本生成行业级测试报告：
```powershell
python .agent/skills/test_case_template_generator/scripts/generate_test_cases.py output/temp/test_case_template_generator/cases.json
```
> 结果位于 `output/`，包含功能/接口/数据/自动化四Sheet。

### 4. 驱动自动化执行 (automated_tester)
启动 `automated_tester` 执行高优先级建议：
> 指令: `use skill automated_tester 执行 output/temp/test_case_template_generator/cases.json 中的核心回归`
> **输出**: 记录执行日志与截图。

// turbo
### 5. 汇总 Bug 与成果
生成最终的缺陷报告与执行汇总：
```powershell
python .agent/skills/automated_tester/scripts/run_automation.py [JSON结果路径]
```

## 交付产物
- `[项目名].xlsx`: 包含全量测试设计。
- `bugList.xlsx`: 自动化发现的缺陷。
- `测试成果.xlsx`: 执行合规性统计。
