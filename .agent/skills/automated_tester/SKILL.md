---
name: automated_tester
description: 使用无头浏览器根据配置文件执行自动化测试，并导出相应的 Bug 列表和测试成果。
---

# 自动化测试器 (Automated Tester)

此 Skill 负责执行由 `prd_analyzer` 和 `test_case_template_generator` 产出的自动化测试建议。它利用 `browser_subagent` 在真实浏览器环境中验证业务逻辑。它从 `input/config.json` 读取环境参数，并对指定的测试用例进行验证。

## 角色定位
你是一位高级自动化测试工程师，具备熟练的浏览器操作能力，能够准确发现系统缺陷并记录详细的重现步骤。

## 配置文件规范 (`input/config.json`)
```json
{
  "test_url": "http://example.com/login", // 测试起始地址
  "username": "admin",                     // 登录用户名
  "password": "password123",               // 登录密码
  "test_level": ["P0", "P1"]               // 仅执行这些级别的测试用例
  "engine": "Playwright" // 推荐引擎
}
```

## 输出规则
### 1. Bug 列表 (`output/bugList.xlsx`)
每个 Bug 必须包含：
- **Bug 标题**: 简洁明了的功能描述。
- **严重程度**: 阻塞/严重/一般/轻微。
- **优先级**: 高/中/低。
- **重现步骤**: 必须包含 `[步骤]`、`[结果]`、`[期望]` 三部分。

### 2. 测试成果 (`output/测试成果.xlsx`)
包含所有测试用例的执行汇总：
- **列名**: 模块 | 用例 ID | 测试项 | 优先级 | 重要程度 | 状态 (通过/失败) | 结果 (具体报错或通过描述)

## 执行流程
1. **加载配置**: 读取 `input/config.json`。
2. **浏览器启动**: 使用 `browser_subagent` 打开 `test_url`。
3. **自动登录**: 若页面包含登录表单，自动填入 `username` 和 `password`。
4. **用例执行**:
    - 获取 `prd_analyzer` 或 `test_case_template_generator` 产出的 JSON 用例。
    - 筛选符合 `test_level` 的用例。
    - 逐一在浏览器中模拟操作并验证结果。
4. **异常捕获**: 若操作失败，自动截图（记录路径）并记录 Bug 详情。
5. **记录与导出**:
    - 将发现的问题记录到 JSON。
    - 提示运行 `python .agent/skills/automated_tester/scripts/run_automation.py` 完成 Excel 转换。

## 指令
- 运行测试前，请确保 `output/temp` 目录下已有最新的用例数据。
- 若遇到登录失败或环境不可达，应立即向用户反馈。
