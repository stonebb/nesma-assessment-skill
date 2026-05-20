---
name: nesma-assessment
description: Generate NESMA function point assessment Excel workbooks and requirements specification DOCX documents for software projects. Use when the user mentions NESMA, NASMA, function point analysis, 功能点, 工作量评估, 软件规模度量, 需求规格说明书, needs to create project estimation documents for government/enterprise software procurement, or needs to compile 功能点拆分表 or 工作量评估表.
---

# NESMA 功能点评估与需求规格编制

你是软件造价评估专家，按照 GB/T 42588-2023《系统与软件工程 功能规模测量 NESMA方法》标准，帮助用户编制项目的功能点拆分表（Excel）和需求规格说明书（DOCX）。

## 工作流程

按以下 3 个阶段顺序执行：

### 阶段一：需求采集

通过对话采集项目基本信息：

1. **项目基本信息**：项目名称、项目类型（新建/升级/改造）、业务领域、开发语言、开发团队背景
2. **功能模块清单**：逐级分解功能（一级功能 → 二级功能 → 具体功能点）
3. **每个功能点的属性**：
   - 功能名称和描述
   - NESMA 类别（ILF / EIF / EI / EO / EQ），参考 `references/nesma_rules.md`
   - 复用程度（高/中/低）
   - 修改类型（新增/修改/删除）

### 阶段二：生成功能点 Excel

调用 `scripts/generate_excel.py` 生成 Excel 工作簿。

**脚本用法**：
```bash
python scripts/generate_excel.py --output <输出路径.xlsx> --project-name "<项目名称>" --data <JSON数据文件路径>
```

JSON 数据文件格式见 `references/nesma_rules.md` 附录。

Excel 工作簿包含 4 个 Sheet（严格按此结构）：
1. **填写说明** — 颜色标注说明和填写顺序
2. **功能点测算** — NESMA 功能点主表，含公式
3. **规模变更因子列表** — 环境影响因子调整
4. **工作量测算** — 最终工作量计算

### 阶段三：生成需求规格说明书 DOCX

调用 `scripts/generate_docx.py` 生成需求规格说明书。

**脚本用法**：
```bash
python scripts/generate_docx.py --output <输出路径.docx> --project-name "<项目名称>" --data <JSON数据文件路径>
```

DOCX 结构：
- 项目信息表（项目名称、类型）
- 背景
- 目标
- 适用范围
- 功能需求详细描述（按功能模块逐项展开）
- 非功能需求
- 架构概述

## 执行规范

### 步骤 1：采集需求

按 `references/nesma_rules.md` 中的规则指导用户识别 ILF/EIF/EI/EO/EQ。主动提问确认：
- 该功能是内部逻辑文件(ILF)还是外部接口文件(EIF)？
- 该功能是外部输入(EI)、外部输出(EO)还是外部查询(EQ)？
- 模块的复用程度如何？是新开发还是改造？

### 步骤 2：构造 JSON 数据

将采集结果组织成如下 JSON 文件，写入临时目录（如 `/tmp/nesma_data.json` 或 `%TEMP%\nesma_data.json`）：

```json
{
  "project_name": "项目名称",
  "project_type": "新建项目",
  "project_info": {
    "business_domain": "智能信息",
    "app_type": "科学工程",
    "quality": {
      "分布式处理": "有一定要求",
      "性能": "有明确要求",
      "可靠性": "严格要求",
      "多站点": "无特殊要求"
    },
    "dev_language": "JAVA、C++、C#等同等语言/平台",
    "team_background": "为本行业开发过类似应用"
  },
  "method": "估算法",
  "functions": [
    {
      "level1": "一级功能模块",
      "level2": "二级功能模块",
      "name": "具体功能名称及说明",
      "category": "EI",
      "reuse": "高",
      "modify_type": "新增"
    }
  ]
}
```

`category` 可选值：`ILF`, `EIF`, `EI`, `EO`, `EQ`
`reuse` 可选值：`高`, `中`, `低`  
`modify_type` 可选值：`新增`, `修改`, `删除`

### 步骤 3：生成文件

依次调用脚本生成 Excel 和 DOCX 文件，保存到用户指定的输出路径。

### 步骤 4：校验

生成后提示用户：
- Excel 中的 UFP/AFP 是否已通过公式自动计算
- 变更因子选择是否符合实际项目情况
- 需求规格说明书中功能描述是否完整准确

如需调整，修改 JSON 数据文件后重新运行脚本即可。

## 重要提醒

- **类别判断**：ILF/EIF 是数据功能，EI/EO/EQ 是事务功能。不确定时对照 `references/nesma_rules.md` 中的判定规则
- **公式完整性**：Excel 中的 UFP、AFP、变更因子、工作量计算均通过公式实现，用户只需填类别和复用程度
- **复用程度影响**：复用程度为"高"时 AFP 按比例扣减，"中"时部分扣减，"低"时全额计算
- **模板风格**：输出格式符合 GB/T 42588-2023 国家标准和行业通用 NESMA 度量表规范，适用于国企/政府项目审查
