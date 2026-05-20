# NESMA 功能点评估技能 (nesma-assessment)

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Claude%20Code%20%7C%20Cline%20%7C%20Cursor%20%7C%20Copilot-green)]()
[![Standard](https://img.shields.io/badge/standard-GB%2FT%2042588--2023-orange)]()

基于 **GB/T 42588-2023**（ISO/IEC 24570:2018 MOD）国家标准的 NESMA 功能点评估 Agent 技能，帮助软件项目自动生成功能点拆分 Excel 和需求规格说明书 DOCX。

## 快速安装

```bash
npx skills add stonebb/nesma-assessment-skill
```

支持平台：Claude Code、Cline、Cursor、GitHub Copilot、CodeBuddy、Codex、Continue、Kimi Code CLI、Trae CN 等 55+ 个 Agent。

## 功能概览

```
用户描述项目需求
      │
      ▼
┌─────────────┐    ┌──────────────┐    ┌──────────────┐
│  阶段一      │ →  │  阶段二       │ →  │  阶段三       │
│  需求采集    │    │  生成 Excel   │    │  生成 DOCX    │
│             │    │             │    │             │
│ · 项目信息   │    │ · 填写说明    │    │ · 项目信息表  │
│ · 功能分解   │    │ · 功能点测算  │    │ · 功能需求    │
│ · NESMA归类 │    │ · 变更因子    │    │ · 非功能需求  │
│             │    │ · 工作量测算  │    │ · 架构概述    │
└─────────────┘    └──────────────┘    └──────────────┘
```

## NESMA 方法简介

NESMA（Netherlands Software Measurement Association，荷兰软件度量协会）功能点分析方法是中国国家标准 GB/T 42588-2023 采用的软件规模度量方法，广泛用于政府和企业信息化项目的**预算编制**、**招投标**和**工作量评估**。

### 三种评估方法

| 方法 | 适用阶段 | 识别范围 | 精度 |
|------|----------|----------|------|
| **预估法** (Indicative) | 早期/招投标，需求模糊 | 仅 ILF + EIF | 粗略 |
| **估算法** (Estimated) | 需求中等明确 | ILF + EIF + EI + EO + EQ | 中等 |
| **详算法** (Detailed) | 需求完全明确 | 全部组件，按复杂度查表 | 精确 |

### 五种功能组件

| 代码 | 类型 | 说明 | 示例 |
|------|------|------|------|
| ILF | 内部逻辑文件 | 本系统维护的数据 | 数据库表、配置库 |
| EIF | 外部接口文件 | 外部系统维护的数据 | 第三方 API 数据 |
| EI | 外部输入 | 数据的增删改操作 | 表单提交、数据导入 |
| EO | 外部输出 | 含计算逻辑的输出 | 分析报表、统计图 |
| EQ | 外部查询 | 无计算逻辑的查询 | 列表查询、详情查看 |

### 关键公式

```
UFP = Σ(各功能点按类别的功能点数)
AFP = Σ(UFP × 复用程度系数 × 修改类型系数)
CF  = 业务领域 × 应用类型 × (1 + 0.025 × Σ质量特征) × 开发语言 × 团队背景
S   = AFP × CF
AE  = S × PDR / 8      (人天)
```

## 生成的 Excel 结构

| Sheet | 名称 | 内容 |
|-------|------|------|
| 1 | 填写说明 | 颜色标注说明、填写顺序指引 |
| 2 | 功能点测算 | 功能点主表（序号/一级功能/二级功能/名称/类别/UFP/复用程度/修改类型/AFP） |
| 3 | 规模变更因子列表 | 业务领域、应用类型、质量特征（分布式/性能/可靠性/多站点）、开发语言、团队背景 |
| 4 | 工作量测算 | AFP → CF → S → PDR → AE 完整计算链 |

## 生成的 DOCX 结构

```
1. 背景
   1.1 基本信息
   1.2 目标
   1.3 适用范围
   1.4 参考文献
2. 功能需求（按模块逐级展开）
3. 非功能需求
   3.1 性能需求
   3.2 可靠性需求
   3.3 安全性需求
4. 架构概述
5. 附录
```

## JSON 数据格式

```json
{
  "project_name": "XX项目",
  "project_type": "新建项目",
  "project_info": {
    "business_domain": "智能信息",
    "app_type": "科学工程",
    "quality": {
      "分布式处理": "无特殊要求",
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
      "level1": "用户管理",
      "level2": "登录注册",
      "name": "用户通过手机号+验证码登录系统",
      "category": "EI",
      "reuse": "高",
      "modify_type": "新增"
    }
  ]
}
```

### 字段说明

| 字段 | 取值 |
|------|------|
| `category` | ILF / EIF / EI / EO / EQ |
| `reuse` | 高(0.33) / 中(0.67) / 低(1.0) |
| `modify_type` | 新增 / 修改 / 删除 |
| `business_domain` / `app_type` | 办公自动化系统 / 科学工程 / 多媒体 / 智能信息 / 系统 / 通信控制 / 过程控制 |
| `quality.*` | 无特殊要求 / 有一定要求 / 严格要求 |
| `dev_language` | C等 / JAVA、C++、C#等 / Python等 / PowerBuilder、ASP等 |
| `team_background` | 无同类项目经验 / 有同类项目经验 / 为本行业开发过类似应用 |

## 环境依赖

- Python 3.8+
- openpyxl（Excel 生成）
- python-docx（DOCX 生成）

```bash
pip install openpyxl python-docx
```

## 使用示例

安装后，在 Agent 中直接对话即可触发：

```
> 帮我为"智慧园区管理平台"项目做NESMA功能点评估

Agent 会自动引导你：
1. 采集项目基本信息（业务领域、开发语言等）
2. 逐模块识别功能点和NESMA类别
3. 生成功能点拆分 Excel 和需求规格说明书 DOCX
```

也可手动调用脚本：

```bash
# 准备 JSON 数据文件
# 生成 Excel
python scripts/generate_excel.py --output 功能点评估.xlsx --data project.json

# 生成 DOCX
python scripts/generate_docx.py --output 需求规格说明书.docx --data project.json
```

## 参考标准

- **GB/T 42588-2023** 系统与软件工程 功能规模测量 NESMA方法
- **ISO/IEC 24570:2018** Software engineering — NESMA functional size measurement method
- **GB/T 36964-2018** 软件工程 软件开发成本度量规范
- **中国软件行业基准数据（SSM-BK）** — 基准生产率（PDR）参考来源

## 项目结构

```
nesma-assessment-skill/
├── SKILL.md                     # 技能主文件（ Agent instructions）
├── README.md                    # 本文件
├── references/
│   └── nesma_rules.md           # NESMA 规则参考（组件判定、赋值表、因子、公式）
├── scripts/
│   ├── generate_excel.py        # Excel 生成脚本（4 Sheet，含 NESMA 公式）
│   └── generate_docx.py         # DOCX 生成脚本（需求规格说明书）
└── assets/                      # 静态资源
```

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request。主要改进方向：
- 支持详算法（基于 DET/RET/FTR 的精确复杂度判定）
- 集成 PDR 基准数据自动查询
- 支持更多行业模板变体
