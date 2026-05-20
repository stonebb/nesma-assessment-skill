"""
Generate NESMA function point assessment Excel workbook.

Usage:
    python generate_excel.py --output <path.xlsx> --data <data.json>
    python generate_excel.py --output output.xlsx --project-name "项目名" --data data.json
"""

import argparse
import json
import sys
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, numbers
    )
    from openpyxl.utils import get_column_letter
except ImportError:
    print("请先安装 openpyxl: pip install openpyxl")
    sys.exit(1)

# ============================================================
# Color definitions (matching the template)
# ============================================================
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
LIGHT_GRAY_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
LIGHT_BLUE_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")

HEADER_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
NORMAL_FONT = Font(name="微软雅黑", size=10)
TITLE_FONT = Font(name="微软雅黑", size=12, bold=True)
BOLD_FONT = Font(name="微软雅黑", size=10, bold=True)
RED_FONT = Font(name="微软雅黑", size=10, color="FF0000")

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)

# NESMA category UFP lookup (估算法 default values)
CATEGORY_UFP = {"ILF": 10, "EIF": 7, "EI": 4, "EO": 5, "EQ": 4}

# Reuse adjustment
REUSE_RATIO = {"高": 0.33, "中": 0.67, "低": 1.0}

# Modify type adjustment
MODIFY_RATIO = {"新增": 1.0, "修改": 0.67, "删除": 0.0}

# Business domain factors
BUSINESS_DOMAIN = {
    "办公自动化系统": 1.0,
    "科学工程": 1.25,
    "多媒体": 1.3,
    "智能信息": 1.7,
    "系统": 1.8,
    "通信控制": 1.9,
    "过程控制": 2.0,
}

# Quality feature factors (each quality dimension gets -1, 0, or 1)
QUALITY_LEVELS = {
    "无": -1,
    "无特殊要求": -1,
    "有一定要求": 0,
    "有明确要求": 0,
    "严格要求": 1,
    "严格要求（高）": 1,
}

# Dev language factors
DEV_LANGUAGE = {
    "C等同等语言/平台": 1.5,
    "JAVA、C++、C#等同等语言/平台": 1.0,
    "Python等同等语言/平台": 0.8,
    "PowerBuilder、ASP等同等语言/平台": 0.6,
}

# Team background factors
TEAM_BACKGROUND = {
    "没有同类项目经验": 1.2,
    "有同类项目经验": 1.0,
    "为本行业开发过类似应用": 0.8,
}

# PDR reference
PDR_VALUES = {
    "智能信息": 10.48,
    "通用": 10.12,
}


def build_sheet_instructions(ws):
    """Sheet 1: 填写说明"""
    ws.title = "1. 填写说明"

    ws.merge_cells("A1:B1")
    ws["A1"] = (
        "模板填写顺序为：\n"
        "1、先填写功能点，填写《2.功能点测算》；\n"
        "2、再选择各调整因子参考《3.规模变更因子列表》取值；\n"
        "3、自动汇总计算后请检查《4.工作量测算》"
    )
    ws["A1"].font = NORMAL_FONT
    ws["A1"].alignment = LEFT_ALIGN

    ws["A4"] = "黄色单元格"
    ws["A4"].fill = YELLOW_FILL
    ws["B4"] = "推荐填写时只填写黄色单元格"
    ws["B4"].font = NORMAL_FONT

    ws["A5"] = "灰色单元格"
    ws["A5"].fill = LIGHT_GRAY_FILL
    ws["B5"] = "模板公式数字，不可修改"
    ws["B5"].font = NORMAL_FONT

    ws["A6"] = "浅灰色单元格"
    ws["A6"].fill = LIGHT_BLUE_FILL
    ws["B6"] = "模板默认数字，默认为行业基准推荐取值，各省可按实际填写，也可按实际情况据实调优，修改须经省公司沟通确认或总部审核同意"
    ws["B6"].font = NORMAL_FONT

    ws["A7"] = "白色单元格"
    ws["B7"] = "自动计算汇总数据，不可修改"
    ws["B7"].font = NORMAL_FONT

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 60


def build_sheet_function_points(ws, data):
    """Sheet 2: 功能点测算"""
    ws.title = "2. 功能点测算"

    functions = data.get("functions", [])
    method = data.get("method", "估算法")

    # Set column widths
    col_widths = {
        "A": 6, "B": 22, "C": 22, "D": 40, "E": 8,
        "F": 10, "G": 10, "H": 10, "I": 10, "J": 15, "K": 8
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Title
    ws.merge_cells("A1:K1")
    ws["A1"] = "1.功能点测算"
    ws["A1"].font = TITLE_FONT

    # Method selector
    ws["A3"] = "功能点测算方法"
    ws["A3"].font = BOLD_FONT
    ws["C3"] = f"{method}（推荐模板）"
    ws["C3"].font = BOLD_FONT
    ws["D3"] = '推荐使用"7/5/4/5/4"估算法'
    ws["D3"].font = NORMAL_FONT

    # Totals
    last_data_row = 7 + len(functions)
    ws["A4"] = "未调整功能点合计"
    ws["A4"].font = BOLD_FONT
    ws["C4"] = f"=SUM(F8:F{last_data_row})"
    ws["C4"].font = BOLD_FONT
    ws["D4"] = "UFP,单位:FP"

    ws["A5"] = "调整后功能点合计"
    ws["A5"].font = BOLD_FONT
    ws["C5"] = f"=SUM(I8:I{last_data_row})"
    ws["C5"].font = BOLD_FONT
    ws["D5"] = "AFP,单位:FP"

    ws["A6"] = "1. 依据项目《XXX需求说明书》等文件进行测算\n2. 本次测算为项目前期，需求未完全明确"
    ws["A6"].font = NORMAL_FONT
    ws["A6"].alignment = LEFT_ALIGN

    # Header row
    headers = ["序号", "一级功能", "二级功能", "对应功能名称/需求",
               "类别", "UFP", "复用程度", "修改类型", "AFP", "备注", ""]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=7, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER

    # Sub-header (说明行)
    sub_headers = ["填写\n说明", "", "", "功能点名称描述",
                   "选择推荐模板类别。", "自动计算未调整的功能点数量",
                   "选择输入复用程度", "选择输入修改类型",
                   "自动计算调整后的功能点数量", "", ""]
    for col_idx, sh in enumerate(sub_headers, 1):
        cell = ws.cell(row=8, column=col_idx, value=sh)
        cell.font = Font(name="微软雅黑", size=8, color="666666")
        cell.fill = LIGHT_BLUE_FILL
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER

    # Data rows
    for i, func in enumerate(functions):
        row = 9 + i
        seq = i + 1

        ws.cell(row=row, column=1, value=seq).font = NORMAL_FONT
        ws.cell(row=row, column=2, value=func.get("level1", "")).font = NORMAL_FONT
        ws.cell(row=row, column=3, value=func.get("level2", "")).font = NORMAL_FONT
        ws.cell(row=row, column=4, value=func.get("name", "")).font = NORMAL_FONT

        category = func.get("category", "")
        cat_cell = ws.cell(row=row, column=5, value=category)
        cat_cell.font = NORMAL_FONT
        cat_cell.fill = YELLOW_FILL
        cat_cell.alignment = CENTER_ALIGN

        # UFP formula: lookup category in a hidden table, or use default
        ufp_val = CATEGORY_UFP.get(category, 0)
        ufp_cell = ws.cell(row=row, column=6, value=ufp_val)
        ufp_cell.font = NORMAL_FONT
        ufp_cell.fill = LIGHT_GRAY_FILL
        ufp_cell.alignment = CENTER_ALIGN

        # Reuse degree
        reuse = func.get("reuse", "低")
        reuse_cell = ws.cell(row=row, column=7, value=reuse)
        reuse_cell.font = NORMAL_FONT
        reuse_cell.fill = YELLOW_FILL
        reuse_cell.alignment = CENTER_ALIGN

        # Modify type
        modify = func.get("modify_type", "新增")
        modify_cell = ws.cell(row=row, column=8, value=modify)
        modify_cell.font = NORMAL_FONT
        modify_cell.fill = YELLOW_FILL
        modify_cell.alignment = CENTER_ALIGN

        # AFP = UFP * reuse_ratio * modify_ratio
        reuse_val = REUSE_RATIO.get(reuse, 1.0)
        modify_val = MODIFY_RATIO.get(modify, 1.0)
        afp_val = round(ufp_val * reuse_val * modify_val, 1)
        afp_cell = ws.cell(row=row, column=9, value=afp_val)
        afp_cell.font = NORMAL_FONT
        afp_cell.fill = WHITE_FILL
        afp_cell.alignment = CENTER_ALIGN

        ws.cell(row=row, column=10, value=func.get("notes", "")).font = NORMAL_FONT

        # Apply borders
        for col in range(1, 12):
            ws.cell(row=row, column=col).border = THIN_BORDER
            ws.cell(row=row, column=col).alignment = CENTER_ALIGN

    # Freeze panes
    ws.freeze_panes = "A9"


def build_sheet_adjustment_factors(ws, data):
    """Sheet 3: 规模变更因子列表"""
    ws.title = "3. 规模变更因子列表"
    project_info = data.get("project_info", {})

    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 15

    ws["A1"] = "规模变更因子列表"
    ws["A1"].font = TITLE_FONT

    # Headers
    for col, h in enumerate(["类别", "子项", "判断标准", "调整因子"], 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER

    row = 4

    # Business domain
    bd = project_info.get("business_domain", "")
    bd_val = BUSINESS_DOMAIN.get(bd, 1.0)
    ws.cell(row=row, column=1, value="业务领域").font = BOLD_FONT
    ws.cell(row=row, column=2, value="").font = NORMAL_FONT
    ws.cell(row=row, column=3, value=bd).font = NORMAL_FONT
    ws.cell(row=row, column=3).fill = YELLOW_FILL
    ws.cell(row=row, column=4, value=bd_val).font = NORMAL_FONT
    ws.cell(row=row, column=4).fill = LIGHT_GRAY_FILL
    for c in range(1, 5):
        ws.cell(row=row, column=c).border = THIN_BORDER
    row += 1

    # Application type (from the same domain categories but different dimension)
    at = project_info.get("app_type", "")
    at_val = BUSINESS_DOMAIN.get(at, 1.0)
    ws.cell(row=row, column=1, value="应用类型").font = BOLD_FONT
    ws.cell(row=row, column=2, value="").font = NORMAL_FONT
    ws.cell(row=row, column=3, value=at).font = NORMAL_FONT
    ws.cell(row=row, column=3).fill = YELLOW_FILL
    ws.cell(row=row, column=4, value=at_val).font = NORMAL_FONT
    ws.cell(row=row, column=4).fill = LIGHT_GRAY_FILL
    for c in range(1, 5):
        ws.cell(row=row, column=c).border = THIN_BORDER
    row += 1

    # Quality features
    quality = project_info.get("quality", {})
    quality_map = {
        "分布式处理": "无分布式处理要求",
        "性能": "无特殊要求",
        "可靠性": "无特殊要求",
        "多站点": "无特殊要求",
    }
    # Update with actual selections
    for k, v in quality.items():
        if v in QUALITY_LEVELS:
            quality_map[k] = v

    q_labels = ["分布式处理", "性能", "可靠性", "多站点"]
    q_total = 0
    for i, ql in enumerate(q_labels):
        qv = quality.get(ql, "无特殊要求")
        q_val = QUALITY_LEVELS.get(qv, -1)
        q_total += q_val
        label = "质量特征" if i == 0 else ""
        ws.cell(row=row, column=1, value=label).font = BOLD_FONT
        ws.cell(row=row, column=2, value=ql).font = NORMAL_FONT
        ws.cell(row=row, column=3, value=qv).font = NORMAL_FONT
        ws.cell(row=row, column=3).fill = YELLOW_FILL
        ws.cell(row=row, column=4, value=q_val).font = NORMAL_FONT
        ws.cell(row=row, column=4).fill = LIGHT_GRAY_FILL
        for c in range(1, 5):
            ws.cell(row=row, column=c).border = THIN_BORDER
        row += 1

    # Dev language
    dl = project_info.get("dev_language", "")
    dl_val = DEV_LANGUAGE.get(dl, 1.0)
    ws.cell(row=row, column=1, value="开发语言").font = BOLD_FONT
    ws.cell(row=row, column=2, value="").font = NORMAL_FONT
    ws.cell(row=row, column=3, value=dl).font = NORMAL_FONT
    ws.cell(row=row, column=3).fill = YELLOW_FILL
    ws.cell(row=row, column=4, value=dl_val).font = NORMAL_FONT
    ws.cell(row=row, column=4).fill = LIGHT_GRAY_FILL
    for c in range(1, 5):
        ws.cell(row=row, column=c).border = THIN_BORDER
    row += 1

    # Team background
    tb = project_info.get("team_background", "")
    tb_val = TEAM_BACKGROUND.get(tb, 1.0)
    ws.cell(row=row, column=1, value="开发团队背景").font = BOLD_FONT
    ws.cell(row=row, column=2, value="").font = NORMAL_FONT
    ws.cell(row=row, column=3, value=tb).font = NORMAL_FONT
    ws.cell(row=row, column=3).fill = YELLOW_FILL
    ws.cell(row=row, column=4, value=tb_val).font = NORMAL_FONT
    ws.cell(row=row, column=4).fill = LIGHT_GRAY_FILL
    for c in range(1, 5):
        ws.cell(row=row, column=c).border = THIN_BORDER
    row += 1

    # Store key values for Sheet 4 formulas
    ws.cell(row=row + 1, column=1, value="CF = 业务领域 × 应用类型 × (1 + 0.025 × Σ质量) × 语言 × 团队")
    ws.cell(row=row + 1, column=1).font = Font(name="微软雅黑", size=9, italic=True)

    # Build CF value
    at_val = BUSINESS_DOMAIN.get(at, 1.0)
    cf_val = bd_val * at_val * (1 + 0.025 * q_total) * dl_val * tb_val
    ws.cell(row=row + 2, column=1, value="规模变更因子(CF)")
    ws.cell(row=row + 2, column=4, value=round(cf_val, 4))
    ws.cell(row=row + 2, column=4).font = BOLD_FONT

    # PDR selection
    pdr = PDR_VALUES.get("智能信息", 10.12) if bd in ("智能信息", "科学工程") else PDR_VALUES.get("通用", 10.12)
    ws.cell(row=row + 3, column=1, value="基准生产率(PDR)")
    ws.cell(row=row + 3, column=4, value=pdr)


def build_sheet_workload(ws, data):
    """Sheet 4: 工作量测算"""
    ws.title = "4. 工作量测算"
    project_info = data.get("project_info", {})
    functions = data.get("functions", [])
    last_data_row = 7 + len(functions)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 40

    ws["A1"] = "工作量测算"
    ws["A1"].font = TITLE_FONT

    row = 3
    headers = ["项目", "", "数值", "说明"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
    row += 1

    # AFP
    ws.cell(row=row, column=1, value="调整后功能点（AFP）").font = NORMAL_FONT
    ws.cell(row=row, column=3, value=f"='2. 功能点测算'!C5").font = BOLD_FONT
    ws.cell(row=row, column=4, value="来自功能点测算表，单位:FP").font = Font(name="微软雅黑", size=9)
    for c in range(1, 5):
        ws.cell(row=row, column=c).border = THIN_BORDER
    row += 1

    # Calculate CF
    bd = project_info.get("business_domain", "")
    at = project_info.get("app_type", "")
    dl = project_info.get("dev_language", "")
    tb = project_info.get("team_background", "")
    quality = project_info.get("quality", {})

    bd_val = BUSINESS_DOMAIN.get(bd, 1.0)
    at_val = BUSINESS_DOMAIN.get(at, 1.0)
    dl_val = DEV_LANGUAGE.get(dl, 1.0)
    tb_val = TEAM_BACKGROUND.get(tb, 1.0)

    q_total = sum(QUALITY_LEVELS.get(quality.get(q, "无特殊要求"), -1)
                  for q in ["分布式处理", "性能", "可靠性", "多站点"])

    cf_val = bd_val * at_val * (1 + 0.025 * q_total) * dl_val * tb_val

    ws.cell(row=row, column=1, value="规模变更因子（CF）").font = NORMAL_FONT
    ws.cell(row=row, column=3, value=round(cf_val, 4)).font = BOLD_FONT
    ws.cell(row=row, column=3).fill = LIGHT_GRAY_FILL
    ws.cell(row=row, column=4, value="来自规模变更因子列表").font = Font(name="微软雅黑", size=9)
    for c in range(1, 5):
        ws.cell(row=row, column=c).border = THIN_BORDER
    row += 1

    # S = AFP * CF
    ws.cell(row=row, column=1, value="调整后规模（S）").font = NORMAL_FONT
    ws.cell(row=row, column=3, value=f"=C{row-2}*C{row-1}").font = BOLD_FONT
    ws.cell(row=row, column=4, value="S = AFP × CF, 单位:FP").font = Font(name="微软雅黑", size=9)
    for c in range(1, 5):
        ws.cell(row=row, column=c).border = THIN_BORDER
    row += 1

    # PDR
    pdr = PDR_VALUES.get("智能信息", 10.12) if bd in ("智能信息", "科学工程") else PDR_VALUES.get("通用", 10.12)
    ws.cell(row=row, column=1, value="基准生产率（PDR）").font = NORMAL_FONT
    ws.cell(row=row, column=3, value=pdr).font = BOLD_FONT
    ws.cell(row=row, column=3).fill = LIGHT_BLUE_FILL
    ws.cell(row=row, column=4, value=f'PDR，参考《中国软件行业基准数据（SSM-BK-202409）》，单位:人时/FP')
    ws.cell(row=row, column=4).font = Font(name="微软雅黑", size=9)
    for c in range(1, 5):
        ws.cell(row=row, column=c).border = THIN_BORDER
    row += 1

    # AE = S * PDR / 8
    ws.cell(row=row, column=1, value="工作量（AE）").font = BOLD_FONT
    ws.cell(row=row, column=3, value=f"=C{row-2}*C{row-1}/8").font = Font(name="微软雅黑", size=12, bold=True)
    ws.cell(row=row, column=4, value="AE = S × PDR / 8, 单位:人天").font = Font(name="微软雅黑", size=9)
    for c in range(1, 5):
        ws.cell(row=row, column=c).border = THIN_BORDER
    row += 1

    # Notes
    row += 1
    ws.cell(row=row, column=1, value="说明：").font = BOLD_FONT
    ws.cell(row=row + 1, column=1, value="1. 基准生产率（PDR）参考中国软件行业基准数据（SSM-BK-202409）").font = Font(name="微软雅黑", size=9)
    ws.cell(row=row + 2, column=1, value="2. 本测算不包含直接非人力成本").font = Font(name="微软雅黑", size=9)
    ws.cell(row=row + 3, column=1, value="3. 本测算适用于项目立项阶段，需求尚未完全明确").font = Font(name="微软雅黑", size=9)


def generate_excel(output_path, data):
    """Generate complete NESMA Excel workbook."""
    wb = Workbook()

    # Sheet 1
    ws1 = wb.active
    build_sheet_instructions(ws1)

    # Sheet 2
    ws2 = wb.create_sheet()
    build_sheet_function_points(ws2, data)

    # Sheet 3
    ws3 = wb.create_sheet()
    build_sheet_adjustment_factors(ws3, data)

    # Sheet 4
    ws4 = wb.create_sheet()
    build_sheet_workload(ws4, data)

    # Save
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f"Excel workbook saved to: {output_path}")

    # Print summary
    functions = data.get("functions", [])
    total_ufp = sum(CATEGORY_UFP.get(f.get("category", ""), 0) for f in functions)
    total_afp = sum(
        CATEGORY_UFP.get(f.get("category", ""), 0)
        * REUSE_RATIO.get(f.get("reuse", "低"), 1.0)
        * MODIFY_RATIO.get(f.get("modify_type", "新增"), 1.0)
        for f in functions
    )
    print(f"  功能点总数: {len(functions)}")
    print(f"  UFP 合计: {total_ufp} FP")
    print(f"  AFP 合计: {round(total_afp, 1)} FP")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate NESMA function point Excel")
    parser.add_argument("--output", required=True, help="Output .xlsx file path")
    parser.add_argument("--data", required=True, help="JSON data file path")
    parser.add_argument("--project-name", help="Project name (optional, can be in data)")
    args = parser.parse_args()

    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    if args.project_name:
        data["project_name"] = args.project_name

    generate_excel(args.output, data)
