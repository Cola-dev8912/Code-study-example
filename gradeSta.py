# -*- coding: utf-8 -*-
"""
统计工作表中「优、良、中、次、差」的个数及各自所占比例。
适用于路面损坏状况评定明细表中的等级列。
"""

import os
import sys

try:
    import openpyxl
except ImportError:
    print("请先安装 openpyxl: pip install openpyxl")
    sys.exit(1)


# 要统计的等级字符
GRADE_CHARS = ("优", "良", "中", "次", "差")
# 支持的 Excel 扩展名（openpyxl 可读取 .xlsx 与 .xlsm）
EXCEL_EXTENSIONS = (".xlsx", ".xlsm")
# 要筛选的工作表指标（需同时满足：含指标名 且 含"百"）
INDICATORS = (
    "PQI", "PCI", "RQI", "RDI", "PBI", "SRI", "PSSI"
)


def count_grades_in_sheet(ws):
    """遍历工作表中所有单元格，统计等级字符出现次数。"""
    counts = {c: 0 for c in GRADE_CHARS}
    total = 0

    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            val = str(cell.value).strip()
            if not val:
                continue
            # 只统计单字符或单元格整体为等级字的情况
            for char in val:
                if char in GRADE_CHARS:
                    counts[char] += 1
                    total += 1

    return counts, total


def _resolve_excel_path(path):
    """若指定路径不存在，尝试同名但另一扩展名（.xlsx / .xlsm）。"""
    if os.path.isfile(path):
        return path
    base, ext = os.path.splitext(path)
    for other in EXCEL_EXTENSIONS:
        if other != ext.lower():
            alt = base + other
            if os.path.isfile(alt):
                return alt
    return None


def _is_target_sheet(name):
    """
    判断工作表名是否需要统计。
    条件：名字中含有 INDICATORS 之一，且名字中含有"百"。
    例如："PCI百上" ✔，"RDI下" ✘，"PCI千下" ✘，"DCB百上二" ✘
    """
    name_upper = name.upper()
    has_indicator = any(ind in name_upper for ind in INDICATORS)
    has_bai = "百" in name
    return has_indicator and has_bai


def _extract_indicator(name):
    """从工作表名中提取指标名（如 "PQI百上" -> "PQI"）"""
    name_upper = name.upper()
    for ind in INDICATORS:
        if ind in name_upper:
            return ind
    return None


def _get_file_path():
    """获取要处理的 Excel 文件路径。"""
    # 优先使用命令行参数
    if len(sys.argv) >= 2:
        path = sys.argv[1]
        resolved = _resolve_excel_path(path)
        if resolved and resolved != path:
            print(f"未找到 {path}，使用同名的 {resolved}\n")
            return resolved
        return path

    # 自动查找当前目录下的 Excel 文件
    candidates = [
        f for f in os.listdir(".")
        if f.lower().endswith(EXCEL_EXTENSIONS) and not f.startswith("~$")
    ]
    if not candidates:
        print("当前目录下未找到 .xlsx / .xlsm 文件")
        sys.exit(1)

    print(f"使用文件: {candidates[0]}\n")
    return candidates[0]


def _load_sheets(path):
    """加载 Excel 文件并筛选目标工作表。"""
    if not os.path.isfile(path):
        print(f"文件不存在: {path}")
        sys.exit(1)

    wb = openpyxl.load_workbook(path, data_only=True)
    sheets = [ws for ws in wb.worksheets if _is_target_sheet(ws.title)]

    if not sheets:
        print("未找到符合条件的工作表")
        wb.close()
        sys.exit(0)

    return wb, sheets


def _collect_data(sheets):
    """收集各指标的统计数据。"""
    data = {}
    for ws in sheets:
        ind = _extract_indicator(ws.title)
        if not ind:
            continue

        counts, total = count_grades_in_sheet(ws)
        if ind not in data:
            data[ind] = {"counts": {c: 0 for c in GRADE_CHARS}, "total": 0, "sheets": []}

        for c in GRADE_CHARS:
            data[ind]["counts"][c] += counts[c]
        data[ind]["total"] += total
        data[ind]["sheets"].append(ws.title)

    return data


def _create_output(path, data):
    """生成统计结果的 Excel 文件。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    wb_out = Workbook()
    ws = wb_out.active
    ws.title = "统计结果"

    # 样式定义
    header_fill = PatternFill("solid", "4472C4")
    indicator_fill = PatternFill("solid", "D9E1F2")
    thin = Side("thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment("center", "center")

    # 表头
    headers = ["指标", "工作表", "优", "良", "中", "次", "差", "合计", "优%", "良%", "中%", "次%", "差%"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.font = Font(bold=True, size=12, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    # 列宽
    for col, w in enumerate([10, 30, 8, 8, 8, 8, 8, 10, 8, 8, 8, 8, 8], 1):
        ws.column_dimensions[chr(64 + col)].width = w

    # 数据行
    for row, ind in enumerate(filter(lambda x: x in data, INDICATORS), 2):
        d = data[ind]
        counts, total = d["counts"], d["total"]
        pct = {c: counts[c] / total * 100 if total else 0 for c in GRADE_CHARS}

        ws.cell(row, 1, ind).fill = indicator_fill
        ws.cell(row, 2, " + ".join(d["sheets"]))

        for i, grade in enumerate(GRADE_CHARS, 3):
            ws.cell(row, i, counts[grade])
            ws.cell(row, i + 5, f"{pct[grade]:.1f}%")

        ws.cell(row, 8, total)

        for col in range(1, 14):
            ws.cell(row, col).border = border
            ws.cell(row, col).alignment = center

    wb_out.save(path)
    print(f"统计完成！结果已保存至: {path}")


def main():
    """主函数：协调整个统计流程。"""
    path = _get_file_path()
    wb, sheets = _load_sheets(path)
    data = _collect_data(sheets)

    output_path = path.replace(".xlsx", "_统计结果.xlsx").replace(".xlsm", "_统计结果.xlsx")
    _create_output(output_path, data)

    wb.close()


if __name__ == "__main__":
    main()
