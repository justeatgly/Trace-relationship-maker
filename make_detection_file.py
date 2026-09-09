# -*- coding: utf-8 -*-
"""
待检测文件生成器
=====================================================================

作用：根据“高级别需求 / 低级别需求 / 追溯结果文件”中的标识，把
      “追溯结果文件”中 LLM判定=“是” 的追溯配对，生成版式与
      “表格格式.xlsx”中 Sheet2 一致的“待检测文件”。

Sheet2 版式说明（本程序将复刻）：
  第 1 行：两栏大标题（左=高级别需求，右=低级别需求），各横向合并 6 列；
  第 2 行：左右各 6 列表头：标识 / 名称 / 描述 / 是否需求 / 是否派生 / 原理；
  第 3 行起：每一行代表一条“高级别—低级别”追溯配对；
            当同一高级别需求连续对应多条低级别需求时，
            左侧 6 列按行数纵向合并（与模板一致）。

填写规则（按你的约定）：
  1. “描述”列：取自源文件的“标题和正文”（若无此列，回退用“描述”）；
  2. “名称 / 是否派生 / 原理”：留空；
  3. “是否需求”：一律填 True。

使用方式：
  1. 修改下方【配置区】的输入/输出文件路径；
  2. 运行：
         python make_detection_file.py
     也可通过命令行临时指定文件（覆盖配置区）：
         python make_detection_file.py 高级别.xlsx 低级别.xlsx 追溯结果.xlsx 待检测.xlsx
"""

import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# =====================================================================
# 配置区（按需修改）
# =====================================================================

# 三个输入文件与输出文件路径（请改成你自己的真实路径）
HIGH_LEVEL_FILE = r"C:\Users\justeatgly\req_trace\sample_data\高级别需求.xlsx"
LOW_LEVEL_FILE = r"C:\Users\justeatgly\req_trace\sample_data\低级别需求.xlsx"
TRACE_RESULT_FILE = r"C:\Users\justeatgly\req_trace\sample_data\追溯结果文件.xlsx"
OUTPUT_FILE = r"C:\Users\justeatgly\req_trace\sample_data\待检测文件.xlsx"

# 列名（程序会忽略表头多余空格后精确匹配）
COL_ID = "标识"             # 高/低级别源文件中的需求标识列
COL_BODY_PREFERRED = "标题和正文"  # 源文件正文列（优先）
COL_BODY_FALLBACK = "描述"        # 若源文件无“标题和正文”，回退用此列
# 追溯结果文件相关列
COL_TRACE_HIGH = "高级别需求标识"
COL_TRACE_LOW = "低级别需求标识"
COL_TRACE_VERDICT = "LLM判定"
VERDICT_KEEP = "是"         # 只把判定为“是”的配对写入待检测文件

# Sheet2 版式参数
GROUP_TITLE_HIGH = "高级别需求"   # 若想与模板完全一致，可改为“软件高级别需求”
GROUP_TITLE_LOW = "低级别需求"     # 同理可改为“下游低级别需求”
HEADERS = ["标识", "名称", "描述", "是否需求", "是否派生", "原理"]
OUTPUT_SHEET_NAME = "Sheet1"

# =====================================================================
# 以下为程序主体
# =====================================================================


def normalize_header(df: pd.DataFrame) -> pd.DataFrame:
    """去除表头中的首尾空格，避免列名匹配失败。"""
    df.columns = [str(c).strip() if pd.notna(c) else c for c in df.columns]
    return df


def resolve_body_col(df: pd.DataFrame, label: str) -> str:
    """解析源文件的正文列：优先“标题和正文”，缺失时回退“描述”。"""
    if COL_BODY_PREFERRED in df.columns:
        return COL_BODY_PREFERRED
    if COL_BODY_FALLBACK in df.columns:
        return COL_BODY_FALLBACK
    raise KeyError(
        f"{label}文件缺少“{COL_BODY_PREFERRED}”或“{COL_BODY_FALLBACK}”列；"
        f"实际表头为：{list(df.columns)}"
    )


def cell_text(value) -> str:
    """单元格值转字符串；NaN/None 返回空串。"""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def load_body_lookup(file: Path, label: str) -> dict:
    """读取需求文件，返回 {标识: 正文文本} 映射（标识转为字符串）。"""
    df = normalize_header(pd.read_excel(file))
    missing = [c for c in (COL_ID,) if c not in df.columns]
    if missing:
        raise KeyError(f"{label}文件缺少列：{missing}；实际表头为：{list(df.columns)}")
    body_col = resolve_body_col(df, label)
    return {str(r[COL_ID]): cell_text(r[body_col]) for _, r in df.iterrows()}


def load_trace_pairs(file: Path) -> list:
    """读取追溯结果文件，返回 LLM判定=“是” 的 (高标识, 低标识) 有序列表。"""
    df = normalize_header(pd.read_excel(file))
    need = [COL_TRACE_HIGH, COL_TRACE_LOW, COL_TRACE_VERDICT]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise KeyError(
            f"追溯结果文件缺少列：{missing}；实际表头为：{list(df.columns)}"
        )
    ok = df[df[COL_TRACE_VERDICT].map(cell_text) == VERDICT_KEEP]
    pairs = []
    seen = set()
    for _, r in ok.iterrows():
        key = (cell_text(r[COL_TRACE_HIGH]), cell_text(r[COL_TRACE_LOW]))
        if key not in seen:      # 去重（保持出现顺序）
            seen.add(key)
            pairs.append(key)
    return pairs


# ---- 生成 Sheet2 版式 ----

def _set_cell(ws, row, col, value, border, wrap=False, bold=False, center=False):
    cell = ws.cell(row=row, column=col)
    cell.value = value
    cell.border = border
    if bold:
        cell.font = Font(bold=True)
    if wrap:
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    if center:
        cell.alignment = Alignment(horizontal="center", vertical="center")


def build_detection_file(pairs, high_lookup, low_lookup, missing_high, missing_low,
                         out_file: Path) -> int:
    """按 Sheet2 版式生成待检测文件，返回写入的配对条数。"""
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = Workbook()
    ws = wb.active
    ws.title = OUTPUT_SHEET_NAME

    # 第 1 行：两栏大标题（横向合并）
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    ws.merge_cells(start_row=1, start_column=7, end_row=1, end_column=12)
    title_fill = PatternFill("solid", fgColor="DDEBF7")
    for c in range(1, 13):
        cell = ws.cell(row=1, column=c)
        cell.border = border
        cell.fill = title_fill
        cell.font = Font(bold=True)
    ws.cell(row=1, column=1).value = GROUP_TITLE_HIGH
    ws.cell(row=1, column=7).value = GROUP_TITLE_LOW

    # 第 2 行：左右各 6 列表头
    for c, h in enumerate(HEADERS, start=1):
        _set_cell(ws, 2, c, h, border, bold=True, center=True)
        _set_cell(ws, 2, c + 6, h, border, bold=True, center=True)

    # 第 3 行起：按“同一高级别连续多条低级别”分组，写入并纵向合并左侧
    row = 3
    n = 0
    i = 0
    while i < len(pairs):
        high_id, _ = pairs[i]
        j = i
        while j < len(pairs) and pairs[j][0] == high_id:
            j += 1          # [i, j) 属于同一高级别的连续组
        group = pairs[i:j]
        high_body = high_lookup.get(str(high_id), "")
        if str(high_id) not in high_lookup:
            missing_high.add(str(high_id))
        for k, (_, low_id) in enumerate(group):
            low_body = low_lookup.get(str(low_id), "")
            if str(low_id) not in low_lookup:
                missing_low.add(str(low_id))
            r = row + k
            # 左侧：仅组首行填值，其余留空（等待纵向合并）
            left_vals = (
                [str(high_id), None, high_body, True, None, None]
                if k == 0
                else [None] * 6
            )
            right_vals = [str(low_id), None, low_body, True, None, None]
            for c, v in enumerate(left_vals, start=1):
                _set_cell(ws, r, c, v, border,
                          wrap=(c == 3), center=(c == 4))
            for c, v in enumerate(right_vals, start=7):
                _set_cell(ws, r, c, v, border,
                          wrap=(c == 9), center=(c == 10))
        # 纵向合并左侧（同一高级别占用多行时）
        if len(group) > 1:
            for c in range(1, 7):
                ws.merge_cells(start_row=row, start_column=c,
                               end_row=row + len(group) - 1, end_column=c)
        n += len(group)
        row += len(group)
        i = j

    # 列宽（正文列加宽并自动换行）
    widths = {1: 18, 2: 20, 3: 70, 4: 12, 5: 12, 6: 45,
              7: 18, 8: 20, 9: 70, 10: 12, 11: 12, 12: 45}
    for c, w in widths.items():
        ws.column_dimensions[chr(64 + c)].width = w
    ws.freeze_panes = "A3"

    wb.save(out_file)
    return n


def resolve_paths():
    """命令行参数（2~4 个）会覆盖配置区；否则使用配置区路径。"""
    if len(sys.argv) >= 3:
        high = sys.argv[1]
        low = sys.argv[2]
        trace = sys.argv[3] if len(sys.argv) >= 4 else TRACE_RESULT_FILE
        out = sys.argv[4] if len(sys.argv) >= 5 else OUTPUT_FILE
        return high, low, trace, out
    return HIGH_LEVEL_FILE, LOW_LEVEL_FILE, TRACE_RESULT_FILE, OUTPUT_FILE


def main():
    high_file, low_file, trace_file, out_file = resolve_paths()
    for f in (high_file, low_file, trace_file):
        if not Path(f).exists():
            raise FileNotFoundError(f"找不到文件：{f}")

    print(f"[读取] 高级别需求：{high_file}")
    print(f"[读取] 低级别需求：{low_file}")
    print(f"[读取] 追溯结果：{trace_file}")

    high_lookup = load_body_lookup(high_file, "高级别需求")
    low_lookup = load_body_lookup(low_file, "低级别需求")
    pairs = load_trace_pairs(trace_file)
    print(f"[筛选] 追溯结果中 LLM判定=“{VERDICT_KEEP}” 的配对 {len(pairs)} 条")

    if not pairs:
        # 仍生成带表头的空模板文件，便于直接查看格式
        build_detection_file([], high_lookup, low_lookup,
                             set(), set(), out_file)
        print(f"[输出] 无匹配配对，已生成空模板：{out_file}")
        return

    missing_high, missing_low = set(), set()
    n = build_detection_file(pairs, high_lookup, low_lookup,
                             missing_high, missing_low, out_file)
    if missing_high:
        print(f"[警告] 未在高级别需求中找到的标识：{sorted(missing_high)}（描述留空）")
    if missing_low:
        print(f"[警告] 未在低级别需求中找到的标识：{sorted(missing_low)}（描述留空）")
    print(f"[输出] 已生成待检测文件（{n} 条配对），写入：{out_file}")


if __name__ == "__main__":
    main()
