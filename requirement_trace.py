# -*- coding: utf-8 -*-
"""
需求追溯关系自动识别程序
=====================================================================

功能（三步）：
  第一步：分别筛选“是否需求”为 TRUE 的高级别需求和低级别需求；
  第二步：从高级别需求表格的“描述”列中提取变量名
          （同时包含大写字母与下划线（_）的组合，长度至少 2 个字符，
           例如 MAX_SPEED、POWER_1；纯大写字母如 MAX 或纯数字组合不提取，
           一条描述中可能有多个变量，全部提取），
          若低级别需求表格的“标题和正文”中能找到相同的变量名，
          则认定这两条需求之间存在追溯关系；
  第三步：将所有查询结果写入一个新的 Excel 文件，
          表头为：高级别需求标识、低级别需求标识。

使用方式：
  1. 安装依赖（只需一次）：
         pip install pandas openpyxl
  2. 修改下方【配置区】中的两个输入文件路径与输出文件路径；
  3. 运行：
         python requirement_trace.py
     也可以临时通过命令行参数指定文件（会覆盖配置区）：
         python requirement_trace.py 高级别.xlsx 低级别.xlsx 结果.xlsx
"""

import re
import sys
from pathlib import Path

import pandas as pd

# =====================================================================
# 配置区（按需修改）
# =====================================================================

# 两个输入文件与输出文件的路径，请改成你自己的真实路径
HIGH_LEVEL_FILE = r"C:\Users\justeatgly\req_trace\sample_data\高级别需求.xlsx"
LOW_LEVEL_FILE = r"C:\Users\justeatgly\req_trace\sample_data\低级别需求.xlsx"
OUTPUT_FILE = r"C:\Users\justeatgly\req_trace\sample_data\追溯结果.xlsx"

# 列名（程序会忽略表头多余空格后精确匹配）
COL_ID = "标识"          # 需求标识列
COL_IS_REQ = "是否需求"  # 是否需求列（值为 TRUE 的行视为“需求”）
COL_DESC = "描述"        # 高级别需求：从此列提取变量名
COL_BODY = "标题和正文"   # 低级别需求：在此列中查找相同变量名

# 被视为“需求”的值（不区分大小写；Excel 中的布尔 TRUE 也会被识别）
TRUE_VALUES = {"TRUE", "YES", "是", "1"}

# 变量名正则：大写字母开头，后跟大写字母 / 下划线 / 数字，总长度至少 2 个字符。
# 使用“非 [A-Z0-9_]”的环视，避免匹配到某个长单词/变量的中间片段。
# 注意：此处先宽松匹配全部大写风格标识符，是否“同时含大写字母与下划线”
#       的过滤在 extract_variables() 中完成。
VARIABLE_PATTERN = re.compile(r"(?<![A-Z0-9_])[A-Z][A-Z0-9_]{1,}(?![A-Z0-9_])")

# 匹配模式：
#   "any" —— 低级别需求只要包含高级别描述中的【任意一个】变量即算追溯关系（默认）
#   "all" —— 低级别需求必须包含高级别描述中的【全部】变量才算追溯关系
MATCH_MODE = "any"

# =====================================================================
# 以下为程序主体
# =====================================================================


def normalize_header(df: pd.DataFrame) -> pd.DataFrame:
    """去除表头中的首尾空格，避免因多余空格导致列名匹配失败。"""
    df.columns = [str(c).strip() if pd.notna(c) else c for c in df.columns]
    return df


def is_requirement(value) -> bool:
    """判断“是否需求”单元格是否为“需求”。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    if isinstance(value, bool):
        return value is True
    return str(value).strip().upper() in TRUE_VALUES


def extract_variables(text):
    """从一段文本中提取变量名，去重并保持出现顺序。

    只保留“同时包含大写字母与下划线（_）”的变量：
    正则已保证以大写字母开头（必含大写字母），
    因此这里再额外要求变量名中含有下划线即可。
    """
    if not isinstance(text, str):
        return []
    seen = set()
    result = []
    for m in VARIABLE_PATTERN.finditer(text):
        var = m.group()
        if "_" in var and any(c.isupper() for c in var):
            if var not in seen:
                seen.add(var)
                result.append(var)
    return result


def _var_in_text(var: str, text: str) -> bool:
    """判断变量名是否作为完整标识符出现在文本中（避免 MAX 命中 MAX_SPEED）。"""
    return re.search(
        r"(?<![A-Z0-9_])" + re.escape(var) + r"(?![A-Z0-9_])", text
    ) is not None


def is_traceable(high_vars, low_text) -> bool:
    """判断低级别需求文本是否与高级别变量构成追溯关系。"""
    if not high_vars or not isinstance(low_text, str):
        return False
    hits = [v for v in high_vars if _var_in_text(v, low_text)]
    if MATCH_MODE == "all":
        return len(hits) == len(high_vars)
    return len(hits) > 0


def resolve_paths():
    """命令行参数（3 个）会覆盖配置区；否则使用配置区路径。"""
    if len(sys.argv) >= 3:
        high, low = sys.argv[1], sys.argv[2]
        out = sys.argv[3] if len(sys.argv) >= 4 else "追溯结果.xlsx"
        return high, low, out
    return HIGH_LEVEL_FILE, LOW_LEVEL_FILE, OUTPUT_FILE


def main():
    high_file, low_file, out_file = resolve_paths()

    for f in (high_file, low_file):
        if not Path(f).exists():
            raise FileNotFoundError(f"找不到文件：{f}")

    # ---------- 读取 ----------
    high_df = normalize_header(pd.read_excel(high_file))
    low_df = normalize_header(pd.read_excel(low_file))

    # 校验必要列是否存在
    for name, df, need in (
        ("高级别", high_df, [COL_ID, COL_IS_REQ, COL_DESC]),
        ("低级别", low_df, [COL_ID, COL_IS_REQ, COL_BODY]),
    ):
        missing = [c for c in need if c not in df.columns]
        if missing:
            raise KeyError(
                f"{name}需求表缺少列：{missing}；实际表头为：{list(df.columns)}"
            )

    # ---------- 第一步：筛选“是否需求”为 TRUE 的行 ----------
    high_req = high_df[high_df[COL_IS_REQ].map(is_requirement)]
    low_req = low_df[low_df[COL_IS_REQ].map(is_requirement)]
    print(f"[第一步] 筛选完成：高级别需求 {len(high_req)} 条，低级别需求 {len(low_req)} 条")

    # ---------- 第二步：提取变量名并建立追溯关系 ----------
    records = []
    high_with_vars = 0
    for _, h in high_req.iterrows():
        h_id = h[COL_ID]
        vars_ = extract_variables(h[COL_DESC])
        if not vars_:
            continue  # “描述”中无变量，无法建立追溯
        high_with_vars += 1
        for _, l in low_req.iterrows():
            if is_traceable(vars_, l[COL_BODY]):
                records.append({"高级别需求标识": h_id, "低级别需求标识": l[COL_ID]})

    print(f"[第二步] 含变量的高级别需求 {high_with_vars} 条，进行追溯匹配…")

    # ---------- 第三步：去重后写入新 Excel ----------
    result = pd.DataFrame(records).drop_duplicates()
    result.to_excel(out_file, index=False)
    print(f"[第三步] 共发现 {len(result)} 条追溯关系，已写入：{out_file}")


if __name__ == "__main__":
    main()
