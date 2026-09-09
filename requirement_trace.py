# -*- coding: utf-8 -*-
"""
需求追溯关系自动识别程序（关键词初筛 + 大语言模型复核）
=====================================================================

功能（三步）：
  第一步：分别筛选“是否需求”为 TRUE 的高级别需求和低级别需求；
  第二步（初筛）：从高级别需求“描述”列中提取变量名（同时包含大写字母与下划线、
          长度至少 2 个字符，如 MAX_SPEED、POWER_1，一条描述可提取多个），
          若低级别需求的正文（优先“标题和正文”，该列缺失时回退用“描述”）中出现
          相同变量名，则记为一条候选追溯关系，去重后写入“追溯结果中间产物.xlsx”
          （表头：高级别需求标识、低级别需求标识）；
  第三步（复核与输出）：读取“追溯结果中间产物”，按高/低级别需求标识从原始需求表
          取回双方描述（高级别取“描述”；低级别优先取“标题和正文”，该列缺失时取“描述”）
          拼入提示词，分批调用 DeepSeek 大语言模型判定每条候选是否真实存在追溯关系，
          将中间产物每一行与其 LLM 判定（是/否/失败、判定理由）合并，
          写入“追溯结果文件.xlsx”。

使用方式：
  1. 安装依赖（只需一次）：
         pip install pandas openpyxl
  2. 在下方【配置区】填写/核对：
         - 输入文件、中间产物、最终结果文件的路径；
         - DeepSeek API Key（LLM_API_KEY，或设置环境变量 DEEPSEEK_API_KEY）；
  3. 运行：
         python requirement_trace.py
     也可以临时通过命令行参数指定文件（会覆盖配置区）：
         python requirement_trace.py 高级别.xlsx 低级别.xlsx 中间产物.xlsx 最终结果.xlsx
"""

import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

import pandas as pd

# =====================================================================
# 配置区（按需修改）
# =====================================================================

# 两个输入文件与输出文件的路径，请改成你自己的真实路径
HIGH_LEVEL_FILE = r"C:\Users\justeatgly\req_trace\sample_data\高级别需求.xlsx"
LOW_LEVEL_FILE = r"C:\Users\justeatgly\req_trace\sample_data\低级别需求.xlsx"
# 关键词初筛后得到的“候选追溯关系”中间产物
INTERMEDIATE_FILE = r"C:\Users\justeatgly\req_trace\sample_data\追溯结果中间产物.xlsx"
# 大模型复核后写入的“最终追溯结果”
FINAL_FILE = r"C:\Users\justeatgly\req_trace\sample_data\追溯结果文件.xlsx"

# 列名（程序会忽略表头多余空格后精确匹配）
COL_ID = "标识"          # 需求标识列
COL_IS_REQ = "是否需求"  # 是否需求列（值为 TRUE 的行视为“需求”）
COL_DESC = "描述"        # 高级别需求：从此列提取变量名
COL_BODY = "标题和正文"   # 低级别需求：在此列中查找相同变量名；若低级别表无此列，自动回退用“描述”列

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
# 大语言模型复核配置（DeepSeek，OpenAI 兼容接口）
# =====================================================================

ENABLE_LLM_REVIEW = True   # False：只生成“追溯结果中间产物”，跳过第二步大模型复核
LLM_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "").strip()  # 若环境变量未设置，请直接在此填入密钥
LLM_BASE_URL = "https://api.deepseek.com"   # OpenAI 兼容接口地址
LLM_MODEL = "deepseek-chat"                 # DeepSeek 模型名
LLM_TIMEOUT = 90             # 单次请求超时（秒）
BATCH_SIZE = 10              # 每次请求合并判定的候选条数（分批合并调用，更省 token）
MAX_RETRY = 2                # 单个批次请求/解析失败后的重试次数
REVIEW_SLEEP = 0.3           # 批次之间与失败重试之间的间隔（秒），避免触发限流

# =====================================================================
# 以下为程序主体
# =====================================================================


def normalize_header(df: pd.DataFrame) -> pd.DataFrame:
    """去除表头中的首尾空格，避免因多余空格导致列名匹配失败。"""
    df.columns = [str(c).strip() if pd.notna(c) else c for c in df.columns]
    return df


def resolve_low_body_column(low_df: pd.DataFrame):
    """解析低级别需求表的“正文来源”列名。

    优先使用“标题和正文”列；若该列不存在，则回退到“描述”列；
    两者都不存在时返回 None（由调用方报错）。
    """
    if COL_BODY in low_df.columns:
        return COL_BODY
    if COL_DESC in low_df.columns:
        return COL_DESC
    return None


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


def _text_of_cell(value) -> str:
    """把单元格内容转为去首尾空白的字符串；NaN / None 返回空串。"""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


# =====================================================================
# 第三步辅助：调用大语言模型复核候选追溯关系
# =====================================================================

SYSTEM_PROMPT = (
    "你是一位需求工程专家，负责判断两条需求之间是否存在追溯关系。"
    "高级别需求（父需求）粒度较粗，低级别需求（子需求）是对其的细化、拆分、实现或验证。"
    "判定标准：仅当低级别需求明确服务于、细化、实现或验证该高级别需求时，才判定存在追溯关系；"
    "若两者只是提到了相同的变量或话题、并不存在上下游关系，则判定为不存在。"
    "只依据给出的描述内容作答，不要臆测，也不要编造理由。"
)


def chat_completion(system: str, user: str) -> str:
    """调用一次 OpenAI 兼容接口（DeepSeek），返回助手回复文本。"""
    if not LLM_API_KEY:
        raise RuntimeError(
            "未配置 DeepSeek API Key：请在配置区 LLM_API_KEY 中填入，"
            "或设置环境变量 DEEPSEEK_API_KEY 后重试"
        )
    payload = {
        "model": LLM_MODEL,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "temperature": 0,   # 判定类任务取 0，减少随机性
    }
    req = urllib.request.Request(
        LLM_BASE_URL.rstrip("/") + "/chat/completions",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": "Bearer " + LLM_API_KEY,
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=LLM_TIMEOUT) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"大模型接口返回错误（HTTP {e.code}）：{detail}") from e
    except Exception as e:  # URLError / TimeoutError / 网络异常等
        raise RuntimeError(f"调用大模型接口失败：{e}") from e
    try:
        return data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as e:
        raise RuntimeError(f"大模型返回结构异常：{data}") from e


def _norm_bool(value):
    """把模型返回的 traceable 字段归一化为 True / False / None。"""
    if isinstance(value, bool):
        return value
    s = str(value).strip().lower()
    if s in {"true", "yes", "y", "1", "是", "有", "存在"}:
        return True
    if s in {"false", "no", "n", "0", "否", "无", "不存在"}:
        return False
    return None


def parse_review(content: str, n: int):
    """把模型返回的多行 JSON（JSONL）解析为长度 n 的列表，按“no”序号对齐。"""
    result = [None] * n
    for line in content.splitlines():
        line = line.strip()
        if not line.startswith("{"):
            continue
        try:
            obj = json.loads(line)
        except json.JSONDecodeError:
            continue
        if not isinstance(obj, dict):
            continue
        no = obj.get("no")
        if not isinstance(no, int) or not 1 <= no <= n:
            continue
        result[no - 1] = {
            "traceable": _norm_bool(obj.get("traceable")),
            "reason": str(obj.get("reason", "")).strip(),
        }
    return result


def review_batch(items):
    """复核一批候选追溯关系。

    items: list[dict]，每个 dict 含 high_id / low_id / high_desc / low_desc。
    返回与 items 等长（同序）的判定列表，元素形如 {"traceable": bool, "reason": str}。
    单批多次失败时抛 RuntimeError，由调用方兜底标记为“失败”。
    """
    lines = []
    for i, it in enumerate(items, start=1):
        lines.append(
            f'{i}. 高级别需求【{it["high_id"]}】：{it["high_desc"] or "（无描述）"}\n'
            f'   低级别需求【{it["low_id"]}】：{it["low_desc"] or "（无正文）"}'
        )
    user = (
        "以下是若干条“高级别需求—低级别需求”候选追溯关系，请逐条判定。\n"
        "输出要求：\n"
        "1) 严格按顺序为每条候选输出一行 JSON（不要使用代码块标记、不要输出其他任何文字）；\n"
        "2) 每行格式：{\"no\": 序号, \"traceable\": true 或 false, \"reason\": \"不超过50字的理由\"}；\n"
        "3) traceable 为布尔值：true 表示存在追溯关系，false 表示不存在。\n\n"
        "候选列表：\n" + "\n\n".join(lines)
    )
    last_err = None
    for _ in range(MAX_RETRY + 1):
        try:
            content = chat_completion(SYSTEM_PROMPT, user)
        except RuntimeError as e:
            last_err = e
            time.sleep(REVIEW_SLEEP)
            continue
        parsed = parse_review(content, len(items))
        if all(p is not None and p["traceable"] is not None for p in parsed):
            return parsed
        last_err = RuntimeError("大模型返回内容不完整或解析失败")
        time.sleep(REVIEW_SLEEP)
    raise last_err


def resolve_paths():
    """命令行参数（2~4 个）会覆盖配置区；否则使用配置区路径。
    用法：python requirement_trace.py 高级别.xlsx 低级别.xlsx [中间产物.xlsx] [最终结果.xlsx]
    """
    if len(sys.argv) >= 3:
        high, low = sys.argv[1], sys.argv[2]
        interm = sys.argv[3] if len(sys.argv) >= 4 else INTERMEDIATE_FILE
        final = sys.argv[4] if len(sys.argv) >= 5 else FINAL_FILE
        return high, low, interm, final
    return HIGH_LEVEL_FILE, LOW_LEVEL_FILE, INTERMEDIATE_FILE, FINAL_FILE


def main():
    high_file, low_file, interm_file, final_file = resolve_paths()

    for f in (high_file, low_file):
        if not Path(f).exists():
            raise FileNotFoundError(f"找不到文件：{f}")

    # ---------- 读取 ----------
    high_df = normalize_header(pd.read_excel(high_file))
    low_df = normalize_header(pd.read_excel(low_file))

    # 校验必要列是否存在
    for name, df, need in (
        ("高级别", high_df, [COL_ID, COL_IS_REQ, COL_DESC]),
        ("低级别", low_df, [COL_ID, COL_IS_REQ]),
    ):
        missing = [c for c in need if c not in df.columns]
        if missing:
            raise KeyError(
                f"{name}需求表缺少列：{missing}；实际表头为：{list(df.columns)}"
            )

    # 低级别需求的“正文来源”列：优先“标题和正文”，缺失时回退到“描述”
    low_body_col = resolve_low_body_column(low_df)
    if low_body_col is None:
        raise KeyError(
            f"低级别需求表缺少“{COL_BODY}”或“{COL_DESC}”列；"
            f"实际表头为：{list(low_df.columns)}"
        )
    print(f"[读取] 低级别需求正文来源列：{low_body_col}")

    # ---------- 第一步：筛选“是否需求”为 TRUE 的行 ----------
    high_req = high_df[high_df[COL_IS_REQ].map(is_requirement)]
    low_req = low_df[low_df[COL_IS_REQ].map(is_requirement)]
    print(f"[第一步] 筛选完成：高级别需求 {len(high_req)} 条，低级别需求 {len(low_req)} 条")

    # ---------- 第二步：提取变量名建立候选追溯关系，写入“中间产物” ----------
    records = []
    high_with_vars = 0
    for _, h in high_req.iterrows():
        h_id = h[COL_ID]
        vars_ = extract_variables(h[COL_DESC])
        if not vars_:
            continue  # “描述”中无变量，无法建立候选
        high_with_vars += 1
        for _, l in low_req.iterrows():
            if is_traceable(vars_, l[low_body_col]):
                records.append({"高级别需求标识": h_id, "低级别需求标识": l[COL_ID]})

    candidates = pd.DataFrame(records).drop_duplicates()
    candidates.to_excel(interm_file, index=False)
    print(f"[第二步] 含变量的高级别需求 {high_with_vars} 条，候选追溯关系 {len(candidates)} 条，"
          f"中间产物已写入：{interm_file}")

    # ---------- 第三步：大模型复核，写入“追溯结果文件” ----------
    if not ENABLE_LLM_REVIEW:
        print("[第三步] ENABLE_LLM_REVIEW=False，已跳过第二步大模型复核，仅生成中间产物。")
        return

    if len(candidates) == 0:
        pd.DataFrame(
            columns=["高级别需求标识", "低级别需求标识", "LLM判定", "判定理由"]
        ).to_excel(final_file, index=False)
        print(f"[第三步] 无候选追溯关系，最终结果为空文件：{final_file}")
        return

    # 按标识建立“标识 -> 描述/正文”映射，供拼入提示词（复核时的双方描述来源）
    high_text = {str(h[COL_ID]): _text_of_cell(h[COL_DESC]) for _, h in high_req.iterrows()}
    low_text = {str(l[COL_ID]): _text_of_cell(l[low_body_col]) for _, l in low_req.iterrows()}

    cand_records = candidates.to_dict("records")
    total = len(cand_records)
    print(f"[第三步] 开始大模型复核：共 {total} 条候选，每批 {BATCH_SIZE} 条，模型 {LLM_MODEL}…")

    final_records = []
    for start in range(0, total, BATCH_SIZE):
        batch = cand_records[start:start + BATCH_SIZE]
        items = [
            {
                "high_id": r["高级别需求标识"],
                "low_id": r["低级别需求标识"],
                "high_desc": high_text.get(str(r["高级别需求标识"]), ""),
                "low_desc": low_text.get(str(r["低级别需求标识"]), ""),
            }
            for r in batch
        ]
        try:
            verdicts = review_batch(items)
        except RuntimeError as e:
            verdicts = [{"traceable": None, "reason": f"调用失败：{e}"}] * len(items)
        for r, v in zip(batch, verdicts):
            flag = v["traceable"]
            final_records.append(
                {
                    "高级别需求标识": r["高级别需求标识"],
                    "低级别需求标识": r["低级别需求标识"],
                    "LLM判定": "是" if flag is True else ("否" if flag is False else "失败"),
                    "判定理由": v["reason"],
                }
            )
        print(f"[第三步] 已复核 {min(start + BATCH_SIZE, total)}/{total} 条…")
        time.sleep(REVIEW_SLEEP)

    final_df = pd.DataFrame(final_records)
    final_df.to_excel(final_file, index=False)
    yes_cnt = int((final_df["LLM判定"] == "是").sum())
    fail_cnt = int((final_df["LLM判定"] == "失败").sum())
    no_cnt = len(final_df) - yes_cnt - fail_cnt
    print(f"[第三步] 复核完成：判定“是” {yes_cnt} 条 / “否” {no_cnt} 条 / “失败” {fail_cnt} 条，"
          f"已写入：{final_file}")


if __name__ == "__main__":
    main()
