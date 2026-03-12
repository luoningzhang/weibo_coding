"""
批量编码微博（云雾 API / claude-sonnet-4-6）

用法:
    python batch_code.py \
        --input  data/no_verified_classified_with_movie.json \
        --output data/coded_output.xlsx \
        --key    YOUR_YUNWU_API_KEY

可选:
    --codebook  data/codebook.xlsx          # 默认
    --batch     20                          # 每次请求的条数，默认 20
    --workers   5                           # 并发数，默认 5
    --checkpoint data/checkpoint.json       # 断点续跑文件，默认同目录

节省 token 的设计：
  1. System prompt 只包含编码手册精简版（编号+名称+识别关键词）
  2. 每条微博只传 user/movie/text 三个字段
  3. 每批 20 条一起发，返回纯 JSON 数字数组，无多余文字
"""

import argparse
import json
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openai import OpenAI

# ── 常量 ──────────────────────────────────────────────────────────────────────
YUNWU_BASE_URL = "https://yunwu.ai/v1"
MODEL = "claude-sonnet-4-6"
MAX_RETRIES = 3


# ── 读取 codebook ──────────────────────────────────────────────────────────────
def load_codebook(path: str) -> list[dict]:
    wb = load_workbook(path)
    ws = wb.active
    codes = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        num, name, definition, criteria, note = (list(row) + [None] * 5)[:5]
        if num is None:
            continue
        codes.append(
            {
                "num": int(num),
                "name": str(name).strip(),
                "criteria": str(criteria).strip() if criteria else "",
                "note": str(note).strip() if note else "",
            }
        )
    return codes


def build_system_prompt(codes: list[dict]) -> str:
    """将编码手册压缩为最精简的 system prompt，节省 token。"""
    lines = [
        "你是一位内容分析员，负责对微博文本按行为类别编码。",
        "每条微博可对应 0 到多个编码。只返回 JSON，不要任何解释。",
        "",
        "## 编码列表（编号 名称 | 识别关键词/特征）",
    ]
    for c in codes:
        # 只保留识别标准中的关键词部分（关键词:之后的内容），进一步压缩
        criteria_short = c["criteria"]
        if "关键词：" in criteria_short:
            kw_part = criteria_short.split("关键词：")[1].split("；")[0]
            criteria_short = kw_part
        elif "行为特征：" in criteria_short:
            kw_part = criteria_short.split("行为特征：")[1].split("；")[0]
            criteria_short = kw_part
        note_part = f"（注：{c['note'][:40]}）" if c["note"] else ""
        lines.append(f"{c['num']} {c['name']} | {criteria_short}{note_part}")

    lines += [
        "",
        "## 输出格式",
        '输入是一个 JSON 数组，每项含 id/user/movie/text 字段。',
        '输出必须是一个 JSON 数组，顺序与输入一一对应，',
        '每项是该微博适用的编码编号数组（没有适用编码则为空数组[]）。',
        "例：输入3条 → 输出 [[5,12],[22],[]]",
    ]
    return "\n".join(lines)


# ── 读取微博数据 ───────────────────────────────────────────────────────────────
def load_posts(path: str) -> list[dict]:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, list):
        return data
    for key in ("data", "posts", "statuses", "items"):
        if key in data and isinstance(data[key], list):
            return data[key]
    raise ValueError(f"无法识别 JSON 结构: {type(data)}")


# ── 调用 API ──────────────────────────────────────────────────────────────────
def call_api(
    client: OpenAI,
    system_prompt: str,
    batch: list[dict],
    batch_index: int,
) -> list[list[int]]:
    """发送一批微博，返回每条的编码列表。失败时重试。"""
    # 只传三个字段：id（批内序号）、user、movie、text
    payload = [
        {
            "id": i + 1,
            "user": p.get("user", {}).get("screen_name", "") if isinstance(p.get("user"), dict) else str(p.get("user", "")),
            "movie": p.get("movie", ""),
            "text": (p.get("text") or p.get("content") or "")[:500],  # 截断超长正文
        }
        for i, p in enumerate(batch)
    ]
    user_msg = json.dumps(payload, ensure_ascii=False)

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_msg},
                ],
                temperature=0,
                max_tokens=512,
            )
            raw = resp.choices[0].message.content.strip()
            # 提取 JSON 数组（防止模型前后有多余文字）
            start, end = raw.find("["), raw.rfind("]")
            if start == -1 or end == -1:
                raise ValueError(f"响应不含 JSON 数组: {raw[:200]}")
            result = json.loads(raw[start : end + 1])
            if len(result) != len(batch):
                raise ValueError(
                    f"返回条数 {len(result)} ≠ 批次条数 {len(batch)}"
                )
            return result
        except Exception as e:
            wait = 2 ** attempt
            print(f"  [批次{batch_index}] 第{attempt}次失败: {e}，{wait}s 后重试…", flush=True)
            if attempt == MAX_RETRIES:
                print(f"  [批次{batch_index}] 放弃，填充空编码", flush=True)
                return [[] for _ in batch]
            time.sleep(wait)


# ── 导出 Excel ────────────────────────────────────────────────────────────────
def export_excel(posts: list[dict], code_results: dict[str, list[int]], codes: list[dict], path: str):
    code_name_map = {c["num"]: c["name"] for c in codes}

    wb = Workbook()
    ws = wb.active
    ws.title = "编码结果"

    headers = ["微博ID", "用户名", "电影", "正文", "适用编码（编号）", "适用编码（名称）"]
    hfill = PatternFill("solid", fgColor="4F81BD")
    hfont = Font(bold=True, color="FFFFFF")
    for col, label in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=label)
        c.fill = hfill
        c.font = hfont
        c.alignment = Alignment(horizontal="center")

    for row_i, post in enumerate(posts, 2):
        pid = str(post.get("id") or post.get("mid") or post.get("idstr", ""))
        user = post.get("user") or {}
        user_name = user.get("screen_name", "") if isinstance(user, dict) else str(user)
        movie = post.get("movie", "")
        text = post.get("text") or post.get("content") or ""

        nums = code_results.get(pid, [])
        nums_str = ", ".join(str(n) for n in nums)
        names_str = ", ".join(f"{n} {code_name_map.get(n, '')}" for n in nums)

        row_vals = [pid, user_name, movie, text, nums_str, names_str]
        for col, val in enumerate(row_vals, 1):
            cell = ws.cell(row=row_i, column=col, value=val)
            if col == 1:  # ID 列保文本
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="left")

    col_widths = [22, 16, 20, 80, 20, 60]
    for col, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"

    wb.save(path)


# ── 主流程 ────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="批量编码微博")
    parser.add_argument("--input", default="data/no_verified_classified_with_movie.json")
    parser.add_argument("--output", default="data/coded_output.xlsx")
    parser.add_argument("--codebook", default="data/codebook.xlsx")
    parser.add_argument("--key", required=True, help="云雾 API Key")
    parser.add_argument("--batch", type=int, default=20, help="每批条数")
    parser.add_argument("--workers", type=int, default=5, help="并发线程数")
    parser.add_argument("--checkpoint", default=None, help="断点文件路径")
    args = parser.parse_args()

    checkpoint_path = args.checkpoint or (Path(args.output).stem + "_checkpoint.json")

    # 初始化客户端
    client = OpenAI(api_key=args.key, base_url=YUNWU_BASE_URL)

    # 加载数据
    print("读取 codebook…")
    codes = load_codebook(args.codebook)
    system_prompt = build_system_prompt(codes)
    print(f"System prompt 长度：{len(system_prompt)} 字符")

    print("读取微博数据…")
    posts = load_posts(args.input)
    print(f"共 {len(posts)} 条")

    # 读取断点
    if os.path.exists(checkpoint_path):
        with open(checkpoint_path, encoding="utf-8") as f:
            code_results: dict[str, list[int]] = json.load(f)
        print(f"断点续跑：已完成 {len(code_results)} 条")
    else:
        code_results = {}

    # 过滤已完成
    todo_posts = [
        p for p in posts
        if str(p.get("id") or p.get("mid") or p.get("idstr", "")) not in code_results
    ]
    print(f"待处理：{len(todo_posts)} 条，批大小={args.batch}，并发={args.workers}")

    # 分批
    batches = [
        todo_posts[i : i + args.batch]
        for i in range(0, len(todo_posts), args.batch)
    ]

    completed = 0
    total = len(todo_posts)

    def process_batch(idx_batch):
        idx, batch = idx_batch
        result = call_api(client, system_prompt, batch, idx)
        return batch, result

    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {executor.submit(process_batch, (i, b)): i for i, b in enumerate(batches)}
        for future in as_completed(futures):
            batch, result = future.result()
            for post, nums in zip(batch, result):
                pid = str(post.get("id") or post.get("mid") or post.get("idstr", ""))
                code_results[pid] = [int(n) for n in nums if isinstance(n, (int, float))]
            completed += len(batch)

            # 保存断点
            with open(checkpoint_path, "w", encoding="utf-8") as f:
                json.dump(code_results, f, ensure_ascii=False)

            pct = completed / total * 100
            print(f"进度：{completed}/{total} ({pct:.1f}%)", end="\r", flush=True)

    print(f"\n编码完成，共 {len(code_results)} 条")

    # 输出 Excel
    print(f"导出 Excel → {args.output}")
    export_excel(posts, code_results, codes, args.output)
    print("完成！")


if __name__ == "__main__":
    main()
