"""
批量编码微博（云雾 API / claude-sonnet-4-6）

用法:
    python batch_code.py

可选参数:
    --input     data/no_verified_classified_with_movie.json
    --output    data/coded_output.json
    --codebook  data/codebook.xlsx
    --batch     20      每次请求的条数
    --workers   5       并发数

API Key 放在 config.json 的 api_key 字段中。
"""

import argparse
import json
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from openpyxl import load_workbook
from openai import OpenAI

YUNWU_BASE_URL = "https://yunwu.ai/v1"
MODEL = "claude-sonnet-4-6"
MAX_RETRIES = 3
CONFIG_PATH = Path(__file__).parent / "config.json"


def load_config() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def load_codebook(path: str) -> list[dict]:
    wb = load_workbook(path)
    ws = wb.active
    codes = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        num, name, definition, criteria, note = (list(row) + [None] * 5)[:5]
        if num is None:
            continue
        codes.append({
            "num": int(num),
            "name": str(name).strip(),
            "criteria": str(criteria).strip() if criteria else "",
            "note": str(note).strip() if note else "",
        })
    return codes


def build_system_prompt(codes: list[dict]) -> str:
    lines = [
        "你是一位内容分析员，负责对微博文本按行为类别编码。",
        "每条微博可对应 0 到多个编码。只返回 JSON，不要任何解释。",
        "",
        "## 编码列表（编号 名称 | 识别关键词/特征）",
    ]
    for c in codes:
        criteria_short = c["criteria"]
        if "关键词：" in criteria_short:
            criteria_short = criteria_short.split("关键词：")[1].split("；")[0]
        elif "行为特征：" in criteria_short:
            criteria_short = criteria_short.split("行为特征：")[1].split("；")[0]
        note_part = f"（注：{c['note'][:40]}）" if c["note"] else ""
        lines.append(f"{c['num']} {c['name']} | {criteria_short}{note_part}")
    lines += [
        "",
        "## 输出格式",
        "输入是一个 JSON 数组，每项含 id/user/movie/text 字段。",
        "输出必须是一个 JSON 数组，顺序与输入一一对应，",
        "每项是该微博适用的编码编号数组（没有适用编码则为空数组[]）。",
        "例：输入3条 → 输出 [[5,12],[22],[]]",
    ]
    return "\n".join(lines)


def load_posts(path: str) -> list[dict]:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, list):
        return data
    for key in ("data", "posts", "statuses", "items"):
        if key in data and isinstance(data[key], list):
            return data[key]
    raise ValueError(f"无法识别 JSON 结构: {type(data)}")


def post_id(post: dict) -> str:
    return str(post.get("id") or post.get("mid") or post.get("idstr", ""))


def call_api(client: OpenAI, system_prompt: str, batch: list[dict], batch_index: int) -> list[list[int]]:
    payload = [
        {
            "id": i + 1,
            "user": p.get("user", {}).get("screen_name", "") if isinstance(p.get("user"), dict) else str(p.get("user", "")),
            "movie": p.get("movie", ""),
            "text": (p.get("text") or p.get("content") or "")[:500],
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
            start, end = raw.find("["), raw.rfind("]")
            if start == -1 or end == -1:
                raise ValueError(f"响应不含 JSON 数组: {raw[:200]}")
            result = json.loads(raw[start:end + 1])
            if len(result) != len(batch):
                raise ValueError(f"返回条数 {len(result)} ≠ 批次条数 {len(batch)}")
            return result
        except Exception as e:
            wait = 2 ** attempt
            print(f"  [批次{batch_index}] 第{attempt}次失败: {e}，{wait}s 后重试…", flush=True)
            if attempt == MAX_RETRIES:
                print(f"  [批次{batch_index}] 放弃，填充空编码", flush=True)
                return [[] for _ in batch]
            time.sleep(wait)


def main():
    parser = argparse.ArgumentParser(description="批量编码微博")
    parser.add_argument("--input",    default="data/no_verified_classified_with_movie.json")
    parser.add_argument("--output",   default="data/coded_output.json")
    parser.add_argument("--codebook", default="data/codebook.xlsx")
    parser.add_argument("--batch",    type=int, default=20)
    parser.add_argument("--workers",  type=int, default=5)
    args = parser.parse_args()

    checkpoint_path = Path(args.output).with_suffix(".checkpoint.json")

    config = load_config()
    client = OpenAI(api_key=config["api_key"], base_url=YUNWU_BASE_URL)

    print("读取 codebook…")
    codes = load_codebook(args.codebook)
    system_prompt = build_system_prompt(codes)
    code_name_map = {c["num"]: c["name"] for c in codes}

    print("读取微博数据…")
    posts = load_posts(args.input)
    print(f"共 {len(posts)} 条")

    # 断点续跑
    if checkpoint_path.exists():
        with open(checkpoint_path, encoding="utf-8") as f:
            code_results: dict[str, list[int]] = json.load(f)
        print(f"断点续跑：已完成 {len(code_results)} 条")
    else:
        code_results = {}

    todo_posts = [p for p in posts if post_id(p) not in code_results]
    print(f"待处理：{len(todo_posts)} 条，批大小={args.batch}，并发={args.workers}")

    if todo_posts:
        batches = [todo_posts[i:i + args.batch] for i in range(0, len(todo_posts), args.batch)]
        completed = 0
        total = len(todo_posts)

        def process_batch(idx_batch):
            idx, batch = idx_batch
            return batch, call_api(client, system_prompt, batch, idx)

        with ThreadPoolExecutor(max_workers=args.workers) as executor:
            futures = {executor.submit(process_batch, (i, b)): i for i, b in enumerate(batches)}
            for future in as_completed(futures):
                batch, result = future.result()
                for p, nums in zip(batch, result):
                    code_results[post_id(p)] = [int(n) for n in nums if isinstance(n, (int, float))]
                completed += len(batch)
                with open(checkpoint_path, "w", encoding="utf-8") as f:
                    json.dump(code_results, f, ensure_ascii=False)
                print(f"进度：{completed}/{total} ({completed/total*100:.1f}%)", end="\r", flush=True)
        print(f"\n编码完成，共 {len(code_results)} 条")

    # 输出 JSON：在原始数据上附加 codes 字段
    output = []
    for p in posts:
        pid = post_id(p)
        nums = code_results.get(pid, [])
        output.append({
            **p,
            "codes": nums,
            "code_names": [f"{n} {code_name_map.get(n, '')}" for n in nums],
        })

    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"已写出 → {args.output}")


if __name__ == "__main__":
    main()
