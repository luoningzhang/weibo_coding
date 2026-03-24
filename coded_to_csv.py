"""
把 *_coded.json 转成 CSV

用法:
    python coded_to_csv.py data/tmp/长安三万里_电影/长安三万里_电影_coded.json
    python coded_to_csv.py data/tmp/长安三万里_电影/长安三万里_电影_coded.json --output 长安.csv
"""

import argparse
import csv
import json
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description="coded.json → CSV")
    parser.add_argument("input", help="*_coded.json 路径")
    parser.add_argument("--output", default="", help="输出 CSV 路径（默认同名 .csv）")
    args = parser.parse_args()

    input_path  = Path(args.input)
    output_path = Path(args.output) if args.output else input_path.with_suffix(".csv")

    with open(input_path, encoding="utf-8") as f:
        posts = json.load(f)
    if isinstance(posts, dict):
        posts = list(posts.values())

    if not posts:
        print("文件为空")
        return

    # 展开 user 字段，codes/code_names 转字符串
    def flatten(p: dict) -> dict:
        row = {}
        for k, v in p.items():
            if k == "user" and isinstance(v, dict):
                for uk, uv in v.items():
                    row[f"user_{uk}"] = uv
            elif k == "codes":
                row["codes"] = ", ".join(str(c) for c in v) if v else ""
            elif k == "code_names":
                row["code_names"] = ", ".join(str(c) for c in v) if v else ""
            else:
                row[k] = v
        return row

    rows = [flatten(p) for p in posts]

    # 保证列顺序稳定
    priority = ["id", "movie", "text", "codes", "code_names",
                "user_screen_name", "user_id", "user_verified",
                "created_at", "reposts_count", "comments_count",
                "attitudes_count", "engagement_total", "predicted_label"]
    all_keys = list(dict.fromkeys(
        [k for k in priority if any(k in r for r in rows)] +
        [k for r in rows for k in r if k not in priority]
    ))

    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=all_keys, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)

    print(f"✅ 已导出 {len(rows)} 条 → {output_path}")


if __name__ == "__main__":
    main()
