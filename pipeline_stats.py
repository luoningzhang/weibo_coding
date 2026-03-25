"""
各电影数据管道各环节统计

用法:
    python pipeline_stats.py                     # 扫描 data/tmp/
    python pipeline_stats.py --dir data/tmp      # 指定根目录
    python pipeline_stats.py --excel pipeline_stats.xlsx  # 导出 Excel
"""

import argparse
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

# 按顺序定义各阶段：(阶段名, 文件后缀)
STAGES = [
    ("①原始",               "_电影.json"),
    ("②预清洗",              "_电影_cleaned.json"),
    ("③去认证用户",           "_电影_cleaned_no_verified.json"),
    ("④去V类",               "_电影_cleaned_no_verified_filtered.json"),
    ("⑤去S类",               "_电影_classified_F.json"),
    ("⑥编码完成",             "_电影_coded.json"),
]


def load_posts(path: Path) -> list[dict]:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, list):
        return data
    for key in ("data", "posts", "items"):
        if key in data and isinstance(data[key], list):
            return data[key]
    return []


def get_user_id(post: dict):
    u = post.get("user")
    if isinstance(u, dict):
        return u.get("id") or u.get("user_id")
    return u or post.get("user_id")


def stage_stats(posts: list[dict]) -> tuple[int, int]:
    """返回 (条数, 不同用户数)"""
    count = len(posts)
    users = len({get_user_id(p) for p in posts if get_user_id(p) is not None})
    return count, users


def find_movies(root: Path) -> list[str]:
    """扫描子目录，返回电影名列表"""
    movies = sorted(
        d.name.replace("_电影", "")
        for d in root.iterdir()
        if d.is_dir() and not d.name.startswith(".")
    )
    return movies


def collect(root: Path, movie: str) -> list[tuple[str, int | None, int | None]]:
    """返回 [(阶段名, 条数, 用户数), ...]，文件缺失时为 None"""
    folder = root / f"{movie}_电影"
    if not folder.exists():
        folder = root / movie  # fallback：直接以电影名命名的目录
    rows = []
    for stage_name, suffix in STAGES:
        path = folder / f"{movie}{suffix}"
        if path.exists():
            posts = load_posts(path)
            cnt, users = stage_stats(posts)
            rows.append((stage_name, cnt, users))
        else:
            rows.append((stage_name, None, None))
    return rows


def print_report(root: Path, movies: list[str]):
    sep = "=" * 72
    hdr = f"  {'阶段':<12} {'条数':>8} {'不同用户':>10} {'条数保留%':>10} {'用户保留%':>10}"
    for movie in movies:
        rows = collect(root, movie)
        print(f"\n{sep}")
        print(f"  电影：{movie}")
        print(sep)
        print(hdr)
        print(f"  {'-'*68}")
        base_cnt = base_usr = None
        for stage, cnt, users in rows:
            if cnt is None:
                print(f"  {stage:<12} {'—':>8} {'—':>10} {'—':>10} {'—':>10}")
                continue
            if base_cnt is None:
                base_cnt, base_usr = cnt, users
            pct_cnt = cnt  / base_cnt * 100 if base_cnt else 0
            pct_usr = users / base_usr * 100 if base_usr else 0
            print(f"  {stage:<12} {cnt:>8,} {users:>10,} {pct_cnt:>9.1f}% {pct_usr:>9.1f}%")
    print(f"\n{sep}\n")


def export_excel(root: Path, movies: list[str], out_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "管道统计"

    # 表头
    header = ["电影", "阶段", "条数", "不同用户", "条数保留%", "用户保留%"]
    ws.append(header)
    hdr_fill = PatternFill("solid", fgColor="4472C4")
    hdr_font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center")

    fill_alt = PatternFill("solid", fgColor="EBF1FA")
    for i, movie in enumerate(movies):
        rows = collect(root, movie)
        base_cnt = base_usr = None
        for stage, cnt, users in rows:
            if cnt is None:
                ws.append([movie, stage, "—", "—", "—", "—"])
            else:
                if base_cnt is None:
                    base_cnt, base_usr = cnt, users
                pct_cnt = round(cnt   / base_cnt * 100, 1) if base_cnt else 0
                pct_usr = round(users / base_usr * 100, 1) if base_usr else 0
                ws.append([movie, stage, cnt, users, pct_cnt, pct_usr])
            if i % 2 == 1:
                for cell in ws[ws.max_row]:
                    cell.fill = fill_alt
        # 空行分隔
        ws.append([""] * 6)

    # 自动列宽
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 30)

    wb.save(out_path)
    print(f"✅ Excel 已导出 → {out_path}")


def main():
    parser = argparse.ArgumentParser(description="数据管道各环节条数与用户数统计")
    parser.add_argument("--dir",   default="data/tmp", help="数据根目录")
    parser.add_argument("--excel", default="",         help="导出 Excel 路径")
    args = parser.parse_args()

    root = Path(args.dir)
    if not root.exists():
        print(f"❌ 目录不存在：{root}")
        return

    movies = find_movies(root)
    if not movies:
        print(f"❌ {root} 下未找到任何子目录")
        return

    print(f"找到 {len(movies)} 部电影：{', '.join(movies)}")
    print_report(root, movies)

    if args.excel:
        export_excel(root, movies, args.excel)


if __name__ == "__main__":
    main()
