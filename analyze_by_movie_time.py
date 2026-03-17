"""
按电影、上映阶段分析编码分布与互动量。

上映日期（院线首映）：
  流浪地球2   2023-01-22
  消失的她    2023-06-22
  封神第一部  2023-07-20
  长安三万里  2023-07-08
  孤注一掷    2023-08-08

时间分段：
  上映前   created_at < 首映日
  上映期   首映日 ≤ created_at ≤ 首映日 + 30天
  长尾期   created_at > 首映日 + 30天

用法：
    python analyze_by_movie_time.py
    python analyze_by_movie_time.py --input data/coded_output.json --codebook data/codebook.xlsx
"""

import argparse
import json
from collections import defaultdict
from datetime import datetime, timedelta

from openpyxl import load_workbook

RELEASE_DATES = {
    "流浪地球2_电影": datetime(2023, 1, 22),
    "消失的她_电影":  datetime(2023, 6, 22),
    "封神_电影":      datetime(2023, 7, 20),
    "长安三万里_电影": datetime(2023, 7, 8),
    "孤注一掷_电影":  datetime(2023, 8, 8),
}
HOT_DAYS = 30  # 上映期窗口


def load_code_names(codebook_path: str) -> dict[int, str]:
    wb = load_workbook(codebook_path)
    ws = wb.active
    names = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        num, name = row[0], row[1]
        if num is not None:
            names[int(num)] = str(name).strip()
    return names


def parse_dt(s: str):
    try:
        return datetime.strptime(s, "%a %b %d %H:%M:%S +0800 %Y")
    except Exception:
        return None


def get_phase(dt: datetime, release: datetime) -> str:
    if dt < release:
        return "上映前"
    elif dt <= release + timedelta(days=HOT_DAYS):
        return "上映期"
    else:
        return "长尾期"


def engagement(post: dict) -> int:
    return (int(post.get("reposts_count") or 0)
            + int(post.get("comments_count") or 0)
            + int(post.get("attitudes_count") or 0))


def print_table(title: str, rows: list, col_header: str):
    """rows: [(label, posts, eng, code_dist_dict)]"""
    print(f"\n{'='*70}")
    print(f"  {title}")
    print(f"{'='*70}")
    print(f"  {'类别':<18} {'帖数':>7} {'互动总量':>12}  编码 Top5（频次）")
    print(f"  {'─'*65}")
    for label, posts, eng, code_dist in rows:
        top5 = sorted(code_dist.items(), key=lambda x: -x[1])[:5]
        top5_str = "  ".join(f"{c}×{n}" for c, n in top5)
        print(f"  {label:<18} {posts:>7,} {eng:>12,}  {top5_str}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input",    default="data/coded_output.json")
    parser.add_argument("--codebook", default="data/codebook.xlsx")
    args = parser.parse_args()

    with open(args.input, encoding="utf-8") as f:
        data = json.load(f)

    code_names = load_code_names(args.codebook)

    # ── 1. 五部电影汇总 ───────────────────────────────────────────
    movie_posts   = defaultdict(int)
    movie_eng     = defaultdict(int)
    movie_codes   = defaultdict(lambda: defaultdict(int))

    # ── 2. 电影 × 阶段 ────────────────────────────────────────────
    phase_posts   = defaultdict(lambda: defaultdict(int))     # [movie][phase]
    phase_eng     = defaultdict(lambda: defaultdict(int))
    phase_codes   = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))

    # ── 3. 全库各编码 × 阶段 ─────────────────────────────────────
    code_phase_posts = defaultdict(lambda: defaultdict(int))  # [code][phase]
    code_phase_eng   = defaultdict(lambda: defaultdict(int))

    PHASES = ["上映前", "上映期", "长尾期"]

    for post in data:
        movie   = post.get("movie", "")
        dt      = parse_dt(post.get("created_at", ""))
        eng     = engagement(post)
        codes   = [int(c) for c in post.get("codes", [])]
        release = RELEASE_DATES.get(movie)

        # 电影维度
        movie_posts[movie] += 1
        movie_eng[movie]   += eng
        for c in codes:
            movie_codes[movie][c] += 1

        # 阶段维度
        if dt and release:
            phase = get_phase(dt, release)
            phase_posts[movie][phase] += 1
            phase_eng[movie][phase]   += eng
            for c in codes:
                phase_codes[movie][phase][c] += 1
                code_phase_posts[c][phase]   += 1
                code_phase_eng[c][phase]     += eng

    # ══════════════════════════════════════════════════════════════
    #  输出 1：五部电影汇总
    # ══════════════════════════════════════════════════════════════
    print("\n" + "="*70)
    print("  【一】五部电影总览")
    print("="*70)
    print(f"  {'电影':<18} {'帖数':>7} {'互动总量':>12}  编码 Top5（频次）")
    print(f"  {'─'*65}")
    for movie in sorted(RELEASE_DATES):
        n   = movie_posts[movie]
        eng = movie_eng[movie]
        top5 = sorted(movie_codes[movie].items(), key=lambda x: -x[1])[:5]
        top5_str = "  ".join(f"{code_names.get(c,c)}×{cnt}" for c, cnt in top5)
        print(f"  {movie:<18} {n:>7,} {eng:>12,}  {top5_str}")

    # ══════════════════════════════════════════════════════════════
    #  输出 2：每部电影的三阶段分布
    # ══════════════════════════════════════════════════════════════
    print("\n\n" + "="*70)
    print("  【二】每部电影 × 上映阶段（帖数 / 互动 / 编码 Top3）")
    for movie in sorted(RELEASE_DATES):
        release = RELEASE_DATES[movie]
        print(f"\n  ── {movie}  首映：{release.strftime('%Y-%m-%d')} ──")
        print(f"  {'阶段':<10} {'帖数':>7} {'互动总量':>12}  编码 Top3（频次）")
        print(f"  {'─'*55}")
        for phase in PHASES:
            n   = phase_posts[movie][phase]
            eng = phase_eng[movie][phase]
            top3 = sorted(phase_codes[movie][phase].items(), key=lambda x: -x[1])[:3]
            top3_str = "  ".join(f"{code_names.get(c,c)}×{cnt}" for c, cnt in top3)
            print(f"  {phase:<10} {n:>7,} {eng:>12,}  {top3_str}")

    # ══════════════════════════════════════════════════════════════
    #  输出 3：各编码在三阶段的分布（帖数）
    # ══════════════════════════════════════════════════════════════
    print("\n\n" + "="*70)
    print("  【三】各编码 × 上映阶段分布（帖数 / 互动）")
    print("="*70)
    print(f"  {'编码':<5} {'名称':<18} " +
          "".join(f"  {p}帖数  {p}互动" for p in PHASES))
    print(f"  {'─'*85}")
    for code in sorted(code_names):
        name = code_names[code]
        row = f"  {code:<5} {name:<18}"
        for phase in PHASES:
            n   = code_phase_posts[code][phase]
            eng = code_phase_eng[code][phase]
            row += f"  {n:>6,}  {eng:>10,}"
        print(row)


if __name__ == "__main__":
    main()
