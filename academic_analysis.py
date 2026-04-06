"""
学术分析脚本：六大行为类别 × 四阶段参与结构

输出：
  Table 1 — 六大类行为分布总表（电影内部占比 + 全样本占比）
  Table 2 — 行为频次与共鸣效率对比（帖数占比 / 互动占比 / 均互动 / 效率比）
  Figure 1 — 四阶段共鸣结构变化折线图（PNG）
  Table 3 — SIPs行为指纹对比（含code17 vs 全样本）
  Table 4 — SIPs跨电影分布

用法：
    python academic_analysis.py
    python academic_analysis.py --dir data/tmp --out academic_results.xlsx --fig figure1.png
"""

import argparse
import json
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

# ── 六大类别定义（二阶编码）────────────────────────────────────────
CATEGORIES = [
    ("内容生产劳动",   "Content Production Labor",        {5, 6, 7, 8}),
    ("扩散与动员劳动", "Diffusion & Mobilization Labor",  {9, 11, 12, 13, 28}),
    ("数据与市场劳动", "Data & Market Labor",              {2, 3, 4, 26}),
    ("策略性介入",     "Strategic Intervention",          {21, 22, 23, 24, 25, 27}),
    ("身份认同表达",   "Identity & Affiliation Expr.",    {1, 10, 17, 18, 19}),
    ("话语防御",       "Discourse Defense",               {14, 15, 16, 20}),
]
CAT_NAMES_ZH = [c[0] for c in CATEGORIES]
CAT_CODES    = [c[2] for c in CATEGORIES]
N_CATS       = len(CATEGORIES)

# SIPs 核心编码
SIPS_CODE = 17
# Table 3 额外高亮的一阶编码
HIGHLIGHT_CODES = [21, 22, 2, 3]

# ── 上映日期 ──────────────────────────────────────────────────────
RELEASE_DATES_RAW = {
    "流浪地球2": datetime(2023, 1, 22),
    "消失的她":  datetime(2023, 6, 22),
    "封神":      datetime(2023, 7, 20),
    "长安三万里": datetime(2023, 7, 8),
    "孤注一掷":  datetime(2023, 8, 8),
}
# 同时支持 "XX_电影" 格式
RELEASE_DATES = {**RELEASE_DATES_RAW,
                 **{k + "_电影": v for k, v in RELEASE_DATES_RAW.items()}}

MOVIE_ORDER = list(RELEASE_DATES_RAW.keys())  # 展示顺序（按上映时间）
MOVIE_ORDER_SORTED = sorted(MOVIE_ORDER, key=lambda m: RELEASE_DATES_RAW[m])

# ── 四阶段定义（相对首映日） ──────────────────────────────────────
PHASES = ["上映前期", "上映初期", "上映中期", "长尾期"]
PHASE_RANGES = [(-30, -1), (1, 14), (15, 30), (31, 120)]   # [days from release]

FILTER_PATTERN = "今年我最爱的#微博年度电影#是"


# ─────────────────────────────────────────────────────────────────
#  工具函数
# ─────────────────────────────────────────────────────────────────

def parse_dt(s: str):
    if not s:
        return None
    try:
        return datetime.strptime(s, "%a %b %d %H:%M:%S +0800 %Y")
    except Exception:
        pass
    try:
        return datetime.fromisoformat(s[:19])
    except Exception:
        return None


def get_phase(dt: datetime, release: datetime):
    delta = (dt - release).days + (1 if (dt - release).seconds > 0 and (dt - release).days < 0 else 0)
    # more precise: use total days difference
    diff = (dt.date() - release.date()).days
    for i, (lo, hi) in enumerate(PHASE_RANGES):
        if lo <= diff <= hi:
            return i
    return None  # outside tracked window


def post_engagement(p: dict) -> int:
    return (int(p.get("reposts_count") or 0)
            + int(p.get("comments_count") or 0)
            + int(p.get("attitudes_count") or 0))


def post_cat_flags(codes: list[int]) -> list[bool]:
    code_set = set(codes)
    return [bool(code_set & cat_codes) for cat_codes in CAT_CODES]


def normalize_movie(raw: str) -> str:
    """Return canonical movie name (without _电影 suffix)."""
    return raw.removesuffix("_电影").strip()


def load_posts(root: Path) -> list[dict]:
    posts = []
    for path in sorted(root.rglob("*_coded.json")):
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            chunk = data
        else:
            chunk = data.get("data", data.get("posts", data.get("items", [])))
        # inject movie name from filename if post has none
        movie_from_file = path.stem.removesuffix("_coded").removesuffix("_电影")
        for p in chunk:
            if not p.get("movie"):
                p["movie"] = movie_from_file
            posts.append(p)
    return posts


# ─────────────────────────────────────────────────────────────────
#  主计算
# ─────────────────────────────────────────────────────────────────

def build_stats(posts: list[dict]):
    """
    Returns a nested stats dict used by all tables / figures.
    """
    # per-movie totals
    movie_posts  = defaultdict(int)           # movie -> total posts
    movie_eng    = defaultdict(int)           # movie -> total engagement

    # category accumulators
    # cat_posts[cat]        -> total posts belonging to this category (global)
    # cat_eng[cat]          -> total engagement of posts in this category (global)
    cat_posts   = [0] * N_CATS
    cat_eng     = [0] * N_CATS

    # per-movie category posts
    # mv_cat_posts[movie][cat] -> posts
    mv_cat_posts = defaultdict(lambda: [0] * N_CATS)

    # four-phase × category
    # phase_cat_posts[phase][cat] -> posts count
    # phase_cat_eng[phase][cat]   -> engagement sum
    phase_cat_posts = [[0] * N_CATS for _ in range(4)]
    phase_cat_eng   = [[0] * N_CATS for _ in range(4)]

    # SIPs accumulators
    sips_posts   = 0
    sips_cat_posts = [0] * N_CATS          # how many SIPs posts touch each cat
    sips_code_posts = defaultdict(int)     # raw code co-occurrence in SIPs
    all_code_posts  = defaultdict(int)     # raw code occurrence in full sample

    # SIPs per movie
    sips_by_movie = defaultdict(int)

    total_posts = 0
    total_eng   = 0

    for p in posts:
        text = p.get("text") or ""
        if FILTER_PATTERN in text:
            continue

        movie_raw = p.get("movie", "")
        movie     = normalize_movie(movie_raw)
        codes     = [int(c) for c in p.get("codes") or []]
        eng       = post_engagement(p)
        flags     = post_cat_flags(codes)

        release = RELEASE_DATES.get(movie) or RELEASE_DATES.get(movie + "_電影")
        dt      = parse_dt(p.get("created_at", ""))
        phase_i = None
        if dt and release:
            phase_i = get_phase(dt, release)

        total_posts += 1
        total_eng   += eng
        movie_posts[movie] += 1
        movie_eng[movie]   += eng

        for i, flag in enumerate(flags):
            if flag:
                cat_posts[i]          += 1
                cat_eng[i]            += eng
                mv_cat_posts[movie][i] += 1
                if phase_i is not None:
                    phase_cat_posts[phase_i][i] += 1
                    phase_cat_eng[phase_i][i]   += eng

        for c in codes:
            all_code_posts[c] += 1

        is_sip = SIPS_CODE in codes
        if is_sip:
            sips_posts += 1
            sips_by_movie[movie] += 1
            for i, flag in enumerate(flags):
                if flag:
                    sips_cat_posts[i] += 1
            for c in codes:
                sips_code_posts[c] += 1

    return {
        "total_posts":       total_posts,
        "total_eng":         total_eng,
        "movie_posts":       dict(movie_posts),
        "movie_eng":         dict(movie_eng),
        "cat_posts":         cat_posts,
        "cat_eng":           cat_eng,
        "mv_cat_posts":      {k: list(v) for k, v in mv_cat_posts.items()},
        "phase_cat_posts":   phase_cat_posts,
        "phase_cat_eng":     phase_cat_eng,
        "sips_posts":        sips_posts,
        "sips_cat_posts":    sips_cat_posts,
        "sips_code_posts":   dict(sips_code_posts),
        "all_code_posts":    dict(all_code_posts),
        "sips_by_movie":     dict(sips_by_movie),
    }


# ─────────────────────────────────────────────────────────────────
#  Excel 输出
# ─────────────────────────────────────────────────────────────────

def pct(num, denom):
    return round(num / denom * 100, 2) if denom else 0.0


def export_excel(stats: dict, out_path: str):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()

    HDR_FILL  = PatternFill("solid", fgColor="2F5597")
    HDR_FONT  = Font(bold=True, color="FFFFFF", size=10)
    SUM_FILL  = PatternFill("solid", fgColor="D6E4F0")
    SUM_FONT  = Font(bold=True, size=10)
    BODY_FONT = Font(size=10)
    CENTER    = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin      = Side(style="thin")
    BORDER    = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_header(ws, row_i):
        for cell in ws[row_i]:
            cell.fill  = HDR_FILL
            cell.font  = HDR_FONT
            cell.alignment = CENTER
            cell.border = BORDER

    def style_body(ws, row_i, bold=False):
        for cell in ws[row_i]:
            cell.font = SUM_FONT if bold else BODY_FONT
            cell.fill = SUM_FILL if bold else PatternFill()
            cell.alignment = CENTER
            cell.border = BORDER

    def auto_width(ws, max_w=25):
        for col in ws.columns:
            length = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(length + 2, max_w)

    T = stats["total_posts"]
    movies = MOVIE_ORDER_SORTED

    # ══════════════════════════════════════════════════════════════
    #  Table 1：六大类行为分布总表
    # ══════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "Table1_行为分布"

    hdr = ["行为类别（二阶编码）", "包含编码"] + \
          [f"{m}\n（内部占比%）" for m in movies] + ["全样本占比%"]
    ws1.append(hdr)
    style_header(ws1, 1)
    ws1.row_dimensions[1].height = 36

    for i, (zh, en, codes) in enumerate(CATEGORIES):
        row = [f"{i+1}. {zh}\n{en}", ", ".join(str(c) for c in sorted(codes))]
        for movie in movies:
            mp    = stats["movie_posts"].get(movie, 0)
            n_cat = stats["mv_cat_posts"].get(movie, [0]*N_CATS)[i]
            row.append(pct(n_cat, mp))
        row.append(pct(stats["cat_posts"][i], T))
        ws1.append(row)
        style_body(ws1, ws1.max_row)

    # 样本行数
    ws1.append(["样本量（帖子数）", "—"] +
               [stats["movie_posts"].get(m, 0) for m in movies] + [T])
    style_body(ws1, ws1.max_row, bold=True)

    ws1.column_dimensions["A"].width = 28
    ws1.column_dimensions["B"].width = 18
    for col_i in range(3, 3 + len(movies) + 1):
        ws1.column_dimensions[get_column_letter(col_i)].width = 16

    # ══════════════════════════════════════════════════════════════
    #  Table 2：行为频次与共鸣效率
    # ══════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Table2_共鸣效率")
    ws2.append(["行为类别（二阶编码）", "帖数", "帖数占比%",
                "互动总量", "互动占比%", "每帖均互动数", "效率比\n(互动%÷帖数%)"])
    style_header(ws2, 1)
    ws2.row_dimensions[1].height = 28

    TE = stats["total_eng"]
    for i, (zh, en, codes) in enumerate(CATEGORIES):
        n   = stats["cat_posts"][i]
        e   = stats["cat_eng"][i]
        pp  = pct(n, T)
        ep  = pct(e, TE)
        avg = round(e / n, 1) if n else 0.0
        eff = round(ep / pp, 3) if pp else 0.0
        ws2.append([f"{i+1}. {zh}", n, pp, e, ep, avg, eff])
        style_body(ws2, ws2.max_row)

    ws2.append(["全样本合计", T, 100.0, TE, 100.0,
                round(TE / T, 1) if T else 0, 1.0])
    style_body(ws2, ws2.max_row, bold=True)

    ws2.column_dimensions["A"].width = 26
    for col_i in range(2, 8):
        ws2.column_dimensions[get_column_letter(col_i)].width = 16

    # ══════════════════════════════════════════════════════════════
    #  Table 3：SIPs行为指纹对比
    # ══════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("Table3_SIPs指纹")

    SP = stats["sips_posts"]

    ws3.append(["指标", "SIPs子集（含code17）", "全样本", "SIPs/全样本比值"])
    style_header(ws3, 1)

    ws3.append(["帖子数", SP, T, round(SP / T, 4) if T else 0])
    style_body(ws3, ws3.max_row, bold=True)

    ws3.append(["── 二阶类别共现率（%） ──", "", "", ""])
    ws3[ws3.max_row][0].font = Font(bold=True, italic=True, size=10)

    for i, (zh, en, codes) in enumerate(CATEGORIES):
        sn  = stats["sips_cat_posts"][i]
        an  = stats["cat_posts"][i]
        s_r = pct(sn, SP)
        a_r = pct(an, T)
        ratio = round(s_r / a_r, 3) if a_r else 0.0
        ws3.append([f"{i+1}. {zh}", f"{s_r}%", f"{a_r}%", ratio])
        style_body(ws3, ws3.max_row)

    ws3.append(["── 关键一阶编码共现率（%） ──", "", "", ""])
    ws3[ws3.max_row][0].font = Font(bold=True, italic=True, size=10)

    for code in HIGHLIGHT_CODES:
        sn  = stats["sips_code_posts"].get(code, 0)
        an  = stats["all_code_posts"].get(code, 0)
        s_r = pct(sn, SP)
        a_r = pct(an, T)
        ratio = round(s_r / a_r, 3) if a_r else 0.0
        ws3.append([f"编码 {code}", f"{s_r}%", f"{a_r}%", ratio])
        style_body(ws3, ws3.max_row)

    ws3.column_dimensions["A"].width = 28
    for col_i in range(2, 5):
        ws3.column_dimensions[get_column_letter(col_i)].width = 22

    # ══════════════════════════════════════════════════════════════
    #  Table 4：SIPs跨电影分布
    # ══════════════════════════════════════════════════════════════
    ws4 = wb.create_sheet("Table4_SIPs电影分布")
    ws4.append(["电影", "首映日", "该电影总帖数", "含code17帖数", "SIPs内部占比%"])
    style_header(ws4, 1)

    for movie in movies:
        rel  = RELEASE_DATES_RAW.get(movie, "—")
        rel_s = rel.strftime("%Y-%m-%d") if isinstance(rel, datetime) else str(rel)
        mp   = stats["movie_posts"].get(movie, 0)
        sn   = stats["sips_by_movie"].get(movie, 0)
        ws4.append([movie, rel_s, mp, sn, pct(sn, mp)])
        style_body(ws4, ws4.max_row)

    total_sips = sum(stats["sips_by_movie"].get(m, 0) for m in movies)
    ws4.append(["合计", "—", T, total_sips, pct(total_sips, T)])
    style_body(ws4, ws4.max_row, bold=True)

    auto_width(ws4)

    wb.save(out_path)
    print(f"✅ Excel 已保存 → {out_path}")


# ─────────────────────────────────────────────────────────────────
#  Figure 1：四阶段共鸣结构折线图
# ─────────────────────────────────────────────────────────────────

def plot_figure1(stats: dict, fig_path: str):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.ticker as mticker
        import numpy as np
    except ImportError:
        print("⚠️  matplotlib 未安装，跳过 Figure 1。请运行: pip install matplotlib")
        return

    # avg engagement per post: phase × category
    phase_cat_posts = stats["phase_cat_posts"]
    phase_cat_eng   = stats["phase_cat_eng"]

    avg_eng = np.zeros((4, N_CATS))
    for ph in range(4):
        for cat in range(N_CATS):
            n = phase_cat_posts[ph][cat]
            e = phase_cat_eng[ph][cat]
            avg_eng[ph, cat] = e / n if n else np.nan

    # Color palette (colorblind-friendly)
    colors  = ["#E63946", "#457B9D", "#2A9D8F", "#E9C46A", "#F4A261", "#9B5DE5"]
    markers = ["o", "s", "^", "D", "v", "P"]

    fig, ax = plt.subplots(figsize=(9, 5.5))

    x = np.arange(4)
    for i, (zh, en, _) in enumerate(CATEGORIES):
        y = avg_eng[:, i]
        ax.plot(x, y, color=colors[i], marker=markers[i], linewidth=2,
                markersize=7, label=zh, zorder=3)

    ax.set_xticks(x)
    ax.set_xticklabels(PHASES, fontsize=11)
    ax.set_ylabel("每帖均互动数（转+评+赞）", fontsize=11)
    ax.set_xlabel("上映阶段", fontsize=11)
    ax.set_title("Figure 1  四阶段共鸣结构变化（全样本加权平均）", fontsize=13, pad=14)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(loc="upper right", fontsize=9, framealpha=0.8)

    plt.tight_layout()
    plt.savefig(fig_path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"✅ Figure 1 已保存 → {fig_path}")


# ─────────────────────────────────────────────────────────────────
#  控制台摘要
# ─────────────────────────────────────────────────────────────────

def print_summary(stats: dict):
    T  = stats["total_posts"]
    TE = stats["total_eng"]
    SP = stats["sips_posts"]
    print(f"\n{'='*64}")
    print(f"  全样本：{T:,} 帖  |  互动总量：{TE:,}  |  SIPs（含code17）：{SP:,}")
    print(f"{'='*64}")

    print("\n【Table 1 预览】六大类行为分布（全样本占比%）")
    print(f"  {'类别':<14} {'帖数':>8} {'全样本占比%':>12}")
    print("  " + "─"*36)
    for i, (zh, _, _) in enumerate(CATEGORIES):
        n = stats["cat_posts"][i]
        print(f"  {zh:<14} {n:>8,} {pct(n,T):>11.2f}%")

    print("\n【Table 2 预览】效率比（>1 = 高于平均共鸣效率）")
    TE = stats["total_eng"]
    print(f"  {'类别':<14} {'帖%':>8} {'互动%':>8} {'均互动':>10} {'效率比':>8}")
    print("  " + "─"*52)
    for i, (zh, _, _) in enumerate(CATEGORIES):
        n   = stats["cat_posts"][i]
        e   = stats["cat_eng"][i]
        pp  = pct(n, T)
        ep  = pct(e, TE)
        avg = round(e / n, 1) if n else 0
        eff = round(ep / pp, 3) if pp else 0
        print(f"  {zh:<14} {pp:>7.2f}% {ep:>7.2f}% {avg:>10,.0f} {eff:>8.3f}")

    print(f"\n【Table 4 预览】SIPs跨电影分布")
    print(f"  {'电影':<12} {'电影总帖':>10} {'SIPs帖':>8} {'占比%':>8}")
    print("  " + "─"*42)
    for movie in MOVIE_ORDER_SORTED:
        mp = stats["movie_posts"].get(movie, 0)
        sn = stats["sips_by_movie"].get(movie, 0)
        print(f"  {movie:<12} {mp:>10,} {sn:>8,} {pct(sn,mp):>7.2f}%")


# ─────────────────────────────────────────────────────────────────
#  主程序
# ─────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="学术分析：六大类 × 四阶段")
    parser.add_argument("--dir", default="data/tmp",
                        help="扫描根目录（递归查找 *_coded.json）")
    parser.add_argument("--out", default="academic_results.xlsx",
                        help="Excel 输出路径")
    parser.add_argument("--fig", default="figure1_phases.png",
                        help="Figure 1 PNG 输出路径")
    args = parser.parse_args()

    root = Path(args.dir)
    if not root.exists():
        print(f"❌ 目录不存在：{root}")
        return

    print(f"正在加载 {root} 下的编码文件...")
    posts = load_posts(root)
    if not posts:
        print("❌ 未找到任何帖子数据")
        return
    print(f"共加载 {len(posts):,} 条帖子，开始计算...")

    stats = build_stats(posts)
    print_summary(stats)
    export_excel(stats, args.out)
    plot_figure1(stats, args.fig)


if __name__ == "__main__":
    main()
