"""
学术分析脚本：六大行为类别 × 三阶段参与结构

输出：
  Table 1 — 六大类 × 28编码行为分布总表（电影内部占比 + 全样本占比）
  Table 2 — 行为频次与共鸣效率对比（帖数占比 / 互动占比 / 均互动 / 效率比）
  Figure 1 — 三阶段共鸣结构变化折线图（PNG）
  Table 3 — SIPs行为指纹对比（含code17 vs 全样本）
  Table 4 — SIPs跨电影分布

用法：
    python academic_analysis.py
    python academic_analysis.py --dir data/tmp --out academic_results.xlsx --fig figure1.png
    python academic_analysis.py --codebook data/codebook.xlsx
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

# ── 三阶段定义（相对首映日） ──────────────────────────────────────
PHASES = ["上映前期", "上映初期", "长尾期"]
PHASE_RANGES = [(-30, -1), (1, 30), (31, 120)]   # [days from release]
N_PHASES = len(PHASES)

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


def load_codebook(path: str) -> dict[int, str]:
    """编码号 → 名称（取 data/codebook.xlsx 第二列）"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path)
        ws = wb.active
        mapping = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            cells = list(row) + [None, None]
            num, name = cells[0], cells[1]
            if num is not None:
                try:
                    mapping[int(num)] = str(name).strip() if name else ""
                except (ValueError, TypeError):
                    pass
        return mapping
    except Exception:
        return {}


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

    # three-phase × category
    # phase_cat_posts[phase][cat] -> posts count
    # phase_cat_eng[phase][cat]   -> engagement sum
    phase_cat_posts = [[0] * N_CATS for _ in range(N_PHASES)]
    phase_cat_eng   = [[0] * N_CATS for _ in range(N_PHASES)]

    # per-movie × three-phase × category
    mv_phase_cat_posts = {m: [[0]*N_CATS for _ in range(N_PHASES)] for m in MOVIE_ORDER_SORTED}
    mv_phase_cat_eng   = {m: [[0]*N_CATS for _ in range(N_PHASES)] for m in MOVIE_ORDER_SORTED}

    # per-movie per-code post counts (for Table 1 code-level detail)
    mv_code_posts = defaultdict(lambda: defaultdict(int))

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
                    if movie in mv_phase_cat_posts:
                        mv_phase_cat_posts[movie][phase_i][i] += 1
                        mv_phase_cat_eng[movie][phase_i][i]   += eng

        codes_set = set(codes)
        for c in codes_set:
            all_code_posts[c] += 1
            mv_code_posts[movie][c] += 1

        is_sip = SIPS_CODE in codes_set
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
        "mv_code_posts":       {k: dict(v) for k, v in mv_code_posts.items()},
        "mv_phase_cat_posts":  mv_phase_cat_posts,
        "mv_phase_cat_eng":    mv_phase_cat_eng,
    }


# ─────────────────────────────────────────────────────────────────
#  Excel 输出
# ─────────────────────────────────────────────────────────────────

def pct(num, denom):
    return round(num / denom * 100, 2) if denom else 0.0


def export_excel(stats: dict, out_path: str, codebook: dict[int, str] | None = None):
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
    cb = codebook or {}

    # 行样式用到的额外颜色
    CAT_FILL   = PatternFill("solid", fgColor="C5D9F1")   # 类别小计行
    CODE_FILL  = PatternFill("solid", fgColor="F2F2F2")   # 一阶编码行（偶数）
    CODE_FILL2 = PatternFill("solid", fgColor="FFFFFF")   # 一阶编码行（奇数）
    CAT_FONT   = Font(bold=True, size=10)
    CODE_FONT  = Font(size=10)
    LEFT_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=False)

    # ══════════════════════════════════════════════════════════════
    #  Table 1：六大类 × 28编码行为分布总表
    # ══════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "Table1_行为分布"

    # 表头：两行（合并第一行为大标题 + 第二行列名）
    hdr = ["行为类别 / 一阶编码", "编码号", "编码名称"] + \
          [f"{m}\n内部占比%" for m in movies] + ["全样本\n帖数", "全样本\n占比%"]
    ws1.append(hdr)
    style_header(ws1, 1)
    ws1.row_dimensions[1].height = 36

    # 冻结首行
    ws1.freeze_panes = "A2"

    for cat_i, (zh, en, cat_codes_set) in enumerate(CATEGORIES):
        cat_codes_sorted = sorted(cat_codes_set)
        code_row_count = 0

        # ── 一阶编码行 ──────────────────────────────────────────
        for code in cat_codes_sorted:
            code_name = cb.get(code, "")
            row = [f"  └ {zh}", str(code), code_name]
            for movie in movies:
                mp = stats["movie_posts"].get(movie, 0)
                n  = stats["mv_code_posts"].get(movie, {}).get(code, 0)
                row.append(pct(n, mp))
            global_n = stats["all_code_posts"].get(code, 0)
            row += [global_n, pct(global_n, T)]
            ws1.append(row)
            r = ws1.max_row
            fill = CODE_FILL if code_row_count % 2 == 0 else CODE_FILL2
            for cell in ws1[r]:
                cell.font      = CODE_FONT
                cell.fill      = fill
                cell.border    = BORDER
                cell.alignment = CENTER
            ws1[r][0].alignment = LEFT_ALIGN   # 类别名左对齐
            code_row_count += 1

        # ── 类别小计行 ──────────────────────────────────────────
        row = [f"▶ {cat_i+1}. {zh}（小计）", "—", "—"]
        for movie in movies:
            mp    = stats["movie_posts"].get(movie, 0)
            n_cat = stats["mv_cat_posts"].get(movie, [0]*N_CATS)[cat_i]
            row.append(pct(n_cat, mp))
        global_n_cat = stats["cat_posts"][cat_i]
        row += [global_n_cat, pct(global_n_cat, T)]
        ws1.append(row)
        r = ws1.max_row
        for cell in ws1[r]:
            cell.font      = CAT_FONT
            cell.fill      = CAT_FILL
            cell.border    = BORDER
            cell.alignment = CENTER
        ws1[r][0].alignment = LEFT_ALIGN
        ws1.row_dimensions[r].height = 18

    # ── 样本量汇总行 ──────────────────────────────────────────
    ws1.append(["【样本量（帖子数）】", "", ""] +
               [stats["movie_posts"].get(m, 0) for m in movies] + [T, 100.0])
    style_body(ws1, ws1.max_row, bold=True)

    # 列宽
    ws1.column_dimensions["A"].width = 24
    ws1.column_dimensions["B"].width = 8
    ws1.column_dimensions["C"].width = 20
    for col_i in range(4, 4 + len(movies)):
        ws1.column_dimensions[get_column_letter(col_i)].width = 14
    ws1.column_dimensions[get_column_letter(4 + len(movies))].width     = 12
    ws1.column_dimensions[get_column_letter(4 + len(movies) + 1)].width = 12

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

    # ══════════════════════════════════════════════════════════════
    #  Figure 1 底层数据：电影 × 阶段 × 类别（帖数 / 互动 / 均互动）
    # ══════════════════════════════════════════════════════════════
    ws5 = wb.create_sheet("Figure1_数据")

    # ── 宽表：列 = 阶段×类别，行 = 电影 ───────────────────────────
    # 第一行：大标题（阶段）
    # 第二行：子标题（类别指标）
    # 后续行：各电影 + 合计

    # 构建列顺序：每个阶段下6个类别，各3个指标（帖数/互动/均互动）
    # 但这会导致很多列。改用"长表"格式更适合论文分析。

    # ── 长表（tidy format）──────────────────────────────────────
    ws5.append(["电影 (Movie)", "电影英文", "阶段 (Phase)", "阶段英文",
                "行为类别 (Category)", "类别英文",
                "帖数 (Posts)", "互动总量 (Total Eng.)", "每帖均互动 (Avg Eng/Post)"])
    style_header(ws5, 1)
    ws5.row_dimensions[1].height = 28

    PHASE_EN_SHORT = ["Pre-Release", "Opening Period (D1-30)", "Long Tail (D31-120)"]
    MOVIE_EN_MAP   = {**MOVIE_EN, "All Films": "All Films (Weighted Avg.)"}

    # 5 movies
    for movie in movies:
        mv_pp = stats["mv_phase_cat_posts"].get(movie, [[0]*N_CATS]*N_PHASES)
        mv_pe = stats["mv_phase_cat_eng"].get(movie,   [[0]*N_CATS]*N_PHASES)
        for ph, (ph_zh, ph_en) in enumerate(zip(PHASES, PHASE_EN_SHORT)):
            for ci, (cat_zh, cat_en, _) in enumerate(CATEGORIES):
                n   = mv_pp[ph][ci]
                e   = mv_pe[ph][ci]
                avg = round(e / n, 2) if n else ""
                ws5.append([movie, MOVIE_EN.get(movie, movie), ph_zh, ph_en,
                             f"{ci+1}. {cat_zh}", cat_en, n, e, avg])
                style_body(ws5, ws5.max_row)

    # total (weighted average across all films)
    for ph, (ph_zh, ph_en) in enumerate(zip(PHASES, PHASE_EN_SHORT)):
        for ci, (cat_zh, cat_en, _) in enumerate(CATEGORIES):
            n   = stats["phase_cat_posts"][ph][ci]
            e   = stats["phase_cat_eng"][ph][ci]
            avg = round(e / n, 2) if n else ""
            ws5.append(["全部电影（合计）", "All Films", ph_zh, ph_en,
                         f"{ci+1}. {cat_zh}", cat_en, n, e, avg])
            r = ws5.max_row
            for cell in ws5[r]:
                cell.font = SUM_FONT
                cell.fill = SUM_FILL
                cell.border = BORDER
                cell.alignment = CENTER

    auto_width(ws5)
    # 加宽类别列
    ws5.column_dimensions["E"].width = 22
    ws5.column_dimensions["F"].width = 28

    wb.save(out_path)
    print(f"✅ Excel 已保存 → {out_path}")


# ─────────────────────────────────────────────────────────────────
#  Figure 1：四阶段共鸣结构折线图
# ─────────────────────────────────────────────────────────────────

# English labels for movies (used in Figure 1)
MOVIE_EN = {
    "流浪地球2":  "Wandering Earth 2",
    "消失的她":   "Lost in the Stars",
    "封神":       "Creation of the Gods",
    "长安三万里": "Chang'an",
    "孤注一掷":   "No More Bets",
}
PHASES_EN = ["Pre-Release\n(d-30 to d-1)", "Opening Period\n(d1–d30)", "Long Tail\n(d31–d120)"]
CAT_EN    = [c[1] for c in CATEGORIES]   # English short names already in CATEGORIES


def _avg_eng_matrix(phase_posts: list, phase_eng: list):
    """Return N_PHASES × N_CATS array of avg engagement (nan where no data)."""
    import numpy as np
    mat = np.zeros((N_PHASES, N_CATS))
    for ph in range(N_PHASES):
        for cat in range(N_CATS):
            n = phase_posts[ph][cat]
            e = phase_eng[ph][cat]
            mat[ph, cat] = e / n if n else float("nan")
    return mat


def _draw_panel(ax, avg_eng, title: str, colors, markers, x):
    """Draw one phase-line panel onto ax."""
    import numpy as np
    import matplotlib.ticker as mticker

    lines = []
    for i, en in enumerate(CAT_EN):
        y = avg_eng[:, i]
        ln, = ax.plot(x, y, color=colors[i], marker=markers[i],
                      linewidth=1.8, markersize=6, zorder=3)
        lines.append(ln)

    ax.set_xticks(x)
    ax.set_xticklabels(PHASES_EN, fontsize=7.5)
    ax.set_ylabel("Avg. Engagement per Post", fontsize=8)
    ax.set_title(title, fontsize=9, pad=5, fontweight="bold")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.grid(axis="y", linestyle="--", alpha=0.35, zorder=0)
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(axis="both", labelsize=7.5)
    return lines


def plot_figure1(stats: dict, fig_path: str):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import numpy as np
    except ImportError:
        print("⚠️  matplotlib not installed. Run: pip install matplotlib numpy")
        return

    plt.rcParams.update({
        "font.family":        "DejaVu Sans",
        "axes.unicode_minus": False,
    })

    colors  = ["#E63946", "#457B9D", "#2A9D8F", "#E9C46A", "#F4A261", "#9B5DE5"]
    markers = ["o", "s", "^", "D", "v", "P"]
    x       = np.arange(N_PHASES)

    # 6 panels: 5 movies + 1 total, arranged 2 rows × 3 columns
    panel_data = []
    for movie in MOVIE_ORDER_SORTED:
        mat = _avg_eng_matrix(stats["mv_phase_cat_posts"][movie],
                              stats["mv_phase_cat_eng"][movie])
        panel_data.append((MOVIE_EN.get(movie, movie), mat))
    # total (weighted average across all movies)
    mat_total = _avg_eng_matrix(stats["phase_cat_posts"], stats["phase_cat_eng"])
    panel_data.append(("All Films (Weighted Avg.)", mat_total))

    # compute per-panel y-max: each panel auto-scales to its own data
    # (keeps within-movie trends clearly readable)

    fig, axes = plt.subplots(2, 3, figsize=(13, 7.5),
                             gridspec_kw={"hspace": 0.55, "wspace": 0.38})
    axes_flat = axes.flatten()

    legend_lines = None
    for idx, (title, mat) in enumerate(panel_data):
        lines = _draw_panel(axes_flat[idx], mat, title, colors, markers, x)
        # let each panel's y-axis scale to its own data range
        if legend_lines is None:
            legend_lines = lines

    # shared legend below all subplots, outside the plot area
    fig.legend(
        legend_lines, CAT_EN,
        loc="lower center",
        ncol=3,
        fontsize=8,
        framealpha=0.9,
        bbox_to_anchor=(0.5, -0.08),
        title="Behavioral Category",
        title_fontsize=8.5,
    )

    plt.savefig(fig_path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"✅ Figure 1 saved → {fig_path}")


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
    parser = argparse.ArgumentParser(description="学术分析：六大类 × 三阶段")
    parser.add_argument("--dir",      default="data/tmp",
                        help="扫描根目录（递归查找 *_coded.json）")
    parser.add_argument("--out",      default="academic_results.xlsx",
                        help="Excel 输出路径")
    parser.add_argument("--fig",      default="figure1_phases.png",
                        help="Figure 1 PNG 输出路径")
    parser.add_argument("--codebook", default="data/codebook.xlsx",
                        help="编码手册路径（用于 Table 1 显示编码名称）")
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

    cb = load_codebook(args.codebook)
    if cb:
        print(f"编码手册已加载：{len(cb)} 个编码名称")
    else:
        print(f"⚠️  未加载编码手册（{args.codebook}），Table1 将只显示编码号")

    stats = build_stats(posts)
    print_summary(stats)
    export_excel(stats, args.out, codebook=cb)
    plot_figure1(stats, args.fig)


if __name__ == "__main__":
    main()
