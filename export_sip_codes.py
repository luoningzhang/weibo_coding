"""
导出 SIP 行为编码帖子明细

每个 SIP 编码单独一个 sheet，最后一个 sheet 汇总所有 SIP 行为帖子（去重）。

SIP 编码范围：
  数据与市场劳动：2, 3, 4, 26
  策略性介入：    21, 22, 23, 24, 25, 27
  身份认同表达：  10, 17, 19
  话语防御：      14, 15, 16, 20

用法：
    python export_sip_codes.py
    python export_sip_codes.py --dir data/tmp --out sip_posts.xlsx
"""

import argparse
import json
from datetime import datetime
from pathlib import Path
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ── SIP 编码定义（顺序即 sheet 顺序）────────────────────────────
SIP_CODES = [
    (2,  "数据与市场劳动", "播报/分析票房数据"),
    (3,  "数据与市场劳动", "预测票房"),
    (4,  "数据与市场劳动", "分析排片/呼吁增加排片"),
    (26, "数据与市场劳动", "跨平台数据监控"),
    (21, "策略性介入",     "提出具体宣发建议"),
    (22, "策略性介入",     "分析竞争对手/市场环境"),
    (23, "策略性介入",     "讨论破圈策略"),
    (24, "策略性介入",     "主动联系/@院线/媒体"),
    (25, "策略性介入",     "地推活动"),
    (27, "策略性介入",     "海外/跨国宣发关注"),
    (10, "身份认同表达",   "组织/参与线下包场"),
    (17, "身份认同表达",   "自称精神股东/自来水"),
    (19, "身份认同表达",   "表达与主创共情"),
    (14, "话语防御",       "反驳差评/黑水"),
    (15, "话语防御",       "解释/维护电影"),
    (16, "话语防御",       "举报恶意言论"),
    (20, "话语防御",       "为电影逆袭/受委屈鸣不平"),
]
SIP_CODE_SET = {c for c, _, _ in SIP_CODES}

MOVIE_ORDER = ["流浪地球2", "消失的她", "封神", "长安三万里", "孤注一掷"]
FILTER_PATTERN = "今年我最爱的#微博年度电影#是"

# ── 样式 ────────────────────────────────────────────────────────
thin     = Side(style="thin")
BDR      = Border(left=thin, right=thin, top=thin, bottom=thin)
HDR_FILL = PatternFill("solid", fgColor="2F5597")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
CAT_FILLS = {
    "数据与市场劳动": PatternFill("solid", fgColor="DDEBF7"),
    "策略性介入":     PatternFill("solid", fgColor="E2EFDA"),
    "身份认同表达":   PatternFill("solid", fgColor="FFF2CC"),
    "话语防御":       PatternFill("solid", fgColor="FCE4D6"),
}
CENTER = Alignment(horizontal="center", vertical="top", wrap_text=False)
WRAP   = Alignment(horizontal="left",   vertical="top", wrap_text=True)
BODY   = Font(size=9)


def normalize_movie(raw: str) -> str:
    return raw.removesuffix("_电影").strip()


def engagement(p: dict) -> int:
    return (int(p.get("reposts_count") or 0)
            + int(p.get("comments_count") or 0)
            + int(p.get("attitudes_count") or 0))


def get_screen_name(p: dict) -> str:
    u = p.get("user")
    return u.get("screen_name", "") if isinstance(u, dict) else ""


def fmt_dt(raw: str) -> str:
    try:
        return datetime.strptime(raw, "%a %b %d %H:%M:%S +0800 %Y").strftime("%Y-%m-%d %H:%M")
    except Exception:
        return raw


def load_posts(root: Path) -> list[dict]:
    posts = []
    for path in sorted(root.rglob("*_coded.json")):
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        chunk = data if isinstance(data, list) else data.get(
            "data", data.get("posts", data.get("items", [])))
        movie_from_file = path.stem.removesuffix("_coded").removesuffix("_电影")
        for p in chunk:
            if not p.get("movie"):
                p["movie"] = movie_from_file
            posts.append(p)
    return posts


def movie_sort_key(p: dict) -> tuple:
    movie = normalize_movie(p.get("movie", ""))
    order = MOVIE_ORDER.index(movie) if movie in MOVIE_ORDER else 99
    return (order, -engagement(p))


HEADERS = ["序号", "电影", "用户昵称", "互动总量", "转发", "评论", "点赞",
           "发帖时间", "该帖全部编码", "帖子内容"]
COL_WIDTHS = [5, 12, 16, 10, 7, 7, 7, 17, 18, 70]


def write_sheet(ws, rows: list[dict], category: str, fill=None):
    ws.append(HEADERS)
    for cell in ws[1]:
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = CENTER
        cell.border = BDR
    ws.row_dimensions[1].height = 18
    ws.freeze_panes = "A2"

    row_fill = fill or CAT_FILLS.get(category, PatternFill())

    for idx, p in enumerate(rows, 1):
        movie  = normalize_movie(p.get("movie", ""))
        codes  = sorted({int(c) for c in (p.get("codes") or []) if str(c).isdigit()})
        text   = (p.get("text") or "").replace("\n", " ").replace("\r", "")
        eng    = engagement(p)

        ws.append([
            idx, movie, get_screen_name(p),
            eng,
            int(p.get("reposts_count") or 0),
            int(p.get("comments_count") or 0),
            int(p.get("attitudes_count") or 0),
            fmt_dt(p.get("created_at", "")),
            ", ".join(str(c) for c in codes),
            text,
        ])
        r = ws.max_row
        for ci, cell in enumerate(ws[r], 1):
            cell.font      = BODY
            cell.fill      = row_fill
            cell.border    = BDR
            cell.alignment = WRAP if ci == len(HEADERS) else CENTER
        ws.row_dimensions[r].height = 14

    for i, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def main():
    parser = argparse.ArgumentParser(description="导出SIP编码帖子明细")
    parser.add_argument("--dir", default="data/tmp")
    parser.add_argument("--out", default="sip_posts.xlsx")
    args = parser.parse_args()

    root = Path(args.dir)
    if not root.exists():
        print(f"❌ 目录不存在：{root}")
        return

    print("加载帖子数据...")
    all_posts = load_posts(root)
    print(f"共 {len(all_posts):,} 条帖子")

    # 过滤模板帖，预处理 codes
    posts = []
    for p in all_posts:
        if FILTER_PATTERN in (p.get("text") or ""):
            continue
        p["_codes_set"] = {int(c) for c in (p.get("codes") or []) if str(c).isdigit()}
        posts.append(p)

    # 按编码建索引
    code_posts: dict[int, list] = defaultdict(list)
    for p in posts:
        for c in p["_codes_set"]:
            if c in SIP_CODE_SET:
                code_posts[c].append(p)

    # 排序每个编码的帖子
    for c in code_posts:
        code_posts[c].sort(key=movie_sort_key)

    # 所有 SIP 帖子（去重，按电影+互动排序）
    seen_ids = set()
    all_sip = []
    for p in sorted(posts, key=movie_sort_key):
        pid = p.get("id") or id(p)
        if p["_codes_set"] & SIP_CODE_SET and pid not in seen_ids:
            seen_ids.add(pid)
            all_sip.append(p)

    wb = Workbook()
    first = True

    # ── 每个编码一个 sheet ───────────────────────────────────────
    for code, category, name in SIP_CODES:
        rows = code_posts.get(code, [])
        # sheet 名：Excel 上限31字符
        sheet_name = f"c{code:02d}_{name}"[:31]
        if first:
            ws = wb.active
            ws.title = sheet_name
            first = False
        else:
            ws = wb.create_sheet(sheet_name)

        ws.append([f"编码 {code}  {category}：{name}  （共 {len(rows)} 条）"])
        ws[1][0].font = Font(bold=True, size=11)
        ws[1][0].fill = CAT_FILLS.get(category, PatternFill())
        ws.row_dimensions[1].height = 20

        ws.append([])  # 空行
        write_sheet(ws, rows, category)
        print(f"  c{code:02d} {name:<16} {len(rows):>5} 条")

    # ── 最后：所有 SIP 行为帖子汇总 ─────────────────────────────
    ws_all = wb.create_sheet("★全部SIP行为帖子")
    ws_all.append([f"全部 SIP 行为帖子汇总（编码含任意SIP编码，共 {len(all_sip)} 条，已去重）"])
    ws_all[1][0].font = Font(bold=True, size=11)
    ws_all[1][0].fill = PatternFill("solid", fgColor="D9D9D9")
    ws_all.row_dimensions[1].height = 20
    ws_all.append([])
    write_sheet(ws_all, all_sip, "")
    print(f"\n  ★ 全部SIP帖子（去重）      {len(all_sip):>5} 条")

    wb.save(args.out)
    print(f"\n✅ 已导出 → {args.out}  （{len(SIP_CODES)} 个编码sheet + 1 个汇总sheet）")


if __name__ == "__main__":
    main()
