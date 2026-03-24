"""
编码结果分析脚本

用法:
    python analyze_codes.py                        # 扫描 data/tmp/ 下所有 *_coded.json
    python analyze_codes.py --dir data/tmp         # 指定根目录
    python analyze_codes.py --file data/out.json   # 单文件
    python analyze_codes.py --codebook data/codebook.xlsx  # 显示编码名称（默认自动加载）
    python analyze_codes.py --excel result_analysis.xlsx   # 导出 Excel
"""

import argparse
import json
from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


def load_codebook(path: str) -> dict[int, str]:
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path)
        ws = wb.active
        mapping = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            num, name = (list(row) + [None, None])[:2]
            if num is not None:
                mapping[int(num)] = str(name).strip() if name else ""
        return mapping
    except Exception:
        return {}


def load_coded_file(path: Path) -> list[dict]:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, list):
        return data
    for key in ("data", "posts", "items"):
        if key in data and isinstance(data[key], list):
            return data[key]
    return []


def find_coded_files(root: Path) -> list[Path]:
    return sorted(root.rglob("*_coded.json"))


def movie_name(path: Path) -> str:
    # 从文件名提取电影名，去掉 _电影_coded 后缀
    stem = path.stem  # e.g. 流浪地球2_电影_coded
    for suffix in ("_电影_coded", "_coded"):
        if stem.endswith(suffix):
            return stem[: -len(suffix)]
    return stem


def get_user_id(post: dict):
    u = post.get("user")
    if isinstance(u, dict):
        return u.get("id") or u.get("user_id")
    return u or post.get("user_id")


def engagement(post: dict) -> int:
    v = post.get("engagement_total")
    if v is not None:
        return int(v)
    return (int(post.get("reposts_count", 0) or 0) +
            int(post.get("comments_count", 0) or 0) +
            int(post.get("attitudes_count", 0) or 0))


def analyze(posts: list[dict]) -> dict:
    total = len(posts)
    no_code = sum(1 for p in posts if not p.get("codes"))
    code_counter: Counter = Counter()
    code_reposts:   dict[int, int] = {}
    code_comments:  dict[int, int] = {}
    code_attitudes: dict[int, int] = {}
    code_engagement: dict[int, int] = {}
    for p in posts:
        rep = int(p.get("reposts_count",   0) or 0)
        com = int(p.get("comments_count",  0) or 0)
        att = int(p.get("attitudes_count", 0) or 0)
        eng = engagement(p)
        for c in p.get("codes", []):
            c = int(c)
            code_counter[c]    += 1
            code_reposts[c]    = code_reposts.get(c, 0)    + rep
            code_comments[c]   = code_comments.get(c, 0)   + com
            code_attitudes[c]  = code_attitudes.get(c, 0)  + att
            code_engagement[c] = code_engagement.get(c, 0) + eng
    unique_users = len({get_user_id(p) for p in posts if get_user_id(p) is not None})
    return {"total": total, "no_code": no_code, "counter": code_counter,
            "unique_users": unique_users,
            "reposts": code_reposts, "comments": code_comments,
            "attitudes": code_attitudes, "engagement": code_engagement}


def print_report(results: list[tuple[str, dict]], code_names: dict[int, str]):
    sep = "=" * 70

    for movie, stats in results:
        print(f"\n{sep}")
        print(f"  电影：{movie}")
        print(sep)
        total    = stats["total"]
        no_code  = stats["no_code"]
        coded    = total - no_code
        print(f"  总条数：{total}　　不同用户：{stats['unique_users']}　　"
              f"已编码：{coded}　　无编码：{no_code}　　"
              f"编码覆盖率：{coded/total*100:.1f}%")
        print()
        if stats["counter"]:
            print(f"  {'编码':<6} {'名称':<16} {'条数':>6} {'占%':>6} {'转发':>10} {'评论':>10} {'点赞':>10} {'转赞评总':>12} {'均转赞评':>10}")
            print(f"  {'-'*90}")
            for code, cnt in sorted(stats["counter"].items()):
                name = code_names.get(code, "")
                pct  = cnt / total * 100
                rep  = stats["reposts"].get(code, 0)
                com  = stats["comments"].get(code, 0)
                att  = stats["attitudes"].get(code, 0)
                eng  = stats["engagement"].get(code, 0)
                avg  = eng / cnt if cnt else 0
                print(f"  {code:<6} {name:<16} {cnt:>6} {pct:>5.1f}% {rep:>10,} {com:>10,} {att:>10,} {eng:>12,} {avg:>10,.0f}")
        else:
            print("  （无编码数据）")

    print(f"\n{sep}")
    # 汇总
    total_all   = sum(s["total"] for _, s in results)
    no_code_all = sum(s["no_code"] for _, s in results)
    all_counter: Counter = Counter()
    for _, s in results:
        all_counter.update(s["counter"])

    unique_users_all = sum(s["unique_users"] for _, s in results)
    print(f"  【全部电影汇总】总条数：{total_all}　　不同用户：{unique_users_all}　　"
          f"无编码：{no_code_all}　　"
          f"覆盖率：{(total_all-no_code_all)/total_all*100:.1f}%")
    print()
    all_reposts:   dict[int, int] = {}
    all_comments:  dict[int, int] = {}
    all_attitudes: dict[int, int] = {}
    all_eng:       dict[int, int] = {}
    for _, s in results:
        for c in s["counter"]:
            all_reposts[c]   = all_reposts.get(c, 0)   + s["reposts"].get(c, 0)
            all_comments[c]  = all_comments.get(c, 0)  + s["comments"].get(c, 0)
            all_attitudes[c] = all_attitudes.get(c, 0) + s["attitudes"].get(c, 0)
            all_eng[c]       = all_eng.get(c, 0)       + s["engagement"].get(c, 0)
    if all_counter:
        print(f"  {'编码':<6} {'名称':<16} {'条数':>6} {'占%':>6} {'转发':>10} {'评论':>10} {'点赞':>10} {'转赞评总':>12} {'均转赞评':>10}")
        print(f"  {'-'*90}")
        for code, cnt in sorted(all_counter.items()):
            name = code_names.get(code, "")
            pct  = cnt / total_all * 100
            rep  = all_reposts.get(code, 0)
            com  = all_comments.get(code, 0)
            att  = all_attitudes.get(code, 0)
            eng  = all_eng.get(code, 0)
            avg  = eng / cnt if cnt else 0
            print(f"  {code:<6} {name:<16} {cnt:>6} {pct:>5.1f}% {rep:>10,} {com:>10,} {att:>10,} {eng:>12,} {avg:>10,.0f}")
    print(sep)


def export_excel(results: list[tuple[str, dict]], code_names: dict[int, str],
                 out_path: str):
    wb = Workbook()

    # ── Sheet1：各电影概览 ──────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "各电影概览"

    # 收集所有出现过的编码
    all_codes = sorted({c for _, s in results for c in s["counter"]})

    header = ["电影", "总条数", "不同用户", "已编码", "无编码", "覆盖率%"] + \
             [f"{c} {code_names.get(c,'')}" for c in all_codes]
    ws1.append(header)

    # 表头样式
    hdr_fill = PatternFill("solid", fgColor="4472C4")
    hdr_font = Font(bold=True, color="FFFFFF")
    for cell in ws1[1]:
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center")

    for movie, stats in results:
        total   = stats["total"]
        no_code = stats["no_code"]
        coded   = total - no_code
        row = [movie, total, stats["unique_users"], coded, no_code,
               round(coded / total * 100, 1) if total else 0]
        for c in all_codes:
            row.append(stats["counter"].get(c, 0))
        ws1.append(row)

    # 汇总行
    all_counter: Counter = Counter()
    total_all = no_code_all = 0
    for _, s in results:
        total_all   += s["total"]
        no_code_all += s["no_code"]
        all_counter.update(s["counter"])
    coded_all = total_all - no_code_all
    sum_row = ["【合计】", total_all, sum(s["unique_users"] for _, s in results),
               coded_all, no_code_all,
               round(coded_all / total_all * 100, 1) if total_all else 0]
    for c in all_codes:
        sum_row.append(all_counter.get(c, 0))
    ws1.append(sum_row)
    for cell in ws1[ws1.max_row]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")

    # 自动列宽
    for col in ws1.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=0)
        ws1.column_dimensions[col[0].column_letter].width = min(max_len + 2, 30)

    # ── Sheet2：编码频率明细（每个电影一小块） ─────────────────────
    ws2 = wb.create_sheet("编码频率明细")
    row_idx = 1
    for movie, stats in results:
        total = stats["total"]
        # 小标题
        ws2.cell(row_idx, 1, f"▶ {movie}（总 {total} 条）").font = Font(bold=True)
        row_idx += 1
        for col, hdr in enumerate(["编码", "名称", "次数", "占总条数%",
                                    "转发总数", "评论总数", "点赞总数", "转赞评总数", "均转赞评"], 1):
            ws2.cell(row_idx, col, hdr).font = Font(bold=True)
            ws2.cell(row_idx, col).fill = PatternFill("solid", fgColor="BDD7EE")
        row_idx += 1
        for code, cnt in sorted(stats["counter"].items()):
            rep = stats["reposts"].get(code, 0)
            com = stats["comments"].get(code, 0)
            att = stats["attitudes"].get(code, 0)
            eng = stats["engagement"].get(code, 0)
            for col, val in enumerate([code, code_names.get(code, ""), cnt,
                                        round(cnt / total * 100, 1) if total else 0,
                                        rep, com, att, eng,
                                        round(eng / cnt, 0) if cnt else 0], 1):
                ws2.cell(row_idx, col, val)
            row_idx += 1
        row_idx += 1  # 空行

    for col in ws2.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=0)
        ws2.column_dimensions[col[0].column_letter].width = min(max_len + 2, 30)

    wb.save(out_path)
    print(f"\n✅ Excel 已导出 → {out_path}")


def main():
    parser = argparse.ArgumentParser(description="编码结果分析")
    parser.add_argument("--dir",      default="data/tmp",
                        help="扫描根目录（递归查找 *_coded.json）")
    parser.add_argument("--file",     help="直接指定单个 coded.json")
    parser.add_argument("--codebook", default="data/codebook.xlsx")
    parser.add_argument("--excel",    default="",
                        help="导出 Excel 路径，如 result_analysis.xlsx")
    args = parser.parse_args()

    code_names = load_codebook(args.codebook)

    if args.file:
        files = [Path(args.file)]
    else:
        root  = Path(args.dir)
        files = find_coded_files(root)
        if not files:
            print(f"❌ 在 {root} 下未找到任何 *_coded.json 文件")
            return

    results = []
    for f in files:
        posts = load_coded_file(f)
        stats = analyze(posts)
        name  = movie_name(f)
        results.append((name, stats))
        print(f"✔ {name}: {stats['total']} 条")

    print_report(results, code_names)

    if args.excel:
        export_excel(results, code_names, args.excel)


if __name__ == "__main__":
    main()
