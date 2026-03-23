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


def analyze(posts: list[dict]) -> dict:
    total = len(posts)
    no_code = sum(1 for p in posts if not p.get("codes"))
    code_counter: Counter = Counter()
    for p in posts:
        for c in p.get("codes", []):
            code_counter[int(c)] += 1
    return {"total": total, "no_code": no_code, "counter": code_counter}


def print_report(results: list[tuple[str, dict]], code_names: dict[int, str]):
    sep = "=" * 70

    for movie, stats in results:
        print(f"\n{sep}")
        print(f"  电影：{movie}")
        print(sep)
        total    = stats["total"]
        no_code  = stats["no_code"]
        coded    = total - no_code
        print(f"  总条数：{total}　　已编码：{coded}　　无编码：{no_code}　　"
              f"编码覆盖率：{coded/total*100:.1f}%")
        print()
        if stats["counter"]:
            print(f"  {'编码':<6} {'名称':<18} {'出现次数':>8} {'占总条数%':>10}")
            print(f"  {'-'*46}")
            for code, cnt in sorted(stats["counter"].items()):
                name = code_names.get(code, "")
                pct  = cnt / total * 100
                print(f"  {code:<6} {name:<18} {cnt:>8} {pct:>9.1f}%")
        else:
            print("  （无编码数据）")

    print(f"\n{sep}")
    # 汇总
    total_all   = sum(s["total"] for _, s in results)
    no_code_all = sum(s["no_code"] for _, s in results)
    all_counter: Counter = Counter()
    for _, s in results:
        all_counter.update(s["counter"])

    print(f"  【全部电影汇总】总条数：{total_all}　　"
          f"无编码：{no_code_all}　　"
          f"覆盖率：{(total_all-no_code_all)/total_all*100:.1f}%")
    print()
    if all_counter:
        print(f"  {'编码':<6} {'名称':<18} {'出现次数':>8} {'占总条数%':>10}")
        print(f"  {'-'*46}")
        for code, cnt in sorted(all_counter.items()):
            name = code_names.get(code, "")
            pct  = cnt / total_all * 100
            print(f"  {code:<6} {name:<18} {cnt:>8} {pct:>9.1f}%")
    print(sep)


def export_excel(results: list[tuple[str, dict]], code_names: dict[int, str],
                 out_path: str):
    wb = Workbook()

    # ── Sheet1：各电影概览 ──────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "各电影概览"

    # 收集所有出现过的编码
    all_codes = sorted({c for _, s in results for c in s["counter"]})

    header = ["电影", "总条数", "已编码", "无编码", "覆盖率%"] + \
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
        row = [movie, total, coded, no_code,
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
    sum_row = ["【合计】", total_all, coded_all, no_code_all,
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
        ws2.cell(row_idx, 1, "编码")
        ws2.cell(row_idx, 2, "名称")
        ws2.cell(row_idx, 3, "次数")
        ws2.cell(row_idx, 4, "占总条数%")
        for cell in ws2[row_idx]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="BDD7EE")
        row_idx += 1
        for code, cnt in sorted(stats["counter"].items()):
            ws2.cell(row_idx, 1, code)
            ws2.cell(row_idx, 2, code_names.get(code, ""))
            ws2.cell(row_idx, 3, cnt)
            ws2.cell(row_idx, 4, round(cnt / total * 100, 1) if total else 0)
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
