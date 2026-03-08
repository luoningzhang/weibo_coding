"""
从微博 JSON 数据中，按照转发+点赞+评论总数从多到少排序，取前 200 条。

用法:
    python sort_weibo_top_posts.py input.json            # 输出 top200.json + top200.xlsx
    python sort_weibo_top_posts.py input.json output     # 输出 output.json + output.xlsx
"""

import json
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


def get_engagement(post: dict) -> int:
    """计算单条微博的转赞评总数。

    兼容常见的微博 JSON 字段命名：
      - 转发: reposts_count / retweet_count / reposts
      - 点赞: attitudes_count / likes_count / like_count / attitudes
      - 评论: comments_count / comment_count / comments
    """
    reposts = (
        post.get("reposts_count")
        or post.get("retweet_count")
        or post.get("reposts")
        or 0
    )
    likes = (
        post.get("attitudes_count")
        or post.get("likes_count")
        or post.get("like_count")
        or post.get("attitudes")
        or 0
    )
    comments = (
        post.get("comments_count")
        or post.get("comment_count")
        or post.get("comments")
        or 0
    )
    return int(reposts) + int(likes) + int(comments)


def sort_top_posts(posts: list[dict], top_n: int = 200) -> list[dict]:
    """按转赞评总数降序排序，返回前 top_n 条。"""
    sorted_posts = sorted(posts, key=get_engagement, reverse=True)
    return sorted_posts[:top_n]


def load_posts(path: str) -> list[dict]:
    """加载 JSON 文件，支持顶层为列表或包含 data/posts 字段的对象。"""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("data", "posts", "statuses", "items"):
            if key in data and isinstance(data[key], list):
                return data[key]
    raise ValueError(
        f"无法识别 JSON 结构，期望顶层为列表或含 data/posts/statuses/items 字段的对象，实际得到: {type(data)}"
    )


def export_excel(top_posts: list[dict], path: str) -> None:
    """将 top_posts 导出为 Excel，id 列强制写为文本防止大数被四舍五入。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Top200"

    headers = [
        ("排名", "_rank"),
        ("微博ID", "id"),
        ("发布时间", "created_at"),
        ("用户名", "_user_name"),
        ("正文", "text"),
        ("转发数", "reposts_count"),
        ("评论数", "comments_count"),
        ("点赞数", "attitudes_count"),
        ("转赞评总计", "_engagement_total"),
        ("标签", "predicted_label"),
        ("电影", "movie"),
    ]

    # 表头样式
    header_fill = PatternFill("solid", fgColor="4F81BD")
    header_font = Font(bold=True, color="FFFFFF")
    for col, (label, _) in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=label)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 数据行
    for row, post in enumerate(top_posts, start=2):
        user = post.get("user") or {}
        user_name = user.get("screen_name", "") if isinstance(user, dict) else str(user)

        row_data = {
            "_rank": post.get("_rank"),
            "id": str(post.get("id") or post.get("mid") or post.get("idstr", "")),
            "created_at": post.get("created_at", ""),
            "_user_name": user_name,
            "text": post.get("text") or post.get("content", ""),
            "reposts_count": post.get("reposts_count", 0),
            "comments_count": post.get("comments_count", 0),
            "attitudes_count": post.get("attitudes_count", 0),
            "_engagement_total": post.get("_engagement_total"),
            "predicted_label": post.get("predicted_label", ""),
            "movie": post.get("movie", ""),
        }

        for col, (_, key) in enumerate(headers, start=1):
            value = row_data[key]
            cell = ws.cell(row=row, column=col, value=value)
            # id 列：明确设为文本格式，彻底防止 Excel 四舍五入
            if key == "id":
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="left")

    # 自适应列宽
    col_widths = [6, 22, 26, 16, 60, 10, 10, 10, 12, 8, 20]
    for col, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    # 冻结首行
    ws.freeze_panes = "A2"

    wb.save(path)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    input_path = sys.argv[1]
    base = sys.argv[2] if len(sys.argv) > 2 else "top200"
    # 去掉用户传入的扩展名，统一处理
    if base.endswith(".json"):
        base = base[:-5]
    json_path = base + ".json"
    xlsx_path = base + ".xlsx"

    posts = load_posts(input_path)
    print(f"共加载 {len(posts)} 条微博")

    top_posts = sort_top_posts(posts, top_n=200)
    print(f"筛选后保留 {len(top_posts)} 条")

    for rank, post in enumerate(top_posts, start=1):
        post["_rank"] = rank
        post["_engagement_total"] = get_engagement(post)

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(top_posts, f, ensure_ascii=False, indent=2)
    print(f"JSON 已写入: {json_path}")

    export_excel(top_posts, xlsx_path)
    print(f"Excel 已写入: {xlsx_path}")

    print("\n前 5 名预览:")
    for post in top_posts[:5]:
        mid = post.get("id") or post.get("mid") or post.get("idstr", "?")
        text = post.get("text") or post.get("content") or ""
        print(
            f"  #{post['_rank']} | 总互动={post['_engagement_total']} | id={mid} | {text[:40]}..."
        )


if __name__ == "__main__":
    main()
