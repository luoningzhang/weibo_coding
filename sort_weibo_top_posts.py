"""
从微博 JSON 数据中，按照转发+点赞+评论总数从多到少排序，取前 200 条。

用法:
    python sort_weibo_top_posts.py input.json output.json
    python sort_weibo_top_posts.py input.json              # 输出到 top200.json
"""

import json
import sys


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


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "top200.json"

    posts = load_posts(input_path)
    print(f"共加载 {len(posts)} 条微博")

    top_posts = sort_top_posts(posts, top_n=200)
    print(f"筛选后保留 {len(top_posts)} 条")

    # 附加排名和转赞评汇总字段，方便后续使用
    for rank, post in enumerate(top_posts, start=1):
        post["_rank"] = rank
        post["_engagement_total"] = get_engagement(post)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(top_posts, f, ensure_ascii=False, indent=2)

    print(f"结果已写入: {output_path}")
    print("\n前 5 名预览:")
    for post in top_posts[:5]:
        mid = post.get("id") or post.get("mid") or post.get("idstr", "?")
        text = post.get("text") or post.get("content") or ""
        print(
            f"  #{post['_rank']} | 总互动={post['_engagement_total']} | id={mid} | {text[:40]}..."
        )


if __name__ == "__main__":
    main()
