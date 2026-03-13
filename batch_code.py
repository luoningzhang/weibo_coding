"""
批量编码微博（云雾 API / claude-sonnet-4-6）

用法:
    python batch_code.py

可选参数:
    --input     data/no_verified_classified_with_movie.json
    --output    data/coded_output.json
    --codebook  data/codebook.xlsx
    --batch     20      每次请求的条数（默认 20）
    --workers   5       并发数
    --show-prompt       打印 system prompt 和示例 user 消息后退出
    --test              跑前200条，与 top200-coding.xlsx 对比，打印每条差异及汇总统计后退出

API Key 放在 config.json 的 api_key 字段中。
"""

import argparse
import json
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from openpyxl import load_workbook
from openai import OpenAI

YUNWU_BASE_URL = "https://yunwu.ai/v1"
MODEL = "claude-sonnet-4-6-thinking"
MAX_RETRIES = 3
MAX_TOKENS = 6000   # thinking budget 4000 + 输出最多约2000（20条JSON编号）
THINKING_BUDGET = 4000  # 分类任务不需要深度推理，4000 thinking token 足够
CONFIG_PATH = Path(__file__).parent / "config.json"


def load_config() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def load_codebook(path: str) -> list[dict]:
    wb = load_workbook(path)
    ws = wb.active
    codes = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        num, name, definition, criteria, note = (list(row) + [None] * 5)[:5]
        if num is None:
            continue
        codes.append({
            "num": int(num),
            "name": str(name).strip(),
            "definition": str(definition).strip() if definition else "",
            "criteria": str(criteria).strip() if criteria else "",
            "note": str(note).strip() if note else "",
        })
    return codes


def load_top200_coding(path: str) -> dict[str, list[int]]:
    """读取 top200-coding.xlsx，返回 {微博ID: [编码编号...]}"""
    import re
    wb = load_workbook(path)
    ws = wb.active
    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        weibo_id = str(row[1]).strip() if row[1] is not None else None
        coding_str = row[12]  # 适用编码列
        if not weibo_id:
            continue
        if coding_str:
            nums = [int(n) for n in re.findall(r'\b(\d+)\b', str(coding_str))]
        else:
            nums = []
        result[weibo_id] = nums
    return result


def build_system_prompt(codes: list[dict]) -> str:
    lines = [
        "你是一位内容分析员，负责对微博文本按行为类别编码。",
        "每条微博可对应 0 到多个编码。只返回 JSON，不要任何解释。",
        "",
        "## 编码列表",
    ]
    for c in codes:
        lines.append(f"\n### {c['num']} {c['name']}")
        lines.append(f"操作定义：{c['definition']}")
        lines.append(f"识别标准：{c['criteria']}")
        if c["note"]:
            lines.append(f"备注：{c['note']}")
    lines += [
        "",
        "## 编码规则与常见易错",
        "",
        "### 开始编码前，先扫描以下高频信号",
        "在对每条微博开始编码之前，先逐一检查以下信号是否存在：",
        '- 【11 安利】正文含"快去看/强推/一定要去/值得一看/安利/带朋友去看"等 → 必须加11',
        '- 【17 自来水】正文含"自来水/精神股东/不是粉丝/路人" → 必须加17',
        "- 【19 共情】正文对导演/演员/幕后有心疼/感动/感谢/敬佩 → 必须加19",
        "- 【22 市场分析】正文同时提及两部或以上影片并进行比较 → 必须加22",
        "",
        "### 各代码精确边界",
        "",
        "**代码8（长评）**：必须是100字以上、有明确论点导向的分析。以下不算：纯情绪感叹、剧情复述、粉丝向角色讨论、列优缺点但无论点。",
        "",
        "**代码12（话题词条）**：使用#话题#标签是微博常规操作，不构成此代码。必须有明确的词条聚合引导意图（如呼吁他人到该话题下跟帖、在超话内组织讨论）才算。",
        "",
        "**代码14 vs 15 vs 22**：",
        "- 有具体批评论点且逐条驳斥，或质疑批评者动机/来源 → 14",
        "- 无具体批评对象，为本片理性澄清立场 → 15",
        "- 涉及不同影片之间的比较（票房/质量/排片） → 22",
        "- 三者可以同时出现",
        "",
        "**代码5 vs 7**：",
        "- 有创作成本（绘图/剪辑/同人文/音频） → 5；发了图/视频链接即打5",
        "- 纯文字谐音梗/改编段子/截图加文字调侃 → 7",
        "",
        "**代码11备注**：必须有推荐动作；带朋友去看、叫朋友一起也算；仅@某人讨论不算。",
        "",
        "**代码6备注**：是作为观众主动发现细节，不是作为宣发/制作方披露。",
        "",
        '**空数组规则**：只有以下情况才允许输出空数组[]：正文是纯转发无评论、纯图片说明、或与电影完全无关的内容。只要正文有实质性的涉片内容，就必须至少编一个代码。禁止把"拿不准"当作不编码的理由。',
        "",
        "## 标注示例（编码时请严格参照）",
        "以下示例来自真实标注数据，每类展示正确编码逻辑。",
        "",
        "### 边界一：代码5（同人创作）vs 代码7（搞笑梗）",
        "判断核心：有无原创创作成本？",
        '✅ [5] "昆仑快递，请您签收[嘻嘻]"（附自制插图） → 有绘图成本，打5。正文只有一句话也打5，关键看有没有发布创作物。',
        '✅ [5] "感谢@xxx帮我完成了这个视频 因为电脑被我下的球2素材源卡死多次..."（附视频链接） → 描述剪辑过程，打5。',
        '✅ [7] "封神观后感：幻视绍伊古不是我的错……！！" → 纯文字联想梗，无创作成本，打7。',
        '✅ [7] "原来是注水的\'注\'" → 一句话谐音梗，打7。',
        '✅ [7] "汉堡王（捂嘴）"（附截图） → 截图+调侃文字，无原创创作，打7。',
        "❌ 常见错误：看到图就打5——截图+调侃文字是7，自制绘图/视频才是5。",
        "",
        "### 边界二：代码11（安利）——高频附带行为，主轴之外也要打",
        "判断核心：正文是否包含向他人推荐观影的动作信号？",
        '✅ [11] "九张图不剧透式安利这电影，，，" → 明确写"安利"，打11。',
        '✅ [1, 11] "今天四刷《消失的她》，带来两个一无所知的朋友……" → 主轴是晒N刷（1），"带来两个朋友"是安利，同时打11。',
        '✅ [11] "他都脱成那样了……电影没看过可以去看看电影🎦，封神第一部，别买错票" → 主体讨论演员，结尾"可以去看看"是安利信号，打11。',
        '✅ [11] "今天又献祭了两个质子，快去影院审视他们……" → "快去影院"是安利信号，打11。',
        '✅ [11, 17] "天呐封神榜也太好看了吧！化身自来水，大家也要去看啊啊啊啊啊" → "大家也要去看"=11，"化身自来水"=17，同时打。',
        '❌ 常见错误：只看帖子主题，忽略结尾/顺带的推荐语——只要有"快去看/去影院/值得一看/安利/别错过"等词，无论主题是什么，都加11。',
        "",
        "### 边界三：代码17（自来水）——身份标签，附带行为",
        "判断核心：发布者是否声明自己是非粉丝的自发支持者？",
        '✅ [17] "首先我不是谁粉丝纯是因为这电影雷点最少……" → "不是谁粉丝"是自来水身份声明，打17。',
        '✅ [17, 20] "这才是观众说了算不是zb说了算的电影，封神值得" → 路人身份为电影发声，17+20同时打。',
        '✅ [11, 17, 22] "纯路人，仅从个人观影体验出发，满江红和流浪地球2，如果纠结看一部的话，选流浪地球[举手]。" → "纯路人"=17，"选流浪地球"=推荐=11，两片比较=22，三个同时打。',
        '❌ 常见错误：只打主要行为代码，漏掉"不是粉丝/纯路人/自来水/精神股东"这类身份声明——看到这些词立刻加17。',
        "",
        "### 边界四：代码19（共情）——对主创的情感表达",
        "判断核心：是否对导演/演员/幕后团队表达心疼、感动、感谢、敬佩？",
        '✅ [19, 28] "@乌尔善导演，你配享太庙！" → 对导演表达敬佩，打19。',
        '✅ [19] "谢谢一龙和其他主创的长沙路演 谢谢热腾腾的票……" → 对主创表达感谢，打19。',
        '✅ [17, 19] "封神也太好了吧！有第二部一定会去看的🥰" → 路人视角=17，对主创的支持和期待=19。',
        "❌ 区分19和11：19是情感冲着主创（演员/导演），11是行为冲着受众（推荐别人去看）。同一条可以同时有。",
        "",
        "### 边界五：代码14（反驳差评）vs 15（维护电影）vs 22（分析市场）",
        '✅ [14, 15] "那些黑乌尔善和封神的其心可诛" → 质疑黑评者动机=14；为电影辩护=15。',
        '✅ [14, 15, 22] "封神粉从孤注一掷点映第一天就开始踩……" → 驳斥踩竞品行为=14；为孤注辩护=15；涉及两部电影比较=22。',
        '✅ [15, 20] "封神这部电影……为的就是打压它" → 无具体批评对象，为电影处境辩护=15；表达被打压的愤慨=20。没有逐条驳斥所以不打14。',
        '✅ [22]（不是20也不是14） "部分媒体评论 新京 澎湃 财经 三联 永远不让你失望" → 列举媒体评价含比较意味，打22；无委屈情绪不打20，无驳斥不打14。',
        '❌ 常见错误：把任何"为电影说话"都打成14——14必须有具体批评对象并驳斥；只是表态支持是15。',
        "",
        "### 边界六：代码20（鸣不平）——情绪化声讨不公正",
        "判断核心：是否对电影受到的不公对待表达强烈情绪？",
        '✅ [2, 20] "我是真的在操心🙏🏻" → 为票房表达焦虑=20；有票房话题=2。',
        '✅ [20] "认认真真拍个电影 到底是招谁惹谁了" → 替电影鸣不平的情绪化表达，打20。',
        '✅ [20]（不是22） "《流浪地球2》和《无名》，值得更高票房" → 替电影抱不平=20；只是顺带提到多部，无比较分析，不打22。',
        '❌ 常见错误：把"提到其他电影"就打22——22要有明确的比较分析。只是顺带提到或只为目标片发声，打20不打22。',
        "",
        "### 边界七：代码8（长评）——100字以上有论点的分析",
        "判断核心：字数够且有明确论点导向？",
        '✅ [8] "都说消失的她里面朱一龙的演技比人生大事更精进了……不是演技更精进了，而是消失的她给了朱一龙更大的发挥空间……" → 有论点，有分析，打8。',
        '✅ [8] "其实姬发的人设真的有那么无敌吗？……可为什么给观众的感觉就是很悬浮呢？" → 有核心论点并展开论据，打8。',
        '✅ [1, 8] "IMAX看的很过瘾……周末准备二刷！费翔老师的纣王集贵气，英武，狠辣，阴鸷于一身……" → 有具体角色评价（100字以上）=8；有观影记录（准备二刷）=1，同时打。',
        '❌ [不打8] "超好看，故事很完整，反转很精彩，朱一龙演技牛批！" → 无论点导向，字数不足，不打8。',
        '❌ [不打8，打17] "这片子功德无量 值得中国影史第十一名 申奥导演剪辑节奏太流畅了 全程无尿点 群戏全员上桌" → 列举优点但无分析论证，不打8；路人自来水视角，打17。',
        "",
        "## 输出格式",
        "输入是一个 JSON 数组，每项含 id/user/movie/text 字段，其中 id 是序号（1开始）。",
        "输出必须是一个 JSON 对象，key 是 id（字符串），value 是该微博适用的编码编号数组。",
        "没有适用编码则为空数组[]。必须为每一条都输出对应的 key。",
        '例：输入3条(id=1,2,3) → 输出 {"1":[5,12],"2":[22],"3":[]}',
    ]
    return "\n".join(lines)


def load_posts(path: str) -> list[dict]:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, list):
        return data
    for key in ("data", "posts", "statuses", "items"):
        if key in data and isinstance(data[key], list):
            return data[key]
    raise ValueError(f"无法识别 JSON 结构: {type(data)}")


def post_id(post: dict) -> str:
    return str(post.get("id") or post.get("mid") or post.get("idstr", ""))


def call_api(client: OpenAI, system_prompt: str, batch: list[dict], batch_index: int) -> list[list[int]]:
    payload = [
        {
            "id": i + 1,
            "user": p.get("user", {}).get("screen_name", "") if isinstance(p.get("user"), dict) else str(p.get("user", "")),
            "movie": p.get("movie", ""),
            "text": (p.get("text") or p.get("content") or "")[:500],
        }
        for i, p in enumerate(batch)
    ]
    user_msg = json.dumps(payload, ensure_ascii=False)

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_msg},
                ],
                max_tokens=MAX_TOKENS,
                extra_body={"thinking": {"type": "enabled", "budget_tokens": THINKING_BUDGET}},
            )
            raw = resp.choices[0].message.content.strip()
            # 提取最外层 JSON 对象
            start, end = raw.find("{"), raw.rfind("}")
            if start == -1 or end == -1:
                raise ValueError(f"响应不含 JSON 对象: {raw[:200]}")
            result_dict = json.loads(raw[start:end + 1])
            # 按顺序组装，缺失的条目填空数组
            result = []
            missing = []
            for i in range(len(batch)):
                key = str(i + 1)
                if key in result_dict:
                    result.append(result_dict[key])
                else:
                    result.append([])
                    missing.append(key)
            if missing:
                print(f"  [批次{batch_index}] 警告：缺失 id={','.join(missing)}，已填空", flush=True)
            return result
        except Exception as e:
            wait = 2 ** attempt
            print(f"  [批次{batch_index}] 第{attempt}次失败: {e}，{wait}s 后重试…", flush=True)
            if attempt == MAX_RETRIES:
                print(f"  [批次{batch_index}] 放弃，填充空编码", flush=True)
                return [[] for _ in batch]
            time.sleep(wait)


def main():
    parser = argparse.ArgumentParser(description="批量编码微博")
    parser.add_argument("--input",    default="data/no_verified_classified_with_movie.json")
    parser.add_argument("--output",   default="data/coded_output.json")
    parser.add_argument("--codebook", default="data/codebook.xlsx")
    parser.add_argument("--batch",        type=int, default=20)
    parser.add_argument("--workers",      type=int, default=5)
    parser.add_argument("--show-prompt",  action="store_true", help="打印 prompt 后退出")
    parser.add_argument("--test",         action="store_true", help="跑前200条并与 top200-coding.xlsx 对比后退出")
    parser.add_argument("--top200",       default="data/top200-coding.xlsx", help="人工编码参考文件")
    args = parser.parse_args()

    checkpoint_path = Path(args.output).with_suffix(".checkpoint.json")

    config = load_config()
    client = OpenAI(api_key=config["api_key"], base_url=YUNWU_BASE_URL)

    print("读取 codebook…")
    codes = load_codebook(args.codebook)
    system_prompt = build_system_prompt(codes)
    code_name_map = {c["num"]: c["name"] for c in codes}

    if args.show_prompt:
        print("=" * 60)
        print("【SYSTEM PROMPT】")
        print(system_prompt)
        print("=" * 60)
        print("【USER 消息示例（第一批前3条）】")
        posts_preview = load_posts(args.input)[:3]
        sample = [
            {
                "id": i + 1,
                "user": p.get("user", {}).get("screen_name", "") if isinstance(p.get("user"), dict) else str(p.get("user", "")),
                "movie": p.get("movie", ""),
                "text": (p.get("text") or p.get("content") or "")[:500],
            }
            for i, p in enumerate(posts_preview)
        ]
        print(json.dumps(sample, ensure_ascii=False, indent=2))
        return

    print("读取微博数据…")
    posts = load_posts(args.input)
    print(f"共 {len(posts)} 条")

    if args.test:
        test_posts = posts[:200]
        print(f"【测试模式】处理前 {len(test_posts)} 条，批大小={args.batch}，并发={args.workers}…")

        # 并发跑批次
        test_results: dict[str, list[int]] = {}

        def process_batch(idx_batch):
            idx, b = idx_batch
            return b, call_api(client, system_prompt, b, idx)

        batches = [test_posts[i:i + args.batch] for i in range(0, len(test_posts), args.batch)]
        with ThreadPoolExecutor(max_workers=args.workers) as executor:
            futures = {executor.submit(process_batch, (i, b)): i for i, b in enumerate(batches)}
            done = 0
            for future in as_completed(futures):
                b, res = future.result()
                for p, nums in zip(b, res):
                    test_results[post_id(p)] = [int(n) for n in nums if isinstance(n, (int, float))]
                done += len(b)
                print(f"  进度：{done}/{len(test_posts)}", end="\r", flush=True)
        print()

        # 与 top200-coding.xlsx 对比
        ref = load_top200_coding(args.top200)
        hit = miss = extra = no_ref = 0
        print("\n" + "=" * 70)
        print(f"{'排名/ID':<20} {'人工编码':<28} {'模型编码':<28} 差异")
        print("=" * 70)
        for rank, p in enumerate(test_posts, 1):
            pid = post_id(p)
            model_set = set(test_results.get(pid, []))
            if pid not in ref:
                no_ref += 1
                continue
            ref_set = set(ref[pid])
            only_ref = ref_set - model_set   # 漏编
            only_model = model_set - ref_set  # 多编
            hit += len(ref_set & model_set)
            miss += len(only_ref)
            extra += len(only_model)

            if only_ref or only_model:
                ref_str = ", ".join(str(n) for n in sorted(ref_set)) or "无"
                model_str = ", ".join(str(n) for n in sorted(model_set)) or "无"
                diff_parts = []
                if only_ref:
                    diff_parts.append(f"漏:{','.join(str(n) for n in sorted(only_ref))}")
                if only_model:
                    diff_parts.append(f"多:{','.join(str(n) for n in sorted(only_model))}")
                print(f"[{rank:03d}]{pid:<16} {ref_str:<28} {model_str:<28} {' | '.join(diff_parts)}")

        total_ref = hit + miss
        print("=" * 70)
        print(f"\n【汇总】参考编码实例总数：{total_ref}，其中：")
        print(f"  命中（有）：{hit}  ({hit/total_ref*100:.1f}%)" if total_ref else "  命中：0")
        print(f"  漏编（没有）：{miss}  ({miss/total_ref*100:.1f}%)" if total_ref else "  漏编：0")
        print(f"  多编（错误）：{extra}")
        if no_ref:
            print(f"  无参考（top200中未找到）：{no_ref} 条")
        return

    # 断点续跑
    if checkpoint_path.exists():
        with open(checkpoint_path, encoding="utf-8") as f:
            code_results: dict[str, list[int]] = json.load(f)
        print(f"断点续跑：已完成 {len(code_results)} 条")
    else:
        code_results = {}

    todo_posts = [p for p in posts if post_id(p) not in code_results]
    print(f"待处理：{len(todo_posts)} 条，批大小={args.batch}，并发={args.workers}")

    if todo_posts:
        batches = [todo_posts[i:i + args.batch] for i in range(0, len(todo_posts), args.batch)]
        completed = 0
        total = len(todo_posts)

        def process_batch(idx_batch):
            idx, batch = idx_batch
            return batch, call_api(client, system_prompt, batch, idx)

        with ThreadPoolExecutor(max_workers=args.workers) as executor:
            futures = {executor.submit(process_batch, (i, b)): i for i, b in enumerate(batches)}
            for future in as_completed(futures):
                batch, result = future.result()
                for p, nums in zip(batch, result):
                    code_results[post_id(p)] = [int(n) for n in nums if isinstance(n, (int, float))]
                completed += len(batch)
                with open(checkpoint_path, "w", encoding="utf-8") as f:
                    json.dump(code_results, f, ensure_ascii=False)
                print(f"进度：{completed}/{total} ({completed/total*100:.1f}%)", end="\r", flush=True)
        print(f"\n编码完成，共 {len(code_results)} 条")

    # 输出 JSON：在原始数据上附加 codes 字段
    output = []
    for p in posts:
        pid = post_id(p)
        nums = code_results.get(pid, [])
        output.append({
            **p,
            "codes": nums,
            "code_names": [f"{n} {code_name_map.get(n, '')}" for n in nums],
        })

    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"已写出 → {args.output}")


if __name__ == "__main__":
    main()
