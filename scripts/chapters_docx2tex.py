# -*- coding: utf-8 -*-
"""chapters_docx2tex.py —— 四章式大报告：多人 Word 稿件 → LaTeX（v2 引擎）

结构（v2 起）：
    第二章  重大问题                     ← 章（第X章，雅黑三号加粗居中）
      1.  结合案例分析泌乳素分泌的调节因素   ← 每份稿件 = 章内一个一级标题
          1.1  病例特点                     ← 稿件里的 1 级标题
              1.1.1  实验室指标             ← 稿件里的 2 级标题
      2.  下一个人的题目

实现委托给 app/pbltex/core.py（与桌面版软件共用同一套引擎：
标题降级、封面、表格自适应宽度、参考文献跨文件去重、上标引用直出）。

用法：
    python scripts/chapters_docx2tex.py 计划.json -o 大报告.tex --compile
    python scripts/chapters_docx2tex.py 提交目录 -o 大报告.tex --auto --compile

计划 JSON（与 v1 兼容，新增 cover）：
{
  "title": "内分泌系统 / PBL 大报告",
  "author": "第11组",
  "date": "2026 年 10 月",
  "cover": {
    "line1": "2023级11组PBL学习",
    "line2": "学习报告",
    "case": "Case 1：难缠的脖子痛",
    "tutor": "马凯歌",
    "chair": "朱俊铭",
    "secretary": "郭育恒",
    "members": ["叶瑞林", "张欣仪", "韩曦", "任欣萌", "王思源", "于子轩", "牛海华", "韩畅"],
    "logo": "pics/xjtu_logo.png"
  },
  "chapters": [
    {"title": "案例回顾", "files": ["提交/01_病例.docx"]},
    {"title": "重大问题", "files": ["提交/02_机制.docx"]},
    {"title": "次要问题", "files": ["提交/03_鉴别.docx"]},
    {"title": "机制图", "figures": [{"path": "pics/mech1.png", "caption": "三条链机制全景"}]}
  ]
}
"""
from __future__ import annotations

import argparse
import glob
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
for cand in (os.path.join(REPO, "app"), REPO):
    if os.path.isdir(os.path.join(cand, "pbltex")) and cand not in sys.path:
        sys.path.insert(0, cand)

try:
    from pbltex.core import CHAPTER_NAMES, convert_chapters, guess_chapter
except ImportError:  # pragma: no cover
    sys.exit("找不到 app/pbltex。请确认本脚本与 app/ 目录在同一仓库里（仓库根目录应含 app/pbltex/core.py）。")

IMG_EXT = (".png", ".jpg", ".jpeg", ".emf", ".wmf", ".tif", ".tiff", ".bmp", ".gif")


def _log(msg):
    print(msg)
    sys.stdout.flush()


def _abs(base, p):
    return p if os.path.isabs(p) else os.path.normpath(os.path.join(base, p))


def plan_from_json(plan_path):
    with open(plan_path, encoding="utf-8-sig") as fh:
        plan = json.load(fh)
    base = os.path.dirname(os.path.abspath(plan_path))
    items, titles, figures = [], {}, {}
    for idx, chap in enumerate(plan.get("chapters") or []):
        title = chap.get("title") or (CHAPTER_NAMES[idx] if idx < len(CHAPTER_NAMES) else "第%d章" % (idx + 1))
        if title != (CHAPTER_NAMES[idx] if idx < len(CHAPTER_NAMES) else title):
            titles[idx] = title
        for f in chap.get("files") or []:
            items.append((_abs(base, f), idx))
        figs = []
        for fg in chap.get("figures") or []:
            p = fg.get("path") if isinstance(fg, dict) else fg
            cap = fg.get("caption", "") if isinstance(fg, dict) else ""
            if p:
                figs.append((_abs(base, p), cap))
        if figs:
            figures[idx] = figs
    meta = {
        "title": plan.get("title", "PBL 大报告"),
        "author": plan.get("author", ""),
        "date": plan.get("date", ""),
        "cover": plan.get("cover"),
        "section_titles": plan.get("section_titles"),
    }
    return items, titles, figures, meta


def plan_from_dir(folder):
    """按关键词自动分章：稿件读题目首行，图片归“机制图”。"""
    docx = [f for f in sorted(glob.glob(os.path.join(folder, "*.docx")))
            if not os.path.basename(f).startswith("~$")]
    imgs = [f for f in sorted(glob.glob(os.path.join(folder, "*")))
            if f.lower().endswith(IMG_EXT)]
    items = []
    for f in docx:
        chap = guess_chapter(f)
        items.append((os.path.abspath(f), chap))
        _log("  %-40s → %s" % (os.path.basename(f), CHAPTER_NAMES[chap]))
    figures = {}
    if imgs:
        figures[3] = [(os.path.abspath(p), os.path.splitext(os.path.basename(p))[0]) for p in imgs]
        _log("  %d 张图片 → 机制图" % len(imgs))
    meta = {"title": "PBL 大报告", "author": "", "date": "", "cover": None}
    return items, {}, figures, meta


def main():
    ap = argparse.ArgumentParser(description="四章式大报告：多人稿件 → LaTeX（v2 引擎）")
    ap.add_argument("source", help="计划 JSON 文件，或稿件目录（配 --auto）")
    ap.add_argument("-o", "--output", default="大报告.tex")
    ap.add_argument("--auto", action="store_true", help="按题目关键词把目录里的稿件自动分章")
    ap.add_argument("--compile", action="store_true", help="生成后用 xelatex 编译")
    ap.add_argument("--title", default="", help="大报告标题（覆盖计划里的 title）")
    ap.add_argument("--author", default="", help="组别 / 作者")
    ap.add_argument("--date", default="", help="日期")
    ap.add_argument("--cover-line1", default="", help="封面首行，如 2023级11组PBL学习")
    ap.add_argument("--cover-line2", default="", help="封面次行，默认“学习报告”")
    ap.add_argument("--case", default="", help="Case 标题，如 Case 1：难缠的脖子痛")
    ap.add_argument("--tutor", default="")
    ap.add_argument("--chair", default="")
    ap.add_argument("--secretary", default="")
    ap.add_argument("--members", default="", help="组员，空格/逗号分隔")
    ap.add_argument("--logo", default="", help="封面校徽图片路径")
    args = ap.parse_args()

    outdir = os.path.dirname(os.path.abspath(args.output)) or "."
    os.makedirs(outdir, exist_ok=True)
    tex_name = os.path.basename(args.output)

    if args.auto or os.path.isdir(args.source):
        items, titles, figures, meta = plan_from_dir(args.source)
    else:
        items, titles, figures, meta = plan_from_json(args.source)

    cover = dict(meta.get("cover") or {})
    for key, val in (("line1", args.cover_line1), ("line2", args.cover_line2), ("case", args.case),
                     ("tutor", args.tutor), ("chair", args.chair), ("secretary", args.secretary),
                     ("members", args.members), ("logo", args.logo)):
        if val:
            cover[key] = val
    if args.date:
        cover.setdefault("date", args.date)

    if args.title:
        meta["title"] = args.title
    if args.author:
        meta["author"] = args.author
    if args.date:
        meta["date"] = args.date

    if not items and not figures:
        sys.exit("没有可用的稿件：%s" % args.source)

    info = convert_chapters(
        items, outdir,
        title=meta["title"], author=meta["author"], date=meta["date"],
        tex_name=tex_name, compile_pdf=args.compile,
        chapter_titles=titles, chapter_figures=figures,
        cover=cover or None, log=_log,
        section_titles=meta.get("section_titles"))

    _log("")
    _log("生成完成：")
    _log("  tex      %s" % info["tex"])
    _log("  分章清单  %s" % info["manifest"])
    _log("  文献      %d 条（跨文件去重后）" % info["refs"])
    if info.get("pdf"):
        _log("  pdf      %s" % info["pdf"])
    elif args.compile:
        _log("  pdf      未生成（看上面的 xelatex 报错）")
    if info.get("skipped"):
        _log("  跳过      %s" % "、".join(info["skipped"]))


if __name__ == "__main__":
    main()

