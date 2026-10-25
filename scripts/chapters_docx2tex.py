# -*- coding: utf-8 -*-
"""
chapters_docx2tex.py —— 把多人稿件按“四章式”大报告组织成 LaTeX

四章固定为：
    第一章 案例回顾
    第二章 重大问题
    第三章 次要问题
    第四章 机制图

章标题：居中、三号(16pt)、微软雅黑、加粗（LaTeX 里用 xeCJK 的微软雅黑族）

用法：
    python chapters_docx2tex.py 计划.json -o 大报告.tex --compile
    python chapters_docx2tex.py 提交目录 -o 大报告.tex --auto     # 关键词自动分章

计划 JSON 格式：
{
  "title": "×××课程 / PBL 大报告",
  "author": "第 × 组",
  "date": "2026 年 10 月",
  "chapters": [
    {"title": "案例回顾", "files": ["提交/01_案例.docx"]},
    {"title": "重大问题", "files": ["提交/02_机制.docx", "提交/03_因果.docx"]},
    {"title": "次要问题", "files": ["提交/04_鉴别.docx"]},
    {"title": "机制图", "figures": [
        {"path": "pics/mech1.png", "caption": "下丘脑—垂体—甲状腺轴（示意）"}
    ]}
  ]
}
"""
import argparse
import glob
import json
import os
import re
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import docx2tex as d2t
from docx import Document

CHAPTERS = ["案例回顾", "重大问题", "次要问题", "机制图"]

# 自动分章用的关键词（可被计划 JSON 覆盖）
MAJOR_KEYS = ["机制", "为什么", "为何", "因果", "关系", "通路", "调控", "怎么", "如何", "原理", "解释"]
MINOR_KEYS = ["鉴别", "概念", "定义", "分类", "补充", "细节", "名词", "区别"]

CHAPTER_PREAMBLE = r"""%% !TEX program = xelatex
%% 由 chapters_docx2tex.py 生成（四章式大报告）
\documentclass[UTF8,zihao=-4,a4paper,fontset=windows,openany]{ctexrep}

\usepackage[top=2.54cm,bottom=2.54cm,left=1.91cm,right=1.91cm]{geometry}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{enumitem}
\usepackage{caption}
\usepackage[hidelinks]{hyperref}
\usepackage{natbib}
\setcitestyle{super,sort&compress,numbers}

\setmainfont{Times New Roman}
%% 微软雅黑：用于章标题
\setCJKfamilyfont{yh}{Microsoft YaHei}
\newcommand{\yahei}{\CJKfamily{yh}}
\pagestyle{plain}

%% ===== 章：居中、三号、微软雅黑、加粗 =====
\ctexset{
  chapter = {
    name = {第,章},
    number = \chinese{chapter},
    format = \centering\yahei\zihao{3}\bfseries,
    aftername = \quad,
    beforeskip = 1.2ex plus .3ex,
    afterskip  = 1.0ex plus .2ex,
  },
  section = {
    format = \heiti\zihao{-4}\bfseries,
    aftername = \hspace{0.55em},
    beforeskip = .9ex plus .2ex, afterskip = .5ex plus .1ex,
  },
  subsection = {
    format = \heiti\zihao{-4}\bfseries,
    aftername = \hspace{0.55em},
    beforeskip = .7ex plus .2ex, afterskip = .4ex plus .1ex,
  },
}
\setcounter{secnumdepth}{3}     %% 章/节/小节，再深用有序列表

\captionsetup{labelsep=period, font=small}
\renewcommand{\figurename}{图}
\renewcommand{\tablename}{表}
\renewcommand{\refname}{}
\providecommand{\bibname}{参考文献}

\setlength{\parindent}{2em}
\linespread{1.45}

\begin{document}
"""


def esc(t):
    return d2t.esc(t)


def chapter_open(title):
    return "\\chapter{%s}\n" % esc(title)


def render_file_blocks(blocks, cap_level=3):
    """文件内标题：最浅一级 -> \\section，下一级 -> \\subsection，再深 -> 加粗小标题/列表"""
    out = []
    i = 0
    while i < len(blocks):
        b = blocks[i]
        if b[0] == "h":
            lvl = b[1]
            if lvl == 2:
                out.append("\\section{%s}" % esc(b[2]))
            elif lvl == 3:
                out.append("\\subsection{%s}" % esc(b[2]))
            else:
                out.append("\\noindent\\textbf{%s}\\par" % esc(b[2]))
            i += 1
            continue
        if b[0] == "img":
            cap = ""
            if i + 1 < len(blocks) and blocks[i + 1][0] == "cap" and blocks[i + 1][1] == "图":
                cap = blocks[i + 1][2]
                i += 1
            out.append("\\begin{figure}[htbp]\n  \\centering")
            for p in b[1]:
                out.append("  \\includegraphics[width=0.8\\linewidth]{%s}" % p)
            if cap:
                out.append("  \\caption{%s}" % esc(cap))
            out.append("\\end{figure}")
            i += 1
            continue
        if b[0] == "tbl":
            cap = ""
            if i + 1 < len(blocks) and blocks[i + 1][0] == "cap" and blocks[i + 1][1] == "表":
                cap = blocks[i + 1][2]
                i += 1
            out.append(d2t.table_latex(b[1], cap))
            i += 1
            continue
        if b[0] == "li":
            items = []
            while i < len(blocks) and blocks[i][0] == "li":
                items.append(blocks[i][1])
                i += 1
            out.append("\\begin{enumerate}[label=(\\arabic*),leftmargin=2em]")
            for it in items:
                out.append("  \\item %s" % esc(it))
            out.append("\\end{enumerate}")
            continue
        if b[0] == "b":
            out.append("\\noindent\\textbf{%s}\\par" % esc(b[1]))
            i += 1
            continue
        if b[0] == "cap":
            out.append("\\begin{center}\\small %s\\end{center}" % esc(
                b[3] if len(b) > 3 else "%s%s" % (b[1], b[2])))
            i += 1
            continue
        out.append(esc(b[1]))
        i += 1
    return out


def normalize_for_chapter(blocks):
    """文件内的标题整体下移一层：最浅一级 -> 2（\\section），深一级 -> 3"""
    levels = [b[1] for b in blocks if b[0] == "h"]
    if not levels:
        return blocks
    top = min(levels)
    shift = 2 - top
    out = []
    for b in blocks:
        if b[0] == "h":
            lvl = max(2, b[1] + shift)
            out.append(("h", lvl, b[2]) if lvl <= 3 else ("b", b[2]))
        else:
            out.append(b)
    return out


def build_plan_from_dir(folder, auto=True):
    files = sorted(glob.glob(os.path.join(folder, "*.docx")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    plan = {"chapters": [{"title": t, "files": []} for t in CHAPTERS]}
    for f in files:
        name = os.path.splitext(os.path.basename(f))[0]
        # 用文件名里的问题名做关键词判断
        score_major = sum(1 for k in MAJOR_KEYS if k in name)
        score_minor = sum(1 for k in MINOR_KEYS if k in name)
        if "案例" in name or "回顾" in name or "病史" in name:
            idx = 0
        elif score_minor > score_major:
            idx = 2
        elif score_major > 0:
            idx = 1
        else:
            idx = 1
        plan["chapters"][idx]["files"].append(f)
    return plan


def main():
    ap = argparse.ArgumentParser(description="四章式大报告：多人稿件 -> LaTeX")
    ap.add_argument("source", help="计划 JSON 文件，或稿件目录（配 --auto）")
    ap.add_argument("-o", "--output", default="大报告.tex")
    ap.add_argument("--auto", action="store_true", help="按关键词把目录里的稿件自动分章")
    ap.add_argument("--compile", action="store_true")
    args = ap.parse_args()

    if args.auto or os.path.isdir(args.source):
        plan = build_plan_from_dir(args.source)
    else:
        with open(args.source, encoding="utf-8-sig") as fh:
            plan = json.load(fh)

    outdir = os.path.dirname(os.path.abspath(args.output)) or "."
    figdir = os.path.join(outdir, "figures")
    os.makedirs(figdir, exist_ok=True)

    body = []
    manifest = []
    state = {"fig": 0, "refs": {}, "ref_list": [], "ref_text": {}}
    CITE_RE = re.compile(r"\[(\d+(?:\s*[-,，–]\s*\d+)*)\]")

    def expand(tok):
        out = []
        for part in re.split(r"[,，]", tok):
            m = re.match(r"^(\d+)\s*[-–]\s*(\d+)$", part.strip())
            if m:
                out.extend(range(int(m.group(1)), int(m.group(2)) + 1))
            elif part.strip().isdigit():
                out.append(int(part.strip()))
        return out

    def fmt(nums):
        nums = sorted(set(nums))
        if not nums:
            return ""
        seg, a, b = [], nums[0], nums[0]
        for n in nums[1:]:
            if n == b + 1:
                b = n
                continue
            seg.append((a, b)); a = b = n
        seg.append((a, b))
        parts = []
        for x, y in seg:
            parts.append(str(x) if x == y else ("%d,%d" % (x, y) if y == x + 1 else "%d-%d" % (x, y)))
        return "[" + ",".join(parts) + "]"

    for ci, chap in enumerate(plan.get("chapters", []), 1):
        title = chap.get("title") or (CHAPTERS[ci - 1] if ci <= len(CHAPTERS) else "第%d章" % ci)
        files = chap.get("files") or []
        figs = chap.get("figures") or []
        if not files and not figs:
            continue
        body.append(chapter_open(title))
        manifest.append("第%d章 %s" % (ci, title))

        for f in files:
            if not os.path.isabs(f):
                f = os.path.join(os.path.dirname(os.path.abspath(args.source))
                                 if args.source.endswith(".json") else "", f)
            if not os.path.exists(f):
                print("[SKIP] 找不到文件：%s" % f)
                continue
            prefix = re.sub(r"[^\w\u4e00-\u9fff]+", "", os.path.splitext(os.path.basename(f))[0])[:12] + "_"
            blocks = d2t.collect(Document(f), figdir, prefix=prefix, start=1)

            refs_here = []
            kept, in_refs = [], False
            for b in blocks:
                if b[0] == "h" and "参考文献" in b[2]:
                    in_refs = True
                    continue
                if in_refs and b[0] == "p" and re.match(r"^\[\d+\]", b[1]):
                    refs_here.append(b[1])
                    continue
                kept.append(b)
            blocks = normalize_for_chapter(kept)

            mapping = {}
            for r in refs_here:
                n = int(re.match(r"^\[(\d+)\]", r).group(1))
                key = re.sub(r"\s+", "", re.sub(r"^\[\d+\]\s*", "", r))
                if key not in state["refs"]:
                    state["refs"][key] = len(state["refs"]) + 1
                    state["ref_list"].append(r)
                    state["ref_text"][key] = re.sub(r"^\[\d+\]\s*", "", r)[:70]
                mapping[n] = state["refs"][key]
            if mapping:
                kept2 = []
                for b in blocks:
                    if b[0] == "p":
                        kept2.append(("p", CITE_RE.sub(
                            lambda m: fmt([mapping.get(x, x) for x in expand(m.group(1))]), b[1])))
                    else:
                        kept2.append(b)
                blocks = kept2

            body.extend(render_file_blocks(blocks))
            manifest.append("    ├─ %s" % os.path.basename(f))

        for fg in figs:
            path = fg.get("path", "")
            if not os.path.isabs(path):
                base = os.path.dirname(os.path.abspath(args.source))
                path = os.path.join(base, path)
            if not os.path.exists(path):
                print("[SKIP] 找不到图片：%s" % path)
                continue
            name = "chap%d_fig%02d%s" % (ci, len(os.listdir(figdir)) + 1, os.path.splitext(path)[1] or ".png")
            with open(path, "rb") as src, open(os.path.join(figdir, name), "wb") as dst:
                dst.write(src.read())
            body.append("\\begin{figure}[htbp]\n  \\centering\n"
                        "  \\includegraphics[width=0.85\\linewidth]{figures/%s}\n"
                        "  \\caption{%s}\n\\end{figure}" % (name, esc(fg.get("caption", ""))))

    if state["ref_list"]:
        body.append("\\chapter*{参考文献}")
        body.append("\\begin{thebibliography}{99}")
        for i, r in enumerate(state["ref_list"], 1):
            body.append("  \\bibitem{ref%d} %s" % (i, esc(re.sub(r"^\[\d+\]\s*", "", r))))
        body.append("\\end{thebibliography}")

    title = plan.get("title", "×××课程 / PBL 大报告")
    author = plan.get("author", "第 × 组")
    date = plan.get("date", "2026 年 10 月")
    head = CHAPTER_PREAMBLE + "\\title{%s}\n\\author{%s}\n\\date{%s}\n\\maketitle\n\n\\tableofcontents\n\\newpage\n\n" % (
        esc(title), esc(author), esc(date))

    with open(args.output, "w", encoding="utf-8") as fh:
        fh.write(head + "\n\n".join(body) + "\n\n\\end{document}\n")

    rep = os.path.join(outdir, "分章清单.txt")
    with open(rep, "w", encoding="utf-8") as fh:
        fh.write("四章式大报告\n\n" + "\n".join(manifest) + "\n")
        fh.write("\n参考文献（去重后 %d 条）\n" % len(state["ref_list"]))
        for k, v in state["refs"].items():
            fh.write("  [%d] %s\n" % (v, state["ref_text"].get(k, k[:70])))

    print("生成: %s\n      %s\n      %s" % (args.output, rep, figdir))

    if args.compile:
        for _ in range(2):
            subprocess.run(["xelatex", "-interaction=nonstopmode", os.path.basename(args.output)],
                           cwd=outdir, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pdf = os.path.splitext(args.output)[0] + ".pdf"
        print("PDF:", pdf, "(OK)" if os.path.exists(pdf) else "(失败)")


if __name__ == "__main__":
    main()
