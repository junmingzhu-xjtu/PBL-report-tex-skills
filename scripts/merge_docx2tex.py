# -*- coding: utf-8 -*-
"""
merge_docx2tex.py —— 把多人提交的 Word 文稿合并成一份 LaTeX（大报告）

用法：
    python merge_docx2tex.py 提交目录 -o 大报告.tex
    python merge_docx2tex.py 提交目录 -o 大报告.tex --compile
    python merge_docx2tex.py a.docx b.docx c.docx -o 大报告.tex

它做的事：
    1) 逐个文件识别：各级标题（1/1.1/1.1.1）、正文、图片与题注、表格、参考文献
    2) 标题级别归一化：每人文中最浅的标题 → 一级，其余据此顺延，最深到三级
       （更深的层级自动降级为 (1)(2)(3) 有序列表）
    3) 图片按“作者_序号”重命名，统一放 figures/，题注“图x.”由 LaTeX 自动编号
    4) 参考文献：跨文件去重合并，正文 [n] 引用自动重编号到全局序号
    5) 输出 .tex、figures/、合并清单.txt；加 --compile 直接调 xelatex 出 PDF
"""
import argparse
import glob
import hashlib
import os
import re
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
try:
    import docx2tex as d2t
    from docx import Document
except ImportError as e:
    sys.exit("依赖缺失：pip install python-docx  （%s）" % e)

CITE_RE = re.compile(r"\[(\d+(?:\s*[-,，–]\s*\d+)*)\]")


# ---------------- 工具 ----------------
def slug(name, limit=12):
    base = os.path.splitext(os.path.basename(name))[0]
    base = re.sub(r"[^\w\u4e00-\u9fff]+", "", base)
    return (base[:limit] or "doc") + "_"


def norm_ref(text):
    t = re.sub(r"^\[\d+\]\s*", "", text.strip())
    t = re.sub(r"\s+", "", t)
    t = t.replace("，", ",").replace("．", ".").replace("：", ":")
    return t


def expand_cite(token):
    """'1-3' / '1,2' -> [1,2,3]"""
    out = []
    for part in re.split(r"[,，]", token):
        part = part.strip()
        m = re.match(r"^(\d+)\s*[-–]\s*(\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            out.extend(range(a, b + 1))
        elif part.isdigit():
            out.append(int(part))
    return out


def fmt_cite(nums):
    nums = sorted(set(nums))
    if not nums:
        return ""
    groups, start, prev = [], nums[0], nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
            continue
        groups.append((start, prev))
        start = prev = n
    groups.append((start, prev))
    parts = []
    for a, b in groups:
        if a == b:
            parts.append(str(a))
        elif b == a + 1:
            parts.append("%d,%d" % (a, b))
        else:
            parts.append("%d-%d" % (a, b))
    return "[" + ",".join(parts) + "]"


# ---------------- 单文件处理 ----------------
def prepare(path, figdir, state, author=None, strip_lead=False):
    """返回该文件的 blocks（已归一化标题层级、引用已重编号）"""
    doc = Document(path)
    prefix = slug(path)
    blocks = d2t.collect(doc, figdir, prefix=prefix, start=state["fig"] * 0 + 1)
    state["fig"] += 1

    # 拆出参考文献条目
    refs_here = []
    kept = []
    in_refs = False
    for b in blocks:
        if b[0] == "h" and "参考文献" in b[2]:
            in_refs = True
            continue
        if in_refs and b[0] == "p" and re.match(r"^\[\d+\]", b[1]):
            refs_here.append(b[1])
            continue
        if b[0] == "p" and re.match(r"^\[\d+\]\s*\S", b[1]) and "et al" not in b[1][:0]:
            # 没有“参考文献”标题时，也把 [n] 开头的段落当文献
            if in_refs:
                refs_here.append(b[1])
                continue
        kept.append(b)
    blocks = kept

    # 去掉第一个标题之前的“个人稿件标题/作者/日期”等段落（可选）
    if strip_lead:
        for idx0, b0 in enumerate(blocks):
            if b0[0] == "h":
                blocks = blocks[idx0:]
                break

    # 标题归一化
    levels = [b[1] for b in blocks if b[0] == "h"]
    base = min(levels) if levels else 1
    norm = []
    for b in blocks:
        if b[0] == "h":
            lvl = b[1] - base + 1
            if lvl > 3:
                norm.append(("li", b[2]))
            else:
                norm.append(("h", lvl, b[2]))
        else:
            norm.append(b)

    # 文件没有一级标题时，才包一层“提交人”节；层级超三级的降为加粗小标题
    h1 = [b for b in norm if b[0] == "h" and b[1] == 1]
    top = min([b[1] for b in norm if b[0] == "h"], default=2)
    if author and not h1:
        shift = 2 - top
        shifted = []
        for b in norm:
            if b[0] == "h":
                lvl = b[1] + shift
                shifted.append(("h", lvl, b[2]) if lvl <= 3 else ("b", b[2]))
            else:
                shifted.append(b)
        norm = [("h", 1, author)] + shifted

    # 引用重编号
    mapping = {}
    for r in refs_here:
        n = int(re.match(r"^\[(\d+)\]", r).group(1))
        key = norm_ref(r)
        if key not in state["refs"]:
            state["refs"][key] = len(state["refs"]) + 1
            state["ref_list"].append(r)
            state["ref_text"][key] = re.sub(r"^\[\d+\]\s*", "", r.strip())[:70]
        mapping[n] = state["refs"][key]

    if mapping:
        def repl(m):
            nums = expand_cite(m.group(1))
            new = [mapping.get(x, x) for x in nums]
            return fmt_cite(new)

        for i, b in enumerate(norm):
            if b[0] == "p":
                norm[i] = ("p", CITE_RE.sub(repl, b[1]))

    return norm, refs_here


# ---------------- 渲染 ----------------
def fig_env(paths, caption):
    lines = [r"\begin{figure}[htbp]", r"  \centering"]
    for p in paths:
        lines.append(r"  \includegraphics[width=0.8\linewidth]{%s}" % p)
    if caption:
        lines.append(r"  \caption{%s}" % d2t.esc(caption))
    lines.append(r"\end{figure}")
    return "\n".join(lines)


def render_blocks(blocks):
    out = []
    i = 0
    while i < len(blocks):
        b = blocks[i]
        if b[0] == "h":
            cmd = {1: "section", 2: "subsection", 3: "subsubsection"}[b[1]]
            out.append("\\%s{%s}" % (cmd, d2t.esc(b[2])))
            i += 1
            continue
        if b[0] == "img":
            cap = ""
            if i + 1 < len(blocks) and blocks[i + 1][0] == "cap" and blocks[i + 1][1] == "图":
                cap = blocks[i + 1][2]
                i += 1
            out.append(fig_env(b[1], cap))
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
            out.append(r"\begin{enumerate}[label=(\arabic*),leftmargin=2em]")
            for it in items:
                out.append(r"  \item %s" % d2t.esc(it))
            out.append(r"\end{enumerate}")
            continue
        if b[0] == "b":
            out.append("\\noindent\\textbf{%s}\\par" % d2t.esc(b[1]))
            i += 1
            continue
        if b[0] == "cap":
            out.append(r"\begin{center}\small %s\end{center}" % d2t.esc(
                b[3] if len(b) > 3 else "%s%s" % (b[1], b[2])))
            i += 1
            continue
        out.append(d2t.esc(b[1]))
        i += 1
    return out


def main():
    ap = argparse.ArgumentParser(description="多人 Word 稿 -> 一份 LaTeX 大报告")
    ap.add_argument("inputs", nargs="+", help="目录或 .docx 文件（可多个）")
    ap.add_argument("-o", "--output", default="大报告.tex")
    ap.add_argument("--margins", default=None, help="上,下,左,右 cm")
    ap.add_argument("--compile", action="store_true", help="调用 xelatex 编译 PDF")
    ap.add_argument("--upcite", action="store_true", help="正文 [n] 转 \\upcite{}")
    ap.add_argument("--toc", dest="toc", action="store_true", default=True, help="生成目录（默认开）")
    ap.add_argument("--no-toc", dest="toc", action="store_false", help="不生成目录")
    ap.add_argument("--strip-lead", action="store_true",
                    help="丢弃每个文件第一个标题之前的标题栏/作者/日期段落")
    a = ap.parse_args()

    # 收集文件
    files = []
    for it in a.inputs:
        if os.path.isdir(it):
            files += sorted(glob.glob(os.path.join(it, "*.docx")))
        else:
            files.append(it)
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files:
        sys.exit("没有找到 .docx 文件")

    outdir = os.path.dirname(os.path.abspath(a.output)) or "."
    figdir = os.path.join(outdir, "figures")
    os.makedirs(figdir, exist_ok=True)

    state = {"fig": 0, "refs": {}, "ref_list": [], "ref_text": {}}
    body = []
    manifest = []

    for idx, f in enumerate(files, 1):
        name = os.path.splitext(os.path.basename(f))[0]
        author = re.sub(r"^[\d\W_]+", "", name).replace("_", " · ").strip() or name
        try:
            blocks, refs = prepare(f, figdir, state, author=author, strip_lead=a.strip_lead)
        except Exception as e:
            print("[SKIP] %s: %s" % (name, e))
            manifest.append("%-40s 失败：%s" % (name, e))
            continue
        n_h = sum(1 for b in blocks if b[0] == "h")
        n_img = sum(len(b[1]) for b in blocks if b[0] == "img")
        n_tbl = sum(1 for b in blocks if b[0] == "tbl")
        body.append("%% ===== %d. %s（%s） =====" % (idx, name, os.path.basename(f)))
        body.extend(render_blocks(blocks))
        manifest.append("%-40s 标题 %2d，图 %2d，表 %2d，参考文献 %2d" % (name, n_h, n_img, n_tbl, len(refs)))
        print("[OK] %s: 标题 %d, 图 %d, 表 %d, 参考 %d" % (name, n_h, n_img, n_tbl, len(refs)))

    # 参考文献
    if state["ref_list"]:
        body.append(r"\section*{参考文献}")
        body.append(r"\begin{thebibliography}{99}")
        for i, r in enumerate(state["ref_list"], 1):
            txt = re.sub(r"^\[\d+\]\s*", "", r.strip())
            if a.upcite:
                pass
            body.append(r"  \bibitem{ref%d} %s" % (i, d2t.esc(txt)))
        body.append(r"\end{thebibliography}")

    margins = [float(x) for x in a.margins.split(",")] if a.margins else [2.54, 2.54, 1.91, 1.91]
    head = d2t.PREAMBLE.format(t="%.2f" % margins[0], b="%.2f" % margins[1],
                               l="%.2f" % margins[2], r="%.2f" % margins[3])
    title_block = ("\n\n\\title{×××课程 / PBL 大报告}\n"
                   "\\author{第 × 组}\n\\date{2026 年 10 月}\n\\maketitle\n\n")
    if a.toc:
        title_block = title_block + "\\tableofcontents\n\\newpage\n\n"
    head = head.replace("\\begin{document}\n", "\\begin{document}" + title_block, 1)
    with open(a.output, "w", encoding="utf-8") as fh:
        fh.write(head + "\n\n".join(body) + "\n\n\\end{document}\n")

    rep = os.path.join(outdir, "合并清单.txt")
    with open(rep, "w", encoding="utf-8") as fh:
        fh.write("合并文件数：%d\n\n" % len(files))
        fh.write("\n".join(manifest))
        fh.write("\n\n参考文献（去重后 %d 条）\n" % len(state["ref_list"]))
        for k, v in state["refs"].items():
            fh.write("  [%d] %s\n" % (v, state["ref_text"].get(k, k[:70])))

    print("\n生成: %s\n      %s\n      %s\\" % (a.output, rep, figdir))

    if a.compile:
        for _ in range(2):
            r = subprocess.run(["xelatex", "-interaction=nonstopmode", os.path.basename(a.output)],
                               cwd=outdir, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        pdf = os.path.splitext(a.output)[0] + ".pdf"
        print("PDF：%s（%s）" % (pdf, "成功" if os.path.exists(pdf) else "失败"))


if __name__ == "__main__":
    main()
