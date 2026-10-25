# -*- coding: utf-8 -*-
"""
docx2tex.py —— 把 Word 版大报告转成 LaTeX（ctexart）

用法：
    python docx2tex.py 大报告.docx
    python docx2tex.py 大报告.docx -o 大报告.tex
    python docx2tex.py 大报告.docx --margins 2.54,2.54,1.91,1.91 --upcite

规则：
    1) 1. / 1.1. / 1.1.1. 三级标题 -> section / subsection / subsubsection
    2) 更深层级 -> 有序列表 (1)(2)(3)，不再生成标题
    3) 图片 + 紧随的“图x. ……”题注 -> figure 环境
    4) 表格 + 紧随的“表x. ……”题注 -> table 环境（booktabs 三线表）
    5) 页边距默认读 docx，可用 --margins 覆盖（上,下,左,右，单位 cm）
    6) --upcite 把正文里的 [1]、[2-3] 转成 \\upcite{}
"""
import argparse
import os
import re
import sys

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    sys.exit("需要 python-docx：pip install python-docx")

SPECIALS = {
    "\\": r"\textbackslash{}", "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#",
    "_": r"\_", "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}
HEADING_RE = re.compile(r"^\s*(\d+(?:\.\d+){0,5})\s*[\.、]?\s+(.+)$")
CAPTION_RE = re.compile(r"^(图|表)\s*\d+\s*[\.、:：]?\s*(.*)$")
LIST_RE = re.compile(r"^[\(（]\s*\d+\s*[\)）]\s*")



def rel_part(part, rid):
    """按 rId 取关系目标部件（兼容不同 python-docx 版本）"""
    rels = getattr(part, "rels", None)
    if rels is None:
        return None
    for getter in ("__getitem__",):
        try:
            rel = rels[rid]
            if rel is not None:
                return rel.target_part
        except Exception:
            pass
    d = getattr(rels, "_rels", None)
    if isinstance(d, dict) and rid in d:
        try:
            return d[rid].target_part
        except Exception:
            return None
    try:
        for rel in rels.values():
            if getattr(rel, "rId", None) == rid:
                return rel.target_part
    except Exception:
        pass
    return None


def esc(s):
    return "".join(SPECIALS.get(ch, ch) for ch in (s or ""))


def heading_level(par):
    st = par.style
    name = (st.name or "").strip()
    sid = (st.style_id or "").strip()
    m = re.match(r"(?i)heading\s*(\d)$", name) or re.match(r"(?i)heading(\d)$", sid)
    if m:
        return int(m.group(1))
    m = re.match(r"标题\s*(\d)$", name)
    if m:
        return int(m.group(1))
    pPr = par._p.pPr
    if pPr is not None:
        ol = pPr.find(qn("w:outlineLvl"))
        if ol is not None:
            try:
                v = int(ol.get(qn("w:val")))
            except Exception:
                v = 9
            # 0..8 才是标题级别；9 = 正文
            if 0 <= v <= 8:
                txt0 = (par.text or "").strip()
                if len(txt0) <= 80 and not txt0.endswith("。"):
                    return v + 1
    txt = (par.text or "").strip()
    m = HEADING_RE.match(txt)
    if m and len(m.group(1).split(".")) <= 3 and len(m.group(2)) <= 40 and not txt.endswith("。"):
        return len(m.group(1).split("."))
    return 0


def is_caption(par):
    if "caption" in (par.style.name or "").lower() or "题注" in (par.style.name or ""):
        return True
    return bool(CAPTION_RE.match((par.text or "").strip()))


def is_list(par):
    if "list" in (par.style.style_id or "").lower() or "列表" in (par.style.name or ""):
        return True
    return bool(LIST_RE.match((par.text or "").strip()))


def para_images(par):
    out = []
    for blip in par._p.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if rid:
            out.append(rid)
    return out


def collect(doc, figdir, prefix="", start=0):
    """把正文分成块：('h', lvl, text) / ('p', text) / ('li', text) /
       ('img', [paths]) / ('cap', kind, text) / ('tbl', latex)"""
    blocks = []
    fig_n = start
    for el in doc.element.body.iterchildren():
        tag = el.tag.split("}")[-1]
        if tag == "tbl":
            blocks.append(("tbl", None))
            blocks[-1] = ("tbl", Table(el, doc))
            continue
        if tag != "p":
            continue
        par = Paragraph(el, doc)
        rids = para_images(par)

        if rids:
            paths = []
            for rid in rids:
                fig_n += 1
                part = rel_part(par.part, rid)
                if part is None:
                    continue
                ext = os.path.splitext(str(part.partname))[1] or ".png"
                name = "%sfig%02d%s" % (prefix, fig_n, ext)
                with open(os.path.join(figdir, name), "wb") as fh:
                    fh.write(part.blob)
                paths.append("figures/" + name)
            if paths:
                blocks.append(("img", paths))
            continue

        txt = (par.text or "").strip()
        if not txt:
            continue
        if is_caption(par):
            m = CAPTION_RE.match(txt)
            kind = m.group(1) if m else "图"
            body = m.group(2) if m else txt
            blocks.append(("cap", kind, body, txt))
            continue
        lvl = heading_level(par)
        if lvl:
            m = HEADING_RE.match(txt)
            clean = m.group(2) if m else txt
            blocks.append(("h", lvl, clean.strip()))
            continue
        if is_list(par):
            blocks.append(("li", LIST_RE.sub("", txt)))
            continue
        blocks.append(("p", txt))
    return blocks


def table_latex(tbl, caption):
    rows = []
    for row in tbl.rows:
        cells = []
        for cell in row.cells:
            cells.append(esc(" ".join(p.text.strip() for p in cell.paragraphs if p.text.strip())))
        rows.append(cells)
    if not rows:
        return ""
    ncol = max(len(r) for r in rows)
    spec = "l" + "c" * (ncol - 1)
    lines = [r"\begin{table}[htbp]", r"  \centering"]
    if caption:
        lines.append(r"  \caption{%s}" % esc(caption))
    lines.append(r"  \begin{tabular}{%s}" % spec)
    lines.append(r"    \toprule")
    for i, r in enumerate(rows):
        r = r + [""] * (ncol - len(r))
        lines.append("    " + " & ".join(r) + r" \\")
        if i == 0:
            lines.append(r"    \midrule")
    lines.append(r"    \bottomrule")
    lines.append(r"  \end{tabular}")
    lines.append(r"\end{table}")
    return "\n".join(lines)


def render(blocks, upcite=False):
    out = []
    i = 0
    in_refs = False
    refs = []

    def fig_env(paths, caption):
        lines = [r"\begin{figure}[htbp]", r"  \centering"]
        for p in paths:
            lines.append(r"  \includegraphics[width=0.85\linewidth]{%s}" % p)
        if caption:
            lines.append(r"  \caption{%s}" % esc(caption))
        lines.append(r"\end{figure}")
        return "\n".join(lines)

    while i < len(blocks):
        b = blocks[i]
        kind = b[0]

        if kind == "h":
            lvl, text = b[1], b[2]
            if "参考文献" in text:
                in_refs = True
                out.append(r"\section*{参考文献}")
                i += 1
                continue
            if lvl == 1:
                out.append(r"\section{%s}" % esc(text))
            elif lvl == 2:
                out.append(r"\subsection{%s}" % esc(text))
            elif lvl == 3:
                out.append(r"\subsubsection{%s}" % esc(text))
            else:
                out.append(r"\begin{enumerate}[label=(\arabic*),leftmargin=2em]")
                out.append(r"  \item %s" % esc(text))
                # 收集其后的正文作为列表项内容
                i += 1
                while i < len(blocks) and blocks[i][0] == "p":
                    out.append(r"  \item %s" % esc(blocks[i][1]))
                    i += 1
                out.append(r"\end{enumerate}")
                continue
            i += 1
            continue

        if kind == "img":
            caption = ""
            nxt = blocks[i + 1] if i + 1 < len(blocks) else None
            if nxt and nxt[0] == "cap" and nxt[1] == "图":
                caption = nxt[2]
                i += 1
            out.append(fig_env(b[1], caption))
            i += 1
            continue

        if kind == "tbl":
            caption = ""
            nxt = blocks[i + 1] if i + 1 < len(blocks) else None
            if nxt and nxt[0] == "cap" and nxt[1] == "表":
                caption = nxt[2]
                i += 1
            elif out and out[-1].startswith(r"\begin{center}\small 表"):
                out.pop()  # 表题在表上方
                prev = blocks[i - 1]
                caption = prev[2]
            out.append(table_latex(b[1], caption))
            i += 1
            continue

        if kind == "li":
            items = []
            while i < len(blocks) and blocks[i][0] == "li":
                items.append(blocks[i][1])
                i += 1
            out.append(r"\begin{enumerate}[label=(\arabic*),leftmargin=2em]")
            for it in items:
                out.append(r"  \item %s" % esc(it))
            out.append(r"\end{enumerate}")
            continue

        if kind == "cap":
            # 没有配到图/表的孤立题注：保留原文
            full = (b[3] if len(b) > 3 else ("%s%s" % (b[1], b[2]))).strip()
            out.append(r"\begin{center}\small %s\end{center}" % esc(full))
            i += 1
            continue

        # 普通段落
        text = b[1]
        if in_refs and re.match(r"^\[\d+\]", text):
            refs.append(text)
        else:
            if upcite:
                text = re.sub(r"\[(\d+(?:\s*[-,–]\s*\d+)*)\]", r"\\upcite{\1}", text)
            out.append(esc(text))
        i += 1

    if refs:
        out.append(r"\begin{thebibliography}{99}")
        for k, r in enumerate(refs, 1):
            r = re.sub(r"^\[\d+\]\s*", "", r)
            out.append(r"  \bibitem{ref%d} %s" % (k, esc(r)))
        out.append(r"\end{thebibliography}")
    return "\n\n".join(out)


PREAMBLE = r"""%% !TEX program = xelatex
%% 由 docx2tex.py 自动生成
\documentclass[UTF8,zihao=-4,a4paper,fontset=windows]{{ctexart}}

%% ===== 页面：上下左右 {t} {b} {l} {r} cm =====
\usepackage[top={t}cm,bottom={b}cm,left={l}cm,right={r}cm]{{geometry}}
\usepackage{{graphicx}}
\usepackage{{booktabs}}
\usepackage{{enumitem}}
\usepackage{{caption}}
\usepackage[hidelinks]{{hyperref}}
\usepackage{{natbib}}
\setcitestyle{{super,sort&compress,numbers}}

\setmainfont{{Times New Roman}}
\pagestyle{{plain}}

%% ===== 标题：1. / 1.1. / 1.1.1.（三级封顶） =====
\renewcommand{{\thesection}}{{\arabic{{section}}.}}
\renewcommand{{\thesubsection}}{{\thesection\arabic{{subsection}}.}}
\renewcommand{{\thesubsubsection}}{{\thesubsection\arabic{{subsubsection}}.}}

\ctexset{{
  section = {{
    format = \heiti\zihao{{4}}\bfseries,
    aftername = \hspace{{0.55em}},
    beforeskip = 1.0ex plus .2ex, afterskip = .6ex plus .1ex,
  }},
  subsection = {{
    format = \heiti\zihao{{-4}}\bfseries,
    aftername = \hspace{{0.55em}},
    beforeskip = .8ex plus .2ex, afterskip = .4ex plus .1ex,
  }},
  subsubsection = {{
    format = \heiti\zihao{{-4}}\bfseries,
    aftername = \hspace{{0.55em}},
    beforeskip = .6ex plus .2ex, afterskip = .3ex plus .1ex,
  }},
}}
\setcounter{{secnumdepth}}{{3}}      %% 四级及更深：改用有序列表

%% ===== 图题：图1. 说明文字（内容） =====
\captionsetup{{labelsep=period, font=small}}
\renewcommand{{\figurename}}{{图}}
\renewcommand{{\tablename}}{{表}}
\renewcommand{{\refname}}{{}}
\providecommand{{\bibname}}{{参考文献}}

\setlength{{\parindent}}{{2em}}
\linespread{{1.5}}

\begin{{document}}

"""


def convert(src, dst, margins=None, upcite=False):
    doc = Document(src)
    sec = doc.sections[0]
    m = margins or [sec.top_margin.cm, sec.bottom_margin.cm,
                    sec.left_margin.cm, sec.right_margin.cm]
    outdir = os.path.dirname(os.path.abspath(dst)) or "."
    figdir = os.path.join(outdir, "figures")
    os.makedirs(figdir, exist_ok=True)

    blocks = collect(doc, figdir)
    body = render(blocks, upcite=upcite)
    head = PREAMBLE.format(t="%.2f" % m[0], b="%.2f" % m[1],
                           l="%.2f" % m[2], r="%.2f" % m[3])
    with open(dst, "w", encoding="utf-8") as fh:
        fh.write(head + body + "\n\n\\end{document}\n")
    return dst


def main():
    ap = argparse.ArgumentParser(description="Word 大报告 -> LaTeX")
    ap.add_argument("input")
    ap.add_argument("-o", "--output")
    ap.add_argument("--margins", help="上,下,左,右（cm）")
    ap.add_argument("--upcite", action="store_true")
    a = ap.parse_args()
    out = a.output or os.path.splitext(a.input)[0] + ".tex"
    margins = [float(x) for x in a.margins.split(",")] if a.margins else None
    convert(a.input, out, margins=margins, upcite=a.upcite)
    print("written:", out)


if __name__ == "__main__":
    main()
