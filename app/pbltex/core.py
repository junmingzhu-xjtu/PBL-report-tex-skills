# -*- coding: utf-8 -*-
"""pbltex.core —— Word 稿件 → LaTeX 的核心转换库

两种产物：
  1. convert_single()  单份稿件 → ctexart（1. / 1.1. 三级标题）
  2. convert_chapters() 多份稿件 → ctexrep 四章式大报告（第X章 + 目录 + 跨文件参考文献去重）

只依赖 python-docx（打包后已内嵌），不依赖外部技能脚本。
"""
from __future__ import annotations

import os
import re
import shutil

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

# --------------------------------------------------------------------------- #
# 文本与转义
# --------------------------------------------------------------------------- #
SPECIALS = {
    "\\": r"\textbackslash{}", "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#",
    "_": r"\_", "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}


def esc(s: str) -> str:
    """转义 LaTeX 特殊字符（含中文全角字符时不处理）。"""
    return "".join(SPECIALS.get(ch, ch) for ch in (s or ""))


def norm(t: str) -> str:
    """去掉全角空格 / 制表符 / 软换行 / 多余空格。"""
    t = (t or "").replace("\u3000", " ").replace("\xa0", " ").replace("\t", " ")
    t = t.replace("\r\n", " ").replace("\r", " ").replace("\n", " ").replace("\x0b", " ")
    t = re.sub(r" {2,}", " ", t)
    return t.strip()


NUM_HEAD_RE = re.compile(r"^\s*(\d{1,2}(?:[.．]\d{1,2}){0,3})(?!\d)\s*[、.．]?\s*(\S.*)$")
REF_ITEM_RE = re.compile(r"^\[(\d+)\]\s*")
CITE_RE = re.compile(r"(?<!\\)\[(\d+(?:\s*[-,–，、]\s*\d+)*)\]")
CAPTION_RE = re.compile(r"^(图|表)\s*(\d+)\s*[.．、:：]?\s*(.*)$")
LIST_MARK_RE = re.compile(r"^[\(（]\s*\d+\s*[\)）]\s*")
BULLET_RE = re.compile(r"^[•·●○◦▪‣※\-\*]\s+")
LABEL_RE = re.compile(r"^([^：:。，,；;\.\(\)（）]{2,14})[：:]\s*(.+)$")


def style_name(par) -> str:
    """段落样式名；WPS 转换的 docx 可能 style=None，这里绝不抛异常。"""
    st = getattr(par, "style", None)
    return (getattr(st, "name", "") or "")


def style_id(par) -> str:
    st = getattr(par, "style", None)
    return (getattr(st, "style_id", "") or "")


# --------------------------------------------------------------------------- #
# docx 元素判定
# --------------------------------------------------------------------------- #
def _style_level(par) -> int:
    """Heading 1-9 / 标题 1-9 样式。"""
    name = style_name(par).strip()
    sid = style_id(par).strip()
    m = re.match(r"(?i)^heading\s*(\d)$", name) or re.match(r"(?i)^heading(\d)$", sid)
    if m:
        return int(m.group(1))
    m = re.match(r"^标题\s*(\d)$", name)
    if m:
        return int(m.group(1))
    return 0


def _outline_level(par) -> int:
    """段落大纲级别（0–8 有效，9 为正文）。"""
    pPr = par._p.pPr
    if pPr is None:
        return 0
    ol = pPr.find(qn("w:outlineLvl"))
    if ol is None:
        return 0
    try:
        v = int(ol.get(qn("w:val")))
    except (TypeError, ValueError):
        return 0
    return v + 1 if 0 <= v <= 8 else 0


def _num_id(par) -> str | None:
    """段落编号 id；numId="0" 表示已取消编号。"""
    pPr = par._p.pPr
    if pPr is None:
        return None
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return None
    numId = numPr.find(qn("w:numId"))
    return numId.get(qn("w:val")) if numId is not None else None


def _all_bold(par) -> bool:
    runs = [r for r in par.runs if (r.text or "").strip()]
    if not runs:
        return False
    for r in runs:
        if r.bold is True:
            continue
        if r.bold is None and r.style is not None and r.style.font.bold is True:
            continue
        return False
    return True


def _clean_heading(txt: str) -> str:
    t = norm(txt)
    t = re.sub(r"^\s*\d{1,2}(?:[.．]\d{1,2}){0,3}\s*[、.．]?\s*", "", t)
    # 软换行合并后常留下“ （作者）”“ ，”这类多余空格
    t = re.sub(r"\s+([（()）、，。；：！？“”\"'《》])", r"\1", t)
    # “情 境 1”这种字间空格（Word 分散对齐留下的）还原成“情境 1”
    t = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", t)
    t = t.rstrip("：: ；;，,、.")
    return t.strip()


def heading_of(par):
    """返回 (层级, 干净标题)；非标题返回 (0, None)。"""
    txt = norm(par.text)
    if not txt:
        return 0, None

    lvl = _style_level(par) or _outline_level(par)
    if lvl:
        # 大纲级别可能被误用：整句、以句号结尾的当正文
        if len(txt) <= 80 and not txt.endswith(("。", "；", ";", "，")):
            clean = _clean_heading(txt)
            if clean:
                return min(lvl, 3), clean
        return 0, None

    m = NUM_HEAD_RE.match(txt)
    if m and len(txt) <= 60 and not txt.endswith(("。", "；", ";", "，")):
        title = m.group(2).strip()
        depth = m.group(1).replace("．", ".").count(".") + 1
        if len(title) >= 2:
            return min(depth, 3), _clean_heading(title)

    # 全加粗、短、不以冒号结尾 → 无编号的小标题
    if _all_bold(par) and len(txt) <= 30 and not txt.endswith(("：", ":", "，", "、")):
        return 2, _clean_heading(txt)

    return 0, None


def is_list_item(par, txt: str) -> bool:
    """真正的自动编号项目（numId 有效）或以 (1) 开头的段落。"""
    nid = _num_id(par)
    if nid is not None and nid not in ("0", ""):
        return True
    return bool(LIST_MARK_RE.match(txt) or BULLET_RE.match(txt))


def is_caption(par, txt: str) -> bool:
    name = style_name(par)
    if "caption" in name.lower() or "题注" in name:
        return True
    return bool(CAPTION_RE.match(txt))


def para_images(par):
    out = []
    for blip in par._p.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if rid:
            out.append(rid)
    return out


def rel_part(part, rid):
    """按 rId 取关系目标部件（兼容不同 python-docx 版本）。"""
    rels = getattr(part, "rels", None)
    if rels is not None:
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
    return None


# --------------------------------------------------------------------------- #
# 解析成块
# --------------------------------------------------------------------------- #
def collect(doc, figdir, prefix="", log=None):
    """把 docx 正文切成块：
        ("h", 层级, 文本) / ("p", 文本) / ("li", 文本)
        ("img", [路径]) / ("cap", 种类, 说明, 原文) / ("tbl", Table)
    """
    blocks = []
    fig_n = 0
    for el in doc.element.body.iterchildren():
        tag = el.tag.split("}")[-1]
        if tag == "tbl":
            blocks.append(("tbl", Table(el, doc)))
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
                if ext.lower() not in (".png", ".jpg", ".jpeg", ".emf", ".wmf", ".gif", ".bmp", ".tif", ".tiff"):
                    ext = ".png"
                name = "%sfig%02d%s" % (prefix, fig_n, ext)
                try:
                    with open(os.path.join(figdir, name), "wb") as fh:
                        fh.write(part.blob)
                except Exception as exc:  # noqa: BLE001
                    if log:
                        log("  [警告] 图片导出失败：%s" % exc)
                    continue
                paths.append("figures/" + name)
            if paths:
                blocks.append(("img", paths))
            continue

        txt = norm(par.text)
        if not txt:
            continue

        if is_caption(par, txt):
            m = CAPTION_RE.match(txt)
            kind = m.group(1) if m else "图"
            body = m.group(3) if m else txt
            blocks.append(("cap", kind, body, txt))
            continue

        lvl, clean = heading_of(par)
        if lvl and clean:
            blocks.append(("h", lvl, clean))
            continue

        if is_list_item(par, txt):
            item = LIST_MARK_RE.sub("", txt)
            item = BULLET_RE.sub("", item)
            blocks.append(("li", item))
            continue

        blocks.append(("p", txt))
    return blocks


def split_lead_and_refs(blocks):
    """拆出：开头标题/姓名行、正文块、参考文献条目、正文内的引用编号。"""
    # 1) 参考文献：出现“参考文献”标题之后、以 [n] 开头的段落
    refs = []
    kept = []
    in_refs = False
    for b in blocks:
        if b[0] == "h" and "参考文献" in b[2]:
            in_refs = True
            continue
        if in_refs:
            if b[0] == "p" and REF_ITEM_RE.match(b[1]):
                refs.append(b[1])
                continue
            if b[0] == "p" and not REF_ITEM_RE.match(b[1]):
                # 参考文献列表之后的杂项（致谢等）丢弃
                continue
        kept.append(b)

    # 2) 开头标题 / 姓名行（第一个标题之前）
    lead = []
    i = 0
    while i < len(kept):
        b = kept[i]
        if b[0] == "h":
            break
        if b[0] == "p" and len(b[1]) <= 60 and not b[1].endswith("。") and len(lead) < 2:
            lead.append(b[1])
            i += 1
            continue
        break
    body = kept[i:]
    return lead, body, refs


def group_label_items(blocks):
    """把连续「标签：内容」段落合成一个有序列表块 ("lab", [(标签, 内容), ...])。"""
    out = []
    i = 0
    while i < len(blocks):
        b = blocks[i]
        if b[0] == "p":
            m = LABEL_RE.match(b[1])
            if m:
                items = []
                j = i
                while j < len(blocks) and blocks[j][0] == "p":
                    m2 = LABEL_RE.match(blocks[j][1])
                    if not m2:
                        break
                    items.append((m2.group(1).strip(), m2.group(2).strip()))
                    j += 1
                if len(items) >= 2:
                    out.append(("lab", items))
                    i = j
                    continue
        out.append(b)
        i += 1
    return out


def fix_heading_levels(blocks):
    """层级防跳级：首次出现的最浅级当作 1 级；不允许一次跳两级以上。"""
    levels = [b[1] for b in blocks if b[0] == "h"]
    if not levels:
        return blocks
    top = min(levels)
    out = []
    prev = 1
    for b in blocks:
        if b[0] != "h":
            out.append(b)
            continue
        lvl = max(1, b[1] - top + 1)
        if lvl > prev + 1:
            lvl = prev + 1
        prev = lvl
        out.append(("h", min(lvl, 3), b[2]))
    return out


# --------------------------------------------------------------------------- #
# LaTeX 渲染
# --------------------------------------------------------------------------- #
PREAMBLE_COMMON = r"""%% !TEX program = xelatex
%% 由 PBL 大报告生成器生成
"""

PREAMBLE_SINGLE = PREAMBLE_COMMON + r"""\documentclass[UTF8,zihao=-4,a4paper,fontset=windows]{ctexart}

%% ===== 页面 =====
\usepackage[top=__TOP__cm,bottom=__BOT__cm,left=__LEFT__cm,right=__RIGHT__cm]{geometry}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{array}
\usepackage{tabularx}
\usepackage{enumitem}
\usepackage{caption}
\usepackage[hidelinks]{hyperref}
%% 引用：正文直接写上标编号（\textsuperscript{[1,2]}），不依赖 natbib 的两遍解析

\setmainfont{Times New Roman}
\pagestyle{plain}

%% 黑体无粗体字重：另立字族并启用伪粗体
\setCJKfamilyfont{heibold}{SimHei}[AutoFakeBold=2]
\newcommand{\heibold}{\CJKfamily{heibold}}

%% ===== 标题：1. / 1.1. / 1.1.1.（三级封顶）=====
\renewcommand{\thesection}{\arabic{section}.}
\renewcommand{\thesubsection}{\thesection\arabic{subsection}.}
\renewcommand{\thesubsubsection}{\thesubsection\arabic{subsubsection}.}

\ctexset{
  section = {
    format = \heibold\zihao{4}\bfseries,
    aftername = \hspace{0.55em},
    beforeskip = 1.0ex plus .2ex, afterskip = .6ex plus .1ex,
  },
  subsection = {
    format = \heibold\zihao{-4}\bfseries,
    aftername = \hspace{0.55em},
    beforeskip = .8ex plus .2ex, afterskip = .4ex plus .1ex,
  },
  subsubsection = {
    format = \heibold\zihao{-4}\bfseries,
    aftername = \hspace{0.55em},
    beforeskip = .6ex plus .2ex, afterskip = .3ex plus .1ex,
  },
}
\setcounter{secnumdepth}{3}

%% 图题「图1. 说明」：标签名与数字之间不留空格
\DeclareCaptionLabelFormat{pbllabel}{\CJKsetecglue{}#1#2}
\captionsetup{labelformat=pbllabel, labelsep=period, font=small}
\renewcommand{\figurename}{图}
\renewcommand{\tablename}{表}

\setlength{\parindent}{2em}
\linespread{1.5}
\setlist[enumerate]{label=(\arabic*),leftmargin=2.4em,itemsep=0.15em,topsep=0.3em}

\begin{document}

"""

PREAMBLE_CHAPTERS = PREAMBLE_COMMON + r"""\documentclass[UTF8,zihao=-4,a4paper,fontset=windows,openany]{ctexrep}

\usepackage[top=__TOP__cm,bottom=__BOT__cm,left=__LEFT__cm,right=__RIGHT__cm]{geometry}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{array}
\usepackage{tabularx}
\usepackage{enumitem}
\usepackage{caption}
\usepackage[hidelinks]{hyperref}
%% 引用：正文直接写上标编号（\textsuperscript{[1,2]}），不依赖 natbib 的两遍解析

%% ===== 每人一个问题 = 一级标题，编号在每个章内从 1. 重来 =====
%% 章：第二章 重大问题；章内：1. 题目 / 1.1 / 1.1.1（三级封顶）
\renewcommand{\thesection}{\arabic{section}.}
\renewcommand{\thesubsection}{\thesection\arabic{subsection}.}
\renewcommand{\thesubsubsection}{\thesubsection\arabic{subsubsection}.}
%% 图/表题全程连续编号：图1. 表1.（不随章重置）
\renewcommand{\thefigure}{\arabic{figure}}
\renewcommand{\thetable}{\arabic{table}}

\setmainfont{Times New Roman}
%% 微软雅黑：章标题用（自定义族名，避免与 ctex 的 \yahei 冲突）
\setCJKfamilyfont{pblyh}{Microsoft YaHei}
\newcommand{\pblyh}{\CJKfamily{pblyh}}
\pagestyle{plain}

%% ===== 章：居中、三号、微软雅黑、加粗 =====
\ctexset{
  chapter = {
    name = {第,章},
    number = \chinese{chapter},
    format = \centering\pblyh\zihao{3}\bfseries,
    aftername = \quad,
    beforeskip = 1.2ex plus .3ex,
    afterskip  = 1.0ex plus .2ex,
  },
  section = {
    format = \heibold\zihao{-4}\bfseries,
    aftername = \hspace{0.55em},
    beforeskip = .9ex plus .2ex, afterskip = .5ex plus .1ex,
  },
  subsection = {
    format = \heibold\zihao{-4}\bfseries,
    aftername = \hspace{0.55em},
    beforeskip = .7ex plus .2ex, afterskip = .4ex plus .1ex,
  },
}
\setcounter{secnumdepth}{3}

\setCJKfamilyfont{heibold}{SimHei}[AutoFakeBold=2]
\newcommand{\heibold}{\CJKfamily{heibold}}
%% 封面标题用宋体加粗（SimSun 没有粗体字重 → 伪粗体）
\setCJKfamilyfont{songbold}{SimSun}[AutoFakeBold=2]
\newcommand{\songbold}{\CJKfamily{songbold}}

%% 图题「图1. 说明」：标签名与数字之间不留空格
\DeclareCaptionLabelFormat{pbllabel}{\CJKsetecglue{}#1#2}
\captionsetup{labelformat=pbllabel, labelsep=period, font=small}
\renewcommand{\figurename}{图}
\renewcommand{\tablename}{表}
%% 参考文献标题由 thebibliography 自己排（ctexrep 里是 \chapter*，再写一遍标题会重复/空页）
\providecommand{\refname}{参考文献}
\providecommand{\bibname}{参考文献}

\setlength{\parindent}{2em}
\linespread{1.45}
\setlist[enumerate]{label=(\arabic*),leftmargin=2.4em,itemsep=0.15em,topsep=0.3em}

\begin{document}

"""


def preamble_single(margins=(2.54, 2.54, 1.91, 1.91)) -> str:
    s = PREAMBLE_SINGLE
    for key, val in zip(("__TOP__", "__BOT__", "__LEFT__", "__RIGHT__"), margins):
        s = s.replace(key, "%.2f" % val)
    return s


def preamble_chapters(margins=(2.54, 2.54, 1.91, 1.91)) -> str:
    s = PREAMBLE_CHAPTERS
    for key, val in zip(("__TOP__", "__BOT__", "__LEFT__", "__RIGHT__"), margins):
        s = s.replace(key, "%.2f" % val)
    return s


def _disp_len(s: str) -> float:
    """按显示宽度估字数：中日韩/全角算 1，其余算 0.55。"""
    n = 0.0
    for ch in s or "":
        n += 1.0 if ord(ch) > 0x2E7F else 0.55
    return n


def _table_rows(tbl):
    """取表行；合并单元格只取一次并记住跨列数。"""
    rows = []
    ncol = 0
    for row in tbl.rows:
        cells, seen = [], set()
        for cell in row.cells:
            key = id(cell._tc)
            if key in seen:            # 合并单元格在 row.cells 里会重复出现
                continue
            seen.add(key)
            try:
                span = max(1, min(int(cell._tc.grid_span), 20))
            except Exception:          # noqa: BLE001
                span = 1
            text = esc(norm(" ".join(p.text for p in cell.paragraphs if p.text.strip())))
            cells.append((text, span))
        rows.append(cells)
        ncol = max(ncol, sum(s for _, s in cells))
    return rows, ncol


def table_tex(tbl, caption="", text_width_em=40.0) -> str:
    """Word 表格 → 三线表。

    宽度自适应，绝不溢出正文：
      · 内容较窄 → 自然宽度 tabular；
      · 列多或字多 → tabularx（长列自动折行，宽度 = \\linewidth）；
      · 全是短列但列特别多 → resizebox 等比缩到页宽。
    """
    rows, ncol = _table_rows(tbl)
    if not rows or ncol <= 0:
        return ""

    # 每列最大显示宽度（合并单元格按跨列分摊）
    colw = [0.0] * ncol
    for cells in rows:
        i = 0
        for text, span in cells:
            if i >= ncol:
                break
            w = _disp_len(text) / span
            for k in range(i, min(i + span, ncol)):
                colw[k] = max(colw[k], w)
            i += span

    est = sum(colw) + 1.0 * ncol + 0.6          # tabcolsep 与竖线余量
    ratio = est / text_width_em if text_width_em else 1.0
    n_long = sum(1 for w in colw if w > 6.0)

    # 有文字列 → 折行（保留字号，最耐读）；纯短列又超宽 → 等比缩小。
    use_x = (not (ratio <= 0.95)) and (
        (1 <= n_long <= 3 and ncol <= 8) or (n_long >= 1 and ratio <= 1.25))

    if ratio <= 0.95:
        spec = "l" + "c" * (ncol - 1)
        env_begin = r"\begin{tabular}{%s}" % spec
        env_end = r"\end{tabular}"
        font = ""
    elif use_x:
        cols = []
        for w in colw:
            cols.append(r">{\raggedright\arraybackslash}X" if w > 6.0 else "c")
        env_begin = r"\begin{tabularx}{\linewidth}{%s}" % "".join(cols)
        env_end = r"\end{tabularx}"
        font = r"\small" if ratio > 1.2 else ""
    else:
        # 列太多 / 太宽：整体等比缩到版心宽（字会小一点，但版面干净、绝不越界）
        spec = "l" + "c" * (ncol - 1)
        env_begin = r"\resizebox{\linewidth}{!}{\begin{tabular}{%s}" % spec
        env_end = r"\end{tabular}}"
        font = ""

    lines = [r"\begin{table}[htbp]", r"  \centering"]
    if font:
        lines.append("  " + font)
    if caption:
        lines.append(r"  \caption{%s}" % esc(caption))
    lines.append(r"  \setlength{\tabcolsep}{4pt}")
    lines.append(r"  \renewcommand{\arraystretch}{1.15}")
    lines.append("  " + env_begin)
    lines.append(r"    \toprule")
    for ri, cells in enumerate(rows):
        parts = []
        for text, span in cells:
            if span > 1:
                parts.append(r"\multicolumn{%d}{c}{%s}" % (span, text))
            else:
                parts.append(text)
        filled = sum(s for _, s in cells)
        if filled < ncol:
            parts.append(r"\multicolumn{%d}{c}{}" % (ncol - filled))
        lines.append("    " + " & ".join(parts) + r" \\")
        if ri == 0:
            lines.append(r"    \midrule")
    lines.append(r"    \bottomrule")
    lines.append("  " + env_end)
    lines.append(r"\end{table}")
    return "\n".join(lines)


def figure_tex(paths, caption="", width=0.85) -> str:
    lines = [r"\begin{figure}[htbp]", r"  \centering"]
    for p in paths:
        lines.append(r"  \includegraphics[width=%s\linewidth]{%s}" % (("%g" % width), p))
    if caption:
        lines.append(r"  \caption{%s}" % esc(caption))
    lines.append(r"\end{figure}")
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# 封面（照 23 级 Case 报告的样式：校徽 + 三行标题 + Tutor/主席/书记/成员）
# --------------------------------------------------------------------------- #
COVER_MEMBER_PER_LINE = 4


def _split_names(members):
    if isinstance(members, str):
        members = re.split(r"[\s,，、;；]+", members)
    return [m.strip() for m in (members or []) if m and m.strip()]


def cover_tex(cover, figdir, log=None) -> str:
    """生成 titlepage。cover 为 None / 空则返回空串（调用方退回 \\maketitle）。

    版式（老师/同学的约定）：
      校徽
      2023级11组PBL学习   一号（26pt）居中，行距 3 倍
      学习报告            小一（24pt）居中，行距 3 倍
      Case：xxx           小初（36pt）居中，行距 3 倍
      （空一行，小初的一行高度）
      Tutor / 主席 / 书记 / 成员   三号（16pt），行距 1.5
      日期（页脚）
    """
    if not cover:
        return ""
    log = log or (lambda *_: None)
    g = dict(cover)

    line1 = (g.get("line1") or "").strip()
    line2 = (g.get("line2") or "学习报告").strip()
    case = (g.get("case") or "").strip()
    tutor = (g.get("tutor") or "").strip()
    chair = (g.get("chair") or "").strip()
    secretary = (g.get("secretary") or "").strip()
    date = (g.get("date") or "").strip()
    members = _split_names(g.get("members"))
    # 主席/书记若混在成员里，先从成员中剔除（“其余为组员”）
    members = [m for m in members if m not in (chair, secretary)]

    # “Case：名称”是约定格式：作者只填名称时自动补前缀
    if case and not re.match(r"^(case|cause)\b", case, re.I) \
            and "：" not in case and ":" not in case:
        case = "Case：" + case

    logo_rel = ""
    logo = (g.get("logo") or "").strip()
    if logo:
        if os.path.isfile(logo):
            ext = os.path.splitext(logo)[1].lower() or ".png"
            name = "cover_logo" + ext
            try:
                shutil.copyfile(logo, os.path.join(figdir, name))
                logo_rel = "figures/" + name
            except OSError as exc:  # noqa: BLE001
                log("  [警告] 封面图片复制失败（%s）" % exc)
        else:
            log("  [警告] 找不到封面图片：%s" % logo)

    def fs(size, multiple):
        """字号 + 行距（行距 = 倍数 × 字号）。"""
        return r"\fontsize{%gpt}{%gpt}\selectfont" % (size, size * multiple)

    out = [r"\begin{titlepage}", r"  \thispagestyle{empty}", r"  \begin{center}"]
    if logo_rel:
        out.append(r"    \includegraphics[width=0.60\linewidth]{%s}\par" % logo_rel)
        out.append(r"    \vspace{2.2em}")

    # ---- 标题三行：一号(26) / 小一(24) / 小初(36)，行距 3 倍 ----
    # 段与段之间的基线距离 = 3 × 字号；\vspace 扣掉下一行的字高
    def title_block(size, text, kind):
        return r"    {\CJKsetecglue{}%s%s %s\par}" % (kind, fs(size, 1.2), esc(text))

    if line1:
        out.append(title_block(26, line1, r"\songbold"))
    if line2:
        if line1:
            out.append(r"    \vspace{%gpt}" % (3 * 26 - 0.85 * 24))
        out.append(title_block(24, line2, r"\songbold"))
    if case:
        if line1 or line2:
            out.append(r"    \vspace{%gpt}" % (3 * 24 - 0.85 * 36))
        out.append(r"    \begin{minipage}{0.94\linewidth}\centering")
        out.append(r"      " + title_block(36, case, r"\bfseries").strip())
        out.append(r"    \end{minipage}")
    out.append(r"  \end{center}")

    # ---- 空一行（小初字号的整整一行）----
    out.append(r"  \vspace{%gpt}" % (36 * 1.2))

    # ---- 小组构成：三号(16pt)，行距 1.5 ----
    info_fs = fs(16, 1.5)
    rows = []
    if tutor:
        rows.append(r"\hspace{2.0em}Tutor：%s" % esc(tutor))
    if chair:
        rows.append(r"\hspace{2.0em}PBL小组主席：%s" % esc(chair))
    if secretary:
        rows.append(r"\hspace{2.0em}PBL小组书记：%s" % esc(secretary))
    if members:
        chunks = [members[i:i + COVER_MEMBER_PER_LINE]
                  for i in range(0, len(members), COVER_MEMBER_PER_LINE)]
        rows.append(r"\hspace{2.0em}PBL小组成员：%s"
                    % r"\hspace{0.85em}".join(esc(m) for m in chunks[0]))
        for ch in chunks[1:]:
            rows.append(r"\hspace*{11.2em}%s" % r"\hspace{0.85em}".join(esc(m) for m in ch))
    if rows:
        out.append(r"  \begin{flushleft}")
        out.append(r"    {%s %s\par}" % (info_fs, (r"\\" + "\n    ").join(rows)))
        out.append(r"  \end{flushleft}")

    if date:
        out.append(r"  \vfill")
        out.append(r"  \begin{center}{\CJKsetecglue{}%s %s\par}\end{center}"
                   % (fs(16, 1.5), esc(date)))
    out.append(r"\end{titlepage}")
    return "\n".join(out)


# --------------------------------------------------------------------------- #
# 引用处理
# --------------------------------------------------------------------------- #
def _expand_cite(tok: str):
    nums = []
    for part in re.split(r"[,，、]", tok):
        part = part.strip()
        m = re.match(r"^(\d+)\s*[-–]\s*(\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if 0 < b - a <= 30:
                nums.extend(range(a, b + 1))
            else:
                nums.extend([a, b])
        elif part.isdigit():
            nums.append(int(part))
    return nums


def _fmt_nums(nums) -> str:
    """[1,2,5,6,7] → '1,2,5-7'（GB/T 7714 的压缩写法）。"""
    nums = sorted({int(n) for n in nums if int(n) > 0})
    if not nums:
        return ""
    runs, start, prev = [], nums[0], nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
            continue
        runs.append((start, prev))
        start = prev = n
    runs.append((start, prev))
    out = []
    for a, b in runs:
        if a == b:
            out.append(str(a))
        elif b == a + 1:
            out.extend([str(a), str(b)])
        else:
            out.append("%d-%d" % (a, b))
    return ",".join(out)


def apply_citations(text: str, mapping: dict) -> str:
    """把正文 [1]、[2-3] 直接写成上标编号，例如 \\textsuperscript{[1,2]}。

    不依赖 \\cite 的两遍解析：即使只编译一遍（或第二遍失败），也不会出现 [?,?,?]。
    映射不到的编号（原稿引了但参考文献里没有）保留原文，方便人工发现。
    """
    def repl(m):
        nums = _expand_cite(m.group(1))
        mapped = [mapping[n] for n in nums if n in mapping]
        if not mapped or len(mapped) != len(nums):
            return m.group(0)
        return "\\textsuperscript{[%s]}" % _fmt_nums(mapped)
    return CITE_RE.sub(repl, text)


# --------------------------------------------------------------------------- #
# 渲染块
# --------------------------------------------------------------------------- #
class RenderCtx:
    def __init__(self, heading_map, cite_map, log=None):
        self.heading_map = heading_map   # 文件内层级 -> output level
        self.cite_map = cite_map
        self.log = log or (lambda *_: None)


def render_blocks(blocks, ctx: RenderCtx) -> list:
    out = []
    i = 0
    while i < len(blocks):
        b = blocks[i]

        if b[0] == "h":
            lvl = ctx.heading_map.get(b[1], 0)
            text = esc(b[2])
            if lvl == 1:
                out.append(r"\section{%s}" % text)
            elif lvl == 2:
                out.append(r"\subsection{%s}" % text)
            elif lvl == 3:
                out.append(r"\subsubsection{%s}" % text)
            else:
                out.append(r"\noindent\textbf{%s}\par" % text)
            i += 1
            continue

        if b[0] == "img":
            cap = ""
            if i + 1 < len(blocks) and blocks[i + 1][0] == "cap" and blocks[i + 1][1] == "图":
                cap = blocks[i + 1][2]
                i += 1
            out.append(figure_tex(b[1], cap))
            i += 1
            continue

        if b[0] == "tbl":
            cap = ""
            if i + 1 < len(blocks) and blocks[i + 1][0] == "cap" and blocks[i + 1][1] == "表":
                cap = blocks[i + 1][2]
                i += 1
            tex = table_tex(b[1], cap)
            if tex:
                out.append(tex)
            i += 1
            continue

        if b[0] == "li":
            items = []
            while i < len(blocks) and blocks[i][0] == "li":
                items.append(blocks[i][1])
                i += 1
            out.append(r"\begin{enumerate}")
            for it in items:
                out.append(r"  \item %s" % apply_citations(esc(it), ctx.cite_map))
            out.append(r"\end{enumerate}")
            continue

        if b[0] == "lab":
            out.append(r"\begin{enumerate}")
            for label, body in b[1]:
                out.append(r"  \item \textbf{%s：}%s" % (esc(label), apply_citations(esc(body), ctx.cite_map)))
            out.append(r"\end{enumerate}")
            i += 1
            continue

        if b[0] == "cap":
            full = b[3] if len(b) > 3 else ("%s%s" % (b[1], b[2]))
            out.append(r"\begin{center}\small %s\end{center}" % esc(full))
            i += 1
            continue

        out.append(apply_citations(esc(b[1]), ctx.cite_map))
        i += 1
    return out


# --------------------------------------------------------------------------- #
# 顶层：单份
# --------------------------------------------------------------------------- #
DEFAULT_MARGINS = (2.54, 2.54, 1.91, 1.91)


def _safe_stem(path: str) -> str:
    stem = os.path.splitext(os.path.basename(path))[0]
    stem = re.sub(r'[\\/:*?"<>|\r\n]+', "_", stem).strip() or "report"
    return stem


def convert_single(docx_path, outdir, title="", author="", compile_pdf=False,
                   margins=DEFAULT_MARGINS, log=None) -> dict:
    """单份 Word → LaTeX（+ 可选 PDF）。输出到 outdir 下同名子文件夹。"""
    log = log or (lambda *_: None)
    if not os.path.isfile(docx_path):
        raise FileNotFoundError("找不到文件：%s" % docx_path)
    if docx_path.lower().endswith(".doc"):
        raise ValueError("不支持 .doc 旧格式，请先在 Word 里“另存为” .docx")

    stem = _safe_stem(docx_path)
    workdir = os.path.join(outdir, stem)
    figdir = os.path.join(workdir, "figures")
    os.makedirs(figdir, exist_ok=True)

    doc = Document(docx_path)
    blocks = collect(doc, figdir, prefix="", log=log)
    lead, body, refs = split_lead_and_refs(blocks)

    if not title:
        title = lead[0] if lead else stem
    if not author:
        author = lead[1] if len(lead) > 1 else ""

    body = fix_heading_levels(group_label_items(body))

    cite_map = {}
    for k in range(1, len(refs) + 1):
        cite_map[k] = k

    ctx = RenderCtx(heading_map={1: 1, 2: 2, 3: 3}, cite_map=cite_map, log=log)
    body_tex = render_blocks(body, ctx)

    head = []
    head.append(r"\begin{center}")
    head.append(r"  {\zihao{5} PBL 大报告}")
    head.append(r"  \\[0.7em]")
    head.append(r"  {\heibold\zihao{3}\bfseries %s}" % esc(title))
    if author:
        head.append(r"  \\[0.8em]")
        head.append(r"  {\zihao{-4} %s}" % esc(author))
    head.append(r"\end{center}")
    head.append(r"")
    head.append(r"\vspace{0.2em}")
    head.append(r"\hrule height 0.6pt")
    head.append(r"\vspace{1.0em}")

    parts = [preamble_single(margins), "\n".join(head), "", "\n\n".join(body_tex)]

    if refs:
        bib = [r"\begin{thebibliography}{99}"]
        for k, r in enumerate(refs, 1):
            text = REF_ITEM_RE.sub("", r)
            bib.append(r"  \bibitem{ref%d} %s" % (k, esc(norm(text))))
        bib.append(r"\end{thebibliography}")
        parts.append("\n".join(bib))

    parts.append(r"\end{document}")

    tex_path = os.path.join(workdir, stem + ".tex")
    with open(tex_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(parts) + "\n")

    info = {
        "tex": tex_path,
        "workdir": workdir,
        "title": title,
        "headings": sum(1 for b in body if b[0] == "h"),
        "refs": len(refs),
        "figures": sum(len(b[1]) for b in body if b[0] == "img"),
        "pdf": "",
    }
    if compile_pdf:
        from .toolchain import compile_tex
        info["pdf"] = compile_tex(tex_path, log=log)
    return info


# --------------------------------------------------------------------------- #
# 顶层：四章式大报告
# --------------------------------------------------------------------------- #
CHAPTER_NAMES = ["案例回顾", "重大问题", "次要问题", "机制图"]


def _looks_like_person(t: str) -> bool:
    """判断某行是不是“姓名+学号”这类作者行。"""
    t = (t or "").strip()
    if re.search(r"\d{4,}", t):
        return True
    return any(k in t for k in ("学号", "规培", "班级", "电话", "邮箱", "成员"))


def _same_text(a: str, b: str) -> bool:
    na = re.sub(r"[\s\W_]+", "", a or "")
    nb = re.sub(r"[\s\W_]+", "", b or "")
    return bool(na) and na == nb


def _pick_section_title(lead, content, base):
    """决定这份稿件在大报告里的“一级标题”，并说明要不要丢掉稿件的第一个标题。

    顺序：开头的题目行 → 稿件第一个标题 → 文件名。
    如果题目行和第一个标题是同一句话，就只留一个标题。
    """
    title = ""
    for t in lead:
        if t and not _looks_like_person(t):
            title = t.strip()
            break
    first_heading = content[0][2] if content and content[0][0] == "h" else ""

    if title and first_heading and _same_text(title, first_heading):
        return title, True
    if title:
        return title, False
    if first_heading:
        return first_heading, True

    stem = os.path.splitext(base)[0]
    stem = re.sub(r"^\s*\d+\s*[-_.、\s]+", "", stem).strip()
    return stem or base, False


def convert_chapters(items, outdir, title="PBL 大报告", author="", date="",
                     tex_name="大报告.tex", compile_pdf=False,
                     margins=DEFAULT_MARGINS, log=None,
                     chapter_titles=None, chapter_figures=None, cover=None,
                     section_titles=None) -> dict:
    """items: [(docx_path, 章序号 0-3), ...]，按章顺序输出 ctexrep 大报告。

    chapter_titles: {章序号: 章标题}，覆盖默认的「案例回顾/重大问题/…」；
    chapter_figures: {章序号: [(图片路径, 题注), ...]}，用于“机制图”章直接放图；
    section_titles: {docx 路径或文件名: 一级标题}，人工指定（界面里双击可改）。
    """
    log = log or (lambda *_: None)
    os.makedirs(outdir, exist_ok=True)
    figdir = os.path.join(outdir, "figures")
    os.makedirs(figdir, exist_ok=True)

    refs = {}          # 规范化文本 -> 全局序号
    ref_list = []
    ref_preview = {}

    titles = dict(chapter_titles or {})
    figs_by_chap = dict(chapter_figures or {})

    body = []
    manifest = []
    groups = {}
    for path, chap in items:
        groups.setdefault(int(chap), []).append(path)
    for chap in figs_by_chap:                 # 只有图的章（如“机制图”）也要出现
        groups.setdefault(int(chap), [])

    warned = []
    for seq, idx in enumerate(sorted(groups), 1):     # 章号按实际输出的顺序编
        cname = titles.get(idx) or (CHAPTER_NAMES[idx] if 0 <= idx < len(CHAPTER_NAMES) else "其他")
        body.append(r"\chapter{%s}" % esc(cname))
        manifest.append("第%d章 %s" % (seq, cname))
        used_titles = set()                            # 本章内已用过的一级标题（查重名）

        for path in groups[idx]:
            base = os.path.basename(path)
            if not os.path.isfile(path):
                warned.append(base)
                log("  [跳过] 找不到文件：%s" % base)
                manifest.append("    ├─ (%s 缺失)" % base)
                continue
            prefix = re.sub(r"[^\w\u4e00-\u9fff]+", "", os.path.splitext(base)[0])[:10] + "_"
            try:
                doc = Document(path)
            except Exception as exc:  # noqa: BLE001
                warned.append(base)
                log("  [跳过] %s 打不开（%s）" % (base, exc))
                continue

            blocks = collect(doc, figdir, prefix=prefix, log=log)
            lead, content, file_refs = split_lead_and_refs(blocks)
            content = fix_heading_levels(group_label_items(content))

            # 每份稿件 = 大报告里的一个一级标题（1. 题目），稿内标题整体降一级：
            # 文件的 1 级 → 1.1，2 级 → 1.1.1，再深 → 加粗小标题 / (1)(2)(3)
            heading_map = {1: 2, 2: 3, 3: 0}
            sec_title, drop_first_heading = _pick_section_title(lead, content, base)
            origin = lead[0] if lead else ""
            if drop_first_heading and content and content[0][0] == "h":
                # 稿件的第一个标题被拿去当“一级标题”了，剩下部分要重新归一化，
                # 否则它下面的二级标题会被当成三级，出现 \section 后直接 \subsubsection 的跳级。
                content = fix_heading_levels(content[1:])

            # 人工指定的一级标题优先（界面里双击可改）
            override = (section_titles or {}).get(path) or (section_titles or {}).get(base)
            if override:
                sec_title = norm(override)
            # 同一章里两条一级标题撞名（比如两份情景稿都叫“内分泌系统与疾病”）→ 用文件名兜底
            elif sec_title in used_titles:
                stem = os.path.splitext(base)[0]
                sec_title = stem if stem and stem != sec_title else "%s（%d）" % (sec_title, len(used_titles) + 1)
            used_titles.add(sec_title)

            body.append("%% 稿件：%s　一级标题：%s" % (base, sec_title))
            body.append(r"\section{%s}" % esc(sec_title))
            manifest.append("    ├─ %s　→ 一级标题：%s" % (base, sec_title))

            # 参考文献去重 + 引用重编号
            mapping = {}
            for r in file_refs:
                m = REF_ITEM_RE.match(r)
                if not m:
                    continue
                old = int(m.group(1))
                text = norm(REF_ITEM_RE.sub("", r))
                key = re.sub(r"\s+", "", text)
                if key not in refs:
                    ref_list.append(text)
                    refs[key] = len(ref_list)
                    ref_preview[key] = text[:80]
                mapping[old] = refs[key]

            ctx = RenderCtx(heading_map=heading_map, cite_map=mapping, log=log)
            body.extend(render_blocks(content, ctx))

            if file_refs:
                log("    %s：%d 个标题、%d 条文献" % (base, sum(1 for b in content if b[0] == "h"), len(file_refs)))
            else:
                log("    %s：%d 个标题" % (base, sum(1 for b in content if b[0] == "h")))

        # 本章附带的图片（如“机制图”章）——编号与题注跟正文一致
        for k, (fpath, fcap) in enumerate(figs_by_chap.get(idx) or [], 1):
            if not os.path.isfile(fpath):
                warned.append(os.path.basename(fpath))
                log("  [跳过] 找不到图片：%s" % fpath)
                manifest.append("    ├─ (%s 缺失)" % os.path.basename(fpath))
                continue
            ext = os.path.splitext(fpath)[1].lower() or ".png"
            name = "chap%d_fig%02d%s" % (seq, k, ext)
            try:
                shutil.copyfile(fpath, os.path.join(figdir, name))
            except OSError as exc:  # noqa: BLE001
                warned.append(os.path.basename(fpath))
                log("  [跳过] 图片复制失败（%s）：%s" % (exc, fpath))
                continue
            body.append(figure_tex(["figures/" + name], fcap, 0.85))
            manifest.append("    ├─ [图] %s　%s" % (os.path.basename(fpath), fcap[:40]))
            log("    [图] %s → figures/%s" % (os.path.basename(fpath), name))

    if not body:
        raise ValueError("没有可用的稿件：请先添加 Word 文件并分配章节")

    if ref_list:
        body.append(r"\addcontentsline{toc}{chapter}{参考文献}")
        body.append(r"\begin{thebibliography}{99}")
        for k, r in enumerate(ref_list, 1):
            body.append(r"  \bibitem{ref%d} %s" % (k, esc(r)))
        body.append(r"\end{thebibliography}")

    head = []
    if cover:
        cv = dict(cover)
        cv.setdefault("date", date)
        head.append(cover_tex(cv, figdir, log))
    else:
        head.extend([
            r"\title{%s}" % esc(title),
            r"\author{%s}" % esc(author),
            r"\date{%s}" % esc(date),
            r"\maketitle",
        ])
    head.extend([r"\tableofcontents", r"\newpage"])
    tex_path = os.path.join(outdir, tex_name)
    with open(tex_path, "w", encoding="utf-8") as fh:
        fh.write(preamble_chapters(margins) + "\n".join(head) + "\n\n" + "\n\n".join(body) + "\n\n\\end{document}\n")

    # 分章清单
    manifest_path = os.path.join(outdir, "分章清单.txt")
    with open(manifest_path, "w", encoding="utf-8") as fh:
        fh.write("四章式大报告\n标题：%s\n组别：%s\n日期：%s\n\n" % (title, author, date))
        fh.write("\n".join(manifest) + "\n")
        fh.write("\n参考文献（跨文件去重后 %d 条）\n" % len(ref_list))
        for k, text in enumerate(ref_list, 1):
            fh.write("  [%d] %s\n" % (k, text[:100]))

    info = {
        "tex": tex_path,
        "manifest": manifest_path,
        "workdir": outdir,
        "refs": len(ref_list),
        "skipped": warned,
        "pdf": "",
    }
    if compile_pdf:
        from .toolchain import compile_tex
        info["pdf"] = compile_tex(tex_path, log=log)
    return info


# --------------------------------------------------------------------------- #
# 自动分章（读题目 + 文件名猜）
# --------------------------------------------------------------------------- #
MAJOR_KEYS = ["机制", "为什么", "为何", "因果", "通路", "调控", "调节", "如何", "原理",
              "解释", "关系", "判读", "影响", "作用", "因素", "调控网络", "生成", "分泌"]
MINOR_KEYS = ["鉴别", "概念", "定义", "分类", "补充", "细节", "名词", "区别", "简介", "概述"]
# 注意：“案例”二字太常见（很多题目叫“结合案例分析××”），只认明确的病例描述词
CASE_KEYS = ["案例回顾", "病史", "现病史", "查体", "病例摘要", "主诉", "影像学表现"]


def guess_chapter(docx_path) -> int:
    """0 案例回顾 / 1 重大问题 / 2 次要问题。

    先按“重点/次要”的关键词判，判不出来再看是不是病例描述，最后默认重大问题。
    """
    name = os.path.basename(docx_path)
    head = ""
    try:
        doc = Document(docx_path)
        for par in doc.paragraphs[:6]:
            t = norm(par.text)
            if t:
                head = t
                break
    except Exception:  # noqa: BLE001
        pass
    text = name + " " + head

    major = sum(1 for k in MAJOR_KEYS if k in text)
    minor = sum(1 for k in MINOR_KEYS if k in text)
    case = sum(1 for k in CASE_KEYS if k in text)

    if major > minor and major > 0:
        return 1
    if minor > major and minor > 0:
        return 2
    if case > 0:
        return 0
    return 1


