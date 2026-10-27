# -*- coding: utf-8 -*-
"""pbltex.pdfread —— 直接读 PDF，还原成与 docx 相同的块结构

产出与 core.collect() 一致的块：
    ("h", 层级, 文本) / ("p", 文本)
    ("img", [图片路径]) / ("cap", 种类, 说明, 原文)

PDF 里的**图**和**表**都按“区域裁图”处理（矢量图也能保住原样），
图题/表题作为紧随其后的块 → 渲染时进入同一个 figure 环境，
保证题注和图/表永远黏在一起。

依赖：pdfplumber（文字与版式）+ pypdfium2（区域渲染），均为宽松许可。
"""
from __future__ import annotations

import os
import re
from collections import Counter

from .core import _clean_heading, norm

CAP_RE = re.compile(r"^(图|表)\s*(\d+|[一二三四五六七八九十]+)\s*[.．、:：]?\s*(.*)$")
REF_WORDS = ("参考文献", "主要参考文献", "References", "REFERENCES")
CJK_RE = re.compile(r"[\u3400-\u9fff]")


def _fam(name: str) -> str:
    return re.sub(r"[-,].*$", "", (name or "").split("+")[-1])


def _line_text(line) -> str:
    return norm("".join(ch["text"] for ch in line["chars"]))


def _line_size(line) -> float:
    sizes = [round(ch["size"], 1) for ch in line["chars"]]
    return Counter(sizes).most_common(1)[0][0] if sizes else 0.0


def _line_cjk_family(line) -> str:
    fams = Counter()
    for ch in line["chars"]:
        if CJK_RE.match(ch["text"]):
            fams[_fam(ch.get("fontname"))] += 1
    return fams.most_common(1)[0][0] if fams else ""


def _line_bold(line) -> bool:
    chs = line["chars"]
    if not chs:
        return False
    b = sum(1 for ch in chs if "Bold" in (ch["fontname"] or ""))
    return b >= max(1, int(len(chs) * 0.6))


def _join(a: str, b: str) -> str:
    if not a:
        return b
    if re.search(r"[A-Za-z0-9,\.\)\]]$", a) and re.match(r"[A-Za-z0-9\(]", b):
        return a + " " + b
    return a + b


def _body_stats(pdf):
    size_cnt, fam_cnt = Counter(), Counter()
    for pg in pdf.pages:
        for ch in pg.chars:
            txt = (ch["text"] or "").strip()
            if not txt:
                continue
            size_cnt[round(ch["size"], 1)] += len(txt)
            if CJK_RE.match(txt):
                fam_cnt[_fam(ch.get("fontname"))] += len(txt)
    body = size_cnt.most_common(1)[0][0] if size_cnt else 10.5
    body_fam = fam_cnt.most_common(1)[0][0] if fam_cnt else ""
    return body, body_fam


class _Cropper:
    """整页渲染后按区域裁剪（pypdfium2）。"""

    def __init__(self, pdf_path, scale=2.4):
        self.scale = scale
        self.doc = None
        try:
            import pypdfium2 as pdfium
            self.doc = pdfium.PdfDocument(pdf_path)
        except Exception:  # noqa: BLE001
            self.doc = None
        self._pno = None
        self._img = None

    def _page_img(self, pno):
        if self.doc is None:
            return None
        if self._pno != pno:
            self._img = self.doc[pno - 1].render(scale=self.scale).to_pil()
            self._pno = pno
        return self._img

    def crop(self, pno, bbox, out_path):
        img = self._page_img(pno)
        if img is None:
            return False
        s = self.scale
        x0, top, x1, bottom = [float(v) for v in bbox]
        box = (max(0, int(x0 * s) - 3), max(0, int(top * s) - 3),
               min(img.width, int(x1 * s) + 3), min(img.height, int(bottom * s) + 3))
        if box[2] - box[0] < 12 or box[3] - box[1] < 12:
            return False
        try:
            img.crop(box).save(out_path)
            return True
        except Exception:  # noqa: BLE001
            return False


def pdf_blocks(pdf_path, figdir, prefix="", log=None):
    """PDF → blocks（与 docx 路线同一套块结构）。"""
    log = log or (lambda *_: None)
    import pdfplumber

    blocks = []
    fig_n = [0]
    crop = _Cropper(pdf_path)

    def add_image(pno, bbox):
        fig_n[0] += 1
        name = "%sfig%02d.png" % (prefix, fig_n[0])
        if crop.crop(pno, bbox, os.path.join(figdir, name)):
            blocks.append(("img", ["figures/" + name]))
            return True
        return False

    with pdfplumber.open(pdf_path) as pdf:
        body, body_fam = _body_stats(pdf)

        size_hist = Counter()
        for pg in pdf.pages:
            for l in pg.extract_text_lines(strip=True):
                t = _line_text(l)
                if t:
                    size_hist[_line_size(l)] += 1
        big = sorted({s for s in size_hist if s >= body * 1.12}, reverse=True)[:3]
        rank = {s: i + 1 for i, s in enumerate(big)}
        sub_level = min((len(big) or 1) + 1, 3)

        for pno, pg in enumerate(pdf.pages, 1):
            # ---- 表格区域（三线表没竖线 → text 竖切 + lines 横切）----
            tboxes = []
            for settings in ({"vertical_strategy": "text", "horizontal_strategy": "lines"},
                             {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
                             {}):
                try:
                    found = pg.find_tables(table_settings=settings)
                except Exception:  # noqa: BLE001
                    found = []
                for t in found:
                    bb = t.bbox
                    if (bb[2] - bb[0]) > 60 and (bb[3] - bb[1]) > 24:
                        tboxes.append(bb)
                if tboxes:
                    break

            def in_box(bb):
                cx, cy = (bb[0] + bb[2]) / 2.0, (bb[1] + bb[3]) / 2.0
                return any(b[0] - 2 <= cx <= b[2] + 2 and b[1] - 2 <= cy <= b[3] + 2 for b in tboxes)

            imgs = []
            for im in pg.images:
                bb = (im["x0"], im["top"], im["x1"], im["bottom"])
                if bb[2] - bb[0] >= 30 and bb[3] - bb[1] >= 30:
                    imgs.append(bb)

            lines = []
            for l in pg.extract_text_lines(strip=True):
                t = _line_text(l)
                if not t:
                    continue
                bb = (l["x0"], l["top"], l["x1"], l["bottom"])
                if in_box(bb):
                    continue
                lines.append({"text": t, "top": l["top"], "bottom": l["bottom"],
                              "size": _line_size(l), "bold": _line_bold(l),
                              "family": _line_cjk_family(l)})

            items = [("t", l["top"], l) for l in lines]
            items += [("img", bb[1], bb) for bb in imgs]
            items += [("img", bb[1], bb) for bb in tboxes]
            items.sort(key=lambda x: (x[1], 0 if x[0] == "t" else 1))

            para = []
            pending_head = [0]

            def flush():
                if not para:
                    return
                txt = para[0]["text"]
                for nxt in para[1:]:
                    txt = _join(txt, nxt["text"])
                para.clear()
                txt = norm(txt)
                if not txt:
                    return
                if pending_head[0]:
                    blocks.append(("h", pending_head[0], _clean_heading(txt)))
                    pending_head[0] = 0
                else:
                    blocks.append(("p", txt))

            prev_bottom = None
            for kind, _top, obj in items:
                if kind == "img":
                    flush()
                    add_image(pno, obj)
                    prev_bottom = obj[3]
                    continue

                line = obj
                txt = line["text"]

                if re.fullmatch(r"[-—–\s]*\d+[-—–\s]*", txt) or re.fullmatch(r"第?\s*\d+\s*页", txt):
                    continue

                if CAP_RE.match(txt):
                    flush()
                    m = CAP_RE.match(txt)
                    blocks.append(("cap", m.group(1), m.group(3), txt))
                    prev_bottom = line["bottom"]
                    continue

                is_head, lvl = False, 0
                if txt in REF_WORDS:
                    is_head, lvl = True, sub_level
                elif line["size"] in rank and len(txt) <= 90 and not txt.endswith("。"):
                    is_head, lvl = True, rank[line["size"]]
                elif line["family"] and body_fam and line["family"] != body_fam \
                        and len(txt) <= 60 and not txt.endswith(("。", "，", "：")):
                    is_head, lvl = True, sub_level
                elif line["bold"] and line["size"] <= body * 1.05 and len(txt) <= 40 \
                        and not txt.endswith(("。", "，", "：")):
                    is_head, lvl = True, sub_level

                if is_head:
                    flush()
                    pending_head[0] = max(1, min(lvl, 3))
                    para = [line]
                    flush()                     # 标题自成一段，别把正文吸进来
                    prev_bottom = line["bottom"]
                    continue

                if para:
                    gap = line["top"] - prev_bottom if prev_bottom is not None else 0
                    if line["size"] != para[-1]["size"] or gap > line["size"] * 1.15 + 5:
                        flush()
                para.append(line)
                prev_bottom = line["bottom"]
            flush()

    # 中文排版里表题通常在表格**上方** → 把 "cap(表) + img" 换序成 "img + cap"，
    # 这样渲染时表格和它的表题会进同一个 table 环境，不会分家。
    fixed = []
    i = 0
    while i < len(blocks):
        b = blocks[i]
        if b[0] == "cap" and b[1] == "表" and i + 1 < len(blocks) and blocks[i + 1][0] == "img":
            fixed.append(blocks[i + 1])
            fixed.append(b)
            i += 2
            continue
        fixed.append(b)
        i += 1
    blocks = fixed
    log("    PDF：%d 段、%d 图/表、%d 条题注" % (
        sum(1 for b in blocks if b[0] in ("h", "p")),
        sum(1 for b in blocks if b[0] == "img"),
        sum(1 for b in blocks if b[0] == "cap")))
    return blocks



