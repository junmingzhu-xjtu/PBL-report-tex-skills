# -*- coding: utf-8 -*-
"""生成“四章式”大报告 Word 模板
章：第一章 案例回顾 / 第二章 重大问题 / 第三章 次要问题 / 第四章 机制图
章标题：居中、三号(16pt)、微软雅黑、加粗
节：1.1、1.1.1 自动编号；三级封顶；再深用 (1)(2)(3)
"""
import io
import sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml

OUT = r"D:\desktop\大四上\PBL\大报告模板-四章.docx"
if len(sys.argv) > 1:
    OUT = sys.argv[1]
PH = r"D:\desktop\大四上\PBL\tools\img\figure_placeholder.png"

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.top_margin = sec.bottom_margin = Cm(2.54)
sec.left_margin = sec.right_margin = Cm(1.91)


def set_fonts(rPr, ascii_font="Times New Roman", east="宋体"):
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts"); rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)
    rFonts.set(qn("w:eastAsia"), east)
    rFonts.set(qn("w:cs"), ascii_font)


def style_font(style, size_pt, bold=False, east="宋体", ascii_font="Times New Roman"):
    f = style.font
    f.name = ascii_font
    f.size = Pt(size_pt)
    f.bold = bold
    f.color.rgb = RGBColor(0, 0, 0)
    set_fonts(style.element.get_or_add_rPr(), ascii_font, east)


def style_para(style, align=None, line=1.5, before=0, after=0,
               first_line_chars=0, keep_next=False, left_chars=0):
    pPr = style.element.get_or_add_pPr()
    pf = style.paragraph_format
    if align is not None:
        pf.alignment = align
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line:
        pf.line_spacing = line
    pf.keep_with_next = keep_next
    if first_line_chars or left_chars:
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind"); pPr.append(ind)
        if first_line_chars:
            ind.set(qn("w:firstLineChars"), str(int(first_line_chars * 100)))
            ind.set(qn("w:firstLine"), str(int(first_line_chars * 240)))
        if left_chars:
            ind.set(qn("w:leftChars"), str(int(left_chars * 100)))
            ind.set(qn("w:left"), str(int(left_chars * 240)))


st = doc.styles
normal = st["Normal"]
style_font(normal, 12, False, "宋体")
style_para(normal, WD_ALIGN_PARAGRAPH.JUSTIFY, line=1.5, first_line_chars=2)

h1 = st["Heading 1"]           # 章
style_font(h1, 16, True, "微软雅黑")     # 三号 16pt，微软雅黑，加粗
style_para(h1, WD_ALIGN_PARAGRAPH.CENTER, line=1.5, before=18, after=12, keep_next=True)

h2 = st["Heading 2"]           # 节 1.1
style_font(h2, 12, True, "黑体")
style_para(h2, WD_ALIGN_PARAGRAPH.LEFT, line=1.5, before=10, after=4, keep_next=True)

h3 = st["Heading 3"]           # 小节 1.1.1
style_font(h3, 12, True, "黑体")
style_para(h3, WD_ALIGN_PARAGRAPH.LEFT, line=1.5, before=8, after=3, keep_next=True)

cap = st["Caption"]
style_font(cap, 10.5, False, "宋体")
style_para(cap, WD_ALIGN_PARAGRAPH.CENTER, line=1.0, before=3, after=6)

# ---------- 编号：章计数器（隐藏显示）+ 1.1 / 1.1.1 ----------
numbering = doc.part.numbering_part.element


def add_abs(abstract_id, levels):
    xml = ['<w:abstractNum %s w:abstractNumId="%d">' % (nsdecls("w"), abstract_id)]
    xml.append('<w:multiLevelType w:val="multilevel"/>')
    for ilvl, (text, fmt) in enumerate(levels):
        xml.append('<w:lvl w:ilvl="%d"><w:start w:val="1"/>' % ilvl)
        xml.append('<w:numFmt w:val="%s"/>' % fmt)
        xml.append('<w:lvlText w:val="%s"/>' % text)
        xml.append('<w:suff w:val="space"/>')
        xml.append('<w:lvlJc w:val="left"/>')
        xml.append('<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr>')
        xml.append('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
                   'w:eastAsia="宋体" w:hint="default"/></w:rPr></w:lvl>')
    xml.append('</w:abstractNum>')
    node = parse_xml("".join(xml))
    nums = numbering.findall(qn("w:num"))
    if nums:
        nums[0].addprevious(node)
    else:
        numbering.append(node)


def add_num(num_id, abstract_id):
    numbering.append(parse_xml('<w:num %s w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>'
                               % (nsdecls("w"), num_id, abstract_id)))


# 0 级 = 章（不显示编号，只做计数器）；1 级 = 1.1；2 级 = 1.1.1
add_abs(90, [("", "decimal"), ("%1.%2", "decimal"), ("%1.%2.%3", "decimal")])
add_num(90, 90)
# (1)(2)(3)
add_abs(91, [("(%1)", "decimal"), ("(%1)", "decimal"), ("(%1)", "decimal")])
add_num(91, 91)

for ilvl, styleX in [(0, h1), (1, h2), (2, h3)]:
    pPr = styleX.element.get_or_add_pPr()
    numPr = OxmlElement("w:numPr")
    il = OxmlElement("w:ilvl"); il.set(qn("w:val"), str(ilvl))
    ni = OxmlElement("w:numId"); ni.set(qn("w:val"), "90")
    numPr.append(il); numPr.append(ni)
    pPr.append(numPr)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "0"); ind.set(qn("w:firstLine"), "0")
    pPr.append(ind)

if "List Paragraph" in [s.name for s in st]:
    lp = st["List Paragraph"]
else:
    lp = st.add_style("List Paragraph", 1)
style_font(lp, 12, False, "宋体")
style_para(lp, WD_ALIGN_PARAGRAPH.JUSTIFY, line=1.5, first_line_chars=0, left_chars=1)
pPr = lp.element.get_or_add_pPr()
numPr = OxmlElement("w:numPr")
il = OxmlElement("w:ilvl"); il.set(qn("w:val"), "0")
ni = OxmlElement("w:numId"); ni.set(qn("w:val"), "91")
numPr.append(il); numPr.append(ni)
pPr.append(numPr)

# ---------- 兼容模式 ----------
settings = doc.settings.element
compat = settings.find(qn("w:compat"))
if compat is None:
    compat = OxmlElement("w:compat"); settings.append(compat)
cs = OxmlElement("w:compatSetting")
cs.set(qn("w:name"), "compatibilityMode")
cs.set(qn("w:uri"), "http://schemas.microsoft.com/office/word")
cs.set(qn("w:val"), "15")
compat.append(cs)

# ---------- 内容 ----------
def p(text="", style=None, bold=False, size=None, align=None, indent=True):
    par = doc.add_paragraph(style=style)
    if text:
        r = par.add_run(text)
        r.bold = bold
        if size:
            r.font.size = Pt(size)
    if align is not None:
        par.alignment = align
    if not indent:
        par.paragraph_format.first_line_indent = Pt(0)
    return par


def heading(text, level):
    return doc.add_heading(text, level=level)


def list_item(text):
    par = doc.add_paragraph(style="List Paragraph")
    par.add_run(text)
    return par


def caption(kind, text):
    par = doc.add_paragraph(style="Caption")
    par.add_run(kind)
    # SEQ 域
    r = par.add_run(); f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin"); r._r.append(f1)
    r = par.add_run(); it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
    it.text = " SEQ %s \\* ARABIC " % kind; r._r.append(it)
    r = par.add_run(); f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "separate"); r._r.append(f2)
    par.add_run("1")
    r = par.add_run(); f3 = OxmlElement("w:fldChar"); f3.set(qn("w:fldCharType"), "end"); r._r.append(f3)
    par.add_run(". " + text)


# 封面标题
t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("×××课程 / PBL 大报告")
r.font.size = Pt(18); r.bold = True; r.font.name = "微软雅黑"
r._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
sub = doc.add_paragraph(); sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sub.add_run("第 × 组　　2026 年 10 月　　成员：×××、×××、×××")
r.font.size = Pt(10.5)
sub.paragraph_format.first_line_indent = Pt(0)

doc.add_paragraph()

# 第一章
heading("第一章　案例回顾", 1)
heading("病例摘要", 2)
p("××岁女性，主诉……；现病史……；既往史……。本节把病例原始信息整理清楚，不展开分析。")
heading("关键指标", 2)
p("表格示例：把 CASE 指标按“指标—检测值—参考范围—判读”列清楚。")
caption("表", "本例主要实验室指标（CASE 指标）")
tbl = doc.add_table(rows=3, cols=4)
tbl.style = "Table Grid"
for i, row in enumerate([["指标", "检测值", "参考范围", "判读"],
                         ["PRL", "30.63 ng/mL", "4.79–23.3", "轻度升高（1.3×）"],
                         ["FT4", "60.47 pmol/L", "12.0–22.0", "明显升高（2.7×）"]]):
    for j, v in enumerate(row):
        c = tbl.cell(i, j); c.text = v
        for par in c.paragraphs:
            par.paragraph_format.first_line_indent = Pt(0)
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER if j else WD_ALIGN_PARAGRAPH.LEFT
            for run in par.runs:
                run.font.size = Pt(10.5)
                if i == 0:
                    run.bold = True

# 第二章
heading("第二章　重大问题", 1)
p("本章回答“必须讲清楚”的问题：机制、因果、判读。每个问题用一个二级标题，标题里带问题名与负责人。")
heading("催乳素为什么会升高（×同学）", 2)
p("一句话结论 → 三条依据 → 一个限定条件。")
heading("下丘脑与垂体的调控基础", 3)
p("正文……（三级到此为止，再往下用有序列表）")
list_item("要点一：……")
list_item("要点二：……")
heading("甲亢与 PRL 的因果链", 2)
p("正文……，见图 1。")
doc.add_picture(PH, width=Cm(12.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.paragraphs[-1].paragraph_format.first_line_indent = Pt(0)
caption("图", "甲亢与应激抬高 PRL 的三条链（机制示意）")

# 第三章
heading("第三章　次要问题", 1)
p("本章放“可以简短回答”的问题：概念澄清、鉴别要点、细节补充。")
heading("溢乳与闭经的鉴别（×同学）", 2)
p("正文……")

# 第四章
heading("第四章　机制图", 1)
p("把全组画的机制图集中放在这里，每张图配一条图题，正文里用“见图 x”引用。")
doc.add_picture(PH, width=Cm(12.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.paragraphs[-1].paragraph_format.first_line_indent = Pt(0)
caption("图", "下丘脑—垂体—甲状腺轴与泌乳素轴的交汇（全景图）")

# 参考文献（不编号）
hpar = doc.add_paragraph()
hpar.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = hpar.add_run("参考文献")
r.bold = True; r.font.size = Pt(14); r.font.name = "微软雅黑"
r._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
hpar.paragraph_format.first_line_indent = Pt(0)
p("按 GB/T 7714—2015 顺序编码制著录：正文上标 [1]、[2-3]；Word 引用样式选 Chinese Std GBT7714 (numeric)，或 EndNote 载入同名 .ens。")
for ref in ["[1] 作者. 题名[M]. 出版地: 出版社, 年: 起止页.",
            "[2] 作者, 作者. 题名[J]. 刊名, 年, 卷(期): 起止页.",
            "[3] GRATTAN D R. 50 years of neuroendocrinology[J]. Journal of Endocrinology, 2015, 226(2): T101-T122."]:
    par = doc.add_paragraph()
    par.paragraph_format.first_line_indent = Pt(0)
    par.paragraph_format.space_after = Pt(0)
    par.add_run(ref).font.size = Pt(10.5)

doc.core_properties.title = "大报告模板（四章式）"
doc.core_properties.comments = "章：微软雅黑 三号 加粗 居中；节 1.1 / 1.1.1 小四加粗；正文小四；图题“图x. 说明（内容）”"
doc.save(OUT)
print("saved", OUT)
