# -*- coding: utf-8 -*-
"""生成大报告 Word 模板：三级标题自动编号 + 图题 + 参考文献（GB/T 7714 顺序编码制）"""
import copy
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml

OUT = r"D:\desktop\大四上\PBL\大报告模板-终版.docx"
if len(sys.argv) > 1:
    OUT = sys.argv[1]

doc = Document()

# ---------- 页面 ----------
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(2.54)
sec.bottom_margin = Cm(2.54)
sec.left_margin = Cm(1.91)
sec.right_margin = Cm(1.91)
sec.header_distance = Cm(1.5)
sec.footer_distance = Cm(1.75)

# ---------- 字体工具 ----------
def set_run_fonts(rPr, ascii_font="Times New Roman", east="宋体"):
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
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
    rPr = style.element.get_or_add_rPr()
    set_run_fonts(rPr, ascii_font, east)

def style_para(style, align=None, line=1.5, before=0, after=0,
               first_line_chars=0, keep_next=False, left_chars=0):
    pPr = style.element.get_or_add_pPr()
    if align is not None:
        style.paragraph_format.alignment = align
    pf = style.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line:
        pf.line_spacing = line
    pf.keep_with_next = keep_next
    # 首行缩进（字符）
    if first_line_chars:
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind"); pPr.append(ind)
        ind.set(qn("w:firstLineChars"), str(int(first_line_chars * 100)))
        ind.set(qn("w:firstLine"), str(int(first_line_chars * 240)))
    if left_chars:
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind"); pPr.append(ind)
        ind.set(qn("w:leftChars"), str(int(left_chars * 100)))
        ind.set(qn("w:left"), str(int(left_chars * 240)))

# ---------- 样式 ----------
st = doc.styles

normal = st["Normal"]
style_font(normal, 12, False, "宋体")          # 正文：小四 12pt
style_para(normal, WD_ALIGN_PARAGRAPH.JUSTIFY, line=1.5, before=0, after=0, first_line_chars=2)

h1 = st["Heading 1"]
style_font(h1, 14, True, "黑体")               # 一级：四号 14pt 加粗
style_para(h1, WD_ALIGN_PARAGRAPH.LEFT, line=1.5, before=12, after=6, keep_next=True)

h2 = st["Heading 2"]
style_font(h2, 12, True, "黑体")               # 二级：小四 12pt 加粗
style_para(h2, WD_ALIGN_PARAGRAPH.LEFT, line=1.5, before=6, after=3, keep_next=True)

h3 = st["Heading 3"]
style_font(h3, 12, True, "黑体")               # 三级：小四 12pt 加粗
style_para(h3, WD_ALIGN_PARAGRAPH.LEFT, line=1.5, before=6, after=3, keep_next=True)

cap = st["Caption"]
style_font(cap, 10.5, False, "宋体")           # 图题：五号
style_para(cap, WD_ALIGN_PARAGRAPH.CENTER, line=1.0, before=3, after=6)

# ---------- 编号定义 ----------
numbering = doc.part.numbering_part.element

def add_abstract_num(abstract_id, levels):
    """levels: [(lvlText, indent_chars_hanging)]"""
    xml = ['<w:abstractNum %s w:abstractNumId="%d">' % (nsdecls("w"), abstract_id)]
    xml.append('<w:multiLevelType w:val="multilevel"/>')
    for ilvl, (text, fmt) in enumerate(levels):
        xml.append('<w:lvl w:ilvl="%d">' % ilvl)
        xml.append('<w:start w:val="1"/>')
        xml.append('<w:numFmt w:val="%s"/>' % fmt)
        xml.append('<w:lvlText w:val="%s"/>' % text)
        xml.append('<w:suff w:val="space"/>')
        xml.append('<w:lvlJc w:val="left"/>')
        xml.append('<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr>')
        xml.append('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
                   'w:eastAsia="宋体" w:hint="default"/></w:rPr>')
        xml.append('</w:lvl>')
    xml.append('</w:abstractNum>')
    node = parse_xml("".join(xml))
    nums = numbering.findall(qn("w:num"))
    if nums:
        nums[0].addprevious(node)
    else:
        numbering.append(node)
    return node

def add_num(num_id, abstract_id):
    node = parse_xml('<w:num %s w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>'
                     % (nsdecls("w"), num_id, abstract_id))
    numbering.append(node)
    return node

# 多级编号：1.  1.1.  1.1.1.（挂在标题样式上）
add_abstract_num(90, [("%1.", "decimal"), ("%1.%2.", "decimal"), ("%1.%2.%3.", "decimal")])
add_num(90, 90)

# 第四级起的“有序列表”：(1) (2) (3)
add_abstract_num(91, [("(%1)", "decimal"), ("(%1)", "decimal"), ("(%1)", "decimal")])
add_num(91, 91)

# 把多级编号挂到 Heading 1/2/3 样式
for ilvl, style in [(0, h1), (1, h2), (2, h3)]:
    pPr = style.element.get_or_add_pPr()
    numPr = OxmlElement("w:numPr")
    ilvl_el = OxmlElement("w:ilvl"); ilvl_el.set(qn("w:val"), str(ilvl))
    numId_el = OxmlElement("w:numId"); numId_el.set(qn("w:val"), "90")
    numPr.append(ilvl_el); numPr.append(numId_el)
    pPr.append(numPr)
    # 不缩进：编号后仅一个空格，正文紧跟其后
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "0"); ind.set(qn("w:firstLine"), "0")
    pPr.append(ind)

# 有序列表样式
if "List Paragraph" in [s.name for s in st]:
    lp = st["List Paragraph"]
else:
    lp = st.add_style("List Paragraph", 1)  # paragraph
style_font(lp, 12, False, "宋体")
style_para(lp, WD_ALIGN_PARAGRAPH.JUSTIFY, line=1.5, before=0, after=0, first_line_chars=0, left_chars=1)
pPr = lp.element.get_or_add_pPr()
numPr = OxmlElement("w:numPr")
ilvl_el = OxmlElement("w:ilvl"); ilvl_el.set(qn("w:val"), "0")
numId_el = OxmlElement("w:numId"); numId_el.set(qn("w:val"), "91")
numPr.append(ilvl_el); numPr.append(numId_el)
pPr.append(numPr)

# ---------- 文档属性 ----------
cp = doc.core_properties
cp.title = "大报告模板"
cp.subject = "PBL 大报告（三级标题 + GB/T 7714 顺序编码制）"
cp.author = "朱俊铭"
cp.comments = "模板：一级标题四号加粗，二级/三级小四；正文小四；图题“图x. 说明（内容）”；引用 GB/T 7714-2015 顺序编码制"

# ---------- 内容工具 ----------
def p(text="", style=None, bold=False, size=None, align=None, first_line=True):
    par = doc.add_paragraph(style=style)
    if text:
        r = par.add_run(text)
        r.bold = bold
        if size:
            r.font.size = Pt(size)
    if align is not None:
        par.alignment = align
    if not first_line and style is None:
        par.paragraph_format.first_line_indent = Pt(0)
    return par

def heading(text, level):
    return doc.add_heading(text, level=level)

def list_item(text):
    par = doc.add_paragraph(style="List Paragraph")
    par.add_run(text)
    return par

def field(paragraph, instr, placeholder=""):
    r = paragraph.add_run()
    fld1 = OxmlElement("w:fldChar"); fld1.set(qn("w:fldCharType"), "begin"); fld1.set(qn("w:dirty"), "true")
    r._r.append(fld1)
    r2 = paragraph.add_run()
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = " %s " % instr
    r2._r.append(it)
    r3 = paragraph.add_run()
    sep = OxmlElement("w:fldChar"); sep.set(qn("w:fldCharType"), "separate")
    r3._r.append(sep)
    if placeholder:
        paragraph.add_run(placeholder)
    r4 = paragraph.add_run()
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    r4._r.append(end)

def caption(kind, text):
    par = doc.add_paragraph(style="Caption")
    par.add_run(kind)
    field(par, "SEQ %s \\* ARABIC" % kind, "1")
    par.add_run(". " + text)
    return par

# ---------- 正文示例 ----------
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = title.add_run("×××课程 / PBL 大报告")
r.font.size = Pt(18); r.bold = True
r.font.name = "黑体"
r._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
title.paragraph_format.space_after = Pt(6)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sub.add_run("第 × 组　　执笔：×××　　成员：×××、×××、×××　　2026 年 10 月")
r.font.size = Pt(10.5)
sub.paragraph_format.first_line_indent = Pt(0)
sub.paragraph_format.space_after = Pt(12)

p("说明：一级标题写“每个人分配的问题”，二级、三级标题是其下的分析层次；第三级以下不再设标题，"
  "改用有序列表 (1)(2)(3)。正文一律小四；图题写作“图1. 说明文字（内容）”；参考文献采用 GB/T 7714—2015 顺序编码制。", first_line=True)

# 目录
h = doc.add_paragraph()
r = h.add_run("目　录")
r.bold = True; r.font.size = Pt(14)
r.font.name = "黑体"; r._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
h.alignment = WD_ALIGN_PARAGRAPH.CENTER
h.paragraph_format.first_line_indent = Pt(0)

toc_p = doc.add_paragraph()
toc_p.paragraph_format.first_line_indent = Pt(0)
toc_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_t = toc_p.add_run("（此处插入目录：引用 → 目录 → 自动目录；本模板标题已按 1. / 1.1. / 1.1.1. 自动编号）")
_t.font.size = Pt(10.5)
_t.italic = True

doc.add_paragraph()

# ---- 1. 问题一 ----
heading("问题一：×××（姓名）", 1)
p("本部分回答第 1 位同学分配的问题。先给出问题界定与结论提要，再分层次展开。范式：一句话结论 → 三条依据 → 一个限定条件。")

heading("研究背景与现状", 2)
p("在此写该问题的背景、概念界定、研究现状。正文小四，1.5 倍行距，首行缩进 2 字符。")

heading("国内研究进展", 3)
p("三级标题下直接写正文；若还需分层，用有序列表，不再增加标题层级：")
list_item("第一层要点：××××××××××××××××。")
list_item("第二层要点：××××××××××××××××。")
list_item("第三层要点：××××××××××××××××。")

heading("机制与证据", 3)
p("此处插入机制图，图题格式为“图1. 说明文字（内容）”，并在正文中先引用“见图1”。")
doc.add_picture(r"D:\desktop\大四上\PBL\tools\img\figure_placeholder.png", width=Cm(12.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.paragraphs[-1].paragraph_format.first_line_indent = Pt(0)
doc.paragraphs[-1].paragraph_format.space_before = Pt(6)
caption("图", "下丘脑—垂体—甲状腺轴（机制示意图）")
p("（把图片插入到本行上方；图题用“引用 → 插入题注”或直接套用 Caption 题注样式，编号随 SEQ 域自动更新。）")

heading("小结", 2)
p("小结本问题：给出可检验的判断与临床落点。")

# ---- 2. 问题二 ----
heading("问题二：×××（姓名）", 1)
p("第二问的问题陈述与结论提要。")

heading("概念与分类", 2)
heading("定义与边界", 3)
p("正文……")
heading("分类依据", 3)
p("正文……")

heading("关键证据", 2)
p("证据一：……[1]。证据二：……[2-3]。证据三：……[4]。")

# ---- 3. 问题三 ----
heading("问题三：×××（姓名）", 1)
p("第三问……")
heading("数据与指标", 2)
p("正文……（表格示例见下，表题写法与图题一致）")
caption("表", "本例主要实验室指标（CASE 指标）")

table = doc.add_table(rows=4, cols=4)
table.style = "Table Grid"
table.alignment = 1
data = [["指标", "检测值", "参考范围", "判读"],
        ["PRL", "30.63 ng/mL", "4.79–23.3", "轻度升高（1.3×）"],
        ["FT4", "60.47 pmol/L", "12.0–22.0", "明显升高（2.7×）"],
        ["TSH", "<0.01 μIU/mL", "0.27–4.2", "被抑制"]]
for i, row in enumerate(data):
    for j, val in enumerate(row):
        cell = table.cell(i, j)
        cell.text = val
        for par in cell.paragraphs:
            par.paragraph_format.first_line_indent = Pt(0)
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER if j else WD_ALIGN_PARAGRAPH.LEFT
            for run in par.runs:
                run.font.size = Pt(10.5)
                if i == 0:
                    run.bold = True

# ---- 4. 问题四（可复制） ----
heading("问题四：×××（姓名）", 1)
p("第四问……（需要更多问题就复制本段标题：选中“1. 问题一”整段 → 回车后会自动接续编号。）")

# ---- 参考文献 ----
heading("参考文献", 1)
p("以下按 GB/T 7714—2015 顺序编码制（numeric）著录；Word 中请在“引用 → 样式”选择 "
  "“Chinese Std GBT7714 (numeric)”，或使用随附的 Chinese Std GBT7714 (numeric).ens（EndNote）样式；"
  "正文引用用上标方括号，如 [1]、[2-3]。")

refs = [
    "[1] 作者. 题名[M]. 版次. 出版地: 出版社, 年: 起止页.",
    "[2] 作者, 作者. 题名[J]. 刊名, 年, 卷(期): 起止页.",
    "[3] 作者. 题名[D]. 城市: 学校, 年.",
    "[4] GRATTAN D R. 50 years of neuroendocrinology: prolactin and the prolactin receptor[J]. "
    "Journal of Endocrinology, 2015, 226(2): T101-T122.",
]
for ref in refs:
    par = doc.add_paragraph()
    par.paragraph_format.first_line_indent = Pt(0)
    par.paragraph_format.left_indent = Pt(0)
    par.paragraph_format.space_after = Pt(0)
    par.add_run(ref).font.size = Pt(10.5)

doc.add_paragraph()
p("附：本模板已设置——页边距 上下 2.54 cm、左右 1.91 cm；A4；正文宋体小四、1.5 倍行距、首行缩进 2 字符；"
  "一级标题黑体四号加粗（自动编号 1. 2. 3.），二、三级黑体小四加粗（1.1. / 1.1.1.）；"
  "第三级以下一律用 (1)(2)(3) 有序列表；图/表题注五号居中，格式“图1. 说明文字（内容）”。")

# 兼容模式 15（非“兼容性模式”打开）
settings = doc.settings.element
compat = settings.find(qn("w:compat"))
if compat is None:
    compat = OxmlElement("w:compat"); settings.append(compat)
cs = OxmlElement("w:compatSetting")
cs.set(qn("w:name"), "compatibilityMode")
cs.set(qn("w:uri"), "http://schemas.microsoft.com/office/word")
cs.set(qn("w:val"), "15")
compat.append(cs)

doc.save(OUT)
print("saved", OUT)
