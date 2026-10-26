---
name: pbl-report-tex
description: >
  把 PBL / 课程大报告从 Word 稿件整理成规范 LaTeX 并编译 PDF：支持多人稿件汇总、
  四章式结构（第一章 案例回顾 / 第二章 重大问题 / 第三章 次要问题 / 第四章 机制图）、
  每份稿件 = 章内一个一级标题（1. 题目，稿内标题依次降为 1.1 / 1.1.1）、图表题注、
  表格宽度自适应（不溢出）、GB/T 7714 顺序编码制参考文献（跨文件去重、正文引用自动重编号）。
  仓库另附 Windows 桌面版（双击即用，给不会用命令行的同学）。
  Use when the user asks to: 汇总/合并大家提交的 docx 成一份大报告, 转 LaTeX,
  统一排版/套模板, 生成四章式报告, 按问题名把内容分到“重大问题/次要问题”,
  或提到 PBL 大报告、多人合并、章节规划、docx2tex、论文/报告排版工具。
metadata:
  version: "2.0.0"
  language: zh-CN
---

# PBL 大报告 → LaTeX 技能（v2）

四章固定结构：**案例回顾 / 重大问题 / 次要问题 / 机制图**。
章标题 **居中 / 三号(16pt) / 微软雅黑 / 加粗**；
**每份稿件 = 章内一个一级标题**（`1. 题目`），稿内标题依次降级（`1.1`、`1.1.1`，三级封顶，再深用 `(1)(2)(3)`）；
图题 `图1. 说明（内容）`；参考文献 GB/T 7714—2015 顺序编码制。

## 一、判断用哪条路

| 用户诉求 | 用什么 | 产物 |
|---|---|---|
| 单份 Word → LaTeX | `scripts/docx2tex.py` | 一个 `.tex` + `figures/` |
| 多份稿件按原结构合并成一份（不做分章） | `scripts/merge_docx2tex.py` | 一份连续编号的 `.tex` |
| **四章式大报告**（推荐，本技能主流程） | `scripts/chapters_docx2tex.py` | 四章 `.tex` + `分章清单.txt` + PDF |
| 同学不会用命令行，要“双击即用” | `app/` 里的桌面版（见第六节） | 免安装文件夹版 exe |
| 只要 Word 空模板 | `scripts/make_report_template4.py` | 四章式 `.docx` 模板 |

## 二、四章式主流程（Agent 按此执行）

```powershell
# 0) 依赖
pip install python-docx        # 需要 xelatex（MiKTeX / TeX Live）

# 1) 看每份稿件有哪些标题（判断重大问题 / 次要问题，并确认题目行）
python scripts\dump_headings.py 提交目录

# 2) 写计划 JSON（章顺序固定：案例回顾、重大问题、次要问题、机制图）
python scripts\chapters_docx2tex.py 计划.json -o 大报告.tex --compile

# 没有计划文件时：按题目关键词自动分章（结果要人工复核）
python scripts\chapters_docx2tex.py 提交目录 -o 大报告.tex --auto --compile
```

**计划 JSON**

```json
{
  "title": "内分泌系统 / PBL 大报告",
  "author": "第 × 组",
  "date": "2026 年 10 月",
  "chapters": [
    {"title": "案例回顾", "files": ["提交/01_病例.docx"]},
    {"title": "重大问题", "files": ["提交/02_机制.docx", "提交/03_因果.docx"]},
    {"title": "次要问题", "files": ["提交/04_鉴别.docx"]},
    {"title": "机制图", "figures": [
      {"path": "pics/mech1.png", "caption": "甲亢与应激抬高 PRL 的三条链（机制示意）"}
    ]}
  ]
}
```

- `files` / `figures` 相对**计划 JSON 所在目录**；JSON 存 **UTF-8 无 BOM**（也兼容带 BOM）；
- 空章（`files` 和 `figures` 都空）不会输出，章号按实际输出的顺序编；
- `chapters[].title` 可自定义章名，用来覆盖默认的“案例回顾/重大问题/…”。

## 三、结构规则（v2 的核心，别搞错）

```
第二章  重大问题                                  ← \chapter（第X章，雅黑三号加粗居中）
  1.  结合案例分析泌乳素分泌的调节因素              ← 每份稿件的题目 → \section（1.）
      1.1  病例特点                               ← 稿件里的 1 级标题 → \subsection
          1.1.1  实验室指标                       ← 稿件里的 2 级标题 → \subsubsection
      1.2  小结
  2.  下一个人的题目
      2.1  …
```

- **一级标题的来源**：稿件开头第一行（若不是“姓名+学号”）→ 稿件第一个标题 → 文件名；
- 题目行与稿件第一个标题是同一句时**只保留一个**；
- 编号：`\thesection=\arabic{section}.`、`\thesubsection=\thesection\arabic{subsection}.`，
  即**每章内从 1. 重新开始**；图/表编号全文连续（`\thefigure=\arabic{figure}`）；
- 三级封顶：第 4 级及以下降为加粗小标题 / `(1)(2)(3)` 有序列表。

**怎么判“重大 vs 次要”**（看题目首行 + 文件名）

- 案例回顾：`案例回顾 / 病史 / 现病史 / 查体 / 病例摘要 / 主诉`——**“案例”二字不算**（“结合案例分析××”很常见，别误判）；
- 重大问题：`机制 / 因果 / 通路 / 调控 / 调节 / 原理 / 解释 / 关系 / 判读 / 影响 / 作用 / 因素 / 如何`；
- 次要问题：`定义 / 概念 / 分类 / 鉴别 / 简介 / 概述 / 区别 / 细节 / 补充`；
- 机制图：全组画的图，放 `figures`；
- 判不出来默认**重大问题**；`--auto` 只是初猜，**必须人工核对**一遍。

## 四、脚本参数速查

```
docx2tex.py           input.docx [-o out.tex] [--margins 上,下,左,右] [--upcite]
merge_docx2tex.py     目录|文件... [-o out.tex] [--strip-lead] [--toc/--no-toc] [--compile]
chapters_docx2tex.py  计划.json|目录 [-o out.tex] [--auto] [--compile] [--title/--author/--date]
make_report_template4.py [输出.docx]
dump_headings.py      目录|文件...
```

`chapters_docx2tex.py` 自 v2 起把转换交给 `app/pbltex/core.py`（与桌面版同一套引擎）。

## 五、识别规则（排错时对照）

| 元素 | 规则 |
|---|---|
| 标题 | 样式 `Heading 1-3` / `标题 1-3` → 大纲级别 0–8（**9 = 正文**）→ 文本编号 `1.` `1.1` `1.2核心…`（无空格也认）→ 全加粗短行 |
| 一级标题 | 稿件题目行（跳过“姓名+学号”行）/ 第一个标题 / 文件名 |
| 层级 | 稿件内 1 级 → `\subsection`(1.1)，2 级 → `\subsubsection`(1.1.1)，3 级及以下 → 加粗小标题 |
| 列表 | 有效 `numId`（≠0）或 `(1)` 开头；连续「标签：内容」段落自动合成 `(1)(2)(3)` 并加粗标签 |
| 图片 | 段落内嵌图 → `figures/`（`文件名_序号`）+ `figure` 环境；紧随的“图x. ……”→ `\caption` |
| 表格 | Word 表 → booktabs 三线表；**宽度自适应**（见下）；合并单元格按 `gridSpan` → `\multicolumn` |
| 文献 | “参考文献”后以 `[n]` 开头的段落 → `thebibliography`；正文 `[n]` → `\cite` 上标；跨文件去重后重编号 |

**表格宽度三档**（保证不溢出正文）

1. 估算宽度 ≤ 版心 → 普通 `tabular`（不拉伸不缩放）；
2. 有文字列且稍宽 → `tabularx` + `X` 列折行，宽度 = `\linewidth`；
3. 列太多 / 太宽 → `\resizebox{\linewidth}{!}{...}` 等比缩到版心。

宽度估算：中文/全角算 1 字、半角算 0.55 字，A4 左右边距 1.91 cm 时版心约 40 字宽。

## 六、桌面版（点击即用，给同学用）

`app/` 是一个 tkinter 图形界面，两种模式（单份转换 / 合并四章），打包成单文件 exe 或免安装文件夹。

```powershell
cd app
pip install pyinstaller python-docx
.\build.ps1                 # 生成图标 → 打包 onefile + onedir → 自检
```

分发建议：**优先给“免安装文件夹版”**（不需要往临时目录解包）；
单文件版在受限电脑上可能报 `Could not create temporary directory!`，那是解包权限问题，不是要管理员权限。

## 七、坑与对策

1. **xelatex 要跑两遍**才有目录页码；`--compile` 已自动跑两遍。
2. 中文缺字（μ 等）：`fontset=windows` + `\setmainfont{Times New Roman}`。
3. 黑体没有粗体字重：`\setCJKfamilyfont{heibold}{SimHei}[AutoFakeBold=2]`，否则 `\bfseries` 回退。
4. 图题要 `图1.`（无空格）：`\DeclareCaptionLabelFormat{...}{\CJKsetecglue{}#1#2}`，
   不处理的话 xeCJK 会在“图”和“1”之间插一条胶。
5. 章标题字体用 `\setCJKfamilyfont{pblyh}{Microsoft YaHei}` —— 不要自己定义 `\yahei`，ctex 已占用。
6. 大纲级别全 9 的文档（PDF 转来的 docx）不会误判成标题。
7. 表格永远用上面的三档策略，别直接 `tabular{lccc}`，长中文会冲出版心。
8. 图片路径、JSON 路径写错会打印 `[SKIP] 找不到文件` —— 按提示修正相对路径。
9. Word 那边要对齐版式：`scripts/make_report_template4.py` 生成模板
   （章=微软雅黑三号加粗居中，节/小节小四加粗，正文小四，边距 2.54/2.54/1.91/1.91）。

## 八、交付前自检

- [ ] `分章清单.txt` 里每份稿件落在正确的章，“→ 一级标题”与题目一致
- [ ] PDF 目录页：`第二章 重大问题` → `1. 题目` → `1.1` → `1.1.1`（章内编号从 1. 开始）
- [ ] 每张图下面都有“图x.”题注，正文有“见图x”引用；图/表编号连续
- [ ] 每张表都没超出正文宽度（渲染成图看一眼最保险）
- [ ] 参考文献编号连续、正文 `[n]` 与文末一一对应
- [ ] 章标题为微软雅黑加粗居中（不是宋体）
