---
name: pbl-report-tex
description: >
  把 PBL / 课程大报告从 Word 稿件整理成规范 LaTeX 并编译 PDF：支持多人稿件汇总、
  四章式结构（第一章 案例回顾 / 第二章 重大问题 / 第三章 次要问题 / 第四章 机制图）、
  标题层级归一化（1. / 1.1 / 1.1.1，三级封顶，更深用有序列表）、图表题注、
  GB/T 7714 顺序编码制参考文献（跨文件去重、正文引用自动重编号）。
  Use when the user asks to: 汇总/合并大家提交的 docx 成一份大报告, 转 LaTeX,
  统一排版/套模板, 生成四章式报告, 按问题名把内容分到“重大问题/次要问题”,
  或提到 PBL 大报告、多人合并、章节规划、docx2tex。
metadata:
  version: "1.0.0"
  language: zh-CN
---

# PBL 大报告 → LaTeX 技能

四章固定结构，章标题 **居中 / 三号(16pt) / 微软雅黑 / 加粗**；章下用 1.1、1.1.1 自动编号（三级封顶，再深用 `(1)(2)(3)` 有序列表）；图片题注 `图1. 说明（内容）`；参考文献 GB/T 7714—2015 顺序编码制。

## 一、判断用哪条路

| 用户诉求 | 用哪个脚本 | 产物 |
|---|---|---|
| 单份 Word → LaTeX | `scripts/docx2tex.py` | 一个 `.tex` + `figures/` |
| 多份稿件按原结构合并成一份（不做分章） | `scripts/merge_docx2tex.py` | 一份连续编号的 `.tex` |
| **四章式大报告**（推荐，本技能主流程） | `scripts/chapters_docx2tex.py` | 四章 `.tex` + `分章清单.txt` + PDF |
| 只要 Word 空模板 | `scripts/make_report_template4.py` | 四章式 `.docx` 模板 |

## 二、四章式主流程（Agent 按此执行）

```powershell
# 0) 依赖
pip install python-docx        # 需要 xelatex（MiKTeX / TeX Live）

# 1) 看每份稿件有哪些标题（用于判断重大问题 / 次要问题）
python scripts\dump_headings.py 提交目录

# 2) 写计划 JSON（chapter 顺序固定：案例回顾、重大问题、次要问题、机制图）
#    —— 按“问题名称”把文件放进对应章；机制图章的 figures 放图片路径
python scripts\chapters_docx2tex.py 计划.json -o 大报告.tex --compile
```

**计划 JSON 格式**

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

- `files` / `figures` 里的相对路径**相对计划 JSON 所在目录**；
- JSON 请存成 **UTF-8 无 BOM**（工具也兼容带 BOM）；
- 某章可以没有内容（空 `files`），空章不会输出。

**怎么判“重大 vs 次要”**（按问题名称）

- 重大问题：机制/因果/调控/通路/“为什么”“如何”“解释”“关系”“判读”类——决定了整个案例的解释框架；
- 次要问题：概念/定义/分类/鉴别/补充/细节/名词解释类——一两段讲清即可；
- 案例回顾：病例、病史、查体、指标这些“原始资料”，不展开分析；
- 机制图：全组画的图，配图题集中放第四章。

没有计划文件时也可以 `--auto`：按文件名关键词（机制/为什么/因果 → 重大问题；鉴别/概念/定义 → 次要问题）粗略分章，**结果要人工复核**。

## 三、脚本参数速查

```
docx2tex.py       input.docx [-o out.tex] [--margins 上,下,左,右] [--upcite]
merge_docx2tex.py 目录|文件... [-o out.tex] [--strip-lead] [--toc/--no-toc] [--compile]
chapters_docx2tex.py 计划.json [-o out.tex] [--auto] [--compile]
make_report_template4.py [输出.docx]
dump_headings.py  目录|文件...
```

## 四、识别规则（排错时对照）

| 元素 | 规则 |
|---|---|
| 标题 | 样式 `Heading 1-3` / `标题 1-3` → 大纲级别 0–2（**9 = 正文**，不算标题）→ 文本 `1.` `1.1.` `1.1.1.` 兜底 |
| 层级 | 四章模式：文件最浅一级 → `\section`（1.1），下一级 → `\subsection`（1.1.1），再深 → 加粗小标题；有序列表用 `(1)(2)(3)` |
| 图片 | 段落内嵌图 → `figures/`（按 `文件名_序号` 命名）+ `figure` 环境；紧随的“图x. ……”→ `\caption` |
| 表格 | Word 表 → booktabs 三线表；表题在上/下都能配对 |
| 文献 | “参考文献”之后以 `[n]` 开头的段落 → `thebibliography`；跨文件按内容去重，正文 `[n]` 自动映射到全局序号 |

## 五、坑与对策

1. **xelatex 要跑两遍**才有目录/交叉引用；`--compile` 已自动跑两遍。
2. 中文缺字（如 μ）：类文件用 `fontset=windows` + `\setmainfont{Times New Roman}`。
3. 章标题字体：`\setCJKfamilyfont{yh}{Microsoft YaHei}`，章格式 `\centering\yahei\zihao{3}\bfseries`（三号=16pt）。
4. 大纲级别全 9 的文档（PDF 转来的 docx）不会误判成标题。
5. 每份稿件开头的“个人标题/作者/日期”会变成正文：合并时用 `--strip-lead` 丢弃第一个标题之前的段落。
6. 图片路径、JSON 路径写错会打印 `[SKIP] 找不到文件`——按提示修正相对路径。
7. Word 那边要对齐同样的版式：用 `scripts/make_report_template4.py` 生成模板（章=微软雅黑三号加粗居中，节/小节小四加粗，正文小四，边距 2.54/2.54/1.91/1.91）。

## 六、交付前自检

- [ ] `分章清单.txt` 里四章的文件是否放对（重大问题/次要问题按问题名核对）
- [ ] PDF 目录页的章号为“第一章/第二章…”，节号 1.1 / 1.1.1
- [ ] 每张图下面都有“图x.”题注，正文有“见图x”引用
- [ ] 参考文献编号连续、正文 `[n]` 与文末一一对应
- [ ] 章节标题字体为微软雅黑加粗居中（不是宋体）
