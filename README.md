# pbl-report-tex · PBL 大报告 → LaTeX 工具链

把 Word 版大报告（尤其是**多人分别撰写、需要合并**的 PBL 报告）整理成规范 LaTeX 并编译 PDF。
既是 [DeepSeek Harness](https://github.com/) 的 **Skill**（`pbl-report-tex`），也是一套可单独使用的命令行脚本。

## 四章固定结构

| 章 | 内容 |
|---|---|
| **第一章　案例回顾** | 病例、病史、查体、CASE 指标（只整理原始资料） |
| **第二章　重大问题** | 机制 / 因果 / 调控 / “为什么”“如何”类问题 |
| **第三章　次要问题** | 概念、定义、分类、鉴别、细节补充 |
| **第四章　机制图** | 全组机制图集中展示，逐张配题注 |

**版式**：章标题 **居中 · 三号(16 pt) · 微软雅黑 · 加粗**；章下 `1.1`、`1.1.1` 自动编号（三级封顶，再深用 `(1)(2)(3)`）；正文小四、1.5 倍行距、首行缩进 2 字符；图题 `图1. 说明（内容）`；参考文献 GB/T 7714—2015 顺序编码制。

## 目录结构

```
SKILL.md                      Skill 入口：触发条件 / 主流程 / 规则 / 踩坑
scripts/
  docx2tex.py                 单份 Word → LaTeX
  merge_docx2tex.py           多份稿件合并（连续编号）
  chapters_docx2tex.py        四章式组装（主流程）
  dump_headings.py            先探每份稿件的标题结构
  make_report_template4.py    生成四章式 Word 模板
  make_report_template.py     生成三级标题式 Word 模板
templates/
  大报告模板-四章.docx          章=微软雅黑三号加粗居中
  大报告模板-三级版.docx        1. / 1.1 / 1.1.1 三级式
assets/figure_placeholder.png 占位图
examples/
  计划.json                    四章计划示例
  四章式演示.pdf                实际编译出的成品（16 页）
references/使用说明.md
```

## 快速开始

```bash
pip install python-docx          # 另需 xelatex（MiKTeX / TeX Live）

# 1) 先看每份稿件有哪些标题
python scripts/dump_headings.py 提交目录

# 2) 按“问题名称”把文件写进四章计划
python scripts/chapters_docx2tex.py examples/计划.json -o 大报告.tex --compile
```

计划 JSON：

```json
{
  "title": "内分泌系统 / PBL 大报告",
  "author": "第 × 组",
  "date": "2026 年 10 月",
  "chapters": [
    {"title": "案例回顾", "files": ["提交/01_病例.docx"]},
    {"title": "重大问题", "files": ["提交/02_机制.docx", "提交/03_因果.docx"]},
    {"title": "次要问题", "files": ["提交/04_鉴别.docx"]},
    {"title": "机制图", "figures": [{"path": "pics/mech1.png", "caption": "三条链机制全景"}]}
  ]
}
```

## 脚本参数

```
docx2tex.py           input.docx [-o out.tex] [--margins 上,下,左,右] [--upcite]
merge_docx2tex.py     目录|文件... [-o out.tex] [--strip-lead] [--toc/--no-toc] [--compile]
chapters_docx2tex.py  计划.json [-o out.tex] [--auto] [--compile]
make_report_template4.py [输出.docx]
dump_headings.py      目录|文件...
```

## 识别规则

| 元素 | 规则 |
|---|---|
| 标题 | 样式 `Heading 1-3` / `标题 1-3` → 大纲级别 0–2（**9 = 正文**）→ 文本 `1.` `1.1.` `1.1.1.` 兜底 |
| 图片 | 段落内嵌图 → `figures/` + `figure` 环境；紧随的“图x. ……”→ `\caption` |
| 表格 | Word 表 → booktabs 三线表 |
| 文献 | “参考文献”后 `[n]` 开头段落 → `thebibliography`；跨文件去重 + 正文引用自动重编号 |

## 作为 Skill 使用

把整个仓库放到 `~/.dsh/skills/pbl-report-tex/`（或软链接），Agent 在遇到
“汇总大报告 / 转 LaTeX / 四章式 / 按问题名分章”等请求时会自动加载 `SKILL.md`。

## 已知坑

- xelatex 要跑两遍才有目录页码（`--compile` 已处理）；
- 中文缺字（μ 等）需 `fontset=windows` + `\setmainfont{Times New Roman}`；
- 章标题字体用 `\setCJKfamilyfont{yh}{Microsoft YaHei}`；
- PDF 转来的 docx 常给每个段落设 `outlineLvl=9`（正文），工具已排除。

## License

MIT
