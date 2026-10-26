# pbl-report-tex · PBL 大报告 → LaTeX 工具链

把 Word 版大报告（尤其是**多人分别撰写、需要合并**的 PBL 报告）整理成规范 LaTeX 并编译 PDF。

三种形态，同一套引擎：

| 形态 | 适合谁 | 入口 |
|---|---|---|
| **桌面版（点击即用）** | 不会用命令行的同学 | `app/` → 免安装文件夹版 exe |
| **Skill** | DeepSeek Harness 里的 Agent | `SKILL.md` |
| **命令行脚本** | 会用终端的人 / 写脚本批量跑 | `scripts/` |

## 四章固定结构

| 章 | 内容 |
|---|---|
| **第一章　案例回顾** | 病例、病史、查体、CASE 指标（只整理原始资料） |
| **第二章　重大问题** | 机制 / 因果 / 调控 / “为什么”“如何”类问题 |
| **第三章　次要问题** | 概念、定义、分类、鉴别、细节补充 |
| **第四章　机制图** | 全组机制图集中展示，逐张配题注 |

**结构（v2）**：每份稿件 = 章内一个一级标题，稿内标题依次降级，编号每章从 1. 重来：

```
第二章  重大问题                               ← 章：居中 · 三号 · 微软雅黑 · 加粗
  1.  结合案例分析泌乳素分泌的调节因素           ← 稿件题目
      1.1  病例特点
          1.1.1  实验室指标
      1.2  小结
  2.  下一个人的题目
```

正文小四、1.5 倍行距、首行缩进 2 字符；图题 `图1. 说明（内容）`、表题在表上方；
**表格宽度自适应**（自然宽度 / `tabularx` 折行 / `\resizebox` 缩放，保证不超出正文）；
参考文献 GB/T 7714—2015 顺序编码制，跨文件去重、正文引用自动重编号。
**封面**照 23 级 Case 报告样式：校徽 + `2023级××组PBL学习`（一号）/ `学习报告`（小一）/ `Case：名称`（小初），
标题行距 3 倍；下面空一行，再是 Tutor / 主席 / 书记 / 成员（三号、行距 1.5）。
在 GUI 或计划 JSON 的 `cover` 里填；`Case` 可只填名称，程序自动补前缀；**引用直接写上标编号**，不依赖两遍编译，不会出现 `[?,?,?]`。

## 目录结构

```
SKILL.md                      Skill 入口：触发条件 / 主流程 / 规则 / 踩坑
app/                          桌面版（tkinter，双击即用）
  app.py                        图形界面 + 命令行入口
  pbltex/core.py                Word→LaTeX 引擎（结构 / 表格 / 文献 / 四章合并）
  pbltex/toolchain.py           查找 xelatex、编译 PDF
  build.ps1                     一键打包（onefile + onedir）
scripts/
  docx2tex.py                 单份 Word → LaTeX
  merge_docx2tex.py           多份稿件合并（连续编号）
  chapters_docx2tex.py        四章式组装（主流程，v2 起调用 app/pbltex）
  dump_headings.py            先探每份稿件的标题结构
  make_report_template4.py    生成四章式 Word 模板
  make_report_template.py     生成三级标题式 Word 模板
templates/                    两份 Word 模板
assets/figure_placeholder.png 占位图
examples/                     计划 JSON + 编译成品 PDF
references/使用说明.md
```

## 快速开始（命令行）

```bash
pip install python-docx          # 另需 xelatex（MiKTeX / TeX Live）

# 1) 先看每份稿件有哪些标题
python scripts/dump_headings.py 提交目录

# 2) 按“问题名称”把文件写进四章计划
python scripts/chapters_docx2tex.py examples/计划.json -o 大报告.tex --compile

# 或者：按题目关键词自动分章（结果要人工复核）
python scripts/chapters_docx2tex.py 提交目录 -o 大报告.tex --auto --compile
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
  ],
  "cover": {"line1": "2023级11组PBL学习", "line2": "学习报告", "case": "Case 1：难缠的脖子痛",
            "tutor": "马凯歌", "chair": "朱俊铭", "secretary": "郭育恒",
            "members": ["叶瑞林", "张欣仪", "韩曦", "任欣萌", "王思源", "于子轩", "牛海华", "韩畅"],
            "logo": "assets/cover_logo.png"}
}
```

## 桌面版（给同学用）

`app/` 里是 tkinter 界面：① 单份转换 ② 合并四章大报告（自动分章 + 手动分配）③ 环境检测。

```powershell
cd app
pip install pyinstaller python-docx
.\build.ps1            # 出 onefile 单文件版 + onedir 免安装文件夹版，并自检
```

分发说明：**优先发免安装文件夹版**（启动时不需要往 `%TEMP%` 解包，最稳）。
单文件版在受限电脑上可能报 `Could not create temporary directory!` —— 那是 PyInstaller 解包权限问题，
不是程序申请了管理员权限（清单是 `asInvoker`）。

## 脚本参数

```
docx2tex.py           input.docx [-o out.tex] [--margins 上,下,左,右] [--upcite]
merge_docx2tex.py     目录|文件... [-o out.tex] [--strip-lead] [--toc/--no-toc] [--compile]
chapters_docx2tex.py  计划.json|目录 [-o out.tex] [--auto] [--compile] [--title/--author/--date]
make_report_template4.py [输出.docx]
dump_headings.py      目录|文件...
```

## 识别规则

| 元素 | 规则 |
|---|---|
| 标题 | 样式 `Heading 1-3` / `标题 1-3` → 大纲级别 0–8（**9 = 正文**）→ 文本 `1.` `1.1.` `1.2核心…` → 全加粗短行 |
| 一级标题（大报告里） | 稿件题目行（跳过“姓名+学号”行）→ 第一个标题 → 文件名 |
| 图片 | 段落内嵌图 → `figures/` + `figure` 环境；紧随的“图x. ……”→ `\caption` |
| 表格 | Word 表 → booktabs 三线表；宽度三档自适应；合并单元格 → `\multicolumn` |
| 文献 | “参考文献”后 `[n]` 开头段落 → `thebibliography`；跨文件去重 + 正文引用自动重编号 |

## 作为 Skill 使用

把整个仓库放到 `~/.dsh/skills/pbl-report-tex/`（或软链接）。`chapters_docx2tex.py` 会从
`../app/pbltex` 导入引擎，所以 `app/` 要跟着一起放（仓库原样拷贝即可）。
Agent 在遇到“汇总大报告 / 转 LaTeX / 四章式 / 按问题名分章”等请求时会自动加载 `SKILL.md`。

## 已知坑

- xelatex 要跑两遍才有目录页码（`--compile` 已处理）；
- 中文缺字（μ 等）需 `fontset=windows` + `\setmainfont{Times New Roman}`；
- 黑体没有粗体字重 → `AutoFakeBold`；图题要 `图1.` 无空格 → `\CJKsetecglue{}`；
- 章标题字体用自定义族名（`pblyh`），别定义 `\yahei`（ctex 已占用）；
- PDF 转来的 docx 常给每个段落设 `outlineLvl=9`（正文），工具已排除；
- 表格别再用裸 `tabular{lccc}`，长中文一定会冲出版心。

## License

MIT
