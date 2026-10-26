# PBL 大报告生成器 · 桌面版（Windows / tkinter）

给不会用命令行的同学用的 **双击即用** 小软件：把 Word（.docx）稿件整理成规范 LaTeX 并编译 PDF。
与仓库里的 `scripts/` 共用同一套转换引擎（本目录的 `pbltex/`），区别只是多了图形界面。

> 界面截图、封面效果见仓库 `app/` 的构建产物；封面默认使用内置校徽 `assets/cover_logo.png`（西安交大）。

## 两种用法

| 形态 | 怎么用 |
|---|---|
| 免安装文件夹版（推荐） | 双击 `PBL大报告生成器.exe`，不需要装 Python，也不用往临时目录解包 |
| 单文件版（备用） | 只有一个 exe 便于传输，但每次启动要先在 `%TEMP%` 解包；受限电脑可能报 `Could not create temporary directory!` |

## 界面里有什么

- **① 单份转换**：添加一份/一批 Word → 每份生成 `.tex` + PDF（1./1.1./1.1.1. 标题、小四、1.5 倍行距、
  图表题注、GB/T 7714 文献、正文 `[1]` 自动上标）。
- **② 合并四章大报告**：添加所有人的稿件 → 自动分章（读题目关键词）→ 选中行点
  「→ 案例回顾 / 重大问题 / 次要问题 / 机制图」核对 → 填标题组别 → 生成
  封面 + 目录 + 四章 + 分章清单 + 参考文献跨文件去重。
- **③ 环境与说明**：检测 MiKTeX/TeX Live，没装就一键跳下载页；可保存日志。

## 打包

```powershell
pip install pyinstaller python-docx

# 单文件版
pyinstaller --noconfirm --clean --onefile --windowed --icon app.ico `
  --name "PBL大报告生成器" --add-data "app.ico;." --add-data "assets;assets" `
  --hidden-import pbltex.core --hidden-import pbltex.toolchain --paths "." app.py

# 免安装文件夹版（推荐分发）
pyinstaller --noconfirm --clean --onedir --windowed --icon app.ico `
  --name "PBL大报告生成器" --add-data "app.ico;." --add-data "assets;assets" `
  --hidden-import pbltex.core --hidden-import pbltex.toolchain --paths "." `
  --distpath dist_app --workpath build_app app.py
```

或直接运行 `build.ps1`（会自动做图标 → 打包两种形态 → 自检）。

## 命令行（同一个 exe）

```powershell
PBL大报告生成器.exe --cli-single 稿件.docx -o 输出目录
PBL大报告生成器.exe --cli-chapters 计划.json
PBL大报告生成器.exe --selftest          # 只检查界面能否启动
```

## 引擎要点（`pbltex/core.py`）

- **结构**：每份稿件 = 章内一个一级标题（`1. 题目`），稿内标题依次降为 `1.1` / `1.1.1`；
  章号按实际输出的顺序编，节号每章重新从 1. 开始；图/表编号全文连续。
- **表格宽度**：先估算内容宽度（中文 1 字、半角 0.55 字），
  装得下 → 普通 `tabular`；稍宽且有文字列 → `tabularx` 折行；列太多 → `\resizebox` 等比缩放，
  保证不超出正文宽度。合并单元格按 `gridSpan` 生成 `\multicolumn`。
- **自动分章**：题目里 `病史/查体/现病史/案例回顾` → 案例回顾；`机制/调节/原理/关系…` → 重大问题；
  `定义/鉴别/简介…` → 次要问题；其余默认重大问题。注意「结合案例分析××」不会被当成案例回顾。

## 封面与引用（v1.3）

- 封面照 23 级 Case 报告样式：校徽 + 2023级××组PBL学习 / 学习报告 / Case x：… + Tutor / 主席 / 书记 / 成员 / 日期。
  界面「② 合并四章」里直接填；logo 默认取打包进 exe 的 ssets/cover_logo.png。
- 引用不再走 natbib：正文 [1]、[2-3] 在转换阶段就算好，直接写成 \textsuperscript{[1,2]}，
  只编译一遍也不会出现 [?,?,?]（以前第二遍没跑完就会全部变成问号）。
- 编译：至少两遍，.log 仍提示 Rerun 时最多补到三遍；最后一遍失败会在日志里明确警告。