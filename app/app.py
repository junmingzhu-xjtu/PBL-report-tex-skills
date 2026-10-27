# -*- coding: utf-8 -*-
"""PBL 大报告生成器 —— 图形界面（双击即用）

三种运行方式：
    PBL大报告生成器.exe                 图形界面
    PBL大报告生成器.exe --cli-single a.docx -o 输出目录
    PBL大报告生成器.exe --cli-chapters plan.json
    PBL大报告生成器.exe --selftest      只检查界面能否启动
"""
from __future__ import annotations

import json
import os
import queue
import re
import sys
import tempfile
import threading
import time
import traceback

APP_NAME = "PBL 大报告生成器"
APP_VERSION = "1.6.0"
LOG_FILE = os.path.join(tempfile.gettempdir(), "pbl_report_gui.log")


# --------------------------------------------------------------------------- #
# 路径工具（兼容 PyInstaller 单文件）
# --------------------------------------------------------------------------- #
def res_path(name: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def exe_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _shell_folder(csidl: int) -> str:
    """取 Windows 的“桌面 / 文档”真实路径（中文系统叫“桌面”也能拿到）。"""
    if os.name != "nt":
        return ""
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(260)
        ctypes.windll.shell32.SHGetFolderPathW(None, csidl, None, 0, buf)
        return buf.value or ""
    except Exception:  # noqa: BLE001
        return ""


def default_outdir() -> str:
    """默认输出到桌面：绝不写进程序文件夹（那里常常没有写权限）。"""
    cands = []
    for csidl in (0x0010, 0x0005):            # 桌面 / 我的文档
        p = _shell_folder(csidl)
        if p:
            cands.append(os.path.join(p, "PBL生成结果"))
    cands.append(os.path.join(os.path.expanduser("~"), "Desktop", "PBL生成结果"))
    cands.append(os.path.join(tempfile.gettempdir(), "PBL生成结果"))
    for d in cands:
        try:
            os.makedirs(d, exist_ok=True)
            return d
        except OSError:
            continue
    return tempfile.mkdtemp(prefix="PBL生成结果_")


def default_cover_logo() -> str:
    """内置校徽（打包后从 _MEIPASS/assets 取）。"""
    for cand in (res_path(os.path.join("assets", "cover_logo.png")),
                 os.path.join(exe_dir(), "assets", "cover_logo.png")):
        if os.path.isfile(cand):
            return cand
    return ""


# --------------------------------------------------------------------------- #
# 命令行模式（供自动化/自检使用；界面模式下不进入）
# --------------------------------------------------------------------------- #
def _cli_out(msg: str):
    try:
        if sys.stdout is not None:
            sys.stdout.write(msg + "\n")
            sys.stdout.flush()
    except Exception:  # noqa: BLE001
        pass
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            fh.write(msg + "\n")
    except OSError:
        pass


def run_cli(argv) -> int:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from pbltex.core import convert_chapters, convert_single

    if argv[0] == "--selftest":
        try:
            import tkinter as tk
            root = tk.Tk()
            root.withdraw()
            root.update()
            root.destroy()
            _cli_out("SELFTEST OK tkinter")
            return 0
        except Exception as exc:  # noqa: BLE001
            _cli_out("SELFTEST FAIL %s" % exc)
            return 2

    if argv[0] == "--cli-single":
        src = argv[1]
        out = default_outdir()
        pdf = True
        if "-o" in argv:
            out = argv[argv.index("-o") + 1]
        if "--no-pdf" in argv:
            pdf = False
        try:
            info = convert_single(src, out, compile_pdf=pdf, log=_cli_out,
                                  layouts=(pdf and "--no-layout" not in argv))
        except Exception as exc:  # noqa: BLE001
            _cli_out("ERROR %s" % exc)
            _cli_out(traceback.format_exc())
            return 2
        _cli_out("DONE tex=%s pdf=%s layout_pdf=%s"
                 % (info["tex"], info["pdf"], info.get("layout_pdf", "")))
        return 0 if (not pdf or info["pdf"]) else 1

    if argv[0] == "--cli-chapters":
        try:
            with open(argv[1], encoding="utf-8-sig") as fh:
                plan = json.load(fh)
        except Exception as exc:  # noqa: BLE001
            _cli_out("ERROR 读不了计划文件：%s" % exc)
            return 2
        base = os.path.dirname(os.path.abspath(argv[1]))
        titles, figures = {}, {}
        if plan.get("items"):
            items = [(os.path.join(base, p) if not os.path.isabs(p) else p, int(c))
                     for p, c in plan["items"]]
        elif plan.get("chapters"):
            # 兼容技能脚本 chapters_docx2tex.py 的计划格式
            from pbltex.core import CHAPTER_NAMES
            items = []
            for idx, ch in enumerate(plan["chapters"]):
                title = ch.get("title")
                if title and (idx >= len(CHAPTER_NAMES) or title != CHAPTER_NAMES[idx]):
                    titles[idx] = title
                for f in ch.get("files") or []:
                    items.append((f if os.path.isabs(f) else os.path.join(base, f), idx))
                figs = []
                for fg in ch.get("figures") or []:
                    p = fg.get("path") if isinstance(fg, dict) else fg
                    cap = fg.get("caption", "") if isinstance(fg, dict) else ""
                    if p:
                        figs.append((p if os.path.isabs(p) else os.path.join(base, p), cap))
                if figs:
                    figures[idx] = figs
        else:
            _cli_out("ERROR 计划 JSON 里既没有 items 也没有 chapters")
            return 2
        outdir = plan.get("outdir") or default_outdir()
        try:
            info = convert_chapters(items, outdir, title=plan.get("title", "PBL 大报告"),
                                    author=plan.get("author", ""), date=plan.get("date", ""),
                                    compile_pdf=bool(plan.get("pdf", True)), log=_cli_out,
                                    chapter_titles=titles or None,
                                    chapter_figures=figures or None,
                                    cover=plan.get("cover"),
                                    section_titles=plan.get("section_titles"))
        except Exception as exc:  # noqa: BLE001
            _cli_out("ERROR %s" % exc)
            _cli_out(traceback.format_exc())
            return 2
        _cli_out("DONE tex=%s pdf=%s manifest=%s" % (info["tex"], info["pdf"], info["manifest"]))
        return 0

    _cli_out("UNKNOWN ARG %s" % argv[0])
    return 2


# --------------------------------------------------------------------------- #
# 图形界面
# --------------------------------------------------------------------------- #
def launch_gui(preload=None) -> int:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import pbltex
    from pbltex.core import CHAPTER_NAMES, convert_chapters, convert_single, guess_chapter

    try:  # 高分屏不糊
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001
        pass

    ACCENT = "#1f6feb"
    ACCENT_DARK = "#1158c7"
    MUTED = "#5b6470"

    class ScrollFrame(ttk.Frame):
        """可滚动的标签页：屏幕小的时候内容能滚动，按钮不会被挤出可视区。"""

        def __init__(self, master):
            super().__init__(master)
            self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, bg="#f6f7f9")
            vs = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
            self.canvas.configure(yscrollcommand=vs.set)
            vs.pack(side="right", fill="y")
            self.canvas.pack(side="left", fill="both", expand=True)
            self.inner = ttk.Frame(self.canvas, padding=10)
            self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
            self.inner.bind("<Configure>", self._on_inner)
            self.canvas.bind("<Configure>", self._on_canvas)
            self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._wheel))
            self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

        def _on_inner(self, _e):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def _on_canvas(self, event):
            self.canvas.itemconfigure(self._win, width=event.width)

        def _wheel(self, event):
            self.canvas.yview_scroll(int(-event.delta / 120), "units")

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("%s  v%s" % (APP_NAME, APP_VERSION))
            self.geometry("1180x900")
            self.minsize(1040, 700)
            self.q = queue.Queue()
            self.busy = False
            self.single_files = []
            self.chapter_items = []      # [(path, chapter_index, 一级标题覆盖)]
            self.entries = {}
            self.xelatex_ok = None
            self._title_cache = {}
            self._init_style()
            self._build()
            self._poll()
            self.after(150, self.check_env)
            self.after(80, self._fit_screen)
            self.protocol("WM_DELETE_WINDOW", self.on_close)
            try:
                self.iconbitmap(res_path("app.ico"))
            except Exception:  # noqa: BLE001
                pass

        # ---------------- 样式 ---------------- #
        def _init_style(self):
            st = ttk.Style(self)
            try:
                st.theme_use("clam")
            except tk.TclError:
                pass
            base = ("Microsoft YaHei UI", 10)
            st.configure(".", font=base, background="#f6f7f9")
            st.configure("TFrame", background="#f6f7f9")
            st.configure("Card.TFrame", background="white", relief="flat")
            st.configure("TLabel", background="#f6f7f9")
            st.configure("Card.TLabel", background="white")
            st.configure("Title.TLabel", font=("Microsoft YaHei UI", 15, "bold"),
                         background="#f6f7f9", foreground="#16202c")
            st.configure("Sub.TLabel", font=("Microsoft YaHei UI", 9),
                         background="#f6f7f9", foreground=MUTED)
            st.configure("H.TLabelframe", background="#f6f7f9")
            st.configure("TLabelframe", background="#f6f7f9", relief="solid", borderwidth=1)
            st.configure("TLabelframe.Label", font=("Microsoft YaHei UI", 10, "bold"),
                         background="#f6f7f9", foreground="#22303f")
            st.configure("TButton", padding=(10, 6), font=("Microsoft YaHei UI", 9))
            st.configure("Accent.TButton", padding=(16, 9), font=("Microsoft YaHei UI", 10, "bold"),
                         background=ACCENT, foreground="white", borderwidth=0)
            st.map("Accent.TButton",
                   background=[("active", ACCENT_DARK), ("disabled", "#a8c0e4")],
                   foreground=[("disabled", "#eef2f8")])
            st.configure("Toolbutton", font=("Microsoft YaHei UI", 9), padding=(12, 5), background="#e7eaef", relief="flat")
            st.map("Toolbutton",
                   background=[("selected", ACCENT_DARK), ("active", "#dbe4f3")],
                   foreground=[("selected", "white")])
            st.configure("Treeview", font=("Microsoft YaHei UI", 9), rowheight=26,
                         background="white", fieldbackground="white")
            st.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"))
            st.configure("TNotebook", background="#f6f7f9", borderwidth=0)
            st.configure("TNotebook.Tab", padding=(18, 9), font=("Microsoft YaHei UI", 10))
            st.map("TNotebook.Tab",
                   background=[("selected", "white"), ("!selected", "#e7eaef")],
                   foreground=[("selected", ACCENT_DARK)])
            st.configure("TEntry", padding=4)
            st.configure("Horizontal.TProgressbar", background=ACCENT, troughcolor="#e7eaef")

        # ---------------- 布局 ---------------- #
        def _build(self):
            top = ttk.Frame(self)
            top.pack(side="top", fill="x", padx=16, pady=(12, 6))
            ttk.Label(top, text=APP_NAME, style="Title.TLabel").pack(side="left")
            ttk.Label(top, text="Word / PDF 稿件 → 规范 LaTeX / PDF　·　四章式大报告一键合成",
                      style="Sub.TLabel").pack(side="left", padx=(12, 0), pady=(6, 0))
            self.env_badge = tk.Label(top, text="正在检测 LaTeX 环境…", font=("Microsoft YaHei UI", 9),
                                      fg=MUTED, bg="#f6f7f9")
            self.env_badge.pack(side="right")

            # 先摆底部（状态栏 + 日志），再让中间的标签页吃掉剩余空间：
            # 这样窗口被拉小/在高分屏上，标签页会自己收缩，日志不会被挤没。
            bottom = ttk.Frame(self)
            bottom.pack(side="bottom", fill="x", padx=16, pady=(0, 10))
            self.status = ttk.Label(bottom, text="就绪", style="Sub.TLabel")
            self.status.pack(side="left")
            self.progress = ttk.Progressbar(bottom, mode="determinate", length=220)
            self.progress.pack(side="right")
            # 合并模式的“生成大报告”按钮放在状态栏：切到 ② 才出现，始终看得见
            self.btn_ch = ttk.Button(bottom, text="生成大报告", style="Accent.TButton",
                                     command=self.do_chapters)

            logf = ttk.LabelFrame(self, text="运行日志")
            logf.pack(side="bottom", fill="both", padx=16, pady=(0, 6))
            self.logbox = tk.Text(logf, height=4, wrap="word", font=("Consolas", 9),
                                  bg="#0f1720", fg="#d7e2ee", insertbackground="#d7e2ee",
                                  relief="flat", padx=8, pady=6)
            sb = ttk.Scrollbar(logf, command=self.logbox.yview)
            self.logbox.configure(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")
            self.logbox.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
            self.logbox.tag_configure("ok", foreground="#7ee2a8")
            self.logbox.tag_configure("warn", foreground="#ffd479")
            self.logbox.tag_configure("err", foreground="#ff8a80")

            self.nb = ttk.Notebook(self)
            self.nb.pack(side="top", fill="both", expand=True, padx=16, pady=(4, 6))
            self._tab_single()
            self._tab_chapters()
            self._tab_env()
            self.nb.bind("<<NotebookTabChanged>>", self._on_tab_change)

        def _tab_single(self):
            sf = ScrollFrame(self.nb)
            self.nb.add(sf, text="① 单份转换")
            f = sf.inner

            box = ttk.LabelFrame(f, text="待转换的 Word 稿件")
            box.pack(fill="both", expand=True)
            cols = ("file", "status")
            self.tv_single = ttk.Treeview(box, columns=cols, show="headings", height=4)
            self.tv_single.heading("file", text="文件")
            self.tv_single.heading("status", text="状态")
            self.tv_single.column("file", width=560, anchor="w")
            self.tv_single.column("status", width=120, anchor="center")
            s1 = ttk.Scrollbar(box, command=self.tv_single.yview)
            self.tv_single.configure(yscrollcommand=s1.set)
            s1.pack(side="right", fill="y", pady=6, padx=(0, 6))
            self.tv_single.pack(fill="both", expand=True, padx=6, pady=6)

            bar = ttk.Frame(f)
            bar.pack(fill="x", pady=(6, 0))
            ttk.Button(bar, text="添加 Word 文件…", command=self.add_single_files).pack(side="left")
            ttk.Button(bar, text="添加整个文件夹…", command=self.add_single_folder).pack(side="left", padx=6)
            ttk.Button(bar, text="移除所选", command=self.remove_single).pack(side="left")
            ttk.Button(bar, text="清空", command=self.clear_single).pack(side="left", padx=6)

            opt = ttk.LabelFrame(f, text="输出与选项")
            opt.pack(fill="x", pady=(12, 0))
            row = ttk.Frame(opt)
            row.pack(fill="x", padx=8, pady=(8, 4))
            ttk.Label(row, text="输出目录：").pack(side="left")
            self.e_single_out = ttk.Entry(row)
            self.e_single_out.insert(0, default_outdir())
            self.e_single_out.pack(side="left", fill="x", expand=True, padx=6)
            ttk.Button(row, text="浏览…", command=lambda: self.pick_dir(self.e_single_out)).pack(side="left")

            row2 = ttk.Frame(opt)
            row2.pack(fill="x", padx=8, pady=(2, 10))
            self.v_pdf = tk.BooleanVar(value=True)
            self.v_open = tk.BooleanVar(value=True)
            ttk.Checkbutton(row2, text="同时出两种 PDF：提交版 + 排版版（需装 MiKTeX / TeX Live）", style="Toolbutton",
                            variable=self.v_pdf).pack(side="left")
            ttk.Checkbutton(row2, text="完成后打开输出文件夹", variable=self.v_open, style="Toolbutton").pack(side="left", padx=18)
            self.btn_single = ttk.Button(row2, text="开始转换", style="Accent.TButton",
                                         command=self.do_single)
            self.btn_single.pack(side="right")

        def _tab_chapters(self):
            sf = ScrollFrame(self.nb)
            self.nb.add(sf, text="② 合并四章大报告")
            f = sf.inner

            box = ttk.LabelFrame(f, text="稿件与分章（选中行后点下方章节按钮即可分配）")
            box.pack(fill="both", expand=True)
            cols = ("idx", "name", "chapter", "title")
            self.tv_ch = ttk.Treeview(box, columns=cols, show="headings", height=4)
            self.tv_ch.heading("idx", text="#")
            self.tv_ch.heading("name", text="文件名")
            self.tv_ch.heading("chapter", text="章节")
            self.tv_ch.heading("title", text="一级标题（双击可改）")
            self.tv_ch.column("idx", width=42, anchor="center")
            self.tv_ch.column("name", width=330, anchor="w")
            self.tv_ch.column("chapter", width=110, anchor="center")
            self.tv_ch.column("title", width=360, anchor="w")
            s2 = ttk.Scrollbar(box, command=self.tv_ch.yview)
            self.tv_ch.configure(yscrollcommand=s2.set)
            s2.pack(side="right", fill="y", pady=6, padx=(0, 6))
            self.tv_ch.pack(fill="both", expand=True, padx=6, pady=6)
            self.tv_ch.bind("<Double-1>", self._edit_section_title)

            bar = ttk.Frame(f)
            bar.pack(fill="x", pady=(6, 0))
            ttk.Button(bar, text="添加 Word 文件…", command=self.add_chapter_files).pack(side="left")
            ttk.Button(bar, text="添加整个文件夹…", command=self.add_chapter_folder).pack(side="left", padx=6)
            ttk.Button(bar, text="移除所选", command=self.remove_chapter).pack(side="left")
            ttk.Button(bar, text="自动分章", command=self.auto_assign).pack(side="left", padx=(18, 0))
            ttk.Label(bar, text="（按题目关键词先猜一遍，再逐份核对）", style="Sub.TLabel").pack(side="left", padx=8)

            bar2 = ttk.Frame(f)
            bar2.pack(fill="x", pady=(4, 0))
            ttk.Label(bar2, text="选中行分配到：").pack(side="left")
            for i, name in enumerate(CHAPTER_NAMES):
                ttk.Button(bar2, text="→ " + name,
                           command=lambda k=i: self.assign_chapter(k)).pack(side="left", padx=(6, 0))
            ttk.Button(bar2, text="→ 不收录", command=lambda: self.assign_chapter(-1)).pack(side="left", padx=(6, 0))

            # ---------------- 封面（照 23 级 Case 报告样式） ----------------
            cov = ttk.LabelFrame(f, text="封面信息（照 23 级 Case 报告的样式：校徽 + 三行标题 + Tutor/主席/书记/成员）")
            cov.pack(fill="x", pady=(10, 0))
            cg = ttk.Frame(cov)
            cg.pack(fill="x", padx=8, pady=(6, 2))
            ttk.Label(cg, text="首行：").grid(row=0, column=0, sticky="w")
            self.e_cov_line1 = ttk.Entry(cg, width=26)
            self.e_cov_line1.insert(0, "2023级11组PBL学习")
            self.e_cov_line1.grid(row=0, column=1, sticky="we", padx=6, pady=2)
            ttk.Label(cg, text="次行：").grid(row=0, column=2, sticky="w", padx=(14, 0))
            self.e_cov_line2 = ttk.Entry(cg, width=12)
            self.e_cov_line2.insert(0, "学习报告")
            self.e_cov_line2.grid(row=0, column=3, sticky="we", padx=6, pady=2)

            ttk.Label(cg, text="Case（作者填）：").grid(row=1, column=0, sticky="w")
            self.e_cov_case = ttk.Entry(cg, width=26)
            self.e_cov_case.insert(0, "Case 1：难缠的脖子痛")
            self.e_cov_case.grid(row=1, column=1, sticky="we", padx=6, pady=2)
            ttk.Label(cg, text="Tutor：").grid(row=1, column=2, sticky="w", padx=(14, 0))
            self.e_cov_tutor = ttk.Entry(cg, width=12)
            self.e_cov_tutor.insert(0, "马凯歌")
            self.e_cov_tutor.grid(row=1, column=3, sticky="we", padx=6, pady=2)

            ttk.Label(cg, text="主席：").grid(row=2, column=0, sticky="w")
            f_chair = ttk.Frame(cg)
            f_chair.grid(row=2, column=1, sticky="we", padx=6, pady=2)
            self.e_cov_chair = ttk.Entry(f_chair, width=12)
            self.e_cov_chair.insert(0, "朱俊铭")
            self.e_cov_chair.pack(side="left", fill="x", expand=True)
            ttk.Label(f_chair, text="　书记：").pack(side="left")
            self.e_cov_sec = ttk.Entry(f_chair, width=12)
            self.e_cov_sec.insert(0, "郭育恒")
            self.e_cov_sec.pack(side="left", fill="x", expand=True)
            ttk.Label(cg, text="日期：").grid(row=2, column=2, sticky="w", padx=(14, 0))
            self.e_cov_date = ttk.Entry(cg, width=12)
            self.e_cov_date.insert(0, time.strftime("%Y 年 %m 月"))
            self.e_cov_date.grid(row=2, column=3, sticky="we", padx=6, pady=2)

            ttk.Label(cg, text="组员：").grid(row=3, column=0, sticky="w")
            self.e_cov_members = ttk.Entry(cg, width=26)
            self.e_cov_members.insert(0, "叶瑞林 张欣仪 韩曦 任欣萌 王思源 于子轩 牛海华 韩畅")
            self.e_cov_members.grid(row=3, column=1, columnspan=3, sticky="we", padx=6, pady=2)
            ttk.Label(cg, text="（组员空格/逗号分隔，主席与书记不必重复写；Case 只填名称会自动补“Case：”）", style="Sub.TLabel").grid(
                row=4, column=1, columnspan=3, sticky="w", padx=6)
            cg.columnconfigure(1, weight=1)
            cg.columnconfigure(3, weight=1)

            lrow = ttk.Frame(cov)
            lrow.pack(fill="x", padx=8, pady=(2, 2))
            ttk.Label(lrow, text="校徽图片：").pack(side="left")
            self.e_cov_logo = ttk.Entry(lrow)
            self.e_cov_logo.insert(0, default_cover_logo())
            self.e_cov_logo.pack(side="left", fill="x", expand=True, padx=6)
            ttk.Button(lrow, text="浏览…", command=self.pick_cover_logo).pack(side="left")
            ttk.Button(lrow, text="不用图", command=lambda: (self.e_cov_logo.delete(0, "end"))).pack(side="left", padx=6)

            orow = ttk.Frame(cov)
            orow.pack(fill="x", padx=8, pady=(2, 2))
            ttk.Label(orow, text="输出目录：").pack(side="left")
            self.e_ch_out = ttk.Entry(orow, width=30)
            self.e_ch_out.insert(0, default_outdir())
            self.e_ch_out.pack(side="left", fill="x", expand=True, padx=6)
            ttk.Button(orow, text="浏览…", command=lambda: self.pick_dir(self.e_ch_out)).pack(side="left")

            row3 = ttk.Frame(cov)
            row3.pack(fill="x", padx=8, pady=(4, 8))
            self.v_pdf2 = tk.BooleanVar(value=True)
            self.v_open2 = tk.BooleanVar(value=True)
            ttk.Checkbutton(row3, text="顺便编译 PDF", variable=self.v_pdf2, style="Toolbutton").pack(side="left")
            ttk.Checkbutton(row3, text="完成后打开输出文件夹", variable=self.v_open2, style="Toolbutton").pack(side="left", padx=18)
            ttk.Label(row3, text="（点右下角“生成大报告”）", style="Sub.TLabel").pack(side="right")

        def _tab_env(self):
            f = ttk.Frame(self.nb, padding=10)
            self.nb.add(f, text="③ 环境与说明")
            wrap = ttk.Frame(f)
            wrap.pack(fill="both", expand=True)
            self.env_text = tk.Text(wrap, wrap="word", font=("Microsoft YaHei UI", 9), height=14, width=70,
                                    bg="white", relief="flat", padx=10, pady=8)
            env_sb = ttk.Scrollbar(wrap, command=self.env_text.yview)
            self.env_text.configure(yscrollcommand=env_sb.set)
            env_sb.pack(side="right", fill="y")
            self.env_text.pack(side="left", fill="both", expand=True)
            bar = ttk.Frame(f)
            bar.pack(fill="x", pady=(8, 0))
            ttk.Button(bar, text="重新检测环境", command=self.check_env).pack(side="left")
            ttk.Button(bar, text="打开 MiKTeX 下载页", command=self.open_miktex).pack(side="left", padx=8)
            ttk.Button(bar, text="保存日志", command=self.save_log).pack(side="left")
            ttk.Button(bar, text="打开输出文件夹", command=self.open_out).pack(side="left", padx=8)
            self._fill_env_text()

        # ---------------- 文本 ---------------- #
        def _fill_env_text(self):
            from pbltex.toolchain import find_xelatex, xelatex_version
            exe = find_xelatex()
            ver = xelatex_version(exe) if exe else ""
            env = [
                "一、本软件做什么",
                "  把每位同学交上来的 Word 稿件，整理成版式统一的 LaTeX（.tex）并编译成 PDF：",
                "  · 单份转换：1. / 1.1. / 1.1.1. 三级标题，正文小四、1.5 倍行距、首行缩进两字符；",
                "  · 四章合并：章 = 案例回顾 / 重大问题 / 次要问题 / 机制图；",
                "    每人的一份稿件 = 章里的一个一级标题（1. 题目、2. 题目…），",
                "    稿件内部自己的标题依次降一级（1.1 / 1.1.1），编号在每个章内从 1. 重新开始；",
                "    自动生成封面、目录、分章清单，参考文献跨文件去重并重新编号。",
                "  · 封面：照 23 级 Case 报告样式（校徽 + “2023级××组PBL学习 / 学习报告 / Case：…”",
                "    + Tutor / PBL小组主席 / 书记 / 成员）；字号：首行一号、次行小一、Case 小初、",
                "    小组构成三号；标题行距 3 倍、小组构成 1.5 倍；Case 名由作者填（只填名称会自动补“Case：”）。",
                "  · 表格自动适配：按内容取自然宽度、或在版心内折行、或整体等比缩放，不会超出正文宽度；",
                "  · 引用：正文 [1]、[2-3] 直接排成上标编号（GB/T 7714—2015 顺序编码制），",
                "    不依赖两遍编译，所以不会再出现 [?,?,?]。",
                "",
                "二、三步上手",
                "  1. 只转自己那份：「① 单份转换」→ 添加 Word 文件 → 开始转换 → 结果在输出目录下各自子文件夹。",
                "  2. 合成大报告：「② 合并四章大报告」→ 添加所有人的稿件 → 点“自动分章”（按题目关键词猜），",
                "     再选中行用“→ 案例回顾 / 重大问题 / 次要问题 / 机制图”按钮核对修正 → 填标题组别 → 生成大报告。",
                "  3. 找不到 PDF？多半是没装 LaTeX。点“打开 MiKTeX 下载页”安装一次（约 250 MB），",
                "     装完点“重新检测环境”，再重新生成即可。没装也能得到 .tex 源码，随时可补编译。",
                "",
                "三、环境检测结果",
                "  python-docx：已内置（无需另外安装）",
                "  xelatex    ：%s" % (exe if exe else "未检测到（只能生成 .tex，不能编译 PDF）"),
                "  版本信息   ：%s" % (ver if ver else "—"),
                "  输出目录   ：%s" % default_outdir(),
                "",
                "四、常见问题",
                "  · 提示“不支持 .doc”：请用 Word 打开后“另存为” .docx 再转换。",
                "  · 提示文件被占用：先关掉 Word 里正在编辑的那份稿件。",
                "  · 参考文献里英文被拆成两截（如 Americ an）：那是原稿从 PDF 复制留下的断词，",
                "    在 Word 里改好后重新转换即可。",
                "  · 标题识别不准：本工具按“大纲级别 → 编号样式（1.1）→ 加粗短行”依次判断；",
                "    若原稿标题没设大纲级别，在 Word 里套用“标题 1/2/3”样式后识别最准。",
                "  · 生成的大报告排版还想微调：直接用记事本打开 .tex 改，或把 .tex 交给会用 LaTeX 的同学。",
                "",
                "五、日志文件",
                "  %s" % LOG_FILE,
            ]
            self.env_text.configure(state="normal")
            self.env_text.delete("1.0", "end")
            self.env_text.insert("1.0", "\n".join(env))
            self.env_text.configure(state="disabled")

        # ---------------- 日志 / 队列 ---------------- #
        def log(self, msg, tag=None):
            self.q.put(("log", str(msg), tag))

        def _append_log(self, msg, tag=None):
            self.logbox.insert("end", msg + "\n", tag or "")
            self.logbox.see("end")
            try:
                with open(LOG_FILE, "a", encoding="utf-8") as fh:
                    fh.write(msg + "\n")
            except OSError:
                pass

        def _poll(self):
            try:
                while True:
                    item = self.q.get_nowait()
                    kind = item[0]
                    if kind == "log":
                        self._append_log(item[1], item[2] if len(item) > 2 else None)
                    elif kind == "status":
                        self.status.configure(text=item[1])
                    elif kind == "progress":
                        self.progress.configure(value=item[1])
                    elif kind == "done":
                        self._job_done(item[1], item[2])
                    elif kind == "error":
                        self._job_error(item[1], item[2], item[3])
                    elif kind == "row":
                        if self.tv_single.exists(item[1]):
                            self.tv_single.set(item[1], "status", item[2])
                    elif kind == "env":
                        self._apply_env(item[1], item[2])
            except queue.Empty:
                pass
            self.after(110, self._poll)

        def _on_tab_change(self, _event=None):
            """② 合并模式才显示“生成大报告”按钮。"""
            try:
                idx = self.nb.index(self.nb.select())
            except Exception:  # noqa: BLE001
                return
            if idx == 1:
                self.btn_ch.pack(side="right", padx=(0, 12), before=self.progress)
            else:
                self.btn_ch.pack_forget()

        def _set_busy(self, busy):
            self.busy = busy
            state = "disabled" if busy else "normal"
            self.btn_single.configure(state=state)
            self.btn_ch.configure(state=state)
            if busy:
                self.progress.configure(mode="indeterminate")
                self.progress.start(12)
            else:
                self.progress.stop()
                self.progress.configure(mode="determinate", value=0)

        def run_job(self, name, fn):
            if self.busy:
                messagebox.showinfo(APP_NAME, "正在处理上一步，请稍候。")
                return
            self._set_busy(True)
            self.status.configure(text=name + "…")
            self.log("")
            self.log("=== %s ===" % name, "ok")

            def worker():
                try:
                    res = fn()
                    self.q.put(("done", name, res))
                except Exception as exc:  # noqa: BLE001
                    self.q.put(("error", name, str(exc), traceback.format_exc()))
            threading.Thread(target=worker, daemon=True).start()

        def _job_done(self, name, res):
            self._set_busy(False)
            self.status.configure(text="完成：" + name)
            self.log("完成：%s" % name, "ok")

        def _job_error(self, name, msg, tb):
            self._set_busy(False)
            self.status.configure(text="出错：" + name)
            self.log("出错：%s" % msg, "err")
            with open(LOG_FILE, "a", encoding="utf-8") as fh:
                fh.write("\n" + tb + "\n")
            messagebox.showerror(APP_NAME, "%s 失败：\n\n%s" % (name, msg))

        # ---------------- 环境检测 ---------------- #
        def check_env(self):
            from pbltex.toolchain import find_xelatex, reset_cache, xelatex_version
            reset_cache()

            def worker():
                exe = find_xelatex()
                ver = xelatex_version(exe) if exe else ""
                self.q.put(("env", exe, ver))
            threading.Thread(target=worker, daemon=True).start()

        def _apply_env(self, exe, ver):
            self.xelatex_ok = bool(exe)
            if exe:
                tool = "MiKTeX" if "miktex" in (ver or "").lower() else "TeX 引擎"
                self.env_badge.configure(text="● LaTeX 就绪（%s）" % tool, fg="#1a7f4b")
                self.log("LaTeX 引擎：%s" % (ver or exe))
            else:
                self.env_badge.configure(text="● 未检测到 xelatex（只能生成 .tex）", fg="#b3261e")
                self.log("未检测到 xelatex：仍可生成 .tex；要 PDF 请安装 MiKTeX。", "warn")
            self._fill_env_text()

        def open_miktex(self):
            import webbrowser
            webbrowser.open("https://miktex.org/download")

        def open_out(self):
            d = self.e_ch_out.get().strip() or default_outdir()
            self._open_folder(d)

        def _open_folder(self, path):
            try:
                if os.path.isdir(path):
                    os.startfile(path)  # noqa: S606
            except Exception as exc:  # noqa: BLE001
                self.log("打不开文件夹：%s" % exc, "warn")

        def save_log(self):
            txt = self.logbox.get("1.0", "end")
            p = filedialog.asksaveasfilename(defaultextension=".txt", initialfile="转换日志.txt",
                                             filetypes=[("文本文件", "*.txt")])
            if p:
                with open(p, "w", encoding="utf-8") as fh:
                    fh.write(txt)
                self.status.configure(text="日志已保存")

        # ---------------- 单份：文件管理 ---------------- #
        def _pick_docx(self):
            files = filedialog.askopenfilenames(
                title="选择 Word 或 PDF 稿件",
                filetypes=[("Word / PDF", "*.docx *.pdf"), ("Word 文档", "*.docx"),
                           ("PDF 文件", "*.pdf"), ("所有文件", "*.*")])
            return list(files or [])

        def add_single_files(self):
            for f in self._pick_docx():
                if f not in self.single_files:
                    self.single_files.append(f)
                    self.tv_single.insert("", "end", iid=f, values=(f, "待转换"))

        def add_single_folder(self):
            d = filedialog.askdirectory(title="选择存放 Word 稿件的文件夹")
            if not d:
                return
            for name in sorted(os.listdir(d)):
                if name.lower().endswith((".docx", ".pdf")) and not name.startswith("~$"):
                    p = os.path.join(d, name)
                    if p not in self.single_files:
                        self.single_files.append(p)
                        self.tv_single.insert("", "end", iid=p, values=(p, "待转换"))

        def remove_single(self):
            for item in self.tv_single.selection():
                vals = self.tv_single.item(item, "values")
                if vals and vals[0] in self.single_files:
                    self.single_files.remove(vals[0])
                self.tv_single.delete(item)

        def clear_single(self):
            self.single_files.clear()
            for item in self.tv_single.get_children():
                self.tv_single.delete(item)

        def pick_dir(self, entry):
            d = filedialog.askdirectory(title="选择输出目录")
            if d:
                entry.delete(0, "end")
                entry.insert(0, d)

        def pick_cover_logo(self):
            p = filedialog.askopenfilename(
                title="选择封面校徽图片",
                filetypes=[("图片", "*.png *.jpg *.jpeg *.bmp"), ("所有文件", "*.*")])
            if p:
                self.e_cov_logo.delete(0, "end")
                self.e_cov_logo.insert(0, p)

        # ---------------- 单份：执行 ---------------- #
        def do_single(self):
            if not self.single_files:
                messagebox.showinfo(APP_NAME, "请先添加要转换的 Word 或 PDF 文件。")
                return
            out = self.e_single_out.get().strip() or default_outdir()
            want_pdf = self.v_pdf.get()
            open_after = self.v_open.get()      # Tk 变量只能在主线程读
            files = list(self.single_files)

            def job():
                os.makedirs(out, exist_ok=True)
                done, failed = 0, 0
                for i, src in enumerate(files, 1):
                    self.q.put(("status", "正在转换 %d/%d：%s" % (i, len(files), os.path.basename(src))))
                    self.log("[%d/%d] %s" % (i, len(files), os.path.basename(src)))
                    try:
                        info = convert_single(src, out, compile_pdf=want_pdf, log=self.log,
                                              layouts=want_pdf)
                    except Exception as exc:  # noqa: BLE001
                        failed += 1
                        self.q.put(("row", src, "失败"))
                        self.log("  [失败] %s" % exc, "err")
                        continue
                    done += 1
                    self.q.put(("row", src, "已生成 2 版" if info.get("pdf") else "已生成 .tex"))
                    pdf = info.get("pdf") or ""
                    self.log("  标题：%s　标题数：%d　文献：%d　图：%d"
                             % (info["title"][:40], info["headings"], info["refs"], info["figures"]))
                    self.log("  .tex → %s" % info["tex"])
                    if pdf:
                        self.log("  提交版 PDF → %s" % pdf, "ok")
                    if info.get("layout_pdf"):
                        self.log("  排版版 PDF → %s" % info["layout_pdf"], "ok")
                if want_pdf and not self.xelatex_ok:
                    self.log("提示：本机未装 LaTeX，本次只生成了 .tex。", "warn")
                if open_after:
                    self._open_folder(out)
                return "%d 份成功%s" % (done, ("，%d 份失败" % failed) if failed else "")

            self.run_job("单份转换", job)

        # ---------------- 四章：文件管理 ---------------- #
        def _auto_title(self, path):
            """自动识别的一级标题（缓存），仅用于界面展示。"""
            if path not in self._title_cache:
                title = ""
                try:
                    if path.lower().endswith(".pdf"):
                        from pbltex.core import pdf_head
                        title = pdf_head(path)[:40]
                    else:
                        import docx
                        doc = docx.Document(path)
                        for par in doc.paragraphs[:6]:
                            t = " ".join((par.text or "").split())
                            if t:
                                title = t[:40]
                                break
                except Exception:  # noqa: BLE001
                    title = "(打不开)"
                self._title_cache[path] = title
            return self._title_cache.get(path, "")

        def _refresh_ch_tree(self):
            sel = list(self.tv_ch.selection())
            for item in self.tv_ch.get_children():
                self.tv_ch.delete(item)
            for i, row in enumerate(self.chapter_items):
                path, chap = row[0], row[1]
                override = row[2] if len(row) > 2 else ""
                chap_name = CHAPTER_NAMES[chap] if 0 <= chap < len(CHAPTER_NAMES) else "不收录"
                shown = override or self._auto_title(path)
                if override:
                    shown = "✎ " + shown
                self.tv_ch.insert("", "end", iid=str(i),
                                  values=(i + 1, os.path.basename(path), chap_name, shown))
            keep = [x for x in sel if x in self.tv_ch.get_children()]
            if keep:
                self.tv_ch.selection_set(keep)

        def _edit_section_title(self, event=None):
            """双击一行，改它在大报告里的一级标题。"""
            row = self.tv_ch.identify_row(event.y) if event is not None else ""
            if not row:
                sel = self.tv_ch.selection()
                if not sel:
                    return
                row = sel[0]
            i = int(row)
            if not (0 <= i < len(self.chapter_items)):
                return
            path = self.chapter_items[i][0]
            cur = (self.chapter_items[i][2] if len(self.chapter_items[i]) > 2 else "") \
                or self._auto_title(path)

            dlg = tk.Toplevel(self)
            dlg.title("改一级标题")
            dlg.transient(self)
            dlg.grab_set()
            dlg.resizable(True, False)
            ttk.Label(dlg, text="文件：" + os.path.basename(path), style="Sub.TLabel").pack(
                anchor="w", padx=12, pady=(10, 2))
            ttk.Label(dlg, text="在大报告里显示为（留空＝自动识别）：").pack(anchor="w", padx=12)
            var = tk.StringVar(value=cur)
            ent = ttk.Entry(dlg, textvariable=var, width=52)
            ent.pack(fill="x", padx=12, pady=6)
            ent.focus_set()
            ent.icursor("end")

            def ok(_e=None):
                val = var.get().strip()
                self.chapter_items[i] = (path, self.chapter_items[i][1],
                                         "" if val == cur and not self.chapter_items[i][2:] else val)
                dlg.destroy()
                self._refresh_ch_tree()
                self.status.configure(text="已更新一级标题")

            bar = ttk.Frame(dlg)
            bar.pack(fill="x", padx=12, pady=(0, 10))
            ttk.Button(bar, text="恢复自动识别",
                       command=lambda: (var.set(""), ok())).pack(side="left")
            ttk.Button(bar, text="确定", style="Accent.TButton", command=ok).pack(side="right")
            ttk.Button(bar, text="取消", command=dlg.destroy).pack(side="right", padx=8)
            ent.bind("<Return>", ok)
            dlg.bind("<Escape>", lambda _e: dlg.destroy())

        def add_chapter_files(self):
            for f in self._pick_docx():
                if all(f != r[0] for r in self.chapter_items):
                    self.chapter_items.append((f, -1, ""))
            self._refresh_ch_tree()

        def add_chapter_folder(self):
            d = filedialog.askdirectory(title="选择存放 Word 稿件的文件夹")
            if not d:
                return
            for name in sorted(os.listdir(d)):
                if name.lower().endswith((".docx", ".pdf")) and not name.startswith("~$"):
                    p = os.path.join(d, name)
                    if all(p != r[0] for r in self.chapter_items):
                        self.chapter_items.append((p, -1, ""))
            self._refresh_ch_tree()

        def remove_chapter(self):
            sel = [int(i) for i in self.tv_ch.selection()]
            for i in sorted(sel, reverse=True):
                if 0 <= i < len(self.chapter_items):
                    self.chapter_items.pop(i)
            self._refresh_ch_tree()

        def assign_chapter(self, chap):
            if not self.tv_ch.selection():
                messagebox.showinfo(APP_NAME, "先在列表里选中要分配的行（可按住 Ctrl 多选）。")
                return
            for i in [int(x) for x in self.tv_ch.selection()]:
                if 0 <= i < len(self.chapter_items):
                    row = self.chapter_items[i]
                    self.chapter_items[i] = (row[0], chap, row[2] if len(row) > 2 else "")
            self._refresh_ch_tree()
            self.status.configure(text="已更新 %d 行" % len(self.tv_ch.selection()))

        def auto_assign(self):
            if not self.chapter_items:
                messagebox.showinfo(APP_NAME, "请先添加稿件。")
                return
            self.log("按题目关键词自动分章：")
            for i, row in enumerate(self.chapter_items):
                path = row[0]
                chap = guess_chapter(path)
                self.chapter_items[i] = (path, chap, row[2] if len(row) > 2 else "")
                self.log("  %s → %s" % (os.path.basename(path), CHAPTER_NAMES[chap]))
            self._refresh_ch_tree()
            self.status.configure(text="自动分章完成，请核对后生成")

        # ---------------- 四章：执行 ---------------- #
        def do_chapters(self):
            items = [(r[0], r[1]) for r in self.chapter_items if r[1] >= 0]
            overrides = {r[0]: r[2] for r in self.chapter_items if len(r) > 2 and r[2]}
            if not items:
                messagebox.showinfo(APP_NAME, "还没有分配章节的稿件：请添加稿件并用下方章节按钮分配。")
                return
            out = self.e_ch_out.get().strip() or default_outdir()
            want_pdf = self.v_pdf2.get()
            open_after = self.v_open2.get()     # Tk 变量只能在主线程读
            cover = {
                "line1": self.e_cov_line1.get().strip(),
                "line2": self.e_cov_line2.get().strip() or "学习报告",
                "case": self.e_cov_case.get().strip(),
                "tutor": self.e_cov_tutor.get().strip(),
                "chair": self.e_cov_chair.get().strip(),
                "secretary": self.e_cov_sec.get().strip(),
                "members": self.e_cov_members.get().strip(),
                "date": self.e_cov_date.get().strip(),
                "logo": self.e_cov_logo.get().strip(),
            }
            title = cover["line1"] or "PBL 大报告"
            author = cover["line1"]
            date = cover["date"]

            def job():
                os.makedirs(out, exist_ok=True)
                self.q.put(("status", "正在合成四章大报告…"))
                info = convert_chapters(items, out, title=title, author=author, date=date,
                                        compile_pdf=want_pdf, log=self.log, cover=cover,
                                        section_titles=overrides)
                self.log("  .tex → %s" % info["tex"])
                self.log("  分章清单 → %s" % info["manifest"])
                if info.get("pdf"):
                    self.log("  .pdf → %s" % info["pdf"], "ok")
                elif want_pdf:
                    self.log("  未能生成 PDF（见上方 xelatex 报错）", "warn")
                self.log("  参考文献去重后：%d 条" % info["refs"])
                if open_after:
                    self._open_folder(out)
                return "四章大报告已生成（%d 份稿件）" % len(items)

            self.run_job("生成四章大报告", job)

        def add_files_preloaded(self, files):
            """命令行把 docx 拖到 exe 上时，直接加进单份转换列表。"""
            for f in files:
                if f.lower().endswith(".docx") and f not in self.single_files:
                    self.single_files.append(f)
                    self.tv_single.insert("", "end", iid=f, values=(f, "待转换"))
            if files:
                self.status.configure(text="已载入 %d 个文件，点“开始转换”即可" % len(self.single_files))

        def _fit_screen(self):
            """按屏幕大小摆好窗口：小屏笔记本上不会跑出桌面"""
            try:
                sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
                w = max(min(1180, sw - 80), 1000)
                h = max(min(900, sh - 110), 640)
                x = max((sw - w) // 2, 0)
                y = max((sh - h) // 3, 0)
                self.geometry("%dx%d+%d+%d" % (w, h, x, y))
            except Exception:  # noqa: BLE001
                pass

        def on_close(self):
            if self.busy and not messagebox.askokcancel(APP_NAME, "还在处理中，确定退出吗？"):
                return
            self.destroy()

    app = App()
    if preload:
        app.after(300, lambda: app.add_files_preloaded(preload))
    app.mainloop()
    return 0


# --------------------------------------------------------------------------- #
def main() -> int:
    argv = sys.argv[1:]
    if argv and argv[0].startswith("--"):
        return run_cli(argv)
    docs = [a for a in argv if a.lower().endswith((".docx", ".doc"))]
    return launch_gui(preload=docs or None)


if __name__ == "__main__":
    try:
        code = main()
    except Exception:  # noqa: BLE001
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as fh:
                fh.write(traceback.format_exc())
        except OSError:
            pass
        raise
    sys.exit(code)































