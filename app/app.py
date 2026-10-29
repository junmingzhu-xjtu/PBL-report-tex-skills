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
APP_VERSION = "1.8.0"
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


        # ---------------- 视觉：配色与样式 ---------------- #
    FONT = "Microsoft YaHei UI"
    INK = "#101828"
    SUB_TXT = "#667085"
    LINE = "#E5E7EB"
    SOFT = "#F7F8FA"
    ACCENT = "#A5223A"
    ACCENT_DK = "#8A1B2F"
    ACCENT_SOFT = "#FAEDEF"
    GREEN = "#067647"
    GREEN_BG = "#ECFDF3"
    RED_TXT = "#B42318"
    RED_BG = "#FEF3F2"

    class ScrollFrame(ttk.Frame):
        """纵向可滚动的标签页（窗口小时内容不会丢，按钮不会被挤没）。"""

        def __init__(self, master):
            super().__init__(master, style="Page.TFrame")
            self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, bg="white")
            self.vs = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview,
                                    style="Slim.Vertical.TScrollbar")
            self.canvas.configure(yscrollcommand=self._sync_bar)
            self._vs_shown = False
            self.canvas.pack(side="left", fill="both", expand=True)
            self.inner = tk.Frame(self.canvas, bg="white", padx=2, pady=2)
            self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
            self.inner.bind("<Configure>", self._on_inner)
            self.canvas.bind("<Configure>", self._on_canvas)
            self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._wheel))
            self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

        def _sync_bar(self, first, last):
            # 内容超出一屏才显示滚动条，平时保持干净
            need = not (float(first) <= 0.0 and float(last) >= 1.0)
            if need and not self._vs_shown:
                self.vs.pack(side="right", fill="y", before=self.canvas)
                self._vs_shown = True
            elif not need and self._vs_shown:
                self.vs.pack_forget()
                self._vs_shown = False
            self.vs.set(first, last)

        def _on_inner(self, _e):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            h = max(self.inner.winfo_reqheight(), self.canvas.winfo_height())
            self.canvas.itemconfigure(self._win, height=h)

        def _on_canvas(self, event):
            self.canvas.itemconfigure(self._win, width=event.width)
            # 让内容至少有 canvas 那么高，列表才能“吃掉”多余空间
            need = max(self.inner.winfo_reqheight(), event.height)
            self.canvas.itemconfigure(self._win, height=need)

        def _wheel(self, event):
            self.canvas.yview_scroll(int(-event.delta / 120), "units")

    class TabView(tk.Frame):
        """自绘标签栏：白色标题条 + 选中项交大红下划线（比 ttk.Notebook 干净）。"""

        def __init__(self, master):
            super().__init__(master, bg="white")
            self._nav = tk.Frame(self, bg="white")
            self._nav.pack(fill="x")
            tk.Frame(self, bg=LINE, height=1).pack(fill="x")
            self._body = tk.Frame(self, bg="white")
            self._body.pack(fill="both", expand=True)
            self._tabs = []
            self._cur = 0
            self._callbacks = []

        def add(self, frame, text=""):
            idx = len(self._tabs)
            item = tk.Frame(self._nav, bg="white")
            item.pack(side="left", padx=(0, 2))
            lbl = tk.Label(item, text=text, bg="white", fg=SUB_TXT,
                           font=(FONT, 10, "bold"), padx=16, pady=12, cursor="hand2")
            lbl.pack()
            bar = tk.Frame(item, bg="white", height=3)
            bar.pack(fill="x")

            def enter(_e, w=lbl, k=idx):
                w.configure(fg=(ACCENT if k == self._cur else INK))

            def leave(_e, w=lbl, k=idx):
                w.configure(fg=(ACCENT if k == self._cur else SUB_TXT))

            lbl.bind("<Button-1>", lambda _e, k=idx: self.select(k))
            lbl.bind("<Enter>", enter)
            lbl.bind("<Leave>", leave)
            self._tabs.append({"frame": frame, "label": lbl, "bar": bar})
            frame.pack(in_=self._body, fill="both", expand=True)
            frame.pack_forget()
            return idx

        def select(self, idx=None):
            if idx is None:
                return self._cur
            if not self._tabs or not (0 <= idx < len(self._tabs)):
                return self._cur
            self._cur = idx
            for k, tab in enumerate(self._tabs):
                on = (k == idx)
                tab["label"].configure(fg=(ACCENT if on else SUB_TXT))
                tab["bar"].configure(bg=(ACCENT if on else "white"))
                if on:
                    tab["frame"].pack(in_=self._body, fill="both", expand=True)
                else:
                    tab["frame"].pack_forget()
            for cb in list(self._callbacks):
                try:
                    cb(None)
                except Exception:  # noqa: BLE001
                    pass
            return idx

        def index(self, what="current"):
            return what if isinstance(what, int) else self._cur

        def bind(self, seq=None, func=None, add=None):  # noqa: A003
            if seq == "<<NotebookTabChanged>>":
                self._callbacks.append(func)
                return ""
            return super().bind(seq, func, add)

    class Toggle(tk.Frame):
        """拨动开关，替代 clam 那个带叉的复选框。"""

        W, H = 40, 22

        def __init__(self, master, text, var):
            super().__init__(master, bg="white")
            self.var = var
            self.cv = tk.Canvas(self, width=self.W, height=self.H, bg="white",
                                highlightthickness=0, cursor="hand2")
            self.cv.pack(side="left")
            self.lbl = tk.Label(self, text=text, bg="white", fg=INK, font=(FONT, 9),
                                cursor="hand2")
            self.lbl.pack(side="left", padx=(9, 0))
            for w in (self.cv, self.lbl):
                w.bind("<Button-1>", self._click)
            try:
                var.trace_add("write", lambda *a: self._draw())
            except Exception:  # noqa: BLE001
                pass
            self._draw()

        def _draw(self):
            on = bool(self.var.get())
            track = ACCENT if on else "#D5D9E0"
            self.cv.delete("all")
            self.cv.create_oval(1, 4, 19, 18, fill=track, outline=track)
            self.cv.create_oval(21, 4, 39, 18, fill=track, outline=track)
            self.cv.create_rectangle(10, 4, 30, 18, fill=track, outline=track)
            x = 30 if on else 10
            self.cv.create_oval(x - 7, 3, x + 7, 19, fill="white", outline="#EDEFF2")

        def _click(self, _e=None):
            self.var.set(not self.var.get())

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("%s  v%s" % (APP_NAME, APP_VERSION))
            self.geometry("1320x920")
            self.minsize(1120, 720)
            self.configure(bg="white")
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
            st.configure(".", font=(FONT, 10), background="white", foreground=INK)
            st.configure("TFrame", background="white")
            st.configure("Page.TFrame", background="white")
            st.configure("TLabel", background="white", foreground=INK)
            st.configure("Muted.TLabel", background="white", foreground=SUB_TXT, font=(FONT, 9))
            st.configure("Field.TLabel", background="white", foreground=SUB_TXT, font=(FONT, 9))

            st.configure("Field.TEntry", padding=(8, 5), fieldbackground="white",
                         bordercolor=LINE, lightcolor=LINE, darkcolor=LINE,
                         insertcolor=INK, foreground=INK)
            st.map("Field.TEntry",
                   bordercolor=[("focus", ACCENT)], lightcolor=[("focus", ACCENT)],
                   darkcolor=[("focus", ACCENT)])

            st.configure("TButton", font=(FONT, 9), padding=(12, 7), background="white",
                         foreground=INK, bordercolor=LINE, lightcolor="white", darkcolor=LINE,
                         relief="flat", focusthickness=0)
            st.map("TButton",
                   background=[("active", SOFT), ("pressed", "#EFF1F4")],
                   bordercolor=[("active", "#D0D5DD")])
            st.configure("Soft.TButton", font=(FONT, 9), padding=(12, 7), background=SOFT,
                         foreground=INK, bordercolor=LINE)
            st.map("Soft.TButton", background=[("active", "#ECEEF2")])
            st.configure("Ghost.TButton", font=(FONT, 9), padding=(10, 6), background="white",
                         foreground=SUB_TXT, bordercolor="white")
            st.map("Ghost.TButton", background=[("active", SOFT)])
            st.configure("Primary.TButton", font=(FONT, 10, "bold"), padding=(20, 10),
                         background=ACCENT, foreground="white", bordercolor=ACCENT,
                         lightcolor=ACCENT, darkcolor=ACCENT_DK, relief="flat")
            st.map("Primary.TButton",
                   background=[("active", ACCENT_DK), ("pressed", ACCENT_DK),
                               ("disabled", "#E4E7EC")],
                   foreground=[("disabled", "#98A2B3")],
                   bordercolor=[("disabled", "#E4E7EC")])

            st.configure("TCheckbutton", background="white", foreground=INK, font=(FONT, 9),
                         focusthickness=0, indicatorcolor="white",
                         indicatorrelief="flat", upperbordercolor="#C9CED6",
                         lowerbordercolor="#C9CED6", focuscolor="white")
            st.map("TCheckbutton", background=[("active", "white")],
                   foreground=[("disabled", "#98A2B3")],
                   indicatorcolor=[("selected", ACCENT), ("pressed", ACCENT)])
            st.configure("Card.TCheckbutton", background="white", foreground=INK,
                         font=(FONT, 9))

            st.configure("Treeview", font=(FONT, 9), rowheight=32, background="white",
                         fieldbackground="white", foreground=INK, borderwidth=0, relief="flat")
            st.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
            st.configure("Treeview.Heading", font=(FONT, 9, "bold"), background=SOFT,
                         foreground=SUB_TXT, relief="flat", borderwidth=0, padding=(8, 9))
            st.map("Treeview.Heading", background=[("active", "#EFF1F4")])
            st.map("Treeview", background=[("selected", ACCENT_SOFT)],
                   foreground=[("selected", ACCENT)])

            st.configure("Slim.Vertical.TScrollbar", background=SOFT, troughcolor="white",
                         bordercolor="white", arrowcolor=SUB_TXT, arrowsize=12)
            st.configure("Slim.Horizontal.TProgressbar", background=ACCENT, troughcolor="#F0F1F4",
                         bordercolor="#F0F1F4", thickness=6)

        # ---------------- 小部件 ---------------- #
        def _hline(self, parent, pady=(0, 0)):
            ln = tk.Frame(parent, bg=LINE, height=1)
            ln.pack(fill="x", pady=pady)
            return ln

        def _section(self, parent, title, hint=""):
            """带交大红竖条的分区标题，返回内容区。"""
            head = tk.Frame(parent, bg="white")
            head.pack(fill="x", pady=(9, 4))
            tk.Frame(head, bg=ACCENT, width=3, height=15).pack(side="left", padx=(0, 9))
            tk.Label(head, text=title, bg="white", fg=INK,
                     font=(FONT, 11, "bold")).pack(side="left")
            if hint:
                tk.Label(head, text=hint, bg="white", fg=SUB_TXT,
                         font=(FONT, 9)).pack(side="left", padx=(10, 0))
            body = tk.Frame(parent, bg="white")
            body.pack(fill="both", expand=True)
            return body

        def _bordered(self, parent, bg="white"):
            """1px 细边框的容器（树形表格、文本框用）。"""
            wrap = tk.Frame(parent, bg=LINE, padx=1, pady=1)
            inner = tk.Frame(wrap, bg=bg)
            inner.pack(fill="both", expand=True)
            return wrap, inner

        # ---------------- 布局 ---------------- #
        def _build(self):
            # ---------- 顶栏 ----------
            header = tk.Frame(self, bg="white")
            header.pack(side="top", fill="x")
            hpad = tk.Frame(header, bg="white")
            hpad.pack(fill="x", padx=22, pady=(14, 12))

            left = tk.Frame(hpad, bg="white")
            left.pack(side="left")
            mark = tk.Canvas(left, width=38, height=38, bg="white", highlightthickness=0)
            mark.pack(side="left", padx=(0, 14))
            mark.create_rectangle(0, 0, 38, 38, fill=ACCENT, outline="")
            mark.create_text(19, 13, text="PBL", fill="white", font=(FONT, 7, "bold"))
            mark.create_text(19, 25, text="TEX", fill="#F3C6CE", font=(FONT, 7, "bold"))
            tbox = tk.Frame(left, bg="white")
            tbox.pack(side="left")
            tk.Label(tbox, text=APP_NAME, bg="white", fg=INK,
                     font=(FONT, 15, "bold")).pack(anchor="w")
            tk.Label(tbox, text="Word / PDF 稿件 → 规范 LaTeX 与两种 PDF　·　四章式大报告一键合成",
                     bg="white", fg=SUB_TXT, font=(FONT, 9)).pack(anchor="w", pady=(3, 0))

            self.env_badge = tk.Label(hpad, text="正在检测 LaTeX 环境…", font=(FONT, 9),
                                      fg=SUB_TXT, bg=SOFT, padx=12, pady=6)
            self.env_badge.pack(side="right")
            tk.Frame(self, bg=LINE, height=1).pack(side="top", fill="x")

            # ---------- 状态栏（先摆底部）----------
            sbar = tk.Frame(self, bg="white")
            sbar.pack(side="bottom", fill="x")
            tk.Frame(sbar, bg=LINE, height=1).pack(fill="x")
            sin = tk.Frame(sbar, bg="white")
            sin.pack(fill="x", padx=22, pady=8)
            self.status = tk.Label(sin, text="就绪", bg="white", fg=SUB_TXT, font=(FONT, 9))
            self.status.pack(side="left")
            self.progress = ttk.Progressbar(sin, mode="determinate", length=180,
                                            style="Slim.Horizontal.TProgressbar")
            self.progress.pack(side="right")
            self.btn_ch = ttk.Button(sin, text="生成大报告", style="Primary.TButton",
                                     command=self.do_chapters)

            # ---------- 运行日志 ----------
            logwrap = tk.Frame(self, bg="white")
            logwrap.pack(side="bottom", fill="both", padx=22, pady=(2, 8))
            lhead = tk.Frame(logwrap, bg="white")
            lhead.pack(fill="x", pady=(8, 5))
            tk.Label(lhead, text="运行日志", bg="white", fg=INK,
                     font=(FONT, 9, "bold")).pack(side="left")
            tk.Label(lhead, text="·  「③ 环境与说明」里可以保存", bg="white", fg=SUB_TXT,
                     font=(FONT, 8)).pack(side="left", padx=(8, 0))
            self._log_open = False
            self.log_toggle = tk.Label(lhead, text="展开 ▾", bg="white", fg=SUB_TXT,
                                       font=(FONT, 8), cursor="hand2")
            self.log_toggle.pack(side="right")
            lbox = tk.Frame(logwrap, bg="#0B1220")
            lbox.pack(fill="both")
            self.logbox = tk.Text(lbox, height=3, wrap="word", font=("Consolas", 9),
                                  bg="#0B1220", fg="#C7D2E0", insertbackground="#C7D2E0",
                                  relief="flat", padx=14, pady=10, borderwidth=0,
                                  highlightthickness=0, spacing1=1, spacing3=2)
            lsb = ttk.Scrollbar(lbox, command=self.logbox.yview,
                                style="Slim.Vertical.TScrollbar")
            self.logbox.configure(yscrollcommand=lsb.set)
            lsb.pack(side="right", fill="y")
            self.logbox.pack(side="left", fill="both", expand=True)
            lbox.pack_forget()          # 默认收起，点右上角“展开”再看日志
            def _toggle_log(_e=None):
                self._log_open = not self._log_open
                if self._log_open:
                    lbox.pack(fill="both")
                    self.log_toggle.configure(text="展开 ▾")
                else:
                    lbox.pack_forget()
                    self.log_toggle.configure(text="展开 ▾")

            self.log_toggle.bind("<Button-1>", _toggle_log)
            self.logbox.tag_configure("ok", foreground="#6EE7A8")
            self.logbox.tag_configure("warn", foreground="#FBBF6A")
            self.logbox.tag_configure("err", foreground="#FB8A8A")

            # ---------- 标签页 ----------
            self.nb = TabView(self)
            self.nb.pack(side="top", fill="both", expand=True, padx=18, pady=(2, 0))
            self._tab_single()
            self._tab_chapters()
            self._tab_env()
            self.nb.bind("<<NotebookTabChanged>>", self._on_tab_change)
            self.nb.select(0)

        def _tab_single(self):
            sf = ScrollFrame(self.nb)
            self.nb.add(sf, text="①  单份转换")
            f = sf.inner
            f.configure(padx=14, pady=10)
            f.columnconfigure(0, weight=1)
            f.rowconfigure(2, weight=1)

            # ---- 稿件列表 ----
            head = tk.Frame(f, bg="white")
            head.grid(row=0, column=0, sticky="we", pady=(0, 8))
            tk.Frame(head, bg=ACCENT, width=3, height=15).pack(side="left", padx=(0, 9))
            tk.Label(head, text="待转换的稿件", bg="white", fg=INK,
                     font=(FONT, 12, "bold")).pack(side="left")
            tk.Label(head, text=".docx 或 .pdf　·　每份稿件独立输出一个文件夹", bg="white",
                     fg=SUB_TXT, font=(FONT, 9)).pack(side="left", padx=(10, 0))

            wrap = tk.Frame(f, bg=LINE, padx=1, pady=1)
            wrap.grid(row=2, column=0, sticky="nsew")
            inner = tk.Frame(wrap, bg="white")
            inner.pack(fill="both", expand=True)
            cols = ("file", "status")
            self.tv_single = ttk.Treeview(inner, columns=cols, show="headings")
            self.tv_single.heading("file", text="文件")
            self.tv_single.heading("status", text="状态")
            self.tv_single.column("file", width=760, anchor="w")
            self.tv_single.column("status", width=150, anchor="center")
            self.tv_single.tag_configure("odd", background="#FCFCFD")
            s1 = ttk.Scrollbar(inner, command=self.tv_single.yview,
                               style="Slim.Vertical.TScrollbar")
            self.tv_single.configure(yscrollcommand=s1.set)
            s1.pack(side="right", fill="y")
            self.tv_single.pack(side="left", fill="both", expand=True)
            self.hint_single = tk.Label(inner, text="还没有添加文件　·　点下面「添加文件…」，或直接把 Word / PDF 拖进来",
                                        bg="white", fg="#98A2B3", font=(FONT, 9))
            self.hint_single.place(relx=0.5, rely=0.5, anchor="center")

            bar = tk.Frame(f, bg="white")
            bar.grid(row=3, column=0, sticky="we", pady=(10, 0))
            ttk.Button(bar, text="添加文件…", style="Soft.TButton",
                       command=self.add_single_files).pack(side="left")
            ttk.Button(bar, text="添加整个文件夹…",
                       command=self.add_single_folder).pack(side="left", padx=8)
            ttk.Button(bar, text="移除所选", style="Ghost.TButton",
                       command=self.remove_single).pack(side="left")
            ttk.Button(bar, text="清空", style="Ghost.TButton",
                       command=self.clear_single).pack(side="left", padx=4)
            ttk.Label(bar, text="转换完成后文件名带「-提交版 / -排版版」",
                      style="Muted.TLabel").pack(side="right")

            # ---- 输出与选项（底栏）----
            tk.Frame(f, bg=LINE, height=1).grid(row=4, column=0, sticky="we", pady=(14, 12))
            opt = tk.Frame(f, bg="white")
            opt.grid(row=5, column=0, sticky="we")
            opt.columnconfigure(1, weight=1)

            ttk.Label(opt, text="输出目录", style="Field.TLabel").grid(row=0, column=0, sticky="w")
            self.e_single_out = ttk.Entry(opt, style="Field.TEntry")
            self.e_single_out.insert(0, default_outdir())
            self.e_single_out.grid(row=0, column=1, sticky="we", padx=(10, 8))
            ttk.Button(opt, text="浏览…", command=lambda: self.pick_dir(self.e_single_out)
                       ).grid(row=0, column=2, sticky="e")

            row2 = tk.Frame(opt, bg="white")
            row2.grid(row=1, column=0, columnspan=3, sticky="we", pady=(12, 6))
            self.v_pdf = tk.BooleanVar(value=True)
            self.v_open = tk.BooleanVar(value=True)
            Toggle(row2, "同时出两种 PDF（提交版 + 排版版）", self.v_pdf).pack(side="left")
            Toggle(row2, "完成后打开输出文件夹", self.v_open).pack(side="left", padx=24)
            self.btn_single = ttk.Button(row2, text="开始转换", style="Primary.TButton",
                                         command=self.do_single)
            self.btn_single.pack(side="right")

        def _tab_chapters(self):
            sf = ScrollFrame(self.nb)
            self.nb.add(sf, text="②  合并四章")
            f = sf.inner

            body = self._section(f, "稿件与分章", "选中行后点下方章节按钮分配；双击可改一级标题")
            wrap, inner = self._bordered(body)
            wrap.pack(fill="both", expand=True)
            cols = ("idx", "name", "chapter", "title")
            self.tv_ch = ttk.Treeview(inner, columns=cols, show="headings", height=2)
            self.tv_ch.heading("idx", text="#")
            self.tv_ch.heading("name", text="文件名")
            self.tv_ch.heading("chapter", text="章节")
            self.tv_ch.heading("title", text="一级标题（双击可改）")
            self.tv_ch.column("idx", width=44, anchor="center")
            self.tv_ch.column("name", width=330, anchor="w")
            self.tv_ch.column("chapter", width=120, anchor="center")
            self.tv_ch.column("title", width=380, anchor="w")
            self.tv_ch.tag_configure("odd", background="#FCFCFD")
            s2 = ttk.Scrollbar(inner, command=self.tv_ch.yview,
                               style="Slim.Vertical.TScrollbar")
            self.tv_ch.configure(yscrollcommand=s2.set)
            s2.pack(side="right", fill="y")
            self.tv_ch.pack(side="left", fill="both", expand=True)
            self.hint_ch = tk.Label(inner, text="还没有添加稿件　·　添加后点「自动分章」，再核对章节",
                                    bg="white", fg="#98A2B3", font=(FONT, 9))
            self.hint_ch.place(relx=0.5, rely=0.55, anchor="center")
            self.tv_ch.bind("<Double-1>", self._edit_section_title)

            bar = tk.Frame(f, bg="white")
            bar.pack(fill="x", pady=(10, 0))
            ttk.Button(bar, text="添加文件…", style="Soft.TButton",
                       command=self.add_chapter_files).pack(side="left")
            ttk.Button(bar, text="添加整个文件夹…",
                       command=self.add_chapter_folder).pack(side="left", padx=8)
            ttk.Button(bar, text="移除所选", style="Ghost.TButton",
                       command=self.remove_chapter).pack(side="left")
            ttk.Button(bar, text="自动分章", style="Soft.TButton",
                       command=self.auto_assign).pack(side="left", padx=(20, 0))
            ttk.Label(bar, text="按题目关键词先猜一遍，再逐份核对", style="Muted.TLabel"
                      ).pack(side="left", padx=10)

            bar2 = tk.Frame(f, bg="white")
            bar2.pack(fill="x", pady=(6, 0))
            ttk.Label(bar2, text="选中行分配到", style="Field.TLabel").pack(side="left", padx=(0, 8))
            for i, name in enumerate(CHAPTER_NAMES):
                ttk.Button(bar2, text=name, style="Soft.TButton",
                           command=lambda k=i: self.assign_chapter(k)).pack(side="left", padx=(0, 6))
            ttk.Button(bar2, text="不收录", style="Ghost.TButton",
                       command=lambda: self.assign_chapter(-1)).pack(side="left", padx=(2, 0))

            # ---------------- 封面 ----------------
            cov = self._section(f, "封面信息", "校徽 + 三行标题 + Tutor / 主席 / 书记 / 成员（组员用空格分隔）")
            cg = tk.Frame(cov, bg="white")
            cg.pack(fill="x")
            for col, w in ((1, 1), (3, 1)):
                cg.columnconfigure(col, weight=w)

            def lab(row, col, text, pad=(0, 0)):
                ttk.Label(cg, text=text, style="Field.TLabel").grid(
                    row=row, column=col, sticky="w", padx=pad, pady=2)

            lab(0, 0, "首行")
            self.e_cov_line1 = ttk.Entry(cg, style="Field.TEntry")
            self.e_cov_line1.insert(0, "2023级11组PBL学习")
            self.e_cov_line1.grid(row=0, column=1, sticky="we", padx=6, pady=3)
            lab(0, 2, "次行", (14, 0))
            self.e_cov_line2 = ttk.Entry(cg, style="Field.TEntry")
            self.e_cov_line2.insert(0, "学习报告")
            self.e_cov_line2.grid(row=0, column=3, sticky="we", padx=6, pady=3)

            lab(1, 0, "Case（作者填）")
            self.e_cov_case = ttk.Entry(cg, style="Field.TEntry")
            self.e_cov_case.insert(0, "Case 1：难缠的脖子痛")
            self.e_cov_case.grid(row=1, column=1, sticky="we", padx=6, pady=3)
            lab(1, 2, "Tutor", (14, 0))
            self.e_cov_tutor = ttk.Entry(cg, style="Field.TEntry")
            self.e_cov_tutor.insert(0, "马凯歌")
            self.e_cov_tutor.grid(row=1, column=3, sticky="we", padx=6, pady=3)

            lab(2, 0, "主席")
            f_chair = tk.Frame(cg, bg="white")
            f_chair.grid(row=2, column=1, sticky="we", padx=6, pady=3)
            self.e_cov_chair = ttk.Entry(f_chair, style="Field.TEntry")
            self.e_cov_chair.insert(0, "朱俊铭")
            self.e_cov_chair.pack(side="left", fill="x", expand=True)
            ttk.Label(f_chair, text="书记", style="Field.TLabel").pack(side="left", padx=(10, 6))
            self.e_cov_sec = ttk.Entry(f_chair, style="Field.TEntry")
            self.e_cov_sec.insert(0, "郭育恒")
            self.e_cov_sec.pack(side="left", fill="x", expand=True)
            lab(2, 2, "日期", (14, 0))
            self.e_cov_date = ttk.Entry(cg, style="Field.TEntry")
            self.e_cov_date.insert(0, time.strftime("%Y 年 %m 月"))
            self.e_cov_date.grid(row=2, column=3, sticky="we", padx=6, pady=3)

            lab(3, 0, "组员")
            self.e_cov_members = ttk.Entry(cg, style="Field.TEntry")
            self.e_cov_members.insert(0, "叶瑞林 张欣仪 韩曦 任欣萌 王思源 于子轩 牛海华 韩畅")
            self.e_cov_members.grid(row=3, column=1, columnspan=3, sticky="we", padx=6, pady=3)


            lrow = tk.Frame(cov, bg="white")
            lrow.pack(fill="x", pady=(6, 0))
            ttk.Label(lrow, text="封面图片（可留空）", style="Field.TLabel").pack(side="left", padx=(0, 8))
            self.e_cov_logo = ttk.Entry(lrow, style="Field.TEntry")
            self.e_cov_logo.insert(0, "")
            self.e_cov_logo.pack(side="left", fill="x", expand=True)
            ttk.Button(lrow, text="浏览…",
                       command=self.pick_cover_logo).pack(side="left", padx=(8, 0))
            ttk.Label(lrow, text="默认不加校徽/院徽，需要时自己选一张",
                      style="Muted.TLabel").pack(side="left", padx=(10, 0))

            # ---------------- 输出 ----------------
            out = self._section(f, "输出")
            orow = tk.Frame(out, bg="white")
            orow.pack(fill="x", pady=(0, 8))
            ttk.Label(orow, text="输出目录", style="Field.TLabel").pack(side="left", padx=(0, 8))
            self.e_ch_out = ttk.Entry(orow, style="Field.TEntry")
            self.e_ch_out.insert(0, default_outdir())
            self.e_ch_out.pack(side="left", fill="x", expand=True)
            ttk.Button(orow, text="浏览…",
                       command=lambda: self.pick_dir(self.e_ch_out)).pack(side="left", padx=(8, 0))

            row3 = tk.Frame(out, bg="white")
            row3.pack(fill="x", pady=(0, 10))
            self.v_pdf2 = tk.BooleanVar(value=True)
            self.v_open2 = tk.BooleanVar(value=True)
            Toggle(row3, "顺便编译 PDF", self.v_pdf2).pack(side="left")
            Toggle(row3, "完成后打开输出文件夹", self.v_open2).pack(side="left", padx=24)
            ttk.Label(row3, text="点右下角「生成大报告」开始",
                      style="Muted.TLabel").pack(side="right")

        def _tab_chapters(self):
            sf = ScrollFrame(self.nb)
            self.nb.add(sf, text="②  合并四章")
            f = sf.inner
            f.configure(padx=14, pady=10)
            f.columnconfigure(0, weight=1)
            f.columnconfigure(1, minsize=470, weight=0)
            f.rowconfigure(0, weight=1)

            # ================= 左：稿件与分章 =================
            left = tk.Frame(f, bg="white")
            left.grid(row=0, column=0, sticky="nsew", padx=(0, 18))
            left.columnconfigure(0, weight=1)
            left.rowconfigure(2, weight=1)

            head = tk.Frame(left, bg="white")
            head.grid(row=0, column=0, sticky="we", pady=(0, 8))
            tk.Frame(head, bg=ACCENT, width=3, height=15).pack(side="left", padx=(0, 9))
            tk.Label(head, text="稿件与分章", bg="white", fg=INK,
                     font=(FONT, 12, "bold")).pack(side="left")
            tk.Label(head, text="选中行分配章节　·　双击可改一级标题", bg="white",
                     fg=SUB_TXT, font=(FONT, 9)).pack(side="left", padx=(10, 0))

            wrap = tk.Frame(left, bg=LINE, padx=1, pady=1)
            wrap.grid(row=2, column=0, sticky="nsew")
            inner = tk.Frame(wrap, bg="white")
            inner.pack(fill="both", expand=True)
            cols = ("idx", "name", "chapter", "title")
            self.tv_ch = ttk.Treeview(inner, columns=cols, show="headings")
            self.tv_ch.heading("idx", text="#")
            self.tv_ch.heading("name", text="文件名")
            self.tv_ch.heading("chapter", text="章节")
            self.tv_ch.heading("title", text="一级标题（双击可改）")
            self.tv_ch.column("idx", width=44, anchor="center")
            self.tv_ch.column("name", width=280, anchor="w")
            self.tv_ch.column("chapter", width=106, anchor="center")
            self.tv_ch.column("title", width=330, anchor="w")
            self.tv_ch.tag_configure("odd", background="#FCFCFD")
            s2 = ttk.Scrollbar(inner, command=self.tv_ch.yview,
                               style="Slim.Vertical.TScrollbar")
            self.tv_ch.configure(yscrollcommand=s2.set)
            s2.pack(side="right", fill="y")
            self.tv_ch.pack(side="left", fill="both", expand=True)
            self.tv_ch.bind("<Double-1>", self._edit_section_title)
            self.hint_ch = tk.Label(inner, text="还没有添加稿件　·　添加后点「自动分章」，再核对章节",
                                    bg="white", fg="#98A2B3", font=(FONT, 9))
            self.hint_ch.place(relx=0.5, rely=0.5, anchor="center")

            bar = tk.Frame(left, bg="white")
            bar.grid(row=3, column=0, sticky="we", pady=(10, 0))
            ttk.Button(bar, text="添加文件…", style="Soft.TButton",
                       command=self.add_chapter_files).pack(side="left")
            ttk.Button(bar, text="添加整个文件夹…",
                       command=self.add_chapter_folder).pack(side="left", padx=8)
            ttk.Button(bar, text="移除所选", style="Ghost.TButton",
                       command=self.remove_chapter).pack(side="left")
            ttk.Button(bar, text="自动分章", style="Soft.TButton",
                       command=self.auto_assign).pack(side="left", padx=(20, 0))
            ttk.Label(bar, text="可先猜后核", style="Muted.TLabel"
                      ).pack(side="left", padx=8)

            bar2 = tk.Frame(left, bg="white")
            bar2.grid(row=4, column=0, sticky="we", pady=(8, 0))
            ttk.Label(bar2, text="选中行分配到", style="Field.TLabel").pack(side="left", padx=(0, 8))
            for i, name in enumerate(CHAPTER_NAMES):
                ttk.Button(bar2, text=name, style="Soft.TButton",
                           command=lambda k=i: self.assign_chapter(k)).pack(side="left", padx=(0, 6))
            ttk.Button(bar2, text="不收录", style="Ghost.TButton",
                       command=lambda: self.assign_chapter(-1)).pack(side="left", padx=(2, 0))

            # ================= 右：封面 + 输出 =================
            right = tk.Frame(f, bg="white", width=470)
            right.grid(row=0, column=1, sticky="nsew")
            right.grid_propagate(False)
            right.columnconfigure(0, weight=1)

            h2 = tk.Frame(right, bg="white")
            h2.grid(row=0, column=0, sticky="we", pady=(0, 8))
            tk.Frame(h2, bg=ACCENT, width=3, height=15).pack(side="left", padx=(0, 9))
            tk.Label(h2, text="封面信息", bg="white", fg=INK,
                     font=(FONT, 12, "bold")).pack(side="left")
            tk.Label(h2, text="组员用空格分隔", bg="white", fg=SUB_TXT,
                     font=(FONT, 9)).pack(side="left", padx=(10, 0))

            form = tk.Frame(right, bg="white")
            form.grid(row=1, column=0, sticky="nsew")
            form.columnconfigure(1, weight=1)

            def make_row(label, value, row, buttons=None):
                ttk.Label(form, text=label, style="Field.TLabel").grid(
                    row=row, column=0, sticky="w", pady=3, padx=(0, 12))
                entry = ttk.Entry(form, style="Field.TEntry")
                entry.insert(0, value)
                if buttons:
                    holder = tk.Frame(form, bg="white")
                    holder.grid(row=row, column=1, sticky="we", pady=4)
                    entry.pack(in_=holder, side="left", fill="x", expand=True)
                    for text, cmd, style in buttons:
                        ttk.Button(holder, text=text, style=style, command=cmd).pack(
                            side="left", padx=(8, 0))
                else:
                    entry.grid(row=row, column=1, sticky="we", pady=3)
                return entry

            self.e_cov_line1 = make_row("首行", "2023级11组PBL学习", 0)
            self.e_cov_line2 = make_row("次行", "学习报告", 1)
            self.e_cov_case = make_row("Case（作者填）", "Case 1：难缠的脖子痛", 2)
            self.e_cov_tutor = make_row("Tutor", "马凯歌", 3)
            self.e_cov_chair = make_row("主席", "朱俊铭", 4)
            self.e_cov_sec = make_row("书记", "郭育恒", 5)
            self.e_cov_date = make_row("日期", time.strftime("%Y 年 %m 月"), 6)
            self.e_cov_members = make_row(
                "组员", "叶瑞林 张欣仪 韩曦 任欣萌 王思源 于子轩 牛海华 韩畅", 7)
            ttk.Label(form, text="封面图片", style="Field.TLabel").grid(
                row=8, column=0, sticky="w", pady=4, padx=(0, 12))
            logo_row = tk.Frame(form, bg="white")
            logo_row.grid(row=8, column=1, sticky="we", pady=4)
            self.e_cov_logo = ttk.Entry(logo_row, style="Field.TEntry")
            self.e_cov_logo.pack(side="left", fill="x", expand=True)
            self.e_cov_logo.insert(0, "")
            ttk.Button(logo_row, text="浏览…", style="Soft.TButton", width=7,
                       command=self.pick_cover_logo).pack(side="left", padx=(8, 0))
            ttk.Label(form, text="默认留空＝不加校徽 / 院徽", style="Muted.TLabel").grid(
                row=9, column=1, sticky="w", pady=(0, 2))

            tk.Frame(right, bg=LINE, height=1).grid(row=2, column=0, sticky="we", pady=(16, 12))

            h3 = tk.Frame(right, bg="white")
            h3.grid(row=3, column=0, sticky="we", pady=(0, 8))
            tk.Frame(h3, bg=ACCENT, width=3, height=15).pack(side="left", padx=(0, 9))
            tk.Label(h3, text="输出", bg="white", fg=INK,
                     font=(FONT, 12, "bold")).pack(side="left")

            orow = tk.Frame(right, bg="white")
            orow.grid(row=4, column=0, sticky="we")
            orow.columnconfigure(1, weight=1)
            ttk.Label(orow, text="目录", style="Field.TLabel").grid(row=0, column=0, sticky="w")
            self.e_ch_out = ttk.Entry(orow, style="Field.TEntry")
            self.e_ch_out.insert(0, default_outdir())
            self.e_ch_out.grid(row=0, column=1, sticky="we", padx=(10, 8))
            ttk.Button(orow, text="浏览…",
                       command=lambda: self.pick_dir(self.e_ch_out)).grid(row=0, column=2, sticky="e")

            row3 = tk.Frame(right, bg="white")
            row3.grid(row=5, column=0, sticky="we", pady=(12, 0))
            self.v_pdf2 = tk.BooleanVar(value=True)
            self.v_open2 = tk.BooleanVar(value=True)
            Toggle(row3, "顺便编译 PDF", self.v_pdf2).pack(side="left")
            Toggle(row3, "完成后打开文件夹", self.v_open2).pack(side="left", padx=20)

        def _tab_env(self):
            f = tk.Frame(self.nb, bg="white", padx=2, pady=2)
            self.nb.add(f, text="③  环境与说明")
            body = self._section(f, "环境与使用说明", "下面的信息自动生成，随时可用「重新检测环境」刷新")
            wrap, inner = self._bordered(body, bg="#FBFCFD")
            wrap.pack(fill="both", expand=True)
            self.env_text = tk.Text(inner, wrap="word", font=(FONT, 9), height=16, width=70,
                                    bg="#FBFCFD", fg=INK, relief="flat", padx=14, pady=12,
                                    borderwidth=0, highlightthickness=0, spacing1=2, spacing3=3)
            env_sb = ttk.Scrollbar(inner, command=self.env_text.yview,
                                   style="Slim.Vertical.TScrollbar")
            self.env_text.configure(yscrollcommand=env_sb.set)
            env_sb.pack(side="right", fill="y")
            self.env_text.pack(side="left", fill="both", expand=True)

            bar = tk.Frame(f, bg="white")
            bar.pack(fill="x", pady=(10, 8))
            ttk.Button(bar, text="重新检测环境", style="Soft.TButton",
                       command=self.check_env).pack(side="left")
            ttk.Button(bar, text="打开 MiKTeX 下载页",
                       command=self.open_miktex).pack(side="left", padx=8)
            ttk.Button(bar, text="保存日志", command=self.save_log).pack(side="left")
            ttk.Button(bar, text="打开输出文件夹", style="Ghost.TButton",
                       command=self.open_out).pack(side="left", padx=8)
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
                "  · 封面：照 23 级 Case 报告样式（“2023级××组PBL学习 / 学习报告 / Case：…”",
                "    + Tutor / PBL小组主席 / 书记 / 成员）；封面图片默认留空（不放校徽，避免版权问题），",
                "    需要时自己选一张；字号：首行一号、次行小一、Case 小初、小组构成三号（1.5 倍行距）。",
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
                    tag = "odd" if len(self.single_files) % 2 == 0 else ""
                    self.tv_single.insert("", "end", iid=f, values=(f, "待转换"),
                                          tags=(tag,) if tag else ())
            self._update_hints()

        def add_single_folder(self):
            d = filedialog.askdirectory(title="选择存放 Word 稿件的文件夹")
            if not d:
                return
            for name in sorted(os.listdir(d)):
                if name.lower().endswith((".docx", ".pdf")) and not name.startswith("~$"):
                    p = os.path.join(d, name)
                    if p not in self.single_files:
                        self.single_files.append(p)
                        tag = "odd" if len(self.single_files) % 2 == 0 else ""
                        self.tv_single.insert("", "end", iid=p, values=(p, "待转换"),
                                              tags=(tag,) if tag else ())
            self._update_hints()

        def remove_single(self):
            for item in self.tv_single.selection():
                vals = self.tv_single.item(item, "values")
                if vals and vals[0] in self.single_files:
                    self.single_files.remove(vals[0])
                self.tv_single.delete(item)
            self._update_hints()

        def clear_single(self):
            self.single_files.clear()
            for item in self.tv_single.get_children():
                self.tv_single.delete(item)
            self._update_hints()

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

        def _update_hints(self):
            try:
                if self.single_files:
                    self.hint_single.place_forget()
                else:
                    self.hint_single.place(relx=0.5, rely=0.58, anchor="center")
                if self.chapter_items:
                    self.hint_ch.place_forget()
                else:
                    self.hint_ch.place(relx=0.5, rely=0.58, anchor="center")
            except Exception:  # noqa: BLE001
                pass

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
                tag = "odd" if i % 2 == 1 else ""
                self.tv_ch.insert("", "end", iid=str(i),
                                  values=(i + 1, os.path.basename(path), chap_name, shown),
                                  tags=(tag,) if tag else ())
            self._update_hints()
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
                w = max(min(1360, sw - 60), 1120)
                h = max(min(940, sh - 90), 680)
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
































