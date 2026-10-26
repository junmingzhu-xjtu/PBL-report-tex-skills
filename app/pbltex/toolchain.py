# -*- coding: utf-8 -*-
"""pbltex.toolchain —— 查找 xelatex 并编译 PDF"""
from __future__ import annotations

import os
import re
import shutil
import subprocess

CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0
_ENV_CACHE = {}


def _candidate_paths():
    env = os.environ
    roots = []
    local = env.get("LOCALAPPDATA") or ""
    if local:
        roots.append(os.path.join(local, "Programs", "MiKTeX", "miktex", "bin"))
    for base in (env.get("ProgramFiles"), env.get("ProgramFiles(x86)")):
        if base:
            roots.append(os.path.join(base, "MiKTeX", "miktex", "bin"))
    roots.append(r"C:\texlive")
    for year in ("2026", "2025", "2024", "2023", "2022"):
        roots.append(os.path.join(r"C:\texlive", year, "bin"))
        roots.append(os.path.join(env.get("ProgramFiles", r"C:\Program Files"), "texlive", year, "bin"))
    out = []
    for root in roots:
        if not root or not os.path.isdir(root):
            continue
        for sub in ("x64", "windows", "win32", ""):
            d = os.path.join(root, sub) if sub else root
            if not os.path.isdir(d):
                continue
            try:
                for entry in os.listdir(d):
                    if entry.lower() in ("xelatex.exe", "xelatex") and not os.path.isdir(os.path.join(d, entry)):
                        out.append(os.path.join(d, entry))
                    elif sub == "" and os.path.isdir(os.path.join(d, entry)) and entry.lower().startswith("win"):
                        inner = os.path.join(d, entry, "xelatex.exe")
                        if os.path.isfile(inner):
                            out.append(inner)
            except OSError:
                continue
    return out


def find_xelatex() -> str:
    """返回 xelatex.exe 路径，找不到返回空串。"""
    if "xelatex" in _ENV_CACHE:
        return _ENV_CACHE["xelatex"]
    p = shutil.which("xelatex")
    if p:
        _ENV_CACHE["xelatex"] = p
        return p
    for cand in _candidate_paths():
        if os.path.isfile(cand):
            _ENV_CACHE["xelatex"] = cand
            return cand
    _ENV_CACHE["xelatex"] = ""
    return ""


def reset_cache():
    _ENV_CACHE.clear()


def xelatex_version(path=None) -> str:
    path = path or find_xelatex()
    if not path:
        return ""
    try:
        r = subprocess.run([path, "--version"], capture_output=True, timeout=30,
                           creationflags=CREATE_NO_WINDOW)
        txt = (r.stdout or b"").decode("utf-8", "replace")
        first = txt.splitlines()[0].strip() if txt else ""
        return first
    except Exception:  # noqa: BLE001
        return ""


def _is_miktex() -> bool:
    if "miktex" in _ENV_CACHE:
        return _ENV_CACHE["miktex"]
    v = xelatex_version().lower()
    is_mk = ("miktex" in v) or ("miktex" in os.environ.get("PATH", "").lower())
    _ENV_CACHE["miktex"] = is_mk
    return is_mk


def _needs_rerun(tex_path) -> bool:
    """看 .log 里是否还提示要再跑一遍（目录页码、交叉引用）。"""
    logf = os.path.splitext(tex_path)[0] + ".log"
    try:
        with open(logf, "r", encoding="utf-8", errors="replace") as fh:
            txt = fh.read()
    except OSError:
        return False
    return ("Rerun to get" in txt) or ("Label(s) may have changed" in txt)


def _log_errors(tex_path, limit=8):
    """从 xelatex 的 .log 里抠报错（比抓 stdout 可靠）。"""
    logf = os.path.splitext(tex_path)[0] + ".log"
    try:
        with open(logf, "r", encoding="utf-8", errors="replace") as fh:
            lines = fh.readlines()
    except OSError:
        return []
    out = []
    for i, line in enumerate(lines):
        s = line.rstrip("\n")
        low = s.lower()
        if s.startswith("!") or "emergency stop" in low or ("not found" in low and (".sty" in low or ".cls" in low or "font" in low)):
            out.append(s.strip())
            for nxt in lines[i + 1:i + 4]:
                t = nxt.rstrip("\n")
                if t.startswith("l."):
                    out.append(t.strip())
                    break
        if len(out) >= limit:
            break
    if not out:                       # 没有以 ! 开头的，就把最后几行有用的捞出来
        for line in lines[-40:]:
            s = line.rstrip("\n")
            if s.strip() and not s.startswith(("(", ")", "[")):
                out.append(s.strip()[:160])
        out = out[-6:]
    return out


def compile_tex(tex_path, log=None, passes=2, max_passes=3, timeout=300) -> str:
    """编译 tex，返回 pdf 路径；失败返回空串。

    至少跑 `passes` 遍；若 .log 仍提示 Rerun（目录页码没稳定），最多补到 max_passes 遍。
    正文引用是直接写的上标编号，不依赖两遍解析，所以只跑一遍也不会出现 [?,?,?]。
    只要 PDF 比 .tex 新，就算 xelatex 返回码非 0 也认为成功（MiKTeX 偶尔这么抽风）。
    """
    log = log or (lambda *_: None)
    exe = find_xelatex()
    if not exe:
        log("  [跳过 PDF] 未检测到 xelatex。可先安装 MiKTeX（见“环境与说明”标签页）。")
        return ""

    workdir = os.path.dirname(os.path.abspath(tex_path)) or "."
    name = os.path.basename(tex_path)
    pdf = os.path.splitext(tex_path)[0] + ".pdf"
    try:
        tex_mtime = os.path.getmtime(tex_path)
    except OSError:
        tex_mtime = 0
    if os.path.exists(pdf):
        try:
            os.remove(pdf)
        except OSError as exc:
            log("  [警告] 删不掉旧 PDF（%s）——若编译报错，请换个输出目录" % exc.strerror)

    base_cmd = [exe, "-interaction=nonstopmode", "-halt-on-error"]
    if _is_miktex():
        base_cmd.append("--enable-installer")   # 允许 MiKTeX 自动补装宏包

    ok_tail = ""
    last_ok = False
    n = 0
    while n < max_passes:
        n += 1
        log("  xelatex 第 %d 遍…" % n)
        try:
            r = subprocess.run(base_cmd + [name], cwd=workdir, capture_output=True,
                               timeout=timeout, creationflags=CREATE_NO_WINDOW)
        except subprocess.TimeoutExpired:
            log("  [失败] xelatex 超时（>%d 秒）" % timeout)
            break
        except Exception as exc:  # noqa: BLE001
            log("  [失败] 无法运行 xelatex：%s" % exc)
            break
        out = (r.stdout or b"").decode("utf-8", "replace") + (r.stderr or b"").decode("utf-8", "replace")
        ok_tail = out
        last_ok = (r.returncode == 0)
        if n >= passes and not _needs_rerun(tex_path):
            break
        if n >= max_passes:
            break

    # 只要 PDF 比 .tex 新，就算成功（MiKTeX 偶尔返回码非 0，但 PDF 是好的）
    pdf_fresh = False
    if os.path.exists(pdf):
        try:
            pdf_fresh = os.path.getmtime(pdf) >= tex_mtime - 1
        except OSError:
            pdf_fresh = True

    if pdf_fresh:
        if not last_ok:
            log("  [提示] xelatex 返回码非 0，但 PDF 已生成，按成功处理")
        if _needs_rerun(tex_path):
            log("  [提示] 目录页码可能还没稳定，建议再编译一次")
        log("  PDF 已生成：%s" % os.path.basename(pdf))
        return pdf

    # 失败：优先从 .log 里抠报错（比抓 stdout 可靠）
    errs = _log_errors(tex_path) or []
    if not errs:
        for line in ok_tail.splitlines():
            s = line.strip()
            if s.startswith("!") or "not found" in s.lower() or "Emergency stop" in s:
                errs.append(s)
            if len(errs) >= 6:
                break
    if errs:
        log("  [失败] xelatex 报错：")
        for e in errs[:8]:
            log("     " + e[:170])
    else:
        log("  [失败] 未生成 PDF；若输出目录在本程序文件夹里，可能没有写权限——"
            "换个输出目录（如桌面）再试")
    return ""


def pdf_info(pdf_path) -> str:
    """粗查 PDF 页数（数 /Type /Page）。"""
    try:
        with open(pdf_path, "rb") as fh:
            data = fh.read()
        n = len(re.findall(rb"/Type\s*/Page[^s]", data))
        return "%d 页" % n if n else ""
    except OSError:
        return ""
