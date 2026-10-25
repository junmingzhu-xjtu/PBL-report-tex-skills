# -*- coding: utf-8 -*-
"""
dump_headings.py —— 快速列出每个 docx 的标题结构，便于分章（重大问题/次要问题）

用法：
    python dump_headings.py 提交目录
    python dump_headings.py a.docx b.docx
"""
import glob
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import docx2tex as d2t
from docx import Document


def dump(path):
    doc = Document(path)
    rows = []
    for par in doc.paragraphs:
        lvl = d2t.heading_level(par)
        if lvl:
            rows.append((lvl, (par.text or "").strip()))
    print("\n=== %s （%d 个标题）" % (os.path.basename(path), len(rows)))
    for lvl, txt in rows[:40]:
        print("   " + "  " * (lvl - 1) + ("%d. " % lvl) + txt[:60])
    if len(rows) > 40:
        print("   ...（共 %d 个）" % len(rows))
    return rows


def main():
    args = sys.argv[1:]
    files = []
    for a in args:
        if os.path.isdir(a):
            files += sorted(glob.glob(os.path.join(a, "*.docx")))
        else:
            files.append(a)
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    for f in files:
        try:
            dump(f)
        except Exception as e:
            print("[ERR] %s: %s" % (f, e))


if __name__ == "__main__":
    main()
