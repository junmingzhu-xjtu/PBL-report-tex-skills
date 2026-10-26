# -*- coding: utf-8 -*-
"""pbltex —— Word 稿件 → LaTeX/PDF 转换核心（PBL 大报告生成器）。"""
from .core import (  # noqa: F401
    CHAPTER_NAMES,
    convert_chapters,
    convert_single,
    guess_chapter,
)
from .toolchain import (  # noqa: F401
    compile_tex,
    find_xelatex,
    pdf_info,
    xelatex_version,
)

__version__ = "1.0.0"
