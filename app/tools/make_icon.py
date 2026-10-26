# -*- coding: utf-8 -*-
"""生成应用图标 app.ico（无需外部素材）"""
import os

from PIL import Image, ImageDraw, ImageFont

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SIZE = 512


def rounded(draw, box, r, fill):
    draw.rounded_rectangle(box, radius=r, fill=fill)


def make():
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # 背景渐变
    grad = Image.new("RGB", (SIZE, SIZE))
    gd = ImageDraw.Draw(grad)
    top = (46, 116, 235)
    bottom = (12, 52, 130)
    for y in range(SIZE):
        t = y / (SIZE - 1)
        gd.line([(0, y), (SIZE, y)], fill=(
            int(top[0] + (bottom[0] - top[0]) * t),
            int(top[1] + (bottom[1] - top[1]) * t),
            int(top[2] + (bottom[2] - top[2]) * t)))
    mask = Image.new("L", (SIZE, SIZE), 0)
    md = ImageDraw.Draw(mask)
    rounded(md, (6, 6, SIZE - 6, SIZE - 6), 96, 255)
    img.paste(grad, (0, 0), mask)

    # 纸张图形
    d.rounded_rectangle((140, 96, 372, 416), radius=18, fill=(255, 255, 255, 235))
    for i, y in enumerate(range(160, 300, 34)):
        w = 170 if i != 3 else 110
        d.rounded_rectangle((172, y, 172 + w, y + 12), radius=6, fill=(190, 205, 226, 255))
    d.rounded_rectangle((172, 330, 262, 342), radius=6, fill=(31, 111, 235, 255))
    d.rounded_rectangle((276, 330, 340, 342), radius=6, fill=(31, 111, 235, 255))

    # 角标 “TeX” 变小，用简洁几何：右上角一个小方块
    d.rounded_rectangle((330, 60, 452, 182), radius=26, fill=(255, 209, 102, 255))
    try:
        font = ImageFont.truetype(r"C:\Windows\Fonts\msyhbd.ttc", 84)
        d.text((352, 78), "T", font=font, fill=(60, 45, 10, 255))
    except OSError:
        pass

    out = os.path.join(BASE, "app.ico")
    img.save(out, sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
    png = os.path.join(BASE, "app_icon.png")
    img.resize((256, 256), Image.LANCZOS).save(png)
    print("written:", out, os.path.getsize(out), "bytes")


if __name__ == "__main__":
    make()
