# -*- coding: utf-8 -*-
"""取色：打印原图指定坐标处的颜色。
用法: python sample_colors.py original.png 300,55 450,60 1300,108
输出: 位置  #RRGGBB
"""
import argparse
from PIL import Image

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("image")
    ap.add_argument("points", nargs="+", help="x,y 像素坐标，可多个")
    a = ap.parse_args()
    im = Image.open(a.image).convert("RGB")
    W, H = im.size
    for pt in a.points:
        x, y = map(int, pt.split(","))
        if not (0 <= x < W and 0 <= y < H):
            print(f"({x},{y})  超出图像范围 ({W}x{H})")
            continue
        r, g, b = im.getpixel((x, y))
        print(f"({x:4d},{y:4d})  #{r:02X}{g:02X}{b:02X}")

if __name__ == "__main__":
    main()
