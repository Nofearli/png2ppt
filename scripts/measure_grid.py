# -*- coding: utf-8 -*-
"""坐标网格放大图：量原图元素像素坐标用。
用法:
  python measure_grid.py original.png                      # 全图
  python measure_grid.py original.png --crop 550,160,720,320 --scale 5   # 局部放大
输出: <原名>.grid.png（或 <原名>.<区域>.grid.png）
"""
import argparse, os
from PIL import Image, ImageDraw

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("image")
    ap.add_argument("--crop", help="x0,y0,x1,y1（原图px）")
    ap.add_argument("--scale", type=int, default=4, help="放大倍数（默认4）")
    ap.add_argument("--step", type=int, default=10, help="网格间距px（默认10）")
    ap.add_argument("--label-step", type=int, default=20, help="标数间距px（默认20）")
    ap.add_argument("-o", "--out")
    a = ap.parse_args()

    src = Image.open(a.image).convert("RGB")
    W, H = src.size
    if a.crop:
        x0, y0, x1, y1 = map(int, a.crop.split(","))
        x0, y0 = max(0, x0), max(0, y0)
        x1, y1 = min(W, x1), min(H, y1)
        base = f"{os.path.splitext(a.image)[0]}.{x0}_{y0}"
    else:
        x0, y0, x1, y1 = 0, 0, W, H
        base = os.path.splitext(a.image)[0]

    c = src.crop((x0, y0, x1, y1)).resize(((x1 - x0) * a.scale, (y1 - y0) * a.scale), Image.LANCZOS)
    d = ImageDraw.Draw(c)
    for gx in range(x0 - x0 % a.step + a.step, x1, a.step):
        px = (gx - x0) * a.scale
        major = gx % a.label_step == 0
        d.line([(px, 0), (px, c.height)], fill=(255, 0, 255) if major else (255, 160, 255), width=1)
        if major:
            d.text((px + 1, 1), str(gx), fill=(180, 0, 180))
    for gy in range(y0 - y0 % a.step + a.step, y1, a.step):
        py = (gy - y0) * a.scale
        major = gy % a.label_step == 0
        d.line([(0, py), (c.width, py)], fill=(0, 180, 255) if major else (150, 220, 255), width=1)
        if major:
            d.text((1, py + 1), str(gy), fill=(0, 110, 200))
    # 边缘刻度
    for gx in range(x0, x1, a.label_step):
        px = (gx - x0) * a.scale
        d.line([(px, c.height - 6), (px, c.height)], fill=(200, 0, 200), width=1)
    for gy in range(y0, y1, a.label_step):
        py = (gy - y0) * a.scale
        d.line([(c.width - 6, py), (c.width, py)], fill=(200, 0, 200), width=1)

    out = a.out or f"{base}.grid.png"
    c.save(out)
    print(f"saved {out}  ({c.width}x{c.height}, region=({x0},{y0})-({x1},{y1}), scale={a.scale})")

if __name__ == "__main__":
    main()
