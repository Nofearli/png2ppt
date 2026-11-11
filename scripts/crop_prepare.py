# -*- coding: utf-8 -*-
"""素材准备：背景修补（inpaint）+ 裁剪 + 白底转透明 + 检查拼图。JSON 配置驱动。

配置示例（prep.json）:
{
  "src": "original.png",
  "out_dir": "assets/icons",
  "patched": "assets/patched_full.png",
  "patches": [
    {"box": [78, 94, 126, 136], "mode": "h"},
    {"box": [696, 172, 714, 236], "mode": "v", "on": "clean"}
  ],
  "crops": {
    "chip_zone.png": {"box": [1222, 103, 1340, 182], "source": "clean"},
    "pin_e1.png":    {"box": [278, 169, 310, 209],  "source": "keep"}
  },
  "alpha": ["pin_e1.png"],
  "alpha_range": [208, 242],
  "blur": 1.2,
  "contact_sheet": {"height": 190}
}

用法: python crop_prepare.py prep.json

要点:
- patches: mode "h"=按行左右线性插值（平滑背景上的文字），"v"=按列上下插值（贴边残字）
- "on": "both"(默认,两种版本都修) | "clean"(仅 clean 版) —— 用于"去针体保徽章"这类区分
- crops.source: "clean"=修补后版本, "keep"=保守修补版本, "original"=原图
- alpha: 白底转透明（lo/hi 阈值默认 208/242）
- 结束后生成 contact_sheet.png，必须打开检查每个图标
"""
import argparse, json, os
from PIL import Image, ImageFilter, ImageDraw
import numpy as np


def inpaint_h(arr, x0, y0, x1, y1, pad=6):
    ys = slice(y0, y1)
    left = arr[ys, max(x0 - pad, 0):x0].mean(axis=1, keepdims=True)
    right = arr[ys, x1:x1 + pad].mean(axis=1, keepdims=True)
    w = np.linspace(0, 1, x1 - x0)[None, :, None]
    arr[ys, x0:x1] = left * (1 - w) + right * w
    return arr


def inpaint_v(arr, x0, y0, x1, y1, pad=8):
    xs = slice(x0, x1)
    top = arr[max(y0 - pad, 0):y0, xs].mean(axis=0, keepdims=True)
    bottom = arr[y1:y1 + pad, xs].mean(axis=0, keepdims=True)
    w = np.linspace(0, 1, y1 - y0)[:, None, None]
    arr[y0:y1, xs] = top * (1 - w) + bottom * w
    return arr


def build_version(a0, patches, blur):
    arr = a0.copy()
    for p in patches:
        x0, y0, x1, y1 = p["box"]
        arr = inpaint_v(arr, x0, y0, x1, y1) if p.get("mode") == "v" else inpaint_h(arr, x0, y0, x1, y1)
    img = Image.fromarray(arr.clip(0, 255).astype(np.uint8))
    blurred = img.filter(ImageFilter.GaussianBlur(blur))
    mask = Image.new("L", img.size, 0)
    md = ImageDraw.Draw(mask)
    for p in patches:
        md.rectangle(p["box"], fill=255)
    return Image.composite(blurred, img, mask)


def white_to_alpha(img, lo=208, hi=242):
    im = img.convert("RGBA")
    arr = np.asarray(im).astype(float)
    mn = arr[..., :3].min(axis=-1)
    t = np.clip((mn - lo) / (hi - lo), 0, 1)
    arr[..., 3] = np.minimum(arr[..., 3], (1 - t) * 255)
    return Image.fromarray(arr.clip(0, 255).astype(np.uint8), "RGBA")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("config", help="JSON 配置文件路径")
    a = ap.parse_args()
    cfg = json.load(open(a.config, encoding="utf-8"))

    src = Image.open(cfg["src"]).convert("RGB")
    a0 = np.asarray(src).astype(float)
    out_dir = cfg.get("out_dir", "assets/icons")
    os.makedirs(out_dir, exist_ok=True)
    blur = cfg.get("blur", 1.2)

    patches = cfg.get("patches", [])
    base = [p for p in patches if p.get("on", "both") in ("both", "keep")]
    extra = [p for p in patches if p.get("on") == "clean"]

    img_keep = build_version(a0, base, blur)
    img_clean = build_version(a0, base + extra, blur) if extra else img_keep
    if cfg.get("patched"):
        img_keep.save(cfg["patched"])

    srcs = {"clean": img_clean, "keep": img_keep, "original": src}
    names = []
    for name, spec in cfg.get("crops", {}).items():
        x0, y0, x1, y1 = spec["box"]
        srcs.get(spec.get("source", "clean"), img_clean).crop((x0, y0, x1, y1)).save(os.path.join(out_dir, name))
        names.append(name)

    lo, hi = cfg.get("alpha_range", [208, 242])
    for name in cfg.get("alpha", []):
        p = os.path.join(out_dir, name)
        if os.path.exists(p):
            white_to_alpha(Image.open(p), lo, hi).save(p)

    if cfg.get("contact_sheet", {}).get("height") and names:
        H = cfg["contact_sheet"]["height"]
        bg = tuple(cfg["contact_sheet"].get("bg", [190, 196, 203]))
        tiles = [Image.open(os.path.join(out_dir, n)) for n in names]
        scaled = [t.resize((max(1, int(t.width * H / t.height)), H)) for t in tiles]
        W = sum(s.width for s in scaled) + 10 * len(scaled)
        sheet = Image.new("RGB", (W, H + 22), bg)
        d = ImageDraw.Draw(sheet)
        x = 0
        for n, s_ in zip(names, scaled):
            sheet.paste(s_, (x, 20), s_ if s_.mode == "RGBA" else None)
            d.text((x + 2, 4), n, fill=(0, 0, 0))
            x += s_.width + 10
        sheet_path = cfg["contact_sheet"].get("out", os.path.join(out_dir, "..", "contact_sheet.png"))
        sheet.save(sheet_path)
        print("contact sheet:", sheet_path)

    print(f"done: {len(names)} crops -> {out_dir}")


if __name__ == "__main__":
    main()
