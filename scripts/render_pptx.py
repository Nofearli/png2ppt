# -*- coding: utf-8 -*-
"""PowerPoint COM 渲染：把 pptx 第 1 页导出为 PNG 预览。
用法: python render_pptx.py <in.pptx> <out.png> [宽px] [高px]
宽高建议 = 原图像素 × 2。需要 Windows + PowerPoint + pywin32。"""
import sys, os

def render(pptx_path, out_png, width=None, height=None):
    import win32com.client
    pptx_path = os.path.abspath(pptx_path)
    out_png = os.path.abspath(out_png)
    pp = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = pp.Presentations.Open(pptx_path, ReadOnly=True, WithWindow=False)
        try:
            if width is None:
                pres.Slides(1).Export(out_png, "PNG")
            else:
                pres.Slides(1).Export(out_png, "PNG", str(width), str(height))
        finally:
            pres.Close()
    finally:
        try:
            if pp.Presentations.Count == 0:
                pp.Quit()
        except Exception:
            pass
    print("exported:", out_png)

if __name__ == "__main__":
    p = sys.argv[1]; o = sys.argv[2]
    w = int(sys.argv[3]) if len(sys.argv) > 3 else None
    h = int(sys.argv[4]) if len(sys.argv) > 4 else None
    render(p, o, w, h)
