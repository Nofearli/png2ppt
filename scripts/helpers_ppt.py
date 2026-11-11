# -*- coding: utf-8 -*-
"""
helpers_ppt.py —— png2ppt 通用助手库
坐标惯例：调用方以原图像素 / 100 = 英寸传参（In(x) 吃英寸）。
调色板：set_palette({...}) 换成本项目采样色；rgb("key") 或 rgb("#RRGGBB") 均可。
"""
import math
from pptx.util import Inches as In, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn

FONT = "Times New Roman"

# ---------- 调色板（示例值；项目开始时用 sample_colors.py 采样后 set_palette 替换） ----------
DEFAULT_PALETTE = dict(
    dark="2B2B2B", navy_text="1F3D5C",
    blue_label="1F4E9C", red_label="C23B35", pin="F14550", glow="F8C9CD",
    amber="FDE07B", blue_row="D0E8FE",
    teal="447354", teal_dark="3E654B", gray="404040",
    star="409C1C", blob_blue="3B6BC7", blob_red="E45C5C", dot="E8A33D", dot_border="C77F1F",
    grid="B7B7B7", white="FFFFFF",
)
_PALETTE = dict(DEFAULT_PALETTE)

def set_palette(d):
    """合并替换调色板（保留未覆盖的默认键）"""
    _PALETTE.update(d)

def reset_palette():
    _PALETTE.clear(); _PALETTE.update(DEFAULT_PALETTE)

def rgb(key_or_hex):
    return RGBColor.from_string(_PALETTE.get(key_or_hex, key_or_hex))

def _ln(shape):
    return shape._element.spPr.get_or_add_ln()

# ---------- 文本 ----------
def _set_run(r, text, size, color, bold=False, italic=False, kind="n"):
    r.text = text
    f = r.font
    f.name = FONT
    f.size = Pt(size * (0.72 if kind in ("sub", "sup") else 1.0))
    f.bold = bold
    f.italic = italic
    f.color.rgb = rgb(color)
    rPr = f._rPr
    if kind == "sub":
        rPr.set("baseline", "-25000")
    elif kind == "sup":
        rPr.set("baseline", "30000")

def _emit_line(p, line, size, color, bold, italic):
    if isinstance(line, str):
        _set_run(p.add_run(), line, size, color, bold, italic)
    else:
        for item in line:
            if len(item) == 3:
                txt, kind, col = item
            else:
                txt, kind = item
                col = color
            _set_run(p.add_run(), txt, size, col,
                     bold or kind in ("b", "bi"),
                     italic or kind in ("i", "bi", "sub", "sup"),
                     "sub" if kind == "sub" else ("sup" if kind == "sup" else "n"))

def _as_lines(parts):
    if isinstance(parts, str):
        return [parts]
    if isinstance(parts, (list, tuple)):
        if len(parts) > 0 and isinstance(parts[0], tuple):
            return [parts]          # 单行多 run
        return list(parts)          # 多行
    return [parts]

def add_text(shapes, x, y, w, h, parts, size=13, color="dark", align=PP_ALIGN.CENTER,
             bold=False, italic=False, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0):
    """独立文本框（margin=0）。parts: "str" | [(txt,kind),...] 单行 | [行1, 行2, ...] 多行
    kind: n 普通 / i 斜体 / b 粗体 / bi 粗斜体 / sub 下标 / sup 上标；元组第三位可指定颜色"""
    tb = shapes.add_textbox(In(x), In(y), In(w), In(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate(_as_lines(parts)):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        _emit_line(p, line, size, color, bold, italic)
    return tb

def set_shape_text(shape, parts, size=13, color="dark", bold=False, italic=False,
                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0):
    """把富文本写进已有形状（卡片/胶囊/徽章）"""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(27432)
    tf.margin_top = tf.margin_bottom = Emu(13716)
    for i, line in enumerate(_as_lines(parts)):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        _emit_line(p, line, size, color, bold, italic)

# ---------- 形状 ----------
def rrect(shapes, x, y, w, h, fill=None, line=None, line_w=1.0, radius=0.12):
    """圆角矩形；radius=短边比例（0~0.5，胶囊条用 0.5）。fill/line=None 为空"""
    sp = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, In(x), In(y), In(w), In(h))
    try:
        sp.adjustments[0] = radius
    except Exception:
        pass
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = rgb(line); sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp

def rect_fill(shapes, x, y, w, h, fill, line=None, line_w=1.0):
    sp = shapes.add_shape(MSO_SHAPE.RECTANGLE, In(x), In(y), In(w), In(h))
    sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = rgb(line); sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp

def grad_fill(sp, c1, c2, angle=90):
    """双色线性渐变；angle 度（90=从上到下，0=从左到右）"""
    sp.fill.gradient()
    stops = sp.fill.gradient_stops
    stops[0].color.rgb = rgb(c1); stops[0].position = 0.0
    stops[1].color.rgb = rgb(c2); stops[1].position = 1.0
    try:
        sp.fill.gradient_angle = angle
    except Exception:
        pass

def oval(shapes, x, y, w, h, fill=None, line=None, line_w=1.0):
    sp = shapes.add_shape(MSO_SHAPE.OVAL, In(x), In(y), In(w), In(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = rgb(line); sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp

def arrow(shapes, x1, y1, x2, y2, color="teal", w=2.5, head=True, dash=None):
    """直连接线，圆头，可选三角箭头/虚线"""
    conn = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, In(x1), In(y1), In(x2), In(y2))
    conn.shadow.inherit = False
    conn.line.color.rgb = rgb(color); conn.line.width = Pt(w)
    if dash:
        conn.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    ln = _ln(conn)
    ln.set("cap", "rnd")
    if head:
        te = ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"})
        ln.append(te)
    return conn

def plain_line(shapes, x1, y1, x2, y2, color="teal", w=2.5):
    conn = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, In(x1), In(y1), In(x2), In(y2))
    conn.shadow.inherit = False
    conn.line.color.rgb = rgb(color); conn.line.width = Pt(w)
    _ln(conn).set("cap", "rnd")
    return conn

def elbow(shapes, pts, color="teal", w=2.5, head_last=True):
    """折线（逐段直连接线），仅最后一段可带箭头"""
    n = len(pts)
    for i in range(n - 1):
        arrow(shapes, pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1],
              color=color, w=w, head=(head_last and i == n - 2))

def pic(shapes, path, x, y, w=None, h=None):
    kw = {}
    if w: kw["width"] = In(w)
    if h: kw["height"] = In(h)
    return shapes.add_picture(path, In(x), In(y), **kw)

def soft_edge(pic_obj, rad_pt=3.0):
    """图片软边缘羽化（消除裁剪接缝）。rad_pt：羽化半径（磅）"""
    spPr = pic_obj._element.spPr
    eff = spPr.find(qn('a:effectLst'))
    if eff is None:
        eff = spPr.makeelement(qn('a:effectLst'), {})
        spPr.append(eff)
    soft = eff.makeelement(qn('a:softEdge'), {'rad': str(int(rad_pt * 12700))})
    eff.append(soft)

def add_shadow(shape, blur=4.0, dist=2.2, direction=5400000, alpha=22000, color="000000"):
    """柔和外阴影（卡片立体感）。direction: 1/60000 度，5400000=向下"""
    spPr = shape._element.spPr
    eff = spPr.find(qn('a:effectLst'))
    if eff is None:
        eff = spPr.makeelement(qn('a:effectLst'), {})
        spPr.append(eff)
    shdw = eff.makeelement(qn('a:outerShdw'), {
        'blurRad': str(int(blur * 12700)), 'dist': str(int(dist * 12700)),
        'dir': str(direction), 'rotWithShape': '0'})
    clr = shdw.makeelement(qn('a:srgbClr'), {'val': color})
    a = clr.makeelement(qn('a:alpha'), {'val': str(alpha)})
    clr.append(a); shdw.append(clr); eff.append(shdw)

# ---------- 自由曲线（原图手绘感 blob 区域） ----------
def blob_points(cx, cy, rx, ry, n=72, a1=0.10, a2=0.05, p1=0.7, p2=2.1):
    """闭合有机曲线点列（3、5 次谐波扰动圆）。cx,cy,rx,ry 单位英寸"""
    pts = []
    for k in range(n):
        th = 2 * math.pi * k / n
        m = 1 + a1 * math.cos(3 * th + p1) + a2 * math.cos(5 * th + p2)
        pts.append((cx + rx * m * math.cos(th), cy + ry * m * math.sin(th)))
    return pts

def freeform(shapes, pts_in, close=True, fill=None, line=None, line_w=2.5):
    """点列 → 自由形状（英寸）"""
    fb = shapes.build_freeform(Emu(int(pts_in[0][0] * 914400)), Emu(int(pts_in[0][1] * 914400)), scale=1.0)
    rest = [(int(px * 914400), int(py * 914400)) for px, py in pts_in[1:]]
    fb.add_line_segments(rest, close=close)
    sp = fb.convert_to_shape()
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = rgb(line); sp.line.width = Pt(line_w)
    _ln(sp).set("cap", "rnd")
    sp.shadow.inherit = False
    return sp

def blob(shapes, cx, cy, rx, ry, fill, line, line_w=2.5, **kw):
    pts = blob_points(cx, cy, rx, ry, **kw)
    return freeform(shapes, pts, close=True, fill=fill, line=line, line_w=line_w), pts

def blob_arc(shapes, pts, i0, i1, color, line_w=2.5):
    """在 blob 点序列上取弧段画开曲线（局部异色描边用）"""
    n = len(pts)
    seg = []
    k = i0
    while True:
        seg.append(pts[k % n])
        if k % n == i1 % n:
            break
        k += 1
    return freeform(shapes, seg, close=False, fill=None, line=color, line_w=line_w)

# ---------- 表格 ----------
def add_grid_table(shapes, x, y, w, h, rows, cols, col_ws, row_hs):
    """建表并剥离自带样式。col_ws/row_hs 单位英寸（列表）"""
    gf = shapes.add_table(rows, cols, In(x), In(y), In(w), In(h))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    for i, cw in enumerate(col_ws):
        tbl.columns[i].width = Emu(int(cw * 914400))
    for i, rh in enumerate(row_hs):
        tbl.rows[i].height = Emu(int(rh * 914400))
    return tbl

def cell_borders(cell, color="B7B7B7", w=6350):
    """单元格四边细边框（w: EMU，6350=0.5pt）"""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for e in tcPr.findall(qn(tag)):
            tcPr.remove(e)
    for i, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
        ln = tcPr.makeelement(qn(tag), {"w": str(w), "cap": "flat", "cmpd": "sng", "algn": "ctr"})
        sf = ln.makeelement(qn("a:solidFill"), {})
        clr = sf.makeelement(qn("a:srgbClr"), {"val": color})
        sf.append(clr)
        ln.append(sf)
        tcPr.insert(i, ln)

def set_cell(cell, parts, size=12, color="dark", bold=False, italic=False,
             align=PP_ALIGN.CENTER, fill=None):
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(fill)
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = cell.margin_right = Emu(18288)
    cell.margin_top = cell.margin_bottom = Emu(9144)
    tf = cell.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = align
    _emit_line(p, parts, size, color, bold, italic)
