# -*- coding: utf-8 -*-
"""最小示例：用本 skill 的 helpers 搭一个"面板+卡片+胶囊+表格+公式+徽章+blob+箭头"骨架。
运行: python build_example.py  → example.pptx（再用 render_pptx.py 渲染检查）
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
from helpers_ppt import *
from pptx import Presentation
from pptx.util import Inches as In
from pptx.enum.shapes import MSO_SHAPE

set_palette({"panel": "EBF3FB", "panel_bd": "9DB8DC", "title": "2A4A78"})

prs = Presentation()
prs.slide_width = In(8.0)
prs.slide_height = In(4.5)
S = prs.slides.add_slide(prs.slide_layouts[6]).shapes

# 面板底
rrect(S, 0.4, 0.5, 7.2, 3.6, fill="panel", line="panel_bd", line_w=1.0, radius=0.05)
# 标题胶囊（跨面板覆盖元素放最后画也行，这里顺序无冲突）
t = rrect(S, 2.6, 0.32, 2.8, 0.40, fill="title", radius=0.5)
set_shape_text(t, "(a) Example panel", size=16, color="FFFFFF", bold=True, italic=True)

# 白卡片（含阴影）
card = rrect(S, 0.7, 1.0, 2.4, 2.6, fill="FFFFFF", line="A9C6F0", line_w=1.0, radius=0.1)
add_shadow(card)
set_shape_text(card, ["Frozen GNN", "+ Residual Head"], size=13, color="navy_text", bold=True, line_spacing=0.95)

# 公式色带（富文本下标）
b1 = rrect(S, 3.4, 1.2, 2.2, 0.38, fill="FBF2E1", line="E4C99A", line_w=0.75, radius=0.25)
set_shape_text(b1, [("d", "i"), ("corr", "sub"), (" = dist(", "n"), ("S", "i"), ("corr", "sub"),
                    (", ", "n"), ("T", "i"), (")", "n")], size=11.5)

# 权重矩阵（2×3 示例）
tbl = add_grid_table(S, 5.9, 1.0, 1.8, 1.2, 3, 3, [0.5, 0.65, 0.65], [0.4, 0.4, 0.4])
cell_borders(tbl.cell(0, 1)); set_cell(tbl.cell(0, 1), [("E", "bi"), ("1", "sub")], size=12, color="red_label")
cell_borders(tbl.cell(0, 2)); set_cell(tbl.cell(0, 2), [("E", "bi"), ("2", "sub")], size=12, color="red_label")
cell_borders(tbl.cell(1, 0)); set_cell(tbl.cell(1, 0), [("P", "bi"), ("1", "sub")], size=12, color="blue_label", fill="blue_row")
cell_borders(tbl.cell(1, 1)); set_cell(tbl.cell(1, 1), [("w", "i"), ("11", "sub")], size=12)
cell_borders(tbl.cell(1, 2)); set_cell(tbl.cell(1, 2), [("w", "i"), ("12", "sub")], size=12, fill="amber")
cell_borders(tbl.cell(2, 0)); set_cell(tbl.cell(2, 0), [("P", "bi"), ("2", "sub")], size=12, color="blue_label", fill="blue_row")
cell_borders(tbl.cell(2, 1)); set_cell(tbl.cell(2, 1), [("w", "i"), ("21", "sub")], size=12)
cell_borders(tbl.cell(2, 2)); set_cell(tbl.cell(2, 2), [("w", "i"), ("22", "sub")], size=12)

# 自由曲线 blob + 局部异色弧
bl, pts = blob(S, 1.9, 3.9, 0.7, 0.55, fill="DCE9FA", line="blob_blue", line_w=2.5, a1=0.12, p1=2.2)
blob_arc(S, pts, 20, 52, "blob_red", line_w=2.5)

# 箭头
arrow(S, 3.1, 2.3, 3.35, 2.3, color="gray", w=2.5)

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "example.pptx")
prs.save(out)
print("saved:", out)
