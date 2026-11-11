---
name: paper-redraw-pptx
description: 把论文框架图/示意图截图转成原生可编辑 PPT。分类混合法：文字/卡片/箭头/表格/公式全部用 PowerPoint 原生形状重建，精致 3D 图标从原图裁剪嵌入（背景修补 + 软边缘羽化），PowerPoint COM 渲染 + 放大视觉验收闭环。凡用户给出论文插图截图并要求"转成可编辑 PPT""图转 PPT""复现成 PPT""把这张图做成可编辑的"时使用。
---

# Paper Redraw PPTX —— 论文图 → 原生可编辑 PPT

一张论文框架图（截图或 PNG），交付一份**原生可编辑**的 .pptx：文字、卡片、箭头、表格、公式全部是 PowerPoint 形状；原图里的精致 3D 渲染图标（小车、芯片、器官、仪器等）以裁剪位图嵌入并羽化融合。**禁止把整张图当作一张图片贴进 PPT。**

## 环境依赖

- Python 3 + `python-pptx`、`Pillow`、`numpy`（Windows 推荐 conda 环境）
- 渲染预览：Windows + Microsoft PowerPoint（COM，需 `pywin32`）；无 PowerPoint 时跳过自动渲染，改为让用户截图验收
- 字体：Times New Roman / 系统字体（`helpers_ppt.py` 默认 serif）

## 核心原则：元素三分类

对原图逐元素清点，每个元素归入三类（决策表）：

| 类别 | 判据 | 做法 |
|---|---|---|
| **原生重建** | 文字、圆角卡片、标题胶囊、箭头、表格、公式、色带 | python-pptx 形状/文本框/表格；双击可改 |
| **裁剪嵌入** | 3D 渲染图标、照片、复杂插画（重画必然劣化） | 修补背景后裁剪成 PNG，`soft_edge()` 羽化嵌入；可移动缩放，不可改色 |
| **原生绘制** | 结构简单的图形：定位针、循环箭头、徽章、星形、盾牌、芯片 | 少量形状组合（见 lessons.md §9）；比裁剪更干净 |

判断口诀：**"重画会变玩具的 → 裁剪；裁剪会留痕迹的 → 原生画；其余全原生。"**

## 六步流程

约定：全程以**原图像素**为单位记坐标，`100 px = 1 英寸`（`helpers_ppt.py` 的 `In(x)` 直接吃英寸值，除以 100 即可）。工具都在本 skill 的 `scripts/` 下。

### S1 量坐标

```bash
python scripts/measure_grid.py original.png                    # 全图网格
python scripts/measure_grid.py original.png --crop 550,160,720,320 --scale 5   # 局部放大
```

生成带品红竖线/青色横线（每 10px 一条，每 20px 标数）的网格图。**逐元素读出像素框**（卡片、胶囊、图标、表格行列、徽章）。放大局部量细节，不要靠肉眼估。

### S2 取色

```bash
python scripts/sample_colors.py original.png 300,55 450,60 1300,108 1400,170
```

在每个关键位置采样：面板底色（多点，判断是否渐变）、标题条两端、高亮格、边框、箭头色。**渐变要采两端点**；卡片底色要采被裁图标周边的底色（羽化融合靠它）。

### S3 素材准备（修补 + 裁剪）

写一个 JSON 配置，跑 `crop_prepare.py`：

```json
{
  "src": "original.png",
  "out_dir": "assets/icons",
  "patched": "assets/patched_full.png",
  "patches": [
    {"box": [78, 94, 126, 136], "mode": "h"},
    {"box": [696, 172, 714, 236], "mode": "v"}
  ],
  "crops": {
    "chip_zone.png": {"box": [1222, 103, 1340, 182], "source": "clean"}
  },
  "alpha": ["pin_e1.png"],
  "contact_sheet": {"height": 190}
}
```

```bash
python scripts/crop_prepare.py prep.json
```

- `patches`：把原图里**烙死的文字/徽章**修补掉（mode `h`=按行左右插值，`v`=按列上下插值），之后用可编辑文本框补回——字是活的、底是干净的
- **标签徽章必须整框修补**（含边框），只补文字会留下边框弧线残迹
- `crops` 的 `source`：`clean`=修补后图（默认）、`original`=原图
- `alpha`：白底转透明（阈值默认 208/242）
- 自动生成 `contact_sheet.png`——**必须打开检查每个图标**再进入下一步

### S4 写 build 脚本

参考 `scripts/helpers_ppt.py`（必须 import 它，不要重写这些函数）：`rrect / oval / arrow / elbow / pic / soft_edge / add_shadow / grad_fill / blob / add_text / set_shape_text / add_table 系列`。

骨架（完整示例见 `references/build_example.py`）：

```python
from helpers_ppt import *
prs = Presentation(); prs.slide_width = In(W_PX/100); prs.slide_height = In(H_PX/100)
S = prs.slides.add_slide(prs.slide_layouts[6]).shapes
def X(px): return px / 100.0          # 像素 → 英寸

# 面板底 → 卡片 → 图标(pic+soft_edge) → 文本/表格/公式 → 箭头 → 跨面板覆盖元素(最后画)
```

要点（细节全部在 `references/lessons.md`，写之前通读一遍）：

- **表格**：`add_table` 后必须 `strip_table_style()` + 逐格 `cell_borders()`，否则自带蓝色样式
- **公式**：rich-run 下标 —— `[[("S","i"),("corr","sub"),(" = {z : ","n"),...]]`，sub/sup 用 baseline ±，字号 ×0.72
- **胶囊条**：`rrect(..., radius=0.5)`；圆角矩形 radius 是短边比例（0~0.5）
- **文字**：全部走 `add_text/set_shape_text`（自动 Times、居中、word_wrap、margin=0）
- **裁剪图标**：`pic()` 后立刻 `soft_edge(p, 2.5~3.0)`；位置不得超出所在卡片边界
- **调色板**：`set_palette({...})` 换成 S2 采到的色；`rgb("key")` 或 `rgb("#RRGGBB")` 均可
- **自由曲线区域**（原图里的手绘感 blob）：`blob()` 生成、`blob_arc()` 分段描边、`grad_fill()` 渐变填充

### S5 渲染预览

```bash
python scripts/render_pptx.py out.pptx preview.png 3168 1048
# 宽 = 原图宽 px × 2，保证预览清晰；高按原图宽高比折算
```

### S6 放大视觉验收（闭环）

对照原图逐区域放大检查（用 S1 的网格图或 PIL crop 放大 3~5 倍）：

1. 有无裁剪残留（弧线、切边、色块接缝、超出卡片边界）
2. 文字无重叠压线、下标正确、颜色正确
3. **同名对象全图配色一致**（如同一个 S_corr 在两个面板里必须同色）
4. 有问题 → 回 S3/S4 修 → 重渲染 → 再验，直到干净

有视觉验收子代理（如 visual-judge）时，把 preview.png + 原图交给它判定；没有就自己放大逐区域核对。

## 输出约定

- `<项目名>.pptx` + `preview.png`，放在用户指定的输出目录
- 保留 `build.py`（再生成器）与 `assets/icons/`（裁剪素材）：后续改字/换色可直接编辑 pptx，或改脚本一键重出

## references/lessons.md

**写 build 脚本前必读**——14 条实战踩坑（徽章残框、卡片接缝、z-order、表格样式、插值方向、羽化阈值等），每条都是返工换来的。
