# paper-redraw-pptx

把**论文框架图 / 示意图截图**转成**原生可编辑的 PowerPoint**——不是把图片贴进 PPT，而是把文字、卡片、箭头、表格、公式全部重建为 PowerPoint 原生形状；原图里的精致 3D 渲染图标（小车、芯片、器官、仪器…）从原图裁剪嵌入并羽化融合，观感与原图几乎一致。

## 方法：分类混合法

对原图逐元素清点，按"重画会不会劣化、裁剪会不会留痕"分成三类：

```
              ┌─ 文字/卡片/箭头/表格/公式/胶囊 ──→  python-pptx 原生形状（双击可改）
原图逐元素 ───┼─ 3D渲染图标/照片/复杂插画 ─────→  背景修补 → 裁剪 PNG → 软边缘羽化嵌入
              └─ 简单图形(图钉/循环箭头/徽章…) ──→  少量原生形状组合绘制
```

工作流六步：**量坐标 → 取色 → 修补+裁剪 → 写 build 脚本 → COM 渲染 → 放大视觉验收**（不通过则回上一步修，直到干净）。

完整方法论见 [SKILL.md](SKILL.md)；实战踩坑 17 条见 [references/lessons.md](references/lessons.md)。

## 目录

```
├── SKILL.md                  # Agent 技能说明（方法论 + 六步流程）
├── scripts/
│   ├── helpers_ppt.py        # python-pptx 助手库：形状/富文本(上下标)/箭头/表格/自由曲线/渐变/阴影/羽化
│   ├── render_pptx.py        # PowerPoint COM 渲染预览（pptx → PNG）
│   ├── measure_grid.py       # 坐标网格放大图（量元素像素坐标）
│   ├── crop_prepare.py       # JSON 配置驱动：背景修补 + 裁剪 + 白底转透明 + 检查拼图
│   └── sample_colors.py      # 关键位置取色
└── references/
    ├── lessons.md            # 实战踩坑清单
    └── build_example.py      # 最小可运行示例
```

## 安装

作为 Agent skill（Claude Code / ZCode 等）：

```bash
git clone https://github.com/Nofearli/paper-redraw-pptx.git
mkdir -p ~/.agents/skills   # 或你的技能目录（如 ~/.claude/skills）
cp -r paper-redraw-pptx ~/.agents/skills/
```

依赖：

```bash
pip install python-pptx pillow numpy
pip install pywin32        # 仅 Windows 渲染预览需要
```

渲染预览需要本机装有 Microsoft PowerPoint（COM 自动化）。

## 用法

把论文图截图丢给接入本 skill 的 Agent，说：

> 把这张图转成可编辑的 PPT

Agent 会自动走六步流程，交付 `<名>.pptx`（原生可编辑）+ `preview.png`（渲染预览），并保留 `build.py` 与图标素材以便后续一键重出。

也可以单独使用工具脚本：

```bash
python scripts/measure_grid.py fig.png --crop 550,160,720,320 --scale 5   # 量坐标
python scripts/sample_colors.py fig.png 300,55 1300,108                    # 取色
python scripts/crop_prepare.py prep.json                                   # 修补+裁剪
python scripts/render_pptx.py out.pptx preview.png 3168 1048               # 渲染
```

## 设计取舍

- **为什么不把整张图矢量化/用图像生成重绘？** 图像生成会改变科学内容，矢量化整图得到的曲线节点不可维护。分类混合法保证：所有"内容"（文字/数据结构）原生可编辑，所有"美术"（3D 图标）原汁原味。
- **为什么图标用裁剪位图而不是矢量重画？** 论文里的 3D 渲染图标重画必然退化为扁平示意，失去质感；裁剪 + 羽化在视觉上无损，且位置大小仍可在 PPT 里调整。
- **修补（inpaint）为什么自己写而不用深度学习模型？** 目标区域都是平滑渐变背景，线性插值 + 局部羽化零依赖、秒级完成、不会引入纹理伪影。

## License

MIT
