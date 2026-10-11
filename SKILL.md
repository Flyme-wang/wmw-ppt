---
name: wmw-ppt
description: Generate professional scientific research PPTs for FEA/numerical simulation comparison. Covers full workflow: wei204 view screenshot capture with unified colorbar, matplotlib charts (bar, radar, change-rate), and python-pptx slide assembly with soft-blue academic style. Use when user asks to create PPT for simulation results, model comparison, mechanism analysis, or research reporting.
---

# WMW Scientific PPT Generator

Generate publication-quality PPTs for numerical simulation (FEA/FDM/DEM) comparison and analysis.

---

## Design Philosophy (HIGHEST PRIORITY)

### Core Principle: Minimalist Academic Style
- **极简克制**：版式干净，拒绝花哨装饰、渐变、卡通元素
- **图表为王**：图表是 PPT 核心视觉，文字为图表服务，不是反过来
- **一页一观点**：每页只传递一个核心结论，避免信息过载
- **留白呼吸**：页面四周预留充足空白，不堆满内容

### Cross-Model Comparison Rule (MANDATORY — 最高优先级)
**多模型/多工况对比（注入速率、网格尺寸、材料参数等）时，所有模型的同一变量必须使用统一 colorbar 范围**：
- 不统一 colorbar = 相同颜色代表不同数值，跨图对比完全失真，无法看出工况间差异
- **基准确定**：先扫描所有模型各变量的最大值，以结果最大的模型（最极端工况）为基准，上限向上取整到整齐数值
- 所有模型截图使用同一组范围，PPT 中需注明统一 colorbar 范围（如"S-Mises 0~1440 MPa"）
- **预期效果**：色标统一后，工况越极端，高值区（红/橙）覆盖范围越大，差异一眼可见

### Text Rules (MANDATORY)
| Rule | Requirement |
|------|-------------|
| 极致精简 | 摒弃大段段落，每条要点 ≤ 2 行 |
| 单页单论点 | 一页 PPT 只传递一个核心观点 |
| 标题即结论 | 标题直抒结论（如"η 随 Q 增大而降低"），禁止空洞宽泛标题（如"结果分析"） |
| 删除冗余 | 描述性、过渡性文字全部删除，只留核心结论和关键参数 |
| 短句要点 | 使用 bullet points，不用完整句子段落 |
| 轻度高亮 | 重点数据/结论用粗体或强调色，**每页最多 1-2 处高亮**，杜绝多处高亮 |

### Chart Rules (MANDATORY)
| Rule | Requirement |
|------|-------------|
| 图表优先 | 能用图表展示的不用文字，图表布局优先于文字 |
| 标注齐全 | 坐标轴标签、图例、单位、关键数值标注必须完整 |
| 尺寸合理 | 图表不能过小（最小宽度 ≥ 5"），确保投影清晰可读 |
| 图注详细 | 每张图附清晰详细图注：模型/工况名 + 关键数值 + 一句话结论解读，读者不读正文也能看懂图 |
| 配色统一 | 图表配色与 PPT 整体低饱和色系一致 |
| 矢量优先 | 禁止模糊示意图，优先规整矢量图表样式 |

### Layout Rules (MANDATORY)
| Rule | Requirement |
|------|-------------|
| 图文分区 | 文字区与图表区明确分隔，文字不覆盖图表 |
| 对齐规整 | 横向/纵向严格对齐，不随意放置 |
| 分栏布局 | 合理使用左右分栏（图左文右 / 图右文左） |
| 留白充足 | 页面边距 ≥ 0.3"，元素间距 ≥ 0.2" |
| 视觉优先位 | 重点结论放在页面上半部或左侧（视觉优先位置） |
| **页码必备** | **每一页都必须有页码**（封面可省略），位于右下角，用**原生文本框**写入（禁止用图片或截图）；序号必须与实际页序一致，插入内容后要重新核对 |

#### 页码实现（2026-10 用户规范，强制执行）

```python
def page_no(slide, n, total=None):
    """每页右下角原生文本框页码。n 为实际页序，必须逐页核对。"""
    txt = str(n) if total is None else "%d / %d" % (n, total)
    tb = slide.shapes.add_textbox(Inches(12.6), Inches(7.08), Inches(0.6), Inches(0.3))
    p = tb.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run(); r.text = txt
    set_font(r, 10, False, RGBColor(0x7F, 0x8C, 0x8D))
```

- **必须逐页调用**，封面页（第 1 页）可留空但**其余每页不得遗漏**
- 用「封面不编号」方案时，正文首页仍写实际页序（即正文首页显示 2），不要重新从 1 开始
- **插入/删除页面后必须重排**：曾经出现过符号页与公式页插到封面之前、导致后续页码整体错位的缺陷
- 交付前用渲染图逐页确认右下角有号，且数值连续无跳号


### Per-Slide Output Format
每页 PPT 必须包含以下要素：
1. **页面标题**（结论式，≤ 15 字）
2. **版式布局方案**（图文位置、分栏方式）
3. **精简文字内容**（bullet points，≤ 6 条）
4. **图表设计建议**（图表类型、数据、配色）
5. **排版注意事项**（对齐、留白、高亮位置）

---

## Design Specifications

### Color Scheme (Low-Saturation Academic)
| Token | Hex | Usage |
|-------|-----|-------|
| Dark Blue | `#2C5F7C` | Titles, emphasis text |
| Medium Blue | `#5A9BC7` | Title bars, primary chart color |
| Soft Blue | `#A8CFE8` | Accents, secondary bars |
| Light Blue | `#D6E8F5` | Placeholder backgrounds |
| Background | `#F0F6FA` | Content slide backgrounds |
| Accent Orange | `#E8A87C` | Positive changes, Model B ( sparingly) |
| Soft Green | `#7BC47F` | Negative changes (sparingly) |
| Dark Text | `#333333` | Body text |
| White | `#FFFFFF` | Text on colored fills |

> **配色纪律**：基础色为黑白灰 + 蓝色系，强调色（橙/绿）每页最多出现 1 种，禁止高饱和、花哨渐变。

### Font Settings
- **Chinese**: 微软雅黑 — via `a:ea` XML attribute
- **English**: Times New Roman — via `a:latin` + `font.name`
- **最小字号**：**18pt**（2026-08 用户规范：正文/要点/图注一般不得小于 18pt；仅密集数据表格、页码等特殊场景可例外）
- **Body text**: 18pt
- **Headers/key points**: 18–20pt bold
- **Slide title bar**: 20pt bold white（页面标题 20pt）
- **Cover title**: 36pt bold dark blue
- **Page number**: 10pt bottom-right
- **图注 caption/note**: 18pt，内容必须清楚详细（模型名 + 关键数值 + 结论解读），禁止含糊的"见左图"式注；文本需控制在单行（约 ≤20 字/6 英寸宽），多模型对比格中每格均需独立图注
- **表格文字**: 13–14pt（密集数据表格的例外，比正文小但保持可读）
- **Legend 字体**: **15pt bold**（Abaqus 截图左侧图例字号，通过 `contourOptions.setValues(legendFontSize=15, legendFontBold=ON)` 设置）
- **行间距**：**优先 1.5 倍行距**（`p.line_spacing = 1.5`），密集表格/图注可用 1.2~1.3 倍
- **结果分析深度（2026-08 用户要求）**：PPT 必须包含参数差异→结果差异的定量对照（变化百分比 vs 基准）与逐模型机理分析页，不能只罗列云图与最终结论

### Layout Rules
- Slide: **16:9** (13.333 × 7.5 in)
- Title bar: 0.52" tall, medium blue, left-padded text
- Images: **equal aspect ratio** (PIL-computed), thin blue border (1.2pt)
- **图片等比例缩放（MANDATORY）**：所有截图/图表插入 PPT 时必须保持原始宽高比，**禁止拉伸变形**。使用 PIL 计算等比缩放尺寸，居中放置于预留位置
- Comparison layout: **1×2** (single variable) or **2×2** (two variables)
- **5 图对比排版（MANDATORY，2026-08 用户要求）**：5 个模型/工况同变量对比时**禁止 5 图横排**（每图仅 2.25" 宽，过小），必须用 **3+2 两行布局**：上行 3 图 + 下行 2 图，每图宽 3.27" 高 2.45"（4:3），面积约为横排的 2 倍，清晰可读
  - 布局参数（16:9，13.333×7.5"）：行1 y=0.85，3 图 x0=(13.333-3w-2gap)/2；行2 y=行1图底+0.64，2 图 x0=(13.333-2w-gap)/2；w=3.27，gap=0.25
  - 每格图注两行：第一行模型名（15pt bold 中蓝），第二行观察结论（12pt），不用独立模型标签
  - 页脚统一色标说明 12pt 居中 y=6.94
  - 代码模板见 Step 5 的 `add_5fig_grid_32`
- Cover & conclusion: **NO background** (white)
- Content slides: light blue background `#F0F6FA`
- **图文分栏**：图表占页面 60-70%，文字占 30-40%
- **边距**：左右 ≥ 0.3"，上下 ≥ 0.5"（含标题栏）
- **表格对齐（MANDATORY，2026-08 用户要求）**：表格单元格文字必须**水平居中 + 上下垂直居中**（`cell.vertical_anchor = MSO_ANCHOR.MIDDLE` + 段落 `PP_ALIGN.CENTER`），禁止默认的顶部对齐；表头同样适用

### File Naming Rules (MANDATORY)
- **版本信息放文件名最前**：所有交付文件（PPT/Word/PDF 等）命名时，版本号必须置于文件名开头，格式 `v{版本号}_{主题描述}.{扩展名}`，如 `v1_威204水平层敏感性分析.pptx`、`v2_威204水平层敏感性分析.docx`
- **禁止后缀式版本号**：不允许 `主题描述_v2.pptx` 这类版本号放末尾的命名
- **版本号递增自动取**：生成脚本用循环探测下一个未占用版本号（v1→v2→v3…），首次交付 v1，重大修改递增；主题描述部分保持稳定不变
- **同主题多格式**：同一主题的 PPT/Word/PDF 版本号一致（v2_主题.pptx 与 v2_主题.docx 为同一版本的不同格式）

### Standard Slide Structure (18 pages)
1. **Cover** — title, subtitle, decorative lines, date
2. **Model Overview** — parameters, scheme description
3. **Pore Pressure** — POR 1×2 or 2×2
4. **Displacement Magnitude** — U_Mag 1×2
5. **Displacement Components** — U1/U2 2×2
6. **Stress Mises** — S_Mises 1×2
7. **Stress Components** — S11/S33 2×2
8. **Strain Components** — PE11/PE33 2×2
9. **Equivalent Plastic Strain** — PEEQ 1×2
10. **Tensile Damage** — DAMAGET 1×2
11. **Compressive Damage** — DAMAGEC 1×2
12. **Yield State** — AC YIELD 1×2
13. **Quantitative Charts** — grouped bar charts
14. **Radar Chart** — multi-dimensional normalized comparison
15. **Change Rate** — horizontal bar chart
16. **Mechanism Analysis** — text-heavy
17. **Engineering Recommendations** — text-heavy
18. **Conclusion** — key findings, future work (no background)

---

## Workflow

### Step 1: Extract ODB Data

```python
import numpy as np
from odbAccess import openOdb

def extract_comprehensive(odb_path):
    odb = openOdb(odb_path, readOnly=True)
    frame = odb.steps['Injection'].frames[-1]  # adapt step name
    data = {}
    # POR (nodal): max, min, mean
    # S (Mises + components): max, min per component
    # PE (PE11/22/33): max abs
    # PEEQ: max, mean, count > threshold
    # U (Mag + components): max
    # DAMAGET/DAMAGEC: max, count > 0.1/0.5/0.9
    # AC YIELD: count > 0
    odb.close()
    return data
```

### Step 2: Unified Colorbar Ranges (CRITICAL — 多模型对比必须统一)

**原则**：多模型/多工况对比时，所有模型的同一变量必须使用**完全相同**的 colorbar 范围。否则相同颜色不代表相同数值，对比失去意义。

**基准确定流程**：
1. 扫描每个模型各变量的最大值（integration point 值）
2. 以结果最大的模型（最极端工况，如最大注入速率/最细网格）为基准
3. 上限 = 基准模型最大值，向上取整到整齐数值（如 1.435e9 → 1.44e9）
4. 所有模型统一使用该基准范围截图

**关键陷阱 — 积分点值 vs 节点外推值**：
- ODB 扫描得到的是**积分点值**；图例显示的是**节点外推值**（Avg: 75%），可能比积分点值高 20%~100%
- 基准范围基于积分点值时，最极端模型的节点外推值可能超出上限
- 处理：`outsideLimitsMode=SPECTRUM` → 超限区域显示灰色，图例顶部标注实际峰值（可接受）；或基准上限再放大 10%~20% 使不超限

```python
# 统一范围示例（基准 = 最大排量模型 0p6，取整）
cb_ranges = {
    'S':       (0.0, 1.44e9),   # 1440 MPa
    'PEEQ':    (0.0, 0.105),
    'DAMAGET': (0.0, 0.5),
    'DAMAGEC': (0.0, 0.5),
}

# 每张截图前必须完整设置（4 个参数缺一不可）：
vp.odbDisplay.contourOptions.setValues(
    minValue=vmin, maxValue=vmax,
    minAutoCompute=OFF, maxAutoCompute=OFF,   # 必须显式 OFF，否则自动计算覆盖手动范围
    numIntervals=9, intervalType=UNIFORM,
    outsideLimitsMode=SPECTRUM)               # 超限显示灰色并标注真实峰值
```

> 经验：节点外推值可能比积分点扫描值高 20%~100%。若图例顶部出现灰色超限条目（标注实际峰值），属正常现象；若要求完全不超限，基准上限应放大 10%~20%。

### Step 3: Capture Screenshots (wei204 view plugin)

**MUST use wei204 view plugin** for standardized visualization:

```python
import os
import wei204_view_kernel as k
from abaqusConstants import (INTEGRATION_POINT, NODAL, COMPONENT,
                             INVARIANT, PNG, CONTOURS_ON_DEF)

shot_dir = 'C:/Users/Dell/Downloads/fea_shots'  # ASCII path only!
os.makedirs(shot_dir, exist_ok=True)

vp = session.viewports['Viewport: 1']
odb = vp.displayedObject

# 1. Apply wei204 view
k.apply_view(use_active=True)

# 2. Set last frame
si = list(odb.steps.keys()).index('Injection')
nf = len(odb.steps['Injection'].frames)
vp.odbDisplay.setFrame(step=si, frame=nf-1)
vp.odbDisplay.display.setValues(plotState=(CONTOURS_ON_DEF,))

# 3. Helper functions
def set_var(vp, label, comp=None, inv=None, pos=INTEGRATION_POINT):
    if inv:
        vp.odbDisplay.setPrimaryVariable(variableLabel=label,
            outputPosition=pos, refinement=(INVARIANT, inv))
    elif comp:
        vp.odbDisplay.setPrimaryVariable(variableLabel=label,
            outputPosition=pos, refinement=(COMPONENT, comp))
    else:
        vp.odbDisplay.setPrimaryVariable(variableLabel=label,
            outputPosition=pos)

def snap(vp, name):
    session.printToFile(fileName=os.path.join(shot_dir, name),
                        format=PNG, canvasObjects=(vp,))
```

**Full variable list** (16 per model):
```python
# Stress (INTEGRATION_POINT)
('S', None, 'Mises'), ('S', 'S11', None), ('S', 'S22', None), ('S', 'S33', None)
# Strain (INTEGRATION_POINT)
('PE', 'PE11', None), ('PE', 'PE22', None), ('PE', 'PE33', None)
('PEEQ', None, None)
# Displacement (NODAL)
('U', None, 'Magnitude'), ('U', 'U1', None), ('U', 'U2', None), ('U', 'U3', None)
# Pore pressure (NODAL)
('POR', None, None)
# Damage (INTEGRATION_POINT)
('DAMAGET', None, None), ('DAMAGEC', None, None)
# Yield (INTEGRATION_POINT)
('AC YIELD', None, None)
```

**Key constraints**:
- Path MUST be ASCII (no Chinese) → use `C:/Users/Dell/Downloads/`
- `k.apply_view(use_active=True)` always acts on **Viewport: 1**
- Multiple ODBs: process sequentially (close → open → apply → capture)
- Naming: `{model}_{variable}.png`

### Step 4: Generate Charts (matplotlib)

```python
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

COLOR_GRID = '#5A9BC7'    # Model A
COLOR_XLINE = '#A8CFE8'   # Model B
COLOR_POS = '#E8A87C'     # positive change
COLOR_NEG = '#7BC47F'     # negative change
BG = '#F8FBFC'
```

**Required chart types**:
1. Grouped bar charts (stress, strain, displacement, POR, damage)
2. Horizontal change-rate chart (all metrics, color-coded ±)
3. **Radar/spider chart** (10-dimension normalized comparison):

```python
fig, ax = plt.subplots(figsize=(9, 9), subplot_kw=dict(polar=True))
# Normalize: each metric / max(both models)
angles = [n/N * 2π for n in range(N)] + [0]
ax.plot(angles, vals_A, 'o-', color=COLOR_GRID)
ax.fill(angles, vals_A, alpha=0.2)
ax.plot(angles, vals_B, 's-', color=COLOR_POS)
ax.fill(angles, vals_B, alpha=0.2)
```

Style rules:
- Remove top/right spines
- `ax.set_facecolor(BG)`
- `dpi=150`, `bbox_inches='tight'`, `facecolor='white'`
- Value labels on bars (fontsize=9)

### Step 5: Build PPT (python-pptx)

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
```

**Font helper** (critical for CJK+Latin):
```python
def set_font(run, size_pt=16, bold=False, color=COLOR_DARK_TEXT):
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = 'Times New Roman'
    rPr = run._r.get_or_add_rPr()
    ea = rPr.makeelement(qn('a:ea'), {})
    ea.set('typeface', '微软雅黑')
    rPr.append(ea)
    latin = rPr.makeelement(qn('a:latin'), {})
    latin.set('typeface', 'Times New Roman')
    rPr.append(latin)
```

**Image placement** (equal aspect ratio):
```python
def add_image_fixed_ratio(slide, img_path, left, top, max_w, max_h):
    img = Image.open(img_path)
    ratio = min(max_w/img.size[0], max_h/img.size[1])
    # 必须保持浮点！int() 截断会破坏比例（如 1.6875"→1" 导致比例失真）
    new_w = img.size[0]*ratio
    new_h = img.size[1]*ratio
    left_adj = left + (max_w - new_w)/2
    top_adj = top + (max_h - new_h)/2
    slide.shapes.add_picture(img_path, Inches(left_adj), Inches(top_adj),
                             Inches(new_w), Inches(new_h))
```

**5 图 3+2 两行排版** (MANDATORY for 5-model comparison, 2026-08):
```python
# 5图3+2两行布局: 上行3图+下行2图, 每格图注=模型名(bold 15pt)+观察(12pt)
# 图宽3.27"高2.45"(4:3), 面积约为5图横排(2.25")的2倍, 更清晰
# 调用: add_5fig_grid_32(slide, lambda m: snap_f(m, face), notes, view_desc)
def add_5fig_grid_32(slide, img_getter, notes, view_desc):
    IMG_W, IMG_H, GAP, TOP_Y = 3.27, 2.45, 0.25, 0.85
    x1 = (13.333 - (3*IMG_W + 2*GAP)) / 2
    x2 = (13.333 - (2*IMG_W + GAP)) / 2
    row1 = [(x1, TOP_Y), (x1+IMG_W+GAP, TOP_Y), (x1+2*(IMG_W+GAP), TOP_Y)]
    row2_y = TOP_Y + IMG_H + 0.04 + 0.46 + 0.14
    row2 = [(x2, row2_y), (x2+IMG_W+GAP, row2_y)]
    for i, ((m, cn), (px, py)) in enumerate(zip(MODELS, row1+row2)):
        add_image_fixed_ratio(slide, img_getter(m), px, py, IMG_W, IMG_H)
        add_text(slide, px, py+IMG_H+0.04, IMG_W, 0.46,
                 [(cn, 15, True, C_MED), (notes[i], 12, False, C_TEXT)],
                 align=PP_ALIGN.CENTER, line_spacing=1.15)
    add_text(slide, 0.3, 6.94, 12.73, 0.35,
             [(view_desc, 12, False, RGBColor(0x88, 0x88, 0x88))],
             align=PP_ALIGN.CENTER, line_spacing=1.0)
```

**Table cells** (水平居中 + 上下垂直居中，MANDATORY):
```python
from pptx.enum.text import MSO_ANCHOR  # 注意：MSO_ANCHOR 在 pptx.enum.text

cell = table.cell(row, col)
cell.vertical_anchor = MSO_ANCHOR.MIDDLE   # 上下垂直居中（默认 TOP）
cell.margin_top = Inches(0.02)
cell.margin_bottom = Inches(0.02)
for p in cell.text_frame.paragraphs:
    p.alignment = PP_ALIGN.CENTER          # 水平居中
```

### Step 6: Verification

1. File size > 500KB (images embedded)
2. Slide count matches expected
3. All image paths resolve
4. Fonts: verify `a:ea`=微软雅黑, `a:latin`=Times New Roman
5. Tables: verify `cell.vertical_anchor == MSO_ANCHOR.MIDDLE` on every cell (文字在框内上下居中)

---

## Reference Files
- `generate_ppt.py` — complete PPT builder (workspace root)
- `generate_charts.py` — chart generation with radar (workspace root)
- `template.md` — layout diagrams & code snippets (this directory)

## Dependencies
- `python-pptx` ≥ 0.6.21
- `matplotlib` ≥ 3.5
- `Pillow` (PIL)
- `numpy`
- Abaqus MCP (for ODB access & screenshots)

## Common Pitfalls
| Issue | Solution |
|-------|----------|
| Chinese text as boxes | Set `a:ea` font via XML |
| Images distorted | PIL aspect-ratio calculation |
| Colorbar mismatch | `minAutoCompute=False` + unified range |
| Unicode path error | Use ASCII-only save path |
| `k.apply_view` wrong viewport | Always use Viewport: 1, process ODBs sequentially |
| Stress variable error | Must specify `refinement=(INVARIANT, 'Mises')` |
| `setFrame` TypeError | Use `setFrame(step=int, frame=int)` not string |
| Multi-model colorbar mismatch | Scan all models first, set unified range from the most extreme model |
| Legend shows gray over-limit value | Nodal extrapolated value > integration point max; use SPECTRUM mode or add 10–20% margin |
| `_SELECTION_CACHE` stale selection | Clear `_SELECTION_CACHE` before opening each new ODB (id(odb) memory-reuse bug) |
| Manual limits not applied | Must set `minAutoCompute=OFF, maxAutoCompute=OFF` explicitly |
| `setPrimaryVariable` TypeError | Tensor: positional `('S', INTEGRATION_POINT, (INVARIANT, 'Mises'))`; scalar: `(variableLabel='PEEQ', outputPosition=INTEGRATION_POINT)` |
| Contact output (CSLIP/CSHEAR/CPRESS) not visible in viewport | Contact variables are not directly available via `setPrimaryVariable`; use matplotlib heatmaps or field-value extraction instead |
| Images distorted in PPT | Always use PIL to compute aspect-ratio-preserving dimensions; never hardcode width/height that differ from source ratio |
| `int()` truncation distorts ratio | `add_image_fixed_ratio` 内 new_w/new_h 必须保持浮点（勿 int()），Inches() 接受浮点；否则 1.6875"→1" 比例失真 |
| 5 图横排每图过小 | 5 模型同变量对比禁止横排（2.25" 过小），必须用 3+2 两行布局（3.27"，add_5fig_grid_32） |
| 页码缺失或错位 | 插入/删除页面后未重排 | 每页调用 page_no()；改版后重新核对实际页序，逐页渲染确认 |
