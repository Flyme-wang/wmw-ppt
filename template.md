# WMW-PPT Template Reference

## Color Constants

```python
from pptx.dml.color import RGBColor

COLOR_DARK = RGBColor(0x2C, 0x5F, 0x7C)       # 深蓝 - 标题
COLOR_MED = RGBColor(0x5A, 0x9B, 0xC7)         # 中蓝 - 标题栏
COLOR_SOFT = RGBColor(0xA8, 0xCF, 0xE8)        # 柔蓝 - 装饰
COLOR_LIGHT = RGBColor(0xD6, 0xE8, 0xF5)       # 浅蓝 - 占位
COLOR_BG = RGBColor(0xF0, 0xF6, 0xFA)          # 背景
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)        # 白色
COLOR_DARK_TEXT = RGBColor(0x33, 0x33, 0x33)    # 正文
COLOR_ACCENT = RGBColor(0xE8, 0xA8, 0x7C)       # 暖橙
COLOR_GREEN = RGBColor(0x7B, 0xC4, 0x7F)        # 柔绿
```

## Matplotlib Colors

```python
COLOR_GRID = '#5A9BC7'      # Model A (primary)
COLOR_XLINE = '#A8CFE8'     # Model B (secondary)
COLOR_POS = '#E8A87C'       # positive change
COLOR_NEG = '#7BC47F'       # negative change
BG_COLOR = '#F8FBFC'        # chart axes background
```

## Slide Layout Diagrams

### Cover (No Background)
```
┌──────────────────────────────────────────────────┐
│ ████████████ (top line 0.08", MED_BLUE)          │
│                                                  │
│        主标题 (36pt bold DARK, centered)           │
│        副标题 (18pt MED_BLUE, centered)            │
│                                                  │
│       ───── 装饰线 (SOFT, 4.3" wide) ─────       │
│                                                  │
│        日期 | 单位 (14pt, centered)                │
│                                                  │
│ ████████████████████ (bottom line 0.06")          │
└──────────────────────────────────────────────────┘
```

### 1×2 Comparison Slide
```
┌──────────────────────────────────────────────────┐
│ ┌────────────────────────────────────────────┐   │
│ │  Title Bar (0.52" tall, MED_BLUE fill)     │   │
│ └────────────────────────────────────────────┘   │
│                                                  │
│ ┌─────────────────┐  ┌─────────────────┐        │
│ │                 │  │                 │        │
│ │  Model A        │  │  Model B        │        │
│ │  (5.9"×5.2")   │  │  (5.9"×5.2")   │        │
│ │                 │  │                 │        │
│ └─────────────────┘  └─────────────────┘        │
│  Caption A              Caption B                │
│                                          (n/N)   │
└──────────────────────────────────────────────────┘
```

### 2×2 Comparison Slide
```
┌──────────────────────────────────────────────────┐
│ ┌────────────────────────────────────────────┐   │
│ │  Title Bar                                 │   │
│ └────────────────────────────────────────────┘   │
│ ┌─────────────────┐  ┌─────────────────┐        │
│ │  Var1 Model A   │  │  Var1 Model B   │        │
│ │  (5.9"×2.7")   │  │  (5.9"×2.7")   │        │
│ └─────────────────┘  └─────────────────┘        │
│  Caption 1             Caption 2                 │
│ ┌─────────────────┐  ┌─────────────────┐        │
│ │  Var2 Model A   │  │  Var2 Model B   │        │
│ └─────────────────┘  └─────────────────┘        │
│  Caption 3             Caption 4                 │
│                                          (n/N)   │
└──────────────────────────────────────────────────┘
```

### Text-Heavy Slide (Mechanism/Recommendations)
```
┌──────────────────────────────────────────────────┐
│ ┌────────────────────────────────────────────┐   │
│ │  Title Bar                                 │   │
│ └────────────────────────────────────────────┘   │
│                                                  │
│  1. Section Header (16pt bold DARK)              │
│     • Point (15pt, DARK_TEXT)                    │
│     • Point (15pt, MED_BLUE for emphasis)        │
│                                                  │
│  2. Section Header                               │
│     • Point                                      │
│                                                  │
│  3. Key Insight (16pt bold ACCENT)               │
│                                                  │
│                                          (n/N)   │
└──────────────────────────────────────────────────┘
```

## Core Helper Functions

```python
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from PIL import Image

def set_slide_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_title_bar(slide, text):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        Inches(0.4), Inches(0.12), Inches(12.5), Inches(0.52))
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLOR_MED
    shape.line.fill.background()
    tf = shape.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = '    ' + text
    set_font(run, 18, True, COLOR_WHITE)

def add_textbox(slide, left, top, width, height, text,
                size=16, bold=False, color=COLOR_DARK_TEXT,
                align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_font(run, size, bold, color)
    return txBox

def add_image_fixed_ratio(slide, img_path, left, top, max_w, max_h):
    if not os.path.exists(img_path):
        return
    img = Image.open(img_path)
    iw, ih = img.size
    ratio = min(float(max_w)/iw, float(max_h)/ih)
    new_w, new_h = int(iw*ratio), int(ih*ratio)
    left_adj = left + (max_w - new_w) // 2
    top_adj = top + (max_h - new_h) // 2
    slide.shapes.add_picture(img_path, left_adj, top_adj, new_w, new_h)

def add_image_with_border(slide, img_path, left, top, width, height):
    border = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        left-Inches(0.02), top-Inches(0.02),
        width+Inches(0.04), height+Inches(0.04))
    border.fill.background()
    border.line.color.rgb = COLOR_MED
    border.line.width = Pt(1.2)
    add_image_fixed_ratio(slide, img_path, left, top, int(width), int(height))

def add_page_number(slide, num, total):
    add_textbox(slide, Inches(12.4), Inches(7.05), Inches(0.8), Inches(0.3),
                f'{num}/{total}', 10, False, COLOR_DARK, PP_ALIGN.RIGHT)

def add_multiline_textbox(slide, left, top, width, height, lines, size=16):
    """lines: list of (text, bold, color)"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, (text, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(4)
        run = p.add_run()
        run.text = text
        set_font(run, size, bold, color)
    return txBox
```

## Unified Colorbar Workflow

```python
# Step 1: Extract ranges from ALL models
def get_variable_ranges(odb_path, step_name='Injection'):
    """Return {var_key: (global_min, global_max)} for one ODB."""
    ...

# Step 2: Merge ranges
unified = {}
for key in all_keys:
    unified[key] = (min(r[key][0] for r in all_ranges),
                    max(r[key][1] for r in all_ranges))

# Step 3: Apply before each snap
def set_limits(vp, vmin, vmax):
    vp.odbDisplay.contourOptions.setValues(
        minAutoCompute=False, maxAutoCompute=False,
        minValue=vmin, maxValue=vmax)
```

## Radar Chart Template

```python
import numpy as np

metrics = {
    '孔压峰值': (val_A, val_B),
    'Mises应力': (...),
    'PEEQ': (...),
    '合位移': (...),
    '拉伸损伤': (...),
    '压缩损伤': (...),
    '屈服单元': (...),
    # 8-12 dimensions recommended
}

categories = list(metrics.keys())
N = len(categories)
vals_A = [metrics[k][0]/max(metrics[k]) for k in categories]
vals_B = [metrics[k][1]/max(metrics[k]) for k in categories]
vals_A += vals_A[:1]  # close polygon
vals_B += vals_B[:1]

angles = [n/N * 2*np.pi for n in range(N)] + [0]

fig, ax = plt.subplots(figsize=(9,9), subplot_kw=dict(polar=True))
ax.set_theta_offset(np.pi/2)
ax.set_theta_direction(-1)
ax.set_thetagrids(np.degrees(angles[:-1]), categories, fontsize=11, fontweight='bold')
ax.plot(angles, vals_A, 'o-', lw=2.2, color='#5A9BC7', label='Model A')
ax.fill(angles, vals_A, alpha=0.2, color='#5A9BC7')
ax.plot(angles, vals_B, 's-', lw=2.2, color='#E8A87C', label='Model B')
ax.fill(angles, vals_B, alpha=0.2, color='#E8A87C')
ax.set_ylim(0, 1.1)
ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
plt.savefig('chart_radar.png', dpi=150, bbox_inches='tight', facecolor='white')
```
