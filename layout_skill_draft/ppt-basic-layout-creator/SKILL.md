---
name: ppt-basic-layout-creator
description: Create basic slide layouts with reliable, cohesive color palettes and clean visual hierarchy for presentations. Use when asked to generate a presentation, slide deck, or PPT layout.
---

# Basic Slide Layout & Color System

Generate clean, well-proportioned basic slide layouts with strong contrast and visual balance. Use the available presentation-authoring skill for file creation, rendering, and validation; apply this skill as the layout and color system.

## 1. Proven Color Palettes

Pick one palette per deck; use Clean Slate by default. Never invent random hex codes. Apply the chosen tokens consistently across all slides. When a palette offers two accents, choose one for the entire deck.

| Token | A: Clean Slate | B: Warm Editorial | C: Midnight Executive | D: Forest & Sage |
| --- | --- | --- | --- | --- |
| Slide background | `#F8FAFC` | `#FAF8F5` | `#0F172A` | `#F4F7F4` |
| Card / surface | `#FFFFFF` | `#FFFFFF` | `#1E293B` | `#FFFFFF` |
| Surface border | `#E2E8F0` | `#E6E2DC` | `#334155` | `#DCE5DC` |
| Primary text / headings | `#0F172A` | `#1C1917` | `#F8FAFC` | `#14261C` |
| Body / muted text | `#475569` | `#57534E` | `#94A3B8` | `#4A5D52` |
| Primary accent | `#2563EB` | `#D97706` or `#C2410C` | `#38BDF8` or `#818CF8` | `#15803D` |

Use subtle 1px-equivalent surface borders. For Warm Editorial, prefer Terracotta (`#C2410C`) when the accent is used for small text; reserve Warm Amber (`#D97706`) for large text and decoration because it does not provide sufficient contrast for small text on these light surfaces. If using Amber, use primary text for small kickers and tags instead.

## 2. Core Basic Slide Layouts

Stick to these five foundational layouts. Use a 16:9 canvas: `13.333" × 7.5"` in `python-pptx` or `1280px × 720px` in HTML/CSS. Keep generous outer margins: `0.8"` in PowerPoint or `64px` in HTML/CSS. Treat these as medium-specific margin choices, not exact unit conversions.

Keep a consistent top header area on content slides. Use approximately `28–32pt` bold slide headings and align all content to the same left and right guides. Use one font family with consistent weights throughout the deck.

Calculate repeated element widths from available space:

`element_width = (canvas_width - 2 * outer_margin - (count - 1) * gap) / count`

Treat the dimensions below as approximate targets. Prioritize equal widths and the outer margins if rounding would push content outside the safe area.

### Layout 1: Title Slide

- Use a full-bleed background from the selected palette and a vertically centered, left-aligned text block.
- Use the selected palette's text colors. For a dark title background, use Midnight Executive for the entire deck; do not place a light palette's dark text on a dark background.
- Category kicker: `12pt`, uppercase, bold, primary accent subject to the small-text contrast rule.
- Main title: `40–48pt`, bold, primary text.
- Subtitle: `18–20pt`, regular, muted text, maximum two lines.

### Layout 2: 2-Column Split

- Place two equal columns below the header, approximately `5.6"` wide with a `0.5"` gap. With the specified PowerPoint canvas and margins, calculate each width as `5.6165"`.
- Use for comparisons, before/after, problem versus solution, or text on the left with a visual/chart on the right.
- Align column tops and keep equivalent hierarchy on both sides.

### Layout 3: 3-Card Row

- Place three equal cards below the header, approximately `3.6"` wide, `4.2"` tall, with `0.4"` gaps. With the specified PowerPoint canvas and margins, calculate each width as approximately `3.6443"`.
- Fill each card with the card/surface color and use a subtle surface border.
- Top accent number or tag: `20pt`, bold, primary accent.
- Card title: `18pt`, bold, primary text.
- Description: `14pt`, regular, muted text, maximum 30 words.
- Use consistent internal padding and align matching elements across all cards.

### Layout 4: 2×2 Grid

- Place four equal rectangular cards below the header in two rows and two columns.
- Target approximately `5.6" × 2.1"` per card, with `0.4"` horizontal and vertical gaps. With the specified PowerPoint canvas and margins, calculate each width as `5.6665"`.
- Use for four pillars, quarterly roadmaps, or grouped takeaways.
- Keep padding, title placement, and text hierarchy consistent across all four cards.

### Layout 5: Big Stat / Callout Slide

- Use one centered callout or a three-column metric showcase.
- Oversized number: `64–72pt`, bold, primary accent.
- Metric label: `18pt`, bold, primary text.
- Context note: `14pt`, regular, muted text, one line.
- Keep units and labels explicit; preserve provided values and avoid inventing statistics.

## 3. Non-Negotiable Styling Rules

1. **Contrast:** Never put dark text on a dark surface or light text on a light surface. Pair primary and muted text only with the selected palette's slide background or card/surface. Check accent text against its actual background as well.
2. **Accent restraint:** Use the primary accent on at most 10–15% of the slide, such as kickers, stat numbers, badges, or thin top card borders. Never use it as a full-bleed background with low-contrast text.
3. **Blank canvas:** When generating `python-pptx` code, always use `prs.slide_layouts[6]` and explicitly place shapes so default placeholders do not clash.
4. **Text caps:** Use one core idea per slide, at most 3–4 bullets per column, and at most 12 words per bullet line. Always enable `tf.word_wrap = True`. Split dense content across slides rather than shrinking type to force a fit.

## 4. Image Sizing & Aspect-Ratio Fitting

When inserting photos, maps, or satellite imagery, inspect each image with `PIL.Image` first. Read its pixel dimensions and compute `img_w / img_h`. Never pass unrelated width and height values to `slide.shapes.add_picture()`.

Fit and center the entire image inside its target layout box without stretching, cropping away evidence, overflowing, or overlapping text. Reserve a separate caption area below the box. Use the actual image provider, acquisition date, and source URL in the caption or slide notes; do not label WorldView imagery as Sentinel-2.

Use consistent units for every box argument, preferably PowerPoint EMUs from `Inches(...)`. Round calculated coordinates to integer EMUs before insertion. Validate positive image and box dimensions. Normalize EXIF orientation before reading dimensions if needed.

```python
from PIL import Image
from pptx.util import Inches


def add_fitted_image(slide, img_path, box_left, box_top, max_w, max_h):
    with Image.open(img_path) as im:
        w_px, h_px = im.size
    if min(w_px, h_px, max_w, max_h) <= 0:
        raise ValueError("Image and bounding box dimensions must be positive")
    aspect = w_px / h_px
    box_aspect = max_w / max_h
    if aspect > box_aspect:
        final_w = max_w
        final_h = max_w / aspect
    else:
        final_h = max_h
        final_w = max_h * aspect
    left = box_left + (max_w - final_w) / 2
    top = box_top + (max_h - final_h) / 2
    return slide.shapes.add_picture(
        img_path, round(left), round(top),
        width=round(final_w), height=round(final_h),
    )
```

Passing both dimensions is appropriate here because both were calculated from the original aspect ratio. After placement, verify that the image stays within its box, preserves its original proportions within rounding tolerance, and leaves the caption readable.

## 5. Final Layout Check

Render and inspect every slide using the presentation-authoring workflow before delivery. Verify that all elements stay inside the canvas and safe margins, no text is clipped or overlaps, repeated elements align, palette tokens remain consistent, and text caps and accent restraint are respected. Fix problems and render again before returning the deck.
