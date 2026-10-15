# Slide Patterns for Academic Presentations

One python-pptx pattern per required slide type, built from the helpers in [pptx_style.py](pptx_style.py). Visual rules are in [STYLE.md](STYLE.md); do not restate fonts or colours in slide code.

Coordinates are inches on the 13.333 × 7.5 in canvas. Give `add_figure` both `w` and `h`: the image is fitted inside that box (aspect kept), so it never spills into the take-home box or off the slide. Markup in any text argument: `**red bold**`, `*red*`.

---

## Global Setup

```python
import sys
sys.path.insert(0, r"<path to this skill>")
import pptx_style as ps

prs = ps.new_presentation(template=None)   # or "template.pptx"
LOGO = None                                 # optional institution logo PNG (top-right)
SECTIONS = ["Background", "Methodology", "Results", "Conclusions"]
page = 0
```

Layout grid used by every pattern:

| Region | y (in) | Helper |
|--------|--------|--------|
| Page title | 0.50 | `content_slide` / `add_page_title` |
| Lead sentence (no section header) | 1.30 | `add_body(..., y=1.3, bullets=False)` |
| First section header | 1.50 | `add_section_header` |
| Body under header | header y + 0.50 | `add_body` |
| Exhibit area | 1.3 – 6.2 | `add_figure`, `add_table`, `add_concept_boxes` |
| Take-home box | bottom edge at 6.85 | `add_takehome_box` |
| Citation | 7.02 | `add_citation` |
| Page number | bottom-right | `content_slide(page=...)` |

Columns for two-column slides: left column x = 0.5, w ≈ 6.0; right column x ≈ 6.9, w ≈ 5.9. For a figure-dominant slide: figure x = 0.5, w ≈ 7.4; text column x = 8.2, w ≈ 4.6.

---

## 1. Title Slide

**Purpose:** Establish the presentation and give complete attribution.

```python
ps.title_slide(prs,
    "A **General Neural Solver** for Efficient Flow Prediction under Arbitrary Inflow",
    "Author A, Author B, Author C*",
    presenter="Author A", affiliation="Example University",
    venue="Conference name · Month Year", logo=LOGO)
page += 1
```

White background, centred title with the key phrase in red, red bars before the presenter and affiliation lines. No take-home box.

---

## 2. Agenda Slide (replaces section dividers and breadcrumb bars)

**Purpose:** Orient the audience at the start of each major section.

```python
page += 1
ps.agenda_slide(prs, SECTIONS, current="Background", page=page, logo=LOGO)
```

Insert after the title slide and before each later section. Current section black with `•`; others gray with `o`. No take-home box.

---

## 3. Motivation / Context Slide

**Purpose:** Situation and complication: why the question matters and what is missing.

```python
page += 1
s = ps.content_slide(prs, "CFD is too slow for iterative early-stage design", page=page, logo=LOGO)
ps.add_section_header(s, "Current practice:")
ps.add_body(s, [
    "One CFD case takes **10³–10⁴ s** of compute",
    "Design iterations need *hundreds* of cases",
], h=1.2)
# Optional: row of images with short centred labels underneath
# ps.add_figure(s, "a.png", 0.8, 3.4, w=3.7, h=2.7, caption=("", "Pedestrian level comfort"))
ps.add_citation(s, "A. Author, B. Author, Title, Journal 1 (2025) 1–10. https://doi.org/...")
ps.add_takehome_box(s, "A **faster assessment tool** is needed during early-stage design.")
```

---

## 4. Research Question / Objective Slide

**Purpose:** Give the audience the anchor for everything that follows. This slide must exist on its own.

```python
page += 1
s = ps.content_slide(prs, "Neural solvers can replace CFD for design screening", page=page, logo=LOGO)
ps.add_body(s, ["**Objective:** replace the numerical solver with a neural solver for wind "
                "assessment in urban design scenarios."], x=0.5, y=1.3, w=12.3, bullets=False)
ps.add_figure(s, "pipeline.png", 0.5, 2.15, w=12.33, h=2.0)   # input → model → output schematic
ps.add_concept_boxes(s, [
    ("Applicability", ["Real-world building layouts", "Direct 3D prediction"]),
    ("Flexibility",   ["Real-time assessment", "Arbitrary inflow conditions"]),
    ("Accuracy",      ["Key flow structures", "Benchmarked vs. CFD and wind tunnel"]),
], y=4.3, h=1.8)
ps.add_takehome_box(s, "Goal: **one model for arbitrary layouts and inflow conditions**.")
```

---

## 5. Methods Slide

**Purpose:** What the audience needs to evaluate the findings, not procedural completeness. 1–2 slides; detail goes to the appendix.

Two-column pattern (labelled blocks left, figure right), as in the reference "model design" slides:

```python
page += 1
s = ps.content_slide(prs, "Modulation injects inflow into the backbone", page=page, logo=LOGO)
ps.add_section_header(s, "Backbone network:", w=6.0)
ps.add_body(s, ["Physics-attention transformer for canonical flow patterns [1]"],
            w=5.7, h=1.1, bullets=False)
ps.add_section_header(s, "Conditional modulation:", y=3.2, w=6.0)
ps.add_body(s, ["Geometry and inflow encoders produce *scale & shift* per layer"],
            y=3.7, w=5.7, h=1.1, bullets=False)
ps.add_figure(s, "architecture.png", 6.9, 1.3, w=5.9, h=4.8)
ps.add_citation(s, "[1] H. Wu et al., Transolver: A Fast Transformer Solver for PDEs on General Geometries (2024).")
ps.add_takehome_box(s, "Geometry and inflow are learned **separately**, then fused by modulation.")
```

Equations: insert them as images rendered at ≥ 20 pt equivalent, or use PowerPoint's equation editor afterwards. Define every symbol on the same slide.

---

## 6. Results Slide: Figure

**Purpose:** One finding, impossible to miss. Figure left (~7.4 in), interpretation right (~4.6 in).

```python
page += 1
s = ps.content_slide(prs, "Error falls 35% below baseline with more data", page=page, logo=LOGO)
ps.add_figure(s, "scaling.png", 0.5, 1.3, w=7.4, h=4.9,
              caption=("Fig. 1.", "Test error vs. training set size."))
ps.add_section_header(s, "Observations:", x=8.2, w=4.6)
ps.add_body(s, ["Error decreases monotonically", "Baseline saturates at **0.176**"],
            x=8.51, w=4.3, size=18)
ps.add_takehome_box(s, "More data keeps helping: **no saturation** up to 3,500 cases.")
```

The figure itself marks the key point in red (arrow, circle, or label such as `RL2E: 0.12`). Build figures with the user's matplotlib block; focal series `#C00000`, baselines `#6B6B6B`.

---

## 7. Results Slide: Table

**Purpose:** Precise numerical comparison where exact values matter.

```python
page += 1
s = ps.content_slide(prs, "Our model has the lowest error on all four metrics", page=page, logo=LOGO)
ps.add_section_header(s, "Benchmark:")
ps.add_table(s, ["Model", "RL2E", "MAE(u)", "MAE(v)", "MAE(w)"],
             [["Baseline A", "0.319", "0.480", "0.163", "0.072"],
              ["Baseline B", "0.176", "0.215", "0.096", "0.044"],
              ["Ours",       "0.115", "0.122", "0.067", "0.031"]],
             x=1.67, y=2.1, w=10.0, col_widths=[3.2, 1.7, 1.7, 1.7, 1.7],
             highlight=[(2, 1), (2, 2), (2, 3), (2, 4)])
ps.add_caption(s, "Tab. 1.", "Errors on 400 held-out layouts.", 1.67, 4.0, 10.0)
ps.add_takehome_box(s, "RL2E drops from 0.176 to **0.115** (-34.6%).")
```

`highlight` takes (data-row, column) indices, header excluded. Do not colour whole rows. Compact tables beside a figure: `size=14, row_h=0.34`.

---

## 8. Discussion / Implications Slide

**Purpose:** Interpret findings, connect to the research question, address the main limitation.

```python
page += 1
s = ps.content_slide(prs, "Accuracy holds beyond the training inflow range", page=page, logo=LOGO)
ps.add_section_header(s, "Interpretation:")
ps.add_body(s, ["OOD error doubles (0.12 → 0.23) but stays below the best baseline"], h=1.1)
ps.add_section_header(s, "Main limitation:", y=3.3)
ps.add_body(s, ["**External validity:** trained on one city; transfer to other morphologies untested"],
            y=3.8, h=1.1)
ps.add_takehome_box(s, "The model **generalizes beyond the training distribution**; transfer across cities is untested.")
```

---

## 9. Conclusions Slide

**Purpose:** Restate 2–4 takeaways. Stays on screen during Q&A. White background, same layout as content slides.

```python
page += 1
s = ps.content_slide(prs, "Conclusions", page=page, logo=LOGO)
ps.add_body(s, ["We introduced a neural surrogate for 3D flow prediction under "
                "**arbitrary inflow conditions**:"], x=0.5, y=1.3, w=12.3, bullets=False)
ps.add_body(s, [
    "Outperforms all baselines in **accuracy**: RL2E 0.115, 34.6% below the best baseline",
    "**Consistent accuracy** across a wide range of inflow conditions, including OOD",
    "Validated against two wind tunnel benchmarks for **real-world assessments**",
], y=2.1, h=2.8)
ps.add_body(s, ["Contact: name@university.edu  |  Preprint: link"], y=5.2, size=16, bullets=False)
ps.add_takehome_box(s, "One model, **any inflow**, CFD-level accuracy in seconds.")
```

Do not follow it with "Thank You" or a blank slide. A QR code (`ps.add_figure(s, "qr.png", 11.3, 4.4, w=1.4, h=1.4)`) may sit on this slide.

---

## 10. References Slide

```python
page += 1
s = ps.content_slide(prs, "References", page=page, logo=LOGO)
ps.add_body(s, [
    "H. Wu, H. Luo, H. Wang, J. Wang, M. Long, Transolver: A Fast Transformer Solver for PDEs on General Geometries, ICML (2024).",
    "X.I.A. Yang, J. Sadique, R. Mittal, C. Meneveau, Exponential roughness layer ..., J. Fluid Mech. 789 (2016) 127–165.",
], x=0.5, y=1.3, w=12.3, h=5.4, size=14, bullets=False, space_after=10)
```

No take-home box.

---

## 11. Appendix Slide

**Purpose:** Pre-built answers for Q&A. Same layout as content slides; the title starts with the appendix label.

```python
page += 1
s = ps.content_slide(prs, "Appendix A: Modulation beats early fusion", page=page, logo=LOGO)
ps.add_section_header(s, "Conditioning mechanism:", w=6.0)
ps.add_table(s, ["Approach", "Train loss", "Test RL2E"],
             [["Early fusion", "0.0591", "0.1356"],
              ["Cross-attention", "0.0812", "0.1193"],
              ["Modulation", "0.0308", "0.1150"]],
             x=0.8, y=2.1, w=5.6, col_widths=[2.4, 1.6, 1.6], highlight=[(2, 1), (2, 2)])
ps.add_figure(s, "loss_curves.png", 6.9, 1.3, w=5.9, h=4.8)
ps.add_takehome_box(s, "Modulation gives the **lowest train and test error**.")
```

---

## Finish

```python
prs.save("deck.pptx")
print(ps.check_takehome(prs, exempt={1, 2, 10}))   # title, agenda(s), references
```

Then render and inspect:

```bash
python render.py deck.pptx deck_render/
```

## Quick Checks Before Delivery

```
□ Title slide: statement title with red key phrase, authors, presenter, affiliation, venue/date
□ Agenda slide before each major section
□ Every content slide: one-line action title + red dashed take-home box
□ Ghost deck test passes on titles and on take-home sentences
□ Results slides: one exhibit each, key finding marked in red on the exhibit
□ Tables: three black rules, zebra rows, key cells red bold, no grid
□ Captions: "Fig. N." / "Tab. N." bold, one line
□ Every borrowed figure/number: full reference at the slide bottom
□ References slide complete; Conclusions slide last main slide with contact info
□ Body text ≥ 20 pt (side panels ≥ 18 pt)
□ Rendered images inspected: no overlap, no wrapping titles, nothing off-slide
```
