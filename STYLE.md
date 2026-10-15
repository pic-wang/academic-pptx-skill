# House Visual Style

This file defines every visual rule for decks built with this skill. It was derived from three reference decks (a conference talk, a group-meeting report, and an invited talk) and replaces the upstream skill's navy/Arial design standards.

`pptx_style.py` implements every rule below. Build slides by calling its helpers; do not set fonts, colours, or borders by hand.

Reproducibility takes priority over creativity. Do not invent alternate palettes, add decorative lines, or switch fonts.

---

## 1. Canvas and font

| Item | Value |
|------|-------|
| Canvas | 13.333 × 7.5 in (16:9), the size of the reference decks. Base on `template.pptx` when one is provided. |
| Font | **Cambria** for every run, including East-Asian (`a:ea`) and complex-script (`a:cs`) runs. |
| Background | White on every slide, including the title and conclusions slides. |

Sizes and positions in §3 were measured from the three reference decks (median over all slides). Where the decks differ, the value used in two of the three decks is chosen and the other value is noted.

## 2. Palette

| Token | Hex | Use |
|-------|-----|-----|
| Text | `1A1A1A` | Body text |
| Black | `000000` | Page titles, table header text, table rules |
| Red accent | `C00000` | Section headers, key phrases, key table cells, take-home border, concept-box borders |
| Gray | `F2F2F2` | Zebra stripe on table data rows |
| Muted | `6B6B6B` | Citations, venue line |
| Agenda inactive | `BFBFBF` | Non-current sections on the agenda slide |

No blue, no green, no orange in slide chrome. Figures follow the same rule where you control them: red `C00000` for the focal series, gray `6B6B6B` for baselines.

## 3. Type scale and positions (13.333 × 7.5 in canvas)

| Element | Size | Weight | Colour | Position |
|---------|------|--------|--------|----------|
| Page title | 28 pt (the 2026-09 conference deck used 32 pt) | Bold | Black | Top-left `(0.5, 0.5)`, single line (≤ ~54 characters); box width 10 in keeps the logo corner free |
| Section header | 20 pt | Bold | Red, with red `•` bullet | `(0.5, 1.5)`; further headers below the body they introduce |
| Lead sentence | 20 pt | Regular; inline label (e.g. "Objective:") red bold | Text | `(0.5, 1.3)` on slides without a section header |
| Body | 20 pt (18 pt in side panels next to a figure) | Regular | Text; key phrases red | Bullet indent 0.31 in, aligned under the section-header text |
| Take-home box | 20 pt | Regular; key phrase red bold | — | Bottom, centred; bottom edge at `y = 6.85` |
| Table cells | 18 pt (compact side tables 14 pt) | Header bold | Black / Text; key cells red bold | Row height 0.45 in (compact 0.34 in) |
| Caption / label under an image | 18 pt | `Fig. N.` / `Tab. N.` bold, rest regular | Text | Directly beneath the exhibit, single line |
| Labels inside diagrams | 14 pt | Regular; "Input:" / "Output:" bold | Text / red | Inside figures |
| Concept box | Header 18 pt, bullets 16 pt | Header bold | Text | Row of boxes, ~1.8–2.2 in tall |
| Citation | 10.5 pt | Regular | Muted | Bottom-left `(0.1, 7.02)`, full reference |
| Page number | 14 pt | Regular | Text | Bottom-right |
| Agenda items | 24 pt | Bold | Black (current) / `BFBFBF` (others) | Line spacing 200 %, box at `(1.9, 1.7)` |
| Deck title | 28 pt | Bold | Black; key phrase red | Centred, `y ≈ 2.6–4.2` |
| Authors / presenter block | 20 pt | Authors regular; "Presenter:" and affiliation bold | Black | Authors centred at `y = 4.6`; presenter block from `(8.0, 6.3)` |
| Logo | Width 2.6 in | — | — | Top-right corner |

## 4. Slide elements

### Page title
Black bold Cambria, top-left, one line. No accent line or rule under it.

### Section header
Red `•` bullet (`<a:buChar char="•"/>` inside `<a:pPr>`, `buClr` red) followed by red bold text, usually ending in a colon in the reference decks ("Benchmark:", "Conditional modulation:"). Body text follows on the next lines, indented.

### Key phrases
Inside body text, mark the two to five words that carry the point in red (`*text*`) or red bold (`**text**`). Do not colour whole sentences red except inside the take-home box when the sentence itself is the key phrase.

### Take-home box (mandatory on every content slide)
- Text box with `<a:ln w="19050">` (1.5 pt), solid fill `C00000` on the line, `<a:prstDash val="dash"/>`, no shape fill.
- One sentence, centred, at the bottom of the slide.
- The key phrase inside is a red bold run; the rest is regular text colour.
- Exempt slides: title, agenda, references, appendix divider.

### Tables: three-line academic style
- Black top rule, black rule under the header row, black bottom rule. No internal grid, no vertical lines.
- Header row: bold, no fill.
- Data rows: zebra stripes, first data row `F2F2F2`, then white, alternating.
- Column 0 left-aligned; columns 1..n centred.
- Key-number cells red bold.
- Never keep a `tableStyleId`; `add_table` removes it.
- OOXML order: in `<a:tcPr>`, `lnL`, `lnR`, `lnT`, `lnB` must come **before** `<a:solidFill>`. Otherwise PowerPoint ignores the borders and draws the default grid. `_apply_three_line_borders` writes them in this order.

### Figures and captions
- **Figures fill the content area** between the title and the take-home box. A small figure inside a mostly empty slide is not acceptable; enlarge the figure or crop its whitespace, and cut words instead.
- Figure on the left, interpretation (section header + 2–3 bullets) on the right, when a slide has one exhibit plus commentary.
- Geometry and partition figures: show representative cases at the same scale so relative sizes are visible, in neatly aligned columns of equal height.
- Caption under the figure: `Fig. N.` bold + one regular sentence. Must fit on one line; shorten it if it wraps.
- Short descriptive labels under a row of images (e.g. "Pedestrian level comfort", "(Lyu et al., 2025)") are 18 pt, centred under each image (`add_caption` with `label=""`).
- Annotate the focal point inside the figure in red (`C00000`): arrow, circle, or a label such as `RL2E: 0.12`.

### Concept boxes
A row of two to four rounded rectangles with a 1.5 pt red border and no fill. Bold centred header, then two or three short bullets. Use them for objective/property summaries (e.g. Applicability · Flexibility · Accuracy).

### Agenda ("Content") slide
Page title "Content". List of the deck's sections; the current section is black bold with a filled `•`, the others are `BFBFBF` bold with a hollow `o` (Courier New) bullet. Insert it before each major section. This replaces the upstream breadcrumb bar and dark section dividers.

### Title slide
White background. Deck title centred, bold, black, with the key phrase in red. Author list centred below (corresponding author marked `*`). Presenter block bottom-right: a red vertical bar before "**Presenter:** Name" and before the bold affiliation line. Institution logo top-right if available. No dark background, no accent rule.

### Citations
Full reference (authors, title, venue, year, DOI) in 10.5 pt muted text at the bottom-left of the slide where the source is used. Short in-text form "(Author et al., Year)" may also appear beside the figure.

## 5. Never

- Accent lines under page titles.
- Filled or coloured table header rows.
- Red table rules (red is for text and box accents only).
- Centre-aligned body paragraphs (take-home box, captions, and the deck title are the only centred text).
- Emojis, icons in coloured circles, clip art, stock decoration.
- Dark backgrounds on any slide.
- A content slide without a take-home box.
- Fonts other than Cambria in slide text.
- Internal codes on slides (run names such as E1 / L120 / B0, batch names, case IDs such as F1b). Describe the item instead: "baseline, width 128 × 6", "Case E (CFD)", "K = 8 window".
- Writing to the user's own edited copy of a deck (e.g. a `*_update.pptx`). Write only the original file; if the user edited the original since your last write, merge instead of regenerating.

## 6. Visual QA (required before reporting done)

Render the deck (`python render.py deck.pptx out_dir`: PowerPoint COM on Windows, LibreOffice + `pdftoppm` elsewhere) and inspect every slide image.

**Never close the user's PowerPoint.** PowerPoint is single-instance: COM attaches to the user's window and `Application.Quit()` closes all of their presentations, including unsaved ones. `render.py` refuses to run while PowerPoint is open, closes only its own file, and quits only an instance it started. If PowerPoint is open, ask the user before rendering, or check the deck through python-pptx (text, positions, sizes) instead.

Inspect every slide image for:

- red dashed take-home box present on each content slide (`check_takehome()` also checks this);
- black three-line tables, no internal grid, zebra stripes;
- red bullets on section headers;
- page titles and captions on one line;
- no element overlap and nothing outside the slide;
- no text below 14 pt except citations.
