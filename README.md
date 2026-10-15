# Academic Presentations Skill for Claude (house style)

A Claude Skill for academic presentations (conference talks, group meetings, seminars, thesis defenses, grant briefings). It combines:

- the **content and argument rules** of the upstream skill [Gabberflast/academic-pptx-skill](https://github.com/Gabberflast/academic-pptx-skill): action titles, ghost deck test, one exhibit per slide, citations, deck architecture, timing;
- a **fixed house visual style** derived from three reference decks: Cambria, red `C00000` accent, red dashed take-home box on every content slide, three-line academic tables with zebra rows, agenda slides with the current section highlighted;
- a **python-pptx helper library** (`pptx_style.py`) that encodes every visual rule, and a renderer (`render.py`) for visual QA.

## Changes from upstream

| Area | Upstream | This fork |
|------|----------|-----------|
| Build tool | PptxGenJS (JavaScript) | python-pptx via `pptx_style.py` |
| Font | Arial / Calibri | Cambria everywhere |
| Palette | Navy `1F4E79`, blue `2E75B6`, yellow highlight | `1A1A1A` text, `000000` rules, `C00000` red accent, `F2F2F2` zebra, `6B6B6B` muted |
| Title / conclusions slides | Dark navy background | White background; red key phrase in the deck title |
| Title rule | Gray divider under every title | No line under titles |
| Slide conclusion | Action title only | Action title + red dashed take-home box on every content slide |
| Section headers | 22 pt blue | Red `•` + 16 pt bold red |
| Tables | Not specified | Three black rules, no grid, zebra rows, key cells red bold |
| Section navigation | Dark section dividers / breadcrumb bar | "Content" agenda slide, current section black, others gray |
| Citations | 12–14 pt muted line | Full reference, 8 pt muted, bottom-left |
| Font sizes | 26 pt title / 20 pt body | 22 pt title / 15 pt body on a 10 in canvas (= 20 pt on 13.33 in) |
| QA | markitdown + images | `check_takehome()` + `render.py` (PowerPoint COM or LibreOffice) |

The content rules in `content_guidelines.md` are kept from upstream, with sizes and annotation colours adjusted to the house style.

## File structure

```
academic-pptx/
├── SKILL.md               # Entry point: workflow and summary
├── content_guidelines.md  # Argument structure, action titles, take-home box, citations, deck architecture
├── STYLE.md               # Complete visual specification and QA list
├── slide_patterns.md      # python-pptx pattern per slide type
├── pptx_style.py          # Helper library implementing STYLE.md
├── render.py              # .pptx -> PNG per slide + contact sheets
├── examples/build_demo.py # Builds a 10-slide demo deck with every helper
└── Academic_Presentation_Best_Practices_...pdf  # Upstream general advice (its colour/font choices are superseded by STYLE.md)
```

## Installation

Claude Code:

```bash
git clone https://github.com/pic-wang/academic-pptx-skill.git ~/.claude/skills/academic-pptx
pip install python-pptx pillow matplotlib pywin32   # pywin32 only on Windows
```

claude.ai: download the repository as a zip and upload it under **Customize → Skills**.

## Usage

Ask naturally, e.g. "make slides for my group meeting on X" or "build a conference deck from this paper". To check the setup:

```bash
python examples/build_demo.py examples/demo.pptx [--template template.pptx] [--logo logo.png]
python render.py examples/demo.pptx examples/demo_render
```

Institution templates and logos are not stored in this repository. Pass them with `--template` / `--logo`, or `new_presentation(template=...)` / `logo=...` in code.

## License

MIT, as upstream.
