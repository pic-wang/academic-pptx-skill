---
name: academic-pptx
description: "Use this skill whenever the user wants to create, update, or improve a .pptx presentation for an academic context: conference talks, seminar slides, group meetings / lab meetings / weekly reports, thesis defenses, grant briefings, invited lectures. Triggers include 'slides', 'deck', 'presentation', 'pptx', 'group meeting', 'conference talk', 'thesis defense', and Chinese requests such as '做slides', '组会', '汇报'. The skill governs CONTENT and STRUCTURE (argument, action titles, deck architecture) and the house VISUAL STYLE (Cambria, red C00000 accent, red dashed take-home box, three-line tables). Build files with python-pptx through pptx_style.py."
license: MIT (see LICENSE)
---

# Academic Presentations Skill (house style)

## How This Skill Works

This skill has three layers. Read all three before writing any code.

| Layer | File | Governs |
|-------|------|---------|
| Content | [content_guidelines.md](content_guidelines.md) | Argument structure, action titles, exhibits, citations, deck architecture, timing |
| Visual style | [STYLE.md](STYLE.md) | Canvas, font, palette, type scale, take-home box, tables, captions, agenda slide, QA |
| Implementation | [slide_patterns.md](slide_patterns.md) + [pptx_style.py](pptx_style.py) | One python-pptx pattern per slide type, built from the helpers |

`examples/build_demo.py` builds a complete demo deck with every helper. Run it once if unsure how the helpers combine.

If an Anthropic `pptx` skill is also installed, use it only for generic file operations (reading existing decks, `markitdown` extraction). Its design defaults do not apply; STYLE.md overrides them.

---

## Step 1: Identify Presentation Type

### Structured Argument (default)
Conference papers, seminar talks, group meetings, thesis defenses, grant briefings, policy or consulting-style research deliverables.

**Priority order: argument structure → data → layout → aesthetics.** Follow content_guidelines.md in full.

### Visual / Narrative
Public engagement talks and lay audiences. Argument structure still matters, but more slides may be image-led. The visual style in STYLE.md still applies.

### When in doubt
Default to Structured Argument. A paper, study, dataset, thesis, grant, or conference implies Structured Argument.

---

## Step 2: Plan the Deck Before Creating Any Slides

Produce a slide-by-slide outline with four columns: page title (action title), section, exhibit, take-home sentence. Confirm it with the user if the deck exceeds 10 slides or the content is complex. Do not build before the outline is agreed.

Run the ghost deck test on the outline: read only the page titles in sequence; they must tell the complete argument. Then read only the take-home sentences; they must also tell it, with the key numbers.

---

## Step 3: Apply the House Style

All visual rules live in [STYLE.md](STYLE.md). Summary:

- 13.333 × 7.5 in (the size of the reference decks), white background, **Cambria** everywhere.
- Palette: `1A1A1A` text, `000000` titles and table rules, `C00000` red accent, `F2F2F2` zebra gray, `6B6B6B` muted. No blue, green, or orange.
- Page title 28 pt bold black top-left, one line, no underline.
- Section header: red `•` + 20 pt bold red text.
- Body 20 pt (18 pt in side panels), key phrases in red / red bold.
- **Every content slide ends with a red dashed take-home box** stating its conclusion, key phrase in red bold.
- Tables: three black rules, no grid, zebra rows, key cells red bold, no `tableStyleId`.
- Captions 18 pt, `Fig. N.` / `Tab. N.` bold, single line. Citations 10.5 pt.
- An agenda ("Content") slide with the current section highlighted opens each major section.

Use the helpers in `pptx_style.py`: `new_presentation`, `title_slide`, `agenda_slide`, `content_slide`, `add_page_title`, `add_section_header`, `add_body`, `add_takehome_box`, `add_table`, `add_figure`, `add_caption`, `add_citation`, `add_concept_boxes`, `check_takehome`.

Institution images: if the untracked folder `assets/` exists (with `assets.json`), call `ps.house_assets()`. Pass its `polyu_logo` path as `logo=` on content slides, and `title_corner`, `title_slogan`, `lab_logo` as `decorations=` to `title_slide`. This reproduces the title page and logo of the reference decks.

Python plotting code must start with the user's matplotlib styling block (see the user's CLAUDE.md; `examples/build_demo.py` shows it).

---

## Step 4: Build and QA

1. Build the deck with a Python script that imports `pptx_style` (add this skill directory to `sys.path`).
2. Run `check_takehome(prs, exempt={...})` and fix any slide it reports.
3. Render: `python render.py deck.pptx out_dir` → one PNG per slide plus contact sheets.
4. Inspect every image against the QA list in STYLE.md §6 and the academic checklist below.
5. Fix and re-render until a full pass finds no new issue.

```
Academic QA checklist:
□ Every content slide has an action title (one line, states the takeaway)
□ Every content slide has a red dashed take-home box
□ Ghost deck test passes on titles and on take-home sentences
□ One exhibit per results slide; key finding annotated in red on the exhibit
□ Every borrowed figure or number has an in-slide citation
□ References slide exists
□ Conclusions slide is the last main slide, with contact info
□ Agenda slide before each major section (decks > 10 slides)
□ Body text ≥ 20 pt (≥ 18 pt in side panels); captions one line
□ No decorative elements that do not carry content
```

---

## Key Principles (Summary)

**Action titles, not topic labels.** Each page title is a short sentence stating the takeaway. The take-home box restates it with the number that supports it.

**One argument, made well.** Do not present the whole paper. Everything else goes in the appendix.

**One insight per slide.** One exhibit per results slide, with the key finding marked in red on the exhibit.

**Slides support speech.** Body text orients the audience; the presenter carries the argument.

**Cite everything borrowed.** Full reference at the bottom of the slide, complete list on the References slide.

**End on conclusions.** The conclusions slide stays on screen during Q&A.

---

## Dependencies

- `pip install python-pptx pillow` (and `matplotlib` for figures)
- Windows: `pywin32` + Microsoft PowerPoint for rendering
- Elsewhere: LibreOffice (`soffice`) and Poppler (`pdftoppm`)
- Optional: `pip install "markitdown[pptx]"` for text extraction from existing decks
