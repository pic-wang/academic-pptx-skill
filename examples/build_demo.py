"""Build a demo deck that exercises every helper in pptx_style.py.

Usage:  python examples/build_demo.py [out.pptx] [--template template.pptx] [--logo logo.png]
All numbers in this deck are placeholder demo data.
"""
import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import numpy as np
from pathlib import Path
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib import font_manager

user_fonts = font_manager.findSystemFonts(fontpaths=None, fontext="ttf")
regular_fonts = [f for f in user_fonts if "Regular" in Path(f).name]
if regular_fonts:
    mpl.rcParams["font.family"] = "Geologica"
plt.rcParams["mathtext.fontset"] = "cm"
mpl.rcParams.update({
    'font.size': 15,
    'axes.labelsize': 15,
    'xtick.labelsize': 15,
    'ytick.labelsize': 15,
    'axes.spines.top': True,
    'axes.spines.right': True,
})

import pptx_style as ps


def make_figure(path):
    sizes = np.array([500, 1000, 2000, 3500])
    err_a = np.array([0.168, 0.152, 0.131, 0.115])
    err_b = np.array([0.205, 0.193, 0.181, 0.176])
    fig, ax = plt.subplots(figsize=(7.4, 4.4), dpi=200)
    ax.plot(sizes, err_b, "o-", color="#6B6B6B", label="Baseline")
    ax.plot(sizes, err_a, "o-", color="#C00000", label="Ours")
    ax.annotate("-35% vs. baseline", xy=(3500, 0.115), xytext=(2300, 0.150),
                color="#C00000", arrowprops=dict(arrowstyle="->", color="#C00000"))
    ax.set_xlabel("Training set size")
    ax.set_ylabel("Relative L2 error")
    ax.legend(frameon=False)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def build(out, template=None, logo=None):
    prs = ps.new_presentation(template)
    sections = ["Background", "Methodology", "Results", "Conclusions"]
    n = 1

    ps.title_slide(prs, "A **General Neural Solver** for Fast Flow Prediction under Arbitrary Inflow",
                   "Author A, Author B, Author C*", presenter="Author A",
                   affiliation="Example University", venue="Demo Conference 2026", logo=logo)
    n += 1
    ps.agenda_slide(prs, sections, current=0, page=n, logo=logo)

    n += 1
    s = ps.content_slide(prs, "CFD is too slow for iterative early-stage design", page=n, logo=logo)
    ps.add_section_header(s, "Current practice")
    ps.add_body(s, ["One CFD case takes **10³–10⁴ s** of compute",
                    "Design iterations need *hundreds* of cases",
                    "Wind tunnel tests give sparse measurements only"], h=1.6)
    ps.add_concept_boxes(s, [("Applicability", ["Real-world layouts"]),
                             ("Flexibility", ["Arbitrary inflow"]),
                             ("Accuracy", ["Benchmarked vs. CFD"])], y=3.85, h=1.3)
    ps.add_citation(s, "A. Author, B. Author, Example reference title, Journal 1 (2025) 1–10.")
    ps.add_takehome_box(s, "A **faster assessment tool** is needed during early-stage design.")

    n += 1
    s = ps.content_slide(prs, "Research question", page=n, logo=logo)
    ps.add_section_header(s, "Question")
    ps.add_body(s, ["Can one neural solver predict 3D flow for *unseen* inflow profiles "
                    "with error below 0.15 relative L2?"], h=1.1, bullets=False)
    ps.add_section_header(s, "Contribution", y=3.2)
    ps.add_body(s, ["Parametric inflow generator covering realistic profiles",
                    "Conditional modulation of a transformer backbone"], y=3.7, h=1.6)
    ps.add_takehome_box(s, "Target: **generalize across inflow conditions** without retraining.")

    n += 1
    ps.agenda_slide(prs, sections, current=2, page=n, logo=logo)

    n += 1
    s = ps.content_slide(prs, "Error falls 35% below baseline with more data", page=n, logo=logo)
    fig = ROOT / "examples" / "_demo_fig.png"
    make_figure(fig)
    ps.add_figure(s, fig, 0.5, 1.3, w=7.4, h=4.9, caption=("Fig. 1.", "Test error vs. training set size."))
    ps.add_section_header(s, "Observations", x=8.2, w=4.6)
    ps.add_body(s, ["Error decreases monotonically", "Baseline saturates at **0.176**"],
                x=8.51, w=4.3, size=18)
    ps.add_takehome_box(s, "More data keeps helping: **no saturation** up to 3,500 cases.")

    n += 1
    s = ps.content_slide(prs, "Our model has the lowest error on all four metrics", page=n, logo=logo)
    ps.add_section_header(s, "Benchmark on the test set")
    ps.add_table(s, ["Model", "RL2E", "MAE(u)", "MAE(v)", "MAE(w)"],
                 [["Baseline A", "0.319", "0.480", "0.163", "0.072"],
                  ["Baseline B", "0.176", "0.215", "0.096", "0.044"],
                  ["Ours", "0.115", "0.122", "0.067", "0.031"]],
                 x=1.67, y=2.1, w=10.0, col_widths=[3.2, 1.7, 1.7, 1.7, 1.7],
                 highlight=[(2, 1), (2, 2), (2, 3), (2, 4)])
    ps.add_caption(s, "Tab. 1.", "Errors on 400 held-out layouts.", 1.67, 4.0, 10.0)
    ps.add_takehome_box(s, "RL2E drops from 0.176 to **0.115** (-34.6%).")

    n += 1
    ps.agenda_slide(prs, sections, current=3, page=n, logo=logo)

    n += 1
    s = ps.content_slide(prs, "Conclusions", page=n, logo=logo)
    ps.add_body(s, ["We introduced a neural solver for flow under **arbitrary inflow conditions**."],
                x=0.5, y=1.3, w=12.3, h=0.6, bullets=False)
    ps.add_body(s, ["Lowest error among baselines: RL2E **0.115**",
                    "Error stays within *2x* on out-of-distribution inflow",
                    "Validated against two wind tunnel benchmarks"], y=2.1, h=2.4)
    ps.add_body(s, ["Contact: author@example.edu  |  Preprint: example.org/paper"],
                y=5.2, h=0.5, size=16, bullets=False)
    ps.add_takehome_box(s, "One model, **any inflow**, CFD-level accuracy in seconds.")

    n += 1
    s = ps.content_slide(prs, "References", page=n, logo=logo)
    ps.add_body(s, ["A. Author, B. Author, Example reference one, Journal 1 (2025) 1–10.",
                    "C. Author, Example reference two, Conference (2024)."],
                x=0.5, y=1.3, w=12.3, size=14, bullets=False)

    prs.save(out)
    missing = ps.check_takehome(prs, exempt={1, 2, 5, 8, n})
    print("saved", out, "| slides missing take-home box:", missing or "none")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("out", nargs="?", default=str(ROOT / "examples" / "demo.pptx"))
    ap.add_argument("--template")
    ap.add_argument("--logo")
    a = ap.parse_args()
    build(a.out, a.template, a.logo)
