"""Render a .pptx to one PNG per slide plus a contact sheet, for visual QA.

Usage:  python render.py deck.pptx [out_dir]
Windows: PowerPoint COM (pywin32). Elsewhere: LibreOffice + pdftoppm.
"""
import glob
import os
import shutil
import subprocess
import sys
from pathlib import Path


def render(pptx, out_dir):
    pptx = Path(os.path.abspath(pptx))
    out = Path(os.path.abspath(out_dir))
    out.mkdir(parents=True, exist_ok=True)
    if os.name == "nt":
        import win32com.client
        app = win32com.client.Dispatch("PowerPoint.Application")
        pres = app.Presentations.Open(str(pptx), True, False, False)
        for i, s in enumerate(pres.Slides, 1):
            s.Export(str(out / f"s{i:02d}.png"), "PNG", 1600, 900)
        pres.Close()
        app.Quit()
    else:
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                        "--outdir", str(out), str(pptx)], check=True)
        pdf = out / (pptx.stem + ".pdf")
        subprocess.run(["pdftoppm", "-png", "-r", "120", str(pdf), str(out / "s")], check=True)
        for f in out.glob("s-*.png"):
            f.rename(out / f"s{int(f.stem.split('-')[1]):02d}.png")
    return contact_sheet(out)


def contact_sheet(out, cols=3, rows=3, w=640, h=360):
    from PIL import Image, ImageDraw
    files = sorted(glob.glob(str(out / "s[0-9][0-9].png")))
    sheets = []
    for k in range(0, len(files), cols * rows):
        im = Image.new("RGB", (w * cols, h * rows), "white")
        d = ImageDraw.Draw(im)
        for j, f in enumerate(files[k:k + cols * rows]):
            x, y = (j % cols) * w, (j // cols) * h
            im.paste(Image.open(f).convert("RGB").resize((w, h)), (x, y))
            d.rectangle([x, y, x + w - 1, y + h - 1], outline="gray")
            d.text((x + 5, y + 5), Path(f).name, fill="blue")
        p = out / f"sheet{k // (cols * rows) + 1}.jpg"
        im.save(p, quality=85)
        sheets.append(p)
    return sheets


if __name__ == "__main__":
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else Path(src).with_suffix("")
    for p in render(src, dst):
        print(p)
