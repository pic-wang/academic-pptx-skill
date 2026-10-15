"""Render a .pptx to one PNG per slide plus contact sheets, for visual QA.

Usage:  python render.py deck.pptx [out_dir] [--allow-running]

Windows: PowerPoint COM (pywin32). PowerPoint runs as a single instance, so
COM attaches to the window the user is working in and Application.Quit()
would close all of their presentations. This script therefore refuses to run
while PowerPoint is open (pass --allow-running to open the file hidden in the
running instance and close only that file), and it calls Quit only when it
started PowerPoint itself and no other presentation is open.
Elsewhere: LibreOffice + pdftoppm.
"""
import glob
import os
import subprocess
import sys
from pathlib import Path


def _powerpoint_running():
    import pythoncom
    import win32com.client
    try:
        win32com.client.GetActiveObject("PowerPoint.Application")
        return True
    except pythoncom.com_error:
        return False


def render(pptx, out_dir, allow_running=False):
    pptx = Path(os.path.abspath(pptx))
    out = Path(os.path.abspath(out_dir))
    out.mkdir(parents=True, exist_ok=True)
    for f in glob.glob(str(out / "s[0-9][0-9].png")) + glob.glob(str(out / "sheet*.jpg")):
        os.remove(f)
    if os.name == "nt":
        import win32com.client
        running = _powerpoint_running()
        if running and not allow_running:
            raise SystemExit("PowerPoint is open (the user may be working in it): not rendering. "
                             "Ask the user to close it, or pass --allow-running to render "
                             "without closing anything.")
        app = win32com.client.Dispatch("PowerPoint.Application")
        pres = app.Presentations.Open(str(pptx), True, False, False)  # read-only, untitled, no window
        try:
            for i, s in enumerate(pres.Slides, 1):
                s.Export(str(out / f"s{i:02d}.png"), "PNG", 1600, 900)
        finally:
            pres.Close()                                   # closes only this file
            if not running and app.Presentations.Count == 0:
                app.Quit()                                 # only an instance this script started
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
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    src = args[0]
    dst = args[1] if len(args) > 1 else Path(src).with_suffix("")
    for p in render(src, dst, allow_running="--allow-running" in sys.argv):
        print(p)
