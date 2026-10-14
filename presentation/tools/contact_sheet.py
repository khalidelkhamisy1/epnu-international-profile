"""Tile rendered slides into contact sheets (6 per sheet) for quick review."""
import sys, glob, os
from PIL import Image
src = sys.argv[1] if len(sys.argv) > 1 else "../render"
files = sorted(glob.glob(os.path.join(src, "slide-*.png")))
per, cols, tw = 6, 2, 900
for k in range(0, len(files), per):
    chunk = files[k:k + per]
    ims = [Image.open(f).convert("RGB") for f in chunk]
    th = int(tw * 9 / 16)
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * tw + (cols + 1) * 10, rows * th + (rows + 1) * 10), "white")
    for i, im in enumerate(ims):
        sheet.paste(im.resize((tw, th)), (10 + (i % cols) * (tw + 10), 10 + (i // cols) * (th + 10)))
    out = os.path.join(src, f"sheet-{k // per + 1}.jpg")
    sheet.save(out, quality=85)
    print(out)
