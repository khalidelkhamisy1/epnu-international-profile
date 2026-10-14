"""Labeled thumbnail sheets (5x4 per sheet) for a folder of images, so a reviewer can scan many images at once.
Usage: python thumb_sheets.py <image_dir> <out_dir> [prefix]"""
import sys, os, glob, json
from PIL import Image, ImageDraw, ImageFont

src, out = sys.argv[1], sys.argv[2]
prefix = sys.argv[3] if len(sys.argv) > 3 else "sheet"
os.makedirs(out, exist_ok=True)
files = sorted(f for f in glob.glob(os.path.join(src, "*")) if f.lower().endswith((".jpg", ".jpeg", ".png", ".webp")))
cols, rows, tw, th = 5, 4, 320, 220
font = ImageFont.load_default()
index = {}
for k in range(0, len(files), cols * rows):
    sheet = Image.new("RGB", (cols * tw, rows * (th + 22)), "white")
    d = ImageDraw.Draw(sheet)
    for i, f in enumerate(files[k:k + cols * rows]):
        n = k + i
        try:
            im = Image.open(f); w, h = im.size; im = im.convert("RGB"); im.thumbnail((tw - 6, th - 6))
        except Exception:
            continue
        x, y = (i % cols) * tw, (i // cols) * (th + 22)
        sheet.paste(im, (x + (tw - im.width) // 2, y + (th - im.height) // 2))
        d.text((x + 4, y + th + 4), f"#{n}  {w}x{h}", fill="black", font=font)
        index[n] = os.path.basename(f)
    sheet.save(os.path.join(out, f"{prefix}-{k // (cols * rows) + 1:02d}.jpg"), quality=80)
json.dump(index, open(os.path.join(out, f"{prefix}-index.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
print(len(files), "images ->", (len(files) + cols * rows - 1) // (cols * rows), "sheets")
