"""Collect images actually shown on EPNU's main pages (homepage, president, faculties, student life...).
Downloads full-size versions to assets/site_images/ and writes site_images.json (file, url, page, w, h)."""
import json, re, sys, time
from pathlib import Path
from urllib.parse import unquote
import requests
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "assets" / "site_images"
OUT.mkdir(parents=True, exist_ok=True)
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0"}
pages = json.loads(Path(r"D:\AI\EPNU\data\raw\en_pages.json").read_text(encoding="utf-8"))
urls = [p["link"] for p in pages]
rx = re.compile(r"https://epnu\.edu\.eg/wp-content/uploads/[^\"'\s)]+?\.(?:jpe?g|png|webp)", re.I)

seen, records = {}, []
for page in urls:
    try:
        html = requests.get(page, headers=H, timeout=60).text
    except Exception as e:
        print("page fail", page, e); continue
    for u in set(rx.findall(html)):
        full = re.sub(r"-\d+x\d+(\.\w+)$", r"\1", u)
        if "/elementor/thumbs/" in full:
            full = u
        seen.setdefault(full, page)
    time.sleep(0.3)

print(len(seen), "unique image urls")
for i, (u, page) in enumerate(seen.items()):
    name = re.sub(r"[^\w.\-]+", "_", unquote(u.rsplit("/", 1)[-1]))[-90:]
    f = OUT / name
    try:
        if not f.exists():
            r = requests.get(u, headers=H, timeout=120); r.raise_for_status()
            f.write_bytes(r.content); time.sleep(0.2)
        with Image.open(f) as im: w, h = im.size
    except Exception as e:
        print("fail", u, e); continue
    if w < 400 and "logo" not in name.lower():
        f.unlink(missing_ok=True); continue
    records.append({"file": name, "url": unquote(u), "page": unquote(page), "w": w, "h": h})
(ROOT / "assets" / "site_images.json").write_text(json.dumps(records, ensure_ascii=False, indent=1), encoding="utf-8")
print(len(records), "kept")
