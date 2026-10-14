import sys, os, json, re, time, html

# Under `python -I` user site-packages are ignored; add it explicitly so requests imports.
USER_SITE = r"C:\Users\user\AppData\Roaming\Python\Python314\site-packages"
if USER_SITE not in sys.path:
    sys.path.append(USER_SITE)

import requests

API = "https://epnu.edu.eg/wp-json/wp/v2/media"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
HEADERS = {"User-Agent": UA}
ALLOWED_MIME = {"image/jpeg", "image/png", "image/webp"}
MIN_WIDTH = 600
PAUSE = 0.3

BASE = r"D:\AI\EPNU\presentation"
IMG_DIR = os.path.join(BASE, "assets", "images")
CATALOG_JSON = os.path.join(BASE, "assets", "images_catalog.json")
CATALOG_MD = os.path.join(BASE, "assets", "images_catalog.md")
RAW = r"D:\AI\EPNU\data\raw"

errors = []


def strip_html(s):
    if not s:
        return ""
    s = re.sub(r"<[^>]+>", " ", s)
    s = html.unescape(s)
    return re.sub(r"\s+", " ", s).strip()


def enumerate_media():
    items = []
    page = 1
    total_pages = None
    while True:
        try:
            r = requests.get(API, params={"per_page": 100, "page": page},
                             headers=HEADERS, timeout=60)
            r.raise_for_status()
        except Exception as e:
            errors.append(f"media page {page}: {e}")
            break
        if total_pages is None:
            total_pages = int(r.headers.get("X-WP-TotalPages", "1"))
        items.extend(r.json())
        page += 1
        if page > total_pages:
            break
        time.sleep(PAUSE)
    return items


def pick_size(m):
    md = m.get("media_details") or {}
    sizes = md.get("sizes") or {}
    if "full" in sizes and sizes["full"].get("source_url"):
        s = sizes["full"]
        return s["source_url"], s.get("width"), s.get("height")
    if sizes:
        best = max(sizes.values(), key=lambda s: (s.get("width") or 0))
        if best.get("source_url"):
            return best["source_url"], best.get("width"), best.get("height")
    return m.get("source_url"), md.get("width"), md.get("height")


def load_titles():
    titles = {}
    for fn in ("en_rendered_posts.json", "ar_posts.json"):
        path = os.path.join(RAW, fn)
        try:
            with open(path, encoding="utf-8") as f:
                for p in json.load(f):
                    pid = str(p.get("id"))
                    t = strip_html(p.get("title", "")) if isinstance(p.get("title"), str) else ""
                    if pid and t and pid not in titles:
                        titles[pid] = t
        except Exception as e:
            errors.append(f"titles {fn}: {e}")
    return titles


def main():
    os.makedirs(IMG_DIR, exist_ok=True)
    os.makedirs(os.path.dirname(CATALOG_JSON), exist_ok=True)

    media = enumerate_media()
    total = len(media)

    imgs = [m for m in media if m.get("mime_type") in ALLOWED_MIME]
    kept = []
    for m in imgs:
        url, w, h = pick_size(m)
        if not url:
            continue
        name_blob = " ".join([
            os.path.basename(url), (m.get("title") or {}).get("rendered", "") or ""
        ]).lower()
        is_logo = "logo" in name_blob
        if (w or 0) < MIN_WIDTH and not is_logo:
            continue
        kept.append({
            "id": m["id"], "source_url": url, "width": w, "height": h,
            "alt": strip_html(m.get("alt_text", "")),
            "title": strip_html((m.get("title") or {}).get("rendered", "")),
            "caption": strip_html((m.get("caption") or {}).get("rendered", "")),
            "post_id": m.get("post"), "date": m.get("date"),
        })

    titles = load_titles()
    catalog = []
    total_bytes = 0
    downloaded = 0
    for rec in kept:
        basename = os.path.basename(requests.utils.urlparse(rec["source_url"]).path)
        fname = f"{rec['id']}_{basename}"
        dest = os.path.join(IMG_DIR, fname)
        try:
            if not os.path.exists(dest):
                r = requests.get(rec["source_url"], headers=HEADERS, timeout=120)
                r.raise_for_status()
                with open(dest, "wb") as f:
                    f.write(r.content)
                time.sleep(PAUSE)
            total_bytes += os.path.getsize(dest)
            downloaded += 1
        except Exception as e:
            errors.append(f"download {rec['id']} {rec['source_url']}: {e}")
            continue
        pid = str(rec["post_id"]) if rec["post_id"] else None
        catalog.append({
            "file": os.path.join("assets", "images", fname).replace("\\", "/"),
            "id": rec["id"], "width": rec["width"], "height": rec["height"],
            "alt": rec["alt"], "title": rec["title"], "caption": rec["caption"],
            "post_id": rec["post_id"], "context_title": titles.get(pid, "") if pid else "",
            "date": rec["date"],
        })

    with open(CATALOG_JSON, "w", encoding="utf-8") as f:
        json.dump(catalog, f, ensure_ascii=False, indent=2)

    lines = ["| File | WxH | Context title |", "|---|---|---|"]
    for c in catalog:
        ctx = (c["context_title"] or "")[:70].replace("|", "/")
        lines.append(f"| {os.path.basename(c['file'])} | {c['width']}x{c['height']} | {ctx} |")
    with open(CATALOG_MD, "w", encoding="utf-8") as f:
        f.write("# EPNU image catalog\n\n" + "\n".join(lines) + "\n")

    print(f"media_total={total}")
    print(f"images_jpeg_png_webp={len(imgs)}")
    print(f"kept={len(kept)}")
    print(f"downloaded={downloaded}")
    print(f"total_MB={total_bytes / (1024 * 1024):.2f}")
    print(f"errors={len(errors)}")
    for e in errors:
        print("ERR", e)


if __name__ == "__main__":
    main()
