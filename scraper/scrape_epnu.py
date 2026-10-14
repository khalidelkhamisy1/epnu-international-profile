"""Scrape EPNU (epnu.edu.eg) via the WordPress REST API.

Pulls every page and post from both the English (/en/) and Arabic (root) sites,
converts HTML to clean text, and saves:
  data/raw/<site>_<type>.json   full API records (title, link, date, text)
  data/text/<site>/<type>/<slug>.txt   one readable text file per item
Usage: python scrape_epnu.py [output_dir]
"""
import json
import re
import sys
import time
from pathlib import Path
from urllib.parse import unquote

import requests
from bs4 import BeautifulSoup

SITES = {
    "en": "https://epnu.edu.eg/en/wp-json/wp/v2",
    "ar": "https://epnu.edu.eg/wp-json/wp/v2",
}
TYPES = ["pages", "posts"]
HEADERS = {"User-Agent": "Mozilla/5.0 (EPNU research scraper)"}


def html_to_text(html):
    soup = BeautifulSoup(html or "", "lxml")
    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()
    text = soup.get_text("\n")
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines()]
    return "\n".join(ln for ln in lines if ln)


def fetch_all(base, kind):
    items, page = [], 1
    while True:
        r = requests.get(
            f"{base}/{kind}",
            params={"per_page": 100, "page": page},
            headers=HEADERS,
            timeout=60,
        )
        if r.status_code == 400:  # past last page
            break
        r.raise_for_status()
        batch = r.json()
        if not batch:
            break
        items.extend(batch)
        if page >= int(r.headers.get("X-WP-TotalPages", 1)):
            break
        page += 1
        time.sleep(0.5)
    return items


def main():
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "data")
    (out / "raw").mkdir(parents=True, exist_ok=True)
    for site, base in SITES.items():
        for kind in TYPES:
            try:
                items = fetch_all(base, kind)
            except Exception as e:  # keep going if one endpoint fails
                print(f"[{site}/{kind}] failed: {e}")
                continue
            records = []
            text_dir = out / "text" / site / kind
            text_dir.mkdir(parents=True, exist_ok=True)
            for it in items:
                title = html_to_text(it.get("title", {}).get("rendered", ""))
                body = html_to_text(it.get("content", {}).get("rendered", ""))
                rec = {
                    "id": it["id"],
                    "date": it.get("date"),
                    "link": unquote(it.get("link", "")),
                    "title": title,
                    "text": body,
                }
                records.append(rec)
                slug = re.sub(r"[^\w\-]+", "_", unquote(it.get("slug", "")))[:80]
                (text_dir / f"{it['id']}_{slug}.txt").write_text(
                    f"TITLE: {title}\nURL: {rec['link']}\nDATE: {rec['date']}\n\n{body}\n",
                    encoding="utf-8",
                )
            (out / "raw" / f"{site}_{kind}.json").write_text(
                json.dumps(records, ensure_ascii=False, indent=1), encoding="utf-8"
            )
            print(f"[{site}/{kind}] {len(records)} items")
    scrape_rendered_en(out)


def scrape_rendered_en(out):
    """The /en/ site is machine-translated at render time (TranslatePress), so the
    API returns Arabic. Fetch each rendered /en/ page to capture the English text."""
    import trafilatura

    for kind in TYPES:
        src = out / "raw" / f"en_{kind}.json"
        if not src.exists():
            continue
        records = json.loads(src.read_text(encoding="utf-8"))
        text_dir = out / "text" / "en_rendered" / kind
        text_dir.mkdir(parents=True, exist_ok=True)
        for rec in records:
            try:
                r = requests.get(rec["link"], headers=HEADERS, timeout=60)
                r.raise_for_status()
                rec["text_en"] = trafilatura.extract(
                    r.text, include_tables=True, include_links=False, favor_recall=True
                ) or ""
            except Exception as e:
                print(f"  render failed {rec['link']}: {e}")
                rec["text_en"] = ""
            (text_dir / f"{rec['id']}.txt").write_text(
                f"TITLE: {rec['title']}\nURL: {rec['link']}\nDATE: {rec['date']}\n\n{rec['text_en']}\n",
                encoding="utf-8",
            )
            time.sleep(0.3)
        (out / "raw" / f"en_rendered_{kind}.json").write_text(
            json.dumps(records, ensure_ascii=False, indent=1), encoding="utf-8"
        )
        print(f"[en_rendered/{kind}] {len(records)} items")


if __name__ == "__main__":
    main()
