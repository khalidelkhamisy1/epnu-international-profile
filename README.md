# East Port Said National University (EPNU) — International Profile

An English PowerPoint profile of East Port Said National University (https://epnu.edu.eg/en/). It is built for International Relations work with foreign partner universities and delegations. Every fact in it is backed by a research base with cited sources.

## Deliverables
- `presentation/EPNU_International_Profile.pptx`: the 27-slide deck, with speaker notes and caveats for the presenter.
- `presentation/EPNU_International_Profile_SHARE.pptx`: the same deck without notes, ready to share with partners.
- `EPNU_International_Relations_Briefing.md`: the fact-checked briefing that serves as the single source of truth (sources S1–S51).

## Structure
| Path | Contents |
|---|---|
| `scraper/` | WordPress REST and rendered-page scraper for epnu.edu.eg (EN and AR) |
| `data/` | Scraped pages and posts (raw JSON and text) and findings from the official Facebook page |
| `research/` | Subagent research: faculty profiles, program-director bios, publication analysis (OpenAlex, Crossref, Europe PMC), extra findings |
| `review/` | Fact-check and completeness reviews of the briefing |
| `presentation/content/` | Slide content (faculty pages, research slide) and speaker notes |
| `presentation/assets/` | Logo and the photos used in the deck, with selection JSON |
| `presentation/tools/` | Deck generator (`build_deck.js`, pptxgenjs), PowerPoint renderer (`render.ps1`), image harvesting and contact-sheet tools |
| `presentation/review/` | Visual and text QA reports |

## Rebuild the deck
```bash
cd presentation/tools
npm install
PPTX_SKILL=<path to pptx skill> NODE_PATH=./node_modules node build_deck.js            # with notes
NO_NOTES=1 PPTX_SKILL=<path to pptx skill> NODE_PATH=./node_modules node build_deck.js # share version
```
The full image library (about 1,000 photos) is not committed. Regenerate it with `tools/harvest_images.py` and `tools/page_images.py`.

## Caveats
- The publication count (81) comes from raw-affiliation matching in open bibliographic databases. It is not a Scopus-verified figure.
- Some program-director credentials are probable matches only. See `research/bio_*.json` and the speaker notes.
