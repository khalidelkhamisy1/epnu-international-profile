# EPNU Presentation — Plan & Design System

**Deliverable:** `EPNU_International_Profile.pptx` (16:9 widescreen, 13.33" × 7.5"), English, ~18 slides.
**Audience:** foreign partners and delegations. The tone is confident and partnership-oriented. Gaps are framed as partnership opportunities, and data caveats go in the speaker notes, not on the slides.
**Single source of truth:** `../EPNU_International_Relations_Briefing.md` (v2, fact-checked). No fact may appear on a slide unless it is in the briefing.

## Team
| Role | Model | Task |
|---|---|---|
| Planner / designer / builder | Opus 5.5 | This plan, the design system, slide copy, the pptxgenjs build, and fixes |
| Image harvester | Haiku 5.5 | Downloads all media from the WP REST API into `assets/images/` and writes `images_catalog.json` |
| Image curator | Sonnet 5.5 | Views the images, picks the best ones per slide, flags poor quality, and writes `assets/image_selection.json` |
| Speaker-notes writer | Sonnet 5.5 | Writes presenter notes per slide from the briefing, including the caveats, in `content/speaker_notes.json` |
| Visual QA | Sonnet 5.5 | Reviews the rendered slide PNGs for overflow, overlap, contrast and alignment |
| Text / consistency QA | Haiku 5.5 | Runs mechanical checks on the extracted text: typos, number consistency against the briefing, leftover placeholders |

## Design system
**University colors** (extracted from the epnu.edu.eg theme CSS by frequency):

| Token | Hex | Use |
|---|---|---|
| Navy (dominant) | `071C35` | Dark slides, titles on light slides |
| Deep blue | `1B3C73` | Secondary surfaces, cards on dark slides |
| Bright blue | `0077FF` | Primary accent: icons, key numbers |
| Cyan | `00D4FF` | Sharp accent, used sparingly: highlights on dark slides |
| Ice | `EEF6FF` | Light card fill |
| Off-white | `F8FAFC` | Light slide background |
| Slate | `64748B` | Captions, muted text |
| Ink | `0F172A` | Body text on light slides |

- **Structure:** a "sandwich". Dark navy is used for the title, section dividers and closing. Light off-white is used for content slides.
- **Fonts:** headings in Arial bold, body in Calibri (both safe fonts).
- **Type sizes:** titles 36pt, section headers 20–22pt, body 14–16pt, captions 11pt.
- **Motif:** **rounded photo frames** and **icons inside bright-blue circles**. Big stat numbers in bright blue. No accent bars or stripes, and no lines under titles.
- **Layouts (masters):** `TITLE_DARK`, `SECTION_DARK`, `CONTENT_LIGHT` (title placeholder top-left at the same position on every slide, a small footer "East Port Said National University", slide number), `CLOSING_DARK`.

## Slide plan
| # | id | Section | Title | Visual approach | Images |
|---|---|---|---|---|---|
| 1 | title | — | East Port Said National University | Half-bleed campus photo, logo, tagline "We Live the Future Today", subtitle "Building International Bridges from the Suez Canal" | campus hero, logo |
| 2 | glance | Who We Are | EPNU at a Glance | 4 big stat callouts (7 faculties · 13+ programs · 5,200+ students · English-medium) plus an identity strip (non-profit national university, Sinai, Egypt Vision 2030) | campus |
| 3 | why | Who We Are | Why Partner with EPNU | 2×2 icon cards: strategic location · international academic standards · future-focused programs · strong national backing | — |
| 4 | sec-gov | divider | Leadership & Governance | Dark divider | — |
| 5 | president | Leadership | Leadership with a Global Outlook | President profile: photo left; credentials as icon rows (PhD Leningrad, Egyptian-Japanese Initiative, Senate committees, NAQAAE reviewer) | president photo |
| 6 | board | Leadership | A Board Bridging Government, Academia & Industry | Chair highlighted card plus a grid of members grouped Government / Academia / Industry & Trade | — |
| 7 | sec-acad | divider | Academics | Dark divider | — |
| 8 | faculties | Academics | Seven Faculties, Future-Ready Programs | 7 faculty cards with icons and program names | — |
| 9 | standards | Academics | Built on International Academic Standards | Icon grid: English-medium · credit hours ≈ 30 ECTS/semester · bilingual certificates · non-degree and auditor registration · IGCSE and foreign certificates · Huawei ICT Academy | lab photo |
| 10 | sec-ir | divider | International Relations | Dark divider | — |
| 11 | mandate | IR | Internationalization Is a National Priority | 3-step timeline Mar → Jun → Jul 2026, plus a "Global Openness" pillar quote | — |
| 12 | upm | IR | First International Partner: Universiti Putra Malaysia | Process flow: signed agreement → international webinar → joint virtual research meeting → 2nd edition announced → next: faculty and student exchange; plus a "praised by Malaysian press" callout | webinar photo if available |
| 13 | network | IR | A Growing Global Network | Hub-and-spoke diagram: EPNU in the center, spokes to Malaysia, Japan, EU, Russia–Arab, UK, Italy, MENA, China (Huawei) | — |
| 14 | industry | IR | Partners with Regional & Global Reach | Cards: Huawei ICT Academy · Saudi German Hospital · pharmaceutical partners (EIPICO, EVA Pharma…) · 57357 | training photos |
| 15 | students | IR | Our Students on the International Stage | Photo collage plus achievement list (ICIAAI hackathon, ITC-Egypt 2nd stage, International Pediatric Conference, ECPC, 7 papers and 2nd place) | student photos |
| 16 | national | Partnerships | A Strong National Partnership Base | Stat "23 research papers with Port Said University" plus a partner grid (PSU, ITI/Creativa, Suez Canal Authority, Arish University, Armed Forces College of Medicine, Healthcare Authority, Port Said Governorate) | — |
| 17 | opportunities | Future | Partnership Opportunities | 6 icon cards: student and staff exchange · joint and dual degrees · joint research and Horizon Europe consortia · visiting professors · international student pathways · AI, health and maritime hubs | — |
| 18 | closing | — | Let's Build the Future Together | Dark slide, campus photo, website and enquiry route | campus |

## Process
1. The Haiku harvest is running. Next, Sonnet curates the images and Sonnet writes speaker notes in parallel.
2. Opus builds `tools/build_deck.js` and applies the theme.
3. Validate, render via PowerPoint COM to PNG, then Sonnet visual QA and Haiku text QA in parallel. Fix and re-render.
