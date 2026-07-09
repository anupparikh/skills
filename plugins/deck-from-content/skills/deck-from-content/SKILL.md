---
name: deck-from-content
description: >
  Build an on-brand PowerPoint deck from content only. A calling agent supplies text,
  bullets, and image paths through a small Python API (deck.py) and the engine renders
  each slide on the client's native branded template — background, logo, theme colours
  and fonts all come from the template, so the agent never touches a colour, font, or
  coordinate and content cannot go off-brand. Two ways to build a slide: ready-made
  semantic methods for the recurring page shapes, and a coordinate-free 12-column
  compose() grid for laying out any page freely. Trigger when turning written content /
  a research summary / an evidence set into a branded deck, "make slides from this", or
  rebuilding a deck on the brand template. NOT for building the template itself (that's
  brand-deck-template) or reading/extracting an existing deck.
---

# deck-from-content

Content in, on-brand deck out. You write Python that calls `deck.py`; the engine owns the
look (palette, type, borders, spacing, logo, footer — all read from the template's theme and
the native layouts). **You never specify a colour, font, size, or coordinate.**

```python
from deck import Deck
d = Deck(footer="© 2026 Precision Cell Systems · For research use only")   # defaults to the PCS template
d.cover("Why researchers reach for the Singulator.", "Nine peer-reviewed studies.")
d.bullets("What labs need", ["Reproducible prep", "Gentle isolation", "Automated QC"])
d.save("out.pptx")
```

`Deck(brand_config=..., template=..., footer=...)` — pass nothing for PCS; pass `template=`
for another client's `brand-deck-template` output, or `brand_config=` to generate one on the fly.

## Recommended path: content spec → resolver (D-034)

For turning a content set into a deck, emit a **brand-agnostic content spec** — a list of slides that
declare *communicative intent* (`cover · section · statement · evidence · enumerate` + `freeform`), not
layout — and let the resolver lay them out. Schema: `references/content-spec.schema.json`.

```bash
python3 scripts/map_guided.py content.spec.json out.pptx \
  --template <template.pptx> --manifest <deck-manifest.json> [--images-dir DIR]
```

`resolve.py` maps each intent to a **fill-first** layout (adapting to content size, splitting
over-long tables/lists across slides), gated by the brand linter's dead-space check so slides fill the
frame. A slide's intent + which fields it carries pick the layout — one intent renders several ways.
This is the seam that keeps *what to say* (the spec, per task) separate from *how it looks* (engine +
template + manifest). Drop to the raw grid only via a `freeform` spec slide, never around the spec.

The two lower-level APIs below (semantic methods, `compose()`) are what the resolver calls and remain
available for hand-driving a one-off deck.

## Two ways to build a slide (the lower-level engine API)

### 1. Semantic methods — one call per common page shape
Each maps to the reference deck's recurring archetypes and fills native placeholders where a
layout exists, drawing only what has no native equivalent:

| Method | Slide |
|---|---|
| `cover(title, subtitle=None, tagline=None)` | dark title cover |
| `section(title, kicker=None)` | dark section divider |
| `bullets(title, points)` | white title + bullets |
| `statement(title)` | white statement |
| `two_column(title, left, right)` / `comparison(title, a_label, a_pts, b_label, b_pts)` | two columns |
| `figure(title, image, caption=None)` / `content_with_figure(title, image, points)` | image slides |
| `agenda(title, groups)` | grouped numbered index; `groups=[(label, [(item, gloss), …]), …]` |
| `problem_flow(title, body, cards, answer=None, caption=None, eyebrow=None)` | framing + step-flow + answer bar; `cards=[(head, desc), …]` |
| `study_intro(title, citation, method=None, challenge_title=None, challenge_points=None, eyebrow=None)` | citation card + method chips + challenge bullets; `citation=(paper_title, meta)` |
| `study_findings(title, *, figure=/steps=, stat=/statement=, findings=, quote=, why=…)` | dark findings: figure or steps left, big stat/statement + KEY FINDINGS right |
| `stat_cards(title, cards)` / `steps(title, steps)` | metric tiles / numbered steps |

### 2. `compose()` — freeform 12-column grid (creative path)
Place on-brand blocks anywhere on a 12-column grid; the engine owns margins, gutters, borders,
logo and footer, so any arrangement stays on-brand and clean.

```python
d.compose(title="One prep, every downstream assay.", eyebrow="PLATFORM AT A GLANCE", dark=False,
          why="Why it matters: …", images_dir="assets", blocks=[
    {"type":"card",   "col":0, "colspan":4, "row":0, "rowspan":2, "heading":"COLD & AUTOMATED", "body":"…"},
    {"type":"bullets","col":4, "colspan":4, "row":0, "rowspan":2, "label":"WHAT YOU GET", "items":["…","…"]},
    {"type":"stat",   "col":8, "colspan":4, "row":0, "rowspan":2, "big":"< 30 min", "body":"…", "size":34},
    {"type":"callout","col":0, "colspan":12,"row":2, "rowspan":1, "text":"The readout reflects the sample, not the method."},
])
```

- `dark=True` → navy slide + reversed logo; `dark=False` → white slide. `title`, `eyebrow`, `why` optional.
- Content area = a grid **12 columns** wide × however many **rows** your blocks span. Each block gives `col` (0–11), `colspan`, `row` (0-based), `rowspan`; the engine computes every coordinate. Blocks don't overlap unless their cells do.
- **Block types** (colours adapt to light/dark automatically):
  `text` `{title?, body}` · `bullets` `{label?, items[]}` · `card` `{heading?, body, variant?:"tint"|"navy", accent?:"accent_bright"|"warn"|"primary"|"green"|"red"}` · `stat` `{big, body, size?}` · `chips` `{items[]}` · `quote` `{text, attrib?}` · `callout` `{text}` · `figure` `{image, caption?}` (filename resolved against `images_dir`) · `table` `{columns[], rows[[…]], highlight_last?}`.

Keep a block's text to what fits its cell — give it more rows or trim the copy if it's tight.

## Worked example / proof

`scripts/example_singulator.py` rebuilds the first 15 slides of the Singulator differentiation
deck from `scripts/example_content.json` + `scripts/assets/` — semantic methods for the recurring
pages, `compose()` for the metric table and the two-way comparison. Read it as the reference for
driving the engine; run it to regenerate the proof:

```bash
python scripts/example_singulator.py out.pptx
```

## Verify (required) — automated linter first, then fresh eyes

**1. Brand linter (mechanical, no LLM) — run on every build.** `scripts/lint_brand.py` reads the
output `.pptx` and flags off-palette colors, off-brand fonts, missing footer/page-number, off-canvas
or colliding shapes, text overflow, **and dead-space** (under-filled slides — the #1 design defect,
per the D-034 experiment). Boolean pass + defect list; exit 0 = clean.

```bash
python3 scripts/lint_brand.py out.pptx --template <template.pptx> --config <brand_config.json> --json
python3 scripts/lint_brand.py --selfcheck   # proves each check still fires
```

Fix every defect and re-render until it exits 0. A dead-space defect means a slide doesn't fill —
give its content more room / larger type / more blocks; overflow means the opposite.

**2. Design critic (soft qualities the linter can't see).** Render to images and score with a
fresh-eyes subagent against `references/design-critic.md` (balance, whitespace, hierarchy,
readability + named defects). Linter-clean ≠ well-designed — both signals are needed.

```bash
soffice --headless --convert-to pdf --outdir /tmp out.pptx   # macOS: /Applications/LibreOffice.app/Contents/MacOS/soffice
pdftoppm -jpeg -r 120 /tmp/out.pdf /tmp/slide
```

## Scope

Owns slide **composition** — mapping content onto native branded layouts and drawing on-brand
components. It reads all brand values from `brand_config` / the template theme, so content can't go
off-brand. The template's look is owned by `brand-deck-template`; regenerate it there if the brand
changes and this engine re-skins automatically.

## Per-brand knobs — the manifest

The few per-brand, non-paint decisions that used to be hardcoded live in a small
`deck-manifest.json` beside each brand's template (see `references/brand-manifest.md`): dark-canvas
layout, logo position + keep-out, footer text, controlled title sizes, and the two neutral colors
(muted, page-number) + quote-contrast pair that aren't in the theme. Pass it via
`Deck(..., manifest=<path>)`. A new brand = new template + manifest, **zero Python** (D-033/D-034).
Omit it and the engine uses PCS-shaped defaults, so existing callers are unchanged.

## Known gap / follow-up

The template ships no native **dark content** layout (only dark cover + section). Dark pages
(`study_findings`, dark `compose`) are composed on a navy canvas + reversed logo inside the engine.
To make dark content a genuine native layout, add dark-content layouts to `brand-deck-template`'s
`scripts/assets/skeleton.pptx`. (Footer uniformity across archetypes was fixed in D-034 — every
non-cover slide now stamps footer + page-number from one shared base.)

## Fonts must be installed

The deck names the brand's fonts (Merriweather, Franklin Gothic). If they aren't installed on the
authoring machine, PowerPoint substitutes silently and the deck drifts off-brand.
