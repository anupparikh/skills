---
name: deck-from-content
description: >
  Build an on-brand PowerPoint deck from content only. You (the author) write a short
  Python script that calls deck.py: for each slide you PICK ONE of the available layouts
  and fill it with your text, bullets, and image paths. The engine renders every slide on
  the client's native branded template — background, logo, theme colours, fonts, and
  auto-fitted titles all come from the template, so you never touch a colour, font, size,
  or coordinate and content cannot go off-brand. Trigger when turning written content /
  a research summary / an evidence set into a branded deck, "make slides from this", or
  rebuilding a deck on the brand template. NOT for building the template itself (that's
  brand-deck-template) or reading/extracting an existing deck.
---

# deck-from-content

Content in, on-brand deck out. You write Python that calls `deck.py`. **You choose the
layout for each slide; the engine owns the look** — palette, type, borders, spacing, logo,
footer, and title sizing all come from the template's theme and native layouts. You never
specify a colour, font, size, or coordinate.

The division of labour is deliberate:

- **You decide** what each slide says and **which layout it uses**. There is no
  recommended layout per kind of content and no auto-picker — the layout is your call.
- **The engine enforces** brand: it renders through native template placeholders and
  on-brand drawn components, auto-fits titles, and stamps the logo and footer. The one
  rule you must follow is **use the native components** (the methods and block types
  below) rather than hand-drawing your own text boxes, bullets, or shapes — that is what
  keeps a slide on-brand.

```python
from deck import Deck
d = Deck(footer="© 2026 Precision Cell Systems · For research use only")   # defaults to the PCS template
d.cover("Why researchers reach for the Singulator.", "Nine peer-reviewed studies.")
d.bullets("What labs need", ["Reproducible prep", "Gentle isolation", "Automated QC"])
d.save("out.pptx")
```

`Deck(brand_config=..., template=..., footer=..., manifest=...)` — pass nothing for PCS;
pass `template=` for another client's `brand-deck-template` output, or `brand_config=` to
generate one on the fly.

## Before you build (do this first)

1. **Find the brand kit.** Look for `clients/<brand>/brand/` — it holds the branded
   `*-template.pptx`, `deck-manifest.json`, and `brand_config.json`. Use that template;
   never invent a palette or pick a generic one.
2. **Read `brand_config.json` before writing a word of copy.** It carries the canonical
   product names and a `voice_and_tone` section with hard-bans (e.g. no eyebrow/kicker
   above headlines, no em-dashes in headlines, SI unit style). Getting these right on
   turn one is far cheaper than a correction pass.
3. **Install the brand fonts** (`fonts/install_fonts.sh`) — see the QA caveat at the
   bottom. Previews are untrustworthy without them.
4. **Pull figures from the client's own source decks** where possible.

## The layout menu — pick ONE per slide, then fill it

Each method below is a layout. Choose whichever fits the content you want on that slide,
call it, and put whatever you want inside it. The engine fills native placeholders where a
layout has them and draws on-brand components where it doesn't. Titles **auto-fit**: each
renders as large as it can and shrinks only enough to stay on one line, so you never set a
title size.

| Layout method | What the slide looks like |
|---|---|
| `cover(title, subtitle=None, tagline=None)` | dark title cover |
| `section(title, kicker=None)` | dark section divider |
| `statement(title)` | one big line on white |
| `bullets(title, points)` | white title + native bullets |
| `two_column(title, left, right)` | two text columns |
| `comparison(title, a_label, a_pts, b_label, b_pts)` | two labelled columns of points |
| `figure(title, image, caption=None)` | one image + caption |
| `content_with_figure(title, image, points)` | bullets beside an image |
| `agenda(title, groups)` | grouped numbered index; `groups=[(label, [(item, gloss), …]), …]` |
| `stat_cards(title, cards)` | row of metric tiles; `cards=[(big, label), …]` |
| `steps(title, steps)` | numbered step row; `steps=[(label, desc), …]` |
| `problem_flow(title, body, cards, answer=None, caption=None, eyebrow=None)` | framing + step-flow + answer bar; `cards=[(head, desc), …]` |
| `study_intro(title, citation, method=None, challenge_title=None, challenge_points=None, eyebrow=None)` | citation card + method chips + challenge bullets; `citation=(paper_title, meta)` |
| `study_findings(title, *, figure=/steps=, stat=/statement=, findings=, quote=, why=…)` | dark findings: figure or steps left, big stat/statement + KEY FINDINGS right |
| `compose(title, ...)` | freeform 12-column grid — for any slide the named layouts don't cover (see below) |

Want simple bullets? Use `bullets(...)` — it renders real native bullets, not boxes. Want
a number to headline? `stat_cards(...)`. Want a picture? `figure(...)`. The choice is
yours; there is no "right" layout the tool will nudge you toward.

## `compose()` — the freeform layout

When no named layout fits, `compose()` gives you a 12-column grid. You place on-brand
blocks; the engine still owns margins, gutters, borders, logo, footer, and colours, so any
arrangement stays on-brand.

```python
d.compose(title="One prep, every downstream assay.", dark=False, images_dir="assets", blocks=[
    {"type":"text",   "col":0, "colspan":7, "row":0, "rowspan":3, "title":"…", "body":"…"},
    {"type":"bullets","col":7, "colspan":5, "row":0, "rowspan":2, "label":"…", "items":["…","…"]},
    {"type":"stat",   "col":7, "colspan":5, "row":2, "rowspan":1, "big":"< 30 min", "body":"…"},
    {"type":"callout","col":0, "colspan":12,"row":3, "rowspan":1, "text":"…"},
])
```

- `dark=True` → navy slide + reversed logo; `dark=False` → white. `title`, `eyebrow`,
  `why` optional.
- Content area = 12 columns wide × however many rows your blocks span. Each block gives
  `col` (0–11), `colspan`, `row` (0-based), `rowspan`; the engine computes every
  coordinate. Blocks don't overlap unless their cells do.
- **Block types** (colours adapt to light/dark automatically):
  `text` `{title?, body}` · `bullets` `{label?, items[]}` · `card` `{heading?, body, variant?:"tint"|"navy", accent?}` · `stat` `{big, body, size?}` · `chips` `{items[]}` · `quote` `{text, attrib?}` · `callout` `{text}` · `figure` `{image, caption?}` · `table` `{columns[], rows[[…]], highlight_last?}`.

Keep a block's text to what fits its cell — give it more rows or trim the copy if it's tight.

## Worked example

`scripts/example_singulator.py` builds a real deck by calling these layouts directly from
`scripts/example_content.json` + `scripts/assets/`. Read it as the reference for driving the
engine; run it to regenerate the proof:

```bash
python scripts/example_singulator.py out.pptx
```

## Verify (required)

**1. Brand linter — run on every build.** `scripts/lint_brand.py` reads the output `.pptx`
and flags off-palette colours, off-brand fonts, missing footer/page-number, off-canvas or
colliding shapes, and text overflow. These are hard failures — fix them and re-render until
it passes. **Dead-space (under-filled slides) is reported as an `advisory`, not a failure**:
an airy slide is a design choice you may keep, so the linter no longer pushes you to pad it
with boxes.

```bash
python3 scripts/lint_brand.py out.pptx --template <template.pptx> --config <brand_config.json> --json
python3 scripts/lint_brand.py --selfcheck   # proves each check still fires
```

`pass` in the JSON reflects only hard defects; `advisories` lists dead-space notes
separately. `--no-deadspace` skips the advisory pass entirely.

**2. Fresh eyes.** Render to images and look at every slide (`references/design-critic.md`
lists what to check — balance, hierarchy, readability). Linter-clean ≠ well-designed.

```bash
soffice --headless --convert-to pdf --outdir /tmp out.pptx
pdftoppm -jpeg -r 120 /tmp/out.pdf /tmp/slide
```

## Scope and knobs

This skill owns slide **composition** — mapping your content onto native branded layouts
and drawing on-brand components. It reads all brand values from `brand_config` / the
template theme, so content can't go off-brand. The template's look is owned by
`brand-deck-template`; regenerate it there if the brand changes and this engine re-skins
automatically.

Per-brand, non-paint knobs live in `deck-manifest.json` beside each brand's template (see
`references/brand-manifest.md`): dark-canvas layout, logo position + keep-out, footer text,
the neutral colours not in the theme, and `title_sizes` — which is now an **auto-fit cap**
per title kind (`content` / `dark` / `statement`), not a fixed size. Raise a cap to let
short titles grow larger; the engine still shrinks any title that would wrap. Pass the
manifest via `Deck(..., manifest=<path>)`; omit it and PCS-shaped defaults apply.

## Fonts must be installed (QA caveat)

The deck names the brand's fonts (Merriweather display, Franklin Gothic body). If they
aren't installed on the authoring machine, PowerPoint/LibreOffice substitute silently — the
deck drifts off-brand **and every rendered preview lies about text fit and overflow**, so
you can't trust visual QA for exactly the failure modes that matter. Install before
building or reviewing:

```bash
bash fonts/install_fonts.sh
```

Merriweather (OFL) ships in `fonts/` and installs directly. Franklin Gothic is proprietary
and is not bundled — install your licensed copy, or regenerate the template in
`brand-deck-template` to use the open substitute Libre Franklin (the engine reads fonts
from the template theme, so no code change here). See `fonts/README.md`.

The engine also guards fit mechanically: `bullets`/`two_column`/`comparison` size body text
to fit the placeholder, titles auto-fit to one line, and `steps` scales rows to the content
band. The linter's overflow/collision checks are the backstop.
