# Per-brand manifest — the non-paint, per-brand knobs (D-034)

The template + its baked theme own **paint** (palette, fonts, logo art). This manifest owns the
handful of per-brand decisions that were previously hardcoded in `deck.py` and blocked a clean brand
swap. New brand/format = new template + this manifest, never editing layout math in `deck.py`.

Lives beside each brand's template: `clients/<client>/brand/deck-manifest.json` (test fixtures put it
next to their template). The engine loads it via `Deck(..., manifest=<path>)`; absent → PCS-shaped
defaults so nothing breaks.

```json
{
  "dark_canvas_layout": "Blank",
  "logo": { "dark_position": "top-right", "dark_size_in": 0.38, "keepout_in": 1.2 },
  "footer_text": "© 2026 Precision Cell Systems · For research use only",
  "title_sizes": { "problem_flow": 30, "study_intro": 25, "dark": 22, "statement": 32, "compose": 26 },
  "colors": {
    "muted":       "#C8D8EC",
    "page_number": "#9AA5B1",
    "quote_on_light": "#4A5568",
    "quote_on_dark":  "#C8D8EC",
    "warn": "#FD7E14"
  }
}
```

Field notes:
- `dark_canvas_layout` — template layout name used as the full-bleed dark base (`_dark_base`).
- `logo.dark_position` — `top-right` | `top-left`; `keepout_in` = width of the reserved no-draw zone
  the grid must not place blocks into (fixes the dark-canvas title↔logo collision).
- `colors.muted` / `page_number` — the two hexes `deck.py` hardcoded (:173, :176) that leaked onto
  brand B. Source of truth for these is `brand_config.presentation.palette.on_navy_muted` /
  `.page_number`; the manifest carries a copy so the engine needs only template + manifest at render
  time (no brand_config dependency). If a brand omits them, fall back to the template theme's muted
  slot / a derived grey.
- `colors.quote_on_light` / `quote_on_dark` — the quote-text color pair (was one fixed light color,
  near-invisible on light slides). Light slides use the on-light (slate) value; dark use on-dark.
- `colors.warn` — the compose() card "warn" accent (was a hardcoded PCS orange `#FD7E14` at
  `deck.py`'s `_draw_block`, found by the same hex-literal grep that caught muted/page_number). Falls
  back to `#FD7E14` if a brand omits it; a brand with no dedicated warning color can point this at an
  existing on-palette accent instead of introducing a new hex.
- `title_sizes` — controlled per-archetype title point sizes (were inline literals). Keys:
  `problem_flow`, `study_intro`, `dark` (the `_dark_base` title), `statement`, `compose`.
