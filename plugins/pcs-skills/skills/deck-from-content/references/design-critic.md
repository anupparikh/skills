# Design critic — scoring rubric for rendered slides

The soft-quality half of the deck verifier (the hard half is `scripts/lint_brand.py`). A vision
agent with fresh eyes — not the agent that built the deck — scores each rendered slide image
against this rubric. Used two ways: as the test oracle when evaluating engine changes, and as the
self-correction signal for a designer agent at build time.

## Protocol

1. Render the deck to per-slide images (LibreOffice → PDF → `pdftoppm -jpeg -r 110`).
2. Give the critic agent the images and the **Scoring prompt** below, one deck per agent. Never
   reuse the agent that authored the deck; fresh context only.
3. Critic returns one JSON object per slide (schema below). No prose outside the JSON.

## Dimensions (score each 1–5)

- **balance** — Is visual weight distributed with intent? 5: the full frame participates; weight
  sits where the eye should go. 3: usable but lopsided — one quadrant crowded or dead. 1: content
  huddles in one region; large accidental voids.
- **whitespace** — Is empty space *shaped*, or leftover? 5: margins and gaps look chosen, breathing
  room around groups. 3: uneven gaps, one cramped seam. 1: elements collide visually or drown in
  arbitrary emptiness.
- **hierarchy** — Does one scan order emerge? 5: instant read: what this slide says, then the
  support. 3: title wins but support elements compete. 1: everything shouts equally.
- **readability** — Type size/contrast/line length at presentation distance. 5: everything legible
  from the back of a room. 3: secondary text marginal (long lines, small captions). 1: any body
  text illegible or clipped.

Score what is rendered, not what was intended. Do not reward decoration; a plain slide that reads
instantly beats a busy one.

## Named defects

Alongside scores, name concrete defects, worst first, using these tags (free-text `detail` each):
`dead-space` · `crowding` · `misalignment` · `orphan` (element visually unattached to any group) ·
`imbalance` · `weak-title` · `text-wall` · `tiny-text` · `low-contrast` · `awkward-crop` (image) ·
`inconsistent` (breaks a pattern the rest of the deck established). Empty list if none.

## Scoring prompt (give verbatim to the critic agent)

> You are a presentation design critic. You did not make these slides; judge only what you see.
> For each slide image, score 1–5 on balance, whitespace, hierarchy, readability per the rubric
> provided, and list named defects (tags from the rubric, worst first, with one-line details).
> Judge each slide alone first, then add deck-level notes for cross-slide inconsistencies.
> Return ONLY JSON: `{"slides": [{"n": 1, "balance": 4, "whitespace": 3, "hierarchy": 5,
> "readability": 5, "defects": [{"tag": "dead-space", "detail": "bottom third empty"}]}, …],
> "deck_notes": ["…"]}`. Be strict: 5 is rare; a slide with any named defect caps at 4 on the
> dimension the defect touches.

## Repeatability

Same deck scored twice by fresh agents should agree on slide *ranking* (which slides are weakest)
even if absolute scores drift ±1. If rankings disagree, the rubric — not the critic — needs
tightening; fix it here.
