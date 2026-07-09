#!/usr/bin/env python3
"""resolve.py — deterministic, fill-first intent -> deck.py compose() layout resolver.

The Guided arm of the layout experiment: no per-slide layout agent, no LLM, no
randomness. `resolve(slide_spec, ctx)` maps one content-spec slide to a list of
"compose calls" (dicts the driver feeds to `Deck.compose()`, or a `{"native":...}`
call for the one archetype — cover — that has no compose() equivalent).

Design principle this file is built around (read `check_deadspace` in lint_brand.py
before changing geometry): a slide's dead-space score is NOT "is the box big
enough" — it is "how much of the canvas has actual ink or a filled shape on it".
Two consequences that shape every preset below:
  1. A `card`/`stat`/`figure` block draws a FILLED background rectangle that counts
     as covered for its whole declared box (up to 35% of the content region's area
     — bigger than that and the linter treats it as a background panel and excludes
     it entirely, see DEADSPACE_CARD_AREA_CAP). So sparse content (a section title,
     a bare statement) gets tiled across 2-4 stacked card/stat bands rather than
     dropped into one big, mostly-empty box.
  2. A bare `text`/`bullets`/`quote` block only counts the ESTIMATED INK height,
     anchored per its box's vertical anchor (top, for everything compose() draws).
     A big box with one short line only "fills" that one line. These block types
     are fine where there's enough real content to generate ink height on their
     own (evidence findings, enumerate item lists) — not as the sole content of a
     sparse slide.

`stack()`/`two_col()` below are the fill-tiling primitives every sparse-content
preset (section, statement, evidence intro) is built from; `safe_colspan()` keeps
every filled block's area under the linter's exclusion cap by construction, so
preset geometry doesn't need to be hand-computed per case.

One more wrinkle only visible by reading `_draw_block` closely: `card`/`callout`/
`chips` write their text into the SAME shape as their filled background (one
autoshape, `box.text_frame`), so the linter's "has text -> ink-only" branch fires
and the fill is never credited — only `stat` (and `figure`'s panel) draw the
background as a separate, permanently-empty-text shape layered under a real
textbox, so ONLY `stat`/`figure` get guaranteed full-box credit regardless of how
short their text is. That is why every sparse-fill preset below reaches for
`stat`, not `card`, as its primary tiling block; `card`/`bullets` are used only
where the content itself is long enough to generate real ink (findings, method
points, challenge sentences).
"""
import math
import os

# ---- grid geometry -------------------------------------------------------
# Mirrors deck.py's compose() constants (ax0/ax1/ay1/gut/ncols) and lint_brand.py's
# dead-space region/cap constants. Duplicated, not imported — deck.py and the
# linter are frozen for this task, and this is public, stable grid math ("12-col
# grid, 0.5-12.83in, 0.22in gutter") not an implementation detail.
AX0, AX1, AY1 = 0.5, 12.83, 6.4
GUT, NCOLS = 0.22, 12
COLW = (AX1 - AX0 - (NCOLS - 1) * GUT) / NCOLS          # ~0.826in

CANVAS_W = 13.333
DEADSPACE_TOP, DEADSPACE_BOTTOM = 0.5, 6.9
REGION_AREA = CANVAS_W * (DEADSPACE_BOTTOM - DEADSPACE_TOP)
CARD_CAP = 0.35 * REGION_AREA
SAFE_CAP = CARD_CAP * 0.85                                # margin below the linter's exclusion cap

CHIP_LEN = 40           # <= this many chars renders fine as a pill; longer needs a line/card


def band_top(dark, has_title, has_logo=True):
    """compose()'s ay0 for this canvas: title present, else dark-canvas logo
    keep-out, else the plain top margin."""
    if has_title:
        return 1.62
    if dark and has_logo:
        return 0.9
    return 0.55


def _colw(colspan):
    return colspan * COLW + (colspan - 1) * GUT


def _rowh(nrows, ay0):
    return (AY1 - ay0 - (nrows - 1) * GUT) / nrows


def safe_colspan(rowspan, nrows, ay0, want=12):
    """Largest colspan <= want whose filled-background area (a block spanning
    `rowspan` of `nrows` uniform rows starting at `ay0`) stays under SAFE_CAP."""
    rh = _rowh(nrows, ay0)
    h = rowspan * rh + (rowspan - 1) * GUT
    for c in range(want, 0, -1):
        if _colw(c) * h <= SAFE_CAP:
            return c
    return 1


# ---- fill-tiling primitives ----------------------------------------------
# `card` is deliberately absent here: card/callout/chips draw their text into
# the SAME shape as their fill (see module docstring), so they never get the
# guaranteed full-box dead-space credit `stat` gets — every sparse-fill preset
# below builds on `stat`, not `card`.

def _trim(s, n):
    """Trim to <= n chars total (ellipsis included), at a word boundary when possible.
    A single word longer than n has no boundary, so it is cut mid-word as a last resort."""
    if len(s) <= n:
        return s
    body = s[:n - 1]                       # reserve 1 char for the ellipsis
    cut = body.rsplit(" ", 1)[0]
    return (cut or body) + "…"


def _stat(big, size=30, body=None):
    b = {"type": "stat", "big": big, "size": size}
    if body:
        b["body"] = body
    return b


def stack(items, dark, has_title, colspan=10):
    """items: block dicts (card/stat — filled-background types) with no col/row.
    Stacks them as N full-width rows tiling the compose grid top to bottom, each
    row sized so its filled background stays under the dead-space area cap."""
    n = max(len(items), 1)
    ay0 = band_top(dark, has_title)
    # stat blocks are text-credited by the linter (never background-excluded), so
    # safe_colspan's area cap doesn't apply here — honor the full requested width;
    # shrinking only strands side gaps that trip the dead-space fill check.
    cs = min(colspan, 12)
    c0 = (12 - cs) // 2
    out = []
    for i, b in enumerate(items):
        b = dict(b)
        b["col"], b["colspan"] = c0, cs
        b["row"], b["rowspan"] = i, 1
        out.append(b)
    return out


def _distribute_rows(n_items, total_rows):
    """n_items slots sharing total_rows row-units as evenly as possible (extra
    units go to the earliest items) — so a shorter side isn't left with unfilled
    trailing rows when the two sides of a two_col() have different item counts."""
    if n_items == 0:
        return []
    base, rem = divmod(total_rows, n_items)
    out, row = [], 0
    for i in range(n_items):
        span = max(base + (1 if i < rem else 0), 1)
        out.append((row, span))
        row += span
    return out


def two_col(left_items, right_items, dark, has_title):
    """Left/right full-height rails, each independently stacked (card/stat blocks
    only — see module docstring). Both sides share one row grid sized to the
    taller side; a shorter side's items are given extra rowspan (not left with
    unfilled trailing rows) via _distribute_rows."""
    ay0 = band_top(dark, has_title)
    n = max(len(left_items), len(right_items), 1)
    lcs = safe_colspan(1, n, ay0, want=6)
    rcs = safe_colspan(1, n, ay0, want=6)
    out = []
    for (row, span), b in zip(_distribute_rows(len(left_items), n), left_items):
        b = dict(b); b["col"], b["colspan"] = 0, lcs; b["row"], b["rowspan"] = row, span
        out.append(b)
    for (row, span), b in zip(_distribute_rows(len(right_items), n), right_items):
        b = dict(b); b["col"], b["colspan"] = 6, rcs; b["row"], b["rowspan"] = row, span
        out.append(b)
    return out


def card_grid(cells, dark, has_title, cols_per_row=None):
    """cells: block dicts (card/stat), no col/row. Tiles them edge-to-edge in a
    grid, `cols_per_row` columns wide (default: min(4, count)), wrapping to as
    many rows as needed — the enumerate/items and evidence/findings-grid preset."""
    n = len(cells)
    cols_per_row = cols_per_row or min(4, max(n, 1))
    nrows = max(1, math.ceil(n / cols_per_row))
    ay0 = band_top(dark, has_title)
    cs = safe_colspan(1, nrows, ay0, want=12 // cols_per_row)
    out = []
    for i, b in enumerate(cells):
        b = dict(b)
        col, row = (i % cols_per_row) * (12 // cols_per_row), i // cols_per_row
        b["col"], b["colspan"] = col, cs
        b["row"], b["rowspan"] = row, 1
        out.append(b)
    return out


# ---- cover ----------------------------------------------------------------

def resolve_cover(s, ctx):
    return [{"native": "cover", "kwargs": {
        "title": s["title"], "subtitle": s.get("subtitle"), "tagline": s.get("tagline")}}]


# ---- section ----------------------------------------------------------------

def resolve_section(s, ctx):
    # a section carrying contact info is the closing slide (light thank-you + card)
    if s.get("contact") is not None:
        return [{"native": "closing", "kwargs": {"title": s["title"], "contact": s["contact"]}}]
    title, kicker = s["title"], s.get("kicker")
    items = [_stat(kicker.upper() if kicker else "", size=16), _stat(title, size=30)]
    blocks = stack(items, dark=True, has_title=False, colspan=11)
    return [{"blocks": blocks, "dark": True}]


# ---- statement ----------------------------------------------------------------

def resolve_statement(s, ctx):
    text = s["text"]
    body, support = s.get("body"), s.get("support")
    attribution, source, eyebrow = s.get("attribution"), s.get("source"), s.get("eyebrow")

    if attribution:
        # quote variant: the quote (as a stat -- see module docstring, a literal
        # `quote` block is ink-only and left this row under-filled) + a
        # guaranteed-fill source line.
        items = [_stat(text, size=32, body=attribution), _stat(source or "", size=14)]
        blocks = stack(items, dark=True, has_title=False, colspan=11)
        return [{"blocks": blocks, "dark": True}]

    if support:
        # Light problem/answer archetype (approved "Sample prep is part of the
        # biology." slide): serif headline + framing body + a horizontal step-flow
        # of the short support chips + a navy callout for the long "answer" line.
        cards = [(x, None) for x in support if len(x) <= 48]
        answer = " ".join(x for x in support if len(x) > 48) or None
        return [{"native": "problem_flow", "kwargs": {
            "title": text, "eyebrow": eyebrow, "body": body or "",
            "cards": cards or [(text, None)], "answer": answer, "caption": source}}]

    # bare: nothing but the claim — one big statement stat + a repeat at a
    # secondary size (no second real field exists to fill the second band with;
    # see brief §gaps — this is the honest limit, not a hidden drop).
    items = [_stat(text, size=34), _stat("", size=14)]
    blocks = stack(items, dark=True, has_title=False, colspan=11)
    return [{"blocks": blocks, "dark": True}]


# ---- evidence ----------------------------------------------------------------

def _method_stats(method):
    """method: list of chip-length words (narrative/singulator) OR a few full
    verbatim-methods sentences (dense) -- content SIZE decides the shape, not
    which spec this is. Short items join into one guaranteed-fill stat row;
    long ones each get their own row so no single row has to swallow 400+
    chars of joined text (that overflow was the dense-corpus failure mode)."""
    if not method:
        return []
    if len(method) <= 6 and all(len(m) <= 60 for m in method):
        return [_stat(" · ".join(method), size=13)]
    return [_stat(_trim(m, 150), size=11) for m in method]


def _citation_intro_items(citation, method, challenge, claim):
    """Left rail: citation + method, as stat blocks (guaranteed-fill background,
    see module docstring — a card here would only count its own short ink).
    Right rail: challenge points, or — when there's no explicit challenge — the
    claim, also as stat blocks so a 1-2-sentence challenge point still fills its
    full row regardless of how little text it has."""
    author = citation.get("author") or citation.get("doi") or ""
    if len(author) > 70:
        author = author.split(",")[0].strip() + " et al."
    left = [_stat(citation["ref"][:90], size=15, body=author)]
    left += _method_stats(method)
    points = (challenge or {}).get("points") or []
    if points:
        right = [_stat(p, size=14) for p in points]
    elif claim:
        right = [_stat(claim, size=15)]
    else:
        right = [_stat("", size=14)]
    return left, right


FINDINGS_CAP = 6   # findings-only cards per slide before a mechanical split


def _findings_items(figure, steps, stat, claim, findings, images_dir):
    """Right/only rail for the findings half of an evidence slide: figure or
    vertical steps on the left if present, stat/claim/findings on the right —
    falls back to a findings-only card grid with no figure/stat/steps. All
    stat blocks (guaranteed fill, see module docstring). `quote` doesn't ride
    here — resolve_evidence folds it into the `why` footnote instead, so it
    doesn't compete with findings for row height (that competition was
    overflowing findings rows on the singulator corpus). Returns a LIST of
    block-lists — one per slide, split mechanically when a long findings list
    (dense corpus: 7-8 items) would otherwise overflow a single grid."""
    right, weights = [], []
    if stat is not None and claim:
        # one block, not two -- stat's big+body pair already carries both, and
        # a separate claim row was pushing 5-item findings slides past what a
        # small row height could hold without overflowing (singulator corpus).
        # It gets weight 2 (a findings-length row's worth more): the big value
        # line plus a multi-line claim needs more than one findings-sized row.
        right.append(_stat(stat["value"], size=28, body=_trim(claim, 170))); weights.append(2)
    elif stat is not None:
        right.append(_stat(stat["value"], size=40, body=stat.get("label"))); weights.append(1)
    elif claim:
        right.append(_stat(claim, size=15)); weights.append(1)
    if findings:
        right += [_stat(_trim(f, 145), size=11) for f in findings]
        weights += [1] * len(findings)
    if not right:
        right, weights = [_stat("", size=13)], [1]

    if figure is not None:
        n = sum(weights)
        ay0 = band_top(True, True)
        lcs = safe_colspan(1, n, ay0, want=4)
        blocks = [{"type": "figure", "image": figure["image"], "caption": figure.get("caption"),
                   "col": 0, "colspan": lcs, "row": 0, "rowspan": n}]
        rcs = safe_colspan(1, n, ay0, want=(12 - lcs))
        row = 0
        for b, w in zip(right, weights):
            b = dict(b); b["col"], b["colspan"] = lcs, rcs; b["row"], b["rowspan"] = row, w
            blocks.append(b); row += w
        return [blocks]

    if steps:
        left = [_stat((f"{i+1}. " + (st.get('label') or '')), size=12, body=st.get("desc"))
                for i, st in enumerate(steps)]
        # right's weights (see above -- the merged stat+claim row needs 2x a
        # findings row) drive the shared row count; left (steps, roughly even
        # length) distributes across the same n rows via _distribute_rows.
        n = max(len(left), sum(weights))
        ay0 = band_top(True, True)
        # steps labels are short (don't need much width); claim/findings on the
        # right do -- give the right rail more of the 12 columns than an even
        # 6/6 split (that was overflowing findings text at 12pt/6-col width).
        lcs = safe_colspan(1, n, ay0, want=5)
        rcs = safe_colspan(1, n, ay0, want=7)
        blocks = []
        for (row, span), b in zip(_distribute_rows(len(left), n), left):
            b = dict(b); b["col"], b["colspan"] = 0, lcs; b["row"], b["rowspan"] = row, span
            blocks.append(b)
        row = 0
        for b, w in zip(right, weights):
            b = dict(b); b["col"], b["colspan"] = 5, rcs; b["row"], b["rowspan"] = row, w
            blocks.append(b); row += w
        return [blocks]

    # single hero (a lone stat+claim, no findings/figure/steps — e.g. the "110x"
    # slide): a 1-cell card_grid gets width-shrunk by safe_colspan and strands the
    # right half. stat is text-credited, never background-excluded, so span it full
    # width as two big stacked rows (value huge, claim below) to fill by ink.
    if len(right) == 1 and right[0].get("type") == "stat":
        b = right[0]
        hero = [_stat(b["big"], size=54)]
        if b.get("body"):
            hero.append(_stat(b["body"], size=18))
        return [stack(hero, dark=True, has_title=True, colspan=11)]

    # findings-only: no figure/steps to pair with — tile as a card grid so a
    # long findings list wraps into rows instead of one narrow stacked column;
    # cap items per slide so long findings text doesn't overflow a narrow cell.
    # 2 cells per row, not the card_grid default of up to 4 -- findings/claim
    # sentences here run 100-200 chars and need real width to avoid overflowing
    # a narrow 4-per-row cell (dense corpus: 8-finding evidence entries).
    chunks = [right[i:i + FINDINGS_CAP] for i in range(0, len(right), FINDINGS_CAP)] or [right]
    return [card_grid(chunk, dark=True, has_title=True, cols_per_row=min(2, len(chunk))) for chunk in chunks]


def _merge_why(why, quote):
    """quote rides in the compose() `why` footnote band (free real estate below
    the grid, see module docstring) instead of competing with findings/claim for
    row height on the findings half of the slide."""
    parts = [p for p in (why, quote.get("text") if quote else None) if p]
    return "  ".join(parts) or None


def resolve_evidence(s, ctx):
    """Light-first, matching the approved case-study decks: an evidence slide is
    EITHER a study *intro* (citation card + METHOD chips + challenge bullets) OR a
    *findings* slide (figure/steps left; big stat + KEY FINDINGS in a navy accent
    card, on white). Renders to the native `study_intro` / `study_findings` (light)
    archetypes rather than the old navy stat-tile compose grid — the tiles passed
    the linter but read as a monotone dark deck, nothing like the approved look."""
    title, eyebrow = s["title"], s.get("eyebrow")
    citation, claim = s.get("citation"), s.get("claim")
    method, challenge = s.get("method"), s.get("challenge")
    figure, stat, findings = s.get("figure"), s.get("stat"), s.get("findings")
    steps, quote, why = s.get("steps"), s.get("quote"), s.get("why")

    has_findings = any(x is not None for x in (figure, stat, findings, steps, quote))
    calls = []

    if citation is not None:
        author = citation.get("author") or citation.get("doi") or ""
        if len(author) > 70:
            author = author.split(",")[0].strip() + " et al."
        cpts = list((challenge or {}).get("points") or [])
        if not cpts and claim and not has_findings:
            cpts = [claim]
        calls.append({"native": "study_intro", "kwargs": {
            "title": title, "eyebrow": eyebrow,
            "citation": (citation["ref"], author),
            "method": method, "challenge_title": None, "challenge_points": cpts}})

    if has_findings:
        fig = None
        cap = None
        if figure is not None:
            fig = figure["image"]
            cap = figure.get("caption")
        stat_val = stat["value"] if stat else None
        stat_lab = (stat.get("label") if stat else None) or (claim if stat else None)
        if stat_lab:
            stat_lab = _trim(stat_lab, 150)
        step_pairs = [(st.get("label") or "", st.get("desc") or "") for st in (steps or [])] or None
        qtxt = quote["text"] if quote else None
        calls.append({"native": "study_findings", "kwargs": {
            "title": title, "eyebrow": eyebrow, "light": True,
            "figure": fig, "figure_caption": cap, "steps": step_pairs,
            "stat": stat_val, "stat_label": stat_lab,
            "statement": (claim if (stat_val is None and not findings) else None),
            "findings": findings, "quote": qtxt, "why": why}})

    if not calls:
        calls.append({"native": "study_findings", "kwargs": {
            "title": title, "eyebrow": eyebrow, "light": True,
            "statement": claim or s.get("body") or "", "why": why}})
    return calls


# ---- enumerate ----------------------------------------------------------------

TABLE_ROW_CAP = 10     # data rows per slide before a row-split
TABLE_COL_CAP = 4      # columns per group before a column-group split


def _table_col_groups(columns, rows):
    n = len(columns)
    if n <= TABLE_COL_CAP:
        return [(columns, rows)]
    n_groups = math.ceil(n / TABLE_COL_CAP)
    size = math.ceil(n / n_groups)
    groups = []
    for i in range(0, n, size):
        gc = columns[i:i + size]
        groups.append((gc, [r[i:i + size] for r in rows]))
    return groups


def _resolve_table(s, ctx):
    title, eyebrow = s["title"], s.get("eyebrow")
    t = s["table"]
    columns, rows = t["columns"], t["rows"]
    intro, quote = s.get("intro"), s.get("quote")

    n_row_groups = math.ceil(len(rows) / TABLE_ROW_CAP) if rows else 1
    row_size = max(1, math.ceil(len(rows) / n_row_groups))   # never 0: empty rows -> header-only table
    row_chunks = [rows[i:i + row_size] for i in range(0, len(rows), row_size)] or [rows]

    calls = []
    for part, chunk in enumerate(row_chunks, start=1):
        col_groups = _table_col_groups(columns, chunk)
        n = len(col_groups)
        cs = 12 // n
        blocks = []
        for i, (gc, gr) in enumerate(col_groups):
            blocks.append({"type": "table", "columns": gc, "rows": gr, "highlight_last": (i == n - 1),
                            "col": i * cs, "colspan": cs, "row": 0, "rowspan": 1})
        why = None
        parts = [p for p in (intro if part == 1 else None,
                              f"({part}/{n_row_groups})" if n_row_groups > 1 else None,
                              quote["text"] if quote and part == n_row_groups else None) if p]
        why = "  ".join(parts) or None
        page_title = title if n_row_groups == 1 else f"{title} ({part}/{n_row_groups})"
        calls.append({"blocks": blocks, "title": page_title, "dark": True, "eyebrow": eyebrow, "why": why})
    return calls


GROUP_ITEM_CAP = 8   # bullets-column items before a group needs splitting (avoids overflow)


def _split_groups(groups):
    """A bullets column of >GROUP_ITEM_CAP full-sentence items overflows its
    box (singulator's 10-item group at 5.32in text vs 4.78in available) --
    split any oversized group into evenly-sized, labeled continuation parts
    before pairing groups onto slides."""
    out = []
    for g in groups:
        items = g["items"]
        if len(items) <= GROUP_ITEM_CAP:
            out.append(g)
            continue
        n_parts = math.ceil(len(items) / GROUP_ITEM_CAP)
        size = math.ceil(len(items) / n_parts)
        for p, i in enumerate(range(0, len(items), size), start=1):
            label = f"{g['label']} ({p}/{n_parts})"
            out.append({"label": label, "items": items[i:i + size]})
    return out


def _bullets_col(label, items, col, colspan, rowspan=4):
    """One bullet column, full-height by default. (Previously reserved the bottom
    band for a "{N} studies" filler stat — removed: the count noun was hardwired to
    the study index and rendered wrong on other group grids, e.g. "6 studies /
    SAMPLE STATES". When the spec carries a `source` summary, the bottom band holds
    that real banner instead (rowspan 3); otherwise the column runs full height.)"""
    desc = [it.get("desc") or it.get("label") or "" for it in items]
    return [{"type": "bullets", "label": label.upper(), "items": desc,
             "col": col, "colspan": colspan, "row": 0, "rowspan": rowspan}]


GROUP_COLS = 3   # up to N bullet columns on ONE slide. Was hardwired to 2, which
                 # split a 3-column capability grid across two slides and synthesized
                 # a junk "{N} publications in {label}" stat on the lone trailing group.


def _resolve_groups(s, ctx):
    title, eyebrow, intro, source = s["title"], s.get("eyebrow"), s.get("intro"), s.get("source")
    groups = _split_groups(s["groups"])
    chunks = [groups[i:i + GROUP_COLS] for i in range(0, len(groups), GROUP_COLS)]
    calls = []
    # pack up to GROUP_COLS full-height columns per slide: a 2-section index and a
    # 3-column capability grid each land on ONE slide; a lone trailing group just
    # gets a wider column (no synthesized filler stat). A `source` summary renders
    # as a real bottom-banner callout on the last slide (the approved deck's banner).
    for ci, chunk in enumerate(chunks):
        last = ci == len(chunks) - 1
        banner = source if (last and source) else None
        span = 12 // len(chunk)
        rs = 3 if banner else 4
        blocks = []
        for j, g in enumerate(chunk):
            blocks += _bullets_col(g["label"], g["items"], j * span, span, rowspan=rs)
        if banner:
            blocks.append({"type": "callout", "text": banner,
                           "col": 0, "colspan": 12, "row": 3, "rowspan": 1})
        calls.append({"blocks": blocks, "title": title, "dark": False,
                      "eyebrow": eyebrow, "why": intro if ci == 0 else None})
    return calls


def _stat_hb(heading, body, size=15):
    """heading+body pair as a stat block (guaranteed-fill, see module docstring)
    — the enumerate/items card-grid equivalent of a labeled card."""
    if heading:
        return _stat(heading, size=size, body=body)
    return _stat(body or "", size=13)


def _resolve_items(s, ctx):
    title, eyebrow, arrangement = s["title"], s.get("eyebrow"), s.get("arrangement")
    items = s["items"]

    # Dueling big-numbers (approved "1 vs 8"): exactly two parallel items each a big
    # value with a caption/desc — two hero cards, not a stat grid.
    if arrangement == "parallel" and len(items) == 2 and all(it.get("value") for it in items):
        def _duel(it):
            return {"label": it.get("label") or "", "value": it["value"],
                    "caption": it.get("desc"), "note": it.get("note")}
        return [{"native": "dueling_numbers",
                 "kwargs": {"title": title, "eyebrow": eyebrow,
                            "left": _duel(items[0]), "right": _duel(items[1]),
                            "footnote": s.get("source")}}]

    if all("value" in it for it in items):
        cells = [_stat(it["value"], size=36, body=it.get("label")) for it in items]
        blocks = card_grid(cells, dark=False, has_title=True)
        return [{"blocks": blocks, "title": title, "dark": False, "eyebrow": eyebrow}]

    # Checklist (approved "What single-cell labs need"): a requirement list — items
    # carry only desc/label text (no value/points), 3+ of them, not a numbered
    # sequence. Green-check rows read as requirements where a card grid reads as tiles.
    if (len(items) >= 3 and arrangement not in ("sequential", "indexed")
            and all(not it.get("value") and not it.get("points") for it in items)
            and all(not (it.get("label") and it.get("desc")) for it in items)):
        labels = [(it.get("desc") or it.get("label") or "") for it in items]
        return [{"native": "checklist",
                 "kwargs": {"title": title, "eyebrow": eyebrow, "intro": s.get("intro"),
                            "items": labels}}]

    # Comparison option cards (approved "How labs handle X today"): >=2 items that
    # each carry a strengths list (`points`) — rendered as side-by-side cards with a
    # colored header band, STRENGTHS list, and a ⚠ caution footer (`warn`).
    if len(items) >= 2 and all(it.get("points") for it in items):
        columns = [{"label": it.get("label") or "", "desc": it.get("desc"),
                    "strengths": it.get("points") or [], "warn": it.get("warn")}
                   for it in items]
        return [{"native": "comparison_columns",
                 "kwargs": {"title": title, "eyebrow": eyebrow, "columns": columns}}]

    if arrangement == "parallel" and len(items) >= 2:
        left, right = items[0], items[1]
        blocks = two_col(
            [_stat_hb(left.get("label"), left.get("desc"))],
            [_stat_hb(right.get("label"), right.get("desc"))],
            dark=False, has_title=True)
        return [{"blocks": blocks, "title": title, "dark": False, "eyebrow": eyebrow,
                  "why": s.get("source")}]

    # sequential/indexed/grid: stat grid, numbered when order matters
    numbered = arrangement in ("sequential", "indexed")
    # label is optional in the schema (a labeledItem may carry only desc/value) — fall
    # back so a numbered item without a label doesn't KeyError (Codex finding, 2026-07-03).
    def _lbl(it):
        return it.get("label") or it.get("desc") or it.get("value") or ""
    cells = [_stat_hb((f"{i+1}. {_lbl(it)}" if numbered else _lbl(it)), it.get("desc"))
             for i, it in enumerate(items)]
    blocks = card_grid(cells, dark=False, has_title=True)
    return [{"blocks": blocks, "title": title, "dark": False, "eyebrow": eyebrow}]


def resolve_enumerate(s, ctx):
    if s.get("table") is not None:
        return _resolve_table(s, ctx)
    if s.get("groups") is not None:
        return _resolve_groups(s, ctx)
    if s.get("items") is not None:
        return _resolve_items(s, ctx)
    raise ValueError("enumerate slide has none of items/groups/table")


# ---- freeform ----------------------------------------------------------------

def resolve_freeform(s, ctx):
    return [{"blocks": s["blocks"], "title": s.get("title"), "dark": s.get("dark", False),
             "eyebrow": s.get("eyebrow"), "why": s.get("why")}]


# ---- dispatcher ----------------------------------------------------------------

RESOLVERS = {
    "cover": resolve_cover,
    "section": resolve_section,
    "statement": resolve_statement,
    "evidence": resolve_evidence,
    "enumerate": resolve_enumerate,
    "freeform": resolve_freeform,
}


def resolve(slide_spec, ctx):
    """slide_spec: one entry of spec["slides"]. ctx: {"images_dir": ...}.
    Returns a list of compose calls: {"blocks","title","dark","eyebrow","why"}
    dicts (only the keys the driver should pass to Deck.compose()), or a single
    {"native": method_name, "kwargs": {...}} for cover."""
    intent = slide_spec["intent"]
    fn = RESOLVERS.get(intent)
    if fn is None:
        raise ValueError(f"unknown intent: {intent!r}")
    return fn(slide_spec, ctx)
