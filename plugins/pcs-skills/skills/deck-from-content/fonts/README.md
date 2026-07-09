# Brand fonts

The deck names two fonts. If they aren't installed, PowerPoint/LibreOffice substitute
silently and the deck drifts off-brand — and, worse, text-fit/overflow QA on rendered
previews becomes untrustworthy (the substitute has different metrics). Install before
building or reviewing a deck.

```bash
bash install_fonts.sh        # copies bundled fonts into your user font dir
```

## What ships here, and what doesn't

| Role | Font | Bundled? | Why |
|---|---|---|---|
| Display (headlines) | **Merriweather** | ✅ yes | SIL Open Font License — free to redistribute. |
| Body (text) | **Franklin Gothic Book** | ❌ no | Monotype proprietary — cannot be redistributed. |

**Franklin Gothic** — install your own licensed copy (it ships with Microsoft Office /
Windows), OR switch the brand to the open, metrically-similar **Libre Franklin** (also
SIL OFL): regenerate the template in `brand-deck-template` with the body font set to
`Libre Franklin`, and this engine re-skins to it automatically (fonts are read from the
template theme — no code change here).

## Licensing

- `Merriweather[opsz,wdth,wght].ttf` — © The Merriweather Project Authors, SIL Open
  Font License 1.1 (full text in `Merriweather-OFL.txt`). This is the unmodified
  upstream variable font from Google Fonts (`ofl/merriweather`); its typographic family
  name is `Merriweather`, which is what the template theme references (weight comes from
  the font's `wght` axis). "Merriweather" is a Reserved Font Name — the OFL permits
  redistributing this original under that name, but a *modified* copy must be renamed.
- Do not add any font to this folder unless its license permits redistribution (OFL,
  Apache, or an explicit grant). Proprietary fonts (Franklin Gothic, most foundry fonts)
  must not be committed.
