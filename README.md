# skills

Private plugin marketplace for Claude Code / Claude Cowork.

## Install

```
/plugin marketplace add anupparikh/skills
/plugin install deck-from-content@anupparikh-skills
```

Requires GitHub access to this repo (private).

## Plugins

| Plugin | What it does |
|---|---|
| `deck-from-content` | Content in, on-brand PowerPoint deck out. Renders slides on the PCS branded template; agent supplies text and image paths only. |

## Adding a plugin

1. Create `plugins/<name>/` with `.claude-plugin/plugin.json` and a `skills/` folder.
2. Add an entry to `.claude-plugin/marketplace.json`.
3. Commit and push. Installed marketplaces pick up changes on `/plugin marketplace update`.
