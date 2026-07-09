# skills

Private plugin marketplace for Claude Code / Claude Cowork.

## Install

```
/plugin marketplace add anupparikh/skills
/plugin install pcs-skills@anupparikh-skills
```

Requires GitHub access to this repo (private).

## Plugins

| Plugin | Skills | What it does |
|---|---|---|
| `pcs-skills` | `deck-from-content` | Content in, on-brand PowerPoint deck out. Renders slides on the PCS branded template; agent supplies text and image paths only. |

## Repo layout

```
.claude-plugin/marketplace.json      # the marketplace index (repo root)
plugins/<plugin>/
  .claude-plugin/plugin.json         # ONLY plugin.json lives here
  skills/<skill>/SKILL.md            # skills/, commands/, agents/, hooks/ at plugin root
```

Only `plugin.json` goes inside a plugin's `.claude-plugin/`; every component dir
(`skills/`, `commands/`, …) sits at the plugin root. A plugin must be self-contained — it
can't reference files outside its own directory (installs are copied to a cache).

## Adding a plugin

1. Create `plugins/<name>/.claude-plugin/plugin.json` (required: `name`; recommended:
   `displayName`, `version`, `description`, `author` as an object, `license`, `keywords`)
   and a `skills/<skill>/SKILL.md` (recommended frontmatter: `description`).
2. Add an entry to `.claude-plugin/marketplace.json` (`name` + `source`; add `version` to
   pin — omit it to auto-update on every commit).
3. **Validate before pushing:** `claude plugin validate . --strict` (checks both the
   marketplace index and each plugin manifest; must pass clean).
4. Commit and push. Bump `version` in both `plugin.json` and the marketplace entry for a
   release — pushing without a bump means installed users see no update. They pull with
   `/plugin marketplace update anupparikh-skills`.
