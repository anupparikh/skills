# skills — private Claude plugin marketplace

Project type: development
Commit policy: commit-freely

Repo = plugin marketplace (`.claude-plugin/marketplace.json`) consumed by Claude Code / Cowork via `/plugin marketplace add anupparikh/skills`. Private repo — installers need GitHub access.

Layout: one plugin per `plugins/<name>/`. ONLY `plugin.json` in the plugin's `.claude-plugin/`; component dirs (`skills/`, `commands/`, `agents/`, `hooks/`) at plugin root. Plugin must be self-contained (installs are cached — no refs outside the plugin dir). New plugin → add folder + marketplace.json entry.

Manifest conventions (enforced by `claude plugin validate`):
- `author` MUST be an object (`{name,email}`), not a string.
- plugin.json: `name` (kebab, no spaces) + `displayName` + `version` (semver) + `description` + `author` + `license` + `keywords`.
- marketplace.json: top-level `name`, `description`, `owner`, `plugins[]` (each `name`+`source`+`version`+`description`+`displayName`).

Release flow (always): edit → `claude plugin validate . --strict` (must pass clean) → bump `version` in BOTH plugin.json and its marketplace entry (they must agree) → commit → `claude plugin tag plugins/<name>` → push `main` + the `refs/tags/<name>--v<ver>` tag. Pushing without a version bump = installed users see no update. Current: pcs-skills 0.3.0.

Do not run production deck-generation work here; this repo is for building/packaging skills only.
