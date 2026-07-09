# skills — private Claude plugin marketplace

Project type: development
Commit policy: commit-freely

Repo = plugin marketplace (`.claude-plugin/marketplace.json`) consumed by Claude Code / Cowork via `/plugin marketplace add anupparikh/skills`. Private repo — installers need GitHub access.

Layout: one plugin per `plugins/<name>/`, each with `.claude-plugin/plugin.json` + `skills/<skill-name>/SKILL.md`. New plugin → add folder + marketplace.json entry.

Do not run production deck-generation work here; this repo is for building/packaging skills only.
