#!/usr/bin/env bash
# Install the bundled brand fonts so PowerPoint/LibreOffice render the deck true to
# metric (no silent font substitution). Copies every .ttf/.otf beside this script into
# the user font dir. Idempotent. macOS + Linux.
set -euo pipefail
here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

case "$(uname -s)" in
  Darwin) dest="$HOME/Library/Fonts" ;;
  *)      dest="$HOME/.local/share/fonts" ;;
esac
mkdir -p "$dest"

n=0
for f in "$here"/*.ttf "$here"/*.otf; do
  [ -e "$f" ] || continue
  cp -f "$f" "$dest/"; n=$((n+1))
done
echo "installed $n font file(s) to $dest"
command -v fc-cache >/dev/null 2>&1 && fc-cache -f "$dest" >/dev/null 2>&1 || true

echo "NOTE: Merriweather (display) is installed. The body font 'Franklin Gothic Book'"
echo "is NOT bundled (proprietary). Install your licensed copy, or regenerate the"
echo "template to use the open substitute 'Libre Franklin'. See fonts/README.md."
