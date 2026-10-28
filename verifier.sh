#!/usr/bin/env bash
# Contrôles à passer avant chaque commit et push, pour chaque userscript du dépôt.
set -e
cd "$(dirname "$0")"

for f in *.user.js; do
  node --check "$f"
  version=$(grep -m1 '^// @version' "$f" | awk '{print $3}')
  echo "$f : syntaxe OK, version $version"

  # La version n'a à bouger que si le fichier a changé depuis le dernier commit.
  if git rev-parse --git-dir > /dev/null 2>&1 && ! git diff --quiet HEAD -- "$f" 2>/dev/null; then
    precedente=$(git show "HEAD:$f" 2>/dev/null | grep -m1 '^// @version' | awk '{print $3}' || echo "")
    if [ -n "$precedente" ] && [ "$precedente" = "$version" ]; then
      echo "ATTENTION : $f modifié mais @version inchangée ($precedente)."
      echo "Violentmonkey ne proposera pas la mise à jour. Incrémenter avant de pousser."
      exit 1
    fi
  fi
done
echo "prêt à pousser"
