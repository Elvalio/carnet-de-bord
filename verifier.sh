#!/usr/bin/env bash
# Contrôles à passer avant chaque push.
set -e
cd "$(dirname "$0")"

node --check carnet-de-bord.user.js
echo "syntaxe : OK"

version=$(grep -m1 '^// @version' carnet-de-bord.user.js | awk '{print $3}')
echo "version déclarée : $version"

if git rev-parse --git-dir > /dev/null 2>&1; then
  precedente=$(git show HEAD:carnet-de-bord.user.js 2>/dev/null | grep -m1 '^// @version' | awk '{print $3}' || echo "")
  if [ -n "$precedente" ] && [ "$precedente" = "$version" ]; then
    echo "ATTENTION : @version inchangée depuis le dernier commit ($precedente)."
    echo "Violentmonkey ne proposera pas la mise à jour. Incrémenter avant de pousser."
    exit 1
  fi
fi
echo "prêt à pousser"
