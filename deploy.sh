#!/usr/bin/env bash
# deploy.sh — push Tenant Guard to GitHub and enable Pages.
# Usage:  ./deploy.sh [repo-name] [public|private]
#         ./deploy.sh tenant-guard
set -euo pipefail

REPO="${1:-tenant-guard}"
VIS="${2:-public}"   # GitHub Pages on the Free plan needs a PUBLIC repo

# --- prerequisites ----------------------------------------------------------
command -v git >/dev/null || { echo "✗ git not installed → https://git-scm.com"; exit 1; }
command -v gh  >/dev/null || { echo "✗ GitHub CLI not installed → https://cli.github.com"; exit 1; }
[ -f index.html ] || { echo "✗ index.html not found. Run this from the folder that holds it."; exit 1; }

# --- auth -------------------------------------------------------------------
gh auth status >/dev/null 2>&1 || gh auth login
OWNER="$(gh api user --jq .login)"

# --- warn if client id not set ---------------------------------------------
if grep -q "REPLACE_WITH_YOUR_CLIENT_ID" index.html; then
  echo "⚠  index.html still contains REPLACE_WITH_YOUR_CLIENT_ID."
  echo "   Set it now (sed) or later via the app's Settings panel — continuing anyway."
fi

# --- git init + commit ------------------------------------------------------
[ -d .git ] || git init -q
git add .
git commit -q -m "Tenant Guard — Entra security & compliance console" || echo "· nothing new to commit"
git branch -M main

# --- create remote (if new) + push -----------------------------------------
if gh repo view "$OWNER/$REPO" >/dev/null 2>&1; then
  git remote add origin "https://github.com/$OWNER/$REPO.git" 2>/dev/null || true
  git push -u origin main
else
  gh repo create "$OWNER/$REPO" --"$VIS" --source=. --remote=origin --push
fi

# --- enable GitHub Pages from main / root ----------------------------------
gh api -X POST "repos/$OWNER/$REPO/pages" -f "source[branch]=main" -f "source[path]=/" >/dev/null 2>&1 \
  && echo "· Pages enabled" \
  || echo "· Pages already on (or enable it in repo Settings → Pages → main / root)"

URL="https://$OWNER.github.io/$REPO/"
echo
echo "✅ Done. Live in ~1 minute at:"
echo "   $URL"
echo
echo "▶ Last step: in your app registration, confirm the SPA redirect URI is exactly:"
echo "   $URL"
