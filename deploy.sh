#!/bin/bash
# BI Publisher Mac — Deploy to GitHub Pages (gh-pages branch)
# Usage: ./deploy.sh
set -e

BASE_URL="https://capitansuat.github.io/bi-publisher-mac"
REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
WORD_ADDIN_DIR=~/Library/Containers/com.microsoft.Word/Data/Library/Application\ Support/Microsoft/Office/16.0/User\ Data/Web\ Add-ins/bi-publisher-mac

echo "🔨 Building production bundle..."
cd "$REPO_DIR"
npm run build -- --env BASE_URL="$BASE_URL"

echo "📋 Generating production manifest..."
sed "s|https://localhost:3000|$BASE_URL|g" manifest.xml > dist/manifest-production.xml
cp dist/manifest-production.xml manifest-production.xml

echo "🚀 Pushing to gh-pages branch..."
TEMP_DIR=$(mktemp -d)
git worktree add "$TEMP_DIR" gh-pages 2>/dev/null || {
  git checkout gh-pages
  cp -r dist/. .
  git add -A
  git commit -m "Deploy $(date '+%Y-%m-%d %H:%M')" || echo "Nothing to commit"
  git push origin gh-pages
  git checkout main
  rm -rf "$TEMP_DIR"
  echo "✅ Deployed to GitHub Pages!"
  # Install locally
  mkdir -p "$WORD_ADDIN_DIR"
  cp manifest-production.xml "$WORD_ADDIN_DIR/manifest.xml"
  echo "✅ Manifest installed to Word add-ins folder"
  exit 0
}

# Worktree approach
cd "$TEMP_DIR"
cp -r "$REPO_DIR/dist/." .
git add -A
git commit -m "Deploy $(date '+%Y-%m-%d %H:%M')" || echo "Nothing to commit"
git push origin gh-pages

# Cleanup
cd "$REPO_DIR"
git worktree remove "$TEMP_DIR" --force 2>/dev/null

echo "✅ Deployed to GitHub Pages: $BASE_URL"

# Install production manifest to Word auto-load folder
mkdir -p "$WORD_ADDIN_DIR"
cp manifest-production.xml "$WORD_ADDIN_DIR/manifest.xml"
echo "✅ Manifest installed locally (Word will load it automatically)"
echo ""
echo "📦 manifest-production.xml → AppSource submission için hazır"
