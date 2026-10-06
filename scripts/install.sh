#!/usr/bin/env sh
# dsh-docx-export installer (macOS, Linux, Git Bash, WSL)
#
# Installs the Word-export plugin into a DSH profile through the official CLI:
#   dsh plugin --profile <profile> add <spec>
#
# Usage:
#   sh scripts/install.sh                          # npm latest into "desktop"
#   sh scripts/install.sh --profile web --version 1.0.0
#   sh scripts/install.sh --from-path /path/to/dsh-docx-export
#   sh scripts/install.sh --from-repo <owner>/<repo>
#   sh scripts/install.sh --dry-run
set -eu

PKG="dsh-docx-export"
PROFILE="desktop"
VERSION="latest"
FROM_PATH=""
FROM_REPO=""
DRY_RUN=0

say()  { printf '[install] %s\n' "$1"; }
warn() { printf '[warn] %s\n' "$1" >&2; }
die()  { printf '[error] %s\n' "$1" >&2; exit 1; }

while [ $# -gt 0 ]; do
  case "$1" in
    --profile) PROFILE="$2"; shift 2 ;;
    --version) VERSION="$2"; shift 2 ;;
    --from-path) FROM_PATH="$2"; shift 2 ;;
    --from-repo) FROM_REPO="$2"; shift 2 ;;
    --dry-run) DRY_RUN=1; shift ;;
    -h|--help) sed -n '2,16p' "$0"; exit 0 ;;
    *) die "unknown argument: $1" ;;
  esac
done

if [ -n "$FROM_PATH" ]; then
  [ -e "$FROM_PATH" ] || die "path not found: $FROM_PATH"
  SPEC="$(cd "$FROM_PATH" && pwd)"
elif [ -n "$FROM_REPO" ]; then
  case "$FROM_REPO" in
    github:*|git+*|http*) SPEC="$FROM_REPO" ;;
    *) SPEC="github:$FROM_REPO" ;;
  esac
elif [ "$VERSION" = "latest" ]; then
  SPEC="$PKG"
else
  SPEC="$PKG@$VERSION"
fi

if [ -n "${DSH_CMD:-}" ]; then
  CLI="$DSH_CMD"
elif command -v dsh >/dev/null 2>&1; then
  CLI="dsh"
elif command -v npx >/dev/null 2>&1; then
  CLI="npx -y --package @deepseek-ai/dsh dsh"
else
  die "no dsh CLI found; install DSH, or set DSH_CMD."
fi

say "about to run: $CLI plugin --profile $PROFILE add $SPEC"
if [ "$DRY_RUN" -eq 1 ]; then
  say "[dry-run] nothing was executed"
  exit 0
fi

# shellcheck disable=SC2086
$CLI plugin --profile $PROFILE add "$SPEC" || die "the install command failed (see the output above)"

DSH_HOME="${DSH_HOME:-$HOME/.dsh}"
MANIFEST="$DSH_HOME/profiles/$PROFILE/package.json"
if [ -f "$MANIFEST" ]; then
  if grep -q "\"$PKG\"" "$MANIFEST"; then
    say "bundle registered: $PKG (mounts on the next DSH start)"
  else
    warn "$PKG is not listed in $MANIFEST — the mount may not be registered."
  fi
fi

say "next step: restart DSH, then ask for a Word export."
