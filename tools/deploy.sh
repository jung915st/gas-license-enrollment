#!/usr/bin/env bash
# Deploy on obsv01: pull from GitHub, run the tests, then upload with clasp.
#   tools/deploy.sh          pull + tests + list the files clasp would upload (no upload)
#   tools/deploy.sh --push   pull + tests + clasp push
# clasp runs through netmon's wrapper (throw-away node container, credentials in ~/.clasp-home).
# The form flow runs the latest pushed code, so no "clasp deploy" is needed for it;
# only the Wordwall confirmation web app needs a new version when confirm.gs/index.html change.
set -euo pipefail
cd "$(dirname "$0")/.."
CLASP="${CLASP_SH:-/opt/stacks/netmon/tools/clasp.sh}"
[ -f .clasp.json ] || { echo "missing .clasp.json - copy .clasp.json.example and set scriptId"; exit 1; }

echo "== 1/3 git pull"
git pull --ff-only

echo "== 2/3 tests (node:22-alpine)"
docker run --rm --user "$(id -u):$(id -g)" -v "$PWD":/w -w /w node:22-alpine node test/run-tests.js

echo "== 3/3 clasp"
export CLASP_WORKDIR="$PWD"
"$CLASP" show-file-status
if [ "${1:-}" = "--push" ]; then
  "$CLASP" push
  echo "pushed $(git log -1 --format='%h %s')"
else
  echo "dry run only - run with --push to upload"
fi
