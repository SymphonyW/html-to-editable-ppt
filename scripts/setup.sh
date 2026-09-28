#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")/.."
npm install
npx playwright install chromium
printf '\nSetup complete. Example:\n  npm run convert -- ./examples/demo.html ./demo.pptx --report ./demo-report.json\n'
