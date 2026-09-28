# Installation

This folder follows the Agent Skills layout: one top-level skill directory containing one `SKILL.md`, with optional `scripts/` and `references/` resources.

## 1. Install runtime dependencies

Node.js 20+ is required.

### Windows PowerShell

```powershell
.\scripts\setup.ps1
```

### Linux / macOS

```bash
./scripts/setup.sh
```

Equivalent manual commands:

```bash
npm install
npx playwright install chromium
```

On Linux hosts that are missing Chromium system libraries, use:

```bash
npx playwright install --with-deps chromium
```

## 2. Make the Skill available to your agent

Use the Skill import/install mechanism supported by your Agent Skills-compatible client, or place this entire folder in a configured skill/capability directory. Keep the directory structure intact.

The ZIP artifact contains exactly one top-level directory (`html-to-editable-ppt/`), so it can also be used with systems that accept a Skill ZIP bundle.

## 3. Convert a presentation

```bash
node scripts/convert.mjs ./deck.html ./deck.pptx --report ./deck-report.json
```

If slide auto-detection is wrong:

```bash
node scripts/convert.mjs ./deck.html ./deck.pptx --selector '.ppt-page' --report ./deck-report.json
```
