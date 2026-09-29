---
name: html-to-editable-ppt
description: Convert an HTML presentation into an editable .pptx. Preserve text as native PowerPoint text, CSS cards/backgrounds as shapes, and images as independent media. Use SVG/canvas/local rasterization only as a fallback for elements that cannot be represented faithfully with native PowerPoint objects.
---

# HTML to Editable PPT

Use this skill when the user provides an HTML presentation and asks to convert it to a PowerPoint file that can be edited again in PowerPoint, WPS Presentation, or compatible software.

## Primary objective

Preserve **editability first**, then visual fidelity.

Do NOT solve the task by taking a full-slide screenshot and placing that screenshot on each PowerPoint slide unless the user explicitly requests a flattened deck.

Convert supported content as follows:

- Headings, paragraphs, labels, code, list items -> native PowerPoint text boxes.
- Solid/transparent cards, panels, backgrounds and borders -> native PowerPoint shapes.
- `<img>` -> independent PowerPoint image objects.
- SVG -> vector image object when possible; it remains independently movable/resizable but SVG internals are not PowerPoint-native.
- `<canvas>` -> local PNG fallback for that canvas region only.
- Complex CSS/WebGL/iframe/video -> preserve via a local image fallback only for the unsupported region if a reliable native conversion is not possible.
- Whole-slide rasterization is a last resort and must be disclosed.

## Preconditions

The runtime must have Node.js 20+.

Install dependencies once in the skill directory:

```bash
npm install
npx playwright install chromium
```

## Conversion command

Run:

```bash
node scripts/convert.mjs <input.html> <output.pptx>
```

Useful options:

```bash
node scripts/convert.mjs <input.html> <output.pptx> \
  --width 1280 \
  --height 720 \
  --layout wide \
  --selector '.slide' \
  --report conversion-report.json
```

Use `--selector` when automatic slide detection is wrong. Example: `--selector '.ppt-page'`.

Supported `--layout` values:

- `wide` -> 13.333 × 7.5 in (16:9)
- `standard` -> 10 × 7.5 in (4:3)

If the HTML uses a different aspect ratio, prefer matching the HTML viewport via `--width` and `--height`, then choose the closest PowerPoint layout or customize `SLIDE_W_IN` / `SLIDE_H_IN` in `scripts/convert.mjs`.

## Slide detection

The converter attempts these structures in order:

1. Reveal.js: `.reveal .slides > section`, including nested vertical sections.
2. Common slide containers: `.slide`, `[data-slide]`, `[data-slide-index]`, `.page`.
3. Large direct children of `<body>` that look like slide-sized pages.
4. Fallback: the entire `<body>` is treated as one slide.

If automatic detection is incorrect, prefer adding one of the following to the source HTML before converting:

```html
<section class="slide">...</section>
```

or:

```html
<div data-slide>...</div>
```

## Workflow for an AI agent

1. Inspect the input HTML before conversion. Check for JS runtime scaling (e.g. a `transform: scale(...)` fit-to-window stage) — the converter maps coordinates relative to each slide's own rect, so scaled stages convert correctly, but the slide element must be the scaled box itself.
2. Identify the slide framework and any unsupported content such as canvas, WebGL, iframe, video, filters, masks, complex gradients or animation-only states.
3. Run the converter. If the output file may be open in PowerPoint/WPS (EBUSY lock), write to a new versioned filename instead of retrying.
4. Read the generated conversion report when present.
5. Verify the resulting deck visually — do not skip this when a renderer is available:
   - Screenshot every HTML slide with Playwright (capture the slide element, hide nav chrome).
   - On Windows with PowerPoint installed, export every PPT slide via COM (`$slide.Export($path, 'PNG', 1280, 720)`); LibreOffice `--convert-to png` works elsewhere.
   - Compare side by side, or compute a per-slide mean pixel diff to find outlier pages, then eyeball the worst pages.
6. If a slide is materially incorrect, prefer targeted fixes in the converter or source HTML rather than flattening the entire slide.
7. Return the `.pptx` and briefly disclose any elements that were rasterized or approximated.

## Fidelity rules

- Keep editable text editable even when the font differs slightly.
- Text containers with block-level children (including `display:block` spans) are layout nodes: extract each child as its own text box, never merge them into one stacked box.
- Preserve per-span colors/weights by emitting rich-text runs for inline descendants, not a single container color.
- Preserve images as separate objects.
- Preserve element positions using browser-computed bounding rectangles, mapped relative to the slide's own rect (immune to page-level CSS/JS scaling).
- Approximate gradients as alpha-weighted average solid fills rather than dropping them or falling back to the page/body background.
- Approximate unsupported CSS rather than destroying editability of the entire slide.
- Preserve z-order as closely as practical.
- Preserve hyperlinks when the source element is an `<a>` with an absolute URL.
- Preserve code blocks as text boxes; do not screenshot them by default.
- Ignore hidden elements (`display:none`, `visibility:hidden`, near-zero opacity, zero-size boxes).
- Ignore scripts, styles, metadata and accessibility-only text.

## Post-conversion QA checklist (field-tested)

On decks with SVG charts, chip/badge pseudo-elements, or activation animations, run this checklist after `convert.mjs` — every item was hit on a real 26-page deck; detailed recipes are in `references/troubleshooting.md` ("Field lessons"):

1. **Fonts**: if a single family is required, rewrite `typeface="..."` in all XML parts; verify by listing the remaining typefaces.
2. **Duplicated chart labels**: the converter bakes SVG text into the embedded chart *and* extracts it as native text boxes. Strip duplicated `<text>` from the SVG media, or bake a text-hidden render.
3. **Chart geometry in WPS/old PowerPoint**: check PNG fallbacks — they may be tiny shared placeholders, and re-encoded SVG aspect can mismatch the pic extent (letterboxing squashes geometry). Safest fix: per-pic PNG cropped from a live render at the chart's measured DOM rect, svgBlip dropped, pic extent set to that rect, labels kept as native text boxes.
4. **Activation animations** (sankey flow-in, dash-draw): a forced all-visible render captures the pre-animation state with content translated off-canvas. Activate the slide in a live session, wait for the animation, and bake the diagram region as one image.
5. **Crop scale**: stage-scaled decks make element screenshots ≠ design pixels (e.g. 1166x656 at 1x, 2332x1312 at DSF2 for a 1280x720 design). Compute crop scale from actual PNG size / design width.
6. **Pseudo-elements** (`::before` dots/badges): invisible to the DOM walk — re-add as small shapes at positions measured from an HTML render (center-group chips put the dot ~20-30 px in, not at left padding).
7. **Inline chips merged into line runs**: capsule shapes empty, tag text off-capsule — split each tag's run into its own centered text box at the tag's DOM rect.
8. **WPS**: renders PNG fallbacks (not svgBlip); rewrites the package on open (media renames, may drop a pic); COM server can fail intermittently; the file locks while the user has it open. Re-read rels before patching media; write versioned copies when locked.
9. **Verify with expectations, not just diff**: whole-slide mean diff finds anomalies but cannot prove correctness — use target-color pixel counts at known positions and live-capture ground truth; treat ±2 diff swings on animated slides as noise.

## Known limitations

Read `references/limitations.md` before claiming pixel-perfect conversion.

For CSS/property mappings, read `references/css-to-ppt.md`.

For framework-specific selectors, read `references/frameworks.md`.

For installation/runtime problems, read `references/troubleshooting.md`.

## Output requirements

A successful run should produce:

- `<output>.pptx`
- Optional JSON report when `--report` is supplied.

The report includes slide count, extracted object counts and unsupported/fallback element counts so an agent can decide whether manual correction is necessary.
