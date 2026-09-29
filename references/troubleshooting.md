# Troubleshooting

## `Cannot find package 'playwright'` or `pptxgenjs`

Run `npm install` in the Skill directory.

## Chromium executable is missing

Run:

```bash
npx playwright install chromium
```

On a minimal Linux server:

```bash
npx playwright install --with-deps chromium
```

## Wrong number of slides

Specify the repeating slide container explicitly:

```bash
node scripts/convert.mjs deck.html deck.pptx --selector '.slide'
```

## Fonts or line wrapping differ

Install the fonts used by the HTML on the machine doing the conversion and on the machine opening PowerPoint. PowerPoint font substitution can alter line wrapping.

## Gradients, filters or pseudo-elements are missing

Check the JSON conversion report. These CSS effects are intentionally reported when they cannot be represented reliably as native PowerPoint objects. Prefer a targeted manual fix or a local-region fallback rather than flattening the entire slide.

## Images are missing

Keep local assets inside the project tree used to run the command. The converter serves the current working directory when the input HTML is inside it, allowing ordinary relative asset paths to resolve.

## `EBUSY: resource busy or locked` when writing the .pptx

The output file is open in PowerPoint/WPS. Do not retry in a loop — write to a new versioned filename (e.g. `_v2.pptx`) and tell the user which file is current.

## Backgrounds come out black

Two known causes, both fixed in the converter; update your copy if you see them:

1. The slide element has a gradient background over a transparent `background-color`, and the converter fell back to the `<body>` background (often a dark viewer chrome color). The converter now approximates gradients with their alpha-weighted average color.
2. The HTML viewer scales the slide stage with JS `transform: scale()`; `getBoundingClientRect()` then returns scaled coordinates while mapping assumed the raw viewport, shrinking every element ~91% and exposing dark edges. The converter now maps coordinates relative to each slide's own rect.

## Text is stacked/merged into one box, or colors inside a paragraph are lost

Older converter versions merged any `H*`/`P`/`LI` element into a single text box, stacking its block children and flattening inline span colors. Current versions treat containers with block-level children (including `display:block` spans) as layout nodes and emit rich-text runs for inline styles.

## `npm install` fails with `notarget` for playwright/pptxgenjs

Some registries (mirrors, internal proxies) do not carry the pinned versions. Install the newest versions the registry actually has, e.g. `npm view playwright versions` then `npm install playwright@<available> pptxgenjs@<available>`.

---

# Field lessons from a 26-page corporate deck (2026-09 session)

Everything below was hit on a real deck and verified with pixel-level checks. Recipes assume a Python + Playwright + PowerPoint/WPS-COM toolkit.

## Chart renders as a stretched blob / vertices misaligned (WPS and older PowerPoint)

Cause: the converter embeds charts as `svgBlip` + a PNG fallback. Observed defects: fallback PNGs were tiny shared placeholders (all 100x119, identical bytes) used by *several different* pics; and re-encoded SVG internal size/aspect mismatched the pic extent (e.g. width/height 949x375 vs viewBox 0 0 1184 420), so `preserveAspectRatio` letterboxing + stretch squashed geometry relative to the native text boxes around it.

Fix (robust): go PNG-only per pic.

1. Measure each chart's true rect in the live deck (Playwright `getBoundingClientRect`, normalized to design px).
2. Hide `svg text` via injected CSS, screenshot the slide, crop the chart rect (this avoids double-rendering labels, see next section).
3. Write the crop as the pic's own media file, point `a:blip/@r:embed` at it, and delete the `svgBlip` extension (otherwise SVG renderers keep using the vector version and ignore your fix).
4. Set the pic `a:off/a:ext` to the measured DOM rect.

Labels stay as native editable text boxes; only the diagram body becomes an image.

## SVG chart labels are duplicated (baked into image AND extracted as text boxes)

The converter serializes the whole SVG (including its `<text>`/`<tspan>` labels) as an image *and* extracts those texts as native text boxes — every label renders twice with an offset. If you keep the vector/PNG chart, strip duplicated `<text>` elements from the embedded SVG media, or bake a text-hidden render (CSS `svg text{display:none}`). Match multi-`tspan` elements part-by-part: each `tspan` may correspond to its own run (e.g. label + value), so a whole-element match fails.

## Charts whose content is drawn by activation animations (sankey flow-in, dash-draw)

A forced all-slides-visible render captures the *pre-animation* state (flows undrawn) and such SVGs are often translated off-canvas when inactive — cropping them yields "nodes but no flows". These must be captured in a live session:

1. Load the deck normally, set the target slide active (e.g. `slides[i].classList.add('active')` after clearing others), wait 3-5 s for the animation.
2. Screenshot the slide element; bake the whole diagram region as one image.
3. Delete the mis-positioned native elements in that region (on such slides per-element offsets are often inconsistent, e.g. -49/-70/-102 px mixed — region baking beats box-by-box patching).

Keep page-level text outside the baked region as native boxes; re-add any element accidentally inside the crop.

## Page-level JS stage scale breaks crop math

If the deck's viewer scales the stage (here `scale(0.91)`), element screenshots come out at stage scale times deviceScaleFactor — for design width 1280 the 1x render was 1166x656 and the DSF2 render 2332x1312. Compute the crop scale from actual PNG size / design width (2332/1280 = 1.821875), never assume `deviceScaleFactor` alone. Otherwise every crop is mis-scaled and mis-offset.

## Deck has breathing/pulse animations

Slides with looping animations (pulsing dots, marquee) make HTML re-renders differ by ±2 whole-slide mean diff. Re-render the same slide twice and compare before believing a pixel-diff regression.

## Inline tags/pills merged into the line's text box; capsule outlines left empty

Inline styled spans (`tag` chips, `@mentions`) inside a line get appended to the parent line's runs — the text lands off its capsule and the capsule shape stays empty. Fix per tag entry:

1. Find the sp holding a run with exactly the tag text at the tag's y (box width > ~300 px, |Δy| < 25).
2. Delete that run (and a following whitespace-only run).
3. Append a new `wrap="none" anchor="ctr"` text box at the tag's DOM rect with the tag color, centered.

Match runs against XML-escaped text (`&amp;`) and normalize whitespace; the same tag text may occur on several rows — one DOM rect entry per occurrence. Accept that you may find more tag instances than a loose audit reports (one deck: audit found 4, DOM dump found 6).

## CSS `::before`/`::after` pseudo-elements are dropped (status dots, chip badges)

A DOM walk cannot see pseudo-elements. `status-chip::before` dots disappear everywhere. Re-add them as small shapes (7 px ellipse) at chip-left positions, colored per variant class parsed from computed styles, and place them by z-order *above* the chip fill (one case needed moving the dot to the end of `spTree`). Positions: measure the dot's centroid from an HTML render rather than assuming left-padding — chips that center `dot+text` as a group put the dot ~20-30 px in.

## WPS specifics (renderer + package rewriting)

- WPS may render the PNG **fallback**, not the `svgBlip`, so garbage fallbacks are user-visible there.
- Opening (and merely viewing) a pptx in WPS can **normalize and rewrite the package**: media renamed (`imageF4_0.png` → `image4.png`), file size changes, and in one case a pic + its media were dropped entirely. Consequences: after any WPS session, re-read the current rels mapping before patching media — do not assume the filename you wrote last time still exists; and verify your earlier edits survived.
- WPS COM (`PowerPoint.Application` CLSID) intermittently fails to start (`CO_E_SERVER_EXEC_FAILURE`) after heavy use. Cooldown ~15 s and retry; if it stays down, verify at media/XML level instead.
- The user having the file open in WPS locks it (`~$` owner file appears, `PermissionError` on replace). Retry with sleeps; if still locked, write a versioned copy and tell the user which file is current — never kill their WPS.

## Unifying fonts to one family after conversion

Rewrite every `typeface="..."` in all XML parts (Python `zipfile` repack, regex over `.xml` entries), then verify the set of remaining typefaces is exactly the target. When repacking: entries added to the package (new media) must be written explicitly — an `infolist()` loop only re-emits entries that already existed.

## Triage: converter bug vs source bug

Before patching the pptx, compare the chart's raw data in the HTML (polygon `points`, circle `cx/cy`) against the intended design. A real deck had radar vertex marker circles hand-placed up to 42 viewBox units off their polygon vertices (a 95% dot appearing outside the hexagon) — a source bug. Fix the HTML, re-render, re-crop; keep source and pptx in sync.

## Verification methodology (what pixel checks can and cannot prove)

- Whole-slide mean diff finds anomalies but cannot prove correctness: a wrongly stretched image can *lower* the diff by covering more area.
- Verify specific expectations: target-color pixel counts at known positions (dots, flow bands), per-label box vs DOM rect offsets, region cross-correlation offsets between HTML and PPT renders.
- Median sampling over small text regions mistakes bold colored text for background — pair it with targeted color counting.
- "Missing text" audit hits are often false positives: parent direct-text vs inline-span interleaving, `&amp;` unescaping, and duplicate texts matching the wrong instance. Confirm with distinctive-fragment greps in the XML before treating anything as dropped.
- Ground truth for animated/deck-JS pages is a live, activated capture — not a forced all-visible render.
