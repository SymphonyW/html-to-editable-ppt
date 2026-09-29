# Limitations

This converter is designed for editability, not browser-perfect reproduction.

## Usually preserved well

- Static 16:9 HTML slide decks.
- Reveal-like or manually paged HTML.
- Headings, body text, bullets, labels and code blocks.
- Solid-color cards, borders and simple rounded rectangles.
- Raster images and SVG assets.
- Typical absolute/flex/grid layouts after the browser computes final positions.

## Best-effort / approximate

- Shadows and glows.
- Gradients.
- Rotated/skewed containers.
- Nested clipping.
- Complex inline typography.
- CSS transforms that depend on animation state.
- Complex tables.

## Fallback content

- Canvas is rasterized to PNG at the canvas region.
- SVG is inserted as a vector image; SVG internals are not native PPT shapes.
- WebGL, embedded webpages, iframe/video and unsupported drawing surfaces require manual or local-region fallback handling.

## Fonts

PowerPoint will substitute fonts that are not installed on the machine opening the deck. This can change line wrapping and spacing.

## Frameworks

Reveal.js, Slidev, Marp and other frameworks may apply runtime transforms. The converter neutralizes common slide transforms, but unusual themes/plugins may require source-specific adjustments.

## Field-reported edge cases (handled)

- Decks whose stage is scaled at runtime with JS `transform: scale()` (fit-to-window viewers): coordinates are mapped relative to each slide's own bounding rect, so scaled stages convert at full size without dark edges. The slide element itself must be the scaled box.
- Spans styled `display:block` (stat numbers, labels): treated as block-level children so they extract as separate text boxes instead of being merged into the parent.
- Slide background is a gradient on a transparent element: uses the gradient's alpha-weighted average rather than falling back to the `<body>` background (which may be near-black in viewer chrome).

## Still approximate

- Gradient fills become solid average colors, not true gradient fills; fix manually in PowerPoint where it matters.
- Elements whose only border is a left/right/bottom accent bar lose that accent.
- Small pill/badge widths and anti-aliasing differ slightly from the browser render.

## Field-reported edge cases (2026-09 session, see troubleshooting.md for recipes)

- SVG charts: embedded `svgBlip` re-encodes can mismatch internal size/aspect (letterboxing squashes geometry vs surrounding native text boxes), and PNG fallbacks may be tiny shared placeholders that WPS actually renders. Safest fix: per-pic PNG cropped from a live render, svgBlip dropped, labels kept as native text boxes.
- SVG text is extracted as native text boxes *in addition to* being baked into the embedded SVG — labels render twice unless stripped or the render hides `svg text`.
- Charts drawn by slide-activation animations (sankey flow-in, dash-draw): a forced all-visible render captures the pre-animation state. Activate the slide in a live session, wait, and bake the diagram region as one image; per-element offsets on such slides are often inconsistent.
- CSS `::before`/`::after` pseudo-elements (status dots, chip badges) are invisible to the DOM walk and must be re-added as shapes.
- Inline styled chips (`tag`, `@mentions`) can end up merged into the parent line's runs while their capsule shapes stay empty — split runs into their own positioned text boxes.
- Decks whose viewer scales the stage: element screenshots are stage-scaled — compute crop scale from actual PNG size / design width, not from `deviceScaleFactor`.
- WPS rewrites the package on open (media renames, occasionally drops a pic+media) and intermittently fails to start its COM server; re-read rels and verify edits survived after any WPS session.
