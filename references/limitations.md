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
