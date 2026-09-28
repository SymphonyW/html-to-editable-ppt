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
