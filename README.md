# html-to-editable-ppt

A reusable Skill package for converting an HTML slide deck into an editable PowerPoint presentation.

## Install

```bash
npm install
npx playwright install chromium
```

## Use

```bash
node scripts/convert.mjs ./deck.html ./deck.pptx --selector '.slide' --report ./deck-report.json
```

The converter uses Chromium to obtain the browser-computed layout, then maps supported elements to native PowerPoint objects with PptxGenJS.

## Design principle

Editable native elements are preferred over pixel-perfect flattening. Unsupported drawing surfaces are rasterized locally rather than flattening an entire slide.
