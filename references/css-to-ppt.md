# CSS to PowerPoint mapping

The converter intentionally implements a conservative subset of CSS.

| HTML/CSS | PowerPoint representation |
|---|---|
| `left/top/width/height` from computed layout | x/y/w/h in slide inches |
| `font-family` | `fontFace` |
| `font-size` | PowerPoint points (`px × 0.75`) |
| `font-weight >= 600` | bold |
| `font-style: italic` | italic |
| `text-decoration: underline` | underline |
| `color` | text color + transparency |
| `text-align` | align |
| `vertical-align` / flex alignment | best-effort vertical anchor |
| `background-color` | shape fill |
| `border` | shape line |
| `border-radius` | rounded rectangle when sufficiently large |
| `opacity` | fill/text/image transparency where possible |
| `<img>` | PowerPoint image |
| inline/external SVG | SVG image object |
| `<canvas>` | PNG image fallback |
| `<a href>` | hyperlink on native text when possible |

Not mapped exactly: box-shadow, CSS filters, blend modes, masks, clip-path, backdrop-filter, complex transforms, pseudo-elements, text gradients, variable-font axes, WebGL and animations.

## Gradient and text-color handling (implemented after field reports)

| HTML/CSS | PowerPoint representation |
|---|---|
| `linear-gradient(...)` fills | solid fill of the alpha-weighted average stop color; average stop alpha becomes fill transparency. `transparent` stops (serialized by browsers as `rgba(0,0,0,0)`) lower the alpha instead of darkening the color |
| `background-clip:text` with transparent fill (text gradients) | text colored with the alpha-weighted gradient average |
| per-span `color`/`font-weight`/`font-size` inside a text block | rich-text runs on one text box (each run keeps its own style) |
| `line-height` | `lineSpacing` in points (`px × 0.75`) |
| single-line text boxes | `wrap: false`, so PowerPoint font-metric differences cannot reflow a title into its neighbors |
| `list-style-type: none` list items | plain text box, no synthetic `•` prefix |
| only `border-left/right/bottom` (accent bars) | **not converted** — only `border-top*` is read; a left accent bar is lost |
