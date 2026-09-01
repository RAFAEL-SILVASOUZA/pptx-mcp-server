export const AUTHORING_GUIDE = `# How to build a presentation for pptx-mcp-server

You (the calling agent) write the ENTIRE presentation yourself as a plain HTML file
(optionally with a separate CSS file and local images alongside it) directly in the
user's workspace. This server does not render, screenshot, or design anything — it
loads your HTML in a real browser purely to read back exact positions and computed
styles, then converts what you marked into native, editable PowerPoint objects
(text boxes, shapes, pictures). No image is ever embedded as a whole-slide
background. Layout freedom (CSS: flexbox, grid, absolute positioning, custom fonts,
any styling you like) is total — the only requirement is a small markup contract so
the converter knows WHICH elements become WHAT in the final .pptx.

## Workflow

1. Write one self-contained \`.html\` file (plus any \`.css\`/image files it references
   via relative paths) into the workspace, following the contract below.
2. Call \`convert_html_to_pptx\` with the absolute path to that HTML file. That's it —
   the server reads the file, measures it in a headless browser, and writes the
   \`.pptx\`. There is no preview/screenshot step; iterate by re-reading and re-editing
   the HTML file yourself.
3. The tool result includes any warnings (e.g. an element with zero size, an image
   that couldn't be loaded, a slide sized differently from the first one) — read
   them and fix the HTML if anything looks wrong, then call the tool again.

## The markup contract

### Slides: \`data-pptx-slide\`

Each direct "slide" is any HTML element carrying a \`data-pptx-slide\` attribute, in
DOM order. Give it an explicit pixel size (this defines the deck's canvas — all
slides should use the same size):

\`\`\`html
<section data-pptx-slide style="width:1280px; height:720px; position:relative;">
  ...
</section>
\`\`\`

1280x720 (16:9) is a good default; 960x720 (4:3) also works. Whatever you pick,
keep every slide the same size. The slide element's own CSS \`background-color\`
becomes the PPTX slide's background fill.

### Content: \`data-pptx="text" | "shape" | "image"\`

Only elements carrying a \`data-pptx\` attribute become objects in the .pptx. Every
other element (wrapper divs, flex/grid containers, spacers) is pure layout
scaffolding — style and nest it however you want; it's read for layout purposes
only and never itself becomes a PPTX object.

**\`data-pptx="text"\`** — becomes an editable PowerPoint text box.
- Position/size: taken from the element's actual rendered box (works correctly
  with flexbox, grid, absolute positioning — whatever produced the layout).
- Content: the element's \`innerText\` (so nested \`<span>\`/\`<b>\` etc. inside it are
  fine — only the combined visible text is used, not per-run rich formatting).
- Style read: \`font-family\` (first font in the stack), \`font-size\`, \`font-weight\`
  (>=600 → bold), \`font-style\` (italic), \`color\`, \`text-align\`.
- Put \`data-pptx="text"\` on the smallest element that wraps exactly the text you
  want as one text box — e.g. put it on the \`<h1>\`, not on a big wrapper \`<div>\`
  that also contains other things.

**\`data-pptx="shape"\`** — becomes a rectangle (or rounded rectangle, if
\`border-radius\` is set).
- Style read: \`background-color\` (fill), \`border\` (color + width, top values used
  for all sides — keep borders uniform), \`border-radius\`, \`opacity\`.
- Gradients, box-shadows, and other advanced fills/effects are NOT read — a shape
  with a CSS gradient background will convert as a flat rectangle using its
  fallback/first solid color. Pick a real solid \`background-color\` when the shape
  itself needs to look right in PowerPoint.
- Great for card backgrounds, dividers, accent bars, colored panels — put
  \`data-pptx="text"\` elements on top of a \`data-pptx="shape"\` to build a labeled
  card (two separate PPTX objects, visually composed by their overlapping
  positions, exactly like the HTML/CSS shows them).

**\`data-pptx="image"\`** — becomes a real picture object.
- Use either \`<img data-pptx="image" src="...">\` or any element with
  \`background-image: url(...)\` and \`data-pptx="image"\`.
- \`src\`/\`url(...)\` can be: a relative path to a local file next to the HTML file, an
  absolute local path, a \`data:\` URI, or a remote \`http(s)://\` URL (fetched at
  conversion time).
- Position/size come from the element's rendered box, same as everything else —
  set \`width\`/\`height\` (or let flex/grid size it) to control how large the picture
  appears.

### What is NOT supported

CSS animations/transitions (irrelevant — only the final static layout matters),
gradients, box-shadow, filters, clip-path, and multi-run rich text formatting
(bold-only-on-part-of-a-line, mixed colors within one text box, etc.) are not
translated — they're either ignored or flattened to their nearest solid
equivalent. If you need fine per-character formatting, split the text into
multiple adjacent \`data-pptx="text"\` elements, each with its own uniform styling.

## Design guidance (not a template — use your judgment)

This is guidance, not a cage — combine these principles however best fits the
content; do not feel limited to any particular layout.

- **One idea per slide.** Presentations read better with less text per slide and
  more slides, not dense walls of text.
- **Clear visual hierarchy.** Use font size and weight (not color alone) to show
  what matters most: one dominant headline, supporting text clearly smaller.
- **A consistent grid across slides.** Reuse the same margins/gutters and a small
  set of anchor positions (e.g. header zone, content zone, footer zone) so the
  deck feels designed as a whole, not as unrelated slides.
- **Restrained palette.** 2-4 colors total (a background, a text color, one or two
  accents) reads as more professional than many colors. Reuse them consistently
  for the same kind of element across slides.
- **Type pairing.** One or two font families for the whole deck (e.g. one for
  headings, one for body) is plenty; avoid mixing many typefaces.
- **Whitespace is a design tool.** Don't fill every pixel — generous margins and
  spacing between elements make content easier to scan.
- **Purposeful imagery.** Use \`data-pptx="image"\` for real content (photos,
  diagrams, logos, icons) rather than as decoration; decorative color blocks/bars
  are cheaper and cleaner as \`data-pptx="shape"\`.
- **Alignment.** Align related elements to shared edges/centers (flexbox/grid make
  this easy) — misaligned elements are one of the fastest ways a slide looks
  unpolished.

## Minimal example (one slide, title + subtitle over an accent bar)

\`\`\`html
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { margin: 0; }
  .slide {
    width: 1280px; height: 720px; position: relative;
    background-color: #0F172A; font-family: Arial, sans-serif;
  }
  .accent-bar {
    position: absolute; left: 80px; top: 260px; width: 120px; height: 8px;
  }
  .title {
    position: absolute; left: 80px; top: 290px; width: 1000px;
    font-size: 56px; font-weight: 700; color: #FFFFFF;
  }
  .subtitle {
    position: absolute; left: 80px; top: 380px; width: 900px;
    font-size: 24px; color: #94A3B8;
  }
</style>
</head>
<body>
  <section data-pptx-slide class="slide">
    <div data-pptx="shape" class="accent-bar" style="background-color:#38BDF8;"></div>
    <h1 data-pptx="text" class="title">Quarterly Results</h1>
    <p data-pptx="text" class="subtitle">Q3 2026 — Revenue up 18% year over year</p>
  </section>
</body>
</html>
\`\`\`
`;
