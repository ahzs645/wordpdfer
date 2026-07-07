# WordPDFer

Turn a **fillable PDF (AcroForm)** into a **fill-in-Word document that keeps the exact form layout** — entirely in the browser. No server, no upload; your PDF never leaves the device.

The idea: Word isn't a fixed-layout format, so converting a form's *content* always risks reflow. Instead this app renders each PDF page to an image, **locks that image as the page background**, and drops an **absolutely-positioned text box over every form field**. The layout physically cannot move, and you can type into every field — in the browser or later in Word.

## How it works

1. **Read** — [pdf.js](https://mozilla.github.io/pdf.js/) rasterizes each page to a PNG and extracts every widget annotation (name, type, rectangle, checkbox on-states, radio groups).
2. **Edit** — the page images render in the browser with an HTML input/checkbox laid over each field at its real coordinates. Fill it in.
3. **Export Word** — [`docx`](https://docx.js.org) builds a `.docx` where:
   - each PDF page is its own Word **section** (page size = PDF page size, zero margins),
   - the page image floats **behind the text**, pinned to the page origin, full-bleed,
   - every field becomes an absolutely-positioned **text frame** (`w:framePr`, page-anchored) holding your value.
4. **Export PDF** (bonus) — [pdf-lib](https://pdf-lib.js.org) fills the *original* AcroForm fields and hands back PDF bytes.

Coordinates convert as: PDF points → top-left origin via the page viewport → twips (`pt × 20`) for Word frames, or pixels (`pt × 96/72`) for the background image.

## Run it

```bash
npm install
npm run dev      # http://localhost:5173
npm run build    # type-check + production bundle in dist/
```

Open the app, click **Try the sample form** (a 2-page, 265-field allergy record) or drop your own PDF, fill it in, and download **Word (.docx)** or **Filled PDF**.

## Notes & limits

- Designed for PDFs that already have AcroForm fields. XFA-only forms aren't supported (pdf.js reads AcroForm, not XFA).
- Word text frames are the most reliable mechanism for fixed absolute placement and print correctly, but Word is not truly fixed-layout — keep the PDF as the authoritative copy for official submissions.
- The `.docx` also stays editable in Word: every field is a real (if invisible-until-clicked) frame. Turn on **Outline fields in Word doc** to make blanks visible.
- Rotated pages are handled via the page viewport; the DOCX assumes upright (rotation 0) pages, which covers virtually all forms.

## Project layout

```
src/
  lib/
    pdf.ts       # pdf.js: render pages + extract widgets
    docx.ts      # docx: sections, background image, positioned text frames
    fillPdf.ts   # pdf-lib: fill the original AcroForm (bonus export)
    fields.ts    # group widgets into fields / radio groups
    types.ts     # shared data model
    download.ts  # blob download helper
  components/
    PageView.tsx     # one page: background + field overlay
    FieldControl.tsx # one positioned input / checkbox / select
  App.tsx            # state, toolbar, export handlers
public/
  sample.pdf         # bundled demo form
```
