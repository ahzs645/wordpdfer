// PDF loading, page rasterization, and AcroForm widget extraction — all in-browser via pdf.js.
import * as pdfjsLib from "pdfjs-dist";
import type { FieldWidget, PageInfo, ParsedPdf, Rect } from "./types";

export interface ParseOptions {
  /** Rasterization scale for the background image. 2 ≈ 144 DPI, crisp for print without bloating the .docx. */
  renderScale?: number;
  /** Public URL for pdf.js' module worker. Required when the host has not configured pdf.js globally. */
  workerSrc?: string;
  /** Optional public URL for pdf.js standard fonts (must end with a slash). */
  standardFontDataUrl?: string;
  onProgress?: (msg: string) => void;
}

/**
 * Parse a PDF: render every page to a PNG and extract every form widget.
 * `raw` is kept intact (pdf.js is handed a copy) so pdf-lib can fill the real AcroForm later.
 */
export async function parsePdf(
  raw: ArrayBuffer,
  fileName: string,
  opts: ParseOptions = {},
): Promise<ParsedPdf> {
  const renderScale = opts.renderScale ?? 2;
  const report = opts.onProgress ?? (() => {});

  if (opts.workerSrc) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = opts.workerSrc;
  }
  if (!pdfjsLib.GlobalWorkerOptions.workerSrc) {
    throw new Error(
      "pdf.js worker is not configured. Pass ParseOptions.workerSrc or set GlobalWorkerOptions.workerSrc before parsing.",
    );
  }

  // pdf.js detaches the buffer it is given; hand it a copy and keep `raw` pristine.
  const doc = await pdfjsLib.getDocument({
    data: new Uint8Array(raw.slice(0)),
    useSystemFonts: true,
    isEvalSupported: false,
    ...(opts.standardFontDataUrl ? { standardFontDataUrl: opts.standardFontDataUrl } : {}),
  }).promise;

  const pages: PageInfo[] = [];
  const widgets: FieldWidget[] = [];

  for (let p = 1; p <= doc.numPages; p++) {
    report(`Rendering page ${p} of ${doc.numPages}…`);
    const page = await doc.getPage(p);
    const baseViewport = page.getViewport({ scale: 1 }); // points, top-left origin, rotation applied

    // --- Rasterize the page to a PNG data URL (the Word background). ---
    const renderViewport = page.getViewport({ scale: renderScale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(renderViewport.width);
    canvas.height = Math.ceil(renderViewport.height);
    const ctx = canvas.getContext("2d", { alpha: false });
    if (!ctx) throw new Error("Could not get a 2D canvas context.");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport: renderViewport }).promise;
    const pngDataUrl = canvas.toDataURL("image/png");

    pages.push({
      index: p - 1,
      widthPt: baseViewport.width,
      heightPt: baseViewport.height,
      rotation: page.rotate,
      pngDataUrl,
      pngWidthPx: canvas.width,
      pngHeightPx: canvas.height,
    });

    // --- Extract widget annotations. ---
    const anns = await page.getAnnotations({ intent: "display" });
    let idx = 0;
    for (const a of anns) {
      if (a.subtype !== "Widget") continue;
      if (a.pushButton) continue; // action buttons carry no fillable value

      const rect = viewportRect(baseViewport, a.rect);
      const common = {
        id: `${p - 1}:${idx}`,
        page: p - 1,
        name: (a.fieldName ?? `field_${p}_${idx}`) as string,
        rect,
        readOnly: Boolean(a.readOnly),
      };
      idx++;

      if (a.fieldType === "Tx") {
        widgets.push({
          ...common,
          kind: "text",
          multiline: Boolean(a.multiLine),
          comb: Boolean(a.comb),
          maxLen: a.maxLen || undefined,
          align: alignFrom(a.textAlignment),
          defaultText: typeof a.fieldValue === "string" ? a.fieldValue : undefined,
        });
      } else if (a.fieldType === "Btn") {
        const exportValue = typeof a.exportValue === "string" ? a.exportValue : "On";
        widgets.push({
          ...common,
          kind: a.radioButton ? "radio" : "checkbox",
          exportValue,
          defaultChecked:
            a.fieldValue != null && a.fieldValue !== "Off" && a.fieldValue === exportValue,
        });
      } else if (a.fieldType === "Ch") {
        const options = Array.isArray(a.options)
          ? a.options.map((o: { displayValue?: string; exportValue?: string }) =>
              String(o.displayValue ?? o.exportValue ?? ""),
            )
          : undefined;
        widgets.push({
          ...common,
          kind: "dropdown",
          options,
          defaultText: typeof a.fieldValue === "string" ? a.fieldValue : undefined,
        });
      }
    }
  }

  report("Done.");
  return { fileName, pages, widgets, raw };
}

/** Convert a PDF-space rect [x0,y0,x1,y1] to a top-left-origin rect in points, honoring page rotation. */
function viewportRect(
  viewport: { convertToViewportRectangle: (r: number[]) => number[] },
  pdfRect: number[],
): Rect {
  const [a, b, c, d] = viewport.convertToViewportRectangle(pdfRect);
  const x = Math.min(a, c);
  const y = Math.min(b, d);
  return { x, y, w: Math.abs(c - a), h: Math.abs(d - b) };
}

function alignFrom(t: number | null | undefined): "left" | "center" | "right" {
  if (t === 1) return "center";
  if (t === 2) return "right";
  return "left";
}
