// Build a .docx that reproduces the PDF's fixed layout:
//   • one Word section per PDF page (page size = PDF page size, zero margins)
//   • the rendered page image floats behind the text, page-anchored, full-bleed
//   • every form field becomes an absolutely-positioned text frame (framePr),
//     so nothing can reflow and the values print exactly where the fields are.
import {
  AlignmentType,
  BorderStyle,
  Document,
  FrameAnchorType,
  FrameWrap,
  HeightRule,
  HorizontalPositionRelativeFrom,
  ImageRun,
  LineRuleType,
  Packer,
  Paragraph,
  SectionType,
  TextRun,
  TextWrappingType,
  VerticalPositionRelativeFrom,
  type IParagraphOptions,
  type ISectionOptions,
} from "docx";
import type {
  DocxOverlayContent,
  DocxOverlayRun,
  DocxOverlayValue,
  DocxOverlayValues,
  FieldWidget,
  PageInfo,
  ParsedPdf,
} from "./types";

export interface DocxOptions {
  /** Word font used for filled text (Helvetica ≈ Arial). */
  fieldFont?: string;
  /** Draw a light outline around every field so blanks are easy to click in Word. */
  outlineFields?: boolean;
  /** Glyph dropped into a checked box. */
  checkGlyph?: string;
}

const TWIP_PER_PT = 20; // 1 pt = 20 twips
const PX96_PER_PT = 96 / 72; // 1 pt = 1.333 px at 96 DPI (Word's image unit)
const PAD_TWIP = 22; // small left/right inset so text doesn't hug the field edge

/** Assemble the Word document model (no packing) — shared by the browser and tests. */
export function buildDocxDocument(
  parsed: ParsedPdf,
  values: DocxOverlayValues,
  options: DocxOptions = {},
): Document {
  const opts: Required<DocxOptions> = {
    fieldFont: options.fieldFont ?? "Arial",
    outlineFields: options.outlineFields ?? false,
    checkGlyph: options.checkGlyph ?? "X",
  };

  const byPage = new Map<number, FieldWidget[]>();
  for (const w of parsed.widgets) {
    const arr = byPage.get(w.page);
    if (arr) arr.push(w);
    else byPage.set(w.page, [w]);
  }

  const sections: ISectionOptions[] = parsed.pages.map((page) =>
    buildSection(page, byPage.get(page.index) ?? [], values, opts),
  );

  return new Document({
    creator: "WordPDFer",
    title: parsed.fileName.replace(/\.pdf$/i, ""),
    description: "Fill-in-Word overlay generated from a PDF form.",
    sections,
  });
}

export async function buildDocx(
  parsed: ParsedPdf,
  values: DocxOverlayValues,
  options: DocxOptions = {},
): Promise<Blob> {
  return Packer.toBlob(buildDocxDocument(parsed, values, options));
}

function buildSection(
  page: PageInfo,
  widgets: FieldWidget[],
  values: DocxOverlayValues,
  opts: Required<DocxOptions>,
): ISectionOptions {
  const children: Paragraph[] = [backgroundParagraph(page)];

  for (const w of widgets) {
    const para = fieldParagraph(w, values[w.id], opts);
    if (para) children.push(para);
  }
  // A trailing empty paragraph anchors the section without adding visible content.
  children.push(new Paragraph({ text: "" }));

  return {
    properties: {
      type: SectionType.NEXT_PAGE,
      page: {
        size: {
          width: Math.round(page.widthPt * TWIP_PER_PT),
          height: Math.round(page.heightPt * TWIP_PER_PT),
        },
        margin: {
          top: 0,
          right: 0,
          bottom: 0,
          left: 0,
          header: 0,
          footer: 0,
          gutter: 0,
        },
      },
    },
    children,
  };
}

/** Full-page page image, floated behind the text and pinned to the page origin. */
function backgroundParagraph(page: PageInfo): Paragraph {
  return new Paragraph({
    children: [
      new ImageRun({
        type: "png",
        data: dataUrlToUint8(page.pngDataUrl),
        transformation: {
          width: Math.round(page.widthPt * PX96_PER_PT),
          height: Math.round(page.heightPt * PX96_PER_PT),
        },
        floating: {
          horizontalPosition: {
            relative: HorizontalPositionRelativeFrom.PAGE,
            offset: 0,
          },
          verticalPosition: {
            relative: VerticalPositionRelativeFrom.PAGE,
            offset: 0,
          },
          allowOverlap: true,
          behindDocument: true,
          lockAnchor: true,
          wrap: { type: TextWrappingType.NONE },
        },
      }),
    ],
  });
}

function fieldParagraph(
  w: FieldWidget,
  value: DocxOverlayValue | undefined,
  opts: Required<DocxOptions>,
): Paragraph | null {
  if (w.kind === "checkbox" || w.kind === "radio") {
    if (typeof value === "boolean" || value === undefined) {
      const checked = value === true;
      if (!checked && !opts.outlineFields) return null;
      return checkParagraph(w, checked ? [{ text: opts.checkGlyph, bold: true }] : [], opts);
    }
    return checkParagraph(w, overlayRuns(value), opts);
  }
  // text / dropdown
  const runs = overlayRuns(value);
  if (runs.length === 0 && !opts.outlineFields) {
    // Still emit an (invisible) frame so the field can be typed into in Word.
    return textParagraph(w, [], opts);
  }
  return textParagraph(w, runs, opts);
}

function textParagraph(
  w: FieldWidget,
  runs: DocxOverlayRun[],
  opts: Required<DocxOptions>,
): Paragraph {
  const xTw = Math.round(w.rect.x * TWIP_PER_PT) + PAD_TWIP;
  const yTw = Math.round(w.rect.y * TWIP_PER_PT);
  const wTw = Math.max(Math.round(w.rect.w * TWIP_PER_PT) - PAD_TWIP * 2, 120);
  const hTw = Math.max(Math.round(w.rect.h * TWIP_PER_PT), 120);
  const sizeHalfPt = fontHalfPt(w.rect.h);
  const multiline = Boolean(w.multiline);

  const base: IParagraphOptions = {
    frame: {
      type: "absolute",
      position: { x: xTw, y: yTw },
      width: wTw,
      height: hTw,
      anchor: { horizontal: FrameAnchorType.PAGE, vertical: FrameAnchorType.PAGE },
      rule: multiline ? HeightRule.ATLEAST : HeightRule.EXACT,
      wrap: FrameWrap.NONE,
    },
    alignment: alignMap(w.align),
    spacing: multiline
      ? { before: 0, after: 0 }
      : { before: 0, after: 0, line: hTw, lineRule: LineRuleType.EXACT },
    border: opts.outlineFields ? outlineBorder() : undefined,
    children: renderOverlayRuns(runs, opts.fieldFont, sizeHalfPt),
  };
  return new Paragraph(base);
}

function checkParagraph(
  w: FieldWidget,
  runs: DocxOverlayRun[],
  opts: Required<DocxOptions>,
): Paragraph {
  const xTw = Math.round(w.rect.x * TWIP_PER_PT);
  const yTw = Math.round(w.rect.y * TWIP_PER_PT);
  const wTw = Math.max(Math.round(w.rect.w * TWIP_PER_PT), 80);
  const hTw = Math.max(Math.round(w.rect.h * TWIP_PER_PT), 80);
  const sizeHalfPt = clamp(Math.round(w.rect.h * 0.95), 6, 16) * 2;

  return new Paragraph({
    frame: {
      type: "absolute",
      position: { x: xTw, y: yTw },
      width: wTw,
      height: hTw,
      anchor: { horizontal: FrameAnchorType.PAGE, vertical: FrameAnchorType.PAGE },
      rule: HeightRule.EXACT,
      wrap: FrameWrap.NONE,
    },
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: hTw, lineRule: LineRuleType.EXACT },
    border: opts.outlineFields ? outlineBorder() : undefined,
    children: renderOverlayRuns(runs, opts.fieldFont, sizeHalfPt),
  });
}

// --- helpers ---

function alignMap(a: FieldWidget["align"]): (typeof AlignmentType)[keyof typeof AlignmentType] {
  if (a === "center") return AlignmentType.CENTER;
  if (a === "right") return AlignmentType.RIGHT;
  return AlignmentType.LEFT;
}

/** Font size (half-points) sized to the field height so text fits on one line. */
function fontHalfPt(hPt: number): number {
  return clamp(Math.round(hPt * 0.72), 7, 11) * 2;
}

function outlineBorder() {
  const edge = { style: BorderStyle.SINGLE, size: 2, color: "9CC3FF", space: 0 };
  return { top: edge, bottom: edge, left: edge, right: edge };
}

function clamp(n: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, n));
}

function overlayRuns(value: DocxOverlayValue | undefined): DocxOverlayRun[] {
  if (typeof value === "string") return value ? [{ text: value }] : [];
  if (value && typeof value === "object" && Array.isArray((value as DocxOverlayContent).runs)) {
    return (value as DocxOverlayContent).runs.filter((run) => run.text.length > 0);
  }
  return [];
}

function renderOverlayRuns(
  runs: DocxOverlayRun[],
  defaultFont: string,
  defaultSize: number,
): TextRun[] {
  return runs.map(
    (run) =>
      new TextRun({
        text: run.text,
        bold: run.bold,
        font: run.font ?? defaultFont,
        size: run.size ?? defaultSize,
      }),
  );
}

export function dataUrlToUint8(dataUrl: string): Uint8Array {
  const comma = dataUrl.indexOf(",");
  const b64 = dataUrl.slice(comma + 1);
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
