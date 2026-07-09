import { buildDocx, type DocxOptions } from "./docx";
import { parsePdf, type ParseOptions } from "./pdf";
import type {
  DocxOverlayValue,
  DocxOverlayValues,
  FieldWidget,
  ParsedPdf,
} from "./types";

export interface PdfToDocxOptions {
  parse?: ParseOptions;
  docx?: DocxOptions;
  /** Values keyed by WordPDFer widget id. */
  values?: DocxOverlayValues;
  /** Resolve positioned content after the PDF widgets have been parsed. */
  resolveValue?: (
    widget: FieldWidget,
    parsed: ParsedPdf,
  ) => DocxOverlayValue | undefined;
}

export interface PdfToDocxResult {
  blob: Blob;
  parsed: ParsedPdf;
}

/** Parse a PDF and build its layout-locked Word counterpart in one call. */
export async function pdfToDocx(
  raw: ArrayBuffer | Uint8Array,
  fileName: string,
  options: PdfToDocxOptions = {},
): Promise<PdfToDocxResult> {
  const parsed = await parsePdf(toArrayBuffer(raw), fileName, options.parse);
  const values: DocxOverlayValues = { ...(options.values ?? {}) };

  if (options.resolveValue) {
    for (const widget of parsed.widgets) {
      const value = options.resolveValue(widget, parsed);
      if (value !== undefined) values[widget.id] = value;
    }
  }

  return {
    blob: await buildDocx(parsed, values, options.docx),
    parsed,
  };
}

function toArrayBuffer(raw: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (raw instanceof ArrayBuffer) return raw.slice(0);
  const copy = new Uint8Array(raw.byteLength);
  copy.set(raw);
  return copy.buffer;
}
