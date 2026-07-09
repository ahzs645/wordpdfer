export { buildDocx, buildDocxDocument, dataUrlToUint8 } from "./lib/docx";
export type { DocxOptions } from "./lib/docx";
export { pdfToDocx } from "./lib/convert";
export type { PdfToDocxOptions, PdfToDocxResult } from "./lib/convert";
export { parsePdf } from "./lib/pdf";
export type { ParseOptions } from "./lib/pdf";
export type {
  DocxOverlayContent,
  DocxOverlayRun,
  DocxOverlayValue,
  DocxOverlayValues,
  FieldKind,
  FieldWidget,
  FillValues,
  PageInfo,
  ParsedPdf,
  Rect,
} from "./lib/types";
