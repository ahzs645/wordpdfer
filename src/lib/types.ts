// Shared data model for the PDF → Word overlay pipeline.

export type FieldKind = "text" | "checkbox" | "radio" | "dropdown";

/** A rectangle in PDF *points*, measured from the page's top-left corner (y grows down). */
export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** One AcroForm widget (a single on-page control). */
export interface FieldWidget {
  /** Stable unique id: `${page}:${indexOnPage}`. */
  id: string;
  /** 0-based page index. */
  page: number;
  /** AcroForm field name (shared across widgets of the same field / radio group). */
  name: string;
  kind: FieldKind;
  rect: Rect;
  readOnly: boolean;

  // Button widgets (checkbox / radio):
  /** The "on" state value for this individual widget (e.g. "Yes", "actual"). */
  exportValue?: string;

  // Text widgets:
  multiline?: boolean;
  maxLen?: number;
  comb?: boolean;
  align?: "left" | "center" | "right";

  // Choice widgets:
  options?: string[];

  // Pre-existing value in the source PDF (used to seed the editor).
  defaultText?: string;
  defaultChecked?: boolean;
}

export interface PageInfo {
  index: number;
  widthPt: number;
  heightPt: number;
  rotation: number;
  /** Rendered page bitmap, as a data URL (PNG). Used as the Word background. */
  pngDataUrl: string;
  pngWidthPx: number;
  pngHeightPx: number;
}

export interface ParsedPdf {
  fileName: string;
  pages: PageInfo[];
  widgets: FieldWidget[];
  /** Original file bytes, kept intact for pdf-lib form filling. */
  raw: ArrayBuffer;
}

/**
 * Fill state, keyed by widget id.
 *  - text / dropdown  → string
 *  - checkbox / radio → boolean (is this widget checked?)
 */
export type FillValues = Record<string, string | boolean>;
