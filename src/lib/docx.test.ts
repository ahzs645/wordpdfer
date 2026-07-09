import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { buildDocx } from "./docx";
import type { ParsedPdf } from "./types";

const ONE_PIXEL_PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

function fixture(): ParsedPdf {
  return {
    fileName: "source.pdf",
    raw: new ArrayBuffer(0),
    pages: [
      {
        index: 0,
        widthPt: 612,
        heightPt: 792,
        rotation: 0,
        pngDataUrl: ONE_PIXEL_PNG,
        pngWidthPx: 1,
        pngHeightPx: 1,
      },
    ],
    widgets: [
      {
        id: "0:0",
        page: 0,
        name: "notes",
        kind: "text",
        rect: { x: 72, y: 72, w: 240, h: 20 },
        readOnly: false,
      },
      {
        id: "0:1",
        page: 0,
        name: "approved",
        kind: "checkbox",
        rect: { x: 72, y: 110, w: 12, h: 12 },
        readOnly: false,
        exportValue: "Yes",
      },
    ],
  };
}

describe("buildDocx", () => {
  it("keeps rich overlay tags in independent Word runs", async () => {
    const blob = await buildDocx(fixture(), {
      "0:0": { runs: [{ text: "<<fields.notes>>" }] },
      "0:1": {
        runs: [
          { text: "<<cs_{fields.approved==true}>>" },
          { text: "X", bold: true },
          { text: "<<es_>>" },
        ],
      },
    });

    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("&lt;&lt;fields.notes&gt;&gt;");
    expect(xml).toContain("&lt;&lt;cs_{fields.approved==true}&gt;&gt;");
    expect(xml).toContain("&lt;&lt;es_&gt;&gt;");
    expect(xml).toMatch(/cs_\{fields\.approved==true\}[^<]*<\/w:t><\/w:r><w:r[^>]*>.*?<w:t[^>]*>X<\/w:t>/s);
  });
});
