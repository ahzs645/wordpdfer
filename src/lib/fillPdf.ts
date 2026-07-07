// Bonus export: fill the *real* AcroForm with pdf-lib and hand back PDF bytes.
// Best-effort — the .docx overlay is the faithful, layout-locked output; this is a
// convenience for anyone who also wants the values in the original PDF fields.
import { PDFDocument } from "pdf-lib";
import { groupWidgets } from "./fields";
import type { FillValues, ParsedPdf } from "./types";

export async function fillPdf(
  parsed: ParsedPdf,
  values: FillValues,
  flatten = false,
): Promise<Uint8Array> {
  const pdf = await PDFDocument.load(parsed.raw);
  const form = pdf.getForm();
  const groups = groupWidgets(parsed.widgets);

  for (const g of groups) {
    try {
      if (g.kind === "text") {
        const w0 = g.widgets[0];
        const v = values[w0.id];
        if (typeof v === "string" && v) {
          const tf = form.getTextField(g.name);
          // Pin a sane size — otherwise pdf-lib auto-fits (font 0) and tall
          // multi-line boxes render enormous text.
          const size = clamp(Math.round(w0.rect.h * 0.72), 7, 11);
          try {
            tf.setFontSize(size);
          } catch {
            // Some fields keep their /DA on the widget, not the field object,
            // which makes setFontSize throw. Give the field its own /DA instead.
            try {
              tf.acroField.setDefaultAppearance(`/Helv ${size} Tf 0 g`);
            } catch {
              /* leave the appearance as authored */
            }
          }
          tf.setText(v); // last — so the value is always written even if sizing failed
        }
      } else if (g.kind === "dropdown") {
        const v = values[g.widgets[0].id];
        if (typeof v === "string" && v) {
          const dd = form.getDropdown(g.name);
          if (!dd.getOptions().includes(v)) dd.addOptions([v]);
          dd.select(v);
        }
      } else {
        // checkbox / radio
        const on = g.widgets.find((w) => values[w.id] === true);
        if (g.exclusive) {
          try {
            const rg = form.getRadioGroup(g.name);
            if (on?.exportValue) rg.select(on.exportValue);
          } catch {
            setCheckboxes(form, g.widgets, values);
          }
        } else {
          setCheckboxes(form, g.widgets, values);
        }
      }
    } catch {
      // Unusual field shape — skip it rather than aborting the whole export.
    }
  }

  if (flatten) form.flatten();
  return pdf.save();
}

function clamp(n: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, n));
}

function setCheckboxes(
  form: ReturnType<PDFDocument["getForm"]>,
  widgets: { id: string; name: string }[],
  values: FillValues,
): void {
  const name = widgets[0].name;
  const cb = form.getCheckBox(name);
  const anyOn = widgets.some((w) => values[w.id] === true);
  if (anyOn) cb.check();
  else cb.uncheck();
}
