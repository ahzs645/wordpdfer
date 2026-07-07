import { memo, type CSSProperties } from "react";
import type { FieldWidget } from "../lib/types";

interface Props {
  widget: FieldWidget;
  value: string | boolean | undefined;
  pxPerPt: number;
  highlight: boolean;
  onText: (id: string, value: string) => void;
  onToggle: (id: string) => void;
}

function fontPx(w: FieldWidget, pxPerPt: number): number {
  const pt = Math.max(7, Math.min(12, w.rect.h * 0.72));
  return pt * pxPerPt;
}

/** One positioned form control layered over the page image. */
export const FieldControl = memo(function FieldControl({
  widget: w,
  value,
  pxPerPt,
  highlight,
  onText,
  onToggle,
}: Props) {
  const box: CSSProperties = {
    left: w.rect.x * pxPerPt,
    top: w.rect.y * pxPerPt,
    width: w.rect.w * pxPerPt,
    height: w.rect.h * pxPerPt,
  };
  const cls = `field${highlight ? " field--hl" : ""}`;

  if (w.kind === "checkbox" || w.kind === "radio") {
    const on = value === true;
    return (
      <button
        type="button"
        style={{ ...box, fontSize: fontPx(w, pxPerPt) * 1.1 }}
        className={`${cls} field--check${w.kind === "radio" ? " is-radio" : ""}${
          on ? " is-on" : ""
        }`}
        title={`${w.name}${w.exportValue ? ` = ${w.exportValue}` : ""}`}
        aria-pressed={on}
        onClick={() => onToggle(w.id)}
      >
        {on ? <span className="field__mark">✕</span> : null}
      </button>
    );
  }

  if (w.kind === "dropdown") {
    return (
      <select
        style={{ ...box, fontSize: fontPx(w, pxPerPt) }}
        className={`${cls} field--select`}
        title={w.name}
        value={typeof value === "string" ? value : ""}
        onChange={(e) => onText(w.id, e.target.value)}
      >
        <option value="" />
        {(w.options ?? []).map((o, i) => (
          <option key={i} value={o}>
            {o}
          </option>
        ))}
      </select>
    );
  }

  const shared = {
    className: cls,
    title: w.name,
    style: { ...box, fontSize: fontPx(w, pxPerPt), textAlign: w.align },
    value: typeof value === "string" ? value : "",
    onChange: (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) =>
      onText(w.id, e.target.value),
  };

  return w.multiline ? (
    <textarea {...shared} className={`${cls} field--multiline`} />
  ) : (
    <input type="text" {...shared} maxLength={w.maxLen || undefined} />
  );
});
