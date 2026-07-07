import type { FieldWidget } from "./types";

export interface FieldGroup {
  name: string;
  kind: FieldWidget["kind"];
  widgets: FieldWidget[];
  /** Only one widget may be "on" at a time (radio, or checkboxes sharing a name with distinct export values). */
  exclusive: boolean;
}

/** Group widgets by field name and decide which button groups are mutually exclusive. */
export function groupWidgets(widgets: FieldWidget[]): FieldGroup[] {
  const byName = new Map<string, FieldWidget[]>();
  for (const w of widgets) {
    const arr = byName.get(w.name);
    if (arr) arr.push(w);
    else byName.set(w.name, [w]);
  }
  const groups: FieldGroup[] = [];
  for (const [name, ws] of byName) {
    const kind = ws[0].kind;
    const isButton = kind === "radio" || kind === "checkbox";
    const distinctExports = new Set(ws.map((w) => w.exportValue));
    const exclusive =
      kind === "radio" || (isButton && ws.length > 1 && distinctExports.size > 1);
    groups.push({ name, kind, widgets: ws, exclusive });
  }
  return groups;
}

/** Map each widget id to its group, for quick lookup during editing. */
export function widgetGroupIndex(groups: FieldGroup[]): Map<string, FieldGroup> {
  const idx = new Map<string, FieldGroup>();
  for (const g of groups) for (const w of g.widgets) idx.set(w.id, g);
  return idx;
}
