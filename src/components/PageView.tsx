import { memo } from "react";
import type { FieldWidget, FillValues, PageInfo } from "../lib/types";
import { FieldControl } from "./FieldControl";

interface Props {
  page: PageInfo;
  widgets: FieldWidget[];
  values: FillValues;
  pxPerPt: number;
  highlight: boolean;
  onText: (id: string, value: string) => void;
  onToggle: (id: string) => void;
}

/** One rendered page: the rasterized background plus its positioned field overlay. */
export const PageView = memo(function PageView({
  page,
  widgets,
  values,
  pxPerPt,
  highlight,
  onText,
  onToggle,
}: Props) {
  const w = Math.round(page.widthPt * pxPerPt);
  const h = Math.round(page.heightPt * pxPerPt);
  return (
    <div className="page-wrap">
      <div className="page-num">
        Page {page.index + 1}
      </div>
      <div className="page" style={{ width: w, height: h }}>
        <img
          className="page-bg"
          src={page.pngDataUrl}
          width={w}
          height={h}
          alt={`Form page ${page.index + 1}`}
          draggable={false}
        />
        <div className="page-overlay">
          {widgets.map((widget) => (
            <FieldControl
              key={widget.id}
              widget={widget}
              value={values[widget.id]}
              pxPerPt={pxPerPt}
              highlight={highlight}
              onText={onText}
              onToggle={onToggle}
            />
          ))}
        </div>
      </div>
    </div>
  );
});
