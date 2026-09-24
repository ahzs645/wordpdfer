import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type DragEvent,
} from "react";
import { PageView } from "./components/PageView";
import { buildDocx } from "./lib/docx";
import { groupWidgets, widgetGroupIndex, type FieldGroup } from "./lib/fields";
import { fillPdf } from "./lib/fillPdf";
import { parsePdf } from "./lib/pdf";
import workerUrl from "pdfjs-dist/legacy/build/pdf.worker.min.mjs?url";
import { downloadBlob } from "./lib/download";
import type { FieldWidget, FillValues, ParsedPdf } from "./lib/types";

const BASE_PX_PER_PT = 96 / 72; // 100% = actual size on a 96-DPI screen

export default function App() {
  const [parsed, setParsed] = useState<ParsedPdf | null>(null);
  const [values, setValues] = useState<FillValues>({});
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [zoom, setZoom] = useState(1);
  const [highlight, setHighlight] = useState(true);
  const [outlineInDoc, setOutlineInDoc] = useState(false);
  const [dragOver, setDragOver] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const groupIndexRef = useRef<Map<string, FieldGroup>>(new Map());

  const groups = useMemo(
    () => (parsed ? groupWidgets(parsed.widgets) : []),
    [parsed],
  );
  useEffect(() => {
    groupIndexRef.current = widgetGroupIndex(groups);
  }, [groups]);

  const pageWidgets = useMemo(() => {
    const map = new Map<number, FieldWidget[]>();
    if (parsed) {
      for (const w of parsed.widgets) {
        const arr = map.get(w.page);
        if (arr) arr.push(w);
        else map.set(w.page, [w]);
      }
    }
    return map;
  }, [parsed]);

  const fieldCount = parsed?.widgets.length ?? 0;
  const pxPerPt = BASE_PX_PER_PT * zoom;

  // --- value editing (stable callbacks so memoized fields don't re-render) ---
  const onText = useCallback((id: string, value: string) => {
    setValues((prev) => ({ ...prev, [id]: value }));
  }, []);

  const onToggle = useCallback((id: string) => {
    setValues((prev) => {
      const group = groupIndexRef.current.get(id);
      const next = { ...prev };
      const turningOn = prev[id] !== true;
      if (group?.exclusive) {
        for (const w of group.widgets) next[w.id] = false;
      }
      next[id] = turningOn;
      return next;
    });
  }, []);

  // --- loading ---
  const load = useCallback(async (raw: ArrayBuffer, fileName: string) => {
    setBusy(true);
    setError(null);
    setStatus("Reading PDF…");
    try {
      const result = await parsePdf(raw, fileName, {
        workerSrc: workerUrl,
        onProgress: setStatus,
      });
      const seed: FillValues = {};
      for (const w of result.widgets) {
        if ((w.kind === "text" || w.kind === "dropdown") && w.defaultText) {
          seed[w.id] = w.defaultText;
        } else if ((w.kind === "checkbox" || w.kind === "radio") && w.defaultChecked) {
          seed[w.id] = true;
        }
      }
      setParsed(result);
      setValues(seed);
      setStatus("");
    } catch (e) {
      console.error(e);
      setError(
        e instanceof Error ? e.message : "Could not read that PDF. Is it a valid file?",
      );
    } finally {
      setBusy(false);
    }
  }, []);

  const onPickFile = useCallback(
    async (file: File | undefined | null) => {
      if (!file) return;
      if (file.type && file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
        setError("Please choose a PDF file.");
        return;
      }
      load(await file.arrayBuffer(), file.name);
    },
    [load],
  );

  const loadSample = useCallback(async () => {
    setBusy(true);
    setError(null);
    setStatus("Fetching sample form…");
    try {
      const res = await fetch(`${import.meta.env.BASE_URL}sample.pdf`);
      if (!res.ok) throw new Error("Sample form is unavailable.");
      const raw = await res.arrayBuffer();
      await load(raw, "Allergy Sensitivity Record.pdf");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not load sample.");
      setBusy(false);
    }
  }, [load]);

  const clearAll = useCallback(() => {
    setParsed(null);
    setValues({});
    setError(null);
    setStatus("");
    setZoom(1);
  }, []);

  // --- exports ---
  const baseName = (parsed?.fileName ?? "form").replace(/\.pdf$/i, "");

  const exportDocx = useCallback(async () => {
    if (!parsed) return;
    setBusy(true);
    setStatus("Building Word document…");
    try {
      const blob = await buildDocx(parsed, values, { outlineFields: outlineInDoc });
      downloadBlob(blob, `${baseName}.docx`);
      setStatus("");
    } catch (e) {
      console.error(e);
      setError(e instanceof Error ? e.message : "Word export failed.");
    } finally {
      setBusy(false);
    }
  }, [parsed, values, outlineInDoc, baseName]);

  const exportPdf = useCallback(async () => {
    if (!parsed) return;
    setBusy(true);
    setStatus("Filling PDF…");
    try {
      const bytes = await fillPdf(parsed, values, false);
      downloadBlob(bytes, `${baseName} (filled).pdf`, "application/pdf");
      setStatus("");
    } catch (e) {
      console.error(e);
      setError(e instanceof Error ? e.message : "PDF export failed.");
    } finally {
      setBusy(false);
    }
  }, [parsed, values, baseName]);

  // --- drag & drop ---
  const onDrop = useCallback(
    (e: DragEvent) => {
      e.preventDefault();
      setDragOver(false);
      onPickFile(e.dataTransfer.files?.[0]);
    },
    [onPickFile],
  );

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark" aria-hidden>
            ▤
          </div>
          <div className="brand-text">
            <h1>WordPDFer</h1>
            <p>Fill a PDF form as a Word document — layout locked.</p>
          </div>
        </div>

        <div className="topbar-actions">
          <input
            ref={fileInputRef}
            type="file"
            accept="application/pdf,.pdf"
            hidden
            onChange={(e) => onPickFile(e.target.files?.[0])}
          />
          <button
            className="btn"
            onClick={() => fileInputRef.current?.click()}
            disabled={busy}
          >
            Open PDF…
          </button>
          {parsed && (
            <button className="btn btn-ghost" onClick={clearAll} disabled={busy}>
              Clear
            </button>
          )}
          <button
            className="btn btn-primary"
            onClick={exportDocx}
            disabled={busy || !parsed}
            title="Download a Word .docx with the form image locked as the background"
          >
            ⬇ Word (.docx)
          </button>
          <button
            className="btn"
            onClick={exportPdf}
            disabled={busy || !parsed}
            title="Also fill the original PDF's form fields"
          >
            ⬇ Filled PDF
          </button>
        </div>
      </header>

      {parsed && (
        <div className="optionbar">
          <div className="opt-group">
            <button
              className="chip"
              onClick={() => setZoom((z) => Math.max(0.5, +(z - 0.1).toFixed(2)))}
              disabled={busy}
              aria-label="Zoom out"
            >
              −
            </button>
            <span className="chip-value">{Math.round(zoom * 100)}%</span>
            <button
              className="chip"
              onClick={() => setZoom((z) => Math.min(2, +(z + 0.1).toFixed(2)))}
              disabled={busy}
              aria-label="Zoom in"
            >
              +
            </button>
          </div>

          <label className="toggle">
            <input
              type="checkbox"
              checked={highlight}
              onChange={(e) => setHighlight(e.target.checked)}
            />
            Highlight fields
          </label>

          <label className="toggle">
            <input
              type="checkbox"
              checked={outlineInDoc}
              onChange={(e) => setOutlineInDoc(e.target.checked)}
            />
            Outline fields in Word doc
          </label>

          <div className="spacer" />
          <div className="meta">
            {parsed.pages.length} page{parsed.pages.length > 1 ? "s" : ""} · {fieldCount} fields
          </div>
        </div>
      )}

      {error && (
        <div className="banner banner-error" role="alert">
          {error}
          <button className="banner-close" onClick={() => setError(null)} aria-label="Dismiss">
            ✕
          </button>
        </div>
      )}

      {busy && status && (
        <div className="banner banner-info" aria-live="polite">
          <span className="spinner" /> {status}
        </div>
      )}

      <main
        className={`stage${dragOver ? " is-dragover" : ""}`}
        onDragOver={(e) => {
          e.preventDefault();
          setDragOver(true);
        }}
        onDragLeave={() => setDragOver(false)}
        onDrop={onDrop}
      >
        {!parsed ? (
          <EmptyState onOpen={() => fileInputRef.current?.click()} onSample={loadSample} busy={busy} />
        ) : (
          <div className="pages">
            {parsed.pages.map((page) => (
              <PageView
                key={page.index}
                page={page}
                widgets={pageWidgets.get(page.index) ?? []}
                values={values}
                pxPerPt={pxPerPt}
                highlight={highlight}
                onText={onText}
                onToggle={onToggle}
              />
            ))}
          </div>
        )}
      </main>

      <footer className="foot">
        Everything runs in your browser — your PDF never leaves this device.
      </footer>
    </div>
  );
}

function EmptyState({
  onOpen,
  onSample,
  busy,
}: {
  onOpen: () => void;
  onSample: () => void;
  busy: boolean;
}) {
  return (
    <div className="empty">
      <div className="empty-card">
        <div className="empty-icon" aria-hidden>
          ▤
        </div>
        <h2>Drop a fillable PDF here</h2>
        <p>
          WordPDFer reads the PDF's form fields, then hands you a Word document with the
          page image locked as a background and a typeable box over every field — so the
          layout can't shift when you print.
        </p>
        <div className="empty-actions">
          <button className="btn btn-primary" onClick={onOpen} disabled={busy}>
            Open a PDF…
          </button>
          <button className="btn" onClick={onSample} disabled={busy}>
            Try the sample form
          </button>
        </div>
        <ol className="empty-steps">
          <li>Load a PDF that already has form fields.</li>
          <li>Fill it in right here in the browser.</li>
          <li>Download it as Word (.docx) or a filled PDF.</li>
        </ol>
      </div>
    </div>
  );
}
