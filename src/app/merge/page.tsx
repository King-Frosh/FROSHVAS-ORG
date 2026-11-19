"use client";
import Link from "next/link";
import { useEffect, useRef, useState } from "react";
import {
  CheckCircle2,
  Database,
  FileSpreadsheet,
  GitMerge,
  Trash2,
  Upload,
  X,
} from "lucide-react";
import {
  Badge,
  Button,
  Card,
  cn,
  EmptyState,
  Field,
  Input,
  Modal,
  ProgressBar,
  Segmented,
  Select,
  Skeleton,
  Toggle,
  useToast,
} from "@/components/ui";
import { vasGrain } from "@/lib/columns";
import {
  CANONICAL_FIELDS,
  exportVasFile,
  exportVasTemplate,
  guessMapping,
  type CanonicalField,
} from "@/lib/excel";
import {
  dateTimeLabel,
  money,
  num,
  parseAmount,
  parseDateValue,
  round2,
  toISODate,
  type DateFormat,
} from "@/lib/format";
import { refreshAll, useApi } from "@/state/filters";

interface PFile {
  id: string;
  name: string;
  size: number;
  file: File;
  sheets: string[];
  sheet: string;
  headerRow: number;
  dateFormat: DateFormat;
  columns: string[];
  rows: Record<string, unknown>[];
  mapping: Partial<Record<CanonicalField, string>>;
}
interface MappedRow {
  date: Date;
  servicePartner: string;
  serviceId: string;
  pricePoint: number;
  productId: string;
  productName: string;
  transaction: string;
  count: number;
  revenue: number;
  file: string;
}
interface Dataset {
  id: number;
  name: string;
  kind: string;
  sources: { file?: string; sheet?: string; rows?: number }[];
  rowCount: number;
  insertedCount: number;
  duplicateCount: number;
  skippedCount: number;
  createdAt: string;
}

function cellString(v: unknown): string {
  if (v == null) return "";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(v);
  return String(v).trim();
}

interface MergeResult {
  out: MappedRow[];
  dups: number;
  perFile: { id: string; total: number; valid: number; unknown: number }[];
  totalRead: number;
}

const EMPTY_MERGE: MergeResult = { out: [], dups: 0, perFile: [], totalRead: 0 };

async function buildMerge(
  files: PFile[],
  registryIds: Set<string>,
  dedupe: boolean,
  keep: "first" | "last" | "sum",
  knownOnly: boolean,
): Promise<MergeResult> {
  const out: MappedRow[] = [];
  const index = new Map<string, number>();
  const perFile: MergeResult["perFile"] = [];
  let dups = 0;
  let totalRead = 0;
  let processed = 0;
  const yieldToBrowser = () => new Promise<void>((resolve) => setTimeout(resolve, 0));

  for (const file of files) {
    const importDate = new Date();
    let valid = 0;
    let unknown = 0;
    totalRead += file.rows.length;

    for (const row of file.rows) {
      const get = (key?: string) => (key ? row[key] : undefined);
      const productId = cellString(get(file.mapping.productId));
      const serviceId = cellString(get(file.mapping.serviceId) || productId).replace(/\s+/g, "");
      const rawDate = file.mapping.date ? get(file.mapping.date) : null;
      const date = rawDate == null || rawDate === "" ? importDate : parseDateValue(rawDate, file.dateFormat);
      processed++;

      if (serviceId && date) {
        valid++;
        if (knownOnly && registryIds.size && !registryIds.has(serviceId)) {
          unknown++;
        } else {
          const pricePoint = parseAmount(get(file.mapping.pricePoint));
          const count = Math.max(0, Math.round(parseAmount(get(file.mapping.count))));
          let revenue = parseAmount(get(file.mapping.revenue));
          if (!revenue && pricePoint && count) revenue = round2(pricePoint * count);
          const mapped: MappedRow = {
            date,
            servicePartner: cellString(get(file.mapping.servicePartner)),
            serviceId,
            pricePoint,
            productId,
            productName: cellString(get(file.mapping.productName)),
            transaction: cellString(get(file.mapping.transaction)),
            count: count || (revenue || pricePoint ? 1 : 0),
            revenue,
            file: file.name,
          };
          const key = vasGrain(mapped);
          if (dedupe && index.has(key)) {
            dups++;
            const existing = index.get(key)!;
            if (keep === "last") out[existing] = mapped;
            else if (keep === "sum") {
              out[existing] = {
                ...out[existing],
                count: out[existing].count + mapped.count,
                revenue: round2(out[existing].revenue + mapped.revenue),
              };
            }
          } else {
            if (dedupe) index.set(key, out.length);
            out.push(mapped);
          }
        }
      }

      // Yield regularly so large imports do not monopolize the browser's UI thread.
      if (processed % 2500 === 0) await yieldToBrowser();
    }
    perFile.push({ id: file.id, total: file.rows.length, valid, unknown });
  }

  return { out, dups, perFile, totalRead };
}

type WorkerParseResult = {
  sheetNames: string[];
  sheet: string;
  headerRow: number;
  columns: string[];
  rows: Record<string, unknown>[];
};

function parseFileInWorker(file: File, sheet?: string, headerRow?: number): Promise<WorkerParseResult> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("../../workers/excel.worker.ts", import.meta.url), { type: "module" });
    worker.onmessage = (event: MessageEvent<{ ok: boolean; result?: WorkerParseResult; error?: string }>) => {
      worker.terminate();
      if (event.data.ok && event.data.result) resolve(event.data.result);
      else reject(new Error(event.data.error || "Unable to read this workbook."));
    };
    worker.onerror = (event) => {
      worker.terminate();
      reject(new Error(event.message || "Excel background worker failed."));
    };
    worker.postMessage({ file, sheet, headerRow });
  });
}

export default function MergePage() {
  const [files, setFiles] = useState<PFile[]>([]);
  const [adding, setAdding] = useState(false);
  const [drag, setDrag] = useState(false);
  const [dedupe, setDedupe] = useState(true);
  const [keep, setKeep] = useState<"first" | "last" | "sum">("sum");
  const [knownOnly, setKnownOnly] = useState(false);
  const [datasetName, setDatasetName] = useState("");
  const [committing, setCommitting] = useState(false);
  const [progress, setProgress] = useState(0);
  const [result, setResult] = useState<{ inserted: number; duplicates: number; skipped: number } | null>(null);
  const [deleteDs, setDeleteDs] = useState<Dataset | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const toast = useToast();
  const pushToast = toast.push;

  const registry = useApi<{ rows: { serviceId: string }[] }>("/api/services?lite=1");
  const datasets = useApi<{ rows: Dataset[] }>("/api/datasets");

  const addFiles = async (list: FileList | File[]) => {
    const arr = Array.from(list);
    if (!arr.length) return;
    setAdding(true);
    try {
      for (const file of arr) {
        try {
          // Excel parsing runs in a Web Worker so large workbooks do not block the UI.
          const parsed = await parseFileInWorker(file);
          setFiles((fs) => [
            ...fs,
            {
              id: `${file.name}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
              name: file.name,
              size: file.size,
              file,
              sheets: parsed.sheetNames,
              sheet: parsed.sheet,
              headerRow: parsed.headerRow,
              dateFormat: "auto",
              columns: parsed.columns,
              rows: parsed.rows,
              mapping: guessMapping(parsed.columns),
            },
          ]);
        } catch (error) {
          toast.push({
            tone: "error",
            title: `Could not read ${file.name}`,
            desc: error instanceof Error ? error.message : "Supported formats: .xlsx, .xls, .csv",
          });
        }
      }
    } finally {
      setAdding(false);
    }
  };

  const updateFile = async (id: string, patch: Partial<PFile>) => {
    const current = files.find((file) => file.id === id);
    if (!current) return;
    const nextSheet = patch.sheet ?? current.sheet;
    const nextHeaderRow = patch.headerRow ?? current.headerRow;
    if (patch.sheet !== undefined || patch.headerRow !== undefined) {
      setAdding(true);
      try {
        const parsed = await parseFileInWorker(current.file, nextSheet, nextHeaderRow);
        setFiles((fs) => fs.map((file) => file.id !== id ? file : {
          ...file,
          ...patch,
          sheet: parsed.sheet,
          headerRow: parsed.headerRow,
          columns: parsed.columns,
          rows: parsed.rows,
          mapping: guessMapping(parsed.columns),
        }));
      } catch (error) {
        toast.push({ tone: "error", title: "Could not update worksheet", desc: error instanceof Error ? error.message : "Please try again." });
      } finally {
        setAdding(false);
      }
      return;
    }
    setFiles((fs) => fs.map((file) => file.id === id ? { ...file, ...patch } : file));
  };

  const setMapping = (fileId: string, field: CanonicalField, col: string) => {
    setFiles((fs) =>
      fs.map((f) => (f.id === fileId ? { ...f, mapping: { ...f.mapping, [field]: col || undefined } } : f)),
    );
  };

  const [merge, setMerge] = useState<MergeResult>(EMPTY_MERGE);
  const [merging, setMerging] = useState(false);

  useEffect(() => {
    let cancelled = false;
    setMerging(files.length > 0);
    if (!files.length) {
      setMerge(EMPTY_MERGE);
      setMerging(false);
      return () => { cancelled = true; };
    }

    const registryIds = new Set((registry.data?.rows ?? []).map((row) => row.serviceId));
    void buildMerge(files, registryIds, dedupe, keep, knownOnly)
      .then((next) => {
        if (!cancelled) setMerge(next);
      })
      .catch((error) => {
        if (!cancelled) pushToast({
          tone: "error",
          title: "Could not prepare merge",
          desc: error instanceof Error ? error.message : "Please check the uploaded workbook.",
        });
      })
      .finally(() => {
        if (!cancelled) setMerging(false);
      });

    return () => { cancelled = true; };
  }, [files, dedupe, keep, knownOnly, registry.data, pushToast]);

  const fullyMapped = files.filter((f) => (f.mapping.serviceId || f.mapping.productId) && (f.mapping.revenue || f.mapping.pricePoint)).length;

  const commit = async () => {
    if (!merge.out.length) {
      toast.push({ tone: "error", title: "Nothing to commit", desc: "Map SERVICE ID (or PRODUCT ID for exports without VAS_NUMBER) and CHARGE/REVENUE (or PRICE POINT). Missing dates default to the import date." });
      return;
    }
    setCommitting(true);
    setProgress(0);
    setResult(null);
    try {
      const rows = merge.out.map((r) => ({
        serviceId: r.serviceId,
        servicePartner: r.servicePartner,
        productId: r.productId,
        productName: r.productName,
        transaction: r.transaction,
        count: r.count,
        pricePoint: r.pricePoint,
        revenue: r.revenue,
        reference: vasGrain(r),
        network: r.servicePartner,
        transactionAt: r.date.toISOString(),
      }));
      let datasetId: number | null = null;
      let inserted = 0;
      let duplicates = 0;
      let skipped = 0;
      for (let i = 0; i < rows.length; i += 800) {
        const chunk = rows.slice(i, i + 800);
        const res = await fetch("/api/import", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            rows: chunk,
            createDataset: i === 0,
            datasetId,
            datasetName:
              datasetName.trim() ||
              `Merge of ${files.map((f) => f.name).join(", ").slice(0, 120)} — ${new Date().toISOString().slice(0, 10)}`,
            sources: files.map((f) => ({ file: f.name, sheet: f.sheet, rows: f.rows.length })),
          }),
        });
        const d = (await res.json().catch(() => ({}))) as {
          datasetId: number | null;
          inserted: number;
          duplicates: number;
          skipped: number;
          error?: string;
          detail?: string;
          hint?: string;
        };
        if (!res.ok) throw new Error([d.error, d.detail, d.hint].filter(Boolean).join(" — ") || "Import failed");
        datasetId = d.datasetId;
        inserted += d.inserted;
        duplicates += d.duplicates;
        skipped += d.skipped;
        setProgress(Math.round(((i + chunk.length) / rows.length) * 100));
      }
      setResult({ inserted, duplicates, skipped });
      toast.push({
        tone: "success",
        title: `Merge committed — ${num(inserted)} rows stored`,
        desc: duplicates ? `${num(duplicates)} duplicate references were skipped automatically.` : undefined,
      });
      refreshAll();
      datasets.reload();
    } catch (err) {
      toast.push({
        tone: "error",
        title: "Commit failed",
        desc: err instanceof Error ? err.message : "The server rejected the batch. Please retry.",
      });
    } finally {
      setCommitting(false);
    }
  };

  const exportPreview = () => {
    exportVasFile(
      merge.out.map((r) => ({
        date: r.date,
        servicePartner: r.servicePartner,
        serviceId: r.serviceId,
        pricePoint: r.pricePoint,
        productId: r.productId,
        productName: r.productName,
        transaction: r.transaction,
        count: r.count,
        revenue: r.revenue,
      })),
      `vas-merged-${toISODate(new Date())}.xlsx`,
    );
    toast.push({
      tone: "success",
      title: "Merged Excel exported",
      desc: "DATE · SERVICE PARTNER · SERVICE ID · PRICE POINT · PRODUCT ID · PRODUCT NAME · TRANSACTION · COUNT · REVENUE",
    });
  };

  return (
    <div className="space-y-4">
      {/* Steps */}
      <div className="grid gap-3 sm:grid-cols-3">
        <Step n={1} title="Add operator files" desc={files.length ? `${files.length} file${files.length === 1 ? "" : "s"} staged` : "Drop .xlsx / .xls / .csv exports"} done={files.length > 0} />
        <Step n={2} title="Map columns" desc={`${fullyMapped}/${files.length || 0} files fully mapped`} done={files.length > 0 && fullyMapped === files.length} />
        <Step n={3} title="Merge & commit" desc={`${num(merge.out.length)} rows ready to store`} done={!!result} />
      </div>

      <div className="grid gap-4 lg:grid-cols-[350px_1fr]">
        {/* Left: files */}
        <div className="space-y-3">
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setDrag(true);
            }}
            onDragLeave={() => setDrag(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDrag(false);
              void addFiles(e.dataTransfer.files);
            }}
            className={cn(
              "rounded-lg border-2 border-dashed p-6 text-center transition-all duration-200",
              drag ? "border-brand-500 bg-brand-50 scale-[1.01]" : "border-line bg-white/70 hover:border-brand-300",
            )}
          >
            <div className="mx-auto w-fit rounded-xl bg-brand-50 p-3 text-brand-700">
              <Upload size={20} strokeWidth={2} className={adding ? "animate-bounce" : ""} />
            </div>
            <p className="mt-3 font-display text-sm font-semibold text-ink-900">Drop Excel files here</p>
            <p className="mx-auto mt-1 max-w-[240px] text-[11px] leading-relaxed text-mute">
              Drop files that use DATE · SERVICE PARTNER · SERVICE ID · PRICE POINT · PRODUCT ID · PRODUCT NAME · TRANSACTION · COUNT · REVENUE.
            </p>
            <Button variant="primary" size="md" className="mt-3" onClick={() => inputRef.current?.click()} disabled={adding}>
              {adding ? "Reading…" : "Browse files"}
            </Button>
            <button
              type="button"
              className="mt-2 block w-full text-[11px] font-medium text-brand-700 hover:underline"
              onClick={() => {
                exportVasTemplate();
                toast.push({ tone: "info", title: "Template downloaded", desc: "Same 9 columns as your operator Excel." });
              }}
            >
              Download blank template
            </button>
            <input
              ref={inputRef}
              type="file"
              multiple
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => {
                void addFiles(e.target.files ?? []);
                e.target.value = "";
              }}
            />
          </div>

          {files.map((f) => {
            const st = merge.perFile.find((p) => p.id === f.id);
            return (
              <div key={f.id} className="rounded-lg border border-line bg-white p-3 shadow-sm transition-shadow hover:shadow-md">
                <div className="flex items-start gap-2">
                  <FileSpreadsheet size={15} className="mt-0.5 shrink-0 text-brand-600" />
                  <div className="min-w-0 flex-1">
                    <p className="truncate text-xs font-semibold text-ink-900" title={f.name}>{f.name}</p>
                    <p className="num text-[10px] text-mute">{(f.size / 1024).toFixed(0)} KB · {num(f.rows.length)} data rows</p>
                  </div>
                  <button onClick={() => setFiles((fs) => fs.filter((x) => x.id !== f.id))} className="rounded p-1 text-mute hover:bg-red-50 hover:text-danger-600" aria-label={`Remove ${f.name}`}>
                    <X size={13} />
                  </button>
                </div>
                <div className="mt-2.5 grid grid-cols-2 gap-2">
                  <Field label="Sheet">
                    <Select value={f.sheet} onChange={(e) => updateFile(f.id, { sheet: e.target.value })} className="w-full">
                      {f.sheets.map((s) => (
                        <option key={s} value={s}>{s}</option>
                      ))}
                    </Select>
                  </Field>
                  <Field label="Header row">
                    <Input
                      type="number"
                      min={0}
                      value={f.headerRow}
                      onChange={(e) => updateFile(f.id, { headerRow: Math.max(0, parseInt(e.target.value || "0", 10)) })}
                    />
                  </Field>
                  <Field label="Date format" className="col-span-2">
                    <Select value={f.dateFormat} onChange={(e) => updateFile(f.id, { dateFormat: e.target.value as DateFormat })} className="w-full">
                      <option value="auto">Auto-detect (DD/MM preferred)</option>
                      <option value="dmy">DD/MM/YYYY</option>
                      <option value="mdy">MM/DD/YYYY</option>
                      <option value="ymd">YYYY-MM-DD</option>
                    </Select>
                  </Field>
                </div>
                <div className="mt-2 flex items-center gap-1.5">
                  <Badge tone={st && st.valid ? "green" : "red"}>
                    {st ? `${num(st.valid)} mappable` : "—"}
                  </Badge>
                  {st && st.unknown > 0 && <Badge tone="amber">{num(st.unknown)} outside registry</Badge>}
                </div>
              </div>
            );
          })}
        </div>

        {/* Right: mapping + merge */}
        <div className="space-y-4">
          <Card
            title="Column mapping"
            subtitle="Auto-guessed from headers — adjust per file if an operator uses different names"
            bodyClass="p-0"
          >
            {files.length === 0 ? (
              <EmptyState icon={FileSpreadsheet} title="No files staged yet" desc="Supports SERVICE_PROVIDER / VAS_NAME / SERVICES NAME / VAS_NUMBER / COUNT / CHARGE and SERVICE_PROVIDER / VAS_NAME / PRODUCT ID / COUNT / CHARGE, as well as the standard VAS layout. Output Excel uses the canonical layout." />
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full min-w-[560px] text-left text-xs">
                  <thead>
                    <tr className="border-b border-line bg-paper/70 text-[10px] uppercase tracking-wider text-mute">
                      <th className="px-3 py-2 font-semibold">Canonical field</th>
                      {files.map((f) => (
                        <th key={f.id} className="px-3 py-2 font-semibold" title={f.name}>
                          <span className="block max-w-[150px] truncate">{f.name}</span>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {CANONICAL_FIELDS.map((cf) => (
                      <tr key={cf.key} className="border-b border-line/60 last:border-0">
                        <td className="px-3 py-2 align-top">
                          <span className="flex items-center gap-1 font-semibold text-ink-900">
                            {cf.label}
                            {cf.required && <span className="text-danger-600">*</span>}
                          </span>
                          <span className="mt-0.5 block max-w-[190px] text-[10px] leading-snug text-mute">{cf.hint}</span>
                        </td>
                        {files.map((f) => {
                          const col = f.mapping[cf.key] ?? "";
                          const sample = col
                            ? f.rows.slice(0, 6).map((r) => cellString(r[col])).find((s) => s !== "")
                            : undefined;
                          return (
                            <td key={f.id} className="px-3 py-2 align-top">
                              <Select
                                value={col}
                                onChange={(e) => setMapping(f.id, cf.key, e.target.value)}
                                className={cn("w-full max-w-[170px]", !col && cf.required && "border-amber-400 bg-amber-50")}
                              >
                                <option value="">— not mapped —</option>
                                {f.columns.map((c) => (
                                  <option key={c} value={c}>{c}</option>
                                ))}
                              </Select>
                              {sample && <span className="num mt-1 block max-w-[150px] truncate text-[10px] text-mute" title={sample}>e.g. {sample}</span>}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </Card>

          <Card title="Merge rules & commit" subtitle="Duplicates match on DATE + SERVICE PARTNER + SERVICE ID + PRICE POINT + PRODUCT ID + TRANSACTION">
            <div className="grid gap-4 md:grid-cols-2">
              <div className="space-y-3">
                <Toggle checked={dedupe} onChange={setDedupe} label="Collapse duplicate settlement lines" />
                <div className={cn("transition-opacity", !dedupe && "pointer-events-none opacity-40")}>
                  <span className="mb-1 block text-[10px] font-semibold uppercase tracking-wider text-mute">On duplicate</span>
                  <Segmented
                    options={[
                      { key: "sum", label: "Sum COUNT & REVENUE" },
                      { key: "first", label: "Keep first" },
                      { key: "last", label: "Keep last" },
                    ]}
                    value={keep}
                    onChange={setKeep}
                  />
                </div>
                <Toggle checked={knownOnly} onChange={setKnownOnly} label="Keep only rows whose Service ID is in the registry" />
                <p className="text-[10px] leading-relaxed text-mute">
                  Registry holds <span className="num font-semibold text-ink-900">{num(registry.data?.rows.length ?? 0)}</span> Service IDs.
                  Rows for unknown IDs are excluded from the merge when enabled.
                </p>
              </div>
              <div className="space-y-3">
                <Field label="Dataset name" hint="Stored with the merge for audit trail">
                  <Input
                    value={datasetName}
                    onChange={(e) => setDatasetName(e.target.value)}
                    placeholder={`Auto: Merge of ${files.length || "…"} files — today`}
                  />
                </Field>
                <div className="grid grid-cols-2 gap-2">
                  <Stat label="Rows read" value={merging ? "Processing…" : num(merge.totalRead)} />
                  <Stat label="Mappable" value={num(merge.perFile.reduce((a, p) => a + p.valid, 0))} />
                  <Stat label="Duplicates removed" value={num(merge.dups)} tone="amber" />
                  <Stat label="Final rows" value={num(merge.out.length)} tone="green" />
                </div>
                {committing && <ProgressBar value={progress} />}
                <div className="flex flex-wrap gap-2">
                  <Button variant="primary" size="md" icon={GitMerge} onClick={commit} disabled={committing || merging || !merge.out.length}>
                    {committing ? `Committing ${progress}%` : merging ? "Preparing merge…" : `Commit merge (${num(merge.out.length)})`}
                  </Button>
                  <Button size="md" onClick={exportPreview} disabled={merging || !merge.out.length}>
                    Export Excel
                  </Button>
                </div>
                {result && (
                  <div className="flex items-start gap-2 rounded-md border border-brand-200 bg-brand-50 px-3 py-2.5 text-xs text-brand-800 animate-[rise_.25s_ease-out]">
                    <CheckCircle2 size={14} className="mt-0.5 shrink-0" />
                    <span>
                      Stored <strong className="num">{num(result.inserted)}</strong> rows ·{" "}
                      <strong className="num">{num(result.duplicates)}</strong> duplicates skipped ·{" "}
                      <strong className="num">{num(result.skipped)}</strong> invalid rows ignored.{" "}
                      <Link href="/transactions" className="font-semibold underline underline-offset-2">View transactions →</Link>
                    </span>
                  </div>
                )}
              </div>
            </div>

            {merge.out.length > 0 && (
              <div className="mt-4 overflow-x-auto rounded-md border border-line">
                <table className="w-full min-w-[640px] text-left text-[11px]">
                  <thead>
                    <tr className="border-b border-line bg-paper/70 text-[10px] uppercase tracking-wider text-mute">
                      <th className="px-3 py-1.5 font-semibold">DATE</th>
                      <th className="px-3 py-1.5 font-semibold">SERVICE PARTNER</th>
                      <th className="px-3 py-1.5 font-semibold">SERVICE ID</th>
                      <th className="px-3 py-1.5 text-right font-semibold">PRICE POINT</th>
                      <th className="px-3 py-1.5 font-semibold">PRODUCT ID</th>
                      <th className="px-3 py-1.5 font-semibold">PRODUCT NAME</th>
                      <th className="px-3 py-1.5 font-semibold">TRANSACTION</th>
                      <th className="px-3 py-1.5 text-right font-semibold">COUNT</th>
                      <th className="px-3 py-1.5 text-right font-semibold">REVENUE</th>
                    </tr>
                  </thead>
                  <tbody>
                    {merge.out.slice(0, 6).map((r, i) => (
                      <tr key={i} className="border-b border-line/60 last:border-0">
                        <td className="num whitespace-nowrap px-3 py-1.5 text-mute">{dateTimeLabel(r.date)}</td>
                        <td className="px-3 py-1.5 text-ink-900">{r.servicePartner || "—"}</td>
                        <td className="num px-3 py-1.5 font-medium text-ink-900">{r.serviceId}</td>
                        <td className="num px-3 py-1.5 text-right text-ink-900">{money(r.pricePoint)}</td>
                        <td className="num px-3 py-1.5 text-mute">{r.productId || "—"}</td>
                        <td className="max-w-[140px] truncate px-3 py-1.5 text-mute" title={r.productName}>{r.productName || "—"}</td>
                        <td className="px-3 py-1.5 text-mute">{r.transaction || "—"}</td>
                        <td className="num px-3 py-1.5 text-right text-ink-900">{num(r.count)}</td>
                        <td className="num px-3 py-1.5 text-right font-semibold text-ink-900">{money(r.revenue)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <p className="border-t border-line bg-paper/50 px-3 py-1.5 text-[10px] text-mute">
                  Preview of first 6 of {num(merge.out.length)} merged rows
                </p>
              </div>
            )}
          </Card>
        </div>
      </div>

      {/* Datasets */}
      <Card
        title="Stored datasets"
        subtitle="Every committed merge is versioned — delete a dataset to remove its rows"
        bodyClass="p-0"
      >
        <div className="overflow-x-auto">
          <table className="w-full min-w-[760px] text-left text-xs">
            <thead>
              <tr className="border-b border-line bg-paper/70 text-[10px] uppercase tracking-wider text-mute">
                <th className="px-3 py-2 font-semibold">Dataset</th>
                <th className="px-3 py-2 font-semibold">Source files</th>
                <th className="px-3 py-2 text-right font-semibold">Rows stored</th>
                <th className="px-3 py-2 text-right font-semibold">Duplicates skipped</th>
                <th className="px-3 py-2 text-right font-semibold">Invalid ignored</th>
                <th className="px-3 py-2 font-semibold">Created</th>
                <th className="px-3 py-2" />
              </tr>
            </thead>
            <tbody>
              {(datasets.data?.rows ?? []).map((d) => (
                <tr key={d.id} className="border-b border-line/60 transition-colors last:border-0 hover:bg-paper/60">
                  <td className="px-3 py-2">
                    <span className="flex items-center gap-2 font-medium text-ink-900">
                      <Database size={12} className="text-teal-700" />
                      <span className="max-w-[260px] truncate" title={d.name}>{d.name}</span>
                      {d.kind === "sample" && <Badge tone="gold">sample</Badge>}
                    </span>
                  </td>
                  <td className="max-w-[220px] truncate px-3 py-2 text-mute" title={(d.sources ?? []).map((s) => s.file).join(", ")}>
                    {(d.sources ?? []).map((s) => s.file).join(", ") || "—"}
                  </td>
                  <td className="num px-3 py-2 text-right font-semibold text-ink-900">{num(d.rowCount)}</td>
                  <td className="num px-3 py-2 text-right text-amber-700">{num(d.duplicateCount)}</td>
                  <td className="num px-3 py-2 text-right text-mute">{num(d.skippedCount)}</td>
                  <td className="num px-3 py-2 text-mute">{dateTimeLabel(d.createdAt)}</td>
                  <td className="px-3 py-2 text-right">
                    <button onClick={() => setDeleteDs(d)} className="rounded p-1 text-mute hover:bg-red-50 hover:text-danger-600" title="Delete dataset and its rows">
                      <Trash2 size={12} />
                    </button>
                  </td>
                </tr>
              ))}
              {!datasets.data && (
                <tr><td colSpan={7} className="p-4"><Skeleton className="h-10" /></td></tr>
              )}
              {datasets.data && datasets.data.rows.length === 0 && (
                <tr><td colSpan={7}><EmptyState icon={Database} title="No datasets yet" desc="Committed merges and generated samples appear here." /></td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>

      <Modal
        open={!!deleteDs}
        onClose={() => setDeleteDs(null)}
        title="Delete dataset?"
        desc={`“${deleteDs?.name ?? ""}” and all of its ${num(deleteDs?.rowCount ?? 0)} transactions will be permanently removed.`}
        footer={
          <>
            <Button onClick={() => setDeleteDs(null)}>Cancel</Button>
            <Button
              variant="danger"
              icon={Trash2}
              onClick={async () => {
                if (!deleteDs) return;
                await fetch(`/api/datasets/${deleteDs.id}`, { method: "DELETE" });
                toast.push({ tone: "info", title: "Dataset deleted" });
                setDeleteDs(null);
                datasets.reload();
                refreshAll();
              }}
            >
              Delete permanently
            </Button>
          </>
        }
      >
        <p className="text-xs leading-relaxed text-mute">
          Reports and dashboards recalculate immediately. The service registry is unaffected.
        </p>
      </Modal>
    </div>
  );
}

function Step({ n, title, desc, done }: { n: number; title: string; desc: string; done: boolean }) {
  return (
    <div className={cn("flex items-center gap-3 rounded-lg border p-3 transition-all", done ? "border-brand-200 bg-brand-50/60" : "border-line bg-white")}>
      <span
        className={cn(
          "num flex h-7 w-7 shrink-0 items-center justify-center rounded-full text-xs font-bold",
          done ? "bg-brand-600 text-white" : "bg-slate-100 text-mute",
        )}
      >
        {done ? <CheckCircle2 size={14} /> : n}
      </span>
      <div className="min-w-0">
        <p className="font-display text-xs font-semibold text-ink-900">{title}</p>
        <p className="truncate text-[10px] text-mute">{desc}</p>
      </div>
    </div>
  );
}

function Stat({ label, value, tone }: { label: string; value: string; tone?: "green" | "amber" }) {
  return (
    <div className="rounded-md border border-line bg-paper/60 px-2.5 py-2">
      <p className="text-[9px] font-semibold uppercase tracking-wider text-mute">{label}</p>
      <p className={cn("num mt-0.5 text-sm font-semibold", tone === "green" ? "text-brand-700" : tone === "amber" ? "text-amber-700" : "text-ink-900")}>
        {value}
      </p>
    </div>
  );
}
