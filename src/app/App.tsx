import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { Upload, FileSpreadsheet, Database, Download, X, AlertTriangle, CheckCircle, ChevronDown, ChevronRight, Trash2 } from "lucide-react";

// ─── Types ───────────────────────────────────────────────────────────────────

interface BomRow {
  partNo: string;
  desc1: string;
  desc2: string;
  drawingRef: string;
  qty: string | number;
  pOrM: string;
  cost: string | number;
  supplier: string;
  leadTime: string | number;
  masterPlanningFamily: string;
  commsClass: string;
  subClass: string;
  unit: string;
  branch: string;
}

interface BomFile {
  id: string;
  fileName: string;
  partNumber: string;
  rows: BomRow[];
}

interface ProcessedPart extends BomRow {
  isNew: boolean;
  hasOwnBom: boolean;
}

interface AssemblyResult {
  assemblyPn: string;
  newParts: ProcessedPart[];
  allParts: ProcessedPart[];
  bomRows: BomRow[];
}

// ─── Column detection ─────────────────────────────────────────────────────────

function findCol(headers: string[], keywords: string[]): number {
  const lower = headers.map((h) => h.toLowerCase().trim().replace(/[\s\-_/]+/g, ""));
  for (const kw of keywords) {
    const idx = lower.findIndex((h) => h.includes(kw));
    if (idx !== -1) return idx;
  }
  return -1;
}

function parseSheetToRows(ws: XLSX.WorkSheet): BomRow[] {
  const raw = XLSX.utils.sheet_to_json<(string | number)[]>(ws, { header: 1, defval: "" });
  if (raw.length < 2) return [];

  // Find header row
  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(10, raw.length); i++) {
    const nonEmpty = (raw[i] as (string | number)[]).filter((c) => String(c).trim() !== "").length;
    if (nonEmpty >= 2) { headerRowIdx = i; break; }
  }

  const headers = (raw[headerRowIdx] as (string | number)[]).map((h) => String(h));

  const colPN   = findCol(headers, ["partno", "partnum", "part", "pn", "item", "number", "code"]);
  const colD1   = findCol(headers, ["description1", "desc1", "description"]);
  const colD2   = findCol(headers, ["description2", "desc2"]);
  const colDRef = findCol(headers, ["drawingref", "drawref", "drawing", "dwg", "ref"]);
  const colQty  = findCol(headers, ["qty", "quantity", "qnty"]);
  const colPoM  = findCol(headers, ["porm", "porm", "pm"]);
  const colCost = findCol(headers, ["updatecost", "cost", "price"]);
  const colSup  = findCol(headers, ["supplier", "vendor"]);
  const colLT   = findCol(headers, ["leadtime", "lead", "weeks", "lt"]);
  const colMPF  = findCol(headers, ["masterplanning", "masterpla", "planningfamily", "mpf"]);
  const colCC   = findCol(headers, ["commsclass", "comms"]);
  const colSC   = findCol(headers, ["subclass", "sub"]);
  const colUnit = findCol(headers, ["unit", "uom"]);
  const colBr   = findCol(headers, ["branch", "br"]);

  const get = (row: (string | number)[], idx: number): string =>
    idx >= 0 ? String(row[idx] ?? "").trim() : "";
  const getNum = (row: (string | number)[], idx: number): string | number => {
    if (idx < 0) return "";
    const v = row[idx];
    return v !== undefined && v !== "" ? v : "";
  };

  const rows: BomRow[] = [];
  for (let i = headerRowIdx + 1; i < raw.length; i++) {
    const r = raw[i] as (string | number)[];
    const pn = colPN >= 0 ? String(r[colPN] ?? "").trim() : "";
    if (!pn) continue;
    rows.push({
      partNo: pn,
      desc1: get(r, colD1) || (colD1 < 0 && colPN >= 0 ? "" : ""),
      desc2: get(r, colD2),
      drawingRef: get(r, colDRef),
      qty: getNum(r, colQty),
      pOrM: get(r, colPoM),
      cost: getNum(r, colCost),
      supplier: get(r, colSup),
      leadTime: getNum(r, colLT),
      masterPlanningFamily: get(r, colMPF),
      commsClass: get(r, colCC),
      subClass: get(r, colSC),
      unit: get(r, colUnit),
      branch: get(r, colBr),
    });
  }
  return rows;
}

async function readFile(file: File): Promise<XLSX.WorkBook> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target?.result, { type: "array" });
      resolve(wb);
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ─── Excel generation ─────────────────────────────────────────────────────────

const NEW_PARTS_HEADERS = [
  "", "Part No.", "Description 1", "Description 2", "Drawing Ref",
  "P or M", "Cost", "Supplier", "Lead-Time\n(weeks)",
  "Master Planning\nFamily", "Comms\nClass", "Sub Class", "Unit", "Branch",
];

const BOM_HEADERS = [
  "Part No.", "Description 1", "Description 2", "Drawing Ref",
  "Qty", "Update Cost", "Supplier", "Lead-Time\n(weeks)",
  "Master Planning\nFamily", "Comms\nClass", "Sub Class", "Unit", "Branch",
];

function rowToNewPartsArr(p: ProcessedPart): (string | number)[] {
  return [
    p.hasOwnBom ? "NEW Part / BOM" : "NEW Part",
    p.partNo, p.desc1, p.desc2, p.drawingRef,
    p.pOrM, p.cost, p.supplier, p.leadTime,
    p.masterPlanningFamily, p.commsClass, p.subClass, p.unit, p.branch,
  ];
}

function rowToBomArr(r: BomRow): (string | number)[] {
  return [
    r.partNo, r.desc1, r.desc2, r.drawingRef,
    r.qty, r.cost, r.supplier, r.leadTime,
    r.masterPlanningFamily, r.commsClass, r.subClass, r.unit, r.branch,
  ];
}

function generateOutputExcel(results: AssemblyResult[], bomFiles: BomFile[]): void {
  const wb = XLSX.utils.book_new();
  const data: (string | number)[][] = [];
  const EMPTY = new Array(14).fill("");

  // ── Section 1: all new parts across every assembly, deduplicated ──
  const seenNew = new Set<string>();
  const allNewParts: ProcessedPart[] = [];
  for (const result of results) {
    for (const p of result.newParts) {
      const key = normalizePN(p.partNo);
      if (!seenNew.has(key)) { seenNew.add(key); allNewParts.push(p); }
    }
  }

  data.push(NEW_PARTS_HEADERS);
  if (allNewParts.length === 0) {
    data.push(["(all parts already in JDGEs)", ...new Array(13).fill("")]);
  } else {
    for (const p of allNewParts) data.push(rowToNewPartsArr(p));
  }

  data.push(EMPTY);
  data.push(EMPTY);

  // ── Section 2: all BOM blocks sorted ascending by assembly PN ──
  // Collect all unique assemblies that have a BOM (top-level results + any sub-assembly BOM files)
  const bomMap = new Map<string, BomRow[]>();
  for (const result of results) {
    bomMap.set(normalizePN(result.assemblyPn), result.bomRows);
  }
  // Also include BOM files that weren't a top-level result (pure sub-assemblies)
  for (const bf of bomFiles) {
    const key = normalizePN(bf.partNumber);
    if (!bomMap.has(key)) bomMap.set(key, bf.rows);
  }

  const sortedPns = [...bomMap.keys()].sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));

  for (const pnKey of sortedPns) {
    const rows = bomMap.get(pnKey)!;
    // Find the display PN (original casing)
    const originalPn = results.find((r) => normalizePN(r.assemblyPn) === pnKey)?.assemblyPn
      ?? bomFiles.find((bf) => normalizePN(bf.partNumber) === pnKey)?.partNumber
      ?? pnKey;

    data.push([originalPn, ...new Array(13).fill("")]);
    data.push(BOM_HEADERS);
    for (const r of rows) data.push(rowToBomArr(r));
    data.push(EMPTY);
  }

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [
    { wch: 16 }, { wch: 14 }, { wch: 30 }, { wch: 20 }, { wch: 14 },
    { wch: 8 }, { wch: 10 }, { wch: 24 }, { wch: 10 },
    { wch: 14 }, { wch: 8 }, { wch: 10 }, { wch: 6 }, { wch: 8 },
  ];

  XLSX.utils.book_append_sheet(wb, ws, "JDGEs Load");

  XLSX.writeFile(wb, "JDGEs_BOM_Load.xlsx");
}

function normalizePN(pn: string): string {
  return pn.trim().toUpperCase().replace(/\s+/g, "");
}

// ─── Drop Zone ────────────────────────────────────────────────────────────────

function DropZone({ label, accept, multiple, onFiles, fileCount, hint }: {
  label: string; accept: string; multiple: boolean;
  onFiles: (files: File[]) => void; fileCount: number; hint?: string;
}) {
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  return (
    <div
      onClick={() => inputRef.current?.click()}
      onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => { e.preventDefault(); setDragging(false); onFiles(Array.from(e.dataTransfer.files)); }}
      className={`relative cursor-pointer border-2 border-dashed p-6 flex flex-col items-center gap-2 select-none transition-colors
        ${dragging ? "border-accent bg-accent/5" : "border-border hover:border-foreground/40"}`}
    >
      <input ref={inputRef} type="file" accept={accept} multiple={multiple} className="hidden"
        onChange={(e) => { onFiles(Array.from(e.target.files ?? [])); e.target.value = ""; }} />
      <Upload className="w-5 h-5 text-muted-foreground" />
      <div className="text-center">
        <p className="font-semibold text-sm">{label}</p>
        {hint && <p className="text-xs text-muted-foreground mt-0.5">{hint}</p>}
      </div>
      {fileCount > 0 && (
        <span className="absolute top-2 right-2 bg-accent text-accent-foreground text-xs font-mono px-1.5 py-0.5">
          {fileCount} file{fileCount !== 1 ? "s" : ""}
        </span>
      )}
    </div>
  );
}

function FileTag({ name, onRemove }: { name: string; onRemove: () => void }) {
  return (
    <div className="flex items-center gap-1.5 bg-secondary border border-border px-2 py-1 text-xs font-mono">
      <FileSpreadsheet className="w-3 h-3 text-muted-foreground shrink-0" />
      <span className="truncate max-w-[180px]">{name}</span>
      <button onClick={(e) => { e.stopPropagation(); onRemove(); }} className="ml-auto text-muted-foreground hover:text-destructive transition-colors shrink-0">
        <X className="w-3 h-3" />
      </button>
    </div>
  );
}

// ─── Results Table ────────────────────────────────────────────────────────────

const NEW_PARTS_COLS = ["Part No.", "Description 1", "Description 2", "Drawing Ref", "P or M", "Cost", "Supplier", "Lead-Time", "Mst.Plan.Family", "Comms Class", "Sub Class", "Unit", "Branch"];
const BOM_COLS       = ["Part No.", "Description 1", "Description 2", "Drawing Ref", "Qty", "Update Cost", "Supplier", "Lead-Time", "Mst.Plan.Family", "Comms Class", "Sub Class", "Unit", "Branch"];

function partToRow(p: ProcessedPart): (string | number)[] {
  return [p.partNo, p.desc1, p.desc2, p.drawingRef, p.pOrM, p.cost, p.supplier, p.leadTime, p.masterPlanningFamily, p.commsClass, p.subClass, p.unit, p.branch];
}
function bomToRow(r: BomRow): (string | number)[] {
  return [r.partNo, r.desc1, r.desc2, r.drawingRef, r.qty, r.cost, r.supplier, r.leadTime, r.masterPlanningFamily, r.commsClass, r.subClass, r.unit, r.branch];
}

function DataTable({ cols, rows, renderLabel }: {
  cols: string[];
  rows: (string | number)[][];
  renderLabel?: (i: number) => React.ReactNode;
}) {
  return (
    <div className="overflow-x-auto">
      <table className="w-full text-xs font-mono border-collapse">
        <thead>
          <tr className="bg-muted border-b border-border">
            {renderLabel && <th className="text-left px-2 py-1.5 font-semibold whitespace-nowrap w-[110px]">Label</th>}
            {cols.map((c) => (
              <th key={c} className="text-left px-2 py-1.5 font-semibold whitespace-nowrap">{c}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.length === 0 && (
            <tr><td colSpan={cols.length + (renderLabel ? 1 : 0)} className="px-2 py-3 text-center text-muted-foreground">No rows</td></tr>
          )}
          {rows.map((row, i) => (
            <tr key={i} className="border-b border-border/40 hover:bg-secondary/50">
              {renderLabel && <td className="px-2 py-1">{renderLabel(i)}</td>}
              {row.map((cell, j) => (
                <td key={j} className={`px-2 py-1 ${j === 0 ? "font-semibold" : "text-muted-foreground"}`}>
                  {cell !== "" && cell !== undefined ? String(cell) : <span className="opacity-30">—</span>}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function ResultsSection({ result, expanded, onToggle }: {
  result: AssemblyResult; expanded: boolean; onToggle: () => void;
}) {
  const [tab, setTab] = useState<"new" | "bom">("new");

  return (
    <div className="border border-border mb-0">
      <button onClick={onToggle}
        className="w-full flex items-center justify-between px-4 py-3 bg-secondary hover:bg-muted transition-colors text-left">
        <div className="flex items-center gap-3">
          {expanded ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
          <span className="font-mono font-bold text-sm">{result.assemblyPn}</span>
          <span className="text-xs text-muted-foreground">{result.bomRows.length} BOM lines</span>
        </div>
        <div className="flex items-center gap-3">
          {result.newParts.length > 0 ? (
            <span className="flex items-center gap-1 text-xs font-mono text-accent font-bold">
              <AlertTriangle className="w-3.5 h-3.5" />{result.newParts.length} new
            </span>
          ) : (
            <span className="flex items-center gap-1 text-xs font-mono text-green-700 font-bold">
              <CheckCircle className="w-3.5 h-3.5" />all in JDGEs
            </span>
          )}
        </div>
      </button>

      {expanded && (
        <div className="border-t border-border">
          <div className="flex border-b border-border">
            {(["new", "bom"] as const).map((t) => (
              <button key={t} onClick={() => setTab(t)}
                className={`px-4 py-2 text-xs font-bold uppercase tracking-wider transition-colors
                  ${tab === t ? "bg-foreground text-primary-foreground" : "text-muted-foreground hover:text-foreground"}`}>
                {t === "new" ? `New Parts (${result.newParts.length})` : `Full BOM (${result.bomRows.length})`}
              </button>
            ))}
          </div>

          {tab === "new" && (
            <DataTable
              cols={NEW_PARTS_COLS}
              rows={result.newParts.map((p) => partToRow(p))}
              renderLabel={(i) => {
                const p = result.newParts[i];
                return p.hasOwnBom
                  ? <span className="bg-yellow-300 text-black px-1 py-0.5 font-bold text-[10px] whitespace-nowrap">NEW Part / BOM</span>
                  : <span className="bg-yellow-200 text-black px-1 py-0.5 font-bold text-[10px] whitespace-nowrap">NEW Part</span>;
              }}
            />
          )}
          {tab === "bom" && (
            <DataTable cols={BOM_COLS} rows={result.bomRows.map((r) => bomToRow(r))} />
          )}
        </div>
      )}
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────

export default function App() {
  const [bomFiles, setBomFiles] = useState<BomFile[]>([]);
  const [jdgesFile, setJdgesFile] = useState<File | null>(null);
  const [jdgesPartNumbers, setJdgesPartNumbers] = useState<Set<string>>(new Set());
  const [jdgesLoaded, setJdgesLoaded] = useState(false);
  const [jdgesCount, setJdgesCount] = useState(0);
  const [processing, setProcessing] = useState(false);
  const [results, setResults] = useState<AssemblyResult[]>([]);
  const [expandedAssemblies, setExpandedAssemblies] = useState<Set<string>>(new Set());
  const [error, setError] = useState<string | null>(null);

  const handleBomFiles = useCallback(async (files: File[]) => {
    setError(null);
    const parsed: BomFile[] = [];
    for (const file of files) {
      try {
        const wb = await readFile(file);
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = parseSheetToRows(ws);
        const partNumber = file.name.replace(/\.(xlsx|xls|csv)$/i, "").trim();
        parsed.push({ id: `${file.name}-${Date.now()}`, fileName: file.name, partNumber, rows });
      } catch {
        setError(`Failed to parse ${file.name}`);
      }
    }
    setBomFiles((prev) => {
      const existingNames = new Set(prev.map((f) => f.fileName));
      return [...prev, ...parsed.filter((f) => !existingNames.has(f.fileName))];
    });
    setResults([]);
  }, []);

  const handleJdgesFile = useCallback(async (files: File[]) => {
    setError(null);
    const file = files[0];
    if (!file) return;
    try {
      const wb = await readFile(file);
      const ws = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json<(string | number)[]>(ws, { header: 1, defval: "" });
      const pns = new Set<string>();
      raw.flat().map((v) => String(v).trim()).filter((v) => v.length > 1).forEach((v) => pns.add(normalizePN(v)));
      setJdgesPartNumbers(pns);
      setJdgesFile(file);
      setJdgesLoaded(true);
      setJdgesCount(pns.size);
      setResults([]);
    } catch {
      setError("Failed to parse JDGEs database file.");
    }
  }, []);

  const handleProcess = useCallback(() => {
    if (bomFiles.length === 0) { setError("Upload at least one BOM file."); return; }
    if (!jdgesLoaded) { setError("Upload the JDGEs database dump first."); return; }
    setProcessing(true);
    setError(null);

    setTimeout(() => {
      try {
        const bomPns = new Set(bomFiles.map((bf) => normalizePN(bf.partNumber)));
        const newResults: AssemblyResult[] = bomFiles.map((bf) => {
          const allParts: ProcessedPart[] = bf.rows.map((row) => ({
            ...row,
            isNew: !jdgesPartNumbers.has(normalizePN(row.partNo)),
            hasOwnBom: bomPns.has(normalizePN(row.partNo)),
          }));
          return {
            assemblyPn: bf.partNumber,
            newParts: allParts.filter((p) => p.isNew),
            allParts,
            bomRows: bf.rows,
          };
        });
        setResults(newResults);
        setExpandedAssemblies(new Set(newResults.map((r) => r.assemblyPn)));
      } catch (e) {
        setError("Processing error: " + String(e));
      } finally {
        setProcessing(false);
      }
    }, 50);
  }, [bomFiles, jdgesLoaded, jdgesPartNumbers]);

  const totalNew = results.reduce((s, r) => s + r.newParts.length, 0);
  const totalParts = results.reduce((s, r) => s + r.bomRows.length, 0);

  return (
    <div className="min-h-screen bg-background font-[Inter,system-ui,sans-serif]">
      <header className="border-b-2 border-foreground bg-foreground text-primary-foreground">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
          <div>
            <h1 className="text-lg font-black tracking-tight uppercase">JDGEs BOM Loader</h1>
            <p className="text-xs text-primary-foreground/60 font-mono mt-0.5">Bill of Materials cross-reference &amp; file generator</p>
          </div>
          <div className="text-xs font-mono text-primary-foreground/50 text-right hidden sm:block">
            <div>Upload BOMs + JDGEs dump → Generate</div>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-8">
        <div className="grid grid-cols-1 lg:grid-cols-[1fr_300px] gap-8">

          <div className="space-y-6">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {/* BOM files */}
              <div>
                <div className="flex items-center gap-2 mb-2">
                  <span className="bg-foreground text-primary-foreground text-xs font-mono font-bold px-1.5 py-0.5">01</span>
                  <h2 className="text-sm font-bold uppercase tracking-wider">BOM Spreadsheets</h2>
                </div>
                <p className="text-xs text-muted-foreground mb-3">
                  Name each file after its part number — e.g. <span className="font-mono">RW9350.xlsx</span>. Upload top-level and all sub-assembly BOMs.
                </p>
                <DropZone label="Drop BOM files here" hint=".xlsx / .xls / .csv — filename = part number"
                  accept=".xlsx,.xls,.csv" multiple onFiles={handleBomFiles} fileCount={bomFiles.length} />
                {bomFiles.length > 0 && (
                  <div className="mt-2 flex flex-wrap gap-1.5">
                    {bomFiles.map((bf) => (
                      <FileTag key={bf.id} name={bf.fileName}
                        onRemove={() => { setBomFiles((p) => p.filter((f) => f.id !== bf.id)); setResults([]); }} />
                    ))}
                  </div>
                )}
              </div>

              {/* JDGEs DB */}
              <div>
                <div className="flex items-center gap-2 mb-2">
                  <span className="bg-foreground text-primary-foreground text-xs font-mono font-bold px-1.5 py-0.5">02</span>
                  <h2 className="text-sm font-bold uppercase tracking-wider">JDGEs Database</h2>
                </div>
                <p className="text-xs text-muted-foreground mb-3">
                  Export all part numbers from JDGEs. All values in the file are scanned for matching part numbers.
                </p>
                <DropZone label="Drop JDGEs database dump" hint=".xlsx / .xls / .csv export from JDGEs"
                  accept=".xlsx,.xls,.csv" multiple={false} onFiles={handleJdgesFile} fileCount={jdgesFile ? 1 : 0} />
                {jdgesFile && (
                  <div className="mt-2 flex items-center gap-2">
                    <FileTag name={jdgesFile.name}
                      onRemove={() => { setJdgesFile(null); setJdgesLoaded(false); setJdgesPartNumbers(new Set()); setJdgesCount(0); setResults([]); }} />
                    <span className="text-xs font-mono text-green-700 font-bold shrink-0">{jdgesCount.toLocaleString()} PNs</span>
                  </div>
                )}
              </div>
            </div>

            {error && (
              <div className="flex items-center gap-2 bg-destructive/10 border border-destructive text-destructive px-4 py-2 text-sm font-mono">
                <AlertTriangle className="w-4 h-4 shrink-0" />{error}
              </div>
            )}

            <div className="flex items-center gap-4 flex-wrap">
              <button onClick={handleProcess}
                disabled={processing || bomFiles.length === 0 || !jdgesLoaded}
                className="bg-foreground text-primary-foreground font-bold uppercase tracking-wider text-sm px-6 py-2.5
                  hover:bg-accent hover:text-accent-foreground transition-colors disabled:opacity-40 disabled:cursor-not-allowed">
                {processing ? "Processing…" : "03 — Cross-Check"}
              </button>
              {results.length > 0 && (
                <button onClick={() => generateOutputExcel(results, bomFiles)}
                  className="flex items-center gap-2 bg-accent text-accent-foreground font-bold uppercase tracking-wider text-sm px-6 py-2.5
                    hover:bg-foreground hover:text-primary-foreground transition-colors">
                  <Download className="w-4 h-4" />04 — Generate Excel
                </button>
              )}
            </div>

            {results.length > 0 && (
              <div>
                <div className="flex items-center gap-3 mb-3 flex-wrap">
                  <div className="flex items-center gap-2">
                    <h2 className="text-sm font-bold uppercase tracking-wider">Results</h2>
                  </div>
                  <span className="text-xs font-mono">
                    <b>{totalParts}</b> parts · <b>{results.length}</b> assemblies
                    {totalNew > 0
                      ? <span className="text-accent font-bold ml-2">▲ {totalNew} new to load</span>
                      : <span className="text-green-700 font-bold ml-2">✓ all in JDGEs</span>}
                  </span>
                </div>
                <div className="border-t-2 border-foreground">
                  {results.map((result) => (
                    <ResultsSection key={result.assemblyPn} result={result}
                      expanded={expandedAssemblies.has(result.assemblyPn)}
                      onToggle={() => setExpandedAssemblies((prev) => {
                        const next = new Set(prev);
                        next.has(result.assemblyPn) ? next.delete(result.assemblyPn) : next.add(result.assemblyPn);
                        return next;
                      })} />
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Right panel */}
          <div className="space-y-4">
            <div className="border border-border p-4">
              <h3 className="text-xs font-bold uppercase tracking-wider mb-3 flex items-center gap-2">
                <Database className="w-3.5 h-3.5" /> How it works
              </h3>
              <ol className="space-y-3 text-xs text-muted-foreground">
                <li className="flex gap-2"><span className="font-mono font-bold text-foreground shrink-0">1.</span>Upload BOM files named after their part number. Both top-level and sub-assembly BOMs.</li>
                <li className="flex gap-2"><span className="font-mono font-bold text-foreground shrink-0">2.</span>Upload the JDGEs dump. Parts found here are marked "in JDGEs."</li>
                <li className="flex gap-2"><span className="font-mono font-bold text-foreground shrink-0">3.</span>Cross-check flags unlisted parts as <span className="text-accent font-semibold">NEW</span>. Sub-assemblies with their own BOM file get the <span className="bg-yellow-300 text-black px-0.5 text-[10px] font-bold">NEW Part / BOM</span> label.</li>
                <li className="flex gap-2"><span className="font-mono font-bold text-foreground shrink-0">4.</span>Generate Excel creates one sheet per assembly — new parts section on top, full BOM below — plus extra sheets for each sub-assembly.</li>
              </ol>
            </div>

            <div className="border border-border p-4">
              <h3 className="text-xs font-bold uppercase tracking-wider mb-2">Output columns</h3>
              <div className="space-y-2">
                <div>
                  <p className="text-[10px] font-bold uppercase text-muted-foreground mb-1">New Parts section</p>
                  <div className="flex flex-wrap gap-1">
                    {["Label", "Part No.", "Desc 1", "Desc 2", "Drawing Ref", "P or M", "Cost", "Supplier", "Lead-Time", "Mst.Plan.Family", "Comms Class", "Sub Class", "Unit", "Branch"].map((c) => (
                      <span key={c} className="text-[10px] font-mono bg-yellow-100 border border-yellow-300 px-1 py-0.5">{c}</span>
                    ))}
                  </div>
                </div>
                <div>
                  <p className="text-[10px] font-bold uppercase text-muted-foreground mb-1">BOM section</p>
                  <div className="flex flex-wrap gap-1">
                    {["Part No.", "Desc 1", "Desc 2", "Drawing Ref", "Qty", "Update Cost", "Supplier", "Lead-Time", "Mst.Plan.Family", "Comms Class", "Sub Class", "Unit", "Branch"].map((c) => (
                      <span key={c} className="text-[10px] font-mono bg-secondary border border-border px-1 py-0.5">{c}</span>
                    ))}
                  </div>
                </div>
              </div>
            </div>

            {results.length > 0 && (
              <div className="border border-border p-4">
                <h3 className="text-xs font-bold uppercase tracking-wider mb-3">Summary</h3>
                <div className="space-y-1.5">
                  {results.map((r) => (
                    <div key={r.assemblyPn} className="flex items-center justify-between text-xs font-mono">
                      <span className="font-semibold truncate max-w-[130px]">{r.assemblyPn}</span>
                      <div className="flex gap-2 items-center">
                        <span className="text-muted-foreground">{r.bomRows.length}</span>
                        {r.newParts.length > 0
                          ? <span className="text-accent font-bold">{r.newParts.length} new</span>
                          : <span className="text-green-700">✓</span>}
                      </div>
                    </div>
                  ))}
                  <div className="border-t border-border pt-1.5 flex items-center justify-between text-xs font-mono font-bold">
                    <span>TOTAL</span>
                    <span>{totalParts} · {totalNew > 0 ? <span className="text-accent">{totalNew} new</span> : "all loaded"}</span>
                  </div>
                </div>
              </div>
            )}

            {(bomFiles.length > 0 || jdgesLoaded) && (
              <button onClick={() => { setBomFiles([]); setJdgesFile(null); setJdgesLoaded(false); setJdgesPartNumbers(new Set()); setJdgesCount(0); setResults([]); setError(null); }}
                className="w-full flex items-center justify-center gap-2 border border-border py-2 text-xs font-mono text-muted-foreground hover:border-destructive hover:text-destructive transition-colors">
                <Trash2 className="w-3.5 h-3.5" />Reset all
              </button>
            )}
          </div>
        </div>
      </main>
    </div>
  );
}
