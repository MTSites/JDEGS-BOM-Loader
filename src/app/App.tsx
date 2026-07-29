import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { Upload, FileSpreadsheet, Database, Download, X, AlertTriangle, CheckCircle, ChevronDown, ChevronRight, Trash2, Plus } from "lucide-react";

// ─── Types ───────────────────────────────────────────────────────────────────

interface BomRow {
  partNo: string;
  desc1: string;
  desc2: string;
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

interface LabourCode {
  code: string; // "" | "MT080" | "MT100" | "MT060"
  qty: string;
}

interface ProcessedPart extends BomRow {
  isNew: boolean;
  hasOwnBom: boolean;
}

interface AssemblyResult {
  assemblyPn: string;
  assemblyIsNew: boolean; // the assembly PN itself not in JDGEs
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

  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(10, raw.length); i++) {
    const nonEmpty = (raw[i] as (string | number)[]).filter((c) => String(c).trim() !== "").length;
    if (nonEmpty >= 2) { headerRowIdx = i; break; }
  }

  const headers = (raw[headerRowIdx] as (string | number)[]).map((h) => String(h));

  const colPN   = findCol(headers, ["partno", "partnum", "part", "pn", "item", "number", "code"]);
  const colD1   = findCol(headers, ["description1", "desc1", "description"]);
  const colD2   = findCol(headers, ["description2", "desc2"]);
  const colQty  = findCol(headers, ["qty", "quantity", "qnty"]);
  const colPoM  = findCol(headers, ["porm", "pm"]);
  const colCost = findCol(headers, ["updatecost", "cost", "price"]);
  const colSup  = findCol(headers, ["supplier", "vendor"]);
  const colLT   = findCol(headers, ["leadtime", "lead", "weeks", "lt"]);
  const colMPF  = findCol(headers, ["masterplanning", "masterpla", "planningfamily", "mpf"]);
  const colCC   = findCol(headers, ["commsclass", "comms"]);
  const colSC   = findCol(headers, ["subclass", "sub"]);
  const colUnit = findCol(headers, ["unit", "uom"]);
  const colBr   = findCol(headers, ["branch", "br"]);

  const get    = (row: (string | number)[], idx: number): string => idx >= 0 ? String(row[idx] ?? "").trim() : "";
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
      desc1: get(r, colD1),
      desc2: get(r, colD2),
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
    reader.onload = (e) => { resolve(XLSX.read(e.target?.result, { type: "array" })); };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function normalizePN(pn: string): string {
  return pn.trim().toUpperCase().replace(/\s+/g, "");
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

// Drawing Ref always mirrors Part No. per spec
function rowToNewPartsArr(p: ProcessedPart): (string | number)[] {
  return [
    p.hasOwnBom ? "NEW BOM" : "NEW Part",
    p.partNo, p.desc1, p.desc2, p.partNo, // drawingRef = partNo
    p.pOrM, p.cost, p.supplier, p.leadTime,
    p.masterPlanningFamily, p.commsClass, p.subClass, p.unit, p.branch,
  ];
}

function rowToBomArr(r: BomRow): (string | number)[] {
  return [
    r.partNo, r.desc1, r.desc2, r.partNo, // drawingRef = partNo
    r.qty, r.cost, r.supplier, r.leadTime,
    r.masterPlanningFamily, r.commsClass, r.subClass, r.unit, r.branch,
  ];
}

// Column indices (1-based) in the new-parts section
const COL_POM = 6;
const COL_MPF = 10;
const COL_CC  = 11;
const COL_SC  = 12;

async function generateOutputExcel(
  results: AssemblyResult[],
  bomFiles: BomFile[],
  labourCodes: Record<string, LabourCode>,
  outputFileName: string,
): Promise<void> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("JDGEs Load");

  ws.columns = [
    { width: 16 }, { width: 14 }, { width: 32 }, { width: 22 }, { width: 14 },
    { width: 8  }, { width: 10 }, { width: 26 }, { width: 10 },
    { width: 18 }, { width: 10 }, { width: 12 }, { width: 6  }, { width: 8  },
  ];

  const headerFill: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF262626" } };
  const newBomheaderFill: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF473B" } };
  const newBomFill: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF473B" } };
  const newFill:    ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC2A68F" } };
  const labourFill: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8E5DE" } };
  const boldWhite: Partial<ExcelJS.Font> = { bold: true, size: 10, color: { argb: "FFFFFFFF" } };
  const specialboldWhite: Partial<ExcelJS.Font> = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
  const boldDark:  Partial<ExcelJS.Font> = { bold: true, size: 10 };
  const baseFont:  Partial<ExcelJS.Font> = { size: 10 };
  

  const newPartDataRows: number[] = [];

  // ── Section 1: all new parts (assembly PNs + their contents), deduplicated ──
  const seenNew = new Set<string>();
  const allNewParts: ProcessedPart[] = [];

  for (const result of results) {
    // Check the assembly PN itself first
    if (result.assemblyIsNew) {
      const key = normalizePN(result.assemblyPn);
      if (!seenNew.has(key)) {
        seenNew.add(key);
        allNewParts.push({
          partNo: result.assemblyPn,
          desc1: "", desc2: "", qty: "", pOrM: "", cost: "",
          supplier: "", leadTime: "", masterPlanningFamily: "",
          commsClass: "", subClass: "", unit: "", branch: "",
          isNew: true,
          hasOwnBom: true,
        });
      }
    }
    // Then parts inside the BOM
    for (const p of result.newParts) {
      const key = normalizePN(p.partNo);
      if (!seenNew.has(key)) { seenNew.add(key); allNewParts.push(p); }
    }
  }

  const hdrRow = ws.addRow(NEW_PARTS_HEADERS);
  hdrRow.font = boldWhite;
  hdrRow.fill = headerFill;
  hdrRow.alignment = { wrapText: true, vertical: "middle" };

  if (allNewParts.length === 0) {
    const noRow = ws.addRow(["(all parts already in JDGEs)"]);
    noRow.font = baseFont;
  } else {
    for (const p of allNewParts) {
      const r = ws.addRow(rowToNewPartsArr(p));
      r.font = baseFont;
      r.getCell(1).fill = p.hasOwnBom ? newBomFill : newFill;
      r.getCell(1).font = boldWhite;
      newPartDataRows.push(r.number);
    }
  }

  ws.addRow([]);
  ws.addRow([]);

// ── Section 2: BOM blocks sorted ascending by PN ─────────────────────────

const bomMap = new Map<string, { rows: BomRow[]; fileId: string }>();

for (const result of results) {
  const bf = bomFiles.find(
    (f) => normalizePN(f.partNumber) === normalizePN(result.assemblyPn)
  );

  bomMap.set(normalizePN(result.assemblyPn), {
    rows: result.bomRows,
    fileId: bf?.id ?? "",
  });
}

for (const bf of bomFiles) {
  const key = normalizePN(bf.partNumber);

  if (!bomMap.has(key)) {
    bomMap.set(key, {
      rows: bf.rows,
      fileId: bf.id,
    });
  }
}

const sortedPns = [...bomMap.keys()].sort((a, b) =>
  a.localeCompare(b, undefined, {
    numeric: true,
    sensitivity: "base",
  })
);

for (const pnKey of sortedPns) {
  const { rows, fileId } = bomMap.get(pnKey)!;

  const originalPn =
    results.find((r) => normalizePN(r.assemblyPn) === pnKey)?.assemblyPn ??
    bomFiles.find((bf) => normalizePN(bf.partNumber) === pnKey)?.partNumber ??
    pnKey;

  // ── Assembly PN + BOM Header Row ─────────────────────────

  const bomHdrRow = ws.addRow([
    originalPn,
    ...BOM_HEADERS,
  ]);

  // Style BOM headers (columns B onward)
  bomHdrRow.font = boldWhite;
  bomHdrRow.fill = headerFill;

  bomHdrRow.alignment = {
    wrapText: true,
    vertical: "middle",
  };

  // Override column A styling for Assembly PN
  bomHdrRow.getCell(1).fill = newBomheaderFill;
  bomHdrRow.getCell(1).font = specialboldWhite;

  bomHdrRow.getCell(1).alignment = {
    horizontal: "center",
    vertical: "middle",
  };

  // ── Labour Row ───────────────────────────────────────────

  const lc = labourCodes[fileId];

  if (lc?.code) {
    const labourDesc: Record<string, string> = {
      MT060: "MACHINE SHOP",
      MT080: "BOILERMAKING",
      MT100: "ELECTRICAL",
    };

    const labourRow = ws.addRow([
      "",
      lc.code,
      labourDesc[lc.code] ?? "",
      "",
      lc.code,
      lc.qty || "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
    ]);

    labourRow.font = boldDark;
    labourRow.fill = labourFill;
  }

  // ── BOM Data Rows ────────────────────────────────────────

  for (const r of rows) {
    const dataRow = ws.addRow([
      "",
      ...rowToBomArr(r),
    ]);

    dataRow.font = baseFont;
  }

  // Blank row between BOMs
  ws.addRow([]);
}
  
  // ── Dropdowns on new-parts rows ───────────────────────────────────────────
  const dv = (rowNum: number, col: number, formula: string) => {
    ws.getCell(rowNum, col).dataValidation = {
      type: "list", allowBlank: true, formulae: [formula],
      showErrorMessage: true, errorStyle: "warning",
    };
  };
  for (const rowNum of newPartDataRows) {
    dv(rowNum, COL_POM, '"P,M"');
    dv(rowNum, COL_MPF, '"ASS"');
    dv(rowNum, COL_CC,  '"MSC"');
    dv(rowNum, COL_SC,  '"REQ,EHT,ELO,KEL,FLF"');
  }

  // ── Write file ────────────────────────────────────────────────────────────
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const safeName = (outputFileName.trim() || "JDGEs_BOM_Load").replace(/\.xlsx$/i, "");
  a.download = `${safeName}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
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
      className={`relative cursor-pointer border-2 border-dashed p-5 flex flex-col items-center gap-2 select-none transition-colors
        ${dragging ? "border-[#ff473b] bg-[#ff473b]/5" : "border-border hover:border-foreground/40"}`}
    >
      <input ref={inputRef} type="file" accept={accept} multiple={multiple} className="hidden"
        onChange={(e) => { onFiles(Array.from(e.target.files ?? [])); e.target.value = ""; }} />
      <Upload className="w-5 h-5 text-muted-foreground" />
      <div className="text-center">
        <p className="font-semibold text-sm">{label}</p>
        {hint && <p className="text-xs text-muted-foreground mt-0.5">{hint}</p>}
      </div>
      {fileCount > 0 && (
        <span className="absolute top-2 right-2 bg-[#ff473b] text-white text-xs font-mono px-1.5 py-0.5">
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
      <span className="truncate max-w-[200px]">{name}</span>
      <button onClick={(e) => { e.stopPropagation(); onRemove(); }}
        className="ml-auto text-muted-foreground hover:text-[#ff473b] transition-colors shrink-0">
        <X className="w-3 h-3" />
      </button>
    </div>
  );
}

// ─── Labour Code Panel ────────────────────────────────────────────────────────

const LABOUR_OPTIONS = ["", "MT080", "MT100", "MT060"] as const;

function LabourCodePanel({ bomFiles, labourCodes, onChange }: {
  bomFiles: BomFile[];
  labourCodes: Record<string, LabourCode>;
  onChange: (fileId: string, lc: LabourCode) => void;
}) {
  if (bomFiles.length === 0) return null;
  return (
    <div className="border border-border">
      <div className="px-3 py-2 bg-secondary border-b border-border flex items-center gap-2">
        <Plus className="w-3.5 h-3.5 text-muted-foreground" />
        <span className="text-xs font-bold uppercase tracking-wider">Labour Code per BOM</span>
        <span className="text-xs text-muted-foreground ml-1">(optional — adds a row at the top of each BOM)</span>
      </div>
      <div className="divide-y divide-border">
        {bomFiles.map((bf) => {
          const lc = labourCodes[bf.id] ?? { code: "", qty: "" };
          return (
            <div key={bf.id} className="flex items-center gap-3 px-3 py-2">
              <span className="font-mono text-xs font-semibold w-[130px] shrink-0 truncate">{bf.partNumber}</span>
              <select
                value={lc.code}
                onChange={(e) => onChange(bf.id, { ...lc, code: e.target.value })}
                className="text-xs font-mono border border-border bg-background px-2 py-1 focus:outline-none focus:border-[#ff473b]"
              >
                {LABOUR_OPTIONS.map((o) => (
                  <option key={o} value={o}>{o || "— none —"}</option>
                ))}
              </select>
              {lc.code && (
                <div className="flex items-center gap-1.5">
                  <span className="text-xs text-muted-foreground">Qty</span>
                  <input
                    type="text"
                    value={lc.qty}
                    onChange={(e) => onChange(bf.id, { ...lc, qty: e.target.value })}
                    placeholder="e.g. 30"
                    className="text-xs font-mono border border-border bg-background px-2 py-1 w-16 focus:outline-none focus:border-[#ff473b]"
                  />
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── Results Table ────────────────────────────────────────────────────────────

const NEW_PARTS_COLS = ["Part No.", "Description 1", "Description 2", "Drawing Ref", "P or M", "Cost", "Supplier", "Lead-Time", "Mst.Plan.Family", "Comms Class", "Sub Class", "Unit", "Branch"];
const BOM_COLS       = ["Part No.", "Description 1", "Description 2", "Drawing Ref", "Qty", "Update Cost", "Supplier", "Lead-Time", "Mst.Plan.Family", "Comms Class", "Sub Class", "Unit", "Branch"];

function partToRow(p: ProcessedPart): (string | number)[] {
  return [p.partNo, p.desc1, p.desc2, p.partNo, p.pOrM, p.cost, p.supplier, p.leadTime, p.masterPlanningFamily, p.commsClass, p.subClass, p.unit, p.branch];
}
function bomToRow(r: BomRow): (string | number)[] {
  return [r.partNo, r.desc1, r.desc2, r.partNo, r.qty, r.cost, r.supplier, r.leadTime, r.masterPlanningFamily, r.commsClass, r.subClass, r.unit, r.branch];
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
        <div className="flex items-center gap-3 min-w-0">
          {expanded ? <ChevronDown className="w-4 h-4 shrink-0" /> : <ChevronRight className="w-4 h-4 shrink-0" />}
          <span className="font-mono font-bold text-sm">{result.assemblyPn}</span>
          <span className="text-xs text-muted-foreground shrink-0">{result.bomRows.length} lines</span>
        </div>
        <div className="flex items-center gap-3 shrink-0 ml-2">
          {result.newParts.length > 0 || result.assemblyIsNew ? (
            <span className="flex items-center gap-1 text-xs font-mono text-[#ff473b] font-bold">
              <AlertTriangle className="w-3.5 h-3.5" />
              {result.newParts.length + (result.assemblyIsNew ? 1 : 0)} new
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
                  ${tab === t ? "bg-[#262626] text-white" : "text-muted-foreground hover:text-foreground"}`}>
                {t === "new"
                  ? `New Parts (${result.newParts.length + (result.assemblyIsNew ? 1 : 0)})`
                  : `Full BOM (${result.bomRows.length})`}
              </button>
            ))}
          </div>

          {tab === "new" && (
            <DataTable
              cols={NEW_PARTS_COLS}
              rows={[
                ...(result.assemblyIsNew ? [{
                  partNo: result.assemblyPn, desc1: "", desc2: "", qty: "", pOrM: "",
                  cost: "", supplier: "", leadTime: "", masterPlanningFamily: "",
                  commsClass: "", subClass: "", unit: "", branch: "",
                  isNew: true, hasOwnBom: true,
                } as ProcessedPart] : []),
                ...result.newParts,
              ].map((p) => partToRow(p))}
              renderLabel={(i) => {
                const parts = [
                  ...(result.assemblyIsNew ? [{ hasOwnBom: true }] : []),
                  ...result.newParts,
                ];
                const p = parts[i];
                return p.hasOwnBom
                  ? <span className="bg-[#ff473b] text-white px-1 py-0.5 font-bold text-[10px] whitespace-nowrap">NEW BOM</span>
                  : <span className="bg-[#c2a68f] text-white px-1 py-0.5 font-bold text-[10px] whitespace-nowrap">NEW Part</span>;
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
  const [labourCodes, setLabourCodes] = useState<Record<string, LabourCode>>({});
  const [outputFileName, setOutputFileName] = useState("JDGEs_BOM_Load");

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
          // Check the assembly PN itself
          const assemblyIsNew = !jdgesPartNumbers.has(normalizePN(bf.partNumber));

          const allParts: ProcessedPart[] = bf.rows.map((row) => ({
            ...row,
            isNew: !jdgesPartNumbers.has(normalizePN(row.partNo)),
            hasOwnBom: bomPns.has(normalizePN(row.partNo)),
          }));

          return {
            assemblyPn: bf.partNumber,
            assemblyIsNew,
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

  const handleReset = useCallback(() => {
    setBomFiles([]); setJdgesFile(null); setJdgesLoaded(false);
    setJdgesPartNumbers(new Set()); setJdgesCount(0);
    setResults([]); setError(null); setLabourCodes({});
    setOutputFileName("JDGEs_BOM_Load");
  }, []);

  const totalNew = results.reduce((s, r) => s + r.newParts.length + (r.assemblyIsNew ? 1 : 0), 0);
  const totalParts = results.reduce((s, r) => s + r.bomRows.length, 0);

  return (
    <div className="min-h-screen bg-background font-[Inter,system-ui,sans-serif]">
      <header className="border-b-2 border-[#262626] bg-[#262626] text-white">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <span className="inline-block w-2.5 h-2.5 rounded-full bg-[#ff473b]" />
            <div>
              <h1 className="text-lg font-black tracking-tight uppercase">JDGEs BOM Loader</h1>
              <p className="text-xs text-white/50 font-mono mt-0.5">Bill of Materials cross-reference &amp; file generator</p>
            </div>
          </div>
          <div className="text-xs font-mono text-white/40 text-right hidden sm:block">
            Upload BOMs + JDGEs dump → Generate
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-8">
        <div className="grid grid-cols-1 lg:grid-cols-[1fr_300px] gap-8">

          <div className="space-y-5">

            {/* Steps 01 + 02 */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <div className="flex items-center gap-2 mb-2">
                  <span className="bg-[#ff473b] text-white text-xs font-mono font-bold px-1.5 py-0.5">01</span>
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

              <div>
                <div className="flex items-center gap-2 mb-2">
                  <span className="bg-[#ff473b] text-white text-xs font-mono font-bold px-1.5 py-0.5">02</span>
                  <h2 className="text-sm font-bold uppercase tracking-wider">JDGEs Database</h2>
                </div>
                <p className="text-xs text-muted-foreground mb-3">
                  Export all part numbers from JDGEs. The entire file is scanned — every cell is checked against your BOMs.
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

            {/* Labour codes */}
            {bomFiles.length > 0 && (
              <LabourCodePanel
                bomFiles={bomFiles}
                labourCodes={labourCodes}
                onChange={(id, lc) => setLabourCodes((prev) => ({ ...prev, [id]: lc }))}
              />
            )}

            {error && (
              <div className="flex items-center gap-2 bg-destructive/10 border border-[#ff473b] text-[#ff473b] px-4 py-2 text-sm font-mono">
                <AlertTriangle className="w-4 h-4 shrink-0" />{error}
              </div>
            )}

            {/* Actions row */}
            <div className="flex items-center gap-3 flex-wrap">
              <button onClick={handleProcess}
                disabled={processing || bomFiles.length === 0 || !jdgesLoaded}
                className="bg-[#262626] text-white font-bold uppercase tracking-wider text-sm px-6 py-2.5
                  hover:bg-[#ff473b] transition-colors disabled:opacity-40 disabled:cursor-not-allowed">
                {processing ? "Processing…" : "03 — Cross-Check"}
              </button>

              {results.length > 0 && (
                <div className="flex items-center gap-2">
                  <div className="flex items-center border border-border bg-background">
                    <span className="text-xs font-mono text-muted-foreground px-2 border-r border-border">filename</span>
                    <input
                      type="text"
                      value={outputFileName}
                      onChange={(e) => setOutputFileName(e.target.value)}
                      className="text-xs font-mono px-2 py-2.5 w-44 focus:outline-none bg-transparent"
                      placeholder="JDGEs_BOM_Load"
                    />
                    <span className="text-xs font-mono text-muted-foreground px-2 border-l border-border">.xlsx</span>
                  </div>
                  <button onClick={() => generateOutputExcel(results, bomFiles, labourCodes, outputFileName)}
                    className="flex items-center gap-2 bg-[#ff473b] text-white font-bold uppercase tracking-wider text-sm px-5 py-2.5
                      hover:bg-[#262626] transition-colors">
                    <Download className="w-4 h-4" />04 — Generate
                  </button>
                </div>
              )}
            </div>

            {/* Results */}
            {results.length > 0 && (
              <div>
                <div className="flex items-center gap-3 mb-3 flex-wrap">
                  <h2 className="text-sm font-bold uppercase tracking-wider">Results</h2>
                  <span className="text-xs font-mono">
                    <b>{totalParts}</b> parts · <b>{results.length}</b> assemblies
                    {totalNew > 0
                      ? <span className="text-[#ff473b] font-bold ml-2">▲ {totalNew} new to load</span>
                      : <span className="text-green-700 font-bold ml-2">✓ all in JDGEs</span>}
                  </span>
                </div>
                <div className="border-t-2 border-[#262626]">
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

            {/* How it works */}
            <div className="border border-border p-4">
              <h3 className="text-xs font-bold uppercase tracking-wider mb-3 flex items-center gap-2">
                <Database className="w-3.5 h-3.5" /> How it works
              </h3>
              <ol className="space-y-3 text-xs text-muted-foreground">
                <li className="flex gap-2">
                  <span className="font-mono font-bold text-[#ff473b] shrink-0">01</span>
                  <span><strong className="text-foreground">Upload your BOM files.</strong> Rename each spreadsheet to match its part number exactly (e.g. <span className="font-mono">RW9350.xlsx</span>). Upload both top-level assemblies and any sub-assembly BOMs — you can drop multiple files at once.</span>
                </li>
                <li className="flex gap-2">
                  <span className="font-mono font-bold text-[#ff473b] shrink-0">02</span>
                  <span><strong className="text-foreground">Upload your JDGEs dump.</strong> Export all existing part numbers from JDGEs as a spreadsheet. Every cell in the file is scanned — no specific column format required.</span>
                </li>
                <li className="flex gap-2">
                  <span className="font-mono font-bold text-[#ff473b] shrink-0">03</span>
                  <span><strong className="text-foreground">Cross-check.</strong> Each assembly PN and every part inside its BOM is compared against JDGEs. Parts not found are flagged: <span className="bg-[#ff473b] text-white px-1 font-bold text-[10px]">NEW BOM</span> for sub-assemblies with their own BOM file, <span className="bg-[#c2a68f] text-white px-1 font-bold text-[10px]">NEW Part</span> for regular new parts.</span>
                </li>
                <li className="flex gap-2">
                  <span className="font-mono font-bold text-[#ff473b] shrink-0">04</span>
                  <span><strong className="text-foreground">Set labour codes (optional).</strong> For each BOM you can choose a labour code (MT080, MT100, MT060) and quantity — this adds a row at the top of that BOM section in the output file.</span>
                </li>
                <li className="flex gap-2">
                  <span className="font-mono font-bold text-[#ff473b] shrink-0">05</span>
                  <span><strong className="text-foreground">Generate Excel.</strong> A single-sheet file is produced: all new parts at the top, then each BOM in ascending part number order. Type the filename before downloading. Dropdown validation is pre-set for P or M, Master Planning Family, Comms Class, and Sub Class.</span>
                </li>
              </ol>
            </div>

            {/* Output columns */}
            <div className="border border-border p-4">
              <h3 className="text-xs font-bold uppercase tracking-wider mb-2">Output columns</h3>
              <div className="space-y-2">
                <div>
                  <p className="text-[10px] font-bold uppercase text-muted-foreground mb-1">New Parts section</p>
                  <div className="flex flex-wrap gap-1">
                    {["Label", "Part No.", "Desc 1", "Desc 2", "Drawing Ref", "P or M ▾", "Cost", "Supplier", "Lead-Time", "Mst.Plan.Family ▾", "Comms Class ▾", "Sub Class ▾", "Unit", "Branch"].map((c) => (
                      <span key={c} className={`text-[10px] font-mono px-1 py-0.5 ${c.endsWith("▾") ? "bg-[#ff473b]/15 border border-[#ff473b]/40 text-[#ff473b] font-bold" : "bg-[#c2a68f]/20 border border-[#c2a68f]/50"}`}>{c}</span>
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
                <p className="text-[10px] text-muted-foreground">Drawing Ref is always set to match Part No.</p>
              </div>
            </div>

            {/* Summary */}
            {results.length > 0 && (
              <div className="border border-border p-4">
                <h3 className="text-xs font-bold uppercase tracking-wider mb-3">Summary</h3>
                <div className="space-y-1.5">
                  {results.map((r) => {
                    const newCount = r.newParts.length + (r.assemblyIsNew ? 1 : 0);
                    return (
                      <div key={r.assemblyPn} className="grid grid-cols-[1fr_auto_auto] gap-2 items-center text-xs font-mono">
                        <span className="font-semibold">{r.assemblyPn}</span>
                        <span className="text-muted-foreground text-right">{r.bomRows.length}</span>
                        {newCount > 0
                          ? <span className="text-[#ff473b] font-bold text-right w-14">{newCount} new</span>
                          : <span className="text-green-700 text-right w-14">✓</span>}
                      </div>
                    );
                  })}
                  <div className="border-t border-border pt-1.5 grid grid-cols-[1fr_auto_auto] gap-2 items-center text-xs font-mono font-bold">
                    <span>TOTAL</span>
                    <span className="text-right">{totalParts}</span>
                    <span className={`text-right w-14 ${totalNew > 0 ? "text-[#ff473b]" : "text-green-700"}`}>
                      {totalNew > 0 ? `${totalNew} new` : "✓"}
                    </span>
                  </div>
                </div>
              </div>
            )}

            {(bomFiles.length > 0 || jdgesLoaded) && (
              <button onClick={handleReset}
                className="w-full flex items-center justify-center gap-2 border border-border py-2 text-xs font-mono text-muted-foreground hover:border-[#ff473b] hover:text-[#ff473b] transition-colors">
                <Trash2 className="w-3.5 h-3.5" />Reset all
              </button>
            )}
          </div>
        </div>
      </main>
    </div>
  );
}
