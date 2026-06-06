import { useEffect, useState } from "react"
import { Loader2, X } from "lucide-react"

type RowKind = "band" | "header" | "produce" | "z-driver" | "data"

interface Row {
  kind: RowKind
  cells: (string | number)[]
}

interface SheetData {
  name: string
  rows: Row[]
  /** 1-based column indices that should render centered (Qty / Price / etc.) */
  centerCols?: number[]
  truncated: boolean
}

interface Props {
  jobId: string
  filename: string
  title: string
  onClose: () => void
}

// Excel-style column letters: A, B, ... Z, AA, AB, ...
function colLetter(index: number): string {
  let n = index
  let label = ""
  do {
    label = String.fromCharCode(65 + (n % 26)) + label
    n = Math.floor(n / 26) - 1
  } while (n >= 0)
  return label
}

export default function SheetViewer({ jobId, filename, title, onClose }: Props) {
  const [sheets, setSheets] = useState<SheetData[]>([])
  const [active, setActive] = useState(0)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    let cancelled = false
    setLoading(true)
    setError(null)
    fetch(`/api/preview/${jobId}/${filename}`)
      .then((r) => {
        if (!r.ok) throw new Error(`Failed to load preview (${r.status})`)
        return r.json()
      })
      .then((data) => {
        if (cancelled) return
        setSheets(data.sheets || [])
        setActive(0)
      })
      .catch((e) => !cancelled && setError(e.message))
      .finally(() => !cancelled && setLoading(false))
    return () => {
      cancelled = true
    }
  }, [jobId, filename])

  const sheet = sheets[active]
  const colCount = sheet
    ? sheet.rows.reduce((m, r) => Math.max(m, r.cells.length), 0)
    : 0

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm p-4 animate-in fade-in duration-200">
      <div className="bg-background rounded-2xl shadow-2xl border border-border/60 w-full max-w-[1300px] h-[85vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-border/60 bg-muted/30">
          <div>
            <h3 className="font-bold text-foreground tracking-tight">{title}</h3>
            <p className="text-[11px] font-mono text-muted-foreground">{filename} · read-only preview</p>
          </div>
          <button
            onClick={onClose}
            className="w-9 h-9 rounded-lg flex items-center justify-center hover:bg-muted text-muted-foreground hover:text-foreground transition-colors"
            aria-label="Close preview"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Sheet tabs */}
        {sheets.length > 0 && (
          <div className="flex items-center gap-1 px-4 pt-3 border-b border-border/40 bg-muted/10 overflow-x-auto">
            {sheets.map((s, i) => (
              <button
                key={s.name + i}
                onClick={() => setActive(i)}
                className={`px-4 py-2 text-[12px] font-bold rounded-t-lg whitespace-nowrap transition-colors ${
                  i === active
                    ? "bg-background text-primary border border-b-0 border-border/60"
                    : "text-muted-foreground hover:text-foreground hover:bg-muted/40"
                }`}
              >
                {s.name}
              </button>
            ))}
          </div>
        )}

        {/* Body */}
        <div className="flex-1 overflow-auto bg-muted/5">
          {loading && (
            <div className="h-full flex items-center justify-center text-muted-foreground gap-2">
              <Loader2 className="w-5 h-5 animate-spin" /> Loading preview…
            </div>
          )}
          {error && (
            <div className="h-full flex items-center justify-center text-red-600 text-sm font-medium">
              {error}
            </div>
          )}
          {!loading && !error && sheet && (
            <table className="border-collapse text-[12px] w-full">
              <thead className="sticky top-0 z-20">
                <tr>
                  <th className="sticky left-0 z-30 w-12 bg-muted border border-border/60 text-[10px] font-bold text-muted-foreground" />
                  {Array.from({ length: colCount }).map((_, c) => (
                    <th
                      key={c}
                      className="bg-muted border border-border/60 px-3 py-1 text-[10px] font-bold text-muted-foreground text-center min-w-[110px]"
                    >
                      {colLetter(c)}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sheet.rows.map((row, r) => {
                  const isBand = row.kind === "band"
                  const isHeader = row.kind === "header"
                  const isProduce = row.kind === "produce"
                  const isZDriver = row.kind === "z-driver"
                  const centerSet = new Set(sheet.centerCols || [])
                  const rowCls = isBand
                    ? "bg-[#1F2937] text-white font-bold"
                    : isHeader
                    ? "bg-[#4472C4] text-white font-bold"
                    : isProduce
                    ? "bg-[#D3D3D3]"
                    : isZDriver
                    ? "bg-[#B5BFC9]"
                    : "hover:bg-muted/30"
                  // Band cells: centered per spec §10 (vendor/driver/customer + date).
                  const bandCellCls = "px-3 py-1.5 align-middle whitespace-nowrap overflow-hidden text-ellipsis text-center"
                  const headerCellBase = "border border-white/20 px-3 py-1 align-middle whitespace-nowrap overflow-hidden text-ellipsis text-center"
                  const dataCellBase = "border border-border/30 px-3 py-1 align-top whitespace-nowrap overflow-hidden text-ellipsis max-w-[320px]"
                  // Band rows render as a single spanning cell so they read like the Excel merge.
                  if (isBand) {
                    const label = row.cells.find((v) => v !== "" && v != null) ?? ""
                    return (
                      <tr key={r} className={rowCls}>
                        <td className="sticky left-0 z-10 w-12 bg-muted/60 border border-border/40 text-[10px] font-mono text-muted-foreground/70 text-center">
                          {r + 1}
                        </td>
                        <td colSpan={colCount} className={bandCellCls} title={String(label)}>
                          {String(label)}
                        </td>
                      </tr>
                    )
                  }
                  return (
                    <tr key={r} className={rowCls}>
                      <td className="sticky left-0 z-10 w-12 bg-muted/60 border border-border/40 text-[10px] font-mono text-muted-foreground/70 text-center">
                        {r + 1}
                      </td>
                      {Array.from({ length: colCount }).map((_, c) => {
                        const isCentered = centerSet.has(c + 1)
                        const base = isHeader ? headerCellBase : dataCellBase
                        const align = isHeader ? "" : (isCentered ? " text-center" : " text-left")
                        return (
                          <td key={c} className={base + align} title={String(row.cells[c] ?? "")}>
                            {String(row.cells[c] ?? "")}
                          </td>
                        )
                      })}
                    </tr>
                  )
                })}
              </tbody>
            </table>
          )}
        </div>

        {/* Footer */}
        {sheet?.truncated && (
          <div className="px-6 py-2 border-t border-border/40 bg-amber-50 text-amber-700 text-[11px] font-medium">
            Preview limited to the first 2000 rows. Download the Excel file for the complete data.
          </div>
        )}
      </div>
    </div>
  )
}
