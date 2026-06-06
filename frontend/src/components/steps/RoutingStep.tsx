import { useState, useMemo } from "react"
import type { PreviewData } from "../../App"
import { Button } from "@/components/ui/button"
import { applyRouting } from "../../lib/api"
import { Loader2, ArrowRight, GripVertical, FileSpreadsheet, AlertTriangle, CheckCircle2, Search } from "lucide-react"
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
} from "@dnd-kit/core"
import type { DragEndEvent } from "@dnd-kit/core"
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  verticalListSortingStrategy,
  useSortable,
} from "@dnd-kit/sortable"
import { CSS } from "@dnd-kit/utilities"

interface Props {
  jobId: string
  previewData: PreviewData
  onNext: () => void
}

export default function RoutingStep({ jobId, previewData, onNext }: Props) {
  const [isApplying, setIsApplying] = useState(false)
  const [activeTab, setActiveTab] = useState<"all_orders" | "jetro_source" | "po" | "shortages" | "drivers">("all_orders")
  
  // Combine all routable rows for the decisions map
  const allRows = useMemo(() => {
    return [
      ...previewData.allOrders.rows,
      ...previewData.jetroSource.rows,
      ...previewData.po.rows,
    ]
  }, [previewData])

  const [sheetDecisions, setSheetDecisions] = useState<Record<string, { sheet: string; vendor: string }>>(() => {
    const initial: Record<string, { sheet: string; vendor: string }> = {}
    allRows.forEach(row => {
      initial[row.id] = { sheet: row.sheet, vendor: row.vendorRoute }
    })
    return initial
  })

  const [wsDecisions, setWsDecisions] = useState<Record<string, string>>(() => {
    const initial: Record<string, string> = {}
    previewData.warehouseShort.rows.forEach(row => {
      initial[row.id] = row.updateVendor || ""
    })
    return initial
  })

  const [drivers, setDrivers] = useState(previewData.drivers)
  const [filter, setFilter] = useState("")
  const [selectedRowId, setSelectedRowId] = useState<string | null>(null)

  // Per-cell edits. Keyed by row.id, then by field name. A field is only
  // present in the map if the user actually typed into it, so the original
  // row value is the fallback in `getVal`.
  const [cellEdits, setCellEdits] = useState<Record<string, Record<string, any>>>({})

  const getVal = (row: any, field: string): any => {
    const edited = cellEdits[row.id]
    if (edited && Object.prototype.hasOwnProperty.call(edited, field)) {
      return edited[field]
    }
    return row[field]
  }

  const setVal = (rowId: string, field: string, value: any) => {
    setCellEdits(prev => ({
      ...prev,
      [rowId]: { ...(prev[rowId] || {}), [field]: value },
    }))
  }

  const activeRows = useMemo(() => {
    switch(activeTab) {
      case "all_orders": return previewData.allOrders.rows
      case "jetro_source": return previewData.jetroSource.rows
      case "po": return previewData.po.rows
      default: return []
    }
  }, [activeTab, previewData])

  const filteredRows = useMemo(() => {
    if (!filter) return activeRows
    return activeRows.filter(r => 
      r.productName.toLowerCase().includes(filter.toLowerCase()) || 
      r.code.toLowerCase().includes(filter.toLowerCase())
    )
  }, [activeRows, filter])

  const selectedRow = useMemo(() => {
    return allRows.find(r => r.id === selectedRowId)
  }, [allRows, selectedRowId])

  // A warehouse short row is "unrouted" when its dropdown is still on the
  // empty placeholder. Block Execute Logic until every shortage has a route.
  const unroutedShortRows = useMemo(() => {
    return previewData.warehouseShort.rows.filter(
      row => !(wsDecisions[row.id] && wsDecisions[row.id].trim())
    )
  }, [previewData.warehouseShort.rows, wsDecisions])
  const hasUnroutedShorts = unroutedShortRows.length > 0

  const handleApply = async () => {
    if (hasUnroutedShorts) {
      setActiveTab("shortages")
      const firstId = unroutedShortRows[0]?.id
      if (firstId) setSelectedRowId(firstId)
      return
    }
    setIsApplying(true)
    try {
      const sheetRoutingPayload = Object.entries(sheetDecisions).map(([id, dec]) => ({
        id,
        ...dec
      }))
      const wsRoutingPayload = Object.entries(wsDecisions).map(([id, updateVendor]) => ({
        id,
        updateVendor
      }))
      // Flatten the nested edits map into [{id, field, value}, ...] for the API.
      const cellEditsPayload = Object.entries(cellEdits).flatMap(([id, fields]) =>
        Object.entries(fields).map(([field, value]) => ({ id, field, value }))
      )

      await applyRouting(jobId, sheetRoutingPayload, wsRoutingPayload, drivers, cellEditsPayload)
      onNext()
    } catch (err) {
      console.error("Failed to apply routing", err)
      setIsApplying(false)
    }
  }

  const sensors = useSensors(
    useSensor(PointerSensor),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  )

  const handleDragEnd = (event: DragEndEvent) => {
    const { active, over } = event
    if (over && active.id !== over.id) {
      setDrivers((items) => {
        const oldIndex = items.indexOf(active.id as string)
        const newIndex = items.indexOf(over.id as string)
        return arrayMove(items, oldIndex, newIndex)
      })
    }
  }

  return (
    <div className="flex flex-col h-full animate-in fade-in duration-700">
      <div className="mb-6">
        <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-primary/10 text-primary text-[10px] font-bold uppercase tracking-wider mb-2">
          Step 3: Data Orchestration
        </div>
        <h2 className="text-2xl font-bold tracking-tight">Routing Workspace</h2>
        <p className="text-xs text-muted-foreground mt-1">
          Excel-style grid for managing sheet assignments and driver sequences.
        </p>
      </div>

      {/* Excel-style Sheet Tabs */}
      <div className="flex items-center bg-[#dee1e6] px-1 h-9 gap-0.5 border-b border-border/60">
        <button 
          onClick={() => setActiveTab("all_orders")}
          className={`px-4 h-8 text-[11px] font-bold transition-all flex items-center gap-2 rounded-t-sm relative ${
            activeTab === "all_orders" 
              ? "bg-white text-primary shadow-[0_-2px_5px_rgba(0,0,0,0.05)] after:absolute after:bottom-[-2px] after:left-0 after:right-0 after:h-[3px] after:bg-white" 
              : "text-muted-foreground hover:bg-white/50"
          }`}
        >
          <FileSpreadsheet className="w-3 h-3" />
          All Orders
          {activeTab === "all_orders" && <div className="absolute bottom-0 left-0 right-0 h-0.5 bg-primary" />}
        </button>
        <button 
          onClick={() => setActiveTab("jetro_source")}
          className={`px-4 h-8 text-[11px] font-bold transition-all flex items-center gap-2 rounded-t-sm relative ${
            activeTab === "jetro_source" 
              ? "bg-white text-primary shadow-[0_-2px_5px_rgba(0,0,0,0.05)] after:absolute after:bottom-[-2px] after:left-0 after:right-0 after:h-[3px] after:bg-white" 
              : "text-muted-foreground hover:bg-white/50"
          }`}
        >
          <FileSpreadsheet className="w-3 h-3" />
          Jetro Source
          {activeTab === "jetro_source" && <div className="absolute bottom-0 left-0 right-0 h-0.5 bg-primary" />}
        </button>
        <button 
          onClick={() => setActiveTab("po")}
          className={`px-4 h-8 text-[11px] font-bold transition-all flex items-center gap-2 rounded-t-sm relative ${
            activeTab === "po" 
              ? "bg-white text-primary shadow-[0_-2px_5px_rgba(0,0,0,0.05)] after:absolute after:bottom-[-2px] after:left-0 after:right-0 after:h-[3px] after:bg-white" 
              : "text-muted-foreground hover:bg-white/50"
          }`}
        >
          <FileSpreadsheet className="w-3 h-3" />
          PO
          {activeTab === "po" && <div className="absolute bottom-0 left-0 right-0 h-0.5 bg-primary" />}
        </button>
        <button 
          onClick={() => setActiveTab("shortages")}
          className={`px-4 h-8 text-[11px] font-bold transition-all flex items-center gap-2 rounded-t-sm relative ${
            activeTab === "shortages" 
              ? "bg-white text-destructive shadow-[0_-2px_5px_rgba(0,0,0,0.05)] after:absolute after:bottom-[-2px] after:left-0 after:right-0 after:h-[3px] after:bg-white" 
              : "text-muted-foreground hover:bg-white/50"
          }`}
        >
          <AlertTriangle className="w-3 h-3" />
          Warehouse Shorts
          {previewData.warehouseShort.rows.length > 0 && (
            <span className="ml-1 px-1.5 py-0.5 rounded-full bg-destructive text-white text-[8px]">{previewData.warehouseShort.rows.length}</span>
          )}
          {activeTab === "shortages" && <div className="absolute bottom-0 left-0 right-0 h-0.5 bg-destructive" />}
        </button>
        <button 
          onClick={() => setActiveTab("drivers")}
          className={`px-4 h-8 text-[11px] font-bold transition-all flex items-center gap-2 rounded-t-sm relative ${
            activeTab === "drivers" 
              ? "bg-white text-primary shadow-[0_-2px_5px_rgba(0,0,0,0.05)] after:absolute after:bottom-[-2px] after:left-0 after:right-0 after:h-[3px] after:bg-white" 
              : "text-muted-foreground hover:bg-white/50"
          }`}
        >
          <GripVertical className="w-3 h-3" />
          Driver Sequence
          {activeTab === "drivers" && <div className="absolute bottom-0 left-0 right-0 h-0.5 bg-primary" />}
        </button>
        <div className="flex-1" />
        <div className="px-3 text-[10px] font-bold text-muted-foreground/40 uppercase tracking-widest">
          Workbook: Night_Shift_Execution
        </div>
      </div>

      {/* Grid Workspace */}
      <div className="flex-1 bg-[#f3f3f3] border-x border-b border-border/60 overflow-hidden relative min-h-[500px] flex flex-col">
        {/* Formula Bar / Status Bar */}
        <div className="h-8 bg-white border-b border-border/60 flex items-center px-2 gap-2 text-[11px]">
          <div className="flex items-center gap-1 bg-muted/30 px-2 py-0.5 rounded border border-border/40 font-mono text-muted-foreground w-16 justify-center">
            {selectedRowId ? allRows.findIndex(r => r.id === selectedRowId) + 1 : ""}
          </div>
          <div className="w-px h-4 bg-border/60 mx-1" />
          <div className="flex-1 flex items-center gap-2 overflow-hidden">
            <span className="font-serif italic text-primary/60 font-bold px-1 text-xs">fx</span>
            <div className="flex-1 truncate font-medium text-foreground/80">
              {selectedRow ? `${selectedRow.productName} [${selectedRow.code}]` : "Select a cell to view details"}
            </div>
          </div>
        </div>

        {(activeTab === "all_orders" || activeTab === "jetro_source" || activeTab === "po") && (
          <div className="flex-1 flex flex-col bg-white overflow-hidden">
            <div className="bg-[#f8f9fa] border-b border-border/60 p-1 flex items-center gap-2">
              <div className="relative flex-1 max-w-xs">
                <Search className="absolute left-2 top-1/2 -translate-y-1/2 w-3 h-3 text-muted-foreground" />
                <input 
                  type="text" 
                  placeholder={`Search ${activeTab.replace("_", " ")}...`} 
                  value={filter}
                  onChange={(e) => setFilter(e.target.value)}
                  className="w-full bg-white border border-border/60 rounded px-7 py-1 text-[11px] focus:ring-1 focus:ring-primary outline-none"
                />
              </div>
              <div className="text-[10px] font-bold text-muted-foreground/40 uppercase tracking-widest px-2">
                Showing {filteredRows.length} items in {activeTab.replace("_", " ")}
              </div>
            </div>
            <div className="flex-1 overflow-auto">
              <table className="w-full border-collapse table-fixed text-[12px]">
                <thead className="sticky top-0 z-30">
                  <tr className="h-6">
                    <th className="w-10 border-r border-b border-border bg-[#e1e3e8] text-[10px] text-muted-foreground font-medium p-0"></th>
                    <th className="w-48 border-r border-b border-border bg-[#e1e3e8] text-[10px] text-muted-foreground font-bold p-0 uppercase">A</th>
                    <th className="w-24 border-r border-b border-border bg-[#e1e3e8] text-[10px] text-muted-foreground font-bold p-0 uppercase">B</th>
                    <th className="w-24 border-r border-b border-border bg-[#e1e3e8] text-[10px] text-muted-foreground font-bold p-0 uppercase">C</th>
                    <th className="w-40 border-r border-b border-border bg-[#e1e3e8] text-[10px] text-muted-foreground font-bold p-0 uppercase">D</th>
                    <th className="w-44 border-b border-border bg-[#e1e3e8] text-[10px] text-muted-foreground font-bold p-0 uppercase">E</th>
                  </tr>
                  <tr className="bg-[#f8f9fa] border-b border-border shadow-sm">
                    <th className="border-r border-border p-1 bg-[#f8f9fa]"></th>
                    <th className="border-r border-border text-left px-2 py-1.5 font-bold text-[10px] text-muted-foreground uppercase tracking-tight">Product Description</th>
                    <th className="border-r border-border text-left px-2 py-1.5 font-bold text-[10px] text-muted-foreground uppercase tracking-tight">Item ID</th>
                    <th className="border-r border-border text-left px-2 py-1.5 font-bold text-[10px] text-muted-foreground uppercase tracking-tight">Location</th>
                    <th className="border-r border-border text-left px-2 py-1.5 font-bold text-[10px] text-muted-foreground uppercase tracking-tight">Routing Decision</th>
                    <th className="text-left px-2 py-1.5 font-bold text-[10px] text-muted-foreground uppercase tracking-tight">Logistics Provider</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.map((row, idx) => (
                    <tr 
                      key={row.id} 
                      onClick={() => setSelectedRowId(row.id)}
                      className={`h-8 border-b border-border/40 hover:bg-primary/[0.03] transition-colors cursor-default group ${
                        selectedRowId === row.id ? "bg-primary/[0.05] ring-1 ring-inset ring-primary/20" : ""
                      }`}
                    >
                      <td className="border-r border-border bg-[#f8f9fa] text-center text-[10px] text-muted-foreground/60 font-medium">{idx + 1}</td>
                      <td className="border-r border-border p-0">
                        <input
                          type="text"
                          value={getVal(row, "productName") ?? ""}
                          onChange={(e) => setVal(row.id, "productName", e.target.value)}
                          onClick={(e) => e.stopPropagation()}
                          title={String(getVal(row, "productName") ?? "")}
                          className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-primary rounded-none px-2 py-1 text-[12px] font-medium outline-none"
                        />
                      </td>
                      <td className="border-r border-border p-0">
                        <input
                          type="text"
                          value={getVal(row, "code") ?? ""}
                          onChange={(e) => setVal(row.id, "code", e.target.value)}
                          onClick={(e) => e.stopPropagation()}
                          className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-primary rounded-none px-2 py-1 font-mono text-[11px] text-muted-foreground/80 outline-none"
                        />
                      </td>
                      <td className="border-r border-border p-0">
                        <input
                          type="text"
                          value={getVal(row, "bin") ?? ""}
                          onChange={(e) => setVal(row.id, "bin", e.target.value)}
                          onClick={(e) => e.stopPropagation()}
                          className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-primary rounded-none px-2 py-1 italic text-muted-foreground/80 outline-none"
                        />
                      </td>
                      <td className={`border-r border-border p-0 relative ${selectedRowId === row.id ? "bg-white shadow-inner" : ""}`}>
                        <select 
                          value={sheetDecisions[row.id]?.sheet || "All Orders"} 
                          onChange={(e) => setSheetDecisions(prev => ({
                            ...prev,
                            [row.id]: { ...prev[row.id], sheet: e.target.value }
                          }))}
                          className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-primary rounded-none px-2 text-[11px] font-bold text-primary appearance-none cursor-pointer outline-none"
                        >
                          <option value="All Orders">All Orders</option>
                          <option value="Jetro source">Jetro source</option>
                          <option value="PO">PO</option>
                        </select>
                        <div className="absolute right-1 top-1/2 -translate-y-1/2 pointer-events-none opacity-20 group-hover:opacity-100 transition-opacity">
                          <ArrowRight className="w-2 h-2 rotate-90" />
                        </div>
                      </td>
                      <td className={`p-0 relative ${selectedRowId === row.id ? "bg-white shadow-inner" : ""}`}>
                        <select 
                          value={sheetDecisions[row.id]?.vendor || "Jetro"} 
                          onChange={(e) => setSheetDecisions(prev => ({
                            ...prev,
                            [row.id]: { ...prev[row.id], vendor: e.target.value }
                          }))}
                          className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-primary rounded-none px-2 text-[11px] font-bold appearance-none cursor-pointer outline-none"
                        >
                          <option value="Jetro">JETRO</option>
                          {previewData.vendors.map(v => (
                            <option key={v} value={v}>{v.toUpperCase()}</option>
                          ))}
                        </select>
                        <div className="absolute right-1 top-1/2 -translate-y-1/2 pointer-events-none opacity-20 group-hover:opacity-100 transition-opacity">
                          <ArrowRight className="w-2 h-2 rotate-90" />
                        </div>
                      </td>
                    </tr>
                  ))}
                  {filteredRows.length === 0 && (
                    <tr>
                      <td colSpan={6} className="h-32 text-center text-muted-foreground italic text-xs">
                        No rows found in this sheet.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {activeTab === "shortages" && (
          <div className="flex-1 flex flex-col bg-white overflow-hidden relative">
            {previewData.warehouseShort.rows.length > 0 ? (
              <div className="flex-1 overflow-auto">
                <table className="border-collapse text-[12px]" style={{ minWidth: "1500px" }}>
                  <colgroup>
                    <col style={{ width: "32px" }} />
                    <col style={{ width: "260px" }} />
                    <col style={{ width: "84px" }} />
                    <col style={{ width: "70px" }} />
                    <col style={{ width: "140px" }} />
                    <col style={{ width: "240px" }} />
                    <col style={{ width: "60px" }} />
                    <col style={{ width: "70px" }} />
                    <col style={{ width: "160px" }} />
                    <col style={{ width: "80px" }} />
                    <col style={{ width: "90px" }} />
                    <col style={{ width: "60px" }} />
                    <col style={{ width: "180px" }} />
                  </colgroup>
                  <thead className="sticky top-0 z-30">
                    <tr className="bg-[#fff5f5] border-b border-destructive/20 shadow-sm">
                      <th className="border-r border-border p-1 bg-[#fff5f5]"></th>
                      <th className="border-r border-border text-left px-2 py-1.5 font-bold text-[10px] text-destructive uppercase tracking-tight">Product Name / Code</th>
                      <th className="border-r border-border text-left px-2 py-1.5 font-bold text-[10px] text-destructive uppercase tracking-tight">Bin</th>
                      <th className="border-r border-border text-left px-2 py-1.5 font-bold text-[10px] text-destructive uppercase tracking-tight">Int. Bin</th>
                      <th className="border-r border-border text-left px-2 py-1.5 font-bold text-[10px] text-destructive uppercase tracking-tight">Vendor</th>
                      <th className="border-r border-border text-left px-2 py-1.5 font-bold text-[10px] text-destructive uppercase tracking-tight">Description</th>
                      <th className="border-r border-border text-center px-2 py-1.5 font-bold text-[10px] text-destructive uppercase tracking-tight">Qty</th>
                      <th className="border-r border-border text-center px-2 py-1.5 font-bold text-[10px] text-destructive uppercase tracking-tight">QOH</th>
                      <th className="border-r border-border text-left px-2 py-1.5 font-bold text-[10px] text-destructive uppercase tracking-tight">Customer</th>
                      <th className="border-r border-border text-left px-2 py-1.5 font-bold text-[10px] text-destructive uppercase tracking-tight">Driver</th>
                      <th className="border-r border-border text-right px-2 py-1.5 font-bold text-[10px] text-destructive uppercase tracking-tight">Shortage</th>
                      <th className="border-r border-border text-center px-2 py-1.5 font-bold text-[10px] text-destructive uppercase tracking-tight">Unit</th>
                      <th className="text-left px-2 py-1.5 font-bold text-[10px] text-destructive uppercase tracking-tight">Emergency Re-Route</th>
                    </tr>
                  </thead>
                  <tbody>
                    {previewData.warehouseShort.rows.map((row, idx) => (
                      <tr
                        key={row.id}
                        onClick={() => setSelectedRowId(row.id)}
                        className={`h-10 border-b border-border/40 hover:bg-destructive/[0.02] transition-colors cursor-default group ${
                          selectedRowId === row.id ? "bg-destructive/[0.04] ring-1 ring-inset ring-destructive/20" : ""
                        }`}
                      >
                        <td className="border-r border-border bg-[#fff5f5] text-center text-[10px] text-destructive/40 font-bold">{idx + 1}</td>
                        <td className="border-r border-border p-0">
                          <input
                            type="text"
                            value={getVal(row, "productName") ?? ""}
                            onChange={(e) => setVal(row.id, "productName", e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                            title={String(getVal(row, "productName") ?? "")}
                            className="w-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 pt-1 text-[12px] font-bold text-foreground outline-none"
                          />
                          <input
                            type="text"
                            value={getVal(row, "code") ?? ""}
                            onChange={(e) => setVal(row.id, "code", e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                            className="w-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 pb-1 text-[9px] text-muted-foreground/60 font-mono uppercase tracking-tighter outline-none"
                          />
                        </td>
                        <td className="border-r border-border p-0">
                          <input
                            type="text"
                            value={getVal(row, "bin") ?? ""}
                            onChange={(e) => setVal(row.id, "bin", e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                            className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 py-1 text-[12px] text-foreground/80 outline-none"
                          />
                        </td>
                        <td className="border-r border-border p-0">
                          <input
                            type="text"
                            value={getVal(row, "internalBin") ?? ""}
                            onChange={(e) => setVal(row.id, "internalBin", e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                            className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 py-1 font-mono text-[11px] text-foreground/70 outline-none"
                          />
                        </td>
                        <td className="border-r border-border p-0">
                          <input
                            type="text"
                            value={getVal(row, "vendor") ?? ""}
                            onChange={(e) => setVal(row.id, "vendor", e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                            className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 py-1 text-[12px] text-foreground/80 outline-none"
                          />
                        </td>
                        <td className="border-r border-border p-0">
                          <input
                            type="text"
                            value={getVal(row, "description") ?? ""}
                            onChange={(e) => setVal(row.id, "description", e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                            className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 py-1 text-[11px] text-muted-foreground/70 outline-none"
                          />
                        </td>
                        <td className="border-r border-border p-0">
                          <input
                            type="text"
                            inputMode="decimal"
                            value={getVal(row, "qty") ?? ""}
                            onChange={(e) => setVal(row.id, "qty", e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                            className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 py-1 text-center tabular-nums text-[12px] text-foreground/80 outline-none"
                          />
                        </td>
                        <td className="border-r border-border p-0">
                          <input
                            type="text"
                            inputMode="decimal"
                            value={getVal(row, "qoh") ?? ""}
                            onChange={(e) => setVal(row.id, "qoh", e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                            className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 py-1 text-center tabular-nums text-[12px] text-foreground/80 outline-none"
                          />
                        </td>
                        <td className="border-r border-border p-0">
                          <input
                            type="text"
                            value={getVal(row, "customer") ?? ""}
                            onChange={(e) => setVal(row.id, "customer", e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                            className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 py-1 text-[12px] text-foreground/80 outline-none"
                          />
                        </td>
                        <td className="border-r border-border p-0">
                          <input
                            type="text"
                            value={getVal(row, "driver") ?? ""}
                            onChange={(e) => setVal(row.id, "driver", e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                            className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 py-1 text-[12px] text-foreground/80 outline-none"
                          />
                        </td>
                        <td className="border-r border-border p-0 bg-destructive/[0.01]">
                          <input
                            type="text"
                            inputMode="decimal"
                            value={getVal(row, "shortage") ?? ""}
                            onChange={(e) => setVal(row.id, "shortage", e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                            className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 py-1 text-right font-black text-destructive tabular-nums text-[12px] outline-none"
                          />
                        </td>
                        <td className="border-r border-border p-0 bg-destructive/[0.01]">
                          <input
                            type="text"
                            value={getVal(row, "unit") ?? ""}
                            onChange={(e) => setVal(row.id, "unit", e.target.value)}
                            onClick={(e) => e.stopPropagation()}
                            className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 py-1 text-center font-bold text-[11px] text-destructive/80 uppercase tracking-wider outline-none"
                          />
                        </td>
                        <td className={`p-0 relative ${selectedRowId === row.id ? "bg-white shadow-inner" : ""}`}>
                          <select
                            value={wsDecisions[row.id] || ""}
                            onChange={(e) => setWsDecisions(prev => ({
                              ...prev,
                              [row.id]: e.target.value
                            }))}
                            className="w-full h-full bg-transparent border-none focus:ring-1 focus:ring-destructive rounded-none px-2 text-[11px] font-black text-destructive appearance-none cursor-pointer outline-none"
                          >
                            <option value="">-- SELECT ROUTE --</option>
                            <option value="Jetro">JETRO BRANCH</option>
                            <option value="WH">WAREHOUSE (WH)</option>
                            <optgroup label="BIN OVERRIDES">
                              <option value="COOLER">COOLER</option>
                              <option value="COOLER-PD">COOLER-PD</option>
                              <option value="DRY">DRY</option>
                              <option value="FREEZER">FREEZER</option>
                              <option value="PICK UP">PICK UP</option>
                            </optgroup>
                            <optgroup label="VENDOR DIRECT">
                              {previewData.vendors.map(v => (
                                <option key={v} value={v}>{v.toUpperCase()}</option>
                              ))}
                            </optgroup>
                          </select>
                          <div className="absolute right-1 top-1/2 -translate-y-1/2 pointer-events-none opacity-20 group-hover:opacity-100 transition-opacity">
                            <ArrowRight className="w-2 h-2 rotate-90 text-destructive" />
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center h-full gap-4 text-muted-foreground animate-in zoom-in duration-500">
                <CheckCircle2 className="w-12 h-12 text-green-500 opacity-20" />
                <p className="font-bold uppercase tracking-widest text-[10px]">No shortages detected in this shift</p>
              </div>
            )}
          </div>
        )}

        {activeTab === "drivers" && (
          <div className="absolute inset-0 bg-muted/10 p-8 flex justify-center overflow-auto">
            <div className="w-full max-w-md">
              <div className="mb-6 p-4 rounded-xl bg-primary/[0.03] border border-primary/10">
                <p className="text-xs font-bold text-primary uppercase tracking-widest mb-1">Driver Sequence Control</p>
                <p className="text-[10px] text-muted-foreground font-medium leading-relaxed">
                  The order below determines the pull sequence. Drivers in the first tier are prioritized for <span className="font-bold text-primary underline decoration-2 underline-offset-2">Freezer One</span>.
                </p>
              </div>

              <DndContext 
                sensors={sensors}
                collisionDetection={closestCenter}
                onDragEnd={handleDragEnd}
              >
                <SortableContext 
                  items={drivers}
                  strategy={verticalListSortingStrategy}
                >
                  <div className="space-y-2">
                    <div className="flex items-center gap-2 mb-4">
                      <div className="h-px flex-1 bg-border/60" />
                      <span className="text-[9px] font-black text-muted-foreground/40 uppercase tracking-[0.3em]">Tier 1: Freezer Priority</span>
                      <div className="h-px flex-1 bg-border/60" />
                    </div>
                    
                    {drivers.slice(0, 3).map((driver, idx) => (
                      <SortableItem key={driver} id={driver} index={idx + 1} isFreezerOne={true} />
                    ))}
                    
                    {drivers.length > 3 && (
                      <>
                        <div className="flex items-center gap-2 my-8">
                          <div className="h-px flex-1 bg-border/60" />
                          <span className="text-[9px] font-black text-muted-foreground/40 uppercase tracking-[0.3em]">Tier 2: Standard Flow</span>
                          <div className="h-px flex-1 bg-border/60" />
                        </div>
                        {drivers.slice(3).map((driver, idx) => (
                          <SortableItem key={driver} id={driver} index={idx + 4} isFreezerOne={false} />
                        ))}
                      </>
                    )}
                    
                    {drivers.length === 0 && (
                      <div className="text-center py-20 border-2 border-dashed rounded-3xl border-muted/40">
                        <p className="text-[10px] font-bold text-muted-foreground/40 uppercase tracking-widest text-center">No driver data available</p>
                      </div>
                    )}
                  </div>
                </SortableContext>
              </DndContext>
            </div>
          </div>
        )}
      </div>

      <div className={`sticky bottom-4 mt-6 flex justify-between items-center backdrop-blur-md p-4 rounded-2xl border shadow-2xl z-20 transition-colors ${
        hasUnroutedShorts
          ? "bg-destructive/[0.04] border-destructive/40 shadow-destructive/10"
          : "bg-card/95 border-border/60 shadow-primary/10"
      }`}>
        <div className="flex items-center gap-6 px-2">
          <button
            type="button"
            onClick={() => hasUnroutedShorts && setActiveTab("shortages")}
            disabled={!hasUnroutedShorts}
            className="flex flex-col text-left disabled:cursor-default"
            title={hasUnroutedShorts ? "Jump to Warehouse Shorts to finish routing" : ""}
          >
            <span className="text-[9px] font-black text-muted-foreground/40 uppercase tracking-widest">Workspace Status</span>
            <div className="flex items-center gap-1.5 mt-0.5">
              {hasUnroutedShorts ? (
                <>
                  <AlertTriangle className="w-3 h-3 text-destructive" />
                  <span className="text-[10px] font-bold uppercase tracking-wider text-destructive">
                    {unroutedShortRows.length} Shortage{unroutedShortRows.length === 1 ? "" : "s"} Unrouted
                  </span>
                </>
              ) : (
                <>
                  <div className="w-1.5 h-1.5 rounded-full bg-green-500 shadow-[0_0_8px_rgba(34,197,94,0.5)]" />
                  <span className="text-[10px] font-bold uppercase tracking-wider">Ready for Execution</span>
                </>
              )}
            </div>
          </button>
          <div className="h-8 w-px bg-border/60" />
          <div className="flex flex-col">
            <span className="text-[9px] font-black text-muted-foreground/40 uppercase tracking-widest">Decisions Logged</span>
            <span className="text-[10px] font-bold tabular-nums mt-0.5">{Object.keys(sheetDecisions).length + Object.keys(wsDecisions).length} Overrides</span>
          </div>
        </div>

        <Button
          size="lg"
          className="gap-2 h-12 px-8 rounded-xl font-bold shadow-xl shadow-primary/20 hover:scale-[1.02] active:scale-[0.98] transition-all disabled:hover:scale-100 disabled:opacity-60 disabled:cursor-not-allowed"
          onClick={handleApply}
          disabled={isApplying || hasUnroutedShorts}
          title={hasUnroutedShorts ? `Select a route for all ${unroutedShortRows.length} warehouse shortage row${unroutedShortRows.length === 1 ? "" : "s"} before running.` : ""}
        >
          {isApplying ? (
            <>
              <Loader2 className="w-4 h-4 animate-spin" />
              Processing Workbook...
            </>
          ) : hasUnroutedShorts ? (
            <>
              <AlertTriangle className="w-4 h-4" />
              Resolve Shortages First
            </>
          ) : (
            <>
              Execute Logic
              <ArrowRight className="w-4 h-4" />
            </>
          )}
        </Button>
      </div>
    </div>
  )
}

function SortableItem({ id, index, isFreezerOne }: { id: string; index: number; isFreezerOne: boolean }) {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id })

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    zIndex: isDragging ? 10 : 0,
    opacity: isDragging ? 0.6 : 1,
  }

  return (
    <div 
      ref={setNodeRef} 
      style={style} 
      className={`flex items-center gap-4 p-4 rounded-xl border transition-all duration-300 group ${
        isDragging 
          ? "shadow-2xl border-primary bg-background ring-4 ring-primary/5" 
          : "bg-card border-border/40 hover:border-primary/30 hover:shadow-md"
      }`}
    >
      <div 
        {...attributes} 
        {...listeners} 
        className="cursor-grab active:cursor-grabbing text-muted-foreground/40 hover:text-primary transition-colors p-1 -ml-1"
      >
        <GripVertical className="w-4 h-4" />
      </div>
      <div className={`w-7 h-7 rounded-lg flex items-center justify-center text-[10px] font-black transition-colors ${
        isFreezerOne ? "bg-primary text-primary-foreground shadow-lg shadow-primary/20" : "bg-muted text-muted-foreground border border-border"
      }`}>
        {index}
      </div>
      <span className="font-bold text-sm tracking-tight text-foreground/80 group-hover:text-primary transition-colors">{id}</span>
    </div>
  )
}
