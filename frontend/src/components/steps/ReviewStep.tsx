import { useState } from "react"
import type { PreviewData } from "../../App"
import { Card, CardHeader, CardTitle, CardContent } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { Badge } from "@/components/ui/badge"
import { FileSpreadsheet, ArrowRight, Eye, AlertTriangle, CheckCircle2 } from "lucide-react"
import SheetViewer from "../SheetViewer"

interface Props {
  jobId: string
  previewData: PreviewData
  onNext: () => void
}

export default function ReviewStep({ jobId, previewData, onNext }: Props) {
  const [viewing, setViewing] = useState(false)

  return (
    <div className="flex flex-col h-full animate-in fade-in duration-700">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-10">
        <div>
          <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-primary/10 text-primary text-[10px] font-bold uppercase tracking-wider mb-2">
            Step 2: Validation
          </div>
          <h2 className="text-3xl font-bold tracking-tight">Review Extraction</h2>
          <p className="text-muted-foreground mt-1">Verify the compiled data and shortages before routing.</p>
        </div>
        <Button
          variant="outline"
          className="gap-2 rounded-xl h-12 px-6 font-bold hover:bg-primary hover:text-primary-foreground transition-all duration-300"
          onClick={() => setViewing(true)}
        >
          <Eye className="w-4 h-4" />
          View Full Workbook
        </Button>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-10">
        <Card className="border-border/40 shadow-sm hover:shadow-md transition-shadow duration-300 rounded-2xl overflow-hidden">
          <CardHeader className="pb-4 bg-muted/30">
            <CardTitle className="text-sm font-bold uppercase tracking-widest flex items-center gap-2 text-muted-foreground">
              <FileSpreadsheet className="w-4 h-4 text-primary" />
              Dataset Summary
            </CardTitle>
          </CardHeader>
          <CardContent className="pt-6">
            <div className="flex items-baseline gap-2">
              <div className="text-5xl font-black tracking-tighter text-primary">{previewData.allOrders.rowCount}</div>
              <div className="text-sm font-bold text-muted-foreground uppercase tracking-wider">Total Rows</div>
            </div>
            <p className="text-xs text-muted-foreground mt-3 font-medium">
              Compiled across item list, inventory, and shopping history.
            </p>
          </CardContent>
        </Card>

        <Card className="border-border/40 shadow-sm hover:shadow-md transition-shadow duration-300 rounded-2xl overflow-hidden">
          <CardHeader className="pb-4 bg-muted/30">
            <CardTitle className="text-sm font-bold uppercase tracking-widest flex items-center gap-2 text-muted-foreground">
              <FileSpreadsheet className="w-4 h-4 text-primary" />
              Automatic Extraction
            </CardTitle>
          </CardHeader>
          <CardContent className="pt-6 flex flex-col justify-between h-[calc(100%-80px)]">
            <div className="space-y-6">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-[10px] font-bold text-muted-foreground uppercase tracking-widest mb-1">Jetro Branch</p>
                  <p className="text-2xl font-black tracking-tight">{previewData.jetroSource.rowCount} <span className="text-xs font-bold text-muted-foreground">Items</span></p>
                </div>
                <div className="w-10 h-10 rounded-xl bg-orange-500/10 flex items-center justify-center text-orange-500 font-bold">J</div>
              </div>
              <div className="w-full h-1.5 bg-muted rounded-full overflow-hidden">
                <div 
                  className="h-full bg-orange-500 rounded-full" 
                  style={{ width: `${(previewData.jetroSource.rowCount / previewData.allOrders.rowCount) * 100}%` }} 
                />
              </div>
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-[10px] font-bold text-muted-foreground uppercase tracking-widest mb-1">Purchase Orders</p>
                  <p className="text-2xl font-black tracking-tight">{previewData.po.rowCount} <span className="text-xs font-bold text-muted-foreground">Items</span></p>
                </div>
                <div className="w-10 h-10 rounded-xl bg-blue-500/10 flex items-center justify-center text-blue-500 font-bold">P</div>
              </div>
              <div className="w-full h-1.5 bg-muted rounded-full overflow-hidden">
                <div 
                  className="h-full bg-blue-500 rounded-full" 
                  style={{ width: `${(previewData.po.rowCount / previewData.allOrders.rowCount) * 100}%` }} 
                />
              </div>
            </div>
          </CardContent>
        </Card>
      </div>

      <Card className="mb-10 border-destructive/20 shadow-lg shadow-destructive/5 rounded-2xl overflow-hidden">
        <div className="bg-destructive/[0.03] px-8 py-5 border-b border-destructive/10 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-destructive/10 flex items-center justify-center text-destructive">
              <AlertTriangle className="w-5 h-5" />
            </div>
            <div>
              <h3 className="font-bold text-lg tracking-tight text-destructive">Warehouse Shortages</h3>
              <p className="text-xs text-destructive/70 font-medium">Items that require immediate routing decisions</p>
            </div>
          </div>
          <Badge variant="destructive" className="px-3 py-1 rounded-lg font-bold shadow-sm">
            {previewData.warehouseShort.rowCount} Short
          </Badge>
        </div>
        <CardContent className="p-0">
          {previewData.warehouseShort.shortages.length > 0 ? (
            <div className="max-h-64 overflow-y-auto">
              <table className="w-full text-sm">
                <thead className="bg-muted/50 text-[10px] font-bold uppercase tracking-widest text-muted-foreground sticky top-0 z-10">
                  <tr>
                    <th className="px-8 py-3 text-left">Item Code</th>
                    <th className="px-8 py-3 text-left">Description</th>
                    <th className="px-8 py-3 text-right">Shortage Qty</th>
                  </tr>
                </thead>
                <tbody className="divide-y border-t">
                  {previewData.warehouseShort.shortages.map((s, i) => (
                    <tr key={i} className="group hover:bg-muted/30 transition-colors">
                      <td className="px-8 py-4 font-bold text-primary">{s.code}</td>
                      <td className="px-8 py-4 font-medium text-foreground/80">{s.productName}</td>
                      <td className="px-8 py-4 text-right">
                        <span className="inline-flex items-center px-2.5 py-1 rounded-md bg-destructive/10 text-destructive font-bold text-xs">
                          {s.shortage} {s.unit}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="py-16 text-center">
              <div className="w-16 h-16 rounded-full bg-green-500/10 flex items-center justify-center text-green-500 mx-auto mb-4">
                <CheckCircle2 className="w-8 h-8" />
              </div>
              <p className="font-bold text-foreground">Zero Shortages Detected</p>
              <p className="text-sm text-muted-foreground mt-1">Inventory levels match all current orders.</p>
            </div>
          )}
        </CardContent>
      </Card>

      <div className="mt-auto flex justify-end pt-8 border-t border-border/50">
        <Button size="lg" className="gap-2 h-14 px-10 rounded-2xl font-bold shadow-xl shadow-primary/20 hover:scale-[1.02] active:scale-[0.98] transition-all" onClick={onNext}>
          Configure Routing
          <ArrowRight className="w-5 h-5" />
        </Button>
      </div>

      {viewing && (
        <SheetViewer
          jobId={jobId}
          filename="part1.xlsx"
          title="Part 1 — Compiled Workbook"
          onClose={() => setViewing(false)}
        />
      )}
    </div>
  )
}
