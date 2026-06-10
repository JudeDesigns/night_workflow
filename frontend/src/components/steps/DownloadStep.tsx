import { useState } from "react"
import { Card, CardHeader, CardTitle, CardContent } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { FileDown, RefreshCcw, FileSpreadsheet, FileText, Eye } from "lucide-react"
import SheetViewer from "../SheetViewer"

interface Props {
  jobId: string
  outputs: any
  onReset: () => void
}

export default function DownloadStep({ jobId, outputs, onReset }: Props) {
  const baseUrl = `/api/files/${jobId}`
  const [viewing, setViewing] = useState<{ filename: string; title: string } | null>(null)

  // PDFs is a list so the Pick Sheets card can surface 3 separate downloads
  // (Dry / Freezer / WH Pickup) while every other card has at most one PDF.
  // downloadOnly suppresses the View button for files that are download-only.
  type Report = {
    title: string
    description: string
    xlsx: string | null | undefined
    pdf?: string | null
    pdfs?: { label: string; file: string | null | undefined }[]
    color: string
    downloadOnly?: boolean
  }

  const reports: Report[] = [
    {
      title: "PO Report",
      description: "Vendor purchase orders",
      xlsx: outputs?.poReport?.xlsx,
      pdf: outputs?.poReport?.pdf,
      color: "bg-blue-500",
    },
    {
      title: "Warehouse Pick Sheets",
      description: "Dry, Freezer, & WH Pickup",
      xlsx: outputs?.dryFreezerWh?.xlsx,
      pdfs: [
        { label: "Dry",      file: outputs?.dryFreezerWh?.dryPdf },
        { label: "Freezer",  file: outputs?.dryFreezerWh?.freezerPdf },
        { label: "WH Pickup", file: outputs?.dryFreezerWh?.whPickupPdf },
      ],
      color: "bg-green-500",
    },
    {
      title: "Jetro Pack",
      description: "Menu, & Produce workbook",
      xlsx: outputs?.jetroWorkbook?.xlsx,
      pdf: null,
      color: "bg-orange-500",
    },
    {
      title: "Jetro PDF Report",
      description: "Landscape printable report",
      xlsx: null,
      pdf: outputs?.jetroPdf?.pdf,
      color: "bg-orange-600",
    },
    {
      title: "All Orders (Routed)",
      description: "All Orders sheet with routing decisions applied",
      xlsx: outputs?.allOrdersRouted?.xlsx,
      color: "bg-purple-500",
      downloadOnly: true,
    },
    {
      title: "Jetro Source (Routed)",
      description: "Jetro Source sheet with routing decisions applied",
      xlsx: outputs?.jetroSourceRouted?.xlsx,
      color: "bg-purple-400",
      downloadOnly: true,
    },
    {
      title: "PO Source (Routed)",
      description: "PO sheet with routing decisions applied",
      xlsx: outputs?.poRouted?.xlsx,
      color: "bg-indigo-500",
      downloadOnly: true,
    },
    {
      title: "WH Shortage (Routed)",
      description: "Warehouse Shortage sheet with routing decisions applied",
      xlsx: outputs?.whShortageRouted?.xlsx,
      color: "bg-purple-600",
      downloadOnly: true,
    },
    {
      title: "Original Upload",
      description: "The source file submitted at the start of this shift",
      xlsx: outputs?.originalUpload?.xlsx,
      color: "bg-slate-400",
      downloadOnly: true,
    },
  ]

  return (
    <div className="flex flex-col h-full animate-in fade-in duration-700">
      <div className="mb-12 text-center">
        <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-green-500/10 text-green-600 text-[10px] font-bold uppercase tracking-wider mb-6">
          Step 5: Completion
        </div>
        <div className="w-20 h-20 rounded-2xl bg-green-500/10 text-green-500 flex items-center justify-center mx-auto mb-6 shadow-xl shadow-green-500/5 rotate-3">
          <FileDown className="w-10 h-10" />
        </div>
        <h2 className="text-3xl font-bold tracking-tight mb-3">Automation Complete</h2>
        <p className="text-muted-foreground max-w-md mx-auto leading-relaxed font-medium">
          All night shift reports have been processed, sorted, and rendered. Your workspace is ready.
        </p>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 mb-12">
        {reports.map((report, idx) => (
          <Card key={idx} className="border-border/40 shadow-sm hover:shadow-xl hover:border-primary/20 transition-all duration-500 rounded-2xl overflow-hidden group">
            <CardHeader className="pb-4 bg-muted/30 border-b border-border/40">
              <div className="flex items-start justify-between">
                <div>
                  <CardTitle className="text-base font-bold tracking-tight text-foreground group-hover:text-primary transition-colors">{report.title}</CardTitle>
                  <p className="text-[10px] font-bold text-muted-foreground uppercase tracking-widest mt-1">{report.description}</p>
                </div>
                <div className={`w-2 h-2 rounded-full ${report.color} animate-pulse`} />
              </div>
            </CardHeader>
            <CardContent className="pt-6 flex flex-wrap gap-2">
              {report.xlsx && (
                <>
                  {!report.downloadOnly && (
                    <Button
                      variant="outline"
                      size="sm"
                      className="flex-1 min-w-[88px] gap-2 h-11 font-bold rounded-xl border-border/60 hover:bg-primary/5 hover:text-primary hover:border-primary/20 transition-all"
                      onClick={() => setViewing({ filename: report.xlsx!, title: report.title })}
                    >
                      <Eye className="w-4 h-4 text-primary" />
                      View
                    </Button>
                  )}
                  <Button
                    variant="outline"
                    size="sm"
                    className="flex-1 min-w-[88px] gap-2 h-11 font-bold rounded-xl border-border/60 hover:bg-green-50 hover:text-green-700 hover:border-green-200 transition-all"
                    onClick={() => window.open(`${baseUrl}/${report.xlsx}`, "_blank")}
                  >
                    <FileSpreadsheet className="w-4 h-4 text-green-600" />
                    Excel
                  </Button>
                </>
              )}
              {report.pdf && (
                <Button
                  variant="outline"
                  size="sm"
                  className="flex-1 min-w-[88px] gap-2 h-11 font-bold rounded-xl border-border/60 hover:bg-red-50 hover:text-red-700 hover:border-red-200 transition-all"
                  onClick={() => window.open(`${baseUrl}/${report.pdf}`, "_blank")}
                >
                  <FileText className="w-4 h-4 text-red-600" />
                  PDF
                </Button>
              )}
              {report.pdfs?.filter(p => p.file).map(p => (
                <Button
                  key={p.label}
                  variant="outline"
                  size="sm"
                  className="flex-1 min-w-[88px] gap-2 h-11 font-bold rounded-xl border-border/60 hover:bg-red-50 hover:text-red-700 hover:border-red-200 transition-all"
                  onClick={() => window.open(`${baseUrl}/${p.file}`, "_blank")}
                  title={`Download the ${p.label} PDF`}
                >
                  <FileText className="w-4 h-4 text-red-600" />
                  {p.label}
                </Button>
              ))}
            </CardContent>
          </Card>
        ))}
      </div>

      <div className="mt-auto bg-primary/[0.03] border border-primary/10 rounded-2xl p-8 flex flex-col sm:flex-row items-center justify-between gap-6 shadow-inner">
        <div className="flex items-center gap-4">
          <div className="w-12 h-12 rounded-xl bg-primary/10 flex items-center justify-center text-primary shadow-sm">
            <FileDown className="w-6 h-6" />
          </div>
          <div>
            <h4 className="font-bold text-foreground">Archive Bundle</h4>
            <p className="text-xs text-muted-foreground font-medium">Download all reports in a single .zip file</p>
          </div>
        </div>
        <div className="flex items-center gap-3 w-full sm:w-auto">
          <Button 
            variant="ghost" 
            onClick={() => window.open(`${baseUrl}/all_outputs.zip`, "_blank")} 
            className="gap-2 h-12 px-6 font-bold text-primary hover:bg-primary/10 rounded-xl flex-1 sm:flex-none"
          >
            Download All (.zip)
          </Button>
          <Button 
            variant="default" 
            onClick={onReset} 
            className="gap-2 h-12 px-8 font-bold rounded-xl shadow-xl shadow-primary/20 hover:scale-[1.02] active:scale-[0.98] transition-all flex-1 sm:flex-none"
          >
            <RefreshCcw className="w-4 h-4" />
            New Shift
          </Button>
        </div>
      </div>

      {viewing && (
        <SheetViewer
          jobId={jobId}
          filename={viewing.filename}
          title={viewing.title}
          onClose={() => setViewing(null)}
        />
      )}
    </div>
  )
}
