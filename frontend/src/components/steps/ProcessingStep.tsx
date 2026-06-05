import { useEffect, useState } from "react"
import { Loader2, CheckCircle2 } from "lucide-react"
import { generateFinalReports } from "../../lib/api"

interface Props {
  jobId: string
  onComplete: (outputs: any) => void
}

export default function ProcessingStep({ jobId, onComplete }: Props) {
  const [status, setStatus] = useState<"processing" | "po_jetro" | "pdfs" | "done">("processing")
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    const run = async () => {
      try {
        // We've already run Part 2 in RoutingStep. Now run Part 3 & 4.
        setStatus("po_jetro")
        const res = await generateFinalReports(jobId)
        setStatus("done")
        onComplete(res.outputs)
      } catch (err: any) {
        setError(err.message || "Final processing failed")
      }
    }
    run()
  }, [jobId, onComplete])

  return (
    <div className="flex flex-col items-center justify-center py-20 text-center h-full animate-in fade-in duration-700">
      <div className="relative mb-12">
        <div className="w-32 h-32 rounded-full border-4 border-muted flex items-center justify-center">
          <Loader2 className="w-12 h-12 text-primary animate-spin opacity-20" />
        </div>
        <div className="absolute inset-0 w-32 h-32 rounded-full border-4 border-primary border-t-transparent animate-spin-slow shadow-[0_0_15px_rgba(37,99,235,0.2)]"></div>
        <div className="absolute inset-4 w-24 h-24 rounded-full border-4 border-primary/30 border-b-transparent animate-spin reverse shadow-inner"></div>
      </div>
      
      <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-primary/10 text-primary text-[10px] font-bold uppercase tracking-wider mb-4">
        Step 4: Final Processing
      </div>
      <h2 className="text-3xl font-bold tracking-tight mb-3">Generating Final Reports</h2>
      <p className="text-muted-foreground max-w-md mx-auto mb-12 leading-relaxed font-medium">
        Building final pick sheets, sorting Jetro data, and rendering all print-ready PDFs.
      </p>
      
      {error ? (
        <div className="flex items-center gap-3 text-destructive bg-destructive/[0.03] border border-destructive/10 px-6 py-4 rounded-2xl animate-in slide-in-from-top-4 duration-500">
          <span className="font-bold tracking-tight">{error}</span>
        </div>
      ) : (
        <div className="w-full max-w-md space-y-4">
          <div className="flex items-center justify-between p-4 rounded-2xl bg-muted/30 border border-border/40 group transition-all duration-300">
            <div className="flex items-center gap-3">
              <div className="w-8 h-8 rounded-lg bg-green-500/10 flex items-center justify-center text-green-500">
                <CheckCircle2 className="w-4 h-4" />
              </div>
              <span className="text-sm font-bold text-foreground/80">Applying routing decisions</span>
            </div>
            <div className="text-[10px] font-black text-green-500 uppercase tracking-widest">Complete</div>
          </div>

          <div className={`flex items-center justify-between p-4 rounded-2xl border transition-all duration-500 ${
            status !== "processing" ? "bg-muted/30 border-border/40" : "bg-primary/[0.02] border-primary/20 shadow-lg shadow-primary/5"
          }`}>
            <div className="flex items-center gap-3">
              <div className={`w-8 h-8 rounded-lg flex items-center justify-center ${
                status === "done" ? "bg-green-500/10 text-green-500" : "bg-primary/10 text-primary"
              }`}>
                {status === "done" ? <CheckCircle2 className="w-4 h-4" /> : <Loader2 className="w-4 h-4 animate-spin" />}
              </div>
              <span className={`text-sm font-bold ${status !== "processing" ? "text-foreground/80" : "text-primary"}`}>
                Generating PO and Jetro sheets
              </span>
            </div>
            <div className={`text-[10px] font-black uppercase tracking-widest ${
              status === "done" ? "text-green-500" : "text-primary animate-pulse"
            }`}>
              {status === "done" ? "Complete" : "In Progress"}
            </div>
          </div>

          <div className={`flex items-center justify-between p-4 rounded-2xl border transition-all duration-500 ${
            status === "done" ? "bg-muted/30 border-border/40" : 
            status === "po_jetro" ? "bg-primary/[0.02] border-primary/20 shadow-lg shadow-primary/5" : "bg-muted/10 border-transparent opacity-40"
          }`}>
            <div className="flex items-center gap-3">
              <div className={`w-8 h-8 rounded-lg flex items-center justify-center ${
                status === "done" ? "bg-green-500/10 text-green-500" : "bg-primary/10 text-primary"
              }`}>
                {status === "done" ? <CheckCircle2 className="w-4 h-4" /> : <Loader2 className="w-4 h-4 animate-spin" />}
              </div>
              <span className={`text-sm font-bold ${status === "done" ? "text-foreground/80" : status === "po_jetro" ? "text-primary" : "text-muted-foreground"}`}>
                Building PDFs
              </span>
            </div>
            <div className={`text-[10px] font-black uppercase tracking-widest ${
              status === "done" ? "text-green-500" : status === "po_jetro" ? "text-primary animate-pulse" : "text-muted-foreground"
            }`}>
              {status === "done" ? "Complete" : status === "po_jetro" ? "In Progress" : "Pending"}
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
