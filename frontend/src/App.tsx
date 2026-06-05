import { useEffect, useState } from "react"
import UploadStep from "./components/steps/UploadStep"
import ReviewStep from "./components/steps/ReviewStep"
import RoutingStep from "./components/steps/RoutingStep"
import ProcessingStep from "./components/steps/ProcessingStep"
import DownloadStep from "./components/steps/DownloadStep"
import { restoreJob } from "./lib/api"
import { Loader2, RotateCcw } from "lucide-react"

const SESSION_KEY = "nightshift_session_v1"

type Session = {
  jobId: string
  step: number
  finalOutputs: any | null
}

export type RowData = {
  id: string
  values: any[]
  productName: string
  code: string
  bin: string
  vendor: string
  sheet: string
  vendorRoute: string
  updateVendor?: string
  shortage?: number
  unit?: string
}

export type PreviewData = {
  allOrders: { rowCount: number; rows: RowData[] }
  jetroSource: { rowCount: number; rows: RowData[] }
  po: { rowCount: number; rows: RowData[] }
  warehouseShort: {
    rowCount: number
    shortages: { code: string; productName: string; unit: string; shortage: number }[]
    rows: RowData[]
  }
  drivers: string[]
  vendors: string[]
}

function App() {
  const [step, setStep] = useState(1)
  const [jobId, setJobId] = useState<string | null>(null)
  const [previewData, setPreviewData] = useState<PreviewData | null>(null)
  const [finalOutputs, setFinalOutputs] = useState<any | null>(null)
  const [restoring, setRestoring] = useState(true)

  const steps = [
    { id: 1, title: "Upload" },
    { id: 2, title: "Review" },
    { id: 3, title: "Routing" },
    { id: 4, title: "Process" },
    { id: 5, title: "Download" },
  ]

  // Rehydrate from localStorage on mount.
  useEffect(() => {
    const raw = localStorage.getItem(SESSION_KEY)
    if (!raw) {
      setRestoring(false)
      return
    }
    let saved: Session | null = null
    try {
      saved = JSON.parse(raw) as Session
    } catch {
      localStorage.removeItem(SESSION_KEY)
      setRestoring(false)
      return
    }
    if (!saved?.jobId) {
      setRestoring(false)
      return
    }
    restoreJob(saved.jobId)
      .then((data) => {
        setJobId(data.jobId)
        setPreviewData(data.preview)
        setFinalOutputs(data.outputs && Object.keys(data.outputs).length ? data.outputs : saved!.finalOutputs)
        // Cap the restored step at what the backend says is reachable.
        const maxStep = data.status === "complete" ? 5 : data.status === "part2_complete" ? 4 : 3
        setStep(Math.min(saved!.step || 2, maxStep))
      })
      .catch(() => {
        localStorage.removeItem(SESSION_KEY)
      })
      .finally(() => setRestoring(false))
  }, [])

  // Persist on changes (skip while we're still restoring to avoid clobbering).
  useEffect(() => {
    if (restoring) return
    if (!jobId) {
      localStorage.removeItem(SESSION_KEY)
      return
    }
    const session: Session = { jobId, step, finalOutputs }
    localStorage.setItem(SESSION_KEY, JSON.stringify(session))
  }, [jobId, step, finalOutputs, restoring])

  const handleReset = () => {
    setStep(1)
    setJobId(null)
    setPreviewData(null)
    setFinalOutputs(null)
    localStorage.removeItem(SESSION_KEY)
  }

  const handleCancelSession = () => {
    if (window.confirm("Cancel this session and start over? Any unsaved routing choices will be lost.")) {
      handleReset()
    }
  }

  if (restoring) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-background text-muted-foreground gap-2">
        <Loader2 className="w-5 h-5 animate-spin" />
        <span className="text-sm font-medium">Restoring session…</span>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-background font-sans text-foreground selection:bg-primary/10">
      <header className="border-b bg-card/50 backdrop-blur-md sticky top-0 z-50">
        <div className="container mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded-lg bg-primary flex items-center justify-center text-primary-foreground shadow-sm">
              <span className="font-bold text-lg">N</span>
            </div>
            <div>
              <h1 className="text-lg font-bold tracking-tight leading-none">Night Shift</h1>
              <p className="text-[10px] text-muted-foreground font-medium uppercase tracking-wider mt-0.5">Report Automator</p>
            </div>
          </div>
          <div className="flex items-center gap-3 text-xs font-medium text-muted-foreground">
            {jobId && step > 1 && (
              <button
                onClick={handleCancelSession}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-full border border-destructive/30 text-destructive hover:bg-destructive/10 transition-colors"
                title="Discard the current job and start a new upload"
              >
                <RotateCcw className="w-3 h-3" />
                <span className="hidden sm:inline">Cancel Session</span>
              </button>
            )}
            <div className="hidden sm:flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-muted/50 border">
              <div className="w-1.5 h-1.5 rounded-full bg-green-500 animate-pulse" />
              System Online
            </div>
          </div>
        </div>
      </header>

      <main className={`container mx-auto px-4 py-12 transition-all duration-500 ${step === 3 ? "max-w-[1400px]" : "max-w-5xl"}`}>
        {/* Progress Tracker */}
        <div className="mb-12 max-w-3xl mx-auto">
          <div className="flex items-center justify-between relative px-2">
            <div className="absolute left-0 top-4 w-full h-[2px] bg-muted rounded-full">
              <div 
                className="h-full bg-primary transition-all duration-700 ease-in-out shadow-[0_0_8px_rgba(37,99,235,0.4)]" 
                style={{ width: `${((step - 1) / (steps.length - 1)) * 100}%` }}
              />
            </div>
            {steps.map((s) => (
              <div 
                key={s.id} 
                className={`relative z-10 flex flex-col items-center gap-3 transition-all duration-300 ${
                  step >= s.id ? "opacity-100" : "opacity-40"
                }`}
              >
                <div 
                  className={`w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold transition-all duration-500 ${
                    step >= s.id 
                      ? "bg-primary text-primary-foreground shadow-lg shadow-primary/25 scale-110" 
                      : "bg-background text-muted-foreground border-2 border-muted"
                  }`}
                >
                  {s.id}
                </div>
                <span className={`text-[10px] font-bold uppercase tracking-widest ${
                  step === s.id ? "text-primary" : "text-muted-foreground"
                }`}>
                  {s.title}
                </span>
              </div>
            ))}
          </div>
        </div>

        {/* Step Content */}
        <div className="bg-card border border-border/50 rounded-2xl shadow-xl shadow-primary/5 min-h-[450px] overflow-hidden">
          <div className="h-full p-8 sm:p-10">
            {step === 1 && (
            <UploadStep 
              onComplete={(id, data) => {
                setJobId(id)
                setPreviewData(data)
                setStep(2)
              }} 
            />
          )}
          {step === 2 && (
            <ReviewStep 
              jobId={jobId!} 
              previewData={previewData!} 
              onNext={() => setStep(3)} 
            />
          )}
          {step === 3 && (
            <RoutingStep 
              jobId={jobId!} 
              previewData={previewData!} 
              onNext={() => setStep(4)} 
            />
          )}
          {step === 4 && (
            <ProcessingStep 
              jobId={jobId!} 
              onComplete={(outputs) => {
                setFinalOutputs(outputs)
                setStep(5)
              }} 
            />
          )}
          {step === 5 && (
              <DownloadStep 
                jobId={jobId!} 
                outputs={finalOutputs}
                onReset={handleReset} 
              />
            )}
          </div>
        </div>
      </main>
    </div>
  )
}

export default App
