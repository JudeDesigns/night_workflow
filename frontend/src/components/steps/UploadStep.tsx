import { useState, useRef } from "react"
import { uploadFile } from "../../lib/api"
import type { PreviewData } from "../../App"
import { UploadCloud, Loader2, AlertCircle } from "lucide-react"

interface Props {
  onComplete: (jobId: string, data: PreviewData) => void
}

export default function UploadStep({ onComplete }: Props) {
  const [isDragging, setIsDragging] = useState(false)
  const [isUploading, setIsUploading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const handleFile = async (file: File) => {
    if (!file.name.endsWith(".xlsx")) {
      setError("Only .xlsx files are supported.")
      return
    }
    
    setError(null)
    setIsUploading(true)
    
    try {
      const response = await uploadFile(file)
      // Assuming response is fast enough and doesn't need polling for Part 1
      onComplete(response.jobId, response.preview)
    } catch (err: any) {
      setError(err.response?.data?.detail || "An error occurred during upload. Please try again.")
    } finally {
      setIsUploading(false)
    }
  }

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(false)
    
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFile(e.dataTransfer.files[0])
    }
  }

  return (
    <div className="flex flex-col items-center justify-center py-10 text-center h-full animate-in fade-in duration-700">
      <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-primary/10 text-primary text-[10px] font-bold uppercase tracking-wider mb-6">
        Step 1: Data Acquisition
      </div>
      <h2 className="text-3xl font-bold tracking-tight mb-3">Upload All Orders</h2>
      <p className="text-muted-foreground mb-10 max-w-md mx-auto leading-relaxed">
        Our system will automatically compile inventory, sales history, and shortages from your master spreadsheet.
      </p>

      <div 
        className={`w-full max-w-xl border-2 border-dashed rounded-[2rem] p-16 transition-all duration-300 flex flex-col items-center justify-center cursor-pointer group relative overflow-hidden ${
          isDragging 
            ? "border-primary bg-primary/[0.03] scale-[1.01] shadow-2xl shadow-primary/5" 
            : "border-border/60 hover:border-primary/40 hover:bg-muted/30"
        } ${isUploading ? "pointer-events-none opacity-80" : ""}`}
        onDragOver={(e) => { e.preventDefault(); setIsDragging(true) }}
        onDragLeave={() => setIsDragging(false)}
        onDrop={handleDrop}
        onClick={() => fileInputRef.current?.click()}
      >
        <input 
          type="file" 
          ref={fileInputRef} 
          className="hidden" 
          accept=".xlsx"
          onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])}
        />
        
        {isUploading ? (
          <div className="flex flex-col items-center gap-6 text-primary animate-in fade-in zoom-in duration-500">
            <div className="relative">
              <Loader2 className="w-16 h-16 animate-spin opacity-20" />
              <div className="absolute inset-0 flex items-center justify-center">
                <div className="w-2 h-2 bg-primary rounded-full animate-ping" />
              </div>
            </div>
            <div className="space-y-1">
              <p className="font-bold text-lg tracking-tight">Analyzing Workbook</p>
              <p className="text-xs text-muted-foreground animate-pulse">Scanning sheets and mapping columns...</p>
            </div>
          </div>
        ) : (
          <div className="flex flex-col items-center gap-6">
            <div className={`w-20 h-20 rounded-2xl flex items-center justify-center transition-all duration-500 ${
              isDragging ? "bg-primary text-primary-foreground rotate-6 scale-110 shadow-xl" : "bg-muted text-primary group-hover:scale-110 group-hover:-rotate-3"
            }`}>
              <UploadCloud className="w-10 h-10" />
            </div>
            <div className="space-y-2">
              <p className="font-bold text-xl text-foreground tracking-tight">
                {isDragging ? "Drop to Process" : "Select Source File"}
              </p>
              <p className="text-sm text-muted-foreground font-medium">
                Drag and drop your <span className="text-primary font-bold">.xlsx</span> report
              </p>
            </div>
            <div className="mt-4 px-6 py-2 rounded-full bg-background border border-border shadow-sm text-xs font-bold text-muted-foreground group-hover:text-primary group-hover:border-primary/30 transition-colors">
              Browse Local Files
            </div>
          </div>
        )}
      </div>

      {error && (
        <div className="mt-8 flex items-center gap-3 text-destructive bg-destructive/[0.03] border border-destructive/10 px-5 py-4 rounded-xl animate-in slide-in-from-top-4 duration-500">
          <AlertCircle className="w-5 h-5 flex-shrink-0" />
          <p className="text-sm font-bold tracking-tight">{error}</p>
        </div>
      )}
    </div>
  )
}
