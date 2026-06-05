import axios from "axios"
import type { PreviewData } from "../App"

const api = axios.create({
  baseURL: "/api", // Proxy defined in vite.config.ts
})

export async function uploadFile(file: File) {
  const formData = new FormData()
  formData.append("file", file)
  
  const response = await api.post<{ jobId: string; status: string; preview: PreviewData }>("/part1", formData, {
    headers: {
      "Content-Type": "multipart/form-data",
    },
  })
  
  return response.data
}

export async function generateFinalReports(jobId: string) {
  const response = await api.post<{ jobId: string; status: string; outputs: any }>(`/part3-4`, {
    jobId,
  })
  return response.data
}

export async function checkStatus(jobId: string) {
  const response = await api.get<{ jobId: string; status: string }>(`/status/${jobId}`)
  return response.data
}

export async function restoreJob(jobId: string) {
  const response = await api.get<{ jobId: string; status: string; preview: PreviewData; outputs: any }>(`/jobs/${jobId}`)
  return response.data
}

export async function applyRouting(
  jobId: string, 
  sheetRoutingDecisions: any[], 
  warehouseShortDecisions: any[], 
  driverPullSequence: string[]
) {
  const response = await api.post<{ jobId: string; status: string; outputs: any }>(`/part2`, {
    jobId,
    sheetRoutingDecisions,
    warehouseShortDecisions,
    driverPullSequence,
  })
  return response.data
}
