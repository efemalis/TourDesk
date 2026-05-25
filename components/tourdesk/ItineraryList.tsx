"use client"

import { useId, useState } from "react"
import { Trash2 } from "lucide-react"
import ExcelJS from "exceljs"
import type { ItineraryRow, TourInfo } from "@/lib/tourdesk-types"
import { hexToArgb } from "@/lib/tourdesk-types"
import { RowColorPicker } from "./RowColorPicker"

interface ItineraryListProps {
  rows: ItineraryRow[]
  notes: string
  tourInfo: TourInfo
  onRowsChange: (rows: ItineraryRow[]) => void
  onNotesChange: (val: string) => void
  onExportSuccess: () => void
}

function newRow(lastTime: string): ItineraryRow {
  return {
    id: crypto.randomUUID(),
    time: lastTime || "12:00",
    activity: "",
    details: "",
  }
}

const COL_COUNT = 3  // ← 4'ten 3'e düşürüldü

function applyBorder(cell: ExcelJS.Cell) {
  const side = { style: "thin" as const, color: { argb: "FFD4D4D4" } }
  cell.border = { top: side, left: side, bottom: side, right: side }
}

export function ItineraryList({ rows, notes, tourInfo, onRowsChange, onNotesChange, onExportSuccess }: ItineraryListProps) {
  const notesId = useId()
  const [fileName, setFileName] = useState("")

  function updateRow(id: string, field: keyof ItineraryRow, value: string) {
    const updatedRows = rows.map((r) => (r.id === id ? { ...r, [field]: value } : r))
    onRowsChange(updatedRows)
  }

  function setRowColor(id: string, color: string | undefined) {
    onRowsChange(rows.map((r) => (r.id === id ? { ...r, rowColor: color } : r)))
  }

  function deleteRow(id: string) {
    onRowsChange(rows.filter((r) => r.id !== id))
  }

  function addRow() {
    const lastTime = rows.length > 0 ? rows[rows.length - 1].time : "12:00"
    onRowsChange([...rows, newRow(lastTime)])
  }

  async function exportToExcel() {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet("Günlük Akış")

    const titleText = `${tourInfo.tourName || "Tur"} — ${tourInfo.artist || "Sanatçı"}  |  Günlük Akış Programı`

    const titleRow = ws.addRow([titleText])
    titleRow.height = 35
    ws.mergeCells(1, 1, 1, COL_COUNT)
    const titleCell = ws.getCell(1, 1)
    titleCell.font = { name: "Segoe UI", bold: true, size: 12, color: { argb: "FF2C3E50" } }
    titleCell.alignment = { horizontal: "center", vertical: "middle" }
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8F9FA" } }
    titleCell.border = { bottom: { style: 'medium', color: { argb: 'FFBDC3C7' } } }

    // ← "Mekan ve Açıklama" kaldırıldı
    const headerRow = ws.addRow(["#", "Saat", "Etkinlik / Durum"])
    headerRow.height = 25
    headerRow.eachCell((cell) => {
      cell.font = { name: "Segoe UI", bold: true, size: 9, color: { argb: "FFFFFFFF" } }
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF262626" } }
      cell.alignment = { horizontal: "center", vertical: "middle" }
      applyBorder(cell)
    })

    rows.forEach((r, i) => {
      // ← details kaldırıldı, sadece activity
      const dr = ws.addRow([i + 1, r.time, r.activity])
      dr.height = 22

      const fillArgb = r.rowColor ? hexToArgb(r.rowColor) : null

      dr.eachCell((cell, colNum) => {
        applyBorder(cell)
        cell.font = { name: "Segoe UI", size: 9.5 }

        if (fillArgb) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fillArgb } }
        }

        if (colNum === 1 || colNum === 2) {
          cell.alignment = { horizontal: "center", vertical: "middle" }
          if (colNum === 1) {
            cell.font = { name: "Segoe UI", bold: true, size: 9, color: { argb: "FF555555" } }
            if (!fillArgb) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4F6F7" } }
          }
        } else {
          // ← Etkinlik sütunu: wrapText açık, sola yasla
          cell.alignment = { horizontal: "left", vertical: "middle", indent: 1, wrapText: true }
        }
      })
    })

    if (notes) {
      ws.addRow([])
      const noteLines = notes.split("\n")
      // ← Satır sayısına göre dinamik yükseklik
      const notesRowHeight = Math.max(30, noteLines.length * 16 + 10)

      const notesRow = ws.addRow([`Önemli Notlar: ${notes}`])
      notesRow.height = notesRowHeight
      ws.mergeCells(notesRow.number, 1, notesRow.number, COL_COUNT)
      const nc = ws.getCell(notesRow.number, 1)
      nc.font = { name: "Segoe UI", size: 9.5, color: { argb: "FFC0392B" }, italic: true }
      // ← wrapText: true eklendi — satır sonları Excel'de görünür
      nc.alignment = { horizontal: "left", vertical: "top", indent: 1, wrapText: true }
      nc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFDF2F0" } }
      const noteBorder = { style: 'thin' as const, color: { argb: 'FFFADBD8' } }
      nc.border = { top: noteBorder, bottom: noteBorder, left: noteBorder, right: noteBorder }
    }

    // ← Sütun genişlikleri: #=5, Saat=10, Etkinlik=65 (geniş)
    ws.columns = [
      { width: 5 },
      { width: 10 },
      { width: 65 },
    ]

    const buffer = await wb.xlsx.writeBuffer()
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
    const url = URL.createObjectURL(blob)
    const a = document.createElement("a")
    a.href = url

    const date = new Date().toISOString().slice(0, 10)
    const defaultName = `GunlukAkis_${(tourInfo.tourName || "Tur").replace(/\s+/g, "_")}_${date}`
    a.download = fileName.trim() ? `${fileName.trim()}.xlsx` : `${defaultName}.xlsx`

    a.click()
    URL.revokeObjectURL(url)
    onExportSuccess()
  }

  const CELL_CLS = "bg-transparent border-0 outline-none w-full text-sm text-foreground placeholder:text-muted-foreground focus:bg-accent/30 rounded px-1 py-0.5 transition-colors"
  const TIME_CLS = "bg-transparent border-0 outline-none w-full text-sm text-foreground focus:bg-accent/30 rounded px-1 py-0.5 transition-colors [color-scheme:dark]"

  return (
    <div className="flex flex-col gap-4">
      <div className="flex justify-between items-center">
        <p className="text-xs text-muted-foreground">{rows.length} akış maddesi</p>

        <div className="flex items-center gap-2">
          <input
            type="text"
            value={fileName}
            onChange={(e) => setFileName(e.target.value)}
            placeholder="Dosya Adı (Opsiyonel)"
            className="bg-card border border-border rounded px-2 py-1.5 text-xs text-foreground placeholder:text-muted-foreground focus:outline-none focus:border-foreground/40 transition-colors w-36 sm:w-48"
          />
          <button onClick={exportToExcel} className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border border-border rounded hover:bg-accent transition-colors text-foreground">
            Excel&apos;e Aktar
          </button>
        </div>
      </div>

      <div className="overflow-visible rounded border border-border pb-10">
        <table className="w-full text-sm border-collapse">
          <thead>
            <tr className="border-b border-border">
              <th className="w-6 px-2 py-2.5" />
              {/* ← "Mekan ve Açıklama" başlığı kaldırıldı */}
              {["#", "Saat", "Etkinlik / Durum", ""].map((h, i) => (
                <th key={i} className="px-3 py-2.5 text-left text-[10px] font-medium uppercase tracking-widest text-muted-foreground whitespace-nowrap">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr key={row.id} className="border-b border-border last:border-0 transition-colors group" style={{ backgroundColor: row.rowColor ?? undefined }}>
                <td className="px-2 py-2">
                  <RowColorPicker value={row.rowColor} onChange={(c) => setRowColor(row.id, c)} />
                </td>
                <td className="px-3 py-2 text-muted-foreground text-xs w-8" style={{ color: row.rowColor ? "#555" : undefined }}>{idx + 1}</td>
                <td className="px-2 py-1.5 w-24">
                  <input type="time" className={TIME_CLS} value={row.time} onChange={(e) => updateRow(row.id, "time", e.target.value)} style={{ color: row.rowColor ? "#111" : undefined }} />
                </td>
                {/* ← Tek geniş etkinlik sütunu, details input'u kaldırıldı */}
                <td className="px-2 py-1.5 w-full">
                  <input className={CELL_CLS} value={row.activity} onChange={(e) => updateRow(row.id, "activity", e.target.value)} placeholder="Örn: Lobi Buluşma, Soundcheck, Sahne — mekan ve detay buraya" style={{ color: row.rowColor ? "#111" : undefined }} />
                </td>
                <td className="px-2 py-1.5 w-8">
                  <button onClick={() => deleteRow(row.id)} className="opacity-0 group-hover:opacity-100 text-muted-foreground hover:text-destructive transition-all">
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="px-3 py-2 border-t border-border">
          <button onClick={addRow} className="text-xs text-muted-foreground hover:text-foreground transition-colors">
            + Satır ekle
          </button>
        </div>
      </div>

      <div className="flex flex-col gap-1.5">
        <label htmlFor={notesId} className="text-[10px] uppercase tracking-widest text-muted-foreground font-medium">Günlük Akış Notları</label>
        <textarea id={notesId} value={notes} onChange={(e) => onNotesChange(e.target.value)} rows={3} placeholder="Günlük akışa ait özel notları buraya yazın…" className="bg-card border border-border rounded px-3 py-2 text-sm text-foreground placeholder:text-muted-foreground focus:outline-none focus:border-foreground/40 resize-y transition-colors" />
      </div>
    </div>
  )
}