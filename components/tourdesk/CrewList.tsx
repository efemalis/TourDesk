"use client"

import { useId, useState } from "react"
import { Trash2 } from "lucide-react"
import ExcelJS from "exceljs"
import type { CrewMember, TourInfo } from "@/lib/tourdesk-types"
import { hexToArgb } from "@/lib/tourdesk-types"
import { RowColorPicker } from "./RowColorPicker"

interface CrewListProps {
  rows: CrewMember[]
  notes: string
  tourInfo: TourInfo
  onRowsChange: (rows: CrewMember[]) => void
  onNotesChange: (val: string) => void
  onExportSuccess: () => void
}

function newRow(): CrewMember {
  return { id: crypto.randomUUID(), fullName: "", role: "", budget: "" }
}

const COL_COUNT = 4 // #, Ad Soyad, Görevi, Bütçe

function applyBorder(cell: ExcelJS.Cell) {
  const side = { style: "thin" as const, color: { argb: "FFD4D4D4" } }
  cell.border = { top: side, left: side, bottom: side, right: side }
}

export function CrewList({ rows, notes, tourInfo, onRowsChange, onNotesChange, onExportSuccess }: CrewListProps) {
  const notesId = useId()
  const [fileName, setFileName] = useState("") // Dosya adı özelliği eklendi

  function updateRow(id: string, field: keyof CrewMember, value: string) {
    onRowsChange(rows.map((r) => (r.id === id ? { ...r, [field]: value } : r)))
  }

  function setRowColor(id: string, color: string | undefined) {
    onRowsChange(rows.map((r) => (r.id === id ? { ...r, rowColor: color } : r)))
  }

  function deleteRow(id: string) {
    onRowsChange(rows.filter((r) => r.id !== id))
  }

  function addRow() {
    onRowsChange([...rows, newRow()])
  }

  async function exportToExcel() {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet("Ekip Listesi")

    const titleText = `${tourInfo.tourName || "Tur"} — ${tourInfo.artist || "Sanatçı"}  |  Ekip Listesi`

    // Ana Başlık Tasarımı
    const titleRow = ws.addRow([titleText])
    titleRow.height = 35
    ws.mergeCells(1, 1, 1, COL_COUNT)
    const titleCell = ws.getCell(1, 1)
    titleCell.font = { name: "Segoe UI", bold: true, size: 12, color: { argb: "FF2C3E50" } }
    titleCell.alignment = { horizontal: "center", vertical: "middle" }
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8F9FA" } }
    titleCell.border = { bottom: { style: 'medium', color: { argb: 'FFBDC3C7' } } }

    // Sütun Başlıkları Tasarımı
    const headerRow = ws.addRow(["#", "Ad Soyad", "Görevi", "Bütçe"])
    headerRow.height = 25
    headerRow.eachCell((cell) => {
      cell.font = { name: "Segoe UI", bold: true, size: 9, color: { argb: "FFFFFFFF" } }
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } } // Şık Koyu Mavi
      cell.alignment = { horizontal: "center", vertical: "middle" }
      applyBorder(cell)
    })

    // Veri Satırları Tasarımı
    rows.forEach((r, i) => {
      const dr = ws.addRow([i + 1, r.fullName, r.role, r.budget])
      dr.height = 22
      const fillArgb = r.rowColor ? hexToArgb(r.rowColor) : null

      dr.eachCell((cell, colNum) => {
        applyBorder(cell)
        cell.font = { name: "Segoe UI", size: 9.5 }

        if (fillArgb) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fillArgb } }
        }

        if (colNum === 1) {
          cell.font = { name: "Segoe UI", bold: true, size: 9, color: { argb: "FF555555" } }
          cell.alignment = { horizontal: "center", vertical: "middle" }
          if (!fillArgb) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4F6F7" } }
        } else if (colNum === 2 || colNum === 3) {
          cell.alignment = { horizontal: "left", vertical: "middle", indent: 1 }
        } else if (colNum === 4) {
          cell.alignment = { horizontal: "right", vertical: "middle" }
        }
      })
    })

    // Notlar Tasarımı
    if (notes) {
      ws.addRow([]) 
      const notesRow = ws.addRow([`Notlar: ${notes}`])
      notesRow.height = 30
      ws.mergeCells(notesRow.number, 1, notesRow.number, COL_COUNT)
      const nc = ws.getCell(notesRow.number, 1)
      nc.font = { name: "Segoe UI", size: 9, color: { argb: "FFE74C3C" }, italic: true }
      nc.alignment = { horizontal: "left", vertical: "middle", indent: 1 }
      nc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFDF2F0" } }
      
      const noteBorder = { style: 'thin' as const, color: { argb: 'FFFADBD8' } }
      nc.border = { top: noteBorder, bottom: noteBorder, left: noteBorder, right: noteBorder }
    }

    ws.columns = [
      { width: 5 }, { width: 26 }, { width: 22 }, { width: 16 },
    ]

    const buffer = await wb.xlsx.writeBuffer()
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
    const url = URL.createObjectURL(blob)
    const a = document.createElement("a")
    a.href = url
    
    // Özel isim kontrolü
    const date = new Date().toISOString().slice(0, 10)
    const defaultName = `EkipListesi_${(tourInfo.tourName || "Tur").replace(/\s+/g, "_")}_${date}`
    a.download = fileName.trim() ? `${fileName.trim()}.xlsx` : `${defaultName}.xlsx`
    
    a.click()
    URL.revokeObjectURL(url)
    onExportSuccess()
  }

  const CELL_CLS = "bg-transparent border-0 outline-none w-full text-sm text-foreground placeholder:text-muted-foreground focus:bg-accent/30 rounded px-1 py-0.5 transition-colors"

  return (
    <div className="flex flex-col gap-4">
      <div className="flex items-center justify-between">
        <p className="text-xs text-muted-foreground">{rows.length} kişi</p>
        
        {/* Dosya Adı Girişi ve Buton */}
        <div className="flex items-center gap-2">
          <input
            type="text"
            value={fileName}
            onChange={(e) => setFileName(e.target.value)}
            placeholder="Dosya Adı (Opsiyonel)"
            className="bg-card border border-border rounded px-2 py-1.5 text-xs text-foreground placeholder:text-muted-foreground focus:outline-none focus:border-foreground/40 transition-colors w-36 sm:w-48"
          />
          <button
            onClick={exportToExcel}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium border border-border rounded hover:bg-accent transition-colors text-foreground"
          >
            Excel&apos;e Aktar
          </button>
        </div>
      </div>

      <div className="overflow-x-auto rounded border border-border">
        <table className="w-full text-sm border-collapse">
          <thead>
            <tr className="border-b border-border">
              <th className="w-6 px-2 py-2.5" />
              {["#", "Ad Soyad", "Görevi", "Bütçe", ""].map((h, i) => (
                <th
                  key={i}
                  className="px-3 py-2.5 text-left text-[10px] font-medium uppercase tracking-widest text-muted-foreground whitespace-nowrap"
                >
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr
                key={row.id}
                className="border-b border-border last:border-0 transition-colors group"
                style={{ backgroundColor: row.rowColor ?? undefined }}
              >
                <td className="px-2 py-2">
                  <RowColorPicker value={row.rowColor} onChange={(c) => setRowColor(row.id, c)} />
                </td>
                <td className="px-3 py-2 text-muted-foreground text-xs w-8" style={{ color: row.rowColor ? "#555" : undefined }}>{idx + 1}</td>
                <td className="px-2 py-1.5">
                  <input className={CELL_CLS} value={row.fullName} onChange={(e) => updateRow(row.id, "fullName", e.target.value)} placeholder="Ad Soyad" style={{ color: row.rowColor ? "#111" : undefined }} />
                </td>
                <td className="px-2 py-1.5">
                  <input className={CELL_CLS} value={row.role} onChange={(e) => updateRow(row.id, "role", e.target.value)} placeholder="Görevi" style={{ color: row.rowColor ? "#111" : undefined }} />
                </td>
                <td className="px-2 py-1.5">
                  <input className={CELL_CLS} value={row.budget} onChange={(e) => updateRow(row.id, "budget", e.target.value)} placeholder="0.00 ₺" style={{ color: row.rowColor ? "#111" : undefined }} />
                </td>
                <td className="px-2 py-1.5 w-8">
                  <button
                    onClick={() => deleteRow(row.id)}
                    className="opacity-0 group-hover:opacity-100 text-muted-foreground hover:text-destructive transition-all"
                    aria-label="Satırı sil"
                  >
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
        <label htmlFor={notesId} className="text-[10px] uppercase tracking-widest text-muted-foreground font-medium">
          Notlar
        </label>
        <textarea
          id={notesId}
          value={notes}
          onChange={(e) => onNotesChange(e.target.value)}
          rows={3}
          placeholder="Notlarınızı buraya yazın…"
          className="bg-card border border-border rounded px-3 py-2 text-sm text-foreground placeholder:text-muted-foreground focus:outline-none focus:border-foreground/40 resize-y transition-colors"
        />
        {notes && <p className="text-sm text-red-500 mt-1 whitespace-pre-wrap">{notes}</p>}
      </div>
    </div>
  )
}