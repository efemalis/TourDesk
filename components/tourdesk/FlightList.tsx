"use client"

import { useId, useState } from "react"
import { Trash2 } from "lucide-react"
import ExcelJS from "exceljs"
import type { FlightRow, TourInfo } from "@/lib/tourdesk-types"
import { hexToArgb } from "@/lib/tourdesk-types"
import { RowColorPicker } from "./RowColorPicker"
import { type Contact, saveContact, searchContacts } from "@/lib/contacts"

interface FlightListProps {
  rows: FlightRow[]
  notes: string
  defaultDepartureTime: string
  defaultReturnTime: string
  tourInfo: TourInfo
  onRowsChange: (rows: FlightRow[]) => void
  onNotesChange: (val: string) => void
  onDefaultDepartureChange: (val: string) => void
  onDefaultReturnChange: (val: string) => void
  onExportSuccess: () => void
}

function newRow(defaultDep: string, defaultRet: string): FlightRow {
  return {
    id: crypto.randomUUID(),
    role: "",
    fullName: "",
    tcKimlikNo: "",
    birthDate: "",
    flightClass: "",
    departureAirport: "",
    departureTime: defaultDep,
    returnAirport: "",
    returnTime: defaultRet,
  }
}

const COL_COUNT = 10 

function applyBorder(cell: ExcelJS.Cell) {
  const side = { style: "thin" as const, color: { argb: "FFD4D4D4" } }
  cell.border = { top: side, left: side, bottom: side, right: side }
}

export function FlightList({
  rows,
  notes,
  defaultDepartureTime,
  defaultReturnTime,
  tourInfo,
  onRowsChange,
  onNotesChange,
  onDefaultDepartureChange,
  onDefaultReturnChange,
  onExportSuccess,
}: FlightListProps) {
  const notesId = useId()
  const depTimeId = useId()
  const retTimeId = useId()

  const [dropdownRowId, setDropdownRowId] = useState<string | null>(null)
  const [filteredContacts, setFilteredContacts] = useState<Contact[]>([])
  const [fileName, setFileName] = useState("") // Yeni dosya adı state'i

  function updateRow(id: string, field: keyof FlightRow, value: string) {
    if (field === "departureTime") onDefaultDepartureChange(value);
    if (field === "returnTime") onDefaultReturnChange(value);

    const updatedRows = rows.map((r) => {
      if (r.id === id) {
        const updatedRow = { ...r, [field]: value }
        
        if (updatedRow.fullName?.length > 2 && updatedRow.tcKimlikNo?.length === 11 && updatedRow.birthDate) {
          saveContact({ 
            fullName: updatedRow.fullName, 
            tcKimlikNo: updatedRow.tcKimlikNo, 
            birthDate: updatedRow.birthDate 
          })
        }
        return updatedRow
      }
      return r
    })
    onRowsChange(updatedRows)
  }

  function handleNameChange(id: string, value: string) {
    updateRow(id, "fullName", value)
    
    if (value.length > 1) {
      const matches = searchContacts(value)
      setFilteredContacts(matches)
      setDropdownRowId(matches.length > 0 ? id : null)
    } else {
      setDropdownRowId(null)
    }
  }

  function handleContactSelect(id: string, contact: Contact) {
    const updatedRows = rows.map((r) => {
      if (r.id === id) {
        return { ...r, fullName: contact.fullName, tcKimlikNo: contact.tcKimlikNo, birthDate: contact.birthDate }
      }
      return r
    })
    onRowsChange(updatedRows)
    setDropdownRowId(null)
  }

  function setRowColor(id: string, color: string | undefined) {
    onRowsChange(rows.map((r) => (r.id === id ? { ...r, rowColor: color } : r)))
  }

  function deleteRow(id: string) {
    onRowsChange(rows.filter((r) => r.id !== id))
  }

  function addRow() {
    onRowsChange([...rows, newRow(defaultDepartureTime, defaultReturnTime)])
  }

  async function exportToExcel() {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet("Uçuş Listesi")

    const titleText = `${tourInfo.tourName || "Tur"} — ${tourInfo.artist || "Sanatçı"}  |  ${tourInfo.startDate || ""} → ${tourInfo.endDate || ""}`

    const titleRow = ws.addRow([titleText])
    titleRow.height = 35
    ws.mergeCells(1, 1, 1, COL_COUNT)
    const titleCell = ws.getCell(1, 1)
    titleCell.font = { name: "Segoe UI", bold: true, size: 12, color: { argb: "FF2C3E50" } }
    titleCell.alignment = { horizontal: "center", vertical: "middle" }
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8F9FA" } }
    titleCell.border = { bottom: { style: 'medium', color: { argb: 'FFBDC3C7' } } }

    const headerRow = ws.addRow([
      "#", "Görevi", "Ad Soyad", "T.C. Kimlik No", "Doğum Tarihi",
      "Sınıf", "Gidiş Havalimanı", "Gidiş Saati", "Dönüş Havalimanı", "Dönüş Saati",
    ])
    headerRow.height = 25
    headerRow.eachCell((cell) => {
      cell.font = { name: "Segoe UI", bold: true, size: 9, color: { argb: "FFFFFFFF" } } 
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } } 
      cell.alignment = { horizontal: "center", vertical: "middle" }
      applyBorder(cell)
    })

    rows.forEach((r, i) => {
      const isBusiness = r.flightClass.toUpperCase() === "BUSINESS"
      const dr = ws.addRow([
        i + 1, r.role, r.fullName, r.tcKimlikNo, r.birthDate,
        r.flightClass, r.departureAirport, r.departureTime, r.returnAirport, r.returnTime,
      ])
      
      dr.height = 20

      const fillArgb = r.rowColor ? hexToArgb(r.rowColor) : isBusiness ? "FFFF0000" : null

      dr.eachCell((cell, colNum) => {
        applyBorder(cell)
        cell.font = { name: "Segoe UI", size: 9 }

        if (fillArgb) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fillArgb } }
        }

        if (colNum === 1) {
          cell.font = { name: "Segoe UI", bold: true, size: 9, color: { argb: "FF555555" } }
          cell.alignment = { horizontal: "center", vertical: "middle" }
          if (!fillArgb) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4F6F7" } }
        } 
        else if (colNum === 2 || colNum === 3) {
          cell.alignment = { horizontal: "left", vertical: "middle", indent: 1 }
        } 
        else {
          cell.alignment = { horizontal: "center", vertical: "middle" }
        }
      })
    })

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
      { width: 5 }, { width: 16 }, { width: 22 }, { width: 14 },
      { width: 12 }, { width: 9 }, { width: 15 }, { width: 10 }, { width: 15 }, { width: 10 },
    ]

    const buffer = await wb.xlsx.writeBuffer()
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
    const url = URL.createObjectURL(blob)
    const a = document.createElement("a")
    a.href = url
    
    // Özel Dosya Adı Mantığı
    const date = new Date().toISOString().slice(0, 10)
    const defaultName = `UcusListesi_${(tourInfo.tourName || "Tur").replace(/\s+/g, "_")}_${date}`
    a.download = fileName.trim() ? `${fileName.trim()}.xlsx` : `${defaultName}.xlsx`
    
    a.click()
    URL.revokeObjectURL(url)
    onExportSuccess()
  }

  const CELL_CLS = "bg-transparent border-0 outline-none w-full text-sm text-foreground placeholder:text-muted-foreground focus:bg-accent/30 rounded px-1 py-0.5 transition-colors"
  const DATE_CLS = "bg-transparent border-0 outline-none w-full text-sm text-foreground focus:bg-accent/30 rounded px-1 py-0.5 transition-colors [color-scheme:dark]"

  return (
    <div className="flex flex-col gap-4">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2">
            <label htmlFor={depTimeId} className="text-[10px] uppercase tracking-widest text-muted-foreground font-medium whitespace-nowrap">
              Varsayılan Gidiş
            </label>
            <input
              id={depTimeId}
              type="time"
              value={defaultDepartureTime}
              onChange={(e) => onDefaultDepartureChange(e.target.value)}
              className="bg-card border border-border rounded px-2 py-1 text-xs text-foreground focus:outline-none focus:border-foreground/40 transition-colors [color-scheme:dark]"
            />
          </div>
          <div className="flex items-center gap-2">
            <label htmlFor={retTimeId} className="text-[10px] uppercase tracking-widest text-muted-foreground font-medium whitespace-nowrap">
              Varsayılan Dönüş
            </label>
            <input
              id={retTimeId}
              type="time"
              value={defaultReturnTime}
              onChange={(e) => onDefaultReturnChange(e.target.value)}
              className="bg-card border border-border rounded px-2 py-1 text-xs text-foreground focus:outline-none focus:border-foreground/40 transition-colors [color-scheme:dark]"
            />
          </div>
        </div>

        <div className="flex items-center gap-3">
          <p className="text-xs text-muted-foreground">{rows.length} uçuş</p>
          
          {/* Dosya Adı ve İndirme Butonu */}
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
      </div>

      <div className="overflow-visible rounded border border-border pb-20">
        <table className="w-full text-sm border-collapse">
          <thead>
            <tr className="border-b border-border">
              <th className="w-6 px-2 py-2.5" />
              {["#", "Görevi", "Ad Soyad", "T.C. Kimlik No", "Doğum Tarihi", "Sınıf", "Gidiş Havalimanı", "Gidiş Saati", "Dönüş Havalimanı", "Dönüş Saati", ""].map((h, i) => (
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
                <td className="px-2 py-1.5 min-w-[100px]">
                  <input className={CELL_CLS} value={row.role} onChange={(e) => updateRow(row.id, "role", e.target.value)} placeholder="Görevi" style={{ color: row.rowColor ? "#111" : undefined }} />
                </td>
                
                <td className="px-2 py-1.5 min-w-[160px] relative">
                  <input 
                    className={CELL_CLS} 
                    value={row.fullName} 
                    onChange={(e) => handleNameChange(row.id, e.target.value)}
                    onFocus={(e) => handleNameChange(row.id, e.target.value)}
                    onBlur={() => setTimeout(() => setDropdownRowId(null), 200)}
                    placeholder="Ad Soyad" 
                    style={{ color: row.rowColor ? "#111" : undefined }} 
                  />
                  {dropdownRowId === row.id && (
                    <div className="absolute z-50 top-full left-0 w-full min-w-[220px] bg-card border border-border shadow-lg rounded-md mt-1 max-h-48 overflow-y-auto">
                      {filteredContacts.map(c => (
                        <div 
                          key={c.fullName} 
                          onMouseDown={(e) => e.preventDefault()}
                          onClick={() => handleContactSelect(row.id, c)} 
                          className="p-2.5 text-xs hover:bg-accent cursor-pointer border-b border-border/50 last:border-0 text-foreground transition-colors"
                        >
                          <span className="font-medium text-sm">{c.fullName}</span> <br/>
                          <span className="text-muted-foreground text-[10px]">{c.tcKimlikNo} • {c.birthDate}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </td>

                <td className="px-2 py-1.5 min-w-[120px]">
                  <input className={CELL_CLS} value={row.tcKimlikNo} onChange={(e) => updateRow(row.id, "tcKimlikNo", e.target.value)} placeholder="12345678901" maxLength={11} style={{ color: row.rowColor ? "#111" : undefined }} />
                </td>
                <td className="px-2 py-1.5 min-w-[120px]">
                  <input type="date" className={DATE_CLS} value={row.birthDate} onChange={(e) => updateRow(row.id, "birthDate", e.target.value)} style={{ color: row.rowColor ? "#111" : undefined }} />
                </td>
                <td className="px-2 py-1.5 min-w-[80px]">
                  <input className={CELL_CLS} value={row.flightClass} onChange={(e) => updateRow(row.id, "flightClass", e.target.value)} placeholder="Economy" style={{ color: row.rowColor ? "#111" : undefined }} />
                </td>
                <td className="px-2 py-1.5 min-w-[100px]">
                  <input className={CELL_CLS} value={row.departureAirport} onChange={(e) => updateRow(row.id, "departureAirport", e.target.value)} placeholder="IST" style={{ color: row.rowColor ? "#111" : undefined }} />
                </td>
                <td className="px-2 py-1.5 min-w-[90px]">
                  <input type="time" className={DATE_CLS} value={row.departureTime} onChange={(e) => updateRow(row.id, "departureTime", e.target.value)} style={{ color: row.rowColor ? "#111" : undefined }} />
                </td>
                <td className="px-2 py-1.5 min-w-[100px]">
                  <input className={CELL_CLS} value={row.returnAirport} onChange={(e) => updateRow(row.id, "returnAirport", e.target.value)} placeholder="SAW" style={{ color: row.rowColor ? "#111" : undefined }} />
                </td>
                <td className="px-2 py-1.5 min-w-[90px]">
                  <input type="time" className={DATE_CLS} value={row.returnTime} onChange={(e) => updateRow(row.id, "returnTime", e.target.value)} style={{ color: row.rowColor ? "#111" : undefined }} />
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
        <label htmlFor={notesId} className="text-[10px] uppercase tracking-widest text-muted-foreground font-medium">Notlar</label>
        <textarea id={notesId} value={notes} onChange={(e) => onNotesChange(e.target.value)} rows={3} placeholder="Notlarınızı buraya yazın…" className="bg-card border border-border rounded px-3 py-2 text-sm text-foreground placeholder:text-muted-foreground focus:outline-none focus:border-foreground/40 resize-y transition-colors" />
        {notes && <p className="text-sm text-red-500 mt-1 whitespace-pre-wrap">{notes}</p>}
      </div>
    </div>
  )
}