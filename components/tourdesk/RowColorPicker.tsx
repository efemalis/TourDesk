"use client"

import { useState, useRef, useEffect } from "react"

const PALETTE = [
  { label: "Beyaz", value: "#FFFFFF" },
  { label: "Sarı", value: "#FFF9C4" },
  { label: "Turuncu", value: "#FFE0B2" },
  { label: "Kırmızı", value: "#FFCDD2" },
  { label: "Yeşil", value: "#C8E6C9" },
  { label: "Mavi", value: "#BBDEFB" },
  { label: "Mor", value: "#E1BEE7" },
  { label: "Gri", value: "#F5F5F5" },
]

interface RowColorPickerProps {
  value?: string
  onChange: (color: string | undefined) => void
}

export function RowColorPicker({ value, onChange }: RowColorPickerProps) {
  const [open, setOpen] = useState(false)
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    function handleClick(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        setOpen(false)
      }
    }
    if (open) document.addEventListener("mousedown", handleClick)
    return () => document.removeEventListener("mousedown", handleClick)
  }, [open])

  return (
    <div ref={ref} className="relative flex items-center justify-center">
      <button
        type="button"
        aria-label="Satır rengini seç"
        onClick={() => setOpen((o) => !o)}
        className="w-4 h-4 rounded-full border border-border/60 flex-shrink-0 transition-transform hover:scale-110 focus:outline-none"
        style={{
          background: value ?? "#2a2a2a",
          boxShadow: value ? `0 0 0 1.5px ${value}55` : undefined,
        }}
      />
      {open && (
        <div
          className="absolute left-5 top-1/2 -translate-y-1/2 z-50 bg-[#1a1a1a] border border-border rounded-lg p-2 shadow-xl flex flex-col gap-1.5 min-w-[96px]"
          role="menu"
        >
          <div className="grid grid-cols-4 gap-1.5">
            {PALETTE.map((p) => (
              <button
                key={p.value}
                type="button"
                aria-label={p.label}
                title={p.label}
                onClick={() => { onChange(p.value); setOpen(false) }}
                className="w-5 h-5 rounded-full border border-black/20 hover:scale-125 transition-transform focus:outline-none"
                style={{ background: p.value }}
              />
            ))}
          </div>
          <button
            type="button"
            onClick={() => { onChange(undefined); setOpen(false) }}
            className="text-[10px] text-muted-foreground hover:text-foreground text-center mt-0.5 transition-colors"
          >
            Sıfırla
          </button>
        </div>
      )}
    </div>
  )
}
