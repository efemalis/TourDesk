"use client"

import type { TourInfo } from "@/lib/tourdesk-types"

interface TopBarProps {
  tourInfo: TourInfo
  onChange: (field: keyof TourInfo, value: string) => void
}

export function TopBar({ tourInfo, onChange }: TopBarProps) {
  return (
    <header className="border-b border-border bg-card px-6 py-4">
      <div className="flex items-center gap-3 mb-4">
        <div className="w-2 h-2 rounded-full bg-foreground opacity-60" />
        <h1 className="text-base font-semibold tracking-widest uppercase text-foreground/70">
          TourDesk
        </h1>
      </div>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <div className="flex flex-col gap-1">
          <label className="text-[10px] uppercase tracking-widest text-muted-foreground font-medium">
            Tur Adı
          </label>
          <input
            type="text"
            value={tourInfo.tourName}
            onChange={(e) => onChange("tourName", e.target.value)}
            placeholder="Dünya Turu 2025"
            className="bg-background border border-border rounded px-3 py-2 text-sm text-foreground placeholder:text-muted-foreground focus:outline-none focus:border-foreground/40 transition-colors"
          />
        </div>
        <div className="flex flex-col gap-1">
          <label className="text-[10px] uppercase tracking-widest text-muted-foreground font-medium">
            Sanatçı / Grup
          </label>
          <input
            type="text"
            value={tourInfo.artist}
            onChange={(e) => onChange("artist", e.target.value)}
            placeholder="Sanatçı Adı"
            className="bg-background border border-border rounded px-3 py-2 text-sm text-foreground placeholder:text-muted-foreground focus:outline-none focus:border-foreground/40 transition-colors"
          />
        </div>
        <div className="flex flex-col gap-1">
          <label className="text-[10px] uppercase tracking-widest text-muted-foreground font-medium">
            Başlangıç Tarihi
          </label>
          <input
            type="date"
            value={tourInfo.startDate}
            onChange={(e) => onChange("startDate", e.target.value)}
            className="bg-background border border-border rounded px-3 py-2 text-sm text-foreground focus:outline-none focus:border-foreground/40 transition-colors [color-scheme:dark]"
          />
        </div>
        <div className="flex flex-col gap-1">
          <label className="text-[10px] uppercase tracking-widest text-muted-foreground font-medium">
            Bitiş Tarihi
          </label>
          <input
            type="date"
            value={tourInfo.endDate}
            onChange={(e) => onChange("endDate", e.target.value)}
            className="bg-background border border-border rounded px-3 py-2 text-sm text-foreground focus:outline-none focus:border-foreground/40 transition-colors [color-scheme:dark]"
          />
        </div>
      </div>
    </header>
  )
}
