# TourDesk
 
**Offline-first tour & artist management platform for real-world live event operations.**
 
Built out of frustration with spreadsheets — TourDesk replaces the manual, error-prone process of managing crew lists, flight schedules, accommodation, and daily show flow with a fast, structured, and exportable dashboard.
 
---
 
## Features
 
### 4 Core Modules
- **Crew List** — Track every team member, their role, and budget allocation
- **Flight Planning** — Manage departure/return flights with default time propagation across rows
- **Accommodation** — Room assignments with roommate matching for double rooms
- **Call Sheet** — Hourly show day schedule with time auto-carry between entries
### Smart Data Entry
- **Default time carry** — Set a departure time once, every new row inherits it automatically (still editable per row)
- **Contact Memory** — Names, T.C. ID numbers, and birth dates are saved to localStorage as you type. Start typing a name on any future list and it auto-fills from your contact history
- **Rehber (Contacts Panel)** — View, search, and insert saved contacts into any active module with one click
### Premium Excel Export
Agency-grade `.xlsx` reports powered by **ExcelJS**:
- Custom fonts, cell heights, and background colors matching professional tour documents
- Color-coded rows (e.g. BUSINESS class highlighted in red)
- Merged title cells with tour name and date range
- Notes section appended in red at the bottom of each sheet
- Per-row color tagging carried into the Excel output
### Row Color Tagging
Assign a color to any row directly in the UI — useful for flagging VIPs, grouping departments, or marking pending confirmations. Colors export to Excel.
 
---
 
## Tech Stack
 
| Layer | Technology |
|-------|-----------|
| Framework | Next.js (App Router) |
| Language | TypeScript |
| Styling | Tailwind CSS |
| Excel Engine | ExcelJS |
| Icons | Lucide React |
| Persistence | localStorage API |
| Rendering | Client-side only (`"use client"`) |
 
---
 
## Architecture
 
```
components/
└── tourdesk/
    ├── TourDesk.tsx        # Parent — holds all state, distributes to modules
    ├── CrewList.tsx         # Crew & budget module
    ├── FlightList.tsx       # Flight planning module
    ├── Accommodation.tsx    # Hotel & room assignment module
    ├── ItineraryList.tsx    # Call sheet module
    ├── TopBar.tsx           # Tour info bar
    ├── RowColorPicker.tsx   # Per-row color selector
    └── ToastNotification.tsx
lib/
└── contacts.ts             # localStorage contact CRUD + fuzzy search
```
 
State lives at the `TourDesk.tsx` level and flows down via props — no external state manager needed. Each row carries a `crypto.randomUUID()` ID for stable React reconciliation.
 
---
 
## Getting Started
 
```bash
git clone https://github.com/efemalis/tourdesk.git
cd tourdesk
npm install
npm run dev
```
 
Open [http://localhost:3000](http://localhost:3000)
 
> No backend, no API keys, no database. Everything runs in the browser.
 
---
 
## Why This Exists
 
I've been working as a freelance tour & artist manager since 2021. Every show involved rebuilding the same Excel files from scratch — crew lists, flight tables, hotel rooming lists, call sheets — formatted just right for venues, airlines, and promoters.
 
TourDesk automates that. The contact memory alone saves 10–15 minutes per document on recurring tours.
 
---
 
## License
 
MIT
