"use client"

import { useState, useEffect } from "react"
import type { TourInfo, CrewMember, FlightRow, AccommodationRow, ItineraryRow } from "@/lib/tourdesk-types"
import { TopBar } from "./TopBar"
import { CrewList } from "./CrewList"
import { FlightList } from "./FlightList"
import { Accommodation } from "./Accommodation"
import { ItineraryList } from "./ItineraryList"
import { ToastNotification } from "./ToastNotification"
import { ContactsModal } from "./ContactsModal"
import type { Contact } from "@/lib/contacts"
import { BookUser } from "lucide-react"

type Tab = "crew" | "flights" | "accommodation" | "itinerary"

const TABS: { id: Tab; label: string }[] = [
  { id: "crew", label: "Ekip Listesi" },
  { id: "flights", label: "Uçuş Listesi" },
  { id: "accommodation", label: "Konaklama" },
  { id: "itinerary", label: "Günlük Akış" },
]

export function TourDesk() {
  const [tourInfo, setTourInfo] = useState<TourInfo>({
    tourName: "",
    artist: "",
    startDate: "",
    endDate: "",
  })

  const [activeTab, setActiveTab] = useState<Tab>("crew")
  const [isContactsOpen, setIsContactsOpen] = useState(false)

  // Crew state
  const [crewRows, setCrewRows] = useState<CrewMember[]>([])
  const [crewNotes, setCrewNotes] = useState("")

  // Flight state
  const [flightRows, setFlightRows] = useState<FlightRow[]>([])
  const [flightNotes, setFlightNotes] = useState("")
  const [defaultDepartureTime, setDefaultDepartureTime] = useState("09:00")
  const [defaultReturnTime, setDefaultReturnTime] = useState("09:00")

  // Accommodation state
  const [accomRows, setAccomRows] = useState<AccommodationRow[]>([])
  const [accomNotes, setAccomNotes] = useState("")

  // Itinerary state (Günlük Akış eklendi)
  const [itineraryRows, setItineraryRows] = useState<ItineraryRow[]>([])
  const [itineraryNotes, setItineraryNotes] = useState("")

  // Toast state
  const [toastVisible, setToastVisible] = useState(false)
  const [toastKey, setToastKey] = useState(0)

  const [hasLoaded, setHasLoaded] = useState(false)

  // 1. Sayfa yüklendiğinde hafızadan getir
  useEffect(() => {
    if (typeof window === "undefined") return

    const savedTourInfo = localStorage.getItem("tourdesk_active_tour_info")
    const savedCrewRows = localStorage.getItem("tourdesk_active_crew_rows")
    const savedCrewNotes = localStorage.getItem("tourdesk_active_crew_notes")
    const savedFlightRows = localStorage.getItem("tourdesk_active_flight_rows")
    const savedFlightNotes = localStorage.getItem("tourdesk_active_flight_notes")
    const savedAccomRows = localStorage.getItem("tourdesk_active_accom_rows")
    const savedAccomNotes = localStorage.getItem("tourdesk_active_accom_notes")
    const savedItinRows = localStorage.getItem("tourdesk_active_itin_rows")
    const savedItinNotes = localStorage.getItem("tourdesk_active_itin_notes")
    const savedDefDep = localStorage.getItem("tourdesk_active_def_dep")
    const savedDefRet = localStorage.getItem("tourdesk_active_def_ret")

    if (savedTourInfo) setTourInfo(JSON.parse(savedTourInfo))
    if (savedCrewRows) setCrewRows(JSON.parse(savedCrewRows))
    if (savedCrewNotes) setCrewNotes(savedCrewNotes)
    if (savedFlightRows) setFlightRows(JSON.parse(savedFlightRows))
    if (savedFlightNotes) setFlightNotes(savedFlightNotes)
    if (savedAccomRows) setAccomRows(JSON.parse(savedAccomRows))
    if (savedAccomNotes) setAccomNotes(savedAccomNotes)
    if (savedItinRows) setItineraryRows(JSON.parse(savedItinRows))
    if (savedItinNotes) setItineraryNotes(savedItinNotes)
    if (savedDefDep) setDefaultDepartureTime(savedDefDep)
    if (savedDefRet) setDefaultReturnTime(savedDefRet)

    setHasLoaded(true)
  }, [])

  // 2. Değişiklikleri hafızaya kaydet
  useEffect(() => { if (hasLoaded) localStorage.setItem("tourdesk_active_tour_info", JSON.stringify(tourInfo)) }, [tourInfo, hasLoaded])
  useEffect(() => { if (hasLoaded) localStorage.setItem("tourdesk_active_crew_rows", JSON.stringify(crewRows)) }, [crewRows, hasLoaded])
  useEffect(() => { if (hasLoaded) localStorage.setItem("tourdesk_active_crew_notes", crewNotes) }, [crewNotes, hasLoaded])
  useEffect(() => { if (hasLoaded) localStorage.setItem("tourdesk_active_flight_rows", JSON.stringify(flightRows)) }, [flightRows, hasLoaded])
  useEffect(() => { if (hasLoaded) localStorage.setItem("tourdesk_active_flight_notes", flightNotes) }, [flightNotes, hasLoaded])
  useEffect(() => { if (hasLoaded) localStorage.setItem("tourdesk_active_accom_rows", JSON.stringify(accomRows)) }, [accomRows, hasLoaded])
  useEffect(() => { if (hasLoaded) localStorage.setItem("tourdesk_active_accom_notes", accomNotes) }, [accomNotes, hasLoaded])
  useEffect(() => { if (hasLoaded) localStorage.setItem("tourdesk_active_itin_rows", JSON.stringify(itineraryRows)) }, [itineraryRows, hasLoaded])
  useEffect(() => { if (hasLoaded) localStorage.setItem("tourdesk_active_itin_notes", itineraryNotes) }, [itineraryNotes, hasLoaded])
  useEffect(() => { if (hasLoaded) localStorage.setItem("tourdesk_active_def_dep", defaultDepartureTime) }, [defaultDepartureTime, hasLoaded])
  useEffect(() => { if (hasLoaded) localStorage.setItem("tourdesk_active_def_ret", defaultReturnTime) }, [defaultReturnTime, hasLoaded])

  function handleExportSuccess() {
    setToastKey((k) => k + 1)
    setToastVisible(true)
    setTimeout(() => setToastVisible(false), 100)
  }

  function updateTourInfo(field: keyof TourInfo, value: string) {
    setTourInfo((prev) => ({ ...prev, [field]: value }))
  }

  function handleContactSelect(contact: Contact) {
    if (activeTab === "crew") {
      setCrewRows([...crewRows, { id: crypto.randomUUID(), fullName: contact.fullName, role: "", budget: "" }])
    } else if (activeTab === "flights") {
      setFlightRows([
        ...flightRows, 
        {
          id: crypto.randomUUID(), role: "", fullName: contact.fullName,
          tcKimlikNo: contact.tcKimlikNo, birthDate: contact.birthDate,
          flightClass: "", departureAirport: "", 
          departureTime: defaultDepartureTime, 
          returnAirport: "", 
          returnTime: defaultReturnTime 
        }
      ])
    } else if (activeTab === "accommodation") {
      setAccomRows([...accomRows, { id: crypto.randomUUID(), fullName: contact.fullName } as any])
    } else if (activeTab === "itinerary") {
      // location alanı buradan çıkarıldı
      setItineraryRows([...itineraryRows, { id: crypto.randomUUID(), time: "12:00", activity: "Ekip Buluşma", details: `${contact.fullName} katılımı.` }])
    }
    setIsContactsOpen(false)
  }

  return (
    <div className="min-h-screen bg-background text-foreground font-sans flex flex-col">
      <TopBar tourInfo={tourInfo} onChange={updateTourInfo} />

      {/* Tabs */}
      <div className="border-b border-border bg-card px-6 flex justify-between items-end">
        <nav className="flex gap-0" role="tablist" aria-label="Tur bölümleri">
          {TABS.map((tab) => (
            <button
              key={tab.id}
              role="tab"
              aria-selected={activeTab === tab.id}
              onClick={() => setActiveTab(tab.id)}
              className={`relative px-4 py-3 text-sm font-medium transition-colors ${
                activeTab === tab.id ? "text-foreground" : "text-muted-foreground hover:text-foreground/70"
              }`}
            >
              {tab.label}
              {activeTab === tab.id && <span className="absolute bottom-0 left-0 right-0 h-px bg-foreground" />}
            </button>
          ))}
        </nav>
        
        <button 
          onClick={() => setIsContactsOpen(true)}
          className="mb-2 flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium bg-primary/10 text-primary border border-primary/20 rounded hover:bg-primary/20 transition-colors"
        >
          <BookUser className="w-3.5 h-3.5" />
          Rehber
        </button>
      </div>

      {/* Tab content */}
      <main className="flex-1 p-6">
        <div className="max-w-7xl mx-auto">
          {activeTab === "crew" && (
            <CrewList rows={crewRows} notes={crewNotes} tourInfo={tourInfo} onRowsChange={setCrewRows} onNotesChange={setCrewNotes} onExportSuccess={handleExportSuccess} />
          )}
          {activeTab === "flights" && (
            <FlightList rows={flightRows} notes={flightNotes} defaultDepartureTime={defaultDepartureTime} defaultReturnTime={defaultReturnTime} tourInfo={tourInfo} onRowsChange={setFlightRows} onNotesChange={setFlightNotes} onDefaultDepartureChange={setDefaultDepartureTime} onDefaultReturnChange={setDefaultReturnTime} onExportSuccess={handleExportSuccess} />
          )}
          {activeTab === "accommodation" && (
            <Accommodation rows={accomRows} notes={accomNotes} tourInfo={tourInfo} onRowsChange={setAccomRows} onNotesChange={setAccomNotes} onExportSuccess={handleExportSuccess} />
          )}
          {activeTab === "itinerary" && (
            <ItineraryList rows={itineraryRows} notes={itineraryNotes} tourInfo={tourInfo} onRowsChange={setItineraryRows} onNotesChange={setItineraryNotes} onExportSuccess={handleExportSuccess} />
          )}
        </div>
      </main>

      <ToastNotification key={toastKey} message="Dışa aktarma başarılı!" visible={toastVisible} />
      <ContactsModal isOpen={isContactsOpen} onClose={() => setIsContactsOpen(false)} onSelect={handleContactSelect} />
    </div>
  )
}