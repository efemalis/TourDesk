export interface TourInfo {
  tourName: string
  artist: string
  startDate: string
  endDate: string
}

export interface CrewMember {
  id: string
  fullName: string
  role: string
  budget: string
  rowColor?: string
}

export interface FlightRow {
  id: string
  role: string
  fullName: string
  tcKimlikNo: string
  birthDate: string
  flightClass: string
  departureAirport: string
  departureTime: string
  returnAirport: string
  returnTime: string
  rowColor?: string
}

export interface AccommodationRow {
  id: string
  fullName: string
  tcKimlikNo: string
  birthDate: string
  roomType: "King Suite" | "Single" | "Double"
  roomLabel: string
  roommate: string
  rowColor?: string
}

// CSS hex → ExcelJS ARGB (e.g. "#FFCDD2" → "FFFFCDD2")
export function hexToArgb(hex: string): string {
  const clean = hex.replace("#", "")
  return "FF" + clean.toUpperCase()
}
export interface ItineraryRow {
  id: string
  time: string
  activity: string
  details: string // location satırı buradan silindi
  rowColor?: string
}