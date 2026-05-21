// lib/contacts.ts
const STORAGE_KEY = "tourdesk_contacts"

export interface Contact {
  fullName: string
  tcKimlikNo: string
  birthDate: string
}

export function getContacts(): Contact[] {
  if (typeof window === "undefined") return []
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    return raw ? JSON.parse(raw) : []
  } catch {
    return []
  }
}

export function saveContact(contact: Contact): void {
  if (!contact.fullName || !contact.tcKimlikNo || !contact.birthDate) return
  const contacts = getContacts()
  const idx = contacts.findIndex(
    (c) => c.fullName.toLowerCase() === contact.fullName.toLowerCase()
  )
  if (idx >= 0) {
    contacts[idx] = contact
  } else {
    contacts.push(contact)
  }
  localStorage.setItem(STORAGE_KEY, JSON.stringify(contacts))
}

export function deleteContact(fullName: string): void {
  const contacts = getContacts().filter(
    (c) => c.fullName.toLowerCase() !== fullName.toLowerCase()
  )
  localStorage.setItem(STORAGE_KEY, JSON.stringify(contacts))
}

export function searchContacts(query: string, limit = 5): Contact[] {
  if (!query.trim()) return []
  const q = query.toLowerCase()
  return getContacts()
    .filter((c) => c.fullName.toLowerCase().includes(q))
    .slice(0, limit)
}