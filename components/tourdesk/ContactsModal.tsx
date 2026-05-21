// components/tourdesk/ContactsModal.tsx
"use client"
import { useState, useEffect } from "react"
import { Trash2, X } from "lucide-react"
import { type Contact, getContacts, deleteContact } from "@/lib/contacts"

interface Props {
  isOpen: boolean
  onClose: () => void
  onSelect: (contact: Contact) => void
}

export function ContactsModal({ isOpen, onClose, onSelect }: Props) {
  const [contacts, setContacts] = useState<Contact[]>([])

  useEffect(() => {
    if (isOpen) setContacts(getContacts())
  }, [isOpen])

  if (!isOpen) return null

  const handleDelete = (fullName: string, e: React.MouseEvent) => {
    e.stopPropagation()
    deleteContact(fullName)
    setContacts(getContacts()) // Listeyi yenile
  }

  return (
    <div className="fixed inset-0 bg-background/80 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <div className="bg-card border border-border rounded-lg shadow-lg w-full max-w-md flex flex-col max-h-[80vh]">
        <div className="flex justify-between items-center p-4 border-b border-border">
          <h2 className="text-lg font-semibold text-foreground">Rehber Kayıtları</h2>
          <button onClick={onClose} className="text-muted-foreground hover:text-foreground transition-colors">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="overflow-y-auto p-4 flex-1">
          {contacts.length === 0 ? (
            <p className="text-sm text-muted-foreground text-center py-4">Kayıtlı kişi bulunamadı.</p>
          ) : (
            <ul className="space-y-2">
              {contacts.map((c) => (
                <li
                  key={c.fullName}
                  onClick={() => onSelect(c)}
                  className="flex justify-between items-center p-3 bg-accent/20 hover:bg-accent/50 rounded-md cursor-pointer border border-border/50 transition-colors group"
                >
                  <div>
                    <p className="text-sm font-medium text-foreground">{c.fullName}</p>
                    <p className="text-xs text-muted-foreground">{c.tcKimlikNo} • {c.birthDate}</p>
                  </div>
                  <button
                    onClick={(e) => handleDelete(c.fullName, e)}
                    className="text-muted-foreground hover:text-destructive p-2 rounded-md opacity-0 group-hover:opacity-100 transition-all"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>
      </div>
    </div>
  )
}