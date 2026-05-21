"use client"

import { useEffect, useState } from "react"
import { CheckCircle } from "lucide-react"

interface ToastProps {
  message: string
  visible: boolean
}

export function ToastNotification({ message, visible }: ToastProps) {
  const [show, setShow] = useState(false)

  useEffect(() => {
    if (visible) {
      setShow(true)
      const timer = setTimeout(() => setShow(false), 3000)
      return () => clearTimeout(timer)
    }
  }, [visible])

  return (
    <div
      className={`fixed bottom-6 right-6 z-50 flex items-center gap-2 px-4 py-3 rounded-md bg-card border border-border text-foreground text-sm font-medium shadow-lg transition-all duration-300 ${
        show ? "opacity-100 translate-y-0" : "opacity-0 translate-y-2 pointer-events-none"
      }`}
    >
      <CheckCircle className="w-4 h-4 text-green-400 shrink-0" />
      {message}
    </div>
  )
}
