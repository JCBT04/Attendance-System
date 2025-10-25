"use client"

import type React from "react"

import { useEffect, useRef, useState } from "react"
import Link from "next/link"
import { Button } from "@/components/ui/button"
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger, DialogFooter } from "@/components/ui/dialog"
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table"
import { Download, Camera, Trash2, Upload, UserPlus } from "lucide-react"
import { Card, CardHeader, CardTitle, CardContent } from "@/components/ui/card"
import { Html5Qrcode } from "html5-qrcode"
import ExcelJS from "exceljs"

export default function Home() {
  // Local UI state and refs
  const [open, setOpen] = useState(false)
  const [attendance, setAttendance] = useState<{ student: string; time: string }[]>([])
  const [presentMales, setPresentMales] = useState<Array<{ student: string; lrn?: string; time?: string }>>([])
  const [presentFemales, setPresentFemales] = useState<Array<{ student: string; lrn?: string; time?: string }>>([])
  const [today, setToday] = useState(() => new Date().toISOString().split("T")[0])
  const scannerRef = useRef<Html5Qrcode | null>(null)
  // Backend base URL. Use NEXT_PUBLIC_API_URL if provided, otherwise default to localhost:8000
  const API_BASE = (process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000").replace(/\/$/, "")

  // Fetch today's attendance from server or fallback to localStorage
  async function fetchTodayAttendance() {
    try {
      const res = await fetch(`${API_BASE}/api/attendance/today/`)
      if (!res.ok) throw new Error("Failed to fetch today attendance")
      const data = await res.json()
      setAttendance(
        (data || []).map((d: any) => ({ student: d.student_name || d.student, time: d.time })) || [],
      )
    } catch (err) {
      // fallback to localStorage recent scans
      try {
        const simple = JSON.parse(localStorage.getItem("attendance_simple") || "[]") as any[]
        // Only load today's scans into attendance state
        const todayOnly = (simple || []).filter((s) => {
          const t = s && s.time ? s.time.toString() : ''
          const datePart = (t.split('T')[0] || t.split(' ')[0] || '').trim()
          return isTodayDate(datePart)
        })
        setAttendance((todayOnly || []).map((s) => ({ student: s.student || s.lrn || JSON.stringify(s), time: s.time || "" })))
      } catch {
        setAttendance([])
      }
    }
  }

  useEffect(() => {
    fetchTodayAttendance()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // Recompute grouped present lists whenever attendance (or localStorage) changes
  useEffect(() => {
    try {
      computeTodayPresent()
    } catch (e) {
      // ignore
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [attendance])

  // Helper: return local YYYY-MM-DD for a Date
  function localYMD(d: Date) {
    const y = d.getFullYear()
    const m = (d.getMonth() + 1).toString().padStart(2, '0')
    const day = d.getDate().toString().padStart(2, '0')
    return `${y}-${m}-${day}`
  }

  // Check whether a date-string or Date refers to today (local)
  function isTodayDate(value: string | Date | undefined) {
    if (!value) return false
    try {
      if (value instanceof Date) {
        return localYMD(value) === localYMD(new Date())
      }
      const s = value.toString()
      // common formats: 'YYYY-MM-DD', 'YYYY-MM-DDTHH:MM:SS', 'YYYY-MM-DD HH:MM:SS'
      const left = s.split('T')[0].split(' ')[0]
      if (/^\d{4}-\d{2}-\d{2}$/.test(left)) return left === localYMD(new Date())
      // fallback to Date parse
      const parsed = new Date(s)
      if (!isNaN(parsed.getTime())) return localYMD(parsed) === localYMD(new Date())
    } catch (e) {
      // ignore
    }
    return false
  }

  // Schedule a daily refresh at local midnight so the table shows only today's scans
  useEffect(() => {
    let t: number | undefined = undefined
    const scheduleNext = () => {
      const now = new Date()
      const tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, 0, 0, 2)
      const ms = tomorrow.getTime() - now.getTime()
      t = window.setTimeout(async () => {
        try {
          await fetchTodayAttendance()
          computeTodayPresent()
        } catch (e) {
          // ignore
        }
        scheduleNext()
      }, ms)
    }
    scheduleNext()
    return () => {
      if (t) clearTimeout(t)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  async function clearAttendance() {
    try {
      const choice = window.prompt("Type 'today' to clear today's attendance, or 'all' to clear all attendance records.", 'today')
      if (!choice) return
      const scope = choice.toLowerCase() === 'all' ? 'all' : 'today'
      const res = await fetch(`${API_BASE}/api/attendance/clear/?scope=${scope}`, { method: 'DELETE' })
      if (!res.ok) {
        const body = await res.text()
        throw new Error(body || 'Failed to clear')
      }

      // Clear in-memory UI
      setAttendance([])

      // Also clear localStorage records so exports do not include stale entries
      try {
        if (scope === 'all') {
          localStorage.removeItem('attendance_simple')
          localStorage.removeItem('attendance_history')
        } else {
          // scope === 'today' -> remove today's entries from attendance_history and attendance_simple
          try {
            const todayKey = new Date().toISOString().split('T')[0]
            // attendance_simple likely contains recent scans with full time strings
            const simpleRaw = localStorage.getItem('attendance_simple')
            if (simpleRaw) {
              const simple = JSON.parse(simpleRaw || '[]') as any[]
              const filtered = (simple || []).filter((s) => {
                const t = (s && s.time) ? (s.time.toString()) : ''
                return !(t.split(' ')[0] === todayKey)
              })
              if (filtered.length) localStorage.setItem('attendance_simple', JSON.stringify(filtered))
              else localStorage.removeItem('attendance_simple')
            }

            const histRaw = localStorage.getItem('attendance_history')
            if (histRaw) {
              const hist = JSON.parse(histRaw || '{}') as Record<string, any[]>
              // Remove the today's key
              delete hist[todayKey]
              // If object is empty, remove the key entirely
              if (Object.keys(hist).length === 0) localStorage.removeItem('attendance_history')
              else localStorage.setItem('attendance_history', JSON.stringify(hist))
            }
          } catch (e) {
            // ignore localStorage trimming errors
          }
        }
      } catch (e) {
        // ignore
      }

      alert(scope === 'all' ? "All attendance cleared." : "Today's attendance cleared.")
    } catch (err) {
      console.error(err)
      alert(`Failed to clear attendance: ${err instanceof Error ? err.message : String(err)}`)
    }
  }

  // Add attendance by sending LRN or student identifier to backend and refresh
  async function addAttendance(lrnOrStudent: string) {
    try {
      let lrn = lrnOrStudent
      const looksLikeName = /[a-zA-Z]/.test(lrnOrStudent) && /\s/.test(lrnOrStudent)
      if (looksLikeName) {
        try {
          const regRes = await fetch(`${API_BASE}/api/registrations/`)
          if (regRes.ok) {
            const regs = await regRes.json()
            const normalized = (lrnOrStudent || "").toString().trim().toLowerCase()
            const found = regs.find((r: any) => (r.student || "").toString().trim().toLowerCase() === normalized)
            if (found && found.lrn) {
              lrn = found.lrn
            } else {
              alert("Scanned name not found in registrations. Please register the student first or include LRN in the QR.")
              return
            }
          }
        } catch (lookupErr) {
          console.error("Lookup by name failed:", lookupErr)
          alert("Could not resolve scanned name to a registered LRN. Ensure QR contains the student LRN or register the student first.")
          return
        }
      }

      const payload = { lrn }
      const res = await fetch(`${API_BASE}/api/attendance/`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      })
      if (!res.ok) {
        let msg = null
        try {
          const body = await res.json()
          msg = body && (body.error || body.detail || body.message)
        } catch {}
        alert(msg || "Failed to record attendance")
        return
      }

      // Refresh list
      await fetchTodayAttendance()
    } catch (err) {
      console.error("Failed to add attendance:", err)
      alert("Failed to add attendance")
    }
  }

  // Compute today's present students grouped by sex, sorted alphabetically.
  async function computeTodayPresent() {
    try {
      // Build a map of present keys -> time for today's scans
      const presentKeyToTime = new Map<string, string>()
      const normalizeName = (s: any) =>
        (s || "")
          .toString()
          .trim()
          .replace(/\s+/g, " ")
          .toLowerCase()

      const normalizeLrn = (s: any) => {
        const str = (s || "").toString()
        const digits = (str.match(/\d+/g) || []).join("")
        return digits || null
      }

      // From already loaded `attendance` state
      (attendance || []).forEach((a) => {
        const nm = normalizeName(a.student)
        if (nm) presentKeyToTime.set(nm, a.time || "")
        const l = normalizeLrn(a.student)
        if (l) presentKeyToTime.set(l, a.time || "")
      })

      // Also include attendance_simple from localStorage for immediate scans, but only today's entries
      try {
        const simpleRaw = localStorage.getItem('attendance_simple')
        if (simpleRaw) {
          const simple = JSON.parse(simpleRaw || '[]') as Array<{ student?: string; time?: string; lrn?: string }>
          simple.forEach((s) => {
            // Determine if this scan is for today (local)
            const t = s && s.time ? s.time.toString() : ''
            // Accept both 'YYYY-MM-DDT..' and 'YYYY-MM-DD ..' shapes
            const datePart = (t.split('T')[0] || t.split(' ')[0] || '').trim()
            if (!datePart || !isTodayDate(datePart)) return
            const keyName = normalizeName(s.student || s.lrn || JSON.stringify(s))
            if (keyName) presentKeyToTime.set(keyName, s.time || '')
            const l = normalizeLrn(s.lrn || s.student)
            if (l) presentKeyToTime.set(l, s.time || '')
          })
        }
      } catch (e) {
        // ignore
      }

      // Fetch registrations from backend (fallback to localStorage)
      let registrations: { student: string; sex: string; lrn: string }[] = []
      try {
        const regRes = await fetch(`${API_BASE}/api/registrations/`)
        if (regRes.ok) registrations = await regRes.json()
        else {
          const regRaw = localStorage.getItem('registrations')
          if (regRaw) registrations = JSON.parse(regRaw)
        }
      } catch (e) {
        try {
          const regRaw = localStorage.getItem('registrations')
          if (regRaw) registrations = JSON.parse(regRaw)
        } catch {}
      }

      const males: Array<{ student: string; lrn?: string; time?: string }> = []
      const females: Array<{ student: string; lrn?: string; time?: string }> = []

      // Build set of matched keys so we can collect unregistered later
      const matchedKeys = new Set<string>()

      registrations.forEach((r) => {
        const nameKey = normalizeName(r.student)
        const lrnKey = normalizeLrn(r.lrn) || ''
        const time = presentKeyToTime.get(nameKey) || presentKeyToTime.get(lrnKey) || ''
        if (time || presentKeyToTime.has(nameKey) || (lrnKey && presentKeyToTime.has(lrnKey))) {
          matchedKeys.add(nameKey)
          if (lrnKey) matchedKeys.add(lrnKey)
          const obj = { student: r.student || r.lrn || 'Unknown', lrn: r.lrn, time }
          if ((r.sex || '').toLowerCase() === 'male') males.push(obj)
          else if ((r.sex || '').toLowerCase() === 'female') females.push(obj)
          else males.push(obj) // default to males if sex missing to still show
        }
      })

      // Sort alphabetically by student name
      const byName = (x: any, y: any) => ('' + (x.student || '')).localeCompare('' + (y.student || ''))
      males.sort(byName)
      females.sort(byName)
      setPresentMales(males)
      setPresentFemales(females)
    } catch (err) {
      console.error('Failed to compute today present:', err)
      alert('Failed to compute today present list')
    }
  }


  function parseStudentFromQr(qrMessage: string): string | null {
    try {
      const obj = JSON.parse(qrMessage)
      return obj.lrn || obj.student || null
    } catch {
      return qrMessage
    }
  }

  // Parse a YYYY-MM-DD date string as a local Date (avoid new Date('YYYY-MM-DD') which is treated as UTC)
  function parseLocalDateFromYMD(dateStr: string): Date {
    const parts = (dateStr || "").split("-").map((p) => Number(p))
    const year = parts[0] || 0
    const month = (parts[1] || 1) - 1
    const day = parts[2] || 1
    return new Date(year, month, day)
  }

  async function startScanner() {
    try {
      const readerElem = document.getElementById("reader")
      if (!readerElem) {
        console.warn("Reader element not found yet.")
        return
      }

      if (!scannerRef.current) {
        scannerRef.current = new Html5Qrcode("reader")
      }

      await scannerRef.current.start(
        { facingMode: "environment" },
        { fps: 10, qrbox: 250 },
        (decoded: string) => {
          const student = parseStudentFromQr(decoded) || "(Unknown)"
          addAttendance(student)
          stopScanner()
          setOpen(false)
        },
        (err: string) => {
          console.warn("QR scan error:", err)
        },
      )
    } catch (e) {
      console.error("Scanner error:", e)
      alert("Could not access camera. You can upload a QR code image instead.")
      stopScanner()
    }
  }

  async function stopScanner() {
    if (scannerRef.current) {
      try {
        await scannerRef.current.stop()
        await scannerRef.current.clear()
      } catch {
        // ignore
      }
    }
  }

  async function handleImageUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return

    try {
      if (scannerRef.current) {
        try {
          await scannerRef.current.stop()
        } catch {
          // ignore if already stopped
        }
      }

      if (!scannerRef.current) {
        scannerRef.current = new Html5Qrcode("reader")
      }

      // Try scanning the file. html5-qrcode's scanFile can return different shapes
      let scanResult: any = null
      try {
        scanResult = await scannerRef.current.scanFile(file, true)
      } catch (scanErr) {
        console.error('scanFile failed:', scanErr)
        alert('Failed to read QR code from image. Check console for details and try a clearer image.')
        setOpen(false)
        return
      }

      // Normalize possible return shapes
      let decodedText: string | null = null
      if (typeof scanResult === 'string') {
        decodedText = scanResult
      } else if (scanResult && typeof scanResult.decodedText === 'string') {
        decodedText = scanResult.decodedText
      } else if (Array.isArray(scanResult) && scanResult.length > 0) {
        // array of results
        const first = scanResult[0]
        decodedText = typeof first === 'string' ? first : first.decodedText || null
      }

      if (!decodedText) {
        console.error('No decoded text returned from scanFile:', scanResult)
        alert('No QR code detected in the selected image. Try a clearer image or take a photo with better lighting.')
        setOpen(false)
        return
      }

      const student = parseStudentFromQr(decodedText) || "(Unknown)"
      addAttendance(student)

      setOpen(false)
    } catch (err) {
      console.error("Image scan failed:", err)
      alert(`Failed to read QR code from image: ${err instanceof Error ? err.message : String(err)}`)
    }
  }
// removed top-level scanner/handler functions; implementations are inside the component so they can
// access state/hooks (scannerRef, setOpen, setAttendance, API_BASE).
// downloadExcel implementation lives inside the component below.
async function downloadExcel(_attendance?: { student: string; time: string }[], includeNames: boolean = false) {
  try {
    const response = await fetch("/attendance_template.xlsx");
    const buffer = await response.arrayBuffer();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];

    const normalizeName = (s: any) =>
      (s || "")
        .toString()
        .trim()
        .replace(/\s+/g, " ")
        .toLowerCase();

    // Stronger key normalizer: strip common punctuation and parentheses so variants like
    // "Name (12345)" or "Name-12345" normalize to the same key. Also produce an LRN-only key when possible.
    const normalizeKey = (s: any) => {
      const str = (s || "").toString();
      const l = str.trim();
      // LRN candidate: digits only
      const digits = (l.match(/\d+/g) || []).join('');
      const lrnKey = digits ? digits : null;
      // Remove parentheses and punctuation, keep letters/numbers and spaces
      let cleaned = l.replace(/[()\[\]{}]|[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
      // Heuristic: only treat cleaned as a matching key if it's reasonably specific
      // (at least 4 chars and contains letters or at least 4 digits). This reduces false positives
      // where very short or generic cleaned strings match multiple students.
      const letters = (cleaned.match(/[a-z]/gi) || []).length;
      const digitsOnly = (cleaned.match(/\d/g) || []).length;
      if (cleaned.length < 4 || (letters === 0 && digitsOnly < 4)) {
        cleaned = '';
      }
      return { raw: l.toLowerCase(), cleaned, lrnKey };
    };

  // 🔹 Load attendance history from localStorage (persistent)
  // Accept arrays of strings or arrays of objects ({ student, lrn })
  let history: Record<string, any[]> = {};
    try {
      history = JSON.parse(localStorage.getItem("attendance_history") || "{}");
    } catch {}

    // 🔹 Build attendanceByStudentMonth from full history
    // Be flexible: history keys may be 'YYYY-MM-DD' or full ISO datetimes; entries may be strings or objects { student, lrn }
    const attendanceByStudentMonth: Record<string, Record<string, Record<number, number>>> = {};
    const AM_FLAG = 1;
    const PM_FLAG = 2;
    const addPresence = (key: string, month: string, day: number, hour: number | null) => {
      if (!key) return;
      if (!attendanceByStudentMonth[key]) attendanceByStudentMonth[key] = {};
      if (!attendanceByStudentMonth[key][month]) attendanceByStudentMonth[key][month] = {};
      const prev = attendanceByStudentMonth[key][month][day] || 0;
      let flag = 0;
      if (hour === null) {
        // unknown time -> mark as both to be safe
        flag = AM_FLAG | PM_FLAG;
      } else if (hour < 12) {
        flag = AM_FLAG;
      } else {
        flag = PM_FLAG;
      }
      attendanceByStudentMonth[key][month][day] = prev | flag;
    };
    Object.keys(history).forEach((dateKey) => {
      // Try to produce a Date for the key.
      let dateObj: Date | null = null;
      if (/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) {
        dateObj = parseLocalDateFromYMD(dateKey);
      } else {
        const parsed = new Date(dateKey);
        if (!isNaN(parsed.getTime())) {
          dateObj = parsed;
        } else {
          // Try taking the left part before 'T'
          const left = (dateKey || '').split('T')[0];
          if (/^\d{4}-\d{2}-\d{2}$/.test(left)) dateObj = parseLocalDateFromYMD(left);
        }
      }
      if (!dateObj) return;
      const month = monthNames[dateObj.getMonth()];
      const day = dateObj.getDate();

      const entries = Array.isArray(history[dateKey]) ? history[dateKey] : [];
      entries.forEach((entry) => {
          let entryHour: number | null = null;
          // Try to extract hour if time provided
          const tval = (entry.time || entry.timestamp || entry.created || '');
          if (tval) {
            try {
              const tstr = tval.toString();
              const isoMatch = tstr.match(/\d{4}-\d{2}-\d{2}T(\d{2}):/);
              if (isoMatch) entryHour = Number(isoMatch[1]);
              else {
                const parsed = new Date(tstr);
                if (!isNaN(parsed.getTime())) entryHour = parsed.getHours();
              }
            } catch {}
          }
        let studentName: string | null = null;
        let entryLrn: string | null = null;
        if (typeof entry === 'string') {
          studentName = entry;
        } else if (entry && typeof entry === 'object') {
          // Accept different shapes
          studentName = (entry.student_name || entry.student || entry.name) || null;
          entryLrn = (entry.lrn || entry.LRN || entry.id) || null;
        }

        const candidates: string[] = [];
        if (studentName) {
          const nk = normalizeKey(studentName);
          if (nk.lrnKey) candidates.push(nk.lrnKey);
          candidates.push(normalizeName(studentName));
        }
        if (entryLrn) {
          const nk = normalizeKey(entryLrn);
          if (nk.lrnKey) candidates.push(nk.lrnKey);
          candidates.push(normalizeName(entryLrn));
        }

        // Also accept raw object-to-string fallback
        if (!studentName && !entryLrn) {
          try {
            const asStr = JSON.stringify(entry);
            if (asStr) {
              // Use only LRN candidate and normalized string. Avoid fuzzy 'cleaned' key to reduce false matches.
              const nk = normalizeKey(asStr);
              if (nk.lrnKey) candidates.push(nk.lrnKey);
              candidates.push(normalizeName(asStr));
            }
          } catch {}
        }

        // Register all candidate keys. Skip generic placeholders like 'unknown student' or empty strings
        candidates.forEach((key) => {
          if (!key) return;
          const lk = (key || '').toString().trim().toLowerCase();
          if (!lk) return;
          if (lk === 'unknown student' || lk === 'unknown' || lk === 'n/a') return;
          addPresence(lk, month, day, entryHour);
        });
      });
    });

    // Also fetch server-side attendance records and merge them (if available)
    try {
      let srvRes = await fetch(`${API_BASE}/api/attendances/`);
      if (!srvRes.ok) srvRes = await fetch(`${API_BASE}/api/attendance/`);
      if (srvRes.ok) {
        const srv = await srvRes.json();
        // srv likely array of { id, student, student_name, time }
        (srv || []).forEach((rec: any) => {
          try {
            const time = rec.time || rec.created || rec.timestamp || null;
            if (!time) return;
            const timeStr = (time || '').toString();
            // Try to extract ISO date (YYYY-MM-DD) first
            let dateObj: Date | null = null;
            const isoMatch = timeStr.match(/\d{4}-\d{2}-\d{2}/);
            if (isoMatch) {
              dateObj = parseLocalDateFromYMD(isoMatch[0]);
            } else {
              // Fallback to Date parsing of common formats (e.g., 'Oct. 20, 2025, 5:08 p.m.')
              const parsed = new Date(timeStr);
              if (!isNaN(parsed.getTime())) dateObj = parsed;
            }
            if (!dateObj) return;
            const month = monthNames[dateObj.getMonth()];
            const day = dateObj.getDate();
            const candidates: string[] = [];
            const studentName = rec.student_name || rec.student || null;
            const studentLrn = rec.student || rec.lrn || null;
            if (studentName) {
              const nk = normalizeKey(studentName);
              if (nk.lrnKey) candidates.push(nk.lrnKey);
              candidates.push(normalizeName(studentName));
            }
            if (studentLrn) {
              const nk = normalizeKey(studentLrn);
              if (nk.lrnKey) candidates.push(nk.lrnKey);
              candidates.push(normalizeName(studentLrn));
            }
            // Determine hour for AM/PM
            let recHour: number | null = null;
            try {
              const isoMatch2 = (time || '').toString().match(/\d{4}-\d{2}-\d{2}T(\d{2}):/);
              if (isoMatch2) recHour = Number(isoMatch2[1]);
              else {
                const p = new Date(time || '');
                if (!isNaN(p.getTime())) recHour = p.getHours();
              }
            } catch {}
            candidates.forEach((rawKey) => {
              if (!rawKey) return;
              const key = (rawKey || '').toString().trim().toLowerCase();
              if (!key) return;
              if (key === 'unknown student' || key === 'unknown' || key === 'n/a') return;
              addPresence(key, month, day, recHour);
            });
          } catch {}
        });
      }
    } catch (e) {
      // ignore server fetch failures
    }

    // Also merge today's recent scans from attendance_simple so immediate scans are recognized
    try {
      const simpleRaw = localStorage.getItem('attendance_simple');
      if (simpleRaw) {
        const simple = JSON.parse(simpleRaw || '[]') as Array<{ student?: string; time?: string; lrn?: string }>;
        simple.forEach((s) => {
          try {
            const t = (s && s.time) ? (s.time.toString()) : '';
            const datePart = t.split(' ')[0] || new Date().toISOString().split('T')[0];
            const dateObj = parseLocalDateFromYMD(datePart);
            const month = monthNames[dateObj.getMonth()];
            const day = dateObj.getDate();
            const candidates: string[] = [];
            if (s.student) {
              const nk = normalizeKey(s.student);
              if (nk.lrnKey) candidates.push(nk.lrnKey);
              candidates.push(normalizeName(s.student));
            }
            if (s.lrn) {
              const nk = normalizeKey(s.lrn);
              if (nk.lrnKey) candidates.push(nk.lrnKey);
              candidates.push(normalizeName(s.lrn));
            }
            if (candidates.length === 0) {
              try {
                const asStr = JSON.stringify(s);
                const nk = normalizeKey(asStr);
                if (nk.lrnKey) candidates.push(nk.lrnKey);
                candidates.push(normalizeName(asStr));
              } catch {}
            }
            // Determine hour from s.time
            let sHour: number | null = null;
            try {
              const t = (s && s.time) ? s.time.toString() : '';
              const m = t.match(/\d{4}-\d{2}-\d{2}T(\d{2}):/);
              if (m) sHour = Number(m[1]);
              else {
                const p = new Date(t || '');
                if (!isNaN(p.getTime())) sHour = p.getHours();
              }
            } catch {}
            candidates.forEach((rawKey) => {
              const key = (rawKey || '').toString().trim().toLowerCase();
              if (!key) return;
              if (key === 'unknown student' || key === 'unknown' || key === 'n/a') return;
              addPresence(key, month, day, sHour);
            });
          } catch {}
        });
      }
    } catch (e) {
      // ignore
    }

    // 🔹 Get registration data from backend, fallback to localStorage
    let registrations: { student: string; sex: string; lrn: string; parent: string; guardian: string }[] = [];
    try {
      const regRes = await fetch(`${API_BASE}/api/registrations/`);
      if (regRes.ok) {
        registrations = await regRes.json();
      } else {
        // fallback to localStorage
        const regRaw = localStorage.getItem("registrations");
        if (regRaw) registrations = JSON.parse(regRaw);
      }
    } catch (e) {
      try {
        const regRaw = localStorage.getItem("registrations");
        if (regRaw) registrations = JSON.parse(regRaw);
      } catch {}
    }

    // 🔹 For each month, update worksheet
    monthNames.forEach((month) => {
      const worksheet = workbook.getWorksheet(month);
      if (!worksheet) return;

      // Helper: normalize a header cell value to a day number (1..31) when possible.
      // Accept numbers, numeric strings ("1", "01"), ordinal strings ("1st"), Date objects,
      // and ExcelJS richText/text/result shapes.
      const getCellDay = (cell: any): number | null => {
        if (!cell) return null;
        const v = cell.value;
        // direct number
        if (typeof v === 'number') return v;

        // plain string (accept "1", "01", "1st", " 1 ")
        if (typeof v === 'string') {
          const s = v.trim();
          const m = s.match(/^(\d{1,2})/);
          if (m) return parseInt(m[1], 10);
          return null;
        }

        // Date object
        if (v instanceof Date) return v.getDate();

        // ExcelJS cell objects (richText/text/formula=result)
        if (v && typeof v === 'object') {
          try {
            // richText (array of runs)
            if ((v as any).richText && Array.isArray((v as any).richText)) {
              const joined = (v as any).richText.map((r: any) => r && r.text ? r.text : '').join('').trim();
              const m = joined.match(/^(\d{1,2})/);
              if (m) return parseInt(m[1], 10);
            }

            // text property
            if ((v as any).text) {
              const s = (v as any).text.toString().trim();
              const m = s.match(/^(\d{1,2})/);
              if (m) return parseInt(m[1], 10);
            }

            // formula/result shapes
            if ((v as any).result !== undefined) {
              const rv = (v as any).result;
              if (typeof rv === 'number') return rv;
              if (typeof rv === 'string') {
                const s = rv.trim();
                const m = s.match(/^(\d{1,2})/);
                if (m) return parseInt(m[1], 10);
              }
              if (rv instanceof Date) return rv.getDate();
            }
          } catch (e) {
            // ignore parse errors
          }
        }
        return null;
      };

  const males = registrations.filter(r => r.sex === "Male");
  const females = registrations.filter(r => r.sex === "Female");

      // Pre-clear merges and fills on date columns for male/female ranges to avoid template overrides
      const dateHeaderRow = worksheet.getRow(10);
      const dateCols: number[] = [];
      dateHeaderRow.eachCell((cell, colNumber) => {
        const d = getCellDay(cell);
        if (d != null) dateCols.push(colNumber);
      });
      // Clear merges/fills for a reasonable block of rows where students exist
      const clearRows = [13, 13 + Math.max(200, males.length), 64, 64 + Math.max(200, females.length)];
      for (const col of dateCols) {
        for (let r = 13; r < 13 + Math.max(200, males.length); r++) {
          try {
            const c = worksheet.getRow(r).getCell(col);
            if (c.master) c.unmerge?.();
            // Preserve borders from the template so cell gridlines remain visible
            (c as any).fill = undefined;
            // do not clear .border here
          } catch {}
        }
        for (let r = 64; r < 64 + Math.max(200, females.length); r++) {
          try {
            const c = worksheet.getRow(r).getCell(col);
            if (c.master) c.unmerge?.();
            (c as any).fill = undefined;
            // do not clear .border here to preserve template cell lines
          } catch {}
        }
      }

      // No green shading: present cells will be left blank (no image overlay)
  // Prepare triangular overlays for AM (top-right) and PM (bottom-left)
  let amImageId: number | null = null;
  let pmImageId: number | null = null;

      // Prepare a dashed black diagonal PNG to overlay each attendance cell (ensures visibility over template)
      let diagImageId: number | null = null;
      try {
        const dCanvas = document.createElement('canvas');
        // Size chosen to roughly match cell drawing area; adjust if needed
        dCanvas.width = 51;
        dCanvas.height = 43;
        const dCtx = dCanvas.getContext('2d');
        if (dCtx) {
          // Transparent background
          dCtx.clearRect(0, 0, dCanvas.width, dCanvas.height);
          dCtx.strokeStyle = '#000000';
          dCtx.lineWidth = 2;
          dCtx.setLineDash([6, 4]);
          dCtx.beginPath();
          dCtx.moveTo(3, dCanvas.height - 3);
          dCtx.lineTo(dCanvas.width - 3, 3);
          dCtx.stroke();

          const dataUrl = dCanvas.toDataURL('image/png');
          const base64 = dataUrl.split(',')[1];
          const binary = atob(base64);
          const len = binary.length;
          const bytes = new Uint8Array(len);
          for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
          diagImageId = workbook.addImage({ buffer: bytes.buffer, extension: 'png' });
        }
      } catch (e) {
        console.warn('Failed to create diagonal image for Excel markers:', e);
        diagImageId = null;
      }

  // ---- Males (row 13)
      let rowIdx = 13;
      males.forEach((reg) => {
        // Derive a stable visible label, but compute matching keys from raw fields to avoid placeholder matches
        const rawStudent = (reg.student && reg.student.toString().trim()) || '';
        const rawLrn = (reg.lrn && reg.lrn.toString().trim()) || '';
        const visibleName = rawStudent || rawLrn || 'Unknown Student';
        const nameCell = worksheet.getCell(`B${rowIdx}`);
        if (includeNames) nameCell.value = visibleName;
  // Force a readable font color and alignment so the name is visible on any template
  (nameCell as any).font = { color: { argb: 'FF000000' }, size: 10 };
        nameCell.alignment = { vertical: 'middle', horizontal: 'left' } as any;
        worksheet.getRow(10).eachCell((cell, colNumber) => {
          const day = getCellDay(cell);
          if (day != null) {
            // Process all months: mark absent/present for every worksheet.
            // For the current month, don't touch future days beyond today.
            const now = new Date();
            const currentMonthName = monthNames[now.getMonth()];
            const todayDay = now.getDate();
            if (month === currentMonthName && day > todayDay) {
              return
            }
            const markCell = worksheet.getRow(rowIdx).getCell(colNumber);

            const nameKey = rawStudent ? normalizeName(rawStudent) : '';
            const keyInfo = normalizeKey(rawLrn || rawStudent || '');
            const lrnKey = keyInfo.lrnKey;
            const studentMonthDates = new Set<number>();
            // Prefer LRN-only match, then cleaned combined key (only if it contains name or LRN), then raw normalized name
            if (lrnKey && attendanceByStudentMonth[lrnKey] && attendanceByStudentMonth[lrnKey][month]) {
              Object.keys(attendanceByStudentMonth[lrnKey][month]).forEach((k) => {
                const num = Number(k);
                if (!isNaN(num)) studentMonthDates.add(num);
              });
            }
            if (nameKey && attendanceByStudentMonth[nameKey] && attendanceByStudentMonth[nameKey][month]) {
              Object.keys(attendanceByStudentMonth[nameKey][month]).forEach((k) => {
                const num = Number(k);
                if (!isNaN(num)) studentMonthDates.add(num);
              });
            }

            // Determine present condition for today (skip accidental marking of day 1)
            const presentCondition = studentMonthDates.has(day) && day !== 1;
            // Apply fill and font first
            if (presentCondition) {
              // ✅ Present: leave blank (no marker) but ensure alignment and clear fill/font
              try {
                markCell.value = "";
                markCell.alignment = { vertical: 'middle', horizontal: 'center' } as any;
                (markCell as any).font = undefined;
                (markCell as any).fill = undefined;
              } catch {}
            } else {
              // ❌ Absent: mark X
              markCell.value = "X";
              markCell.alignment = { vertical: 'middle', horizontal: 'center' } as any;
              markCell.font = { color: { argb: "FF000000" }, bold: false };
              markCell.fill = { type: "pattern", pattern: "none" };
            }

            // Now construct the border (diagonal) after fill is applied so shading remains visible.
            // If the cell is shaded/present, force the diagonal to white so it remains visible over the green marker.
            const diagonalBorder: any = {
              up: true,
              down: false,
              style: "dashed",
              color: { argb: "FF000000" }
            };

            (markCell as any).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
              diagonal: diagonalBorder
            };
          }
        });
        rowIdx++;
      });

  // ---- Females (row 64)
      rowIdx = 64;
      females.forEach((reg) => {
        // Derive visible label but compute matching keys from raw fields to avoid placeholder matches
        const rawStudent = (reg.student && reg.student.toString().trim()) || '';
        const rawLrn = (reg.lrn && reg.lrn.toString().trim()) || '';
        const visibleName = rawStudent || rawLrn || 'Unknown Student';
        const nameCell = worksheet.getCell(`B${rowIdx}`);
        if (includeNames) nameCell.value = visibleName;
        (nameCell as any).font = { color: { argb: 'FF000000' }, size: 10 };
        nameCell.alignment = { vertical: 'middle', horizontal: 'left' } as any;
        worksheet.getRow(10).eachCell((cell, colNumber) => {
          const day = getCellDay(cell);
          if (day != null) {
            // Process all months: mark absent/present for every worksheet.
            // For the current month, skip future days beyond today.
            const now = new Date();
            const currentMonthName = monthNames[now.getMonth()];
            const todayDay = now.getDate();
            if (month === currentMonthName && day > todayDay) {
              return
            }
            const markCell = worksheet.getRow(rowIdx).getCell(colNumber);

            const nameKey = rawStudent ? normalizeName(rawStudent) : '';
            const keyInfoF = normalizeKey(rawLrn || rawStudent || '');
            const lrnKeyF = keyInfoF.lrnKey;
            const studentMonthDates = new Set<number>();
            if (lrnKeyF && attendanceByStudentMonth[lrnKeyF] && attendanceByStudentMonth[lrnKeyF][month]) {
              Object.keys(attendanceByStudentMonth[lrnKeyF][month]).forEach((k) => {
                const num = Number(k);
                if (!isNaN(num)) studentMonthDates.add(num);
              });
            }
            if (nameKey && attendanceByStudentMonth[nameKey] && attendanceByStudentMonth[nameKey][month]) {
              Object.keys(attendanceByStudentMonth[nameKey][month]).forEach((k) => {
                const num = Number(k);
                if (!isNaN(num)) studentMonthDates.add(num);
              });
            }

            // Determine present condition for today (skip accidental marking of day 1)
            const presentCondition = studentMonthDates.has(day) && day !== 1;
            // Apply fill and font first
            if (presentCondition) {
              // ✅ Present: leave blank (no marker) but ensure alignment and clear fill/font
              try {
                markCell.value = "";
                markCell.alignment = { vertical: 'middle', horizontal: 'center' } as any;
                (markCell as any).font = undefined;
                (markCell as any).fill = undefined;
              } catch {}
            } else {
              markCell.value = "X";
              markCell.alignment = { vertical: 'middle', horizontal: 'center' } as any;
              markCell.font = { color: { argb: "FF000000" }, bold: false };
              markCell.fill = { type: "pattern", pattern: "none" };
            }

            // If the cell is shaded/present, force the diagonal to white so it remains visible over the green marker.
            const diagonalBorder: any = {
              up: true,
              down: false,
              style: "dashed",
              color: { argb: "FF000000" }
            };
            (markCell as any).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
              diagonal: diagonalBorder
            };
          }
        });
        rowIdx++;
      });

      // Ensure every attendance cell in the specified template date range has a dashed diagonal (preserve other borders)
      try {
        const diagDef = { up: true, down: false, style: 'dashed', color: { argb: 'FF000000' } };
        // Columns G..AE => 7..31 (A=1)
        const colStart = 7;
        const colEnd = 31;
        // Male rows: 13..62 (inclusive start, exclusive end below uses <= so include 62)
        const maleStart = 13;
        const maleEnd = 62;
        // Female rows: 64..113
        const femaleStart = 64;
        const femaleEnd = 113;

        for (let col = colStart; col <= colEnd; col++) {
          for (let r = maleStart; r <= maleEnd; r++) {
            try {
              const cell = worksheet.getRow(r).getCell(col) as any;
              const existing = cell.border || {};
              cell.border = { ...existing, diagonal: diagDef };
            } catch {}
          }
          for (let r = femaleStart; r <= femaleEnd; r++) {
            try {
              const cell = worksheet.getRow(r).getCell(col) as any;
              const existing = cell.border || {};
              cell.border = { ...existing, diagonal: diagDef };
            } catch {}
          }
        }
      } catch (e) {
        // ignore
      }
    });

    // 🔹 Download Excel
    // 🔹 Add a 'Scans' worksheet summarizing raw scan records for traceability
    try {
      const scansSheet = workbook.addWorksheet('Scans');
      scansSheet.columns = [
        { header: 'Date', key: 'date', width: 15 },
        { header: 'Student/LRN', key: 'student', width: 30 },
        { header: 'LRN', key: 'lrn', width: 20 },
        { header: 'Time', key: 'time', width: 20 },
        { header: 'Source', key: 'source', width: 20 }
      ];

      // attendance_simple: array of { student, time }
      try {
        const simpleRaw = localStorage.getItem('attendance_simple');
        if (simpleRaw) {
          const simple = JSON.parse(simpleRaw || '[]');
          (simple || []).forEach((s: any) => {
            let date = '';
            let time = '';
            if (s && s.time) {
              const t = s.time.toString();
              date = t.split(' ')[0];
              time = t.split(' ').slice(1).join(' ');
            }
            const studentVal = (s && (s.student || s.lrn)) ? (s.student || s.lrn) : JSON.stringify(s);
            scansSheet.addRow({ date, student: studentVal, lrn: s?.lrn || '', time, source: 'attendance_simple' });
          });
        }
      } catch (e) { /* ignore */ }

      // attendance_history: { 'YYYY-MM-DD'|'iso': [strings|objects] }
      try {
        const histRaw = localStorage.getItem('attendance_history');
        if (histRaw) {
          const hist = JSON.parse(histRaw || '{}') as Record<string, any[]>;
          Object.keys(hist).forEach((k) => {
            const entries = Array.isArray(hist[k]) ? hist[k] : [];
            entries.forEach((ent) => {
              let student = '';
              let lrn = '';
              if (typeof ent === 'string') student = ent;
              else if (ent && typeof ent === 'object') {
                student = ent.student || ent.student_name || JSON.stringify(ent);
                lrn = ent.lrn || ent.LRN || '';
              }
              // Normalize key to date portion
              const datePart = (k || '').split('T')[0];
              scansSheet.addRow({ date: datePart, student, lrn, time: '', source: 'attendance_history' });
            });
          });
        }
      } catch (e) { /* ignore */ }
      // server attendance records
      try {
        const srvRes = await fetch(`${API_BASE}/api/attendance/`);
        if (srvRes.ok) {
          const srv = await srvRes.json();
          (srv || []).forEach((rec: any) => {
            try {
              const time = rec.time || rec.created || rec.timestamp || '';
              const date = (time || '').toString().split('T')[0] || '';
              const studentVal = rec.student_name || rec.student || '';
              const lrnVal = rec.student || rec.lrn || '';
              scansSheet.addRow({ date, student: studentVal, lrn: lrnVal, time, source: 'server' });
            } catch {}
          });
        }
      } catch (e) {}
    } catch (e) {
      console.warn('Failed to add Scans worksheet', e);
    }

    const blob = await workbook.xlsx.writeBuffer();

    // Try to upload to backend first. If backend is not available, fall back to client download.
    const uploadUrl = `${API_BASE}/api/attendance/upload_excel/`;
    try {
      const form = new FormData();
      const fileName = `attendance_export_${new Date().toISOString().replace(/[:.]/g, '-')}.xlsx`;
      form.append('file', new Blob([blob], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName);

      const upRes = await fetch(uploadUrl, {
        method: 'POST',
        body: form,
      });

      if (upRes.ok) {
  const body = await upRes.json();
  // Server copy saved; trigger client download (no alert)
        const url = URL.createObjectURL(new Blob([blob]));
        const link = document.createElement("a");
        link.href = url;
        link.download = fileName || `attendance_export.xlsx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        // If the server returned a URL, show a small non-blocking banner with a clickable link
        if (body && body.url) {
          try {
            const full = body.url.startsWith('http') ? body.url : `${API_BASE.replace(/\/$/, '')}${body.url}`;

            // Remove existing banner if present
            const existing = document.getElementById('export-saved-banner');
            if (existing) existing.remove();

            const banner = document.createElement('div');
            banner.id = 'export-saved-banner';
            banner.style.position = 'fixed';
            banner.style.right = '16px';
            banner.style.bottom = '16px';
            banner.style.background = 'rgba(0,0,0,0.8)';
            banner.style.color = 'white';
            banner.style.padding = '10px 14px';
            banner.style.borderRadius = '8px';
            banner.style.boxShadow = '0 6px 18px rgba(0,0,0,0.3)';
            banner.style.zIndex = '9999';
            banner.style.fontSize = '13px';

            const text = document.createElement('span');
            text.textContent = `Export saved to server:`;
            text.style.marginRight = '8px';

            const a = document.createElement('a');
            a.href = full;
            a.textContent = body.filename || 'Download';
            a.style.color = '#9AE6B4';
            a.style.textDecoration = 'underline';
            a.target = '_blank';

            const closeBtn = document.createElement('button');
            closeBtn.textContent = '✕';
            closeBtn.style.marginLeft = '12px';
            closeBtn.style.background = 'transparent';
            closeBtn.style.color = 'white';
            closeBtn.style.border = 'none';
            closeBtn.style.cursor = 'pointer';

            closeBtn.onclick = () => banner.remove();

            banner.appendChild(text);
            banner.appendChild(a);
            banner.appendChild(closeBtn);
            document.body.appendChild(banner);

            // Auto-hide after 15s
            setTimeout(() => {
              try { banner.remove(); } catch {}
            }, 15000);
          } catch (e) {
            console.warn('Failed to show server download link', e);
          }
        }
      } else {
        console.warn('Upload failed, falling back to client download', upRes.status, upRes.statusText);
        // fallback to download
        const url = URL.createObjectURL(new Blob([blob]));
        const link = document.createElement("a");
        link.href = url;
        link.download = `attendance_export.xlsx`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }
    } catch (e) {
      console.warn('Upload attempt threw, falling back to client download', e);
      const url = URL.createObjectURL(new Blob([blob]));
      const link = document.createElement("a");
      link.href = url;
      link.download = `attendance_export.xlsx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }

  } catch (err) {
    console.error("Excel export error:", err);
    alert("Failed to generate Excel file.");
  }
}
// PDF export removed per user request

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-slate-900 to-slate-800 p-6 text-white">
      <Card className="max-w-5xl w-full border-yellow-500 border-2 bg-white/10 backdrop-blur-lg text-white">
        <CardHeader className="flex flex-col sm:flex-row sm:items-center sm:justify-between">
          <CardTitle className="text-center text-yellow-400 text-2xl">Student Attendance Dashboard</CardTitle>

          <Link href="/" passHref>
            <Button variant="secondary" className="mt-2 sm:mt-0">
              <UserPlus className="h-4 w-4 mr-2" />
              Register Student
            </Button>
          </Link>
        </CardHeader>

        <CardContent className="space-y-6">
          <div className="flex items-center justify-between">
            <p className="text-muted-foreground">Scan a QR code to log attendance instantly.</p>
            <div className="flex gap-2">
              <Dialog
                open={open}
                onOpenChange={(val) => {
                  setOpen(val)
                  if (val) {
                    setTimeout(() => startScanner(), 300)
                  } else {
                    stopScanner()
                  }
                }}
              >
                <DialogTrigger asChild>
                  <Button variant="default">
                    <Camera className="h-4 w-4 mr-2" />
                    Scan QR
                  </Button>
                </DialogTrigger>
                <DialogContent className="max-w-md bg-white/10 backdrop-blur-lg border border-yellow-400/50">
                  <DialogHeader>
                    <DialogTitle className="text-yellow-400">Scan QR Code</DialogTitle>
                  </DialogHeader>
                  <div
                    id="reader"
                    className="w-full h-64 bg-black/30 rounded-md flex items-center justify-center text-sm text-gray-400"
                  >
                    Camera feed
                  </div>

                  <div className="mt-10">
                    <label className="flex items-center gap-2 cursor-pointer text-yellow-400 hover:underline">
                      <Upload className="h-4 w-4" />
                      Upload QR Image
                      <input type="file" accept="image/*" onChange={handleImageUpload} className="hidden" />
                    </label>
                  </div>

                  <DialogFooter>
                    <Button variant="destructive" onClick={stopScanner}>
                      Stop
                    </Button>
                  </DialogFooter>
                </DialogContent>
              </Dialog>

              <Button variant="secondary" onClick={() => downloadExcel(attendance, true)}>
                <Download className="h-4 w-4 mr-2" />
                Download Excel
              </Button>



              <Button variant="destructive" onClick={clearAttendance}>
                <Trash2 className="h-4 w-4 mr-2" />
                Clear
              </Button>
            </div>
          </div>

          <div className="mt-6 grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <h3 className="text-yellow-400 font-semibold flex items-center gap-3">Male <span className="text-sm bg-yellow-400/20 text-yellow-300 px-2 py-1 rounded">{presentMales.length}</span></h3>
              <div className="overflow-x-auto mt-2">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead className="text-yellow-400">Name</TableHead>
                      <TableHead className="text-yellow-400">LRN</TableHead>
                      <TableHead className="text-yellow-400">Time</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {presentMales.length === 0 ? (
                      <TableRow>
                        <TableCell colSpan={3} className="text-muted-foreground">No present male students</TableCell>
                      </TableRow>
                    ) : (
                      presentMales.map((p, i) => (
                        <TableRow key={`m-${i}`}>
                          <TableCell>{p.student}</TableCell>
                          <TableCell>{p.lrn || ''}</TableCell>
                          <TableCell>{p.time || ''}</TableCell>
                        </TableRow>
                      ))
                    )}
                  </TableBody>
                </Table>
              </div>
            </div>

            <div>
              <h3 className="text-yellow-400 font-semibold flex items-center gap-3">Female <span className="text-sm bg-yellow-400/20 text-yellow-300 px-2 py-1 rounded">{presentFemales.length}</span></h3>
              <div className="overflow-x-auto mt-2">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead className="text-yellow-400">Name</TableHead>
                      <TableHead className="text-yellow-400">LRN</TableHead>
                      <TableHead className="text-yellow-400">Time</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {presentFemales.length === 0 ? (
                      <TableRow>
                        <TableCell colSpan={3} className="text-muted-foreground">No present female students</TableCell>
                      </TableRow>
                    ) : (
                      presentFemales.map((p, i) => (
                        <TableRow key={`f-${i}`}>
                          <TableCell>{p.student}</TableCell>
                          <TableCell>{p.lrn || ''}</TableCell>
                          <TableCell>{p.time || ''}</TableCell>
                        </TableRow>
                      ))
                    )}
                  </TableBody>
                </Table>
              </div>
            </div>

            {/* Unregistered column removed per request */}
          </div>
        </CardContent>
      </Card>
    </div>
  )
}
