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
import jsPDF from "jspdf"


export default function Home() {
  const [open, setOpen] = useState(false)
  const [attendance, setAttendance] = useState<{ student: string; time: string }[]>([])
  const [today, setToday] = useState(() => {
    const n = new Date();
    return n.toISOString().split("T")[0];
  });
  const scannerRef = useRef<Html5Qrcode | null>(null)

  // Backend base URL. Use NEXT_PUBLIC_API_URL if provided, otherwise default to localhost:8000
  const API_BASE = (process.env.NEXT_PUBLIC_API_URL || 'http://localhost:8000').replace(/\/$/, '');


  // Load attendance for today only

  async function fetchTodayAttendance() {
    try {
    const res = await fetch(`${API_BASE}/api/attendance/today/`);
      if (!res.ok) throw new Error('Failed to fetch today attendance');
      const data = await res.json();
      // data: [{ id, student, student_name, time }]
      setAttendance(data.map((d: any) => ({ student: d.student_name || d.student, time: d.time })));
    } catch (err) {
      console.warn('Failed to fetch today attendance, falling back to localStorage', err);
      // fallback to previous localStorage method
      const raw = localStorage.getItem('attendance_simple');
      if (raw) {
        try {
          const all = JSON.parse(raw) || [];
          const todayStr = new Date().toISOString().split('T')[0];
          setAttendance(all.filter((a: { time: string }) => (a.time || '').slice(0, 10) === todayStr));
        } catch {
          setAttendance([]);
        }
      } else {
        setAttendance([]);
      }
    }
  }

  useEffect(() => {
    fetchTodayAttendance();

    // Set up timer to refresh at midnight
    const now = new Date();
    const msToMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, 0, 0, 1).getTime() - now.getTime();
    const midnightTimeout = setTimeout(() => {
      setToday(new Date().toISOString().split("T")[0]);
      fetchTodayAttendance();
    }, msToMidnight);

    // Also listen for storage changes (in case another tab adds attendance)
    const onStorage = (e: StorageEvent) => {
      if (e.key === 'attendance_simple') fetchTodayAttendance();
    };
    window.addEventListener("storage", onStorage);

    return () => {
      clearTimeout(midnightTimeout);
      window.removeEventListener("storage", onStorage);
    };
  }, [today]);

  // Save attendance (all records, not just today)
  useEffect(() => {
    // Only update if attendance is not empty and the latest record is for today
    if (attendance.length > 0) {
      const raw = localStorage.getItem("attendance_simple");
      let all = [];
      try { all = JSON.parse(raw || "[]"); } catch {}
      // Merge today's attendance with previous days
      const todayStr = new Date().toISOString().split("T")[0];
      const notToday = all.filter((a: { time: string }) => (a.time || "").slice(0, 10) !== todayStr);
      localStorage.setItem("attendance_simple", JSON.stringify([...notToday, ...attendance]));
    }
  }, [attendance]);

  function nowTime(): string {
  const n = new Date()
  let h = n.getHours()
  const m = String(n.getMinutes()).padStart(2, "0")
  const s = String(n.getSeconds()).padStart(2, "0")
  const ampm = h >= 12 ? "PM" : "AM"
  h = h % 12 || 12
  // Use local date instead of UTC
  const year = n.getFullYear()
  const month = String(n.getMonth() + 1).padStart(2, "0")
  const day = String(n.getDate()).padStart(2, "0")
  const date = `${year}-${month}-${day}`
  return `${date} ${String(h).padStart(2, "0")}:${m}:${s} ${ampm}`
  }

//   function addAttendance(student: string) {
//     const now = new Date();
//     const dateStr = now.toISOString().split("T")[0]; // YYYY-MM-DD only
//     setAttendance((prev) => {
//       // Only add if not already present for today
//       const todayStr = dateStr;
//       const already = prev.some((a) => a.student === student && (a.time || "").slice(0, 10) === todayStr);
//       if (already) return prev;
//       const updated = [...prev, { student, time: nowTime() }];
//       // Do not update localStorage here, let useEffect handle it
//       return updated;
//     });

//     // Update persistent history (by date)
//     let history: Record<string, string[]> = {};
//     try {
//       history = JSON.parse(localStorage.getItem("attendance_history") || "{}");
//     } catch {}
//     if (!history[dateStr]) history[dateStr] = [];
//     if (!history[dateStr].includes(student)) {
//       history[dateStr].push(student);
//     }
//     localStorage.setItem("attendance_history", JSON.stringify(history));
//   }

// The backend expects { lrn } in POST /api/attendance/
async function addAttendance(lrnOrStudent: string) {
  try {
    let lrn = lrnOrStudent;

    // Heuristic: if the scanned value looks like a name (contains letters/spaces) instead of an LRN,
    // try to find the registration by student name.
    const looksLikeName = /[a-zA-Z]/.test(lrnOrStudent) && /\s/.test(lrnOrStudent);
    if (looksLikeName) {
      try {
        const regRes = await fetch(`${API_BASE}/api/registrations/`);
        if (regRes.ok) {
          const regs = await regRes.json();
          const normalized = (lrnOrStudent || '').toString().trim().toLowerCase();
          const found = regs.find((r: any) => (r.student || '').toString().trim().toLowerCase() === normalized);
          if (found && found.lrn) {
            lrn = found.lrn;
          } else {
            throw new Error('No registration found for scanned name.');
          }
        } else {
          throw new Error('Failed to fetch registrations for name lookup');
        }
      } catch (lookupErr) {
        console.error('Lookup by name failed:', lookupErr);
        alert('Could not resolve scanned name to a registered LRN. Ensure QR contains the student LRN or register the student first.');
        return;
      }
    }

    const payload: Record<string, string> = { lrn };
    const res = await fetch(`${API_BASE}/api/attendance/`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      let msg: string | null = null;
      try {
        const body = await res.json();
        if (body && (body.error || body.detail || body.message)) {
          msg = body.error || body.detail || body.message || JSON.stringify(body);
        } else {
          msg = JSON.stringify(body);
        }
      } catch (_) {
        try {
          msg = await res.text();
        } catch (__)
        {
          msg = null;
        }
      }
      console.error('Attendance POST failed', { status: res.status, statusText: res.statusText, message: msg });
      throw new Error(msg || 'Failed to add attendance');
    }
    // success: try to persist to local attendance_history so exports see the new entry even if backend/listing lags
    try {
      let body: any = null;
      try { body = await res.json(); } catch (_) { body = null; }

      const returnedName = (body && (body.student_name || body.student)) ? (body.student_name || body.student) : null;
      let studentToRecord = returnedName || lrn;

      // If we only have an LRN, try to resolve it to the registered student name so exported sheets match registrations
      if (!returnedName && studentToRecord) {
        try {
          let regs: any[] = [];
          try {
            const regRes = await fetch(`${API_BASE}/api/registrations/`);
            if (regRes.ok) regs = await regRes.json();
            else {
              const regRaw = localStorage.getItem('registrations');
              if (regRaw) regs = JSON.parse(regRaw);
            }
          } catch (e) {
            const regRaw = localStorage.getItem('registrations');
            if (regRaw) regs = JSON.parse(regRaw);
          }
          if (Array.isArray(regs) && regs.length) {
            const found = regs.find(r => (r.lrn || '').toString() === (lrn || '').toString());
            if (found && found.student) studentToRecord = found.student;
          }
        } catch (e) {
          // ignore lookup failures
        }
      }

      const now = new Date();
      const dateStr = now.toISOString().split('T')[0];
      let history: Record<string, string[]> = {};
      try { history = JSON.parse(localStorage.getItem('attendance_history') || '{}'); } catch {}
      if (!history[dateStr]) history[dateStr] = [];
      if (!history[dateStr].includes(studentToRecord)) {
        history[dateStr].push(studentToRecord);
        try { localStorage.setItem('attendance_history', JSON.stringify(history)); } catch (e) { console.warn('Failed to persist attendance_history locally', e); }
      }
    } catch (e) {
      console.warn('Failed to update local attendance history after POST', e);
    }

    // success: re-fetch today's attendance
    await fetchTodayAttendance();
  } catch (err: any) {
    console.error('Add attendance error:', err);
    const message = (err && err.message) ? err.message : String(err);
    alert(`Failed to log attendance: ${message}`);
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

//   function clearAttendance() {
//     setAttendance([])
//     localStorage.removeItem("attendance_simple")
//   }
async function clearAttendance() {
  try {
    const res = await fetch(`${API_BASE}/api/registrations/clear_all/`, { method: 'DELETE' });
    if (!res.ok) throw new Error('Failed to clear');
    setAttendance([]);
  } catch (err) {
    console.error(err);
    alert('Failed to clear attendance');
  }
}


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

  // 🔹 Load attendance history from localStorage (persistent)
  // Accept arrays of strings or arrays of objects ({ student, lrn })
  let history: Record<string, any[]> = {};
    try {
      history = JSON.parse(localStorage.getItem("attendance_history") || "{}");
    } catch {}

    // 🔹 Build attendanceByStudentMonth from full history
    // Be flexible: history keys may be 'YYYY-MM-DD' or full ISO datetimes; entries may be strings or objects { student, lrn }
    const attendanceByStudentMonth: Record<string, Record<string, Set<number>>> = {};
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
        if (studentName) candidates.push(normalizeName(studentName));
        if (entryLrn) candidates.push(normalizeName(entryLrn));

        // Also accept raw object-to-string fallback
        if (!studentName && !entryLrn) {
          try {
            const asStr = JSON.stringify(entry);
            if (asStr) candidates.push(normalizeName(asStr));
          } catch {}
        }

        candidates.forEach((key) => {
          if (!key) return;
          if (!attendanceByStudentMonth[key]) attendanceByStudentMonth[key] = {};
          if (!attendanceByStudentMonth[key][month]) attendanceByStudentMonth[key][month] = new Set<number>();
          attendanceByStudentMonth[key][month].add(day);
        });
      });
    });

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
            if (s.student) candidates.push(normalizeName(s.student));
            if (s.lrn) candidates.push(normalizeName(s.lrn));
            if (candidates.length === 0) candidates.push(normalizeName(JSON.stringify(s)));
            candidates.forEach((key) => {
              if (!attendanceByStudentMonth[key]) attendanceByStudentMonth[key] = {};
              if (!attendanceByStudentMonth[key][month]) attendanceByStudentMonth[key][month] = new Set<number>();
              attendanceByStudentMonth[key][month].add(day);
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
      // Accept numbers, numeric strings ("1", "01"), and Date objects. Also handle ExcelJS formula/result shapes.
      const getCellDay = (cell: any): number | null => {
        if (!cell) return null;
        const v = cell.value;
        if (typeof v === 'number') return v;
        if (typeof v === 'string') {
          const s = v.trim();
          if (/^\d+$/.test(s)) return parseInt(s, 10);
        }
        if (v instanceof Date) return v.getDate();
        if (v && typeof v === 'object' && (v as any).result !== undefined) {
          const rv = (v as any).result;
          if (typeof rv === 'number') return rv;
          if (typeof rv === 'string' && /^\d+$/.test(rv.trim())) return parseInt(rv, 10);
          if (rv instanceof Date) return rv.getDate();
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

      // Prepare a 51x43 px green square image and add to workbook (for present markers)
      let greenImageId: number | null = null;
      try {
        // Create canvas and draw a hollow (stroked) green square to match requested appearance
        const canvas = document.createElement('canvas');
        canvas.width = 51; // px
        canvas.height = 43; // px
        const ctx = canvas.getContext('2d');
        if (ctx) {
          // Transparent background
          ctx.clearRect(0, 0, canvas.width, canvas.height);
          // Draw stroked rounded rectangle
          const pad = 6;
          const w = canvas.width - pad * 2;
          const h = canvas.height - pad * 2;
          const r = 4;
          ctx.strokeStyle = '#00B050';
          ctx.lineWidth = 3;
          // rounded rect path
          ctx.beginPath();
          ctx.moveTo(pad + r, pad);
          ctx.arcTo(pad + w, pad, pad + w, pad + r, r);
          ctx.arcTo(pad + w, pad + h, pad + w - r, pad + h, r);
          ctx.arcTo(pad, pad + h, pad, pad + h - r, r);
          ctx.arcTo(pad, pad, pad + r, pad, r);
          ctx.closePath();
          ctx.stroke();

          // Convert to ArrayBuffer via dataURL
          const dataUrl = canvas.toDataURL('image/png');
          const base64 = dataUrl.split(',')[1];
          const binary = atob(base64);
          const len = binary.length;
          const bytes = new Uint8Array(len);
          for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
          greenImageId = workbook.addImage({ buffer: bytes.buffer, extension: 'png' });
        }
      } catch (e) {
        console.warn('Failed to create green image for Excel markers, will fallback to glyph.', e);
        greenImageId = null;
      }

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
        // Ensure visible label even if student name is empty
        const visibleName = (reg.student && reg.student.toString().trim()) || (reg.lrn && reg.lrn.toString().trim()) || 'Unknown Student';
  const nameCell = worksheet.getCell(`B${rowIdx}`);
  if (includeNames) nameCell.value = visibleName;
  // Force a readable font color and alignment so the name is visible on any template
  (nameCell as any).font = { color: { argb: 'FF000000' }, size: 10 };
        nameCell.alignment = { vertical: 'middle', horizontal: 'left' } as any;
        worksheet.getRow(10).eachCell((cell, colNumber) => {
          const day = getCellDay(cell);
          if (day != null) {
            const markCell = worksheet.getRow(rowIdx).getCell(colNumber);

            const nameKey = normalizeName(visibleName);
            const lrnKey = normalizeName(reg.lrn);
            const studentMonthDates = new Set<number>();
            // Merge dates stored under the student's name and under their LRN (some histories store LRN)
            if (attendanceByStudentMonth[nameKey] && attendanceByStudentMonth[nameKey][month]) {
              attendanceByStudentMonth[nameKey][month].forEach((d: number) => studentMonthDates.add(d));
            }
            if (lrnKey && attendanceByStudentMonth[lrnKey] && attendanceByStudentMonth[lrnKey][month]) {
              attendanceByStudentMonth[lrnKey][month].forEach((d: number) => studentMonthDates.add(d));
            }

            // Apply fill and font first
            if (studentMonthDates.has(day)) {
              // ✅ Present: clear existing styles first
              try {
                (markCell as any).numFmt = undefined;
                (markCell as any).font = undefined;
                (markCell as any).alignment = { vertical: 'middle', horizontal: 'center' };
                (markCell as any).border = undefined;
                (markCell as any).fill = undefined;
              } catch {}

              if (greenImageId != null) {
                // Compute cell range for adding image. ExcelJS places images by tl/br in col,row coordinates
                // We'll place the image inside the single cell (colNumber,rowIdx). Use small offset to center.
                // ExcelJS expects zero-based column/row for positioning in 'from'/'to' with { col, row, offsetX, offsetY }
                try {
                  // Use A1 range to place image inside the single cell
                  const colLetter = worksheet.getColumn(colNumber).letter || (() => {
                    // fallback simple conversion
                    let n = colNumber; let s = '';
                    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
                    return s;
                  })();
                  // Present: leave the logical cell blank (visual marker is the image overlay)
                  markCell.value = "";
                  markCell.alignment = { vertical: 'middle', horizontal: 'center' } as any;
                  (markCell as any).font = undefined;
                  // Place the image using top-left / bottom-right coordinates (zero-based cols/rows)
                  worksheet.addImage(greenImageId, {
                    tl: { col: colNumber - 1, row: rowIdx - 1 },
                    ext: { width: 46, height: 38 }
                  });
                } catch (e) {
                  // Present fallback: leave logical cell blank. Image couldn't be added so there will be no visual marker.
                  markCell.value = "";
                  markCell.alignment = { vertical: "middle", horizontal: "center" } as any;
                  (markCell as any).font = undefined;
                }
              } else {
                // Fallback when image isn't available: leave the cell blank for present
                markCell.value = "";
                markCell.alignment = { vertical: "middle", horizontal: "center" } as any;
                (markCell as any).font = undefined;
              }
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
              color: { argb: studentMonthDates.has(day) ? "FFFFFFFF" : "FF000000" }
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
        // Ensure visible label even if student name is empty
        const visibleName = (reg.student && reg.student.toString().trim()) || (reg.lrn && reg.lrn.toString().trim()) || 'Unknown Student';
  const nameCell = worksheet.getCell(`B${rowIdx}`);
  if (includeNames) nameCell.value = visibleName;
        (nameCell as any).font = { color: { argb: 'FF000000' }, size: 10 };
        nameCell.alignment = { vertical: 'middle', horizontal: 'left' } as any;
        worksheet.getRow(10).eachCell((cell, colNumber) => {
          const day = getCellDay(cell);
          if (day != null) {
            const markCell = worksheet.getRow(rowIdx).getCell(colNumber);

            const nameKey = normalizeName(visibleName);
            const lrnKey = normalizeName(reg.lrn);
            const studentMonthDates = new Set<number>();
            if (attendanceByStudentMonth[nameKey] && attendanceByStudentMonth[nameKey][month]) {
              attendanceByStudentMonth[nameKey][month].forEach((d: number) => studentMonthDates.add(d));
            }
            if (lrnKey && attendanceByStudentMonth[lrnKey] && attendanceByStudentMonth[lrnKey][month]) {
              attendanceByStudentMonth[lrnKey][month].forEach((d: number) => studentMonthDates.add(d));
            }

            // Apply fill and font first
            if (studentMonthDates.has(day)) {
              // ✅ Present: clear existing styles and shade green (attendance cell only)
              try {
                (markCell as any).numFmt = undefined;
                (markCell as any).font = undefined;
                (markCell as any).alignment = { vertical: 'middle', horizontal: 'center' };
                (markCell as any).border = undefined;
                (markCell as any).fill = undefined;
              } catch {}
              if (greenImageId != null) {
                try {
                  // Use A1 range to place image inside the single cell
                  const colLetter = worksheet.getColumn(colNumber).letter || (() => {
                    let n = colNumber; let s = '';
                    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
                    return s;
                  })();
                  // Present: leave the logical cell blank (image overlay provides visual marker)
                  markCell.value = "";
                  markCell.alignment = { vertical: 'middle', horizontal: 'center' } as any;
                  (markCell as any).font = undefined;
                  worksheet.addImage(greenImageId, {
                    tl: { col: colNumber - 1, row: rowIdx - 1 },
                    ext: { width: 46, height: 38 }
                  });
                } catch (e) {
                // Fallback: leave blank for present
                markCell.value = "";
                markCell.alignment = { vertical: "middle", horizontal: "center" } as any;
                (markCell as any).font = undefined;
                }
              } else {
                // Fallback when image isn't available: leave the cell blank for present
                markCell.value = "";
                markCell.alignment = { vertical: "middle", horizontal: "center" } as any;
                (markCell as any).font = undefined;
              }
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
              color: { argb: studentMonthDates.has(day) ? "FFFFFFFF" : "FF000000" }
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
    } catch (e) {
      console.warn('Failed to add Scans worksheet', e);
    }

    const blob = await workbook.xlsx.writeBuffer();
    const url = URL.createObjectURL(new Blob([blob]));
    const link = document.createElement("a");
    link.href = url;
    link.download = `attendance_export.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

  } catch (err) {
    console.error("Excel export error:", err);
    alert("Failed to generate Excel file.");
  }
}



  async function downloadPDF(_attendance: { student: string; time: string }[]) {
    try {
      const { PDFDocument, rgb } = await import('pdf-lib');
      const templateBytes = await fetch('/attendance_template.pdf').then(res => res.arrayBuffer());
      const pdfDoc = await PDFDocument.load(templateBytes);
      const page = pdfDoc.getPages()[0];

      // Excel grid positions (same as Excel export)
      const excelNameStartX = 70;
      const excelNameStartY = 164;
      const excelNameHeight = 13;
      const excelCellStartX = 208;
      const excelCellStartY = 166;
      const excelCellWidth = 15;
      const excelCellHeight = 13;
      const pdfHeight = page.getHeight();

      // Month and date logic (same as Excel)
      const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
      const dates = [1, 2, 3, 4, 5, 8, 9, 10, 11, 12, 15, 16, 17, 18, 19, 22, 23, 24, 25, 26, 29, 30];

      // Normalize student name
      const normalizeName = (s: any) => (s || "").toString().trim().replace(/\s+/g, " ").toLowerCase();

      // Load attendance history from localStorage
      let history: Record<string, string[]> = {};
      try {
        history = JSON.parse(localStorage.getItem("attendance_history") || "{}");
      } catch {}

      // Build attendanceByStudentMonth from full history
      const attendanceByStudentMonth: Record<string, Record<string, Set<number>>> = {};
      Object.keys(history).forEach((dateStr) => {
        const dateObj = parseLocalDateFromYMD(dateStr);
        const month = monthNames[dateObj.getMonth()];
        const day = dateObj.getDate();
        history[dateStr].forEach((student) => {
          const key = normalizeName(student);
          if (!attendanceByStudentMonth[key]) attendanceByStudentMonth[key] = {};
          if (!attendanceByStudentMonth[key][month]) attendanceByStudentMonth[key][month] = new Set<number>();
          attendanceByStudentMonth[key][month].add(day);
        });
      });

      // Get registration data
      let registrations: { student: string; sex: string; lrn: string; parent: string; guardian: string }[] = [];
      try {
        const regRes = await fetch(`${API_BASE}/api/registrations/`);
        if (regRes.ok) {
          registrations = await regRes.json();
        } else {
          const regRaw = localStorage.getItem("registrations");
          if (regRaw) registrations = JSON.parse(regRaw);
        }
      } catch (e) {
        try {
          const regRaw = localStorage.getItem("registrations");
          if (regRaw) registrations = JSON.parse(regRaw);
        } catch {}
      }

  // List of all students (registered) with visible fallback
  const allStudents = registrations.map(r => (r.student && r.student.toString().trim()) || (r.lrn && r.lrn.toString().trim()) || 'Unknown Student');

      // For October (current month)
      const month = "OCT";

      // Overlay student names and attendance marks
      for (let i = 0; i < allStudents.length; i++) {
        const student = allStudents[i];
        const studentKey = normalizeName(student);
        const studentMonthDates = attendanceByStudentMonth[studentKey]?.[month] || new Set<number>();
        // Draw student name
        const y = pdfHeight - (excelNameStartY + i * excelNameHeight);
        page.drawText(student, {
          x: excelNameStartX,
          y: y,
          size: 8,
          color: rgb(0, 0, 0),
        });

        // Attendance marks
        let presentCount = 0;
        let absentCount = 0;
        for (let j = 0; j < dates.length; j++) {
          const day = dates[j];
          const x = excelCellStartX + j * excelCellWidth;
          const markY = pdfHeight - (excelCellStartY + i * excelCellHeight);
          if (studentMonthDates.has(day)) {
            // Present: use 'P' (supported by WinAnsi)
            page.drawText("P", {
              x: x,
              y: markY + (excelCellHeight / 2) - 4,
              size: 8,
              color: rgb(0, 176/255, 80/255),
            });
            presentCount++;
          } else {
            // Absent: X
            page.drawText("X", {
              x: x,
              y: markY + (excelCellHeight / 2) - 4,
              size: 8,
              color: rgb(0, 0, 0),
            });
            absentCount++;
          }
        }
        // Write totals to ABSENT and PRESENT columns
        const absentX = excelCellStartX + dates.length * excelCellWidth + 10;
        const presentX = absentX + 28;
        page.drawText(absentCount.toString(), {
          x: absentX,
          y: y + (excelCellHeight / 2) - 4,
          size: 8,
          color: rgb(0, 0, 0),
        });
        page.drawText(presentCount.toString(), {
          x: presentX,
          y: y + (excelCellHeight / 2) - 4,
          size: 8,
          color: rgb(0, 0, 0),
        });
      }

      // Save and download the PDF
      const pdfBytes = await pdfDoc.save();
      const arrayBuffer = pdfBytes instanceof Uint8Array ? pdfBytes.slice().buffer : pdfBytes;
      const blob = new Blob([arrayBuffer], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'SF2_Daily_Attendance_Template.pdf';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('[v0] SF2 PDF export error:', err);
      alert('Failed to generate SF2 PDF file.');
    }
  }

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

              {/* New PDF Export Button */}
              <Button variant="secondary" onClick={() => downloadPDF(attendance)}>
                <Download className="h-4 w-4 mr-2" />
                Download PDF
              </Button>

              <Button variant="destructive" onClick={clearAttendance}>
                <Trash2 className="h-4 w-4 mr-2" />
                Clear
              </Button>
            </div>
          </div>

          <div className="overflow-x-auto">
            <Table id="attendanceTable">
              <TableHeader>
                <TableRow>
                  <TableHead className="text-yellow-400">Student</TableHead>
                  <TableHead className="text-yellow-400">Time In</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {attendance.map((a, i) => (
                  <TableRow key={i}>
                    <TableCell>{a.student}</TableCell>
                    <TableCell>{a.time}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </div>
        </CardContent>
      </Card>
    </div>
  )
}
