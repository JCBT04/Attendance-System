"use client";

import { useEffect, useRef, useState } from "react";
import Link from "next/link"; // 👈 import Link
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
  DialogFooter,
} from "@/components/ui/dialog";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Download, Camera, Trash2, Upload, UserPlus } from "lucide-react";
import { Card, CardHeader, CardTitle, CardContent } from "@/components/ui/card";
import { Html5Qrcode } from "html5-qrcode";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

export default function Home() {
  const [open, setOpen] = useState(false);
  const [attendance, setAttendance] = useState<
    { student: string; time: string }[]
  >([]);
  const scannerRef = useRef<Html5Qrcode | null>(null);

  // Load attendance
  useEffect(() => {
    const raw = localStorage.getItem("attendance_simple");
    if (raw) {
      try {
        setAttendance(JSON.parse(raw));
      } catch {
        console.error("Invalid saved attendance");
      }
    }
  }, []);

  // Save attendance
  useEffect(() => {
    localStorage.setItem("attendance_simple", JSON.stringify(attendance));
  }, [attendance]);

  function nowTime(): string {
    const n = new Date();
    let h = n.getHours();
    const m = String(n.getMinutes()).padStart(2, "0");
    const s = String(n.getSeconds()).padStart(2, "0");
    const ampm = h >= 12 ? "PM" : "AM";
    h = h % 12 || 12;
    const date = n.toISOString().split("T")[0];
    return `${date} ${String(h).padStart(2, "0")}:${m}:${s} ${ampm}`;
  }

  function addAttendance(student: string) {
    setAttendance((prev) => [...prev, { student, time: nowTime() }]);
  }

  function parseStudentFromQr(qrMessage: string): string | null {
    try {
      const obj = JSON.parse(qrMessage);
      return obj.student || null;
    } catch {
      return qrMessage;
    }
  }

  async function startScanner() {
    try {
      const readerElem = document.getElementById("reader");
      if (!readerElem) {
        console.warn("Reader element not found yet.");
        return;
      }

      if (!scannerRef.current) {
        scannerRef.current = new Html5Qrcode("reader");
      }

      await scannerRef.current.start(
        { facingMode: "environment" },
        { fps: 10, qrbox: 250 },
        (decoded: string) => {
          const student = parseStudentFromQr(decoded) || "(Unknown)";
          addAttendance(student);
          stopScanner();
          setOpen(false);
        },
        (err: string) => {
          console.warn("QR scan error:", err);
        }
      );
    } catch (e) {
      console.error("Scanner error:", e);
      alert("Could not access camera. You can upload a QR code image instead.");
      stopScanner();
    }
  }

  async function stopScanner() {
    if (scannerRef.current) {
      try {
        await scannerRef.current.stop();
        await scannerRef.current.clear();
      } catch {
        // ignore
      }
    }
  }

  // async function handleImageUpload(e: React.ChangeEvent<HTMLInputElement>) {
  //   const file = e.target.files?.[0];
  //   if (!file) return;

  //   try {
  //     if (!scannerRef.current) {
  //       scannerRef.current = new Html5Qrcode("reader");
  //     }
  //     const result = await scannerRef.current.scanFile(file, true);
  //     const student = parseStudentFromQr(result) || "(Unknown)";
  //     addAttendance(student);
  //     setOpen(false);
  //   } catch (err) {
  //     console.error("Image scan failed:", err);
  //     alert("Failed to read QR code from image.");
  //   }
  // }
  async function handleImageUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      // If camera scan is ongoing, stop it before scanning file
      if (scannerRef.current) {
        try {
          await scannerRef.current.stop();
        } catch {
          // ignore if already stopped
        }
      }

      if (!scannerRef.current) {
        scannerRef.current = new Html5Qrcode("reader");
      }

      const result = await scannerRef.current.scanFile(file, true);
      const student = parseStudentFromQr(result) || "(Unknown)";
      addAttendance(student);

      setOpen(false);
    } catch (err) {
      console.error("Image scan failed:", err);
      alert("Failed to read QR code from image.");
    }
  }


  function clearAttendance() {
    setAttendance([]);
    localStorage.removeItem("attendance_simple");
  }

  async function downloadExcel(attendance: { student: string; time: string }[]) {
  try {
    const response = await fetch("/attendance_template.xlsx");
    const buffer = await response.arrayBuffer();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const worksheet = workbook.getWorksheet("AUG");
    if (!worksheet) {
      throw new Error("Worksheet 'AUG' not found in template");
    }

    const studentRowMap: Record<string, number> = {};

    attendance.forEach((entry, index) => {
      const row = index + 13;
      worksheet.getCell(`B${row}`).value = entry.student;
      studentRowMap[entry.student] = row;

      const datePart = entry.time.split(" ")[0];
      const day = new Date(datePart).getDate();

      // 🔹 Instead of 4+day, find column in row 10 that matches the day
      let colIndex = -1;
      worksheet.getRow(10).eachCell((cell, colNumber) => {
        if (cell.value === day) {
          colIndex = colNumber;
        }
      });

      if (colIndex !== -1) {
        worksheet.getRow(row).getCell(colIndex).value = "✔";
      }
    });

    const today = new Date();
    const currentDay = today.getDate();

    Object.values(studentRowMap).forEach((row) => {
      worksheet.getRow(10).eachCell((cell, colNumber) => {
        if (typeof cell.value === "number" && cell.value < currentDay) {
          const markCell = worksheet.getRow(row).getCell(colNumber);
          if (!markCell.value) {
            markCell.value = "X";
          }
        }
      });
    });

    const blob = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([blob]), "attendance_filled.xlsx");
  } catch (err) {
    console.error("Excel export error:", err);
    alert("Failed to generate Excel file.");
  }
}

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-slate-900 to-slate-800 p-6 text-white">
      <Card className="max-w-5xl w-full border-yellow-500 border-2 bg-white/10 backdrop-blur-lg text-white">
        <CardHeader className="flex flex-col sm:flex-row sm:items-center sm:justify-between">
          <CardTitle className="text-center text-yellow-400 text-2xl">
            Student Attendance Dashboard
          </CardTitle>

          {/* 👇 Link to registration page */}
          <Link href="/" passHref>
          {/* <Button variant="secondary">
              <LayoutDashboard className="h-4 w-4 mr-2" />
              Go to Dashboard
            </Button> */}
            <Button variant="secondary" className="mt-2 sm:mt-0">
              <UserPlus className="h-4 w-4 mr-2" />
              Register Student
            </Button>
          </Link>
        </CardHeader>

        <CardContent className="space-y-6">
          <div className="flex items-center justify-between">
            <p className="text-muted-foreground">
              Scan a QR code to log attendance instantly.
            </p>
            <div className="flex gap-2">
              <Dialog
                open={open}
                onOpenChange={(val) => {
                  setOpen(val);
                  if (val) {
                    setTimeout(() => startScanner(), 300);
                  } else {
                    stopScanner();
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
                    <DialogTitle className="text-yellow-400">
                      Scan QR Code
                    </DialogTitle>
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
                      <input
                        type="file"
                        accept="image/*"
                        onChange={handleImageUpload}
                        className="hidden"
                      />
                    </label>
                  </div>

                  <DialogFooter>
                    <Button variant="destructive" onClick={stopScanner}>
                      Stop
                    </Button>
                  </DialogFooter>
                </DialogContent>
              </Dialog>

              {/* <Button variant="secondary" onClick={downloadCsv}>
                <Download className="h-4 w-4 mr-2" />
                Download CSV
              </Button> */}

              <Button variant="secondary" onClick={() => downloadExcel(attendance)}>
                <Download className="h-4 w-4 mr-2" />
                Download Excel
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
  );
}

