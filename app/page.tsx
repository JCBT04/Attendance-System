"use client";

import { useState, useEffect } from "react";
import Link from "next/link"; // ✅ Import Next.js Link
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Download, Trash2, QrCode, Save, LayoutDashboard } from "lucide-react";
import { QRCodeCanvas } from "qrcode.react";

interface Registration {
  lrn: string;
  student: string;
  parent: string;
  guardian: string;
}

export default function RegisterPage() {
  const [student, setStudent] = useState("");
  const [parent, setParent] = useState("");
  const [guardian, setGuardian] = useState("");
  const [lrn, setLrn] = useState("");
  const [registrations, setRegistrations] = useState<Registration[]>([]);
  const [qrValue, setQrValue] = useState<string>("");

  // load from localStorage
  useEffect(() => {
    try {
      const saved = localStorage.getItem("registrations");
      if (saved) setRegistrations(JSON.parse(saved));
    } catch (e) {
      console.error("Error loading registrations", e);
    }
  }, []);

  // persist to localStorage
  useEffect(() => {
    localStorage.setItem("registrations", JSON.stringify(registrations));
  }, [registrations]);

  const generateQR = () => {
    if (!student || !parent || !guardian || !lrn) {
      alert("Please fill all fields.");
      return;
    }

    const data: Registration = { lrn, student, parent, guardian };
    const payload = JSON.stringify(data);

    setQrValue(payload);
    setRegistrations((prev) => [...prev, data]);
    alert("QR generated and registration saved.");
  };

  const saveManual = () => {
    if (!student || !parent || !guardian || !lrn) {
      alert("Please fill all fields.");
      return;
    }
    setRegistrations((prev) => [...prev, { lrn, student, parent, guardian }]);
    alert("Registration saved.");
  };

  const downloadQR = () => {
    const canvas = document.querySelector(
      "canvas"
    ) as HTMLCanvasElement | null;
    if (!canvas) {
      alert("Generate a QR first.");
      return;
    }
    const imageURL = canvas.toDataURL("image/png");
    const link = document.createElement("a");
    link.href = imageURL;
    link.download = "student_qrcode.png";
    link.click();
  };

  const downloadAll = () => {
    if (!registrations.length) {
      alert("No registrations yet.");
      return;
    }
    let csv = "LRN,Student Name,Parent Name,Guardian Name\n";
    registrations.forEach((r) => {
      csv += `${r.lrn},"${r.student}","${r.parent}","${r.guardian}"\n`;
    });
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `registrations_${new Date().toISOString().split("T")[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const clearAll = () => {
    if (confirm("Are you sure you want to clear all registrations?")) {
      setRegistrations([]);
      localStorage.removeItem("registrations");
      setQrValue("");
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-slate-900 to-slate-800 p-6 text-white">
      <Card className="max-w-5xl w-full border-yellow-500 border-2 bg-white/10 backdrop-blur-lg text-white">
        <CardHeader className="flex justify-between items-center px-7">
          <CardTitle className="text-center text-yellow-400 text-2xl">
            Student QR Code Registration
          </CardTitle>

          <Link href="/dashboard">
            <Button variant="secondary">
              <LayoutDashboard className="h-4 w-4 mr-2" />
              Go to Dashboard
            </Button>
          </Link>
        </CardHeader>

        <CardContent>
          <div className="grid md:grid-cols-2 gap-6">
            {/* QR Panel */}
            <div className="flex flex-col items-center gap-4 p-4 border rounded-xl border-yellow-400/50">
              <div className="bg-white p-3 rounded-lg">
                {qrValue ? (
                  <QRCodeCanvas value={qrValue} size={220} includeMargin />
                ) : (
                  <p className="text-sm text-gray-400">No QR generated yet</p>
                )}
              </div>
              <div className="flex flex-col gap-2 w-full">
                <Button onClick={downloadQR} className="w-full">
                  <Download className="h-4 w-4 mr-2" /> Download QR
                </Button>
              </div>
            </div>

            {/* Form Panel */}
            <div className="space-y-3">
              <label className="block text-sm">Student Name</label>
              <Input value={student} onChange={(e) => setStudent(e.target.value)} />

              <label className="block text-sm">Parent Name</label>
              <Input value={parent} onChange={(e) => setParent(e.target.value)} />

              <label className="block text-sm">Guardian Name</label>
              <Input value={guardian} onChange={(e) => setGuardian(e.target.value)} />

              <label className="block text-sm">LRN Number</label>
              <Input value={lrn} onChange={(e) => setLrn(e.target.value)} />

              <div className="flex flex-wrap gap-2 mt-3">
                <Button onClick={saveManual} variant="secondary">
                  <Save className="h-4 w-4 mr-2" /> Save
                </Button>
                <Button onClick={generateQR}>
                  <QrCode className="h-4 w-4 mr-2" /> Generate & Save
                </Button>
                <Button onClick={downloadAll} variant="secondary">
                  <Download className="h-4 w-4 mr-2" /> Download CSV
                </Button>
                <Button onClick={clearAll} variant="destructive">
                  <Trash2 className="h-4 w-4 mr-2" /> Clear All
                </Button>
              </div>

              <p className="text-muted-foreground text-sm mt-3">
                QR payload format: {"{"}"lrn":"...","student":"...","parent":"...","guardian":"..."{"}"}
              </p>
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
