"use client";

import { useState, useEffect } from "react";
import Link from "next/link";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Download, Trash2, QrCode, Save, LayoutDashboard } from "lucide-react";
import { QRCodeCanvas } from "qrcode.react";

interface Registration {
  lrn: string;
  student: string;
  sex: string;
  parent: string;
  guardian: string;
}

export default function RegisterPage() {
  const [student, setStudent] = useState("");
  const [sex, setSex] = useState("");
  const [parent, setParent] = useState("");
  const [guardian, setGuardian] = useState("");
  const [lrn, setLrn] = useState("");
  const [qrValue, setQrValue] = useState<string>("");

  // ✅ Initialize from localStorage immediately
  const [registrations, setRegistrations] = useState<Registration[]>(() => {
    try {
      const saved = localStorage.getItem("registrations");
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  function refreshRegistrations() {
    try {
      const saved = localStorage.getItem("registrations");
      if (!saved) {
        setRegistrations([]);
        return;
      }
      const parsed: Registration[] = JSON.parse(saved);
      const normalized = parsed.map((r: Registration) => ({
        ...r,
        student: (r.student || "").trim(),
        parent: (r.parent || "").trim(),
        guardian: (r.guardian || "").trim(),
        sex: ((): string => {
          const s = (r.sex || "").toString().trim().toLowerCase();
          if (s.startsWith("m")) return "Male";
          if (s.startsWith("f")) return "Female";
          return r.sex || "";
        })(),
      }));
      normalized.sort((a, b) => a.student.localeCompare(b.student));
      setRegistrations(normalized);
    } catch (e) {
      console.error("Error loading registrations", e);
      setRegistrations([]);
    }
  }

  useEffect(() => {
    refreshRegistrations(); // refresh on mount
    const handleFocus = () => refreshRegistrations();
    window.addEventListener("focus", handleFocus);
    return () => {
      window.removeEventListener("focus", handleFocus);
    };
  }, []);

  // Persist to localStorage whenever registrations change
  useEffect(() => {
    try {
      localStorage.setItem("registrations", JSON.stringify(registrations));
    } catch (e) {
      console.error("Failed to persist registrations", e);
    }
  }, [registrations]);

//   const generateQR = () => {
//     if (!student || !sex || !parent || !guardian || !lrn) {
//       alert("Please fill all fields.");
//       return;
//     }
//     const data: Registration = {
//       lrn: lrn.trim(),
//       student: student.trim(),
//       sex: sex,
//       parent: parent.trim(),
//       guardian: guardian.trim(),
//     };
//     const payload = JSON.stringify(data);
//     setQrValue(payload);
//     setRegistrations((prev) => {
//       const next = [...prev, data];
//       next.sort((a, b) => a.student.localeCompare(b.student));
//       try {
//         localStorage.setItem("registrations", JSON.stringify(next));
//       } catch (e) {
//         console.error("Failed to persist registrations immediately", e);
//       }
//       return next;
//     });
//     setStudent("");
//     setSex("");
//     setParent("");
//     setGuardian("");
//     setLrn("");
//     alert("QR generated and registration saved.");
//   };

const generateQR = async () => {
  if (!student || !sex || !parent || !guardian || !lrn) {
    alert("Please fill all fields.");
    return;
  }

  const data: Registration = {
    lrn: lrn.trim(),
    student: student.trim(),
    sex,
    parent: parent.trim(),
    guardian: guardian.trim(),
  };

  try {
    const response = await fetch("http://localhost:8000/api/registrations/", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    });

    if (response.ok) {
      const payload = JSON.stringify(data);
      setQrValue(payload);
      setRegistrations((prev) => {
        const next = [...prev, data].sort((a, b) => a.student.localeCompare(b.student));
        localStorage.setItem("registrations", JSON.stringify(next));
        return next;
      });
      setStudent("");
      setSex("");
      setParent("");
      setGuardian("");
      setLrn("");
      alert("QR generated and registration saved to backend!");
    } else {
      const result = await response.json();
      alert("Error: " + (result.error || "Failed to register."));
    }
  } catch (error) {
    console.error("Backend connection error:", error);
    alert("Unable to connect to backend.");
  }
};


//   const saveManual = () => {
//     if (!student || !sex || !parent || !guardian || !lrn) {
//       alert("Please fill all fields.");
//       return;
//     }
//     setRegistrations((prev) => {
//       const entry: Registration = {
//         lrn: lrn.trim(),
//         student: student.trim(),
//         sex: sex,
//         parent: parent.trim(),
//         guardian: guardian.trim(),
//       };
//       const next = [...prev, entry];
//       next.sort((a, b) => a.student.localeCompare(b.student));
//       try {
//         localStorage.setItem("registrations", JSON.stringify(next));
//       } catch (e) {
//         console.error("Failed to persist registrations immediately", e);
//       }
//       return next;
//     });
//     setStudent("");
//     setSex("");
//     setParent("");
//     setGuardian("");
//     setLrn("");
//     alert("Registration saved.");
//   };

const saveManual = async () => {
  if (!student || !sex || !parent || !guardian || !lrn) {
    alert("Please fill all fields.");
    return;
  }

  const entry: Registration = {
    lrn: lrn.trim(),
    student: student.trim(),
    sex,
    parent: parent.trim(),
    guardian: guardian.trim(),
  };

  try {
    const response = await fetch("http://localhost:8000/api/registrations/", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(entry),
    });

    if (response.ok) {
      alert("Student registered successfully!");
      setRegistrations((prev) => {
        const next = [...prev, entry].sort((a, b) => a.student.localeCompare(b.student));
        localStorage.setItem("registrations", JSON.stringify(next));
        return next;
      });
      setStudent("");
      setSex("");
      setParent("");
      setGuardian("");
      setLrn("");
    } else {
      const data = await response.json();
      alert("Error: " + (data.error || "Failed to register."));
    }
  } catch (error) {
    console.error("Backend connection error:", error);
    alert("Unable to connect to backend.");
  }
};

  const downloadQR = () => {
    const canvas = document.querySelector("canvas") as HTMLCanvasElement | null;
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

//   const downloadAll = () => {
//     // Always pull fresh data from localStorage
//     const raw = localStorage.getItem("registrations");
//     if (!raw) {
//       alert("No registrations yet.");
//       return;
//     }
//     let parsed: Registration[];
//     try {
//       parsed = JSON.parse(raw);
//     } catch {
//       alert("No registrations yet.");
//       return;
//     }
//     if (!parsed.length) {
//       alert("No registrations yet.");
//       return;
//     }

//     const sorted = parsed.slice().sort((a, b) => a.student.localeCompare(b.student));
//     let csv = "LRN,Student Name,Sex,Parent Name,Guardian Name\n";
//     sorted.forEach((r) => {
//       const lrnCell = r.lrn ?? "";
//       const studentCell = (r.student ?? "").replace(/"/g, '""');
//       const sexCell = r.sex ?? "";
//       const parentCell = (r.parent ?? "").replace(/"/g, '""');
//       const guardianCell = (r.guardian ?? "").replace(/"/g, '""');
//       csv += `${lrnCell},"${studentCell}","${sexCell}","${parentCell}","${guardianCell}"\n`;
//     });

//     const males = sorted.filter((r) => (r.sex || "").toLowerCase().startsWith("m"));
//     const females = sorted.filter((r) => (r.sex || "").toLowerCase().startsWith("f"));

//     csv += "\n--- Male Students ---\n";
//     males.forEach((r, i) => {
//       csv += `${i + 1}. ${r.student}\n`;
//     });

//     csv += "\n--- Female Students ---\n";
//     females.forEach((r, i) => {
//       csv += `${i + 1}. ${r.student}\n`;
//     });

//     const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
//     const url = URL.createObjectURL(blob);
//     const a = document.createElement("a");
//     a.href = url;
//     a.download = `registrations_${new Date().toISOString().split("T")[0]}.csv`;
//     a.click();
//     URL.revokeObjectURL(url);
//   };

const downloadAll = async () => {
  try {
    const response = await fetch("http://localhost:8000/api/registrations/");
    if (!response.ok) {
      alert("Failed to fetch registrations from backend.");
      return;
    }

    const data: Registration[] = await response.json();
    if (!data.length) {
      alert("No registrations found.");
      return;
    }

    const sorted = data.slice().sort((a, b) => a.student.localeCompare(b.student));

    let csv = "LRN,Student Name,Sex,Parent Name,Guardian Name\n";
    sorted.forEach((r) => {
      const lrnCell = r.lrn ?? "";
      const studentCell = (r.student ?? "").replace(/"/g, '""');
      const sexCell = r.sex ?? "";
      const parentCell = (r.parent ?? "").replace(/"/g, '""');
      const guardianCell = (r.guardian ?? "").replace(/"/g, '""');
      csv += `${lrnCell},"${studentCell}","${sexCell}","${parentCell}","${guardianCell}"\n`;
    });

    const males = sorted.filter((r) => (r.sex || "").toLowerCase().startsWith("m"));
    const females = sorted.filter((r) => (r.sex || "").toLowerCase().startsWith("f"));

    csv += "\n--- Male Students ---\n";
    males.forEach((r, i) => {
      csv += `${i + 1}. ${r.student}\n`;
    });

    csv += "\n--- Female Students ---\n";
    females.forEach((r, i) => {
      csv += `${i + 1}. ${r.student}\n`;
    });

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `registrations_${new Date().toISOString().split("T")[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  } catch (error) {
    console.error("Error downloading CSV:", error);
    alert("Unable to connect to backend.");
  }
};

//   const clearAll = () => {
//     if (confirm("Are you sure you want to clear all registrations?")) {
//       setRegistrations([]);
//       localStorage.removeItem("registrations");
//       setQrValue("");
//     }
//   };

const clearAll = async () => {
  if (!window.confirm("Are you sure you want to delete all registrations from the backend?")) {
    return;
  }

  try {
    const response = await fetch("http://localhost:8000/api/registrations/clear_all/", {
      method: "DELETE",
    });

    if (response.ok) {
      alert("All registrations deleted successfully!");
      setRegistrations([]);
      localStorage.removeItem("registrations");
    } else {
      alert("Failed to clear registrations on backend.");
    }
  } catch (error) {
    console.error("Error clearing registrations:", error);
    alert("Unable to connect to backend.");
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
              <label className="block text-sm">Full Name</label>
              <Input value={student} onChange={(e) => setStudent(e.target.value)} />
              <p className="text-xs text-yellow-300 mt-1">
                Format: Last Name, First Name Middle Initial
              </p>

              <label className="block text-sm mt-2">Sex</label>
              <select
                value={sex}
                onChange={(e) => setSex(e.target.value)}
                className="w-full p-2 rounded bg-slate-900 text-white border border-yellow-400"
              >
                <option value="">Select...</option>
                <option value="Male">Male</option>
                <option value="Female">Female</option>
              </select>

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
                QR payload format: {"{"}"lrn":"...","student":"Last Name, First Name Middle Initial","sex":"Male/Female","parent":"...","guardian":"..."{"}"}
              </p>
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
