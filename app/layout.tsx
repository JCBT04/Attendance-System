import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "Student Attendance System",
  description: "Attendance management system for students",
};

export default function RootLayout({
  children,
}: Readonly<{ children: React.ReactNode }>) {
  // Ensure the variables are always defined
  const sansClass = geistSans.variable ?? "";
  const monoClass = geistMono.variable ?? "";

  return (
    <html lang="en">
      <body className={`${sansClass} ${monoClass} antialiased`}>
        {children}
      </body>
    </html>
  );
}
