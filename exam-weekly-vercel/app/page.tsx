'use client';
import { useState } from "react";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!file) return;

    const formData = new FormData();
    formData.append("file", file);

    const res = await fetch("/api/process", {
      method: "POST",
      body: formData,
    });

    if (!res.ok) {
      alert("Processing failed");
      return;
    }

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "exam_weekly.xlsx";
    link.click();
  };

  return (
    <div style={{ padding: "2rem" }}>
      <h1>Excel Weekly Report Generator</h1>
      <form onSubmit={handleSubmit}>
        <input
          type="file"
          accept=".xlsx"
          onChange={(e) => setFile(e.target.files?.[0] || null)}
          required
        />
        <button type="submit">Upload and Process</button>
      </form>
    </div>
  );
}
