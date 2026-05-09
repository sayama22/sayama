"use client";

import { useCallback, useRef, useState } from "react";
import { Button } from "@/components/ui/button";
import { CloudUpload, FileSpreadsheet, Loader2, X } from "lucide-react";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback((f: File) => {
    if (f.type !== "application/pdf") {
      setError("PDFファイルのみ対応しています");
      return;
    }
    setError(null);
    setFile(f);
  }, []);

  const onDrop = useCallback(
    (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      setIsDragging(false);
      const dropped = e.dataTransfer.files[0];
      if (dropped) handleFile(dropped);
    },
    [handleFile]
  );

  const onInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selected = e.target.files?.[0];
    if (selected) handleFile(selected);
  };

  const handleConvert = async () => {
    if (!file) return;
    setIsConverting(true);
    setError(null);
    try {
      const formData = new FormData();
      formData.append("file", file);

      const res = await fetch("/api/convert", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const json = await res.json().catch(() => ({}));
        throw new Error(json.error ?? "変換に失敗しました");
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = file.name.replace(/\.pdf$/i, ".xlsx");
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      setError(err instanceof Error ? err.message : "変換に失敗しました");
    } finally {
      setIsConverting(false);
    }
  };

  const removeFile = () => {
    setFile(null);
    setError(null);
    if (inputRef.current) inputRef.current.value = "";
  };

  return (
    <div className="min-h-screen bg-white flex flex-col">
      {/* Header */}
      <header className="border-b border-gray-100 px-6 py-4">
        <div className="flex items-center gap-3">
          <span className="text-xl font-bold tracking-tight text-gray-900">
            SAYAMA
          </span>
          <span className="text-sm text-gray-400">PDF → Excel</span>
        </div>
      </header>

      {/* Main */}
      <main className="flex-1 flex flex-col items-center justify-center px-4 py-16">
        <div className="w-full max-w-lg space-y-3">
          <h1 className="text-2xl font-bold text-gray-900">PDF取り込み</h1>
          <p className="text-sm text-gray-500">
            PDFファイルをExcel形式に変換します
          </p>

          {/* Drop zone */}
          <div
            className={`mt-6 relative border-2 border-dashed rounded-2xl p-12 flex flex-col items-center justify-center gap-3 transition-colors cursor-pointer select-none
              ${
                isDragging
                  ? "border-gray-400 bg-gray-50"
                  : "border-gray-200 hover:border-gray-300 hover:bg-gray-50/50"
              }`}
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={onDrop}
            onClick={() => inputRef.current?.click()}
          >
            <input
              ref={inputRef}
              type="file"
              accept="application/pdf"
              className="hidden"
              onChange={onInputChange}
            />

            {file ? (
              <>
                <FileSpreadsheet className="w-10 h-10 text-green-500" />
                <div className="flex items-center gap-2">
                  <span className="text-sm font-medium text-gray-700 max-w-xs truncate">
                    {file.name}
                  </span>
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      removeFile();
                    }}
                    className="text-gray-400 hover:text-gray-600 transition-colors"
                    aria-label="ファイルを削除"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>
                <span className="text-xs text-gray-400">
                  {(file.size / 1024).toFixed(1)} KB
                </span>
              </>
            ) : (
              <>
                <CloudUpload className="w-10 h-10 text-gray-300" />
                <p className="text-sm font-medium text-gray-600">
                  PDFをドラッグ＆ドロップ
                </p>
                <p className="text-xs text-gray-400">
                  または クリックしてファイルを選択
                </p>
              </>
            )}
          </div>

          {/* Error */}
          {error && (
            <p className="text-sm text-red-500 text-center">{error}</p>
          )}

          {/* Convert button */}
          {file && (
            <Button
              className="w-full mt-2"
              size="lg"
              onClick={handleConvert}
              disabled={isConverting}
            >
              {isConverting ? (
                <>
                  <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                  変換中...
                </>
              ) : (
                <>
                  <FileSpreadsheet className="w-4 h-4 mr-2" />
                  Excelをダウンロード
                </>
              )}
            </Button>
          )}
        </div>
      </main>
    </div>
  );
}
