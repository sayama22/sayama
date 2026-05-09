"use client";

import { useCallback, useRef, useState } from "react";
import { Button } from "@/components/ui/button";
import {
  CloudUpload,
  FileSpreadsheet,
  Loader2,
  X,
  CheckCircle2,
  AlertCircle,
} from "lucide-react";

type Status = "idle" | "selected" | "converting" | "done" | "error";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [status, setStatus] = useState<Status>("idle");
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback((f: File) => {
    if (f.type !== "application/pdf") {
      setErrorMsg("PDFファイルのみ対応しています");
      setStatus("error");
      return;
    }
    setErrorMsg(null);
    setFile(f);
    setStatus("selected");
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
    setStatus("converting");
    setErrorMsg(null);
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
      setStatus("done");
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : "変換に失敗しました");
      setStatus("error");
    }
  };

  const reset = () => {
    setFile(null);
    setErrorMsg(null);
    setStatus("idle");
    if (inputRef.current) inputRef.current.value = "";
  };

  const isConverting = status === "converting";
  const showDropzone = status !== "converting";

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
        <div className="w-full max-w-lg space-y-4">
          <div>
            <h1 className="text-2xl font-bold text-gray-900">PDF取り込み</h1>
            <p className="text-sm text-gray-500 mt-1">
              PDFファイルをExcel形式に変換します
            </p>
          </div>

          {/* ---- Converting overlay ---- */}
          {isConverting && (
            <div className="border-2 border-dashed border-blue-200 rounded-2xl p-14 flex flex-col items-center justify-center gap-4 bg-blue-50/30">
              <Loader2 className="w-10 h-10 text-blue-400 animate-spin" />
              <div className="text-center">
                <p className="text-sm font-medium text-gray-700">変換中...</p>
                <p className="text-xs text-gray-400 mt-1">
                  テーブル構造を解析してExcelに書き出しています
                </p>
              </div>
              <div className="w-48 h-1.5 bg-blue-100 rounded-full overflow-hidden">
                <div className="h-full bg-blue-400 rounded-full animate-[pulse_1.2s_ease-in-out_infinite]" />
              </div>
            </div>
          )}

          {/* ---- Drop zone (hidden while converting) ---- */}
          {showDropzone && (
            <div
              className={`relative border-2 border-dashed rounded-2xl p-12 flex flex-col items-center justify-center gap-3 transition-colors cursor-pointer select-none
                ${
                  isDragging
                    ? "border-blue-300 bg-blue-50/40"
                    : status === "done"
                    ? "border-green-200 bg-green-50/20"
                    : status === "error"
                    ? "border-red-200 bg-red-50/10"
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

              {status === "done" ? (
                <>
                  <CheckCircle2 className="w-10 h-10 text-green-500" />
                  <p className="text-sm font-medium text-green-700">
                    変換完了・ダウンロード済み
                  </p>
                  <p className="text-xs text-gray-400">
                    クリックして別のPDFを選択
                  </p>
                </>
              ) : file ? (
                <>
                  <FileSpreadsheet className="w-10 h-10 text-blue-400" />
                  <div className="flex items-center gap-2">
                    <span className="text-sm font-medium text-gray-700 max-w-xs truncate">
                      {file.name}
                    </span>
                    <button
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        reset();
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
          )}

          {/* ---- Error message ---- */}
          {status === "error" && errorMsg && (
            <div className="flex items-start gap-2 rounded-xl bg-red-50 border border-red-100 px-4 py-3">
              <AlertCircle className="w-4 h-4 text-red-500 mt-0.5 shrink-0" />
              <p className="text-sm text-red-600">{errorMsg}</p>
            </div>
          )}

          {/* ---- Action buttons ---- */}
          {status === "selected" && (
            <Button className="w-full" size="lg" onClick={handleConvert}>
              <FileSpreadsheet className="w-4 h-4 mr-2" />
              Excelをダウンロード
            </Button>
          )}

          {status === "done" && (
            <Button
              className="w-full"
              size="lg"
              variant="outline"
              onClick={reset}
            >
              別のPDFを変換する
            </Button>
          )}

          {status === "error" && file && (
            <div className="flex gap-2">
              <Button
                className="flex-1"
                size="lg"
                onClick={handleConvert}
              >
                <FileSpreadsheet className="w-4 h-4 mr-2" />
                再試行
              </Button>
              <Button
                className="flex-1"
                size="lg"
                variant="outline"
                onClick={reset}
              >
                別のファイルを選択
              </Button>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}
