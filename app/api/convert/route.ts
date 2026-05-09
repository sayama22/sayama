import { NextRequest, NextResponse } from "next/server";
import { PDFParse } from "pdf-parse";
import ExcelJS from "exceljs";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------
interface PositionedText {
  text: string;
  x: number;
  y: number;
  width: number;
}

interface TextRow {
  items: PositionedText[];
  y: number;
}

// ---------------------------------------------------------------------------
// Helper: extract text items with coordinates using pdfjs-dist
// ---------------------------------------------------------------------------
async function extractPositionedText(
  buffer: Buffer
): Promise<Map<number, TextRow[]>> {
  const pdfjsLib = await import("pdfjs-dist");
  pdfjsLib.GlobalWorkerOptions.workerSrc = "";

  const pdf = await pdfjsLib
    .getDocument({ data: new Uint8Array(buffer), verbosity: 0 })
    .promise;

  const pageMap = new Map<number, TextRow[]>();

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();

    const items: PositionedText[] = content.items
      .filter(
        (item): item is typeof item & { str: string } =>
          "str" in item && typeof item.str === "string" && item.str.trim() !== ""
      )
      .map((item) => {
        const t = (item as { transform: number[] }).transform;
        return {
          text: (item as { str: string }).str,
          x: Math.round(t[4]),
          y: Math.round(viewport.height - t[5]), // flip: PDF y goes up
          width: (item as { width: number }).width,
        };
      });

    pageMap.set(p, groupIntoRows(items, 3));
  }

  return pageMap;
}

// ---------------------------------------------------------------------------
// Helper: group text items into rows by y-coordinate
// ---------------------------------------------------------------------------
function groupIntoRows(items: PositionedText[], tolerance = 3): TextRow[] {
  if (items.length === 0) return [];
  const sorted = [...items].sort((a, b) => a.y - b.y || a.x - b.x);
  const rows: TextRow[] = [];
  let current: PositionedText[] = [sorted[0]];
  let currentY = sorted[0].y;

  for (let i = 1; i < sorted.length; i++) {
    const item = sorted[i];
    if (Math.abs(item.y - currentY) <= tolerance) {
      current.push(item);
    } else {
      rows.push({ items: current.sort((a, b) => a.x - b.x), y: currentY });
      current = [item];
      currentY = item.y;
    }
  }
  if (current.length > 0) {
    rows.push({ items: current.sort((a, b) => a.x - b.x), y: currentY });
  }
  return rows;
}

// ---------------------------------------------------------------------------
// Helper: cluster x-values into column buckets
// ---------------------------------------------------------------------------
function buildColBuckets(rows: TextRow[], tolerance = 15): number[] {
  const xs = rows.flatMap((r) => r.items.map((i) => i.x));
  if (xs.length === 0) return [];
  const sorted = [...new Set(xs)].sort((a, b) => a - b);
  const buckets: number[] = [sorted[0]];
  for (const x of sorted) {
    if (x - buckets[buckets.length - 1] > tolerance) {
      buckets.push(x);
    }
  }
  return buckets;
}

// ---------------------------------------------------------------------------
// Helper: assign item to nearest bucket index
// ---------------------------------------------------------------------------
function assignCol(x: number, buckets: number[]): number {
  let best = 0;
  let bestDist = Math.abs(x - buckets[0]);
  for (let i = 1; i < buckets.length; i++) {
    const d = Math.abs(x - buckets[i]);
    if (d < bestDist) {
      bestDist = d;
      best = i;
    }
  }
  return best;
}

// ---------------------------------------------------------------------------
// Helper: detect if a set of rows looks like a table
// (multi-column rows with consistent column alignment)
// ---------------------------------------------------------------------------
function isTableLike(rows: TextRow[]): boolean {
  if (rows.length < 2) return false;
  const multiCol = rows.filter((r) => r.items.length >= 2).length;
  return multiCol >= Math.ceil(rows.length * 0.4);
}

// ---------------------------------------------------------------------------
// Main: write positioned rows to worksheet, returns updated currentRow
// ---------------------------------------------------------------------------
function writePositionedRows(
  ws: ExcelJS.Worksheet,
  rows: TextRow[],
  startRow: number
): number {
  if (rows.length === 0) return startRow;

  const colBuckets = buildColBuckets(rows);
  const isTable = colBuckets.length >= 2 && isTableLike(rows);
  let cur = startRow;

  // Update column widths
  colBuckets.forEach((_, idx) => {
    const excelColIdx = idx + 1;
    const maxLen = Math.max(
      ...rows.map((r) => {
        const item = r.items.find((i) => assignCol(i.x, colBuckets) === idx);
        return item ? item.text.length : 0;
      }),
      8
    );
    const col = ws.getColumn(excelColIdx);
    const newWidth = Math.min(maxLen + 2, 50);
    if (!col.width || col.width < newWidth) {
      col.width = newWidth;
    }
  });

  rows.forEach((row, rowIdx) => {
    const excelRow = ws.getRow(cur);

    row.items.forEach((item) => {
      const colIdx = assignCol(item.x, colBuckets) + 1;
      const cell = excelRow.getCell(colIdx);
      // Concatenate if multiple items map to same cell
      cell.value = cell.value ? `${cell.value} ${item.text}` : item.text;

      if (isTable) {
        cell.border = {
          top: { style: "thin", color: { argb: "FFCCCCCC" } },
          left: { style: "thin", color: { argb: "FFCCCCCC" } },
          bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
          right: { style: "thin", color: { argb: "FFCCCCCC" } },
        };
        if (rowIdx === 0) {
          cell.font = { bold: true };
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFF0F4FF" },
          };
        }
      }
    });

    excelRow.commit();
    cur++;
  });

  return cur;
}

// ---------------------------------------------------------------------------
// Main: write a structured table (from pdf-parse getTable) to worksheet
// ---------------------------------------------------------------------------
function writeStructuredTable(
  ws: ExcelJS.Worksheet,
  table: string[][],
  startRow: number
): number {
  if (table.length === 0) return startRow;

  const colCount = Math.max(...table.map((r) => r.length));

  // Auto-fit column widths
  for (let c = 0; c < colCount; c++) {
    const maxLen = Math.max(
      ...table.map((r) => (r[c] ?? "").length),
      8
    );
    const excelCol = ws.getColumn(c + 1);
    const newWidth = Math.min(maxLen + 2, 50);
    if (!excelCol.width || excelCol.width < newWidth) {
      excelCol.width = newWidth;
    }
  }

  let cur = startRow;
  table.forEach((row, rowIdx) => {
    const excelRow = ws.getRow(cur);
    for (let c = 0; c < colCount; c++) {
      const cell = excelRow.getCell(c + 1);
      cell.value = row[c] ?? "";
      cell.border = {
        top: { style: "thin", color: { argb: "FFAAAAAA" } },
        left: { style: "thin", color: { argb: "FFAAAAAA" } },
        bottom: { style: "thin", color: { argb: "FFAAAAAA" } },
        right: { style: "thin", color: { argb: "FFAAAAAA" } },
      };
      if (rowIdx === 0) {
        cell.font = { bold: true };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFD6E4F7" },
        };
      }
    }
    excelRow.commit();
    cur++;
  });

  return cur;
}

// ---------------------------------------------------------------------------
// Route handler
// ---------------------------------------------------------------------------
export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file");

    if (!file || typeof file === "string") {
      return NextResponse.json(
        { error: "ファイルが見つかりません" },
        { status: 400 }
      );
    }

    const buffer = Buffer.from(await file.arrayBuffer());

    // --- Parse PDF ---
    const parser = new PDFParse({ data: buffer });

    const [tableResult, infoResult, positionedPages] = await Promise.all([
      parser.getTable().catch(() => null),
      parser.getInfo().catch(() => null),
      extractPositionedText(buffer),
    ]);

    await parser.destroy();

    // --- Build Excel ---
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "SAYAMA";
    workbook.created = new Date();

    const ws = workbook.addWorksheet("PDF内容");
    ws.properties.defaultRowHeight = 15;

    let currentRow = 1;
    const numPages = positionedPages.size;

    for (let p = 1; p <= numPages; p++) {
      // Page separator (except first page)
      if (p > 1) {
        ws.getRow(currentRow).getCell(1).value = `── ページ ${p} ──`;
        ws.getRow(currentRow).getCell(1).font = {
          color: { argb: "FF999999" },
          italic: true,
          size: 9,
        };
        ws.getRow(currentRow).commit();
        currentRow++;
      }

      const pageTables =
        tableResult?.pages.find((pg) => pg.num === p)?.tables ?? [];
      const textRows = positionedPages.get(p) ?? [];

      if (pageTables.length > 0) {
        // --- Page has detected tables ---
        // Write tables with full structure
        for (const table of pageTables) {
          if (!table || table.length === 0) continue;
          currentRow = writeStructuredTable(ws, table, currentRow);
          currentRow++; // blank row after each table
        }

        // Also write any remaining text rows (non-table text)
        // Use positioned text, but skip content already covered by tables
        const tableTexts = new Set(
          pageTables
            .flat()
            .flat()
            .filter(Boolean)
            .map((s: string) => s.trim())
        );
        const remainingRows = textRows.filter(
          (row) =>
            !row.items.every((item) => tableTexts.has(item.text.trim()))
        );
        if (remainingRows.length > 0) {
          currentRow = writePositionedRows(ws, remainingRows, currentRow);
        }
      } else {
        // --- No tables: use coordinate-based layout ---
        currentRow = writePositionedRows(ws, textRows, currentRow);
      }

      currentRow++; // blank row between pages
    }

    // --- Metadata sheet ---
    const metaWs = workbook.addWorksheet("メタデータ");
    metaWs.columns = [
      { header: "項目", key: "k", width: 20 },
      { header: "値", key: "v", width: 60 },
    ];
    const metaHeader = metaWs.getRow(1);
    metaHeader.font = { bold: true };
    metaHeader.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFDDE8F0" },
    };
    metaHeader.commit();

    const info = infoResult?.info as Record<string, unknown> | undefined;
    metaWs.addRow({ k: "ページ数", v: numPages });
    metaWs.addRow({ k: "変換日時", v: new Date().toLocaleString("ja-JP") });
    if (info?.Title) metaWs.addRow({ k: "タイトル", v: info.Title });
    if (info?.Author) metaWs.addRow({ k: "作成者", v: info.Author });
    if (info?.Subject) metaWs.addRow({ k: "件名", v: info.Subject });
    if (info?.Creator) metaWs.addRow({ k: "作成ソフト", v: info.Creator });

    const xlsxBuffer = await workbook.xlsx.writeBuffer();

    return new NextResponse(xlsxBuffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="converted.xlsx"',
      },
    });
  } catch (err) {
    console.error("変換エラー:", err);
    return NextResponse.json(
      { error: "変換処理中にエラーが発生しました" },
      { status: 500 }
    );
  }
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
