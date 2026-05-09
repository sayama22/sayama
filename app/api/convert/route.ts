import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

// ---------------------------------------------------------------------------
// pdf-parse v1: uses pdfjs-dist v2, runs in the main thread (no workers),
// no DOMMatrix requirement → reliable in all Node.js / serverless environments.
// We access it via the npm alias "pdf-parse-v1" defined in package.json.
// ---------------------------------------------------------------------------
// eslint-disable-next-line @typescript-eslint/no-require-imports
const pdfParse = require("pdf-parse-v1") as (
  data: Buffer,
  options?: { pagerender?: (pageData: unknown) => Promise<string> }
) => Promise<{ numpages: number; info: Record<string, unknown> }>;

// ---------------------------------------------------------------------------
// Normalize trailing-minus sign: "1234-" → "-1234", "123.45-" → "-123.45"
// Returns a negative number if the pattern matches, otherwise the raw string.
// ---------------------------------------------------------------------------
function normalizeSign(raw: string): string | number {
  const trimmed = raw.trim();
  // Match digits (with optional thousand-separators/decimal) followed by "-"
  const m = trimmed.match(/^([\d,]+(\.\d+)?)-$/);
  if (m) {
    const numStr = m[1].replace(/,/g, "");
    const n = parseFloat(numStr);
    if (!isNaN(n)) return -n;
  }
  return raw;
}

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------
interface PositionedItem {
  text: string;
  x: number;
  y: number; // "screen" y: increases from top to bottom
}

interface TextRow {
  items: PositionedItem[];
  y: number;
}

// ---------------------------------------------------------------------------
// Extract text items with coordinates for every page (via pagerender hook)
// ---------------------------------------------------------------------------
async function extractPages(buffer: Buffer): Promise<{
  pages: PositionedItem[][];
  numpages: number;
  info: Record<string, unknown>;
}> {
  const pages: PositionedItem[][] = [];

  const result = await pdfParse(buffer, {
    async pagerender(pageData: unknown) {
      const pd = pageData as {
        getTextContent(): Promise<{
          items: Array<{ str: string; transform: number[] }>;
        }>;
      };
      const tc = await pd.getTextContent();
      const items: PositionedItem[] = tc.items
        .filter((it) => it.str && it.str.trim() !== "")
        .map((it) => ({
          text: it.str,
          x: Math.round(it.transform[4]),
          // PDF y-axis: 0 at bottom, increases upward → negate for screen order
          y: -Math.round(it.transform[5]),
        }));
      pages.push(items);
      return tc.items.map((i) => i.str).join(" ");
    },
  });

  return { pages, numpages: result.numpages, info: result.info ?? {} };
}

// ---------------------------------------------------------------------------
// Group positioned items into rows (similar y within tolerance)
// ---------------------------------------------------------------------------
function groupIntoRows(items: PositionedItem[], tolerance = 3): TextRow[] {
  if (items.length === 0) return [];
  const sorted = [...items].sort((a, b) => a.y - b.y || a.x - b.x);
  const rows: TextRow[] = [];
  let current: PositionedItem[] = [sorted[0]];
  let currentY = sorted[0].y;

  for (let i = 1; i < sorted.length; i++) {
    const it = sorted[i];
    if (Math.abs(it.y - currentY) <= tolerance) {
      current.push(it);
    } else {
      rows.push({ items: current.sort((a, b) => a.x - b.x), y: currentY });
      current = [it];
      currentY = it.y;
    }
  }
  if (current.length > 0) {
    rows.push({ items: current.sort((a, b) => a.x - b.x), y: currentY });
  }
  return rows;
}

// ---------------------------------------------------------------------------
// Build column buckets from x positions
// ---------------------------------------------------------------------------
function buildColBuckets(rows: TextRow[], tolerance = 15): number[] {
  const xs = rows.flatMap((r) => r.items.map((i) => i.x));
  if (xs.length === 0) return [];
  const sorted = [...new Set(xs)].sort((a, b) => a - b);
  const buckets: number[] = [sorted[0]];
  for (const x of sorted) {
    if (x - buckets[buckets.length - 1] > tolerance) buckets.push(x);
  }
  return buckets;
}

// ---------------------------------------------------------------------------
// Assign x to nearest bucket index
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
// Does this set of rows form a tabular structure?
// ---------------------------------------------------------------------------
function isTabular(rows: TextRow[]): boolean {
  if (rows.length < 2) return false;
  const multiCol = rows.filter((r) => r.items.length >= 2).length;
  return multiCol >= Math.ceil(rows.length * 0.4);
}

// ---------------------------------------------------------------------------
// Split rows into blocks: consecutive "tabular" runs vs single-column text
// ---------------------------------------------------------------------------
type Block =
  | { type: "table"; rows: TextRow[] }
  | { type: "text"; rows: TextRow[] };

function splitBlocks(rows: TextRow[]): Block[] {
  if (rows.length === 0) return [];
  const blocks: Block[] = [];
  let i = 0;

  while (i < rows.length) {
    // Try to extend a table block starting at i
    let j = i + 1;
    while (j < rows.length) {
      const candidate = rows.slice(i, j + 1);
      if (getColCount(candidate) >= 2) {
        j++;
      } else {
        break;
      }
    }
    const candidate = rows.slice(i, j);
    if (candidate.length >= 2 && isTabular(candidate)) {
      blocks.push({ type: "table", rows: candidate });
      i = j;
    } else {
      // Collect consecutive single-column text rows
      const textRows: TextRow[] = [];
      while (i < rows.length) {
        const next = rows.slice(i, i + 2);
        if (next.length >= 2 && getColCount(next) >= 2 && isTabular(next)) break;
        textRows.push(rows[i]);
        i++;
      }
      if (textRows.length > 0) blocks.push({ type: "text", rows: textRows });
    }
  }
  return blocks;
}

function getColCount(rows: TextRow[]): number {
  return buildColBuckets(rows).length;
}

// ---------------------------------------------------------------------------
// Write a "table" block to worksheet
// ---------------------------------------------------------------------------
function writeTableBlock(
  ws: ExcelJS.Worksheet,
  rows: TextRow[],
  startRow: number
): number {
  const buckets = buildColBuckets(rows);
  const colCount = buckets.length;

  // Auto-fit column widths
  buckets.forEach((_, idx) => {
    const maxLen = Math.max(
      ...rows.map((r) => {
        const it = r.items.find((i) => assignCol(i.x, buckets) === idx);
        return it ? it.text.length : 0;
      }),
      6
    );
    const col = ws.getColumn(idx + 1);
    const w = Math.min(maxLen + 2, 50);
    if (!col.width || col.width < w) col.width = w;
  });

  let cur = startRow;
  rows.forEach((row, rowIdx) => {
    const excelRow = ws.getRow(cur);
    const cells = new Array<string>(colCount).fill("");
    row.items.forEach((it) => {
      const idx = assignCol(it.x, buckets);
      cells[idx] = cells[idx] ? `${cells[idx]} ${it.text}` : it.text;
    });
    cells.forEach((val, idx) => {
      const cell = excelRow.getCell(idx + 1);
      cell.value = normalizeSign(val) as ExcelJS.CellValue;
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
          fgColor: { argb: "FFD6E4F7" },
        };
      }
    });
    excelRow.commit();
    cur++;
  });
  return cur;
}

// ---------------------------------------------------------------------------
// Write a "text" block (with optional indent as column offset)
// ---------------------------------------------------------------------------
function writeTextBlock(
  ws: ExcelJS.Worksheet,
  rows: TextRow[],
  startRow: number
): number {
  const maxLen = Math.max(
    ...rows.flatMap((r) => r.items.map((i) => i.text.length)),
    10
  );
  const col1 = ws.getColumn(1);
  const w = Math.min(maxLen + 4, 100);
  if (!col1.width || col1.width < w) col1.width = w;

  let cur = startRow;
  rows.forEach((row) => {
    const excelRow = ws.getRow(cur);
    // Use indent: items starting far right → indent to column 2
    const firstX = row.items[0]?.x ?? 0;
    const pageXs = rows.flatMap((r) => r.items.map((i) => i.x));
    const minX = Math.min(...pageXs);
    const colIdx = firstX - minX > 30 ? 2 : 1;
    const cell = excelRow.getCell(colIdx);
    const joined = row.items.map((i) => i.text).join(" ");
    cell.value = normalizeSign(joined) as ExcelJS.CellValue;
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
    const { pages, numpages, info } = await extractPages(buffer);

    // Build Excel
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "SAYAMA";
    workbook.created = new Date();

    const ws = workbook.addWorksheet("PDF内容");
    ws.properties.defaultRowHeight = 15;

    let currentRow = 1;

    for (let p = 0; p < pages.length; p++) {
      const pageNum = p + 1;
      const items = pages[p];

      if (p > 0) {
        const sep = ws.getRow(currentRow);
        sep.getCell(1).value = `── ページ ${pageNum} ──`;
        sep.getCell(1).font = {
          color: { argb: "FF999999" },
          italic: true,
          size: 9,
        };
        sep.commit();
        currentRow++;
      }

      const rows = groupIntoRows(items);
      const blocks = splitBlocks(rows);

      for (const block of blocks) {
        if (block.type === "table") {
          currentRow = writeTableBlock(ws, block.rows, currentRow);
          currentRow++; // blank row after table
        } else {
          currentRow = writeTextBlock(ws, block.rows, currentRow);
        }
      }

      currentRow++; // blank row between pages
    }

    // Metadata sheet
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

    metaWs.addRow({ k: "ページ数", v: numpages });
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
    const msg =
      err instanceof Error ? err.message : "変換処理中にエラーが発生しました";
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
