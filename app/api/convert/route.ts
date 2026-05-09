import { NextRequest, NextResponse } from "next/server";
import { PDFParse } from "pdf-parse";
import ExcelJS from "exceljs";

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file");

    if (!file || typeof file === "string") {
      return NextResponse.json({ error: "ファイルが見つかりません" }, { status: 400 });
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    // PDF解析 (pdf-parse v2 API)
    const parser = new PDFParse({ data: buffer });
    const textResult = await parser.getText();
    const infoResult = await parser.getInfo();

    const rawText = textResult.text ?? "";
    const numpages = infoResult.info?.numPages ?? 0;

    // テキストを行に分割して整形
    const lines = rawText
      .split("\n")
      .map((line) => line.trim())
      .filter((line) => line.length > 0);

    // Excelワークブック作成
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "SAYAMA";
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet("PDF内容");

    worksheet.columns = [
      { header: "行番号", key: "lineNumber", width: 10 },
      { header: "テキスト", key: "text", width: 100 },
    ];

    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE2EFDA" },
    };
    headerRow.commit();

    lines.forEach((line, index) => {
      worksheet.addRow({ lineNumber: index + 1, text: line });
    });

    // メタデータシート
    const metaSheet = workbook.addWorksheet("メタデータ");
    metaSheet.columns = [
      { header: "項目", key: "key", width: 20 },
      { header: "値", key: "value", width: 60 },
    ];
    const metaHeader = metaSheet.getRow(1);
    metaHeader.font = { bold: true };
    metaHeader.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFDDE8F0" },
    };
    metaHeader.commit();

    metaSheet.addRow({ key: "ページ数", value: numpages });
    metaSheet.addRow({ key: "総行数", value: lines.length });
    metaSheet.addRow({ key: "変換日時", value: new Date().toLocaleString("ja-JP") });

    const pdfInfo = infoResult.info as Record<string, unknown> | undefined;
    if (pdfInfo?.Title) metaSheet.addRow({ key: "タイトル", value: pdfInfo.Title });
    if (pdfInfo?.Author) metaSheet.addRow({ key: "作成者", value: pdfInfo.Author });

    await parser.destroy();

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
