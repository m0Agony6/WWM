import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import multer from "multer";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { GoogleGenAI } from "@google/genai";
import cors from "cors";

const app = express();
const PORT = 3000;
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json());
app.use(cors());

// Initialize AI
const genAI = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

// --- Processing Helpers ---

function normalizeText(val: any): string {
  if (val === null || val === undefined) return "";
  return String(val).replace(/\u3000/g, " ").replace(/\xa0/g, " ").trim();
}

function parseExcelDate(val: any): Date | null {
  if (val instanceof Date) return val;
  if (typeof val === "number") {
    // Excel date to JS date
    return new Date(Math.round((val - 25569) * 86400 * 1000));
  }
  if (typeof val === "string") {
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function parseExcelTime(val: any): { h: number, m: number } | null {
  if (val instanceof Date) {
    return { h: val.getHours(), m: val.getMinutes() };
  }
  if (typeof val === "number") {
    const totalMinutes = Math.round(val * 24 * 60);
    return { h: Math.floor(totalMinutes / 60), m: totalMinutes % 60 };
  }
  if (typeof val === "string") {
    const match = val.match(/(\d{1,2})[:：](\d{1,2})/);
    if (match) return { h: parseInt(match[1]), m: parseInt(match[2]) };
  }
  return null;
}

function getReportCode(val: any): string {
  const text = normalizeText(val);
  const matches = text.match(/G[0-9A-Za-z]+/g);
  if (matches) return matches[matches.length - 1];
  const lines = text.split("\n").filter(l => l.trim());
  return lines.length > 0 ? lines[lines.length - 1].trim() : "";
}

function serialMatches(cellValue: any, targetSerials: string[]): boolean {
  const left = normalizeText(cellValue);
  if (!left) return false;
  return targetSerials.some(s => normalizeText(s) === left);
}

function sanitizeFileName(name: string): string {
  return name.replace(/[<>:"/\\|?*]/g, '_').trim();
}

// --- Interface ---

interface LedgerRecord {
  sheetName: string;
  serial: string;
  ring: string;
  steelModel: string;
  turnCode: string;
  specialNote: string;
  productionDate: Date | null;
  curingStart: { h: number, m: number } | null;
  curingEnd: { h: number, m: number } | null;
  sprayDate: Date | null;
  report6: string;
  report8: string;
  report10: string;
  report12: string;
  report20: string;
}

// --- API Routes ---

app.post("/api/process", upload.fields([
  { name: 'ledger', maxCount: 1 },
  { name: 'templateLeft', maxCount: 1 },
  { name: 'templateRight', maxCount: 1 }
]), async (req, res) => {
  try {
    const files = req.files as { [fieldname: string]: Express.Multer.File[] };
    if (!files.ledger || !files.templateLeft || !files.templateRight) {
      return res.status(400).json({ error: "Missing required files" });
    }

    const { serials, config, sourceSheet } = JSON.parse(req.body.data || "{}");
    const serialList: string[] = [];
    String(serials).split(/[,\s，、;；]/).filter(s => s.trim()).forEach(part => {
        if (part.includes("-")) {
            const rangeParts = part.split("-");
            if (rangeParts.length === 2) {
                const start = parseInt(rangeParts[0].replace(/\D/g, "").trim());
                const end = parseInt(rangeParts[1].replace(/\D/g, "").trim());
                if (!isNaN(start) && !isNaN(end)) {
                    const min = Math.min(start, end);
                    const max = Math.max(start, end);
                    if (max - min > 2000) throw new Error("Range too large (max 2000)");
                    for (let i = min; i <= max; i++) serialList.push(String(i));
                } else {
                    serialList.push(part);
                }
            } else {
                serialList.push(part);
            }
        } else {
            serialList.push(part);
        }
    });

    if (serialList.length === 0) return res.status(400).json({ error: "No serial numbers provided" });

    // Load Ledger
    const ledgerWorkbook = new ExcelJS.Workbook();
    await ledgerWorkbook.xlsx.load(files.ledger[0].buffer);
    
    // Find sheets to process
    const sheetsToProcess: ExcelJS.Worksheet[] = [];
    if (sourceSheet) {
        const s = ledgerWorkbook.getWorksheet(sourceSheet);
        if (s) sheetsToProcess.push(s);
    }
    
    if (sheetsToProcess.length === 0) {
        for (const s of ledgerWorkbook.worksheets) {
            const row2 = s.getRow(2);
            let hasSerialCol = false;
            row2.eachCell((cell) => {
                if (normalizeText(cell.value).includes("资料流水号")) hasSerialCol = true;
            });
            if (hasSerialCol) {
                sheetsToProcess.push(s);
                if (!sourceSheet) break; 
            }
        }
    }
    
    if (sheetsToProcess.length === 0) {
        return res.status(400).json({ error: "Could not find a valid ledger sheet." });
    }

    const sheet = sheetsToProcess[0];
    const records: LedgerRecord[] = [];
    
    // Find serial column
    let serialColIndex = 3;
    const headerRow = sheet.getRow(2);
    headerRow.eachCell((cell, colNumber) => {
        if (normalizeText(cell.value).includes("资料流水号")) serialColIndex = colNumber;
    });

    sheet.eachRow((row, rowNum) => {
        if (rowNum < 3 || row.hidden) return;
        const cellValue = row.getCell(serialColIndex).value;
        const normalizedCell = normalizeText(cellValue);
        
        if (serialMatches(cellValue, serialList)) {
            const ringValue = row.getCell(serialColIndex + 1).value; 
            if (!ringValue) return;

            records.push({
                sheetName: sheet.name,
                serial: normalizedCell,
                ring: normalizeText(ringValue),
                steelModel: normalizeText(row.getCell(serialColIndex + 2).value),
                turnCode: normalizeText(row.getCell(serialColIndex + 3).value),
                specialNote: normalizeText(row.getCell(serialColIndex + 4).value),
                productionDate: parseExcelDate(row.getCell(serialColIndex + 5).value),
                curingStart: parseExcelTime(row.getCell(serialColIndex + 11).value),
                curingEnd: parseExcelTime(row.getCell(serialColIndex + 12).value),
                sprayDate: parseExcelDate(row.getCell(serialColIndex + 13).value),
                report6: getReportCode(row.getCell(serialColIndex + 15).value),
                report8: getReportCode(row.getCell(serialColIndex + 16).value),
                report10: getReportCode(row.getCell(serialColIndex + 17).value),
                report12: getReportCode(row.getCell(serialColIndex + 18).value),
                report20: getReportCode(row.getCell(serialColIndex + 20).value),
            });
        }
    });

    if (records.length === 0) return res.status(404).json({ error: "No records found matching serial numbers." });

    const zip = new JSZip();
    let successCount = 0;
    
    for (const record of records) {
        try {
            const isLeft = record.turnCode.toUpperCase() === "L" || normalizeText(record.specialNote).includes("左转");
            const templateBuffer = isLeft ? files.templateLeft[0].buffer : files.templateRight[0].buffer;
            
            const templateWorkbook = new ExcelJS.Workbook();
            await templateWorkbook.xlsx.load(templateBuffer);
            
            const ws = (idx: number) => templateWorkbook.getWorksheet(idx);
            const rws = ws(1), sws = ws(2), hws = ws(3), pws = ws(4), mws = ws(5), cws = ws(7), fws = ws(8);

            const ringPadded = record.ring.replace(/\D/g, "").padStart(5, '0');
            const ringNumeric = parseInt(record.ring.replace(/\D/g, "")) || record.ring;
            
            if(rws) {
                rws.getCell("G1").value = record.serial;
                const ringDesc = normalizeText(record.specialNote).includes("标准") ? "直线环" : (normalizeText(record.specialNote) || "直线环");
                rws.getCell("D4").value = `第${ringPadded}环${record.steelModel}型${ringDesc}`;
                const prefixes = config.reportPrefixes || { "G8": "07062500C700308", "G9": "07062500C700309", "G10": "07062500C700311", "G11": "07062500C700310", "G12": "07062500C500100", "G15": "07062500C511100" };
                Object.entries(prefixes).forEach(([cell, pre]) => { rws.getCell(cell).value = `${pre}${ringPadded}`; });
                const supervisors = config.supervisors || {};
                if (supervisors[record.sheetName]) rws.getCell("B28").value = supervisors[record.sheetName];
            }
            if(sws) { sws.getCell("D4").value = ringNumeric; sws.getCell("P3").value = record.productionDate; }
            if(hws) {
                hws.getCell("A14").value = `规格型号:HRB400E 20 复试报告编号：${record.report20}`;
                hws.getCell("A15").value = `规格型号:HRB400E 12 复试报告编号：${record.report12}`;
                hws.getCell("A16").value = `规格型号：HPB300 10 复试报告编号：${record.report10}`;
                hws.getCell("A17").value = `规格型号：HPB300 8  复试报告编号：${record.report8}`;
                hws.getCell("A18").value = `规格型号：HPB300 6  复试报告编号：${record.report6}`;
                hws.getCell("I3").value = record.productionDate; hws.getCell("I4").value = ringNumeric;
            }
            if(pws) {
                pws.getCell("B10").value = `第${ringPadded}环`; pws.getCell("B12").value = ringNumeric; pws.getCell("B22").value = `${ringPadded}-02`; pws.getCell("B32").value = ringNumeric;
                pws.getCell("B6").value = record.productionDate; pws.getCell("B16").value = record.productionDate; pws.getCell("B20").value = `第${ringPadded}环`; pws.getCell("B26").value = record.productionDate; pws.getCell("B30").value = `第${ringPadded}环`;
            }
            if(mws) {
                const moldSuffix = isLeft ? "L" : (record.turnCode.toUpperCase() === "R" || normalizeText(record.specialNote).includes("右转") ? "R" : "");
                const moldPrefixes = config.moldPrefixes || { "D4": "B1", "F4": "B2", "H4": "B3", "J4": "L1", "L4": "L2", "O4": "F" };
                Object.entries(moldPrefixes).forEach(([cell, pre]) => { mws.getCell(cell).value = `${pre}${moldSuffix}-01`; });
                mws.getCell("O3").value = record.productionDate;
            }
            if(cws && record.sprayDate && record.curingStart) {
                const startDt = new Date(record.sprayDate);
                startDt.setHours(record.curingStart.h, record.curingStart.m, 0, 0);
                for(let i = 0; i < 5; i++) {
                    const row = 6 + i; const dt = new Date(startDt.getTime() + i * 60 * 60 * 1000);
                    cws.getCell(`A${row}`).value = dt.getFullYear(); cws.getCell(`B${row}`).value = dt.getMonth() + 1; cws.getCell(`C${row}`).value = dt.getDate();
                    cws.getCell(`D${row}`).value = dt; cws.getCell(`D${row}`).numFmt = "h:mm";
                }
            }
            if(fws) { fws.getCell("P2").value = record.sprayDate; }

            const outBuffer = await templateWorkbook.xlsx.writeBuffer();
            const fileName = sanitizeFileName(`管片检验批-${record.serial}-${ringNumeric}.xlsx`);
            zip.file(fileName, outBuffer as Buffer);
            successCount++;
        } catch (e) {
            console.error(`Error processing ${record.serial}:`, e);
        }
    }

    if (successCount === 0) return res.status(404).json({ error: "Processing failed for all records." });

    const zipBuffer = await zip.generateAsync({ 
        type: "nodebuffer",
        compression: "DEFLATE",
        compressionOptions: { level: 6 }
    });
    
    res.set({
        'Content-Type': 'application/zip',
        'Content-Disposition': `attachment; filename="Pipe_Reports_${Date.now()}.zip"`,
        'Content-Length': zipBuffer.length
    });
    res.send(zipBuffer);

  } catch (error: any) {
    console.error(error);
    res.status(500).json({ error: error.message });
  }
});

// AI Assistant
app.post("/api/chat", async (req, res) => {
  try {
    const { message, context } = req.body;
    const model = (genAI as any).getGenerativeModel({ model: "gemini-2.0-flash" });
    const prompt = `Assistant context: ${JSON.stringify(context)}\nUser: ${message}\nAnswer in Chinese.`;
    const result = await model.generateContent(prompt);
    res.json({ reply: result.response.text() });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
});

async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({ server: { middlewareMode: true }, appType: "spa" });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => res.sendFile(path.join(distPath, "index.html")));
  }
  app.listen(PORT, "0.0.0.0", () => console.log(`Server running on http://localhost:${PORT}`));
}

startServer();
