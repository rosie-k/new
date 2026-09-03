/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import ExcelJS from "exceljs";
import { GoogleGenAI, Type } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = 3000;

// Enable JSON body parsing up to 50mb (for handling large Excel Base64 strings)
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ limit: "50mb", extended: true }));

// Initialize Gemini SDK lazily to avoid crashing on startup if key is missing
let aiClient: GoogleGenAI | null = null;
function getGeminiClient(): GoogleGenAI | null {
  if (!aiClient) {
    const key = process.env.GEMINI_API_KEY;
    if (key && key !== "MY_GEMINI_API_KEY") {
      aiClient = new GoogleGenAI({
        apiKey: key,
        httpOptions: {
          headers: {
            "User-Agent": "aistudio-build",
          },
        },
      });
    }
  }
  return aiClient;
}

// Target standard facilities
const TARGET_FACILITIES = [
  "테니스장A",
  "테니스장B",
  "테니스장C",
  "테니스장D",
  "복합운동장",
  "농구장A",
  "농구장B",
  "농구장C",
];

interface Booking {
  id: string;
  bookingNo?: string;
  facilityName: string;
  resolvedFacility: string;
  bookerName: string;
  groupName?: string;
  date: string;
  startTime: string;
  endTime: string;
  phoneNumber?: string;
  status?: string;
  isNightHighlighted?: boolean;
  isDuplicate?: boolean;
  isExceeded?: boolean;
  exceededNote?: string;
}

function computeExceededBookings(bookingsList: Booking[]): Booking[] {
  const groups: Record<string, Record<string, Booking[]>> = {};

  bookingsList.forEach(b => {
    const d = (b.date || "").trim();
    const id = (b.bookerName || "").trim();
    if (!groups[d]) groups[d] = {};
    if (!groups[d][id]) groups[d][id] = [];
    groups[d][id].push(b);
  });

  const exceededMap: Record<string, { isExceeded: boolean; exceededNote: string }> = {};

  Object.keys(groups).forEach(d => {
    Object.keys(groups[d]).forEach(id => {
      const list = groups[d][id];
      list.sort((a, b) => {
        const timeA = a.startTime || "00:00";
        const timeB = b.startTime || "00:00";
        return timeA.localeCompare(timeB);
      });

      let accumulatedHours = 0;
      list.forEach(b => {
        const statusVal = b.status || "";
        const isCancelled = statusVal.includes("취소") || statusVal.includes("반려") || statusVal.includes("불허");

        if (isCancelled) {
          exceededMap[b.id] = {
            isExceeded: false,
            exceededNote: ""
          };
          return; // Skip accumulating duration for cancelled bookings
        }

        const startParts = (b.startTime || "00:00").split(":");
        const endParts = (b.endTime || "00:00").split(":");
        const startH = parseInt(startParts[0] || "0", 10);
        const startM = parseInt(startParts[1] || "0", 10);
        const endH = parseInt(endParts[0] || "0", 10);
        const endM = parseInt(endParts[1] || "0", 10);

        const duration = Math.round(((endH - startH) + (endM - startM) / 60) * 10000) / 10000;

        const roundedAccumulated = Math.round(accumulatedHours * 10000) / 10000;
        const roundedTotal = Math.round((accumulatedHours + duration) * 10000) / 10000;

        if (roundedAccumulated >= 4) {
          exceededMap[b.id] = {
            isExceeded: true,
            exceededNote: "초과예약분"
          };
        } else if (roundedTotal > 4) {
          exceededMap[b.id] = {
            isExceeded: true,
            exceededNote: "초과예약분"
          };
        } else {
          exceededMap[b.id] = {
            isExceeded: false,
            exceededNote: ""
          };
        }
        accumulatedHours += duration;
      });
    });
  });

  return bookingsList.map(b => {
    const calc = exceededMap[b.id];
    return {
      ...b,
      isExceeded: calc ? calc.isExceeded : false,
      exceededNote: calc ? calc.exceededNote : ""
    };
  });
}

// Helper to normalize dates (returns YYYY-MM-DD or empty)
function normalizeDate(cellValue: any): string {
  if (!cellValue) return "";

  let dateObj: Date | null = null;

  if (cellValue instanceof Date) {
    dateObj = cellValue;
  } else if (typeof cellValue === "number" && cellValue > 30000 && cellValue < 60000) {
    // Excel numeric date serial representation
    const days = cellValue - (cellValue > 60 ? 2 : 1);
    dateObj = new Date(Date.UTC(1900, 0, days));
  } else if (typeof cellValue === "object") {
    // ExcelJS formula result parsing
    if (cellValue.result instanceof Date) {
      dateObj = cellValue.result;
    } else if (typeof cellValue.result === "number" && cellValue.result > 30000 && cellValue.result < 60000) {
      const days = cellValue.result - (cellValue.result > 60 ? 2 : 1);
      dateObj = new Date(Date.UTC(1900, 0, days));
    } else if (cellValue.result) {
      const str = cellValue.result.toString().trim().replace(/\s+/g, "");
      const match = str.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})\.?/);
      if (match) {
        const y = match[1];
        const m = match[2].padStart(2, "0");
        const d = match[3].padStart(2, "0");
        return `${y}-${m}-${d}`;
      }
    }
  }

  if (dateObj && !isNaN(dateObj.getTime())) {
    // Prevent timezone-shift by using getUTC* methods for parsed dates
    const y = dateObj.getUTCFullYear();
    const m = String(dateObj.getUTCMonth() + 1).padStart(2, "0");
    const d = String(dateObj.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  
  // Remove all whitespaces first to handle things like "2026. 7. 3." or "2026 .07 .03"
  const str = cellValue.toString().trim().replace(/\s+/g, "");
  
  // Try to parse YYYY-MM-DD, YYYY.MM.DD, YYYY/MM/DD (optionally ending with a dot)
  const match = str.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})\.?/);
  if (match) {
    const y = match[1];
    const m = match[2].padStart(2, "0");
    const d = match[3].padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  // Try simple MM-DD (optionally ending with a dot)
  const simpleMatch = str.match(/(\d{1,2})[-./](\d{1,2})\.?/);
  if (simpleMatch) {
    const currentYear = new Date().getFullYear();
    const m = simpleMatch[1].padStart(2, "0");
    const d = simpleMatch[2].padStart(2, "0");
    return `${currentYear}-${m}-${d}`;
  }
  return str;
}

// Timezone-agnostic local date parser to prevent day-of-week shift issues
function parseLocalDate(dateStr: string): Date {
  if (!dateStr) return new Date();
  const cleanedStr = dateStr.toString().trim().replace(/\s+/g, "").replace(/[./]$/, "");
  const cleaned = cleanedStr.replace(/[./]/g, "-");
  const parts = cleaned.split("-");
  if (parts.length === 3) {
    const year = Number(parts[0]);
    const month = Number(parts[1]);
    const day = Number(parts[2]);
    return new Date(year, month - 1, day);
  }
  if (parts.length === 2) {
    const year = new Date().getFullYear();
    const month = Number(parts[0]);
    const day = Number(parts[1]);
    return new Date(year, month - 1, day);
  }
  const parsed = new Date(dateStr);
  if (!isNaN(parsed.getTime())) return parsed;
  return new Date();
}

// Helper to extract clean text/string value from ExcelJS Cell (supporting formulas and Rich Text runs)
function getCellValueAsString(cell: ExcelJS.Cell | any): string {
  if (!cell) return "";
  const val = cell.value;
  if (val === undefined || val === null) return "";
  if (typeof val === "object") {
    if ("richText" in val && Array.isArray(val.richText)) {
      return val.richText.map((run: any) => run.text || "").join("");
    }
    if ("result" in val) {
      return val.result !== undefined && val.result !== null ? val.result.toString().trim() : "";
    }
    if ("text" in val) {
      return val.text !== undefined && val.text !== null ? val.text.toString().trim() : "";
    }
    return JSON.stringify(val);
  }
  return val.toString().trim();
}

function formatPhoneNumber(phone: any): string {
  if (!phone) return "";
  let clean = phone.toString().trim().replace(/[^0-9]/g, "");
  
  if (clean.startsWith("82")) {
    clean = "0" + clean.slice(2);
  }
  if (clean.length === 10 && clean.startsWith("1")) {
    clean = "0" + clean;
  }

  if (clean.length === 11) {
    return `${clean.slice(0, 3)}-${clean.slice(3, 7)}-${clean.slice(7)}`;
  } else if (clean.length === 10) {
    if (clean.startsWith("02")) {
      return `${clean.slice(0, 2)}-${clean.slice(2, 6)}-${clean.slice(6)}`;
    }
    return `${clean.slice(0, 3)}-${clean.slice(3, 6)}-${clean.slice(6)}`;
  } else if (clean.length === 9 && clean.startsWith("02")) {
    return `${clean.slice(0, 2)}-${clean.slice(2, 5)}-${clean.slice(5)}`;
  }
  return phone.toString().trim();
}

// Helper to parse date-time from combined Excel cells
function parseDateTimeCell(cellValue: any): { date: string; time: string } | null {
  if (!cellValue) return null;

  let dateObj: Date | null = null;

  if (cellValue instanceof Date) {
    dateObj = cellValue;
  } else if (typeof cellValue === "number" && cellValue > 30000 && cellValue < 60000) {
    const datePart = Math.floor(cellValue);
    const timePart = cellValue - datePart;
    const msInDay = 24 * 60 * 60 * 1000;
    const baseDateMs = new Date(Date.UTC(1900, 0, datePart - (datePart > 60 ? 2 : 1))).getTime();
    dateObj = new Date(baseDateMs + Math.round(timePart * msInDay));
  } else if (typeof cellValue === "object") {
    if (cellValue.result instanceof Date) {
      dateObj = cellValue.result;
    } else if (typeof cellValue.result === "number" && cellValue.result > 30000 && cellValue.result < 60000) {
      const datePart = Math.floor(cellValue.result);
      const timePart = cellValue.result - datePart;
      const msInDay = 24 * 60 * 60 * 1000;
      const baseDateMs = new Date(Date.UTC(1900, 0, datePart - (datePart > 60 ? 2 : 1))).getTime();
      dateObj = new Date(baseDateMs + Math.round(timePart * msInDay));
    } else if (cellValue.result) {
      return parseDateTimeString(cellValue.result.toString());
    }
  } else {
    return parseDateTimeString(cellValue.toString());
  }

  if (dateObj && !isNaN(dateObj.getTime())) {
    // Prevent timezone offset by formatting using local or UTC values carefully
    // Since Excel dates are timezone-naive, treating them UTC-aligned or keeping them as is prevents day shifting
    const y = dateObj.getUTCFullYear();
    const m = String(dateObj.getUTCMonth() + 1).padStart(2, "0");
    const d = String(dateObj.getUTCDate()).padStart(2, "0");
    const hh = String(dateObj.getUTCHours()).padStart(2, "0");
    const mm = String(dateObj.getUTCMinutes()).padStart(2, "0");
    return {
      date: `${y}-${m}-${d}`,
      time: `${hh}:${mm}`
    };
  }

  return null;
}

function parseDateTimeString(str: string): { date: string; time: string } | null {
  const cleaned = str.trim();
  const dateMatch = cleaned.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  const timeMatch = cleaned.match(/(\d{1,2}):(\d{2})/);

  let dateStr = "";
  let timeStr = "";

  if (dateMatch) {
    const y = dateMatch[1];
    const m = dateMatch[2].padStart(2, "0");
    const d = dateMatch[3].padStart(2, "0");
    dateStr = `${y}-${m}-${d}`;
  }

  if (timeMatch) {
    const hh = timeMatch[1].padStart(2, "0");
    const mm = timeMatch[2];
    timeStr = `${hh}:${mm}`;
  }

  if (dateStr || timeStr) {
    return {
      date: dateStr || "2026-07-03",
      time: timeStr || "09:00"
    };
  }

  return null;
}

// Helper to parse time ranges (e.g., "09:00~11:00", "18:00 - 20:00")
function parseTimeRange(timeStr: string): { startTime: string; endTime: string } {
  const defaultTime = { startTime: "09:00", endTime: "10:00" };
  if (!timeStr) return defaultTime;

  // Clean spacing and normalize
  const cleaned = timeStr.replace(/\s+/g, "");
  // Extract numbers
  const numbers = cleaned.match(/\d+/g);
  if (!numbers || numbers.length < 2) return defaultTime;

  let startH = 9, startM = 0, endH = 10, endM = 0;

  if (numbers.length >= 4) {
    startH = parseInt(numbers[0], 10);
    startM = parseInt(numbers[1], 10);
    endH = parseInt(numbers[2], 10);
    endM = parseInt(numbers[3], 10);
  } else if (numbers.length === 2) {
    startH = parseInt(numbers[0], 10);
    endH = parseInt(numbers[1], 10);
  } else if (numbers.length === 3) {
    startH = parseInt(numbers[0], 10);
    startM = parseInt(numbers[1], 10);
    endH = parseInt(numbers[2], 10);
  }

  // Constrain hours to 00-24, minutes to 00-59
  startH = Math.min(23, Math.max(0, startH));
  startM = Math.min(59, Math.max(0, startM));
  endH = Math.min(24, Math.max(0, endH));
  endM = Math.min(59, Math.max(0, endM));

  const pad = (n: number) => String(n).padStart(2, "0");
  return {
    startTime: `${pad(startH)}:${pad(startM)}`,
    endTime: `${pad(endH)}:${pad(endM)}`,
  };
}

// Rule-based mapping of facility names
function mapFacilityName(rawName: string): string {
  if (!rawName) return "";
  const name = rawName.toLowerCase().replace(/\s+/g, "");
  if (name.includes("테니스")) {
    if (name.includes("a") || name.includes("1") || name.includes("가")) return "테니스장A";
    if (name.includes("b") || name.includes("2") || name.includes("나")) return "테니스장B";
    if (name.includes("c") || name.includes("3") || name.includes("다")) return "테니스장C";
    if (name.includes("d") || name.includes("4") || name.includes("라")) return "테니스장D";
    return "테니스장A"; // fallback
  }
  if (name.includes('운동장') || name.includes("축구") || name.includes("풋살") || name.includes("잔디") || name.includes("구장")) {
    return "복합운동장";
  }
  if (name.includes("농구")) {
    if (name.includes("a") || name.includes("1") || name.includes("가")) return "농구장A";
    if (name.includes("b") || name.includes("2") || name.includes("나")) return "농구장B";
    if (name.includes("c") || name.includes("3") || name.includes("다")) return "농구장C";
    return "농구장A"; // fallback
  }
  return "";
}

// API: Parse uploaded Gongyoonuri Excel file
app.post("/api/convert", async (req, res) => {
  try {
    const { fileBase64 } = req.body;
    if (!fileBase64) {
      return res.status(400).json({ error: "파일 데이터(Base64)가 누락되었습니다." });
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(fileBase64, "base64"));
    
    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
      return res.status(400).json({ error: "엑셀 시트를 찾을 수 없습니다." });
    }

    let headerRowIdx = -1;
    let colMap: Record<string, number> = {};

    // Search for headers in first 25 rows
    for (let r = 1; r <= 25; r++) {
      const row = worksheet.getRow(r);
      let hasFacility = false;
      let hasBooker = false;
      let hasTime = false;

      row.eachCell((cell, colNumber) => {
        const val = getCellValueAsString(cell);
        if (val.includes("시설") || val.includes("자원") || val.includes("장소") || val.includes("시설명")) {
          hasFacility = true;
          colMap["facility"] = colNumber;
        }
        const isBookerCandidate = (
          val.includes("예약자") ||
          val.includes("신청자") ||
          val.includes("신청인") ||
          val.includes("예약인") ||
          val.includes("이용인") ||
          val.includes("성명") ||
          val.includes("이름") ||
          val.includes("이용자") ||
          val.includes("대표자") ||
          val.includes("사용자")
        );
        const hasContactKeyword = (
          val.includes("전화") ||
          val.includes("휴대폰") ||
          val.includes("연락처") ||
          val.includes("번호") ||
          val.includes("소속") ||
          val.includes("단체")
        );
        const hasExclusionKeyword = (
          val.includes("인원") ||
          val.includes("자수") ||
          val.includes("건수") ||
          val.includes("구분") ||
          val.includes("금액") ||
          val.includes("요금") ||
          val.includes("비용") ||
          val.includes("료") ||
          val.includes("아이디") ||
          val.includes("id") ||
          val.includes("ID") ||
          val === "명" ||
          val.includes("(명)")
        );
        if (isBookerCandidate && !hasContactKeyword && !hasExclusionKeyword) {
          hasBooker = true;
          colMap["booker"] = colNumber;
        }
        if (val.includes("시간") || val.includes("이용시간") || val.includes("예약시간")) {
          hasTime = true;
          colMap["time"] = colNumber;
        }
        if (val.includes("일자") || val.includes("이용일자") || val.includes("날짜") || val.includes("예약일자") || val.includes("이용일") || val.includes("신청일") || val.includes("예약일") || val.includes("신청일자")) {
          colMap["date"] = colNumber;
        }
        if (val.includes("시작일시") || val.includes("사용시작") || val.includes("시작시간")) {
          if (!val.includes("종료")) {
            colMap["startDateTime"] = colNumber;
          }
        }
        if (val.includes("종료일시") || val.includes("사용종료") || val.includes("종료시간") || val.includes("종료시작") || val.includes("종료")) {
          colMap["endDateTime"] = colNumber;
        }
        if (
          ((val.includes("번호") || val.includes("예약번호") || val.includes("신청번호")) && 
           !val.includes("전화") && !val.includes("휴대폰") && !val.includes("연락처") && !val.includes("예약자") && !val.includes("신청자"))
        ) {
          colMap["bookingNo"] = colNumber;
        }
        if (val.includes("단체") || val.includes("소속")) {
          colMap["group"] = colNumber;
        }
        if (
          val.includes("연락처") || 
          val.includes("전화") || 
          val.includes("휴대폰") || 
          val.includes("연락처명") || 
          val.includes("예약자번호") || 
          val.includes("신청자번호") || 
          val.includes("예약자 번호") || 
          val.includes("신청자 번호")
        ) {
          colMap["phone"] = colNumber;
        }
        if (val.includes("상태") || val.includes("승인")) {
          colMap["status"] = colNumber;
        }
      });

      if (hasFacility && (hasBooker || hasTime || colMap["startDateTime"])) {
        headerRowIdx = r;
        break;
      }
    }

    // Default column fallback indices if we couldn't detect a clear header row
    if (headerRowIdx === -1) {
      headerRowIdx = 1;
      colMap = {
        bookingNo: 1,
        facility: 2,
        booker: 3,
        group: 4,
        date: 5,
        time: 6,
        phone: 7,
        status: 8,
      };
    }

    const bookings: any[] = [];
    const rowsCount = worksheet.rowCount;

    for (let r = headerRowIdx + 1; r <= rowsCount; r++) {
      const row = worksheet.getRow(r);
      const facilityRaw = getCellValueAsString(row.getCell(colMap["facility"] || 2));
      let bookerRaw = getCellValueAsString(row.getCell(colMap["booker"] || 3));
      if (facilityRaw && !bookerRaw) {
        bookerRaw = "예약자";
      }

      let dateRaw = "";
      let startTime = "09:00";
      let endTime = "10:00";

      // Detect and parse combined date-time columns if present
      if (colMap["startDateTime"] || colMap["endDateTime"]) {
        const startVal = colMap["startDateTime"] ? row.getCell(colMap["startDateTime"]).value : null;
        const endVal = colMap["endDateTime"] ? row.getCell(colMap["endDateTime"]).value : null;

        const parsedStart = parseDateTimeCell(startVal);
        const parsedEnd = parseDateTimeCell(endVal);

        if (parsedStart) {
          dateRaw = parsedStart.date;
          startTime = parsedStart.time;
        }
        if (parsedEnd) {
          endTime = parsedEnd.time;
          if (!dateRaw) {
            dateRaw = parsedEnd.date;
          }
        }
      } else {
        const timeRaw = getCellValueAsString(row.getCell(colMap["time"] || 6));
        const times = parseTimeRange(timeRaw);
        startTime = times.startTime;
        endTime = times.endTime;

        const dateValue = row.getCell(colMap["date"] || 5).value;
        dateRaw = normalizeDate(dateValue);
      }

      if (!facilityRaw && !bookerRaw) continue; // Skip empty rows

      const resolvedFac = mapFacilityName(facilityRaw);

      // Night check (19:00 - 22:00)
      // Any overlap with [19:00, 22:00]
      const startHour = parseInt(startTime.split(":")[0], 10);
      const endHour = parseInt(endTime.split(":")[0], 10);
      const isNightHighlighted = (startHour < 22 && endHour > 19);

      bookings.push({
        id: `row-${r}-${Date.now()}`,
        bookingNo: getCellValueAsString(row.getCell(colMap["bookingNo"] || 1)),
        facilityName: facilityRaw,
        resolvedFacility: resolvedFac || "테니스장A", // Default fallback if not matched
        bookerName: bookerRaw,
        groupName: getCellValueAsString(row.getCell(colMap["group"] || 4)),
        date: dateRaw || "2026-07-03", // Friday fallback for preview
        startTime,
        endTime,
        phoneNumber: formatPhoneNumber(getCellValueAsString(row.getCell(colMap["phone"] || 7))),
        status: getCellValueAsString(row.getCell(colMap["status"] || 8)) || "예약완료",
        isNightHighlighted,
      });
    }

    const processedBookings = computeExceededBookings(bookings);
    return res.json({ bookings: processedBookings });
  } catch (error: any) {
    console.error("Error converting Excel file:", error);
    return res.status(500).json({ error: error.message || "파일 가공 중 에러가 발생했습니다." });
  }
});

// API: AI Facility Name Mapper (Exception Handling 4.2)
app.post("/api/gemini/resolve-facilities", async (req, res) => {
  try {
    const { rawNames } = req.body;
    if (!rawNames || !Array.isArray(rawNames)) {
      return res.status(400).json({ error: "rawNames 배열이 필요합니다." });
    }

    const ai = getGeminiClient();
    if (!ai) {
      // Fallback local rule-based mapping if Gemini key is missing
      const mapping = rawNames.map((name) => ({
        raw: name,
        resolved: mapFacilityName(name) || "테니스장A",
      }));
      return res.json({ mapping, isFallback: true });
    }

    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: `당신은 공유누리 예약 관리 시스템의 핵심 매핑 엔진입니다.
제시된 원본 시설명들을 아래의 8가지 고정 표준 시설명 중 가장 적절한 하나로 매핑해 주세요.
반드시 대소문자와 형식을 맞춰서 매핑해야 합니다.

[표준 시설명 8가지]
- 테니스장A
- 테니스장B
- 테니스장C
- 테니스장D
- 복합운동장
- 농구장A
- 농구장B
- 농구장C

[원본 목록]
${JSON.stringify(rawNames)}

반드시 정확히 매핑 결과를 JSON 배열로만 응답해 주세요. 예외적인 텍스트나 사족을 절대 달지 마세요.`,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              raw: { type: Type.STRING },
              resolved: { type: Type.STRING },
            },
            required: ["raw", "resolved"],
          },
        },
      },
    });

    const mapping = JSON.parse(response.text?.trim() || "[]");
    return res.json({ mapping, isFallback: false });
  } catch (error: any) {
    console.error("Gemini mapping error:", error);
    // Fallback on error
    const fallbackMap = (req.body.rawNames || []).map((name: string) => ({
      raw: name,
      resolved: mapFacilityName(name) || "테니스장A",
    }));
    return res.json({ mapping: fallbackMap, error: error.message, isFallback: true });
  }
});

// API: Download a beautifully styled Excel Timetable (Requirement 2.4 & 2.5)
app.post("/api/export", async (req, res) => {
  try {
    const { bookings, weekLabel, weekDates } = req.body;
    if (!bookings || !Array.isArray(bookings)) {
      return res.status(400).json({ error: "예약 목록(bookings)이 유효하지 않습니다." });
    }

    // Helper to calculate or parse the Friday, Saturday, Sunday dates
    const getExportWeekDates = (bList: any[], provided: any) => {
      if (provided && provided.Friday && provided.Saturday && provided.Sunday) {
        return provided;
      }
      let referenceDateStr = "";
      if (bList && bList.length > 0) {
        const found = bList.find(x => x.date && x.date.trim() !== "");
        if (found) referenceDateStr = found.date;
      }
      if (!referenceDateStr) {
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, "0");
        const day = String(now.getDate()).padStart(2, "0");
        referenceDateStr = `${year}-${month}-${day}`;
      }
      const d = parseLocalDate(referenceDateStr);
      const day = d.getDay();
      let offset = 0;
      if (day === 0) offset = -2;
      else if (day === 6) offset = -1;
      else offset = 5 - day;

      const fri = new Date(d);
      fri.setDate(d.getDate() + offset);

      const sat = new Date(fri);
      sat.setDate(fri.getDate() + 1);

      const sun = new Date(fri);
      sun.setDate(fri.getDate() + 2);

      const format = (dateObj: Date) => {
        const y = dateObj.getFullYear();
        const m = String(dateObj.getMonth() + 1).padStart(2, "0");
        const dStr = String(dateObj.getDate()).padStart(2, "0");
        return `${y}-${m}-${dStr}`;
      };

      return {
        Friday: format(fri),
        Saturday: format(sat),
        Sunday: format(sun),
      };
    };

    const formatMMDD = (dateStr: string): string => {
      if (!dateStr) return "";
      const parts = dateStr.split("-");
      if (parts.length === 3) {
        return `${parts[1]}.${parts[2]}.`;
      }
      return "";
    };

    const wDates = getExportWeekDates(bookings, weekDates);

    const processedBookings = computeExceededBookings(bookings);
    const workbook = new ExcelJS.Workbook();
    
    // Create ONE unified worksheet instead of separate sheets
    const sheet = workbook.addWorksheet("공유누리 시간표");
    sheet.views = [{ showGridLines: true }];

    // Set standard column keys and widths
    sheet.columns = [
      { key: "time", width: 15 },
      { key: "tennis_a", width: 18 },
      { key: "tennis_b", width: 18 },
      { key: "tennis_c", width: 18 },
      { key: "tennis_d", width: 18 },
      { key: "stadium", width: 22 },
      { key: "basket_a", width: 18 },
      { key: "basket_b", width: 18 },
      { key: "basket_c", width: 18 },
    ];

    let currentRowNum = 1;

    // 1. Master Title Block (Using formal, soft pastel steel blue style for public offices)
    sheet.getRow(currentRowNum).values = [`공유누리 시간표(${formatMMDD(wDates.Friday)}~${formatMMDD(wDates.Sunday)})`];
    sheet.mergeCells(currentRowNum, 1, currentRowNum, 9);
    const titleCell = sheet.getCell(currentRowNum, 1);
    titleCell.font = { name: "Malgun Gothic", size: 16, bold: true, color: { argb: "FF1E293B" } };
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { theme: 0, tint: -0.15 } as any, // White, Background 1, 15% Darker (Theme-based for easy editing)
    } as ExcelJS.Fill;
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(currentRowNum).height = 40;
    currentRowNum++;

    // 2. Subtitle Block (Deleted as requested)
    currentRowNum++; // Spacing row

    // Group bookings by day
    const days = [
      { name: "금요일", key: "Friday", date: wDates.Friday },
      { name: "토요일", key: "Saturday", date: wDates.Saturday },
      { name: "일요일", key: "Sunday", date: wDates.Sunday },
    ];

    for (const day of days) {
      // 3. Day Section Header Row
      const dayHeaderRow = sheet.getRow(currentRowNum);
      dayHeaderRow.values = [`■ ${day.name} (${formatMMDD(day.date)})`];
      sheet.mergeCells(currentRowNum, 1, currentRowNum, 9);
      const dCell = sheet.getCell(currentRowNum, 1);
      dCell.font = { name: "Malgun Gothic", size: 11, bold: true, color: { argb: "FF0F172A" } };
      dCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { theme: 0, tint: -0.05 } as any, // White, Background 1, 5% Darker
      } as ExcelJS.Fill;
      dCell.alignment = { horizontal: "left", vertical: "middle" };
      dayHeaderRow.height = 28;
      currentRowNum++;

      // 4. Column Headers (for this day's grid)
      const gridHeaderRow = sheet.getRow(currentRowNum);
      gridHeaderRow.values = ["시간", "테니스장A", "테니스장B", "테니스장C", "테니스장D", "복합운동장", "농구장A", "농구장B", "농구장C"];
      gridHeaderRow.height = 26;
      for (let col = 1; col <= 9; col++) {
        const cell = gridHeaderRow.getCell(col);
        cell.font = { name: "Malgun Gothic", size: 10, bold: true, color: { argb: "FF334155" } };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { theme: 0, tint: -0.15 } as any, // White, Background 1, 15% Darker
        } as ExcelJS.Fill;
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin" },
          bottom: { style: "medium" },
          left: { style: "thin" },
          right: { style: "thin" },
        };
      }
      currentRowNum++;

      // 5. Build Hours Grid
      const hours: number[] = [];
      if (day.name === "금요일") {
        hours.push(6, 7, 8, 18, 19, 20, 21);
      } else {
        for (let h = 7; h <= 19; h++) {
          hours.push(h);
        }
      }

      const hourToRowIdx: Record<number, number> = {};
      hours.forEach((hour) => {
        const rowNum = currentRowNum;
        const startStr = String(hour).padStart(2, "0") + ":00";
        const endStr = String(hour + 1).padStart(2, "0") + ":00";
        const row = sheet.getRow(rowNum);
        row.getCell(1).value = `${startStr}~${endStr}`;
        row.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
        row.getCell(1).font = { name: "Malgun Gothic", size: 9, bold: true, color: { argb: "FF475569" } };
        row.getCell(1).border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" },
        };
        row.height = 24;

        // Apply borders and backgrounds to other columns
        for (let c = 2; c <= 9; c++) {
          const cell = row.getCell(c);
          cell.border = {
            top: { style: "thin" },
            bottom: { style: "thin" },
            left: { style: "thin" },
            right: { style: "thin" },
          };
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.font = { name: "Malgun Gothic", size: 10 };
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { theme: 0 }, // White, Background 1
          } as ExcelJS.Fill;
        }

        hourToRowIdx[hour] = rowNum;
        currentRowNum++;
      });

      // 6. Place Bookings on Grid
      const dayBookings = processedBookings.filter((b) => {
        if (!b.date) return false;
        if (b.date === day.date) return true;
        const bDate = parseLocalDate(b.date);
        const targetDayOfWeek = day.name === "금요일" ? 5 : day.name === "토요일" ? 6 : 0;
        return bDate.getDay() === targetDayOfWeek;
      });

      const facToColIdx: Record<string, number> = {
        "테니스장A": 2,
        "테니스장B": 3,
        "테니스장C": 4,
        "테니스장D": 5,
        "복합운동장": 6,
        "농구장A": 7,
        "농구장B": 8,
        "농구장C": 9,
      };

      const gridOccupied: Record<string, boolean> = {}; // key: "col-row"

      dayBookings.forEach((b) => {
        const colIdx = facToColIdx[b.resolvedFacility];
        if (!colIdx) return;

        const startH = parseInt(b.startTime.split(":")[0], 10);
        const endH = parseInt(b.endTime.split(":")[0], 10);

        const bookingHours = hours.filter((h) => h >= startH && h < endH);
        if (bookingHours.length === 0) return;

        const rowIndices = bookingHours.map((h) => hourToRowIdx[h]).filter(Boolean);
        if (rowIndices.length === 0) return;

        const firstRow = Math.min(...rowIndices);
        const lastRow = Math.max(...rowIndices);

        let isOccupied = false;
        for (let rIdx = firstRow; rIdx <= lastRow; rIdx++) {
          if (gridOccupied[`${colIdx}-${rIdx}`]) {
            isOccupied = true;
            break;
          }
        }

        const isNightHighlight = b.isNightHighlighted || (startH < 22 && endH > 19);
        const isExceeded = b.isExceeded;

        // Elegant pastel office coloring
        const cellFont = {
          name: "Malgun Gothic",
          size: 10,
          bold: true,
          color: isExceeded
            ? { argb: "FF991B1B" } // Soft dark red
            : isNightHighlight
            ? { argb: "FF92400E" } // Soft dark amber
            : { argb: "FF1E293B" },
        };

        const cellFill: ExcelJS.Fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { theme: 0, tint: -0.05 } as any, // White, Background 1, 5% Darker
        };

        const nameColor = isExceeded
          ? { argb: "FF991B1B" }
          : isNightHighlight
          ? { argb: "FF92400E" }
          : { argb: "FF1E293B" };

        const phoneColor = isExceeded
          ? { argb: "FFB91C1C" }
          : isNightHighlight
          ? { argb: "FFB45309" }
          : { argb: "FF64748B" };

        const nameFont = { name: "Malgun Gothic", size: 11, bold: true, color: nameColor };
        const phoneFont = { name: "Malgun Gothic", size: 8, color: phoneColor };

        if (isOccupied) {
          const pivotCell = sheet.getRow(firstRow).getCell(colIdx);
          let currentRichText: any[] = [];
          if (pivotCell.value && typeof pivotCell.value === "object" && "richText" in pivotCell.value) {
            currentRichText = [...(pivotCell.value as any).richText];
          } else if (pivotCell.value) {
            currentRichText = [{ text: pivotCell.value.toString() }];
          }

          const hasBooker = currentRichText.some(item => item.text && item.text.includes(b.bookerName));
          if (!hasBooker) {
            if (currentRichText.length > 0) {
              currentRichText.push({ text: "\n/ ", font: { name: "Malgun Gothic", size: 9, bold: true, color: { argb: "FF475569" } } });
            }
            currentRichText.push({ text: b.bookerName, font: nameFont });
            if (b.phoneNumber) {
              currentRichText.push({ text: "\n(" + formatPhoneNumber(b.phoneNumber) + ")", font: phoneFont });
            }
            pivotCell.value = { richText: currentRichText };
          }
          
          // pivotCell.font = cellFont; // Removed to prevent ExcelJS from overwriting/flattening individual richText run fonts
          pivotCell.fill = cellFill;
          return;
        }

        for (let rIdx = firstRow; rIdx <= lastRow; rIdx++) {
          gridOccupied[`${colIdx}-${rIdx}`] = true;
        }

        if (firstRow !== lastRow) {
          sheet.mergeCells(firstRow, colIdx, lastRow, colIdx);
        }

        const bookingCell = sheet.getRow(firstRow).getCell(colIdx);
        const cellRuns: any[] = [{ text: b.bookerName, font: nameFont }];
        if (b.phoneNumber) {
          cellRuns.push({ text: "\n(" + formatPhoneNumber(b.phoneNumber) + ")", font: phoneFont });
        }
        bookingCell.value = { richText: cellRuns };
        bookingCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        // bookingCell.font = cellFont; // Removed to prevent ExcelJS from overwriting/flattening individual richText run fonts
        bookingCell.fill = cellFill;

        for (let rIdx = firstRow; rIdx <= lastRow; rIdx++) {
          const cell = sheet.getRow(rIdx).getCell(colIdx);
          cell.border = {
            top: { style: "thin" },
            bottom: { style: "thin" },
            left: { style: "thin" },
            right: { style: "thin" },
          };
        }
      });

      // 7. Secondary Table at the bottom of the day's grid (Exceeded bookings)
      const dayExceededBookings = dayBookings.filter(b => b.isExceeded);
      currentRowNum++;

      const excTitleRow = sheet.getRow(currentRowNum);
      excTitleRow.values = [`⚠️ 1일 4시간 초과 예약 명단 (${day.name} - 총 ${dayExceededBookings.length}건)`];
      sheet.mergeCells(currentRowNum, 1, currentRowNum, 8);
      const excTitleCell = sheet.getCell(currentRowNum, 1);
      excTitleCell.font = { name: "Malgun Gothic", size: 10, bold: true, color: { argb: "FF991B1B" } };
      excTitleRow.height = 24;
      currentRowNum++;

      if (dayExceededBookings.length > 0) {
        const excHeaderRow = sheet.getRow(currentRowNum);
        excHeaderRow.values = [
          "No.",
          "예약번호",
          "예약자명",
          "단체명",
          "이용시간",
          "표준시설",
          "연락처",
          "비고 (초과여부)"
        ];
        excHeaderRow.height = 22;
        for (let col = 1; col <= 8; col++) {
          const cell = excHeaderRow.getCell(col);
          cell.font = { name: "Malgun Gothic", size: 9, bold: true, color: { argb: "FF7F1D1D" } };
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { theme: 0, tint: -0.15 } as any // White, Background 1, 15% Darker
          };
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.border = {
            top: { style: "thin" },
            bottom: { style: "thin" },
            left: { style: "thin" },
            right: { style: "thin" }
          };
        }
        currentRowNum++;

        dayExceededBookings.forEach((b, idx) => {
          const r = sheet.getRow(currentRowNum);
          r.values = [
            idx + 1,
            b.bookingNo || "-",
            b.bookerName,
            b.groupName || "-",
            `${b.startTime}~${b.endTime}`,
            b.resolvedFacility,
            b.phoneNumber || "-",
            "초과예약분"
          ];
          r.height = 22;

          for (let col = 1; col <= 8; col++) {
            const cell = r.getCell(col);
            cell.font = { name: "Malgun Gothic", size: 9, color: { argb: "FF334155" } };
            cell.alignment = { horizontal: "center", vertical: "middle" };
            cell.border = {
              top: { style: "thin" },
              bottom: { style: "thin" },
              left: { style: "thin" },
              right: { style: "thin" }
            };

            if (col === 8) {
              cell.font = { name: "Malgun Gothic", size: 9, bold: true, color: { argb: "FF991B1B" } };
              cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { theme: 0, tint: -0.05 } as any // White, Background 1, 5% Darker
              };
            }
          }
          currentRowNum++;
        });
      } else {
        const infoRow = sheet.getRow(currentRowNum);
        infoRow.values = ["※ 해당 일자에 1일 4시간 초과 예약 내역이 존재하지 않습니다."];
        sheet.mergeCells(currentRowNum, 1, currentRowNum, 8);
        const infoCell = sheet.getCell(currentRowNum, 1);
        infoCell.font = { name: "Malgun Gothic", size: 9, italic: true, color: { argb: "FF64748B" } };
        infoCell.alignment = { horizontal: "left", vertical: "middle" };
        infoRow.height = 22;
        infoCell.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" }
        };
        currentRowNum++;
      }

      // Separator gap
      currentRowNum += 2;
    }

    // Generate Base64
    const buffer = await workbook.xlsx.writeBuffer();
    const base64 = Buffer.from(buffer as ArrayBuffer).toString("base64");

    return res.json({ excelBase64: base64 });
  } catch (error: any) {
    console.error("Error exporting Excel:", error);
    return res.status(500).json({ error: error.message || "엑셀 생성 중 에러가 발생했습니다." });
  }
});

// API: Send beautiful styled Excel Timetable to Situation Room via Email
app.post("/api/send-email", async (req, res) => {
  try {
    const { excelBase64, weekLabel, recipientEmail, subject, emailBody, fileName } = req.body;
    if (!excelBase64) {
      return res.status(400).json({ error: "첨부할 엑셀 데이터(excelBase64)가 필요합니다." });
    }

    const targetRecipient = recipientEmail || "secom1112@naver.com";
    const targetSubject = subject || `[공유누리] 주말 예약자 시설 시간표 전달`;
    const targetBody = emailBody || "안녕하세요, 상황실 담당자님.\n\n주말 공유누리 시설 예약자 시간표를 첨부하여 전달해 드립니다.\n확인 부탁드립니다.\n\n감사합니다.";

    const attachmentBuffer = Buffer.from(excelBase64, "base64");

    // Get SMTP Configuration
    const host = process.env.SMTP_HOST || "";
    const port = parseInt(process.env.SMTP_PORT || "587", 10);
    const user = process.env.SMTP_USER || "";
    const pass = process.env.SMTP_PASS || "";

    const finalFileName = fileName || `${(weekLabel || "timetable").replace(/[\s~-]/g, "")}_공유누리_예약자시간표.xlsx`;

    if (!host || !user || !pass) {
      // SMTP configuration is missing, return success: false with preview data
      return res.json({
        success: false,
        configMissing: true,
        message: "서버의 SMTP 설정(SMTP_HOST, SMTP_USER, SMTP_PASS)이 완료되지 않았습니다.",
        preview: {
          recipient: targetRecipient,
          subject: targetSubject,
          body: targetBody,
          fileName: finalFileName
        }
      });
    }

    const nodemailer = await import("nodemailer");
    const transporter = nodemailer.createTransport({
      host,
      port,
      secure: port === 465,
      auth: {
        user,
        pass,
      },
    });

    const mailOptions = {
      from: `"공유누리 자동화 시스템" <${user}>`,
      to: targetRecipient,
      subject: targetSubject,
      text: targetBody,
      attachments: [
        {
          filename: finalFileName,
          content: attachmentBuffer,
        },
      ],
    };

    await transporter.sendMail(mailOptions);

    return res.json({
      success: true,
      message: `${targetRecipient} 상황실로 시간표 메일이 성공적으로 전송되었습니다.`
    });
  } catch (error: any) {
    console.error("Error sending email:", error);
    return res.status(500).json({ error: error.message || "메일 전송 중 에러가 발생했습니다." });
  }
});

// API: Generate Mock Raw Gongyoonuri Excel for demo testing
app.get("/api/mock-template", async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("예약조회결과");

    // Standard column headers in raw Gongyoonuri export
    sheet.columns = [
      { header: "예약번호", key: "bookingNo", width: 18 },
      { header: "자원명", key: "facilityName", width: 25 },
      { header: "신청자", key: "booker", width: 12 },
      { header: "단체명/소속", key: "group", width: 20 },
      { header: "이용일자", key: "date", width: 14 },
      { header: "이용시간", key: "time", width: 18 },
      { header: "연락처", key: "phone", width: 15 },
      { header: "예약상태", key: "status", width: 12 },
    ];

    // Helper to get dates for current week's Fri, Sat, Sun
    const getWeekDateStr = (dayOffset: number) => {
      // 5: Fri, 6: Sat, 7: Sun (or 0)
      const d = new Date();
      const currentDay = d.getDay(); // 0 is Sun, 1 is Mon...
      const distanceToFriday = 5 - currentDay;
      const targetDate = new Date();
      targetDate.setDate(d.getDate() + distanceToFriday + dayOffset);
      const y = targetDate.getFullYear();
      const m = String(targetDate.getMonth() + 1).padStart(2, "0");
      const day = String(targetDate.getDate()).padStart(2, "0");
      return `${y}-${m}-${day}`;
    };

    const friDate = getWeekDateStr(0);
    const satDate = getWeekDateStr(1);
    const sunDate = getWeekDateStr(2);

    // Mock records
    const mockData = [
      // 금요일 (Friday)
      {
        bookingNo: "GY-20260703-01",
        facilityName: "테니스장 A코트",
        booker: "김철수",
        group: "일요테니스회",
        date: friDate,
        time: "06:00~08:00",
        phone: "010-1111-2222",
        status: "예약완료",
      },
      {
        bookingNo: "GY-20260703-02",
        facilityName: "대운동장 (천연잔디)",
        booker: "이영희",
        group: "FC 한마음",
        date: friDate,
        time: "18:00~21:00", // overlaps 19:00~21:00 night highlight!
        phone: "010-3333-4444",
        status: "예약완료",
      },
      {
        bookingNo: "GY-20260703-03",
        facilityName: "테니스장 B코트",
        booker: "민수인",
        group: "",
        date: friDate,
        time: "20:00~22:00", // night highlight!
        phone: "010-1212-3434",
        status: "예약완료",
      },
      // 토요일 (Saturday)
      {
        bookingNo: "GY-20260704-01",
        facilityName: "농구장 1코트",
        booker: "박민수",
        group: "바스켓러버",
        date: satDate,
        time: "10:00~12:00",
        phone: "010-5555-6666",
        status: "예약완료",
      },
      {
        bookingNo: "GY-20260704-02",
        facilityName: "테니스장 C",
        booker: "최재웅",
        group: "강서테니스클럽",
        date: satDate,
        time: "18:00~20:00", // overlaps night 19:00
        phone: "010-7777-8888",
        status: "예약완료",
      },
      {
        bookingNo: "GY-20260704-03",
        facilityName: "테니스장 D",
        booker: "한유미",
        group: "",
        date: satDate,
        time: "14:00~16:00",
        phone: "010-9999-0000",
        status: "예약완료",
      },
      {
        bookingNo: "GY-20260704-04",
        facilityName: "테니스장 D",
        booker: "정지민",
        group: "주말복식회",
        date: satDate,
        time: "15:00~17:00", // Overlaps with 한유미 at 15:00-16:00! Trigger duplicate handling.
        phone: "010-4444-5555",
        status: "예약완료",
      },
      // 일요일 (Sunday)
      {
        bookingNo: "GY-20260705-01",
        facilityName: "축구 보조경기장 (인조잔디)",
        booker: "정재민",
        group: "청년축구회",
        date: sunDate,
        time: "08:00~11:00",
        phone: "010-4321-8765",
        status: "예약완료",
      },
      {
        bookingNo: "GY-20260705-02",
        facilityName: "농구장 C코트",
        booker: "윤민석",
        group: "덩크슛동호회",
        date: sunDate,
        time: "18:00~20:00", // night highlight!
        phone: "010-8765-4321",
        status: "승인대기",
      },
    ];

    mockData.forEach((row) => sheet.addRow(row));

    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Disposition", "attachment; filename=gongyoonuri_raw_export.xlsx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    return res.send(buffer);
  } catch (error: any) {
    console.error("Error generating mock template:", error);
    return res.status(500).send("에러 발생: " + error.message);
  }
});

// Export app for serverless platforms like Vercel
export default app;

// Serve Vite client files or setup dev server middleware
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

if (process.env.VERCEL !== "1") {
  startServer();
}
