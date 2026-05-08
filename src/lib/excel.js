import * as XLSX from "xlsx";

/* =========================
   Helpers
========================= */

// Normalize text
const normalize = (str) => str?.toString().toLowerCase().trim() || "";

/**
 * Detect Sheet Type (Smart)
 */
function detectSheetType(sheetName, sampleRows) {
  const name = normalize(sheetName);

  if (name.includes("power") || name.includes("rectifier")) return "power";
  if (
    name.includes("down") ||
    name.includes("critical") ||
    name.includes("link") ||
    name.includes("om")
  )
    return "down";
  if (name.includes("ht") || name.includes("temp")) return "high_temp";
  if (name.includes("gen") || name.includes("fuel")) return "generator";
  if (name.includes("door")) return "door";

  const sampleText = sampleRows
    .slice(0, 20)
    .map((row) => Object.values(row).join(" ").toLowerCase())
    .join(" ");

  if (sampleText.includes("mains") || sampleText.includes("power"))
    return "power";
  if (
    sampleText.includes("down") ||
    sampleText.includes("bts") ||
    sampleText.includes("link")
  )
    return "down";
  if (sampleText.includes("temp")) return "high_temp";
  if (sampleText.includes("generator") || sampleText.includes("gen set"))
    return "generator";
  if (sampleText.includes("door")) return "door";
  return "unknown";
}

/**
 * Detect Columns dynamically
 */
function detectColumns(headerRow, isAutin = false) {
  const mapping = {};
  const entries = Object.entries(headerRow);

  if (isAutin) {
    mapping.siteCode = "B";
  } else {
    // 1. Find Site Code (Priority: 'site code' > 'site id' > 'code' > 'site')
    for (const [key, value] of entries) {
      const text = normalize(value);
      if (text.includes("sitecode") || text.includes("siteid")) {
        mapping.siteCode = key;
        break;
      }
    }
    if (!mapping.siteCode) {
      for (const [key, value] of entries) {
        const text = normalize(value);
        if (
          text.includes("code") ||
          (text.includes("site") && !text.includes("name"))
        ) {
          mapping.siteCode = key;
          break;
        }
      }
    }
    // Fallback to anything with 'site' if still not found
    if (!mapping.siteCode) {
      for (const [key, value] of entries) {
        if (normalize(value).includes("site")) {
          mapping.siteCode = key;
          break;
        }
      }
    }

    // Hard fallback to Column F (requested by user) if still nothing found
    if (!mapping.siteCode && headerRow["F"]) {
      mapping.siteCode = "F";
    }
  }

  for (const [key, value] of entries) {
    const text = normalize(value);

    if (
      !mapping.alarmName &&
      (text.includes("alarm") || text.includes("event"))
    )
      mapping.alarmName = key;

    if (
      !mapping.duration &&
      (text.includes("duration") || text.includes("elapsed"))
    )
      mapping.duration = key;

    if (text.includes("date")) {
      if (text.includes("start") || text.includes("occur"))
        mapping.startDate = key;
      if (text.includes("end") || text.includes("clear")) mapping.endDate = key;
    }
    if (text.includes("time")) {
      if (text.includes("start") || text.includes("occur"))
        mapping.startTime = key;
      if (text.includes("end") || text.includes("clear")) mapping.endTime = key;
    }

    if (
      !mapping.startTime &&
      (text.includes("start") || text.includes("occur")) &&
      !text.includes("date")
    )
      mapping.startTime = key;

    if (
      !mapping.endTime &&
      (text.includes("end") || text.includes("clear")) &&
      !text.includes("date")
    )
      mapping.endTime = key;

    if (
      !mapping.siteName &&
      text.includes("site") &&
      text.includes("name")
    )
      mapping.siteName = key;
  }

  return mapping;
}

/**
 * Parse Duration → ALWAYS minutes
 */
function parseDuration(durationStr, startTime, endTime) {
  // PREFERRED: Absolute Time Mathematics
  if (startTime && endTime) {
    const start = new Date(startTime);
    const end = new Date(endTime);
    if (
      !isNaN(start) &&
      !isNaN(end) &&
      start.getTime() > 0 &&
      end.getTime() > 0
    ) {
      return Math.abs(end - start) / (1000 * 60);
    }
  }

  if (durationStr) {
    const str = durationStr.toString().trim();

    // "D days HH:MM:SS"
    const daysMatch = str.match(/(\d+)\s*(?:days?|d)\s+(\d+):(\d+):?(\d+)?/i);
    if (daysMatch) {
      const days = parseInt(daysMatch[1]) || 0;
      const hours = parseInt(daysMatch[2]) || 0;
      const minutes = parseInt(daysMatch[3]) || 0;
      return days * 24 * 60 + hours * 60 + minutes;
    }

    // "HH:MM:SS"
    const timeMatch = str.match(/(\d+):(\d+):?(\d+)?/);
    if (timeMatch) {
      const hours = parseInt(timeMatch[1]) || 0;
      const minutes = parseInt(timeMatch[2]) || 0;
      return hours * 60 + minutes;
    }

    const num = parseFloat(str);
    if (!isNaN(num)) {
      // If Excel exported a raw time fraction (e.g., 0.045 represents 1 hour and 5 mins)
      if (num > 0 && num < 10) {
        // It's likely a fractional day representation from Excel
        return num * 24 * 60;
      }
      return num; // literal minutes
    }
  }

  if (startTime && endTime) {
    const start = new Date(startTime);
    const end = new Date(endTime);
    if (!isNaN(start) && !isNaN(end)) {
      return Math.abs(end - start) / (1000 * 60);
    }
  }

  return 0;
}

/**
 * Excel Date Serial to JS Date
 */
export function excelToDate(serial) {
  if (!serial) return null;
  if (serial instanceof Date) return serial;

  const str = serial.toString().trim();
  const num = parseFloat(str);

  // 1. Handle Excel Serial Numbers (most common)
  if (!isNaN(num) && /^\d+(\.\d+)?$/.test(str)) {
    const date = new Date(Math.round((num - 25569) * 86400 * 1000));
    return isNaN(date.getTime()) ? null : date;
  }

  // 2. Handle common DD/MM/YYYY or DD-MM-YYYY formats that JS Date() misses
  const dmyMatch = str.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1]);
    const month = parseInt(dmyMatch[2]) - 1;
    let year = parseInt(dmyMatch[3]);
    if (year < 100) year += 2000;
    const hour = parseInt(dmyMatch[4] || 0);
    const minute = parseInt(dmyMatch[5] || 0);
    const second = parseInt(dmyMatch[6] || 0);
    const d = new Date(year, month, day, hour, minute, second);
    return isNaN(d.getTime()) ? null : d;
  }

  // 3. Handle H:MM:SS or HH:MM:SS time strings
  const timeMatch = str.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
  if (timeMatch) {
    const d = new Date();
    d.setHours(parseInt(timeMatch[1]), parseInt(timeMatch[2]), parseInt(timeMatch[3]), 0);
    return d;
  }

  // 4. Default JS parsing
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Deterministic Classification (ONE category)
 */
function classifyAlarm(alarmText, sheetType) {
  const text = normalize(alarmText);

  // 1. If we have a trusted sheet type, USE IT
  if (sheetType === "power") return "power";
  if (sheetType === "down") return "down";
  if (sheetType === "high_temp") return "high_temp";
  if (sheetType === "generator") return "generator";
  if (sheetType === "door") return "door";

  // 2. Fallback to keyword heuristics for unknown sheets
  if (
    text.includes("down") ||
    text.includes("link") ||
    text.includes("bts") ||
    text.includes("o&m")
  )
    return "down";
  if (
    text.includes("mains") ||
    text.includes("rectifier") ||
    text.includes("dc low") ||
    text.includes("power")
  )
    return "power";
  if (text.includes("temp") || text.includes("heat")) return "high_temp";
  if (text.includes("gen") || text.includes("fuel") || text.includes("dg"))
    return "generator";
  if (text.includes("door")) return "door";

  return "other";
}

/**
 * Format Duration
 */
export function formatDuration(minutes) {
  if (!minutes || minutes <= 0) return "0m";

  const days = Math.floor(minutes / (24 * 60));
  const hours = Math.floor((minutes % (24 * 60)) / 60);
  const mins = Math.round(minutes % 60);

  if (days > 0) return `${days}d ${hours}h`;
  if (hours > 0) return `${hours}h ${mins}m`;
  return `${mins}m`;
}

/* =========================
   MAIN PARSER
========================= */

export const parseAlarmsExcel = (file, forcedType = null) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {
          type: "array",
          cellDates: true,
        });

        const allData = [];

        workbook.SheetNames.forEach((sheetName, idx) => {
          let sheetType = forcedType;
          if (!sheetType) {
            // First try strict index mapping
            const idxType = {
              0: "power",
              1: "down",
              3: "high_temp",
              4: "generator",
            }[idx];

            // Then try smart detection from name
            const worksheet = workbook.Sheets[sheetName];
            const sampleRows = XLSX.utils
              .sheet_to_json(worksheet, { header: "A", range: 0, defval: "" })
              .slice(0, 10);
            const nameType = detectSheetType(sheetName, sampleRows);

            // Result: prioritize index for Energy Master, but fallback to name detection
            sheetType = idxType || (nameType !== "unknown" ? nameType : null);

            if (!sheetType) return;
          }

          const worksheet = workbook.Sheets[sheetName];
          const rawRows = XLSX.utils.sheet_to_json(worksheet, {
            header: "A",
            defval: "",
          });

          if (rawRows.length < 2) return;
          console.log(
            `[Parser] Extracting Sheet ${idx + 1} ("${sheetName}") as ${sheetType}`,
          );

          // Find the header row (scan first 10 rows for site/alarm keywords)
          let headerRowIndex = 0;
          let columns = {};
          const isAutin = !!forcedType;

          for (let i = 0; i < Math.min(10, rawRows.length); i++) {
            const detected = detectColumns(rawRows[i], isAutin);
            if (detected.siteCode || detected.alarmName) {
              headerRowIndex = i;
              columns = detected;
              break;
            }
          }

          const formattedData = rawRows
            .slice(headerRowIndex + 1)
            .map((row) => {
              let siteCode = row[columns.siteCode]?.toString().trim() || null;

              if (!siteCode) return null;

              // Robust normalization for matching: remove all non-alphanumeric characters
              siteCode = siteCode
                .toString()
                .replace(/[^a-zA-Z0-9]/g, "")
                .toUpperCase();

              const alarmName = row[columns.alarmName]?.toString() || "Unknown";
              const siteNameFromSheet = row[columns.siteName]?.toString() || "";

              // Smart combining of Date + Time columns
              const combineDT = (dateKey, timeKey) => {
                const dVal = row[dateKey];
                const tVal = row[timeKey];

                const d = excelToDate(dVal);
                const t = excelToDate(tVal);

                if (!d) return t;
                if (!t) return d;

                // Create a clean date and merge time
                const combined = new Date(d);
                combined.setHours(
                  t.getHours(),
                  t.getMinutes(),
                  t.getSeconds(),
                  0,
                );
                return combined;
              };

              const startTime = combineDT(columns.startDate, columns.startTime);
              const endTime = combineDT(columns.endDate, columns.endTime);

              const durationRaw = row[columns.duration];
              const durationMinutes = parseDuration(
                durationRaw,
                startTime,
                endTime,
              );

              const category = classifyAlarm(alarmName, sheetType);

              return {
                sheet: sheetName,
                detectedType: sheetType,
                siteCode,
                siteName: siteNameFromSheet,
                alarmName,
                durationRaw: durationRaw?.toString() || "",
                durationMinutes,
                durationFormatted: formatDuration(durationMinutes),
                category, // used in some places
                categories: [category], // ✅ Required by analyzeAlarms
                startTime,
                endTime,
                raw: row,
                startTimeKey: columns.startTime,
                endTimeKey: columns.endTime,
              };
            })
            .filter(Boolean);

          for (const item of formattedData) {
            allData.push(item);
          }
        });

        console.log("==========");
        console.log("TOTAL RECORDS:", allData.length);

        // Category stats
        const stats = {};
        allData.forEach((a) => {
          stats[a.category] = (stats[a.category] || 0) + 1;
        });

        console.log("CATEGORY DISTRIBUTION:", stats);

        resolve(allData);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

/* =========================
   SITE LIST PARSER
========================= */

export const parseSiteListExcel = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {
          type: "array",
          cellDates: true,
        });

        const worksheet = workbook.Sheets["SiteList"];
        if (!worksheet) {
          reject(new Error("Sheet 'SiteList' not found"));
          return;
        }

        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
          header: "A",
        });

        const formattedData = jsonData
          .slice(1)
          .map((row) => ({
            shortCode: row["A"]?.toString().trim(),
            siteName: row["B"],
            region: row["D"],
            raw: row,
          }))
          .filter((item) => item.shortCode);

        resolve(formattedData);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

/* =========================
   EXPORT UTILITY
========================= */

export const exportToExcel = (data, fileName) => {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Report");
  XLSX.writeFile(workbook, `${fileName}.xlsx`);
};
