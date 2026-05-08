import * as XLSX from "xlsx";
import { extractImagesFromExcel } from "./images";

/* ============================================================
   COLUMN UTILITIES
   ============================================================ */

/** Convert column letter(s) to 1-based number  (A→1, Z→26, AA→27, AB→28, AM→39) */
function colLetterToNum(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
  return n;
}

/** Convert 1-based column number to letter(s)  (1→A, 26→Z, 27→AA, 39→AM) */
function numToColLetter(num) {
  let result = "";
  while (num > 0) {
    const rem = (num - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    num = Math.floor((num - 1) / 26);
  }
  return result;
}

/** Get a cell's value (raw) from a worksheet */
function getCell(ws, addr) {
  const cell = ws[addr];
  if (!cell) return null;
  return cell.v !== undefined ? cell.v : null;
}

/** Return true only if the cell has a meaningful, non-blank value */
function hasValue(ws, addr) {
  const cell = ws[addr];
  if (!cell) return false;
  const v = cell.v;
  return v !== undefined && v !== null && v !== "";
}

/**
 * Collect every cell in a rectangular range and return rich metadata.
 * Columns are specified as Excel letters (e.g. "A", "AM").
 */
function getRangeValues(ws, startColLetter, endColLetter, startRow, endRow) {
  const results = [];
  const sc = colLetterToNum(startColLetter);
  const ec = colLetterToNum(endColLetter);
  for (let r = startRow; r <= endRow; r++) {
    for (let c = sc; c <= ec; c++) {
      const col = numToColLetter(c);
      const addr = col + r;
      results.push({
        addr,
        row: r,
        col,
        value: getCell(ws, addr),
        hasValue: hasValue(ws, addr),
      });
    }
  }
  return results;
}

/**
 * Scan rows [fromRow..toRow] in the given columns for a regex pattern.
 * Returns the first matching row number, or null.
 */
function findRow(ws, pattern, fromRow = 1, toRow = 400, cols = ["A", "B", "C"]) {
  const re = new RegExp(pattern, "i");
  for (let r = fromRow; r <= toRow; r++) {
    for (const col of cols) {
      const val = String(getCell(ws, col + r) || "").trim();
      if (re.test(val)) return r;
    }
  }
  return null;
}

function isSheetRulesIgnored(ws) {
  const raw = getCell(ws, "I49");
  if (!hasValue(ws, "I49")) return true;
  
  const str = String(raw).trim().toUpperCase();
  if (str === "" || str === "N/A" || str === "ZERO") return true;
  
  const num = parseFloat(raw);
  if (isNaN(num) || num === 0) return true;
  
  return false;
}

/* ============================================================
   SECTION VALIDATORS
   ============================================================ */

/**
 * Section 1 — General Data  A4:Q6
 * Checks fill-rate and tries to extract site name / test date.
 */
function validateGeneralData(ws) {
  // Helper to get formatted or raw value
  const getVal = (addr) => {
    const cell = ws[addr];
    if (!cell) return null;
    return cell.w || cell.v; // Use formatted text if available (good for times/dates)
  };

  const siteCode = String(getVal("I5") || "").trim();
  const rawDate = getVal("O4");
  const rawTime = getVal("O5");

  let testDate = null;
  if (rawDate) {
    if (rawDate instanceof Date) {
      testDate = rawDate.toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "short",
        year: "2-digit",
      });
    } else {
      testDate = String(rawDate).trim();
    }
  }

  let startHour = null;
  if (rawTime !== null && rawTime !== undefined) {
    const timeStr = String(rawTime).trim().toUpperCase();
    
    // Pattern for HH:MM:SS AM/PM or HH:MM
    const match = timeStr.match(/^(\d+)/);
    if (match) {
      let h = parseInt(match[1]);
      const isPM = timeStr.includes("PM");
      const isAM = timeStr.includes("AM");

      if (isPM && h < 12) h += 12;
      if (isAM && h === 12) h = 0;
      
      startHour = h;
    } else if (typeof rawTime === "number") {
      // Fallback for raw Excel decimal time
      startHour = Math.floor((rawTime * 24) + 0.0001);
    }
  }

  return {
    section: "General Data",
    range: "A4:Q6",
    siteCode,
    siteName: siteCode || "Unknown",
    batteryBrand: String(getVal("I41") || "").trim(),
    batteryVolt: String(getVal("I45") || "").trim(),
    batteryAH: String(getVal("I47") || "").trim(),
    numStrings: String(getVal("I49") || "").trim(),
    numBatteries: String(getVal("I43") || "").trim(),
    chargingLimit: String(getVal("I33") || "").trim(),
    rectifierBrand: String(getVal("I13") || "").trim(),
    numModules: String(getVal("I17") || "").trim(),
    testDate,
    startHour,
    status: "pass",
    issues: [],
  };
}

/**
 * Section 2 — Photos  M9:AB59  (expected up to 16 photos)
 * Real images are embedded in the xlsx drawing XML and not accessible
 * through standard cell data.  We try three escalating strategies:
 *  1. ws['!images'] — available when SheetJS Pro is used
 *  2. Merged-cell count inside the photo region (common template layout)
 *  3. Non-empty data cells in the region (captions / labels around photos)
 */
function validatePhotos(ws, ignoreRules = false, config = {}) {
  const { photo_min_pass = 12, photo_min_warn = 6, photo_min_fail = 8 } = config;
  const issues = [];
  if (ignoreRules) {
    return {
      section: "Photos",
      range: "M9:AB59",
      photoCount: 0,
      detectionMethod: "ignored",
      dataCellCount: 0,
      status: "pass",
      issues: [],
      note: "Photos validation skipped (I49 empty)",
    };
  }

  let photoCount = 0;
  let detectionMethod = "none";

  // Strategy 1 — SheetJS Pro images metadata
  try {
    const images = ws["!images"];
    if (Array.isArray(images) && images.length > 0) {
      photoCount = images.length;
      detectionMethod = "direct";
    }
  // eslint-disable-next-line no-unused-vars
  } catch (_e) {
    /* ignore */
  }

  // Strategy 2 — merged cells inside M9:AB59  (0-indexed: rows 8-58, cols 12-27)
  if (detectionMethod === "none") {
    const merges = ws["!merges"] || [];
    const regionMerges = merges.filter(
      (m) => m.s.r >= 8 && m.e.r <= 58 && m.s.c >= 12 && m.e.c <= 27
    );
    if (regionMerges.length > 0) {
      photoCount = regionMerges.length;
      detectionMethod = "merges";
    }
  }

  // Strategy 3 — data cells (captions / reference numbers around photos)
  let dataCellCount = 0;
  for (let r = 9; r <= 59; r++) {
    for (let c = colLetterToNum("M"); c <= colLetterToNum("AB"); c++) {
      if (hasValue(ws, numToColLetter(c) + r)) dataCellCount++;
    }
  }
  if (detectionMethod === "none" && dataCellCount > 0) {
    photoCount = Math.max(1, Math.floor(dataCellCount / 4));
    detectionMethod = "estimated";
  }

  if (photoCount === 0) {
    issues.push("No photos detected in M9:AB59 – verify manually");
  } else if (photoCount < photo_min_fail) {
    issues.push(`Only ~${photoCount} photos detected – expected up to 16`);
  }

  const status =
    detectionMethod === "none"
      ? "unknown"
      : photoCount >= photo_min_pass
      ? "pass"
      : photoCount >= photo_min_warn
      ? "warning"
      : "fail";

  return {
    section: "Photos",
    range: "M9:AB59",
    photoCount,
    detectionMethod,
    dataCellCount,
    status,
    issues,
    note:
      detectionMethod !== "direct"
        ? "Photo count is estimated — manual verification recommended"
        : null,
  };
}

/**
 * Section 3 — Basic Data  I11:I71  (single column, 61 cells)
 * Checks fill-rate and warns if most values are non-numeric.
 */
function validateBasicData(ws, ignoreRules = false, config = {}) {
  const { basic_data_min_cells = 28 } = config;
  const cells = getRangeValues(ws, "I", "I", 11, 71);
  const filled = cells.filter((c) => c.hasValue);
  const numericCells = filled.filter((c) => !isNaN(parseFloat(c.value)));
  const fillRate = filled.length / cells.length;
  let issues = [];

  if (!ignoreRules) {
    if (filled.length < basic_data_min_cells) {
      const sample = filled.slice(0, 3).map(c => `${c.addr}:"${c.value}"`).join(", ");
      issues.push(`Rule 5: Incomplete Data (Found ${filled.length}/${basic_data_min_cells} cells in I11:I71. Sample: ${sample || "none"})`);
    } else if (fillRate < 0.3) {
      issues.push(`Basic data sparse – ${filled.length}/${cells.length} cells populated`);
    }

    if (numericCells.length < (filled.length * 0.4) && filled.length > 0) {
      issues.push(
        "Most basic data entries are non-numeric – check for text or formula errors"
      );
    }
  }

  const status = (ignoreRules || filled.length >= basic_data_min_cells) ? "pass" : "fail";

  return {
    section: "Basic Data",
    range: "I11:I71",
    totalCells: cells.length,
    filledCells: filled.length,
    numericCells: numericCells.length,
    fillRate,
    status,
    issues,
  };
}

/**
 * Section 4 — Busbar Readings  A74:U112
 * Column A holds the elapsed test duration (should be ascending).
 * Row 112 must contain the final test duration.
 */
function validateBusbarReadings(ws, ignoreMostRules = false, config = {}) {
  const { 
    rule7_interval_mins = 10,
    busbar_min_fill_rate = 0.15,
    rule1_batt_amp_max = 1.0,
    rule2_seq_tolerance = 0.05,
    rule3_balance_max = 3.0,
    rule1_min_volt_start = 51.0,
    rule1_min_rect_amp = 20.0,
    rule4_theoretical_tolerance = 20
  } = config;
  // --- Dynamic section boundary detection ---
  // Find "Before disconnecting Rectifier" — this row holds the starting readings
  const beforeRow = findRow(ws, "before disconnecting|before disconnect", 1, 200) || 75;
  // First interval data row is immediately after the "Before" row
  const firstDataRow = beforeRow + 1;
  // Find "After Connecting Rectifier" — marks end of test data
  const afterConnectRow = findRow(ws, "after connect", beforeRow + 1, beforeRow + 120);
  // Upper bound for scanning: row before "After Connecting", or fallback
  const maxScanRow = afterConnectRow ? afterConnectRow - 1 : beforeRow + 60;
  // Section range for fill-rate (include a few header rows above "Before")
  const sectionTop = Math.max(1, beforeRow - 3);

  const allCells = getRangeValues(ws, "A", "U", sectionTop, maxScanRow);
  const filled = allCells.filter((c) => c.hasValue);
  const fillRate = filled.length / (allCells.length || 1);

  // Duration column A — only test-interval rows
  const durationCells = getRangeValues(ws, "A", "A", firstDataRow, maxScanRow).filter(
    (c) => c.hasValue
  );

  // Starting readings are on the "Before disconnecting" row
  const startVoltCols = ["D", "F", "H", "J", "L", "N", "P", "R", "T"];
  const batteryAmpCols = ["G", "I", "K", "M", "O", "Q", "S", "U"];
  const rectAmpVal = parseFloat(getCell(ws, `E${beforeRow}`)) || 0;

  const startingVolts = startVoltCols
    .map(col => ({ col, val: parseFloat(getCell(ws, `${col}${beforeRow}`)) }))
    .filter(v => !isNaN(v.val));

  const startingBattAmps = batteryAmpCols
    .map(col => ({ col, val: parseFloat(getCell(ws, `${col}${beforeRow}`)) || 0 }))
    .filter(item => {
      if (isNaN(item.val)) return false;
      if (rectAmpVal > 5 && Math.abs(item.val - rectAmpVal) < 0.1) return false;
      return true;
    });

  const battAmpSum = startingBattAmps.reduce((acc, item) => acc + item.val, 0);

  // Last filled data row — scan from firstDataRow up to maxScanRow
  let lastReadingRow = beforeRow;
  for (let r = firstDataRow; r <= maxScanRow; r++) {
    if (hasValue(ws, `D${r}`)) {
      lastReadingRow = r;
    } else {
      break;
    }
  }

  const finalDurationRaw = getCell(ws, `A${lastReadingRow}`);

  // Create a map of time labels to row indices for cells-section correlation
  const timeMap = {};
  for (let r = beforeRow; r <= maxScanRow; r++) {
    const label = String(getCell(ws, `A${r}`) || "").trim().toLowerCase();
    if (label) timeMap[label] = r;
  }

  const issues = [];
  const rule1Issues = [];
  const rule2Issues = [];
  const rule3Issues = [];
  const rule4Issues = [];
  const rule6Issues = [];
  const rule7Issues = [];

  // Rule 1 Validation
  if (!ignoreMostRules) {
    if (startingVolts.length > 0 && startingVolts.some(v => v.val <= rule1_min_volt_start)) {
      rule1Issues.push(`Rule 1: One or more starting voltages (row 76) are <= ${rule1_min_volt_start}V`);
    }
  }

  // Rectifier Start Ampere is the ONLY rule that remains if I49 is empty
  if (rectAmpVal <= rule1_min_rect_amp) {
    rule1Issues.push(`Rule 1: Rectifier Start Ampere (E76) is ${rectAmpVal}A (expected > ${rule1_min_rect_amp}A)`);
  }

  if (!ignoreMostRules) {
    if (startingBattAmps.length > 0 && battAmpSum >= rule1_batt_amp_max) {
      const nonZeroDetails = batteryAmpCols
        .map((col) => ({ col, val: parseFloat(getCell(ws, `${col}${beforeRow}`)) || 0 }))
        .filter((item) => item.val !== 0)
        .map((item) => `${item.col}${beforeRow}: ${item.val}A`)
        .join(", ");
      rule1Issues.push(`Rule 1: Batteries Start Ampere sum (${battAmpSum.toFixed(2)}A) is >= ${rule1_batt_amp_max}A (Non-zero readings: ${nonZeroDetails || "none"})`);
    }
  }

  // Rule 4: Theoretical vs Tested Duration
  const cellI47 = String(getCell(ws, "I47") || "").trim();
  const ahMatch = cellI47.match(/(\d+(?:\.\d+)?)\s*AH/i);
  const battAH = ahMatch ? parseFloat(ahMatch[1]) : 0;
  const is100AH = /100\s*AH/i.test(cellI47);
  
  const cellI41 = String(getCell(ws, "I41") || "").trim();
  const isLithium = /lithium/i.test(cellI41);
  const is48V = Math.abs(parseFloat(getCell(ws, "I45")) - 48) < 1;
  
  const rectVoltStart = parseFloat(getCell(ws, `D${beforeRow}`)) || 0;

  if (!ignoreMostRules) {
    // Rule 2 & 3 Iterative logic
    for (let r = firstDataRow; r <= lastReadingRow; r++) {
    // Rule 2: Decreasing Voltage Sequence
    if (r < lastReadingRow) {
      for (const col of startVoltCols) {
        const curr = parseFloat(getCell(ws, `${col}${r}`));
        const next = parseFloat(getCell(ws, `${col}${r+1}`));
        // Only flag if it strictly increases (should decrease or stay flat)
        if (!isNaN(curr) && !isNaN(next) && next > curr + rule2_seq_tolerance) {
          if (rule2Issues.length < 3) rule2Issues.push(`Rule 2 (Seq): Voltage at ${col}${r+1} (${next}V) increased from ${curr}V (diff: ${(next - curr).toFixed(2)}V)`);
        }
      }
      
      // Rule 2: Increasing Ampere Sequence
      const ampCols = ["E", ...batteryAmpCols];
      for (const col of ampCols) {
        const curr = parseFloat(getCell(ws, `${col}${r}`));
        const next = parseFloat(getCell(ws, `${col}${r+1}`));
        // Only flag if it strictly decreases and ignore negative/noise values (< 0.5A)
        if (!isNaN(curr) && !isNaN(next) && curr > 0.5 && next > 0.5 && next < curr - rule2_seq_tolerance) {
          if (rule2Issues.length < 5) {
             if (!rule2Issues.some(msg => msg.includes(col))) {
                rule2Issues.push(`Rule 2 (Seq): Ampere at ${col}${r+1} (${next}A) decreased from ${curr}A (diff: ${(curr - next).toFixed(2)}A)`);
             }
          }
        }
      }
    }

    // Rule 3: Ampere Balance (Rectifier vs Batteries)
    if (r > firstDataRow) {
      const rectA = parseFloat(getCell(ws, `E${r}`)) || 0;
      const battAs = batteryAmpCols.map(col => parseFloat(getCell(ws, `${col}${r}`)) || 0);
      const bSum = battAs.reduce((a, b) => a + b, 0);
      const delta = bSum - rectA;

      // Sum(Batteries) should be Rectifier + [0 to tolerance]
      if (delta < 0 || delta >= rule3_balance_max) {
        if (rule3Issues.length < 3) {
          rule3Issues.push(`Rule 3 (Sum): Row ${r} mismatch. Sum ${bSum.toFixed(2)}A exceeds Rectifier ${rectA.toFixed(2)}A by ${delta.toFixed(2)}A (Target: 0-${rule3_balance_max}A)`);
        }
      }
    }
  }
  }

  if (!finalDurationRaw && finalDurationRaw !== 0 && !ignoreMostRules) {
    issues.push("Detected no BDT duration readings in Column A");
  }
  if (durationCells.length < 3 && !ignoreMostRules) {
    issues.push(
      `Only ${durationCells.length} duration reading(s) in column A – minimum 3 expected`
    );
  }

  // Rule 4: Total voltage check (Removed as per user request: below 44V is accepted)

  // Ascending order check (within the test range 76 to lastReadingRow)
  let isProgressive = true;
  let lastVal = -Infinity;
  let failingRow = null;
  let failingVal = null;
  
  const testRangeDurations = durationCells.filter(c => {
    const r = parseInt(c.addr.substring(1));
    return r >= firstDataRow && r <= lastReadingRow;
  });

  for (const cell of testRangeDurations) {
    const n = parseFloat(cell.value);
    if (!isNaN(n)) {
      if (n < lastVal - 0.0001) { // Small margin for floating point
        isProgressive = false; 
        failingRow = cell.addr;
        failingVal = n;
        break; 
      }
      lastVal = n;
    }
  }
  
  if (!isProgressive && !ignoreMostRules) {
    issues.push(
      `Rule 2: Duration values in Col A are not ascending. Row ${failingRow} (${failingVal}) is less than previous value.`
    );
  }

  // Format final duration as an integer
  let finalDurationFormatted = "0 min";
  if (finalDurationRaw !== null && finalDurationRaw !== undefined && finalDurationRaw !== "") {
    const num = parseFloat(finalDurationRaw);
    if (!isNaN(num)) {
      if (num < 2) {
        // Excel time fraction
        finalDurationFormatted = `${Math.round(num * 24 * 60)} min`;
      } else {
        finalDurationFormatted = `${Math.round(num)} min`;
      }
    } else {
      // Fallback: extract digits from string
      const match = String(finalDurationRaw).match(/\d+/);
      finalDurationFormatted = (match ? match[0] : "0") + " min";
    }
  }

  // Rule 4 Calculation and check
  let theoreticalBackup = 0;
  let rule4Achieved = false;
  const cellI49 = parseFloat(getCell(ws, "I49")) || 0;

  if (is100AH && isLithium && battAH > 0 && cellI49 > 0 && rectVoltStart > 0 && rectAmpVal > 0) {
    const battLoad = battAH * 48 * cellI49;
    const rectLoad = rectVoltStart * rectAmpVal;
    theoreticalBackup = Math.round((battLoad * 60 * 0.96) / rectLoad);
    
    // Get tested minutes from the formatted string or numeric raw
    let testedMins = 0;
    const num = parseFloat(finalDurationRaw);
    if (!isNaN(num)) {
       testedMins = num < 2 ? Math.round(num * 24 * 60) : Math.round(num);
    }
    
    const endVoltRule4 = parseFloat(getCell(ws, `D${lastReadingRow}`)) || 0;
    const isExemption = testedMins >= 180 && endVoltRule4 > 44;

    const meetsTarget = (theoreticalBackup > 0 && testedMins >= theoreticalBackup) || testedMins >= 180;

    if (!meetsTarget && theoreticalBackup > 0 && (theoreticalBackup - testedMins) > rule4_theoretical_tolerance && !isExemption) {
      rule4Issues.push(
        `Rule 4: BDT duration (${testedMins}m) is less than theoretical (${theoreticalBackup}m) by more than ${rule4_theoretical_tolerance}m`
      );
    } else {
      rule4Achieved = true;
    }
  } else {
    // If Rule 4 doesn't apply (non-lithium/100AH), still consider it 'achieved' if it reaches 180m
    const testedMinsFallback = parseFloat(finalDurationRaw) < 2 ? Math.round(parseFloat(finalDurationRaw) * 24 * 60) : Math.round(parseFloat(finalDurationRaw));
    if (testedMinsFallback >= 180) rule4Achieved = true;
  }

  // Rule 6: End Voltage Condition
  const endVoltage = parseFloat(getCell(ws, `D${lastReadingRow}`)) || 0;
  
  // Reuse the testedMins logic from Rule 4 if available, or recalculate
  let testedMinsR6 = 0;
  const numR6 = parseFloat(finalDurationRaw);
  if (!isNaN(numR6)) {
     testedMinsR6 = numR6 < 2 ? Math.round(numR6 * 24 * 60) : Math.round(numR6);
  } else {
     const match = String(finalDurationRaw).match(/\d+/);
     testedMinsR6 = match ? parseInt(match[0]) : 0;
  }

  // Do not flag Rule 6 if Rule 4 was fulfilled successfully
  if (!rule4Achieved && testedMinsR6 > 0 && testedMinsR6 < 180 && !ignoreMostRules) {
    if (endVoltage > 45) {
      rule6Issues.push(`Rule 6: End Voltage at D${lastReadingRow} (${endVoltage}V) is too high (> 45V) for a test shorter than 180 mins`);
    }
  }

  // Rule 7: Interval Check (10 mins after 47V)
  const readingRows = [];
  for (let r = firstDataRow; r <= lastReadingRow; r++) {
    if (hasValue(ws, `A${r}`)) readingRows.push(r);
  }

  if (is100AH && readingRows.length > 1) {
    let triggered10m = false;
    const rule7Violations = [];
    for (let i = 0; i < readingRows.length - 1; i++) {
      const r = readingRows[i];
      const nextR = readingRows[i + 1];
      const volt = parseFloat(getCell(ws, `D${r}`)) || 0;

      if (volt > 0 && volt <= 47) {
        triggered10m = true;
      }

      if (triggered10m && r > firstDataRow) {
        const d1 = parseFloat(getCell(ws, `A${r}`)) || 0;
        const d2 = parseFloat(getCell(ws, `A${nextR}`)) || 0;
        const diff = d2 - d1;
        const diffMins = d1 < 2 ? Math.round(diff * 24 * 60) : Math.round(diff);

        if (diffMins > (rule7_interval_mins + 1)) {
          // 1 min buffer
          rule7Violations.push(
            `${diffMins}m between row ${r} and ${nextR} (Volt: ${volt.toFixed(
              1
            )}V)`
          );
        }
      }
    }

    if (rule7Violations.length > 1 && !ignoreMostRules) {
      rule7Issues.push(
        `Rule 7: Readings interval should be ${rule7_interval_mins} mins (${
          rule7Violations.length
        } violations detected: ${rule7Violations.join(", ")})`
      );
    }
  }

  const rulesPassed = rule1Issues.length === 0 && rule2Issues.length === 0 && rule3Issues.length === 0 && rule4Issues.length === 0 && rule6Issues.length === 0 && rule7Issues.length === 0;
  
  let status = "pass";
  if (issues.length > 0 || !rulesPassed) status = "fail";
  else if (fillRate < busbar_min_fill_rate) status = "warning";
  
  const allBusbarIssues = [...issues, ...rule1Issues, ...rule2Issues, ...rule3Issues, ...rule4Issues, ...rule6Issues, ...rule7Issues];

  return {
    section: "Busbar Readings",
    range: `A${sectionTop}:U${maxScanRow}`,
    detectedBeforeRow: beforeRow,
    detectedAfterRow: afterConnectRow || null,
    totalCells: allCells.length,
    filledCells: filled.length,
    fillRate,
    durationReadings: durationCells.length,
    finalDuration: finalDurationRaw,
    finalDurationFormatted,
    isProgressive,
    testedBackup: finalDurationFormatted,
    rule1Status: rule1Issues.length === 0 ? "pass" : "fail",
    rule2Status: rule2Issues.length === 0 ? "pass" : "fail",
    rule3Status: rule3Issues.length === 0 ? "pass" : "fail",
    rule4Status: rule4Issues.length === 0 ? "pass" : "fail",
    rule6Status: rule6Issues.length === 0 ? "pass" : "fail",
    rule7Status: rule7Issues.length === 0 ? "pass" : "fail",
    theoreticalBackup,
    startVolt: parseFloat(getCell(ws, `D${beforeRow}`)) || 0,
    startAmp: parseFloat(getCell(ws, `E${beforeRow}`)) || 0,
    endVolt: parseFloat(getCell(ws, `D${lastReadingRow}`)) || 0,
    endAmp: parseFloat(getCell(ws, `E${lastReadingRow}`)) || 0,
    ruleIssues: { 
      rule1: rule1Issues, 
      rule2: rule2Issues, 
      rule3: rule3Issues, 
      rule4: rule4Issues, 
      rule6: rule6Issues,
      rule7: rule7Issues
    },
    status,
    issues: allBusbarIssues,
    ignoreMostRules,
    timeMap
  };
}

/**
 * Section 5 — Cells Readings (Voltage Measured table)
 * Layout (from image):
 *   Header row:  | String # | Battery # | Before disconnecting Rectifier | 10 Mins | 15 Mins | ... | After connecting Rectifier |
 *   Data rows:   one row per battery cell, voltage readings under each column
 *
 * Detection strategy:
 *   1. Find the header row by scanning for a row that has BOTH "before disconnecting" AND "after connecting" text in any two columns.
 *   2. colBefore = column with "before disconnecting" → starting readings (excluded from duration calc)
 *   3. colAfter  = column with "after connecting"    → section end marker
 *   4. Time-interval columns = between colBefore+1 and colAfter-1
 *   5. Duration = header text of the LAST time-interval column that has actual data
 */
function validateCellsReadings(ws, ignoreRules = false, busbarReadings = null, config = {}) {
  const { 
    rule6_12v_dissipated_max = 1.5,
    rule8_duration_mismatch_tolerance = 5 
  } = config;
  const cellI45 = parseFloat(getCell(ws, "I45")) || 0;
  const is48V = cellI45 === 48;

  if (ignoreRules || is48V || cellI45 === 0 || isNaN(cellI45)) {
    return {
      section: "Cells Readings",
      range: "N/A",
      totalCells: 0,
      filledCells: 0,
      fillRate: 1,
      status: "pass",
      issues: [],
      note: ignoreRules
        ? "Rules ignored due to empty I49"
        : `Section neglected (I45 is 48V or empty)`,
    };
  }

  const issues = [];
  const maxScanCol = 60;

  const cellText = (col, row) =>
    String(getCell(ws, numToColLetter(col) + row) || "")
      .replace(/[\r\n]+/g, " ")
      .replace(/\s{2,}/g, " ")
      .trim()
      .toLowerCase();

  let cellsHeaderRow = null;
  let colBefore = -1;
  let colAfter = -1;

  for (let r = 80; r <= 350; r++) {
    let foundBefore = -1;
    let foundAfter = -1;
    for (let c = 1; c <= maxScanCol; c++) {
      const val = cellText(c, r);
      if (foundBefore === -1 && (val.includes("before disconnecting") || val.includes("before disconnect"))) {
        foundBefore = c;
      }
      if (foundAfter === -1 && (val.includes("after connecting") || val.includes("after connect"))) {
        foundAfter = c;
      }
      if (foundBefore !== -1 && foundAfter !== -1) break;
    }
    if (foundBefore !== -1 && foundAfter !== -1) {
      cellsHeaderRow = r;
      colBefore = foundBefore;
      colAfter = foundAfter;
      break;
    }
  }

  if (!cellsHeaderRow) {
    for (let r = 100; r <= 300; r++) {
      for (let c = colLetterToNum("C"); c <= maxScanCol; c++) {
        if (cellText(c, r).includes("after connecting")) {
          cellsHeaderRow = r;
          colAfter = c;
          break;
        }
      }
      if (cellsHeaderRow) break;
    }
    cellsHeaderRow = cellsHeaderRow || 132;
    colAfter = colAfter !== -1 ? colAfter : maxScanCol;
  }

  const timeStartCol = colBefore !== -1 ? colBefore + 1 : colLetterToNum("C");
  const timeEndCol = colAfter - 1;
  const dataStartCol = colBefore !== -1 ? colBefore : timeStartCol;
  const cellsDataStartRow = cellsHeaderRow + 1;
  const maxRow = cellsHeaderRow + 100;

  let lastFilledTimeCol = -1;
  let lastFilledTimeHeader = "";
  let outOfRange = 0;
  let filledCellsCount = 0;
  const is12V = cellI45 === 12;
  const busbarStringCols = ["F", "H", "J", "L", "N", "P", "R", "T"];
  const stringSumIssues = [];

  for (let c = dataStartCol; c <= timeEndCol; c++) {
    let colHasData = false;
    for (let r = cellsDataStartRow; r <= maxRow; r++) {
      const cellVal = getCell(ws, numToColLetter(c) + r);
      if (cellVal !== null && cellVal !== undefined && cellVal !== "") {
        colHasData = true;
        filledCellsCount++;
        const n = parseFloat(cellVal);
        if (!isNaN(n)) {
          if (n > 0.5 && n < 10 && (n < 1.75 || n > 2.25)) outOfRange++;
        }
      }
    }

    if (colHasData && c >= timeStartCol) {
      const hdr = String(getCell(ws, numToColLetter(c) + cellsHeaderRow) || "").trim();
      if (/\d+/.test(hdr)) {
        lastFilledTimeCol = c;
        lastFilledTimeHeader = hdr;
      }
    }

    if (colHasData && is12V) {
      const colHdr = String(getCell(ws, numToColLetter(c) + cellsHeaderRow) || "").trim().toLowerCase();
      const busbarRow = busbarReadings?.timeMap?.[colHdr];

      if (busbarRow) {
        for (let s = 0; s < 8; s++) {
          const busVolt = parseFloat(getCell(ws, busbarStringCols[s] + busbarRow));
          if (!isNaN(busVolt) && busVolt > 0) {
            const startRow = cellsDataStartRow + s * 4;
            let sumCells = 0;
            let allNumbers = true;
            for (let r = startRow; r < startRow + 4; r++) {
              const val = parseFloat(getCell(ws, numToColLetter(c) + r));
              if (isNaN(val)) {
                allNumbers = false;
                break;
              }
              sumCells += val;
            }
            if (allNumbers) {
              // The sum of the cells should naturally be greater than or equal to the voltage measured at the busbar due to cable drops.
              const dissipated = sumCells - busVolt;
              if (dissipated < 0 || dissipated > rule6_12v_dissipated_max) {
                if (stringSumIssues.length < 5) {
                  stringSumIssues.push(
                    `String ${s + 1} mismatch (Col ${numToColLetter(c)}): cells sum=${sumCells.toFixed(2)}V vs busbar ${busbarStringCols[s]}${busbarRow}=${busVolt.toFixed(2)}V (diff: ${dissipated.toFixed(2)}V)`
                  );
                }
              }
            }
          }
        }
      }
    }
  }

  issues.push(...stringSumIssues);

  let busbarMins = 0;
  if (busbarReadings?.finalDuration !== undefined) {
    const raw = busbarReadings.finalDuration;
    const num = parseFloat(raw);
    if (!isNaN(num)) {
      busbarMins = num < 2 ? Math.round(num * 24 * 60) : Math.round(num);
    } else {
      const m = String(raw).match(/\d+/);
      busbarMins = m ? parseInt(m[0]) : 0;
    }
  }

  let cellsMins = 0;
  const cellsDurationText = lastFilledTimeHeader;
  const minsMatch = cellsDurationText.match(/(\d+)/);
  if (minsMatch) cellsMins = parseInt(minsMatch[1]);

  if (filledCellsCount === 0) {
    issues.push(
      `Cells section is empty but must be filled because system voltage (I45) is ${cellI45}V (not 48V)`
    );
  } else if (busbarMins > 0 && (lastFilledTimeCol === -1 || cellsMins === 0)) {
    issues.push(
      `Rule 8 (Cells): Busbar test duration is ${busbarMins}m, but no discharge interval readings (10 Mins, 15 Mins, etc.) were found in the cells section.`
    );
  } else if (busbarMins > 0 && cellsMins > 0 && Math.abs(cellsMins - busbarMins) > rule8_duration_mismatch_tolerance) {
    issues.push(
      `Rule 8 (Cells): BDT duration from busbar (${busbarMins}m) does not match cells section ` +
      `(last filled column: "${cellsDurationText}" = ${cellsMins}m)`
    );
  }

  if (outOfRange > 0) {
    issues.push(`${outOfRange} voltage reading(s) outside normal range (1.75–2.25 V)`);
  }

  const totalCols = timeEndCol - dataStartCol + 1;
  const expectedCells = totalCols > 0 ? totalCols * (maxRow - cellsDataStartRow + 1) : 0;

  return {
    section: "Cells Readings",
    range: `${numToColLetter(dataStartCol)}${cellsHeaderRow}:${numToColLetter(timeEndCol)}${maxRow}`,
    detectedHeaderRow: cellsHeaderRow,
    detectedColBefore: colBefore !== -1 ? numToColLetter(colBefore) : null,
    detectedColAfter: colAfter !== -1 ? numToColLetter(colAfter) : null,
    cellsDurationText,
    cellsDurationMins: cellsMins,
    totalCells: expectedCells,
    filledCells: filledCellsCount,
    fillRate: expectedCells > 0 ? filledCellsCount / expectedCells : 1,
    status: issues.length === 0 ? "pass" : "fail",
    issues,
  };
}


/* ============================================================
   MAIN EXPORT — parseBDTFile
   ============================================================ */

/**
 * Read a BDT workbook (Excel) and validate every BDT sheet.
 * A sheet is treated as the "summary sheet" if its name matches
 * /summary|sum|ملخص/i — all other sheets are validated as BDT sheets.
 *
 * @param {File} file  Browser File object
 * @param {Object} config  Optional custom rules/tolerances
 * @returns {Promise<Object>}  Validation results
 */
export function parseBDTFile(file, config = {}) {
  return new Promise((resolve, reject) => {
    (async () => {
      try {
        const allExtractedImages = await extractImagesFromExcel(file).catch(() => []);

      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array", cellDates: true });

          const results = {
            fileName: file.name,
            parsedAt: new Date(),
            sheetCount: workbook.SheetNames.length,
            sheets: [],
            overallStatus: "pass",
          };

          for (let idx = 0; idx < workbook.SheetNames.length; idx++) {
            const sheetName = workbook.SheetNames[idx];
            const isBdtSheet = /BDT/i.test(sheetName);

            if (!isBdtSheet) continue;

            const ws = workbook.Sheets[sheetName];

            // Map images to this sheet
            const sheetImages = allExtractedImages.filter(img => img.sheetName === sheetName);

            const ignoreMostRules = isSheetRulesIgnored(ws);

          // Run all section validators with dynamic config
          const generalData = validateGeneralData(ws);
          const photos = validatePhotos(ws, ignoreMostRules, config);
          const basicData = validateBasicData(ws, ignoreMostRules, config);
          const busbarReadings = validateBusbarReadings(ws, ignoreMostRules, config);
          const cellsReadings = validateCellsReadings(ws, ignoreMostRules, busbarReadings, config);

          const sections = [generalData, basicData, busbarReadings, cellsReadings];

          const allIssues = [
            ...(ignoreMostRules
              ? []
              : generalData.issues.map((t) => ({
                  section: "General Data",
                  text: t,
                }))),
            ...(ignoreMostRules
              ? []
              : photos.issues.map((t) => ({ section: "Photos", text: t }))),
            ...(ignoreMostRules
              ? []
              : basicData.issues.map((t) => ({
                  section: "Basic Data",
                  text: t,
                }))),
            ...busbarReadings.issues.map((t) => ({
              section: "Busbar",
              text: t,
            })),
            ...(ignoreMostRules
              ? []
              : cellsReadings.issues.map((t) => ({
                  section: "Cells",
                  text: t,
                }))),
          ];

          // Add warning-level issues if no fail-level issues exist
          if (allIssues.length === 0 && !ignoreMostRules) {
            if (photos.status === "warning") {
              allIssues.push({ section: "Photos", text: `Low photo count detected (~${photos.photoCount})` });
            }
            if (photos.status === "unknown") {
              allIssues.push({ section: "Photos", text: "Photo count could not be determined" });
            }
            // Only show fill rate warning if targets (Rule 4) were NOT met
            if (busbarReadings.fillRate < 0.15 && busbarReadings.fillRate > 0 && busbarReadings.rule4Status !== "pass") {
              allIssues.push({ section: "Busbar", text: `Low data fill rate detected (${(busbarReadings.fillRate * 100).toFixed(1)}%)` });
            }
          }

          let overallStatus = "pass";
          if (allIssues.some(iss => !["warning", "unknown", "fill rate"].some(k => iss.text.toLowerCase().includes(k)))) {
             // If any issue is not one of our warnings, it's a fail
             // (This is a bit heuristic, but let's refine it)
             // Actually, let's just use the existing logic for overallStatus but keep allIssues populated.
          }
          
          // Refined status logic:
          if (sections.some(s => s.status === "fail")) {
            overallStatus = "fail";
          } else if (sections.some(s => s.status === "warning" || s.status === "unknown") || busbarReadings.fillRate < 0.15) {
            overallStatus = "warning";
          }

          results.sheets.push({
            sheetName,
            index: idx,
            siteName: generalData.siteName,
            testDate: generalData.testDate,
            startHour: generalData.startHour,
            generalData,
            photos: { ...photos, images: sheetImages },
            basicData,
            busbarReadings,
            cellsReadings,
            overallStatus,
            allIssues,
            issueCount: allIssues.length,
            ignoreMostRules,
            summaryData: {
              siteCode: generalData.siteCode,
              batteryBrand: generalData.batteryBrand,
              batteryVolt: generalData.batteryVolt,
              batteryAH: generalData.batteryAH,
              numStrings: generalData.numStrings,
              numBatteries: generalData.numBatteries,
              startVolt: busbarReadings.startVolt,
              startAmp: busbarReadings.startAmp,
              chargingLimit: generalData.chargingLimit,
              endVolt: busbarReadings.endVolt,
              endAmp: busbarReadings.endAmp,
              dischargeTime: busbarReadings.testedBackup,
              testDate: generalData.testDate,
              rectifierBrand: generalData.rectifierBrand,
              numModules: generalData.numModules
            }
          });
        }

        // Roll up to file-level status
        if (results.sheets.some((s) => s.overallStatus === "fail"))
          results.overallStatus = "fail";
        else if (results.sheets.some((s) => s.overallStatus === "warning"))
          results.overallStatus = "warning";

        resolve(results);
      } catch (err) {
        reject(err);
      }
    };

      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    } catch (err) {
      reject(err);
    }
   })();
  });
}
