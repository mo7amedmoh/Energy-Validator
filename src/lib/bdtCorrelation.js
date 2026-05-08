/**
 * Rule 8: BDT & Alarm Correlation
 * Cross-references BDT tests with actual network alarm history.
 */

const clean = (str) =>
  str
    ?.toString()
    .replace(/[^a-zA-Z0-9]/g, "")
    .toUpperCase() || "";

/**
 * Robust date parser to handle different browser and locale formats
 */
const safeDate = (input) => {
  if (!input) return null;
  if (input instanceof Date) return isNaN(input.getTime()) ? null : input;
  const d = new Date(input);
  return isNaN(d.getTime()) ? null : d;
};

/**
 * Converts duration strings like "3h 7m" or raw numbers into total minutes
 */
const parseDurationMins = (dur) => {
  if (typeof dur === "number") return Math.round(dur);
  if (!dur || typeof dur !== "string") return 0;

  let total = 0;
  const hMatch = dur.match(/(\d+)\s*h/i);
  const mMatch = dur.match(/(\d+)\s*m/i);

  if (hMatch) total += parseInt(hMatch[1]) * 60;
  if (mMatch) total += parseInt(mMatch[1]);

  if (total === 0) {
    const rawNum = parseInt(dur);
    if (!isNaN(rawNum)) total = rawNum;
  }
  return total;
};

const isSameDay = (d1, d2) => {
  if (!d1 || !d2) return false;
  return (
    d1.getFullYear() === d2.getFullYear() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate()
  );
};

/**
 * Correlates BDT results with standard alarm history (currentAlarms).
 */
export function correlateBDTWithAlarms(bdtResults, currentAlarms) {
  if (!currentAlarms || currentAlarms.length === 0) return bdtResults;

  return bdtResults.map((fileResult) => {
    const updatedSheets = fileResult.sheets.map((sheet) => {
      // Basic requirements for correlation
      if (!sheet.siteName || !sheet.testDate || sheet.startHour === null) {
        return sheet;
      }

      const siteId = clean(sheet.siteName);
      
      // 0. Filter alarms for this site
      const siteAlarms = currentAlarms.filter(
        (a) => a.siteCode && clean(a.siteCode) === siteId
      );

      // If no alarms found at all for the site, we can't apply Rule 8
      if (siteAlarms.length === 0 || sheet.ignoreMostRules) return sheet;

      // Parse BDT timing
      const bdtDate = safeDate(sheet.testDate);
      if (!bdtDate) return sheet; // Date parsing failed
      
      const bdtDuration = parseInt(sheet.busbarReadings?.testedBackup) || 0;

      const rule8 = {
        status: "pass",
        issues: [],
      };

      // --- 1. Door Alarm Verification ---
      // If door alarm occurred on the same day => accepted
      const hasDoor = siteAlarms.some((a) => {
        const ast = safeDate(a.startTime);
        if (!ast) return false;
        const isDoor = a.category === "door" || a.detectedType === "door" || (a.alarmName || "").toLowerCase().includes("door");
        return isDoor && isSameDay(ast, bdtDate);
      });

      if (!hasDoor) {
        rule8.issues.push(
          "Rule 8: No Door Alarm found on the BDT date"
        );
      }

      // --- 2. Power Alarm Matching by Duration on Same Day ---
      const powerAlarmsOnDay = siteAlarms.filter((a) => {
        const ast = safeDate(a.startTime);
        if (!ast) return false;
        const isPower = a.category === "power" || a.detectedType === "power" || (a.alarmName || "").toLowerCase().includes("mains") || (a.alarmName || "").toLowerCase().includes("ac fail");
        return isPower && isSameDay(ast, bdtDate);
      });

      if (powerAlarmsOnDay.length === 0) {
        rule8.issues.push(
          "Rule 8: No Power Alarm found on the BDT date"
        );
      }

      // Find a power alarm that matches BDT duration with tolerance
      let matchingPower = null;
      let matchedPowerDuration = 0;

      for (const a of powerAlarmsOnDay) {
        const pDur = parseDurationMins(
          a.durationMinutes || a.duration || a.durationRaw
        );
        // Tolerance: Power alarm can be slightly shorter (-5m) or longer (+15m) than BDT
        if (pDur >= bdtDuration - 5 && pDur <= bdtDuration + 15) {
          matchingPower = a;
          matchedPowerDuration = pDur;
          break;
        }
      }

      let actualBdtStartTime = null;
      let actualBdtEndTime = null;

      if (powerAlarmsOnDay.length > 0) {
        if (!matchingPower) {
           rule8.issues.push(
             `Rule 8: Found Power alarms on BDT date, but none matched the BDT duration (${bdtDuration}m)`
           );
        } else {
           actualBdtStartTime = safeDate(matchingPower.startTime);
           actualBdtEndTime = new Date(actualBdtStartTime.getTime() + matchedPowerDuration * 60000);
        }
      }

      // --- 3. Down Alarm Check using the Matched Power Alarm Window ---
      if (actualBdtStartTime && actualBdtEndTime) {
        const siteWentDown = siteAlarms.some((a) => {
          const ast = safeDate(a.startTime);
          if (!ast) return false;
          const isDown = a.category === "down" || a.detectedType === "down" || (a.alarmName || "").toLowerCase().includes("down");
          return isDown && ast >= actualBdtStartTime && ast <= actualBdtEndTime;
        });

        if (siteWentDown) {
          rule8.issues.push(
            "Rule 8: Critical rejection – Site Down alarm detected during matched BDT time"
          );
        }
      }

      if (rule8.issues.length > 0) {
        rule8.status = "fail";
      }

      const allIssues = [
        ...sheet.allIssues,
        ...rule8.issues.map((i) => ({ section: "Network Correlation", text: i })),
      ];

      return {
        ...sheet,
        rule8,
        overallStatus: rule8.status === "fail" ? "fail" : sheet.overallStatus,
        allIssues,
        issueCount: allIssues.length,
      };
    });

    return {
      ...fileResult,
      sheets: updatedSheets,
      overallStatus: updatedSheets.some((s) => s.overallStatus === "fail") ? "fail" : "pass",
    };
  });
}
