import { excelToDate } from "./excel";

/**
 * Alarm Analysis Engine
 * Computes top 10 sites by alarm count and by total duration
 * for each category: Power, Down, High Temp, Generator
 */

const CATEGORIES = [
  {
    id: "power",
    name: "Power",
    short: "Power",
    label: "Power Alarms",
    color: "bg-rose-500",
    icon: "⚡",
  },
  {
    id: "down",
    name: "Down",
    short: "Down",
    label: "Down Alarms",
    color: "bg-amber-500",
    icon: "📡",
  },
  {
    id: "high_temp",
    name: "High Temp",
    short: "High Temp",
    label: "High Temp Alarms",
    color: "bg-orange-500",
    icon: "🌡️",
  },
  {
    id: "generator",
    name: "Generator",
    short: "Generator",
    label: "Generator Alarms",
    color: "bg-purple-500",
    icon: "⛽",
  },

  {
    id: "door",
    name: "Door",
    short: "Door",
    label: "Door Alarms",
    color: "bg-teal-500",
    icon: "🚪",
  },
];

export { CATEGORIES };

/**
 * Aggregates alarm data per site per category.
 * @param {Array} alarms - Mapped alarm data
 * @returns {Object} - { power: { bySiteCount, bySiteDuration }, ... }
 */
/**
 * Logic to calculate measured backup for a specific set of alarms.
 */
export function calculateSiteBackup(siteAlarms) {
  const powerAlarms = siteAlarms.filter(
    (a) => a.category === "power" || a.detectedType === "power",
  );
  const downAlarms = siteAlarms.filter(
    (a) => a.category === "down" || a.detectedType === "down",
  );

  if (powerAlarms.length === 0 || downAlarms.length === 0) return null;

  const processedPower = powerAlarms
    .map((p) => ({
      ...p,
      parsedStart: p.startTime instanceof Date ? p.startTime : excelToDate(p.startTime),
      parsedEnd: (p.endTime instanceof Date ? p.endTime : excelToDate(p.endTime)) || new Date(8640000000000000), // Far future for active
    }))
    .filter((p) => p.parsedStart)
    .sort((a, b) => a.parsedStart - b.parsedStart);

  const processedDown = downAlarms
    .map((d) => ({
      ...d,
      parsedStart: d.startTime instanceof Date ? d.startTime : excelToDate(d.startTime),
    }))
    .filter((d) => d.parsedStart)
    .sort((a, b) => a.parsedStart - b.parsedStart);

  const calculatedBackups = [];

  processedPower.forEach((p) => {
    // Avoid measuring from "Low Voltage" precursors if possible
    const name = (p.alarmName || "").toLowerCase();
    if (name.includes("low") || name.includes("voltage") || name.includes("vlt")) return;

    const matchedDown = processedDown.find((d) => {
      // Site went down AFTER power cut, but BEFORE power restored
      return d.parsedStart >= p.parsedStart && d.parsedStart <= p.parsedEnd;
    });

    if (matchedDown) {
      const backup = (matchedDown.parsedStart - p.parsedStart) / (1000 * 60);
      // Filter out noise / false matches (less than 1 minute)
      if (backup > 1) {
        calculatedBackups.push(backup);
      }
    }
  });

  if (calculatedBackups.length === 0) return null;
  return calculatedBackups.reduce((a, b) => a + b, 0) / calculatedBackups.length;
}

export function analyzeAlarms(alarms) {
  const result = {};

  for (const cat of CATEGORIES) {
    const catAlarms = alarms.filter((a) => a.categories?.includes(cat.id));
    const siteAgg = {};
    const regions = {};

    for (const alarm of catAlarms) {
      const code = alarm.siteCode;
      const region = alarm.region || "Unknown";

      if (!siteAgg[code]) {
        siteAgg[code] = {
          siteCode: code,
          siteName: alarm.siteName || code,
          region: region,
          bdt: alarm.bdt || 0,
          count: 0,
          maxDuration: 0,
        };
      }
      siteAgg[code].count += 1;
      const dur = alarm.durationMinutes || 0;
      if (dur > siteAgg[code].maxDuration) siteAgg[code].maxDuration = dur;

      regions[region] = (regions[region] || 0) + 1;
    }

    const sites = Object.values(siteAgg);
    const topOffenders = [...sites]
      .sort((a, b) => b.count - a.count)
      .slice(0, 15);

    result[cat.id] = {
      total: catAlarms.length,
      uniqueSites: sites.length,
      topOffenders,
      regions,
    };
  }

  return result;
}

/**
 * BBI Analysis: Battery Backup Indicator
 * Logic for "Backup mismatched"
 */
export function analyzeBBI(alarms, siteList) {
  const clean = (str) =>
    str
      ?.toString()
      .replace(/[^a-zA-Z0-9]/g, "")
      .toUpperCase() || "";

  // 1. Group alarms by site
  const siteAlarms = {};
  alarms.forEach((a) => {
    const code = clean(a.siteCode);
    if (!code) return;
    if (!siteAlarms[code]) siteAlarms[code] = [];
    siteAlarms[code].push(a);
  });
  const siteMap = new Map();
  siteList.forEach((s) => {
    const code = clean(s.shortCode);
    if (code) siteMap.set(code, s);
  });

  // 3. Process each site
  const bbiResults = [];

  for (const [code, siteAlarmsList] of Object.entries(siteAlarms)) {
    const siteInfo = siteMap.get(code);

    // Condition: BDT must be present and >= 60
    if (!siteInfo || siteInfo.bdt < 60) continue;

    // Filter alarms >= 10 min (Sensitive audit)
    const powerAlarms = siteAlarmsList.filter(
      (a) => a.category === "power" && a.durationMinutes >= 10,
    );
    const downAlarms = siteAlarmsList.filter(
      (a) => a.category === "down" && a.durationMinutes >= 10,
    );

    // Condition: Site MUST have at least one down alarm to be analyzed for BBI mismatch
    if (downAlarms.length === 0) continue;

    const avgBackup = calculateSiteBackup(siteAlarmsList);

    // Aggregations
    if (!avgBackup) continue; // Skip sites with no valid measured discharge cycles

    const totalPowerCount = powerAlarms.length;
    const totalPowerDuration = powerAlarms.reduce(
      (a, b) => a + b.durationMinutes,
      0,
    );
    const avgPowerDuration =
      totalPowerCount > 0 ? totalPowerDuration / totalPowerCount : 0;

    const totalDownCount = downAlarms.length;
    const totalDownDuration = downAlarms.reduce(
      (a, b) => a + b.durationMinutes,
      0,
    );
    const avgDownDuration =
      totalDownCount > 0 ? totalDownDuration / totalDownCount : 0;

    const tolerance = siteInfo.bdt - avgBackup;
    let autoComment = "";

    // PLD Abnormal Trigger Logic
    if (siteInfo.pld) {
      if (tolerance > 60) {
        autoComment = "PLD value seems very high";
      } else if (tolerance < -60) {
        autoComment = "PLD value seems very low";
      }
    }

    bbiResults.push({
      shortCode: siteInfo.shortCode,
      siteName: siteInfo.siteName,
      scOffice: siteInfo.scOffice,
      nodalDeg: siteInfo.nodalDeg,
      vip: siteInfo.vip,
      pld: siteInfo.pld,
      powerAlarmsCount: totalPowerCount,
      downAlarmsCount: totalDownCount,
      avgPowerDuration,
      avgDownDuration,
      totalDownDuration,
      avgBackup,
      bdt: siteInfo.bdt,
      batteries: siteInfo.batteries,
      tolerance,
      comment: autoComment,
    });
  }

  // Categories for the UI
  const backupMismatched = bbiResults
    .filter((s) => s.pld === false && s.tolerance > 30) // Strictly must be false
    .sort((a, b) => b.tolerance - a.tolerance);

  const pldAbnormal = bbiResults
    .filter((s) => s.pld === true && Math.abs(s.tolerance) > 60) // Strictly must be true
    .sort((a, b) => b.tolerance - a.tolerance);

  return {
    all: bbiResults.sort((a, b) => b.tolerance - a.tolerance),
    backupMismatched,
    pldAbnormal,
  };
}

/**
 * Get a summary overview across all categories.
 */
export function getAnalysisSummary(analysis) {
  return CATEGORIES.map((cat) => ({
    ...cat,
    total: analysis[cat.id]?.total || 0,
    uniqueSites: analysis[cat.id]?.uniqueSites || 0,
  }));
}
