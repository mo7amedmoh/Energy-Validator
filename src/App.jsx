import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import {
  BarChart3,
  Settings,
  AlertCircle,
  CheckCircle2,
  Activity,
  Zap,
  TrendingUp,
  Download,
  ArrowRight,
  RefreshCw,
  Plus,
  ArrowUpRight,
  ArrowDownRight,
  Battery,
  MessageSquare,
  LayoutDashboard,
  Filter,
  Save,
  RotateCcw,
  ExternalLink,
  ChevronRight,
  Clock,
  ShieldCheck,
  AlertTriangle,
  FileSpreadsheet,
  Search,
  Calendar,
  ArrowLeft,
  Camera,
  ChevronDown,
  ClipboardCheck,
  Database,
  DoorOpen,
  History,
  Layers,
  RadioTower,
  XCircle,
  Eye,
  EyeOff,
  Moon,
  Sun,
  Palette,
} from "lucide-react";
import {
  PieChart,
  Pie,
  Cell,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  Legend,
} from "recharts";
import { parseAlarmsExcel, formatDuration } from "./lib/excel";
import { fetchSiteList } from "./lib/siteList";
import {
  analyzeAlarms,
  CATEGORIES,
  analyzeBBI,
  calculateSiteBackup,
} from "./lib/analysis";
import { parseBDTFile } from "./lib/bdt";
import BDTPhotoInspector from "./components/BDTPhotoInspector";
import { correlateBDTWithAlarms } from "./lib/bdtCorrelation";
import TrialActivation from "./components/TrialActivation";

const Card = ({ children, className = "" }) => (
  <div
    className={`bg-white dark:bg-premium-900 border border-premium-200 dark:border-premium-800 rounded-[2rem] shadow-sm transition-all duration-300 hover:shadow-xl hover:border-premium-300 dark:hover:border-premium-700 ${className}`}
  >
    {children}
  </div>
);

const getShortIssueText = (issue) => {
  const text = issue.text || "";
  const section = issue.section || "";
  const tLower = text.toLowerCase();

  if (section === "Summary Sync") {
    const match = text.match(/^(.*?)\s+mismatch:/i);
    if (match) return `${match[1]} mismatched in summary`;
    return "Data mismatched in summary";
  }

  if (tLower.includes("photos detected") || tLower.includes("photo count"))
    return "Missing Photos";
  if (
    tLower.includes("incomplete data") ||
    tLower.includes("basic data sparse")
  )
    return "Missing Basic data";
  if (tLower.includes("starting voltages"))
    return "Battery under float voltage";
  if (tLower.includes("rectifier start ampere")) return "Wrong Start ampere";
  if (tLower.includes("batteries start ampere sum"))
    return "Batteries seems not charged";
  if (tLower.includes("voltage") && tLower.includes("increased"))
    return "Volt increase over time";
  if (tLower.includes("ampere") && tLower.includes("decreased"))
    return "Ampere decrease over time";
  if (tLower.includes("sum") && tLower.includes("exceeds rectifier"))
    return "Wrong ampere readings";
  if (tLower.includes("deviates from theoretical"))
    return "Backup is less than expected";
  if (tLower.includes("end voltage") && tLower.includes("too high"))
    return "Incomplete BDT";
  if (tLower.includes("readings interval should be"))
    return "Wrong BDT intervals";
  if (tLower.includes("cells sum") && tLower.includes("vs busbar"))
    return "Wrong cells voltage readings";
  if (tLower.includes("cells section is empty"))
    return "Missing cells voltage readings";
  if (
    tLower.includes("no discharge interval readings") ||
    tLower.includes("does not match cells section")
  )
    return "battery Cells readings mismatched with BDT duration";

  return text.replace(/Rule \d+:/g, "").trim();
};

const App = () => {
  const [siteList, setSiteList] = useState([]);

  // Environment detection: Bypass locks in Electron, enable in Web browser
  const isElectron = typeof window !== 'undefined' && window.process && window.process.versions && window.process.versions.electron;

  const [trialLockEnabled, setTrialLockEnabled] = useState(() => {
    if (isElectron) return false;
    const saved = localStorage.getItem("energy_review_trial_lock_enabled");
    return saved !== null ? JSON.parse(saved) : true;
  });

  const [isActivated, setIsActivated] = useState(() => {
    const saved = localStorage.getItem("energy_review_activated");
    return saved !== null ? JSON.parse(saved) : false;
  });
  const [siteListLoading, setSiteListLoading] = useState(true);
  const [siteListError, setSiteListError] = useState(null);
  const [lastSyncTime, setLastSyncTime] = useState(null);
  const [history, setHistory] = useState([]);
  const [loading, setLoading] = useState(false);
  const [dragActive, setDragActive] = useState(false);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [bbiData, setBbiData] = useState(null);
  const [activeBbiTab, setActiveBbiTab] = useState("mismatched");
  const [dataSourceMode, setDataSourceMode] = useState("energy"); // 'energy' or 'autin'
  const [autinFiles, setAutinFiles] = useState({
    power: [],
    down: [],
    high_temp: [],
    generator: [],
    door: [],
  });
  const [currentAlarms, setCurrentAlarms] = useState([]);
  const [searchSiteCode, setSearchSiteCode] = useState("");
  const [selectedSite, setSelectedSite] = useState(null);
  const [dateFilter, setDateFilter] = useState({ start: "", end: "" });
  const [spotlightCategory, setSpotlightCategory] = useState(null);
  const [dashboardCategory, setDashboardCategory] = useState("high_temp");

  // Theme state
  const [nightMode, setNightMode] = useState(() => {
    const saved = localStorage.getItem("energy_review_night_mode");
    return saved !== null ? JSON.parse(saved) : true;
  });
  const [appTheme, setAppTheme] = useState(() => {
    return localStorage.getItem("energy_review_theme") || "default";
  });

  // Apply themes
  useEffect(() => {
    document.documentElement.classList.toggle("dark", nightMode);
    localStorage.setItem("energy_review_night_mode", JSON.stringify(nightMode));
  }, [nightMode]);

  useEffect(() => {
    document.documentElement.setAttribute("data-theme", appTheme);
    localStorage.setItem("energy_review_theme", appTheme);
  }, [appTheme]);

  // BDT Validator state
  const [bdtResults, setBdtResults] = useState([]);
  const [bdtSummary, setBdtSummary] = useState(null);
  const [oldSummary, setOldSummary] = useState(null);
  const [bdtStatusFilter, setBdtStatusFilter] = useState(null); // null | 'pass' | 'warning' | 'fail'
  const [bdtLoading, setBdtLoading] = useState(false);
  const [bdtDragActive, setBdtDragActive] = useState(false);
  const [bdtSummaryDragActive, setBdtSummaryDragActive] = useState(false);
  const [oldSummaryDragActive, setOldSummaryDragActive] = useState(false);
  const [expandedBdtSheet, setExpandedBdtSheet] = useState(null);
  const [reportLoading, setReportLoading] = useState(false);
  const [reviewedBdtSheets, setReviewedBdtSheets] = useState(new Set());

  // Persistence: Alarm Memory
  const [alarmDatabases, setAlarmDatabases] = useState(() => {
    try {
      const saved = localStorage.getItem("alarm_databases");
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  // Load alarms and history from storage
  useEffect(() => {
    try {
      const savedAlarms = localStorage.getItem("current_alarms");
      if (savedAlarms) setCurrentAlarms(JSON.parse(savedAlarms));

      const savedHistory = localStorage.getItem("analysis_history");
      if (savedHistory) setHistory(JSON.parse(savedHistory));

      const savedBbi = localStorage.getItem("bbi_data");
      if (savedBbi) setBbiData(JSON.parse(savedBbi));
    } catch (e) {
      console.warn("Could not load from storage", e);
    }
  }, []);

  // Sync to storage
  useEffect(() => {
    try {
      localStorage.setItem("alarm_databases", JSON.stringify(alarmDatabases));
      localStorage.setItem("current_alarms", JSON.stringify(currentAlarms));
      localStorage.setItem("analysis_history", JSON.stringify(history));
      localStorage.setItem("bbi_data", JSON.stringify(bbiData));
    } catch (e) {
      if (e.name === "QuotaExceededError") {
        console.warn("Storage quota exceeded. Some data may not be saved.");
      }
    }
  }, [alarmDatabases, currentAlarms, history, bbiData]);

  const DEFAULT_CONFIG = {
    spreadsheetId: "1XwQTSMI5Nz0WuKwSow06dVmTlnoyZOuj",
    sheetName: "SiteList",
    mapping: {
      shortCode: 0,
      siteName: 1,
      scOffice: 2,
      region: 3,
      nodalDeg: 4,
      pld: 18,
      batteries: 20,
      bdt: 27,
      vip: 33,
    },
  };

  const DEFAULT_BDT_CONFIG = {
    photo_min_pass: 12,
    photo_min_warn: 6,
    photo_min_fail: 8,
    basic_data_min_cells: 28,
    rule4_min_volt_48v: 44,
    rule7_interval_mins: 10,
    rule8_duration_mismatch_tolerance: 5,
    rule6_12v_dissipated_max: 1.5,
    busbar_min_fill_rate: 0.15,
    rule1_batt_amp_max: 1.0,
    rule2_seq_tolerance: 0.05,
    rule3_balance_max: 3.0,
    rule1_min_volt_start: 51.0,
    rule1_min_rect_amp: 20.0,
    rule4_theoretical_tolerance: 20,
  };

  const [sysConfig, setSysConfig] = useState(() => {
    try {
      const saved = localStorage.getItem("energy_review_config");
      return saved ? JSON.parse(saved) : DEFAULT_CONFIG;
    } catch {
      return DEFAULT_CONFIG;
    }
  });

  const loadSiteList = async () => {
    setSiteListLoading(true);
    setSiteListError(null);
    try {
      console.time("SiteList Fetch");
      const sites = await fetchSiteList(sysConfig);
      console.timeEnd("SiteList Fetch");
      setSiteList(sites);
      setLastSyncTime(new Date());
    } catch (err) {
      console.error("Failed to fetch SiteList:", err);
      setSiteListError(err.message);
    } finally {
      setSiteListLoading(false);
    }
  };

  const [bdtConfig, setBdtConfig] = useState(() => {
    try {
      const saved = localStorage.getItem("bdt_review_config");
      return saved ? JSON.parse(saved) : DEFAULT_BDT_CONFIG;
    } catch {
      return DEFAULT_BDT_CONFIG;
    }
  });

  const [bdtConfigSaved, setBdtConfigSaved] = useState(false);

  // Auto-persist BDT config to localStorage on change
  useEffect(() => {
    try {
      localStorage.setItem("bdt_review_config", JSON.stringify(bdtConfig));
    } catch (e) {
      console.warn("Failed to persist BDT config", e);
    }
  }, [bdtConfig]);

  useEffect(() => {
    loadSiteList();
  }, []);

  const correlateWithSummary = (results, summary) => {
    if (!summary || !summary.rows) return results;

    return results.map((fileResult) => {
      let fileHasFailure = false;
      const updatedSheets = fileResult.sheets.map((sheet) => {
        // Search from the bottom of the summary sheet to find the most recent record
        const summaryRow = [...summary.rows].reverse().find((row) => {
          const shortCodeKey = Object.keys(row).find((k) => {
            const kl = k.toLowerCase();
            return (
              (kl.includes("short code") ||
                kl.includes("site code") ||
                kl.includes("site identity")) &&
              !kl.includes("name") &&
              !kl.includes("address")
            );
          });
          if (!shortCodeKey || !sheet.summaryData?.siteCode) return false;

          const sVal = clean(row[shortCodeKey]).replace(/^0+/, "");
          const bdtVal = clean(sheet.summaryData.siteCode).replace(/^0+/, "");
          if (!sVal || !bdtVal) return false;

          return sVal === bdtVal;
        });

        if (!summaryRow) return sheet;

        const summaryIssues = [];
        const mapping = [
          {
            key: "batteryVolt",
            label: "Battery Volt",
            patterns: ["Battery Volt", "Volt"],
          },
          {
            key: "batteryAH",
            label: "Battery Ampere Hour",
            patterns: ["Battery Ampere Hour", "Ampere", " AH"],
          },
          {
            key: "numStrings",
            label: "No of String",
            patterns: ["No of String", "String"],
          },
          {
            key: "numBatteries",
            label: "No of Batteries",
            patterns: ["No of Batteries", "Batteri"],
          },
          { key: "startVolt", label: "Start Volt", patterns: ["Start Volt"] },
          { key: "startAmp", label: "Start Amp", patterns: ["Start Amp"] },
          { key: "endVolt", label: "End Volt", patterns: ["End Volt"] },
          { key: "endAmp", label: "End Amp", patterns: ["End Amp"] },
          {
            key: "numModules",
            label: "# of Modules",
            patterns: ["# of Modules", "Modules"],
          },
          {
            key: "dischargeTime",
            label: "Discharge time( Mins)",
            patterns: ["Discharge time", "Duration"],
          },
        ];

        const bdtBackupMins =
          parseFloat(
            String(sheet.summaryData?.dischargeTime || "0").replace(
              /[^\d.]/g,
              "",
            ),
          ) || 0;

        mapping.forEach((m) => {
          // Skip end volt and end amp checks if backup time is 0
          if (
            (m.key === "endVolt" || m.key === "endAmp") &&
            bdtBackupMins === 0
          )
            return;

          const summaryKey = Object.keys(summaryRow).find((k) =>
            m.patterns.some((p) => k.toLowerCase().includes(p.toLowerCase())),
          );
          if (summaryKey) {
            let summaryVal = summaryRow[summaryKey];
            let bdtVal = String(sheet.summaryData?.[m.key] || "").trim();

            // Handle Excel serial dates or JS Date objects
            const isDateCol =
              summaryKey.toLowerCase().includes("date") &&
              !summaryKey.toLowerCase().includes("pm");
            const isTimeCol =
              summaryKey.toLowerCase().includes("time") &&
              !summaryKey.toLowerCase().includes("discharge");

            if (
              (isDateCol || isTimeCol) &&
              (summaryVal instanceof Date ||
                (typeof summaryVal === "number" && summaryVal > 40000))
            ) {
              try {
                let date;
                if (summaryVal instanceof Date) {
                  date = summaryVal;
                } else {
                  // Excel dates are days since 1900-01-01
                  date = new Date(Math.round((summaryVal - 25569) * 864e5));
                }
                summaryVal = date.toLocaleDateString("en-GB", {
                  day: "2-digit",
                  month: "short",
                  year: "2-digit",
                });
              } catch (_e) {
                summaryVal = String(summaryVal);
              }
            } else {
              summaryVal = String(summaryVal || "").trim();
            }

            if (!summaryVal || !bdtVal) return;

            let isMatched = false;

            // Numeric comparison with tolerance
            const sNum = parseFloat(summaryVal.replace(/[^\d.]/g, ""));
            const bNum = parseFloat(bdtVal.replace(/[^\d.]/g, ""));

            if (!isNaN(sNum) && !isNaN(bNum)) {
              isMatched = Math.abs(sNum - bNum) < 0.1;
            } else {
              const cleanMatch = clean(summaryVal) === clean(bdtVal);
              const partialMatch =
                summaryVal.toLowerCase().includes(bdtVal.toLowerCase()) ||
                bdtVal.toLowerCase().includes(summaryVal.toLowerCase());
              isMatched = cleanMatch || partialMatch;
            }

            if (!isMatched) {
              summaryIssues.push({
                section: "Summary Sync",
                text: `${m.label} mismatch: BDT says "${bdtVal}" vs Summary says "${summaryVal}" (Col: "${summaryKey}")`,
              });
            }
          }
        });

        if (summaryIssues.length > 0) {
          fileHasFailure = true;
          return {
            ...sheet,
            allIssues: [
              ...sheet.allIssues.filter(
                (iss) => iss.section !== "Summary Sync",
              ),
              ...summaryIssues,
            ],
            overallStatus: "fail",
          };
        }

        return sheet;
      });

      return {
        ...fileResult,
        sheets: updatedSheets,
        overallStatus: fileHasFailure ? "fail" : fileResult.overallStatus,
      };
    });
  };

  const correlateWithOldSummary = (results, oldSummaryData) => {
    if (!oldSummaryData || !oldSummaryData.rows) return results;

    return results.map((fileResult) => {
      let fileHasFailure = false;
      const updatedSheets = fileResult.sheets.map((sheet) => {
        const oldSummaryRow = [...oldSummaryData.rows].reverse().find((row) => {
          const shortCodeKey = Object.keys(row).find((k) => {
            const kl = k.toLowerCase();
            return (
              (kl.includes("short code") ||
                kl.includes("site code") ||
                kl.includes("site identity")) &&
              !kl.includes("name") &&
              !kl.includes("address")
            );
          });
          if (!shortCodeKey || !sheet.summaryData?.siteCode) return false;

          const sVal = clean(row[shortCodeKey]).replace(/^0+/, "");
          const bdtVal = clean(sheet.summaryData.siteCode).replace(/^0+/, "");
          return sVal && bdtVal && sVal === bdtVal;
        });

        if (!oldSummaryRow) return sheet;

        const oldModuleKey = Object.keys(oldSummaryRow).find((k) => {
          const kl = k.toLowerCase();
          return kl.includes("# of modules") || kl.includes("modules");
        });

        if (!oldModuleKey) return sheet;

        const oldModules = parseInt(
          String(oldSummaryRow[oldModuleKey]).replace(/[^\d]/g, ""),
          10,
        );
        const currentModules = parseInt(
          String(sheet.summaryData?.numModules || "0").replace(/[^\d]/g, ""),
          10,
        );

        if (
          !isNaN(oldModules) &&
          !isNaN(currentModules) &&
          currentModules < oldModules
        ) {
          const newSheet = { ...sheet };
          if (!newSheet.allIssues) newSheet.allIssues = [];

          newSheet.allIssues.push({
            section: "Assets",
            text: `Assets loss found (Old Modules: ${oldModules}, Current: ${currentModules}, Diff: ${oldModules - currentModules})`,
          });
          newSheet.overallStatus = "fail";
          fileHasFailure = true;
          return newSheet;
        }

        return sheet;
      });

      return {
        ...fileResult,
        sheets: updatedSheets,
        overallStatus: fileHasFailure
          ? "fail"
          : fileResult.overallStatus === "fail"
            ? "fail"
            : fileResult.overallStatus,
      };
    });
  };

  useEffect(() => {
    if (bdtResults.length > 0 && bdtSummary) {
      setBdtResults((prev) => correlateWithSummary(prev, bdtSummary));
    }
  }, [bdtSummary]);

  useEffect(() => {
    if (bdtResults.length > 0 && oldSummary) {
      setBdtResults((prev) => correlateWithOldSummary(prev, oldSummary));
    }
  }, [oldSummary]);

  useEffect(() => {
    if (bdtResults.length > 0 && currentAlarms.length > 0) {
      setBdtResults((prev) => correlateBDTWithAlarms(prev, currentAlarms));
    }
  }, [currentAlarms]);

  const handleManualOverride = (fileIdx, sheetIdx, overrideDecision) => {
    setBdtResults((prev) => {
      const newResults = [...prev];
      const newSheets = [...newResults[fileIdx].sheets];
      newSheets[sheetIdx] = { ...newSheets[sheetIdx] };
      newSheets[sheetIdx].manualOverride =
        newSheets[sheetIdx].manualOverride === overrideDecision
          ? null
          : overrideDecision;
      newResults[fileIdx] = { ...newResults[fileIdx], sheets: newSheets };
      return newResults;
    });
  };

  const handleMemoryAnalysis = (dbId = null) => {
    // Determine target alarms: either a specific database or all stored databases
    const targetAlarms = dbId
      ? currentAlarms.filter((a) => a.dbId === dbId)
      : currentAlarms;

    if (targetAlarms.length === 0) {
      alert("No data found for analysis.");
      return;
    }
    setLoading(true);
    setTimeout(() => {
      try {
        const bbiResults = analyzeBBI(targetAlarms, siteList);
        const results = analyzeAlarms(targetAlarms);

        const dbInfo = dbId ? alarmDatabases.find((d) => d.id === dbId) : null;
        const fileName = dbInfo
          ? `Stored: ${dbInfo.fileName}`
          : "Stored Database Analysis";

        setBbiData(bbiResults);
        setHistory((prev) => [
          {
            timestamp: new Date().toISOString(),
            siteCount: siteList.length,
            results,
            fileName,
          },
          ...prev.slice(0, 4),
        ]);
        setLoading(false);
        setActiveTab("dashboard");
      } catch (err) {
        console.error("Memory Analysis Error:", err);
        alert("Analysis failed: " + err.message);
        setLoading(false);
      }
    }, 500);
  };

  const handleFileUpload = async (file) => {
    if (!file) return;
    if (siteListLoading) {
      alert("Database is syncing. Please wait...");
      return;
    }

    setLoading(true);

    // Safety Force Stop: If it takes more than 20s, something is wrong
    const forceStopTimer = setTimeout(() => {
      setLoading(false);
    }, 20000);

    // Use timeout to allow UI to show loading state
    setTimeout(async () => {
      try {
        console.time("Analysis_Phase");

        console.log("[Engine] Starting Parse...");
        const parsedAlarms = await parseAlarmsExcel(file);

        const clean = (str) =>
          str
            ?.toString()
            .replace(/[^a-zA-Z0-9]/g, "")
            .toUpperCase() || "";

        console.log("[Engine] Mapping Sites...");
        const siteMap = new Map();
        siteList.forEach((site) => {
          if (site.shortCode) siteMap.set(clean(site.shortCode), site);
        });

        const mappedAlarms = [];
        for (const alarm of parsedAlarms) {
          const siteInfo = siteMap.get(clean(alarm.siteCode));
          if (siteInfo) mappedAlarms.push({ ...alarm, ...siteInfo });
        }

        if (mappedAlarms.length === 0) {
          throw new Error("No sites from your Database found in this file.");
        }

        console.log("[Engine] Running Algorithms...");
        const bbiResults = analyzeBBI(
          [...currentAlarms, ...mappedAlarms],
          siteList,
        );
        const results = analyzeAlarms([...currentAlarms, ...mappedAlarms]);

        const dbId = Date.now().toString();
        const newDbEntry = {
          id: dbId,
          fileName: file.name,
          timestamp: new Date().toISOString(),
          alarmCount: mappedAlarms.length,
          type: "NET Energizer",
          source: "file",
        };

        // 1. UPDATE UI IMMEDIATELY
        setAlarmDatabases((prev) => [...prev, newDbEntry]);
        setCurrentAlarms((prev) => [
          ...prev,
          ...mappedAlarms.map((a) => ({ ...a, dbId })),
        ]);
        setBbiData(bbiResults);
        setHistory((prev) => [
          {
            timestamp: new Date().toISOString(),
            siteCount: siteList.length,
            results,
            fileName: file.name,
          },
          ...prev.slice(0, 4), // Keep last 5 history records
        ]);

        setLoading(false);
        setActiveTab("dashboard");
        clearTimeout(forceStopTimer);
        console.timeEnd("Analysis_Phase");
      } catch (err) {
        console.error("Analysis Error:", err);
        alert(err.message);
        setLoading(false);
        clearTimeout(forceStopTimer);
      }
    }, 100);
  };
  const handleAutinAnalysis = async () => {
    setLoading(true);
    const forceStopTimer = setTimeout(() => setLoading(false), 30000);

    setTimeout(async () => {
      try {
        console.time("AUTIN_Analysis");
        let allParsedAlarms = [];

        // Parse each category
        for (const [cat, files] of Object.entries(autinFiles)) {
          for (const file of files) {
            console.log(`[AUTIN] Parsing ${cat}: ${file.name}`);
            const data = await parseAlarmsExcel(file, cat);
            for (const item of data) {
              allParsedAlarms.push(item);
            }
          }
        }

        if (allParsedAlarms.length === 0) {
          throw new Error("No data provided in any of the categories.");
        }

        const clean = (str) =>
          str
            ?.toString()
            .replace(/[^a-zA-Z0-9]/g, "")
            .toUpperCase() || "";
        const siteMap = new Map();
        siteList.forEach((site) => {
          if (site.shortCode) siteMap.set(clean(site.shortCode), site);
        });

        const mappedAlarms = [];
        for (const alarm of allParsedAlarms) {
          const siteInfo = siteMap.get(clean(alarm.siteCode));
          if (siteInfo) mappedAlarms.push({ ...alarm, ...siteInfo });
        }

        if (mappedAlarms.length === 0) {
          throw new Error("No matching sites found in the uploaded files.");
        }

        const bbiResults = analyzeBBI(
          [...currentAlarms, ...mappedAlarms],
          siteList,
        );
        const results = analyzeAlarms([...currentAlarms, ...mappedAlarms]);

        const dbId = Date.now().toString();
        const newDbEntry = {
          id: dbId,
          fileName: "AUTIN Multi-Source",
          timestamp: new Date().toISOString(),
          alarmCount: mappedAlarms.length,
          type: "Autin Audit",
          source: "multi",
        };

        setAlarmDatabases((prev) => [...prev, newDbEntry]);
        setCurrentAlarms((prev) => [
          ...prev,
          ...mappedAlarms.map((a) => ({ ...a, dbId })),
        ]);
        setBbiData(bbiResults);
        setHistory((prev) => [
          {
            timestamp: new Date().toISOString(),
            siteCount: siteList.length,
            results,
            fileName: "AUTIN_Multi_Source_Audit",
          },
          ...prev.slice(0, 4),
        ]);

        setLoading(false);
        setActiveTab("dashboard");
        clearTimeout(forceStopTimer);
        console.timeEnd("AUTIN_Analysis");
      } catch (err) {
        console.error("AUTIN Error:", err);
        alert(err.message);
        setLoading(false);
        clearTimeout(forceStopTimer);
      }
    }, 100);
  };

  const handleBdtUpload = async (files) => {
    if (!files || files.length === 0) return;
    setBdtLoading(true);
    try {
      const fileArray = Array.from(files);
      const allResults = [];
      for (const file of fileArray) {
        const result = await parseBDTFile(file, bdtConfig);
        allResults.push(result);
      }
      const correlated = correlateBDTWithAlarms(allResults, currentAlarms);
      const summaryCorrelated = bdtSummary
        ? correlateWithSummary(correlated, bdtSummary)
        : correlated;

      const finalResults = summaryCorrelated.map((fileResult) => ({
        ...fileResult,
        sheets: fileResult.sheets.map((sheet) => {
          const newSheet = { ...sheet };
          let downgraded = false;

          if (!newSheet.allIssues) newSheet.allIssues = [];

          if (!bdtSummary) {
            newSheet.allIssues = [
              ...newSheet.allIssues,
              { section: "Summary", text: "Summary not reviewed" },
            ];
            downgraded = true;
          }

          if (!currentAlarms || currentAlarms.length === 0) {
            newSheet.allIssues = [
              ...newSheet.allIssues,
              { section: "Alarms", text: "Alarms not reviewed" },
            ];
            downgraded = true;
          }

          if (downgraded && newSheet.overallStatus === "pass") {
            newSheet.overallStatus = "warning";
          }

          return newSheet;
        }),
      }));

      // Update the file overallStatus too
      finalResults.forEach((file) => {
        if (file.sheets.some((s) => s.overallStatus === "fail"))
          file.overallStatus = "fail";
        else if (file.sheets.some((s) => s.overallStatus === "warning"))
          file.overallStatus = "warning";
      });

      const oldSummaryCorrelated = oldSummary
        ? correlateWithOldSummary(finalResults, oldSummary)
        : finalResults;

      setBdtResults((prev) => [...prev, ...oldSummaryCorrelated]);
    } catch (err) {
      console.error("BDT Error:", err);
      alert("BDT Parse Error: " + err.message);
    } finally {
      setBdtLoading(false);
    }
  };

  const handleBdtSummaryUpload = async (files) => {
    if (!files || files.length === 0) return;
    const file = files[0];
    setBdtLoading(true);
    try {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array", cellDates: true });
          const sheetName =
            workbook.SheetNames.find(
              (n) => !n.toLowerCase().includes("dc count"),
            ) || workbook.SheetNames[0];
          const ws = workbook.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws);

          setBdtSummary({
            fileName: file.name,
            sheetName,
            rowCount: rows.length,
            rows: rows,
          });
          setBdtLoading(false);
        } catch (innerErr) {
          console.error("Summary Parse Error:", innerErr);
          alert("Error parsing summary: " + innerErr.message);
          setBdtLoading(false);
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      console.error("Summary Upload Error:", err);
      alert("Summary Upload Error: " + err.message);
      setBdtLoading(false);
    }
  };

  const handleOldSummaryUpload = async (files) => {
    if (!files || files.length === 0) return;
    const file = files[0];
    setBdtLoading(true);
    try {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array", cellDates: true });
          const sheetName =
            workbook.SheetNames.find(
              (n) => !n.toLowerCase().includes("dc count"),
            ) || workbook.SheetNames[0];
          const ws = workbook.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws);

          setOldSummary({
            fileName: file.name,
            sheetName,
            rowCount: rows.length,
            rows: rows,
          });
          setBdtLoading(false);
        } catch (innerErr) {
          console.error("Old Summary Parse Error:", innerErr);
          alert("Error parsing old summary: " + innerErr.message);
          setBdtLoading(false);
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      console.error("Old Summary Upload Error:", err);
      alert("Old Summary Upload Error: " + err.message);
      setBdtLoading(false);
    }
  };

  const handleGenerateReport = async (acceptanceFile) => {
    if (!acceptanceFile || bdtResults.length === 0) return;
    setReportLoading(true);
    try {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(worksheet);

          // Map BDT results for quick lookup
          const bdtMap = new Map();
          bdtResults.forEach((fileResult) => {
            fileResult.sheets.forEach((sheet) => {
              if (sheet.siteName) {
                bdtMap.set(clean(sheet.siteName), sheet);
              }
            });
          });

          const reportData = rows.map((row) => {
            // Helper to find column by flexible names
            const findVal = (patterns) => {
              const key = Object.keys(row).find((k) =>
                patterns.some((p) => k.toLowerCase().includes(p.toLowerCase())),
              );
              return key ? row[key] : "N/A";
            };

            const formatDateHelper = (val) => {
              if (!val || val === "N/A") return val;
              if (val instanceof Date) {
                return val.toLocaleDateString("en-GB", {
                  day: "2-digit",
                  month: "short",
                  year: "numeric",
                });
              }
              // Excel serial date (number of days since 1900-01-01)
              if (typeof val === "number" && val > 20000 && val < 70000) {
                try {
                  const date = new Date(Math.round((val - 25569) * 864e5));
                  return date.toLocaleDateString("en-GB", {
                    day: "2-digit",
                    month: "short",
                    year: "numeric",
                  });
                } catch (e) {
                  return val;
                }
              }
              return val;
            };

            const siteCode = findVal(["Site Code", "Code", "SiteID"]);
            const siteName = findVal(["Site Name", "Name"]);
            const zone = findVal(["Zone", "Area", "Region"]);
            const pmDate = formatDateHelper(findVal(["PM Date"]));
            const bdtDate = formatDateHelper(findVal(["BDT Date"]));
            const weekNumber = findVal(["Week"]);

            const bdtInfo =
              bdtMap.get(clean(siteCode)) || bdtMap.get(clean(siteName));

            let status = "Not Found";
            let comments = "No BDT File";

            if (bdtInfo) {
              const baseStatusStr =
                bdtInfo.overallStatus === "fail" ? "Rejected" : "Accepted";
              if (bdtInfo.manualOverride) {
                status =
                  bdtInfo.manualOverride === "pass"
                    ? "Accepted (Manual)"
                    : "Rejected (Manual)";
              } else {
                status = baseStatusStr;
              }

              if (status.includes("Accepted") || status === "Warning") {
                comments = "";
              } else {
                const mappedIssues = bdtInfo.allIssues.map(getShortIssueText);
                comments = Array.from(new Set(mappedIssues)).join(", ");
              }
            }

            return {
              "Site Code": siteCode,
              "Site Name": siteName,
              Zone: zone,
              "PM Date": pmDate,
              "BDT Date": bdtDate,
              "Week Number": weekNumber,
              "BDT Status": status,
              Comments: comments,
            };
          });

          const newWb = new ExcelJS.Workbook();
          const newWs = newWb.addWorksheet("BDT_Report");

          if (reportData.length > 0) {
            const headers = Object.keys(reportData[0]);

            newWs.addTable({
              name: "BDTReportTable",
              ref: "A1",
              headerRow: true,
              style: {
                theme: "TableStyleMedium9",
                showRowStripes: true,
              },
              columns: headers.map((h) => ({ name: h, filterButton: true })),
              rows: reportData.map((dataRow) => headers.map((h) => dataRow[h])),
            });

            const headerRow = newWs.getRow(1);
            headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
            headerRow.alignment = { vertical: "middle", horizontal: "center" };

            newWs.columns.forEach((column, index) => {
              column.width = 20;
            });
            newWs.getColumn(headers.indexOf("Comments") + 1).width = 45;
            newWs.getColumn(headers.indexOf("Site Name") + 1).width = 25;
            newWs.getColumn(headers.indexOf("Site Code") + 1).width = 15;

            newWs.eachRow((row) => {
              row.eachCell((cell) => {
                if (row.number > 1) {
                  cell.alignment = { vertical: "middle", wrapText: true };
                }
              });
            });
          }

          // Add Analysis Sheet
          const analysisWs = newWb.addWorksheet("Analysis");

          let acceptedCount = 0;
          let rejectedCount = 0;

          reportData.forEach((row) => {
            if (row["BDT Status"].includes("Accepted")) acceptedCount++;
            if (row["BDT Status"].includes("Rejected")) rejectedCount++;
          });

          const total = acceptedCount + rejectedCount;
          const acceptedPct = total > 0 ? acceptedCount / total : 0;
          const rejectedPct = total > 0 ? rejectedCount / total : 0;

          analysisWs.columns = [
            { header: "Status", key: "status", width: 25 },
            { header: "Count", key: "count", width: 20 },
            { header: "Percentage", key: "percentage", width: 20 },
          ];

          analysisWs.addRow({
            status: "Accepted",
            count: acceptedCount,
            percentage: acceptedPct,
          });
          analysisWs.addRow({
            status: "Rejected",
            count: rejectedCount,
            percentage: rejectedPct,
          });
          analysisWs.addRow({ status: "Total", count: total, percentage: 1 });

          // Style Analysis Sheet
          const analysisHeader = analysisWs.getRow(1);
          analysisHeader.font = { bold: true, color: { argb: "FFFFFFFF" } };
          analysisHeader.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF0f172a" },
          };
          analysisHeader.alignment = {
            vertical: "middle",
            horizontal: "center",
          };

          analysisWs.getColumn("percentage").numFmt = "0.0%";
          analysisWs.getColumn("count").alignment = { horizontal: "center" };
          analysisWs.getColumn("percentage").alignment = {
            horizontal: "center",
          };

          analysisWs.eachRow((row) => {
            row.eachCell((cell) => {
              cell.border = {
                top: { style: "thin", color: { argb: "FFe2e8f0" } },
                left: { style: "thin", color: { argb: "FFe2e8f0" } },
                bottom: { style: "thin", color: { argb: "FFe2e8f0" } },
                right: { style: "thin", color: { argb: "FFe2e8f0" } },
              };
            });
          });

          const buffer = await newWb.xlsx.writeBuffer();
          const blob = new Blob([buffer], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          });
          const url = URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          a.download = `BDT_Audit_Report_${new Date().toISOString().split("T")[0]}.xlsx`;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);

          setReportLoading(false);
        } catch (innerErr) {
          console.error("Inner Report Error:", innerErr);
          alert("Error processing Acceptance form: " + innerErr.message);
          setReportLoading(false);
        }
      };
      reader.readAsArrayBuffer(acceptanceFile);
    } catch (err) {
      console.error("Report Generation Error:", err);
      alert("Failed to generate report: " + err.message);
      setReportLoading(false);
    }
  };

  const clean = (str) =>
    str
      ?.toString()
      .replace(/[^a-zA-Z0-9]/g, "")
      .toUpperCase() || "";

  const handleExportTopOffenders = async () => {
    console.log("[Export] Starting Top Offenders export...");
    try {
      const results = history[0]?.results || {};
      const categoryName =
        CATEGORIES.find((c) => c.id === dashboardCategory)?.name || "Category";
      const offenders = results[dashboardCategory]?.topOffenders || [];

      console.log(
        `[Export] Category: ${categoryName}, Records: ${offenders.length}`,
      );

      if (offenders.length === 0) {
        alert("No offenders data to export.");
        return;
      }

      const reportData = offenders.slice(0, 10).map((site, idx) => {
        const filteredAlarms = currentAlarms.filter(
          (a) => clean(a.siteCode) === clean(site.siteCode),
        );
        const calc = calculateSiteBackup(filteredAlarms);
        const currBackup = calc ? formatDuration(calc) : "N/A";

        return {
          Rank: idx + 1,
          "Site Code": site.siteCode || "N/A",
          "Site Name": site.siteName || "Unknown",
          Region: site.zone || "Unknown",
          "Freq.": site.count || 0,
          "Max DT": formatDuration(site.maxDurationMins),
          "Current Backup": currBackup,
        };
      });

      console.log("[Export] Report data prepared, creating workbook...");

      const newWb = new ExcelJS.Workbook();
      const sheetName =
        `Top_${categoryName.replace(/[^a-zA-Z0-9]/g, "_")}`.substring(0, 20);
      const newWs = newWb.addWorksheet(sheetName);

      const headers = Object.keys(reportData[0]);

      console.log("[Export] Adding table to worksheet...");

      newWs.addTable({
        name: "OffendersTable",
        ref: "A1",
        headerRow: true,
        style: {
          theme: "TableStyleMedium9",
          showRowStripes: true,
        },
        columns: headers.map((h) => ({ name: h, filterButton: true })),
        rows: reportData.map((dataRow) => headers.map((h) => dataRow[h])),
      });

      const headerRow = newWs.getRow(1);
      headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
      headerRow.alignment = { vertical: "middle", horizontal: "center" };

      newWs.columns.forEach((column) => {
        column.width = 18;
      });
      newWs.getColumn(headers.indexOf("Site Name") + 1).width = 35;

      newWs.eachRow((row) => {
        row.eachCell((cell) => {
          if (row.number > 1) {
            cell.alignment = { vertical: "middle", wrapText: true };
          }
        });
      });

      console.log("[Export] Generating buffer...");
      const buffer = await newWb.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `Top_${categoryName.replace(/[^a-zA-Z0-9]/g, "_")}_Offenders_${new Date().toISOString().split("T")[0]}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      console.log("[Export] Success!");
    } catch (err) {
      console.error("[Export] Error:", err);
      alert("Export failed: " + err.message);
    }
  };

  const renderSiteSpotlight = () => {
    if (!selectedSite) return null;

    const siteAlarms = currentAlarms.filter((a) => {
      const codeMatch = clean(a.siteCode) === clean(selectedSite);
      if (!codeMatch) return false;

      const alarmDate = a.startTime ? new Date(a.startTime) : null;
      if (
        dateFilter.start &&
        alarmDate &&
        alarmDate < new Date(dateFilter.start)
      )
        return false;
      if (dateFilter.end && alarmDate && alarmDate > new Date(dateFilter.end))
        return false;

      return true;
    });

    const siteMetrics = CATEGORIES.map((cat) => ({
      ...cat,
      count: siteAlarms.filter((a) => a.categories?.includes(cat.id)).length,
      duration: siteAlarms
        .filter((a) => a.categories?.includes(cat.id))
        .reduce((acc, curr) => acc + (curr.durationMinutes || 0), 0),
    }));

    const filteredAlarms = siteAlarms.filter((a) => {
      if (!spotlightCategory) return true;
      return a.categories?.includes(spotlightCategory);
    });

    return (
      <div className="fixed inset-0 z-[60] bg-premium-950/95 backdrop-blur-2xl overflow-y-auto animate-fade-in p-8 lg:p-16">
        <div className="max-w-6xl mx-auto space-y-12">
          <header className="flex flex-col md:flex-row md:items-center justify-between gap-8">
            <div className="flex items-center gap-6">
              <button
                onClick={() => {
                  setSelectedSite(null);
                  setSpotlightCategory(null);
                }}
                className="w-14 h-14 rounded-2xl bg-white/5 text-premium-400 hover:bg-white/10 hover:text-white flex items-center justify-center transition-all"
              >
                <ArrowLeft size={24} />
              </button>
              <div className="flex-1">
                <div className="flex items-center gap-3 mb-4">
                  <ShieldCheck className="text-emerald-500" size={16} />
                  <span className="text-xs font-black uppercase tracking-[0.3em] text-premium-400">
                    Site Analysis
                  </span>
                </div>

                {(() => {
                  const siteDetails = siteList.find(
                    (s) => clean(s.shortCode) === clean(selectedSite),
                  );
                  const calculatedBackup = calculateSiteBackup(siteAlarms);

                  return (
                    <div className="relative">
                      {/* Background Watermark Icon */}
                      <div className="absolute -right-16 -top-24 opacity-[0.07] pointer-events-none select-none overflow-hidden">
                        <RadioTower
                          size={480}
                          strokeWidth={0.5}
                          className="text-white rotate-12"
                        />
                      </div>

                      <div className="relative z-10 flex flex-col md:flex-row items-start justify-between gap-8">
                        <div className="space-y-6 flex-1">
                          <div className="flex flex-col">
                            <h2 className="text-6xl md:text-8xl font-black tracking-tighter text-white uppercase leading-none">
                              {selectedSite}
                            </h2>
                            <p className="text-xl md:text-3xl font-bold text-premium-400 mt-2 md:mt-4 uppercase tracking-tight">
                              {siteDetails?.siteName || "Unknown Site Identity"}
                            </p>
                          </div>

                          <div className="flex items-center gap-4 bg-white/5 border border-white/10 rounded-2xl p-5 w-fit group hover:border-blue-500/50 transition-all shadow-xl">
                            <div className="w-12 h-12 rounded-xl bg-blue-500/10 flex items-center justify-center text-blue-400 group-hover:scale-110 transition-transform">
                              <Clock size={24} />
                            </div>
                            <div>
                              <div className="text-[10px] font-black uppercase tracking-widest text-premium-500 mb-1">
                                Calculated Avg Backup
                              </div>
                              <div className="text-2xl font-black text-white">
                                {calculatedBackup
                                  ? `${Math.round(calculatedBackup)} Mins`
                                  : "Insufficient Data"}
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })()}
              </div>
            </div>

            <div className="flex items-center gap-4 bg-white/10 p-2 rounded-2xl border border-white/10">
              <div className="relative">
                <Calendar
                  className="absolute left-4 top-1/2 -translate-y-1/2 text-premium-500"
                  size={16}
                />
                <input
                  type="date"
                  value={dateFilter.start}
                  onChange={(e) =>
                    setDateFilter({ ...dateFilter, start: e.target.value })
                  }
                  className="bg-transparent border-none text-[10px] font-black uppercase text-white pl-10 pr-4 outline-none w-36"
                />
              </div>
              <div className="text-premium-700 font-bold px-2">TO</div>
              <div className="relative">
                <Calendar
                  className="absolute left-4 top-1/2 -translate-y-1/2 text-premium-500"
                  size={16}
                />
                <input
                  type="date"
                  value={dateFilter.end}
                  onChange={(e) =>
                    setDateFilter({ ...dateFilter, end: e.target.value })
                  }
                  className="bg-transparent border-none text-[10px] font-black uppercase text-white pl-10 pr-4 outline-none w-36"
                />
              </div>
            </div>
          </header>

          <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-6">
            {siteMetrics.map((cat) => (
              <button
                key={cat.id}
                onClick={() =>
                  setSpotlightCategory(
                    spotlightCategory === cat.id ? null : cat.id,
                  )
                }
                className={`p-8 rounded-[2rem] border transition-all text-left group ${
                  spotlightCategory === cat.id
                    ? "bg-blue-600 border-blue-400 shadow-2xl shadow-blue-600/40 scale-105"
                    : "bg-white/5 border-white/5 hover:border-white/20 hover:bg-white/10"
                }`}
              >
                <div className="flex items-center gap-4 mb-4">
                  <div
                    className={`text-2xl transition-transform ${spotlightCategory === cat.id ? "scale-125" : "group-hover:scale-110"}`}
                  >
                    {cat.icon}
                  </div>
                  <h4
                    className={`text-[10px] font-black uppercase tracking-widest ${spotlightCategory === cat.id ? "text-blue-100" : "text-premium-400"}`}
                  >
                    {cat.label}
                  </h4>
                </div>
                <div className="flex items-end justify-between">
                  <div>
                    <div
                      className={`text-3xl font-black ${spotlightCategory === cat.id ? "text-white" : "text-white"}`}
                    >
                      {cat.count}
                    </div>
                    <div
                      className={`text-[9px] font-bold uppercase tracking-widest ${spotlightCategory === cat.id ? "text-blue-200" : "text-premium-500"}`}
                    >
                      Alarm Sessions
                    </div>
                  </div>
                  <div className="text-right">
                    <div
                      className={`text-sm font-black ${spotlightCategory === cat.id ? "text-white" : "text-premium-300"}`}
                    >
                      {formatDuration(cat.duration)}
                    </div>
                    <div
                      className={`text-[9px] font-bold uppercase tracking-widest ${spotlightCategory === cat.id ? "text-blue-200" : "text-premium-500"}`}
                    >
                      Total DT
                    </div>
                  </div>
                </div>
              </button>
            ))}
          </div>

          <div className="bg-white/5 rounded-[2rem] border border-white/5 p-8 relative overflow-hidden">
            <div className="flex items-center justify-between mb-8">
              <h4 className="text-sm font-black uppercase tracking-widest flex items-center gap-3">
                <Activity className="text-blue-500" size={18} />
                Detailed Activity Timeline
                {spotlightCategory && (
                  <span className="ml-4 px-3 py-1 bg-blue-500/20 text-blue-400 text-[10px] rounded-lg animate-pulse uppercase tracking-widest">
                    Filtering: {spotlightCategory.replace("_", " ")}
                  </span>
                )}
              </h4>
              {spotlightCategory && (
                <button
                  onClick={() => setSpotlightCategory(null)}
                  className="text-[10px] font-black uppercase tracking-widest text-premium-500 hover:text-white transition-colors"
                >
                  Reset Filter
                </button>
              )}
            </div>

            <div className="overflow-x-auto">
              <table className="w-full">
                <thead>
                  <tr className="text-left border-b border-white/10">
                    <th className="pb-4 text-[10px] font-black uppercase tracking-widest text-premium-500">
                      Alarm Category
                    </th>
                    <th className="pb-4 text-[10px] font-black uppercase tracking-widest text-premium-500">
                      Site Name
                    </th>
                    <th className="pb-4 text-[10px] font-black uppercase tracking-widest text-premium-500">
                      Occurred At
                    </th>
                    <th className="pb-4 text-right text-[10px] font-black uppercase tracking-widest text-premium-500">
                      Duration
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-white/5">
                  {filteredAlarms
                    ?.sort(
                      (a, b) => new Date(b.startTime) - new Date(a.startTime),
                    )
                    .map((a, i) => (
                      <tr
                        key={i}
                        className="hover:bg-white/5 transition-colors group"
                      >
                        <td className="py-4">
                          <span className="text-[10px] font-black uppercase tracking-widest text-blue-400">
                            {a.category}
                          </span>
                        </td>
                        <td className="py-4 font-bold text-white text-sm uppercase">
                          {a.siteName || "N/A"}
                        </td>
                        <td className="py-4 text-[10px] font-black text-premium-400 uppercase">
                          {a.startTime
                            ? new Date(a.startTime).toLocaleString()
                            : "N/A"}
                        </td>
                        <td className="py-4 text-right font-mono text-sm font-black text-white">
                          {formatDuration(a.durationMinutes)}
                        </td>
                      </tr>
                    ))}
                  {filteredAlarms.length === 0 && (
                    <tr>
                      <td
                        colSpan="4"
                        className="py-20 text-center font-bold text-premium-500 uppercase tracking-widest"
                      >
                        No matching alarms found for this selection.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderDashboard = () => {
    if (history.length === 0)
      return (
        <div className="flex flex-col items-center justify-center py-32 animate-fade-in">
          <div className="w-24 h-24 bg-blue-500/10 rounded-[2rem] flex items-center justify-center mb-8">
            <Zap size={48} className="text-blue-500 shadow-glow" />
          </div>
          <h3 className="text-3xl font-black mb-3">Initialize Review</h3>
          <p className="text-premium-400 max-w-sm text-center font-medium">
            Upload your weekly energy master to generate active intelligence.
          </p>
        </div>
      );

    const data = history[0].results;

    return (
      <div className="animate-slide-up space-y-12 pb-32">
        <div className="flex flex-col md:flex-row md:items-end justify-between gap-8">
          <div>
            <div className="flex items-center gap-3 mb-4">
              <span className="flex h-2 w-2 rounded-full bg-emerald-500"></span>
              <span className="text-[10px] font-black uppercase tracking-[0.3em] text-emerald-500">
                Active intelligence session
              </span>
            </div>
            <h1 className="text-5xl font-black tracking-tight leading-none bg-gradient-to-r from-premium-900 to-premium-600 dark:from-white dark:to-premium-400 bg-clip-text text-transparent">
              Network Snapshot
            </h1>
            <div className="flex items-center gap-4 mt-6">
              <button
                onClick={() => {
                  setHistory([]);
                  setCurrentAlarms([]);
                  setBbiData(null);
                  setAutinFiles({
                    power: [],
                    down: [],
                    high_temp: [],
                    generator: [],
                    door: [],
                  });
                }}
                className="flex items-center gap-2 px-6 py-2.5 bg-rose-500/10 text-rose-600 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-rose-500 hover:text-white transition-all shadow-sm"
              >
                <RotateCcw size={14} /> New Analysis
              </button>
              <button
                onClick={() => {
                  setHistory([]); // Just clear the 'results' view to show upload screen, but keep currentAlarms?
                  // Wait, to "upload extra sheet" we need'to go back to upload but KEEP existing data.
                  // This requires changing handleFileUpload to append.
                }}
                className="flex items-center gap-2 px-6 py-2.5 bg-blue-500/10 text-blue-600 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-blue-500 hover:text-white transition-all shadow-sm"
              >
                <Plus size={14} /> Upload Extra
              </button>
            </div>
          </div>
          <div className="flex flex-col md:flex-row md:items-center gap-6">
            <div className="relative group min-w-[320px]">
              <Search
                className="absolute left-5 top-1/2 -translate-y-1/2 text-premium-400 group-focus-within:text-blue-500 transition-colors"
                size={18}
              />
              <input
                type="text"
                placeholder="Lookup Site Code..."
                value={searchSiteCode}
                onChange={(e) => setSearchSiteCode(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    setSelectedSite(searchSiteCode.toUpperCase());
                    setSearchSiteCode("");
                  }
                }}
                className="w-full bg-white dark:bg-premium-900 border-2 border-premium-100 dark:border-premium-800 rounded-[1.5rem] py-4 pl-14 pr-6 font-black text-xs outline-none focus:border-blue-500 shadow-sm focus:ring-4 ring-blue-500/5 transition-all"
              />
            </div>
            <button
              onClick={() => {
                const ws = XLSX.utils.json_to_sheet([
                  { Category: "Total Database Sites", Value: siteList.length },
                  ...CATEGORIES.map((cat) => ({
                    Category: cat.name,
                    Value: data[cat.id]?.total || 0,
                    "Affected Sites": data[cat.id]?.uniqueSites || 0,
                  })),
                ]);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Executive_Summary");
                XLSX.writeFile(
                  wb,
                  `Energy_Review_${new Date().toLocaleDateString()}.xlsx`,
                );
              }}
              className="flex items-center gap-2 px-8 py-4 bg-premium-900 dark:bg-white text-white dark:text-premium-950 rounded-2xl font-bold shadow-2xl hover:scale-105 active:scale-95 transition-all"
            >
              <Download size={20} /> Export Summary
            </button>
          </div>
        </div>

        {/* Global KPIs */}
        <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-6 xl:gap-8">
          {CATEGORIES.map((cat) => (
            <button
              key={cat.id}
              onClick={() => setDashboardCategory(cat.id)}
              className={`p-8 rounded-[2rem] border transition-all text-left group relative overflow-hidden ${
                dashboardCategory === cat.id
                  ? "bg-blue-600 border-blue-400 shadow-2xl scale-105"
                  : "bg-white dark:bg-premium-900 border-premium-100 dark:border-premium-800 hover:border-blue-400 hover:bg-premium-50"
              }`}
            >
              <div className="flex justify-between items-start mb-6 relative z-10">
                <div
                  className={`w-14 h-14 rounded-2xl flex items-center justify-center shadow-lg transition-transform ${
                    dashboardCategory === cat.id
                      ? "bg-white text-blue-600 scale-110"
                      : `${cat.color} text-white group-hover:scale-110`
                  }`}
                >
                  {cat.icon}
                </div>
                <div className="text-right">
                  <div
                    className={`text-3xl font-black tracking-tighter ${dashboardCategory === cat.id ? "text-white" : ""}`}
                  >
                    {data[cat.id]?.total || 0}
                  </div>
                  <div
                    className={`text-[10px] font-black uppercase mt-1 ${dashboardCategory === cat.id ? "text-blue-200" : "text-premium-400"}`}
                  >
                    Total Events
                  </div>
                </div>
              </div>
              <h3
                className={`font-black text-xl mb-4 relative z-10 ${dashboardCategory === cat.id ? "text-white" : ""}`}
              >
                {cat.name}
              </h3>
              <div
                className={`flex items-center justify-between pt-4 border-t relative z-10 ${
                  dashboardCategory === cat.id
                    ? "border-blue-400"
                    : "border-premium-100 dark:border-premium-800"
                }`}
              >
                <span
                  className={`text-xs font-bold ${dashboardCategory === cat.id ? "text-blue-100" : "text-premium-500"}`}
                >
                  Impacted Sites
                </span>
                <span
                  className={`text-lg font-black ${dashboardCategory === cat.id ? "text-white" : ""}`}
                >
                  {data[cat.id]?.uniqueSites || 0}
                </span>
              </div>
            </button>
          ))}
        </div>

        {/* Intelligence Insights Section */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          <Card className="lg:col-span-2 overflow-hidden flex flex-col">
            <div className="p-8 border-b border-premium-100 dark:border-premium-800 bg-premium-50/50 dark:bg-white/5 flex items-center justify-between">
              <div className="flex items-center gap-4">
                <TrendingUp className="text-rose-500" size={24} />
                <h3 className="font-black text-2xl uppercase tracking-tighter">
                  Top {CATEGORIES.find((c) => c.id === dashboardCategory)?.name}{" "}
                  Offenders
                </h3>
              </div>
              <button
                onClick={handleExportTopOffenders}
                className="flex items-center gap-2 px-4 py-2 bg-indigo-500/10 text-indigo-600 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-indigo-500 hover:text-white transition-all"
              >
                <Download size={14} /> Export Table
              </button>
            </div>
            <div className="overflow-x-auto min-h-[400px]">
              <table className="w-full text-left">
                <thead>
                  <tr className="text-[11px] font-black uppercase tracking-widest text-premium-400 bg-premium-50/30">
                    <th className="p-6">Site Identity</th>
                    <th className="p-6">Region</th>
                    <th className="p-6 text-center">
                      {CATEGORIES.find((c) => c.id === dashboardCategory)?.name}{" "}
                      Freq.
                    </th>
                    <th className="p-6 text-right">
                      {
                        CATEGORIES.find((c) => c.id === dashboardCategory)
                          ?.short
                      }{" "}
                      Max DT
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-premium-100 dark:divide-premium-800">
                  {(data[dashboardCategory]?.topOffenders || [])
                    .slice(0, 10)
                    .map((site, idx) => (
                      <tr
                        key={idx}
                        className="group hover:bg-premium-50/50 transition-colors"
                      >
                        <td className="p-6">
                          <div className="flex items-center gap-4">
                            <div className="font-black text-premium-200 text-xl tracking-tighter">
                              {(idx + 1).toString().padStart(2, "0")}
                            </div>
                            <div className="flex flex-col">
                              <span
                                className="font-black text-lg text-premium-900 dark:text-white uppercase cursor-pointer hover:text-blue-600 leading-none"
                                onClick={() => setSelectedSite(site.siteCode)}
                              >
                                {site.siteCode}
                              </span>
                              <span className="text-[10px] font-bold text-premium-500 truncate max-w-[200px] mt-1">
                                {site.siteName}
                              </span>
                              <div className="flex items-center gap-1.5 mt-2 text-[9px] font-black text-emerald-500 dark:text-emerald-400 uppercase tracking-widest bg-emerald-500/10 w-fit px-2 py-0.5 rounded-md border border-emerald-500/20">
                                <Clock size={10} />
                                <span>
                                  {(() => {
                                    const calc = calculateSiteBackup(
                                      currentAlarms.filter(
                                        (a) =>
                                          clean(a.siteCode) ===
                                          clean(site.siteCode),
                                      ),
                                    );
                                    return calc
                                      ? `${Math.round(calc)} Min Backup`
                                      : "No BBI Data";
                                  })()}
                                </span>
                              </div>
                            </div>
                          </div>
                        </td>
                        <td className="p-6 font-bold text-sm text-premium-600 dark:text-premium-400 capitalize">
                          {site.region || "No Region"}
                        </td>
                        <td className="p-6 text-center">
                          <div className="flex flex-col items-center">
                            <span className="px-3 py-1 bg-rose-500 text-white rounded-xl text-xs font-black shadow-lg shadow-rose-500/20">
                              {site.count}
                            </span>
                            <span className="text-[8px] font-black uppercase text-premium-400 mt-2">
                              Sessions
                            </span>
                          </div>
                        </td>
                        <td className="p-6 text-right">
                          <div className="flex flex-col items-end">
                            <span className="font-mono text-sm font-black text-rose-600">
                              {formatDuration(site.maxDuration)}
                            </span>
                            <span className="text-[8px] font-black uppercase text-premium-400 mt-1">
                              Downtime
                            </span>
                          </div>
                        </td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>
          </Card>

          <div className="space-y-8">
            <Card className="p-8">
              <div className="flex items-center gap-4 mb-8">
                <Filter size={24} className="text-premium-400" />
                <h3 className="font-black text-2xl">Regional Focus</h3>
              </div>
              <div className="space-y-8">
                {Object.entries(data[dashboardCategory]?.regions || {})
                  .sort((a, b) => b[1] - a[1])
                  .slice(0, 6)
                  .map(([region, count], idx) => {
                    const max = Math.max(
                      ...Object.values(
                        data[dashboardCategory]?.regions || { 0: 1 },
                      ),
                    );
                    const percent = (count / max) * 100;
                    return (
                      <div key={idx} className="space-y-3">
                        <div className="flex justify-between items-end">
                          <span className="text-sm font-black uppercase tracking-tight text-premium-700 dark:text-premium-300">
                            {region}
                          </span>
                          <span className="text-xs font-black text-rose-500">
                            {count} Events
                          </span>
                        </div>
                        <div className="h-2.5 w-full bg-premium-100 dark:bg-premium-800 rounded-full overflow-hidden">
                          <div
                            className="h-full bg-gradient-to-r from-rose-400 to-rose-600 rounded-full shadow-inner transition-all duration-1000"
                            style={{ width: `${percent}%` }}
                          ></div>
                        </div>
                      </div>
                    );
                  })}
              </div>
            </Card>

            <Card className="p-8 bg-premium-900 dark:bg-white text-white dark:text-premium-900 border-none relative overflow-hidden">
              <div className="relative z-10">
                <div className="text-[10px] font-black uppercase tracking-widest mb-2 opacity-60">
                  System Recommendation
                </div>
                <h4 className="text-xl font-bold mb-4 leading-tight">
                  Focus audits on top 3 sites showing recurrent Link Failures.
                </h4>
                <button className="flex items-center gap-2 font-black text-xs uppercase tracking-widest group">
                  View full details{" "}
                  <ChevronRight
                    size={14}
                    className="group-hover:translate-x-1 transition-transform"
                  />
                </button>
              </div>
              <div className="absolute -right-10 -bottom-10 opacity-10">
                <ShieldCheck size={120} />
              </div>
            </Card>
          </div>
        </div>

        {/* Battery Backup Indicator (BBI) Section */}
        {bbiData && (
          <div className="space-y-8">
            <header className="flex flex-col md:flex-row md:items-center justify-between border-b-2 border-premium-100 dark:border-premium-900 pb-8 gap-6">
              <div className="flex items-center gap-6">
                <div className="w-16 h-16 bg-amber-500 text-white rounded-[1.5rem] flex items-center justify-center shadow-2xl shadow-amber-500/30">
                  <Battery size={32} />
                </div>
                <div>
                  <h2 className="text-4xl font-black tracking-tight">
                    BBI Analysis
                  </h2>
                  <p className="font-medium text-premium-500">
                    Calculated Backup Precision & PLD Audit
                  </p>
                </div>
              </div>
              <div className="flex bg-premium-100 dark:bg-premium-800 p-1.5 rounded-2xl shadow-inner">
                <button
                  onClick={() => setActiveBbiTab("mismatched")}
                  className={`px-8 py-3 rounded-xl text-xs font-black transition-all ${activeBbiTab === "mismatched" ? "bg-white dark:bg-premium-700 shadow-xl text-blue-600" : "text-premium-400 hover:text-premium-600"}`}
                >
                  Mismatched
                </button>
                <button
                  onClick={() => setActiveBbiTab("pld")}
                  className={`px-8 py-3 rounded-xl text-xs font-black transition-all ${activeBbiTab === "pld" ? "bg-white dark:bg-premium-700 shadow-xl text-blue-600" : "text-premium-400 hover:text-premium-600"}`}
                >
                  PLD Abnormal
                </button>
              </div>
            </header>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-8 font-poppins">
              <Card className="p-8 border-l-8 border-l-rose-500 bg-gradient-to-br from-white to-rose-50 dark:from-premium-900 dark:to-rose-900/10">
                <div className="flex items-center justify-between mb-2">
                  <h4 className="font-black text-rose-500 uppercase tracking-widest text-xs">
                    Mismatch Alert
                  </h4>
                  <span className="text-4xl font-black tracking-tighter">
                    {bbiData.backupMismatched.length}
                  </span>
                </div>
                <p className="text-sm text-premium-600 font-medium">
                  Sites falling &gt;30m below their BDT capability.
                </p>
              </Card>
              <Card className="p-8 border-l-8 border-l-amber-500 bg-gradient-to-br from-white to-amber-50 dark:from-premium-900 dark:to-amber-900/10">
                <div className="flex items-center justify-between mb-2">
                  <h4 className="font-black text-amber-500 uppercase tracking-widest text-xs">
                    PLD Logic Warning
                  </h4>
                  <span className="text-4xl font-black tracking-tighter">
                    {bbiData.pldAbnormal.length}
                  </span>
                </div>
                <p className="text-sm text-premium-600 font-medium">
                  Active PLD sites showing erratic (+/- 60m) deviations.
                </p>
              </Card>
            </div>

            <Card className="overflow-hidden border-2 border-premium-100">
              <div className="p-8 border-b border-premium-100 dark:border-premium-900 flex items-center justify-between bg-premium-50/20">
                <h3 className="font-black text-xl flex items-center gap-3">
                  <Database size={24} className="text-premium-400" />
                  BBI Analysis Workbench
                </h3>
                <button
                  onClick={async () => {
                    const list =
                      activeBbiTab === "mismatched"
                        ? bbiData.backupMismatched
                        : bbiData.pldAbnormal;

                    if (list.length === 0) return;

                    const reportData = list.map((item) => ({
                      "Short Code": item.shortCode,
                      "Site Name": item.siteName,
                      "BDT (Baseline)": item.bdt,
                      "AVG Backup (Real)": Math.round(item.avgBackup),
                      "Deviation (Mins)": Math.round(item.tolerance),
                      "PLD Status": item.pld ? "YES" : "NO",
                      "Technical Comment": item.comment || "",
                    }));

                    const workbook = new ExcelJS.Workbook();
                    const worksheet = workbook.addWorksheet("BBI_Analysis");
                    const headers = Object.keys(reportData[0]);

                    worksheet.addTable({
                      name: "BBITable",
                      ref: "A1",
                      headerRow: true,
                      style: {
                        theme: "TableStyleMedium9",
                        showRowStripes: true,
                      },
                      columns: headers.map((h) => ({
                        name: h,
                        filterButton: true,
                      })),
                      rows: reportData.map((row) => headers.map((h) => row[h])),
                    });

                    // Styles
                    const headerRow = worksheet.getRow(1);
                    headerRow.font = {
                      bold: true,
                      color: { argb: "FFFFFFFF" },
                    };
                    headerRow.alignment = {
                      vertical: "middle",
                      horizontal: "center",
                    };

                    worksheet.columns.forEach((column, i) => {
                      const header = headers[i];
                      if (header === "Site Name") column.width = 35;
                      else if (header === "Technical Comment")
                        column.width = 40;
                      else column.width = 18;
                    });

                    worksheet.eachRow((row) => {
                      row.eachCell((cell) => {
                        cell.alignment = {
                          vertical: "middle",
                          wrapText: true,
                          horizontal: "center",
                        };
                        if (row.number === 1)
                          cell.alignment.horizontal = "center";
                      });
                    });

                    const buffer = await workbook.xlsx.writeBuffer();
                    const blob = new Blob([buffer], {
                      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    });
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement("a");
                    a.href = url;
                    a.download = `Energy_Review_BBI_${activeBbiTab}_${new Date().toISOString().split("T")[0]}.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                  }}
                  className="px-6 py-2.5 bg-blue-600 text-white rounded-xl text-xs font-black shadow-lg shadow-blue-500/20 hover:scale-105 active:scale-95 transition-all"
                >
                  Generate BBI Report
                </button>
              </div>
              <div className="overflow-x-auto min-h-[500px]">
                <table className="w-full text-left">
                  <thead className="bg-premium-50/50 dark:bg-white/5 border-b border-premium-100 dark:border-premium-900">
                    <tr className="text-[10px] font-black uppercase tracking-widest text-premium-400">
                      <th className="p-6">Site Identity</th>
                      <th className="p-6 text-center">Config</th>
                      <th className="p-6 text-center">Summary</th>
                      <th className="p-6 text-center">AVG Backup</th>
                      <th className="p-6 text-center">BDT Ratio</th>
                      <th className="p-6">Technical Comment</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-premium-100 dark:divide-premium-900">
                    {(activeBbiTab === "mismatched"
                      ? bbiData.backupMismatched
                      : bbiData.pldAbnormal
                    ).map((item, idx) => (
                      <tr
                        key={idx}
                        className="group hover:bg-premium-50/30 transition-all"
                      >
                        <td className="p-6">
                          <div className="mb-2">
                            <span className="font-black text-premium-950 dark:text-white block tracking-tighter">
                              {item.shortCode}
                            </span>
                            <span className="text-[10px] font-bold text-premium-400 uppercase tracking-tight truncate max-w-[200px]">
                              {item.siteName}
                            </span>
                          </div>
                          <div className="flex gap-1.5">
                            {(() => {
                              const nodal = parseInt(item.nodalDeg) || 0;
                              const label = item.vip
                                ? `VIP+${nodal}`
                                : `1+${nodal}`;
                              return (
                                <span
                                  className={`px-2 py-0.5 rounded-md text-[9px] font-black shadow-sm ${item.vip ? "bg-amber-500 text-white" : "bg-blue-600 text-white"}`}
                                >
                                  {label}
                                </span>
                              );
                            })()}
                            {item.pld && (
                              <span className="px-2 py-0.5 bg-emerald-100 text-emerald-600 rounded-md text-[9px] font-black">
                                PLD
                              </span>
                            )}
                          </div>
                        </td>
                        <td className="p-6 text-center">
                          <div className="flex flex-col items-center">
                            <span className="text-xs font-bold text-premium-600">
                              {item.batteries || "---"}
                            </span>
                            <span className="text-[9px] font-black text-premium-400 uppercase">
                              Assets
                            </span>
                          </div>
                        </td>
                        <td className="p-6 text-center">
                          <div className="flex flex-col items-center">
                            <span className="font-black text-premium-800 dark:text-premium-200">
                              {item.powerAlarmsCount}/{item.downAlarmsCount}
                            </span>
                            <span className="text-[9px] font-bold text-premium-400 uppercase">
                              Power vs Down
                            </span>
                          </div>
                        </td>
                        <td className="p-6 text-center">
                          <div className="flex flex-col items-center">
                            <span className="text-lg font-black text-amber-600 tracking-tighter">
                              {Math.round(item.avgBackup)}m
                            </span>
                            <div className="flex items-center gap-1">
                              {item.tolerance > 0 ? (
                                <ArrowDownRight
                                  size={10}
                                  className="text-rose-500"
                                />
                              ) : (
                                <ArrowUpRight
                                  size={10}
                                  className="text-emerald-500"
                                />
                              )}
                              <span
                                className={`text-[9px] font-black ${item.tolerance > 0 ? "text-rose-500" : "text-emerald-500"}`}
                              >
                                {Math.abs(item.tolerance).toFixed(0)}m Δ
                              </span>
                            </div>
                          </div>
                        </td>
                        <td className="p-6 text-center font-black text-premium-900 dark:text-white text-lg tracking-tighter">
                          {item.bdt}m
                        </td>
                        <td className="p-6">
                          <div className="relative group/input">
                            <MessageSquare
                              size={14}
                              className="absolute left-3 top-1/2 -translate-y-1/2 text-premium-400 pointer-events-none"
                            />
                            <input
                              type="text"
                              value={item.comment}
                              onChange={(e) => {
                                const updatedAll = [...bbiData.all];
                                const siteIdx = updatedAll.findIndex(
                                  (s) => s.shortCode === item.shortCode,
                                );
                                if (siteIdx !== -1) {
                                  updatedAll[siteIdx].comment = e.target.value;
                                  const mismatched = updatedAll
                                    .filter(
                                      (s) =>
                                        s.pld === false && s.tolerance > 30,
                                    )
                                    .sort((a, b) => b.tolerance - a.tolerance);
                                  const pldAbnormal = updatedAll
                                    .filter(
                                      (s) =>
                                        s.pld === true &&
                                        Math.abs(s.tolerance) > 60,
                                    )
                                    .sort((a, b) => b.tolerance - a.tolerance);
                                  setBbiData({
                                    all: updatedAll,
                                    backupMismatched: mismatched,
                                    pldAbnormal: pldAbnormal,
                                  });
                                }
                              }}
                              className="w-full bg-premium-100/50 dark:bg-white/5 border border-premium-200 dark:border-premium-900 rounded-xl pl-10 pr-4 py-3 text-xs font-medium focus:ring-2 ring-blue-500/20 outline-none transition-all placeholder:text-premium-300"
                              placeholder={
                                activeBbiTab === "pld"
                                  ? "System recommendation active..."
                                  : "Specify action taken..."
                              }
                              disabled={activeBbiTab === "pld"}
                            />
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Card>
          </div>
        )}
      </div>
    );
  };

  const renderSettings = () => (
    <div className="animate-slide-up max-w-6xl mx-auto space-y-12 pb-32">
      <header className="text-center">
        <h3 className="text-5xl font-black mb-3 tracking-tighter">
          System Console
        </h3>
        <p className="text-premium-500 font-bold uppercase text-[10px] tracking-[0.4em]">
          Hardware & Database configuration
        </p>
      </header>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-10">
        <div className="lg:col-span-2 space-y-10">
          <Card className="p-10">
            <div className="flex items-center gap-5 mb-10">
              <div className="w-14 h-14 bg-emerald-500/10 text-emerald-500 rounded-2xl flex items-center justify-center shadow-inner">
                <Database size={28} />
              </div>
              <div>
                <h4 className="text-2xl font-black">Connection Schema</h4>
                <p className="text-sm font-medium text-premium-400">
                  Manage your Google Cloud Database integration
                </p>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
              <div className="space-y-3">
                <label className="text-[10px] font-black uppercase text-premium-400 tracking-widest pl-1">
                  Spreadsheet ID
                </label>
                <input
                  type="text"
                  value={sysConfig.spreadsheetId}
                  onChange={(e) =>
                    setSysConfig({
                      ...sysConfig,
                      spreadsheetId: e.target.value,
                    })
                  }
                  className="w-full bg-premium-50 dark:bg-white/5 border-2 border-premium-100 dark:border-premium-900 rounded-2xl px-5 py-4 font-mono text-sm focus:border-emerald-500/50 focus:ring-4 ring-emerald-500/5 outline-none transition-all"
                  placeholder="Sheet ID..."
                />
              </div>
              <div className="space-y-3">
                <label className="text-[10px] font-black uppercase text-premium-400 tracking-widest pl-1">
                  Target Worksheet
                </label>
                <input
                  type="text"
                  value={sysConfig.sheetName}
                  onChange={(e) =>
                    setSysConfig({ ...sysConfig, sheetName: e.target.value })
                  }
                  className="w-full bg-premium-50 dark:bg-white/5 border-2 border-premium-100 dark:border-premium-900 rounded-2xl px-5 py-4 font-bold text-sm focus:border-emerald-500/50 focus:ring-4 ring-emerald-500/5 outline-none transition-all"
                  placeholder="e.g. MasterSiteList"
                />
              </div>
            </div>

            <div className="mt-12 flex items-center gap-6">
              <button
                onClick={() => {
                  localStorage.setItem(
                    "energy_review_config",
                    JSON.stringify(sysConfig),
                  );
                  alert(
                    "Configuration successfully committed to localStorage.",
                  );
                }}
                className="flex items-center gap-3 bg-emerald-600 hover:bg-emerald-700 text-white px-8 py-4 rounded-2xl font-black text-sm shadow-xl shadow-emerald-500/20 active:scale-95 transition-all"
              >
                <Save size={20} /> Commit Settings
              </button>
              <button
                onClick={() => {
                  if (
                    window.confirm(
                      "Restore factory Alexandria database settings? This cannot be undone.",
                    )
                  ) {
                    setSysConfig(DEFAULT_CONFIG);
                    localStorage.removeItem("energy_review_config");
                  }
                }}
                className="flex items-center gap-2 text-rose-500 font-black text-xs uppercase tracking-widest hover:bg-rose-50 px-4 py-2 rounded-xl transition-colors"
              >
                <RotateCcw size={16} /> Factory Reset
              </button>
            </div>
          </Card>

          <Card className="p-10">
            <div className="flex items-center gap-5 mb-10">
              <div className="w-14 h-14 bg-blue-500/10 text-blue-500 rounded-2xl flex items-center justify-center">
                <Activity size={28} />
              </div>
              <div>
                <h4 className="text-2xl font-black">Data Property Mapper</h4>
                <p className="text-sm font-medium text-premium-400">
                  Map properties to zero-indexed spreadsheet columns (A=0, B=1,
                  etc.)
                </p>
              </div>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-3 gap-8">
              {Object.entries(sysConfig.mapping).map(([key, val]) => (
                <div key={key} className="space-y-3">
                  <label className="text-[10px] font-black uppercase text-premium-500 tracking-tight pl-1">
                    {key.replace(/([A-Z])/g, " $1")}
                  </label>
                  <input
                    type="number"
                    value={val}
                    onChange={(e) =>
                      setSysConfig({
                        ...sysConfig,
                        mapping: {
                          ...sysConfig.mapping,
                          [key]: parseInt(e.target.value) || 0,
                        },
                      })
                    }
                    className="w-full bg-premium-50 dark:bg-white/5 border border-premium-100 dark:border-premium-900 rounded-xl px-4 py-3 font-mono text-sm focus:ring-4 ring-blue-500/10 outline-none transition-all text-center"
                  />
                </div>
              ))}
            </div>
          </Card>
        </div>

        <div className="space-y-10">
          <Card className="p-8 bg-premium-50/50 border-2 border-premium-100">
            <h5 className="text-[10px] font-black uppercase tracking-[0.2em] text-premium-400 mb-8">
              Network Status
            </h5>
            <div className="flex items-center justify-between mb-10">
              <div className="flex flex-col">
                <span
                  className={`text-xl font-black ${siteListLoading ? "text-amber-500" : siteListError ? "text-rose-500" : "text-emerald-500"}`}
                >
                  {siteListLoading
                    ? "SYNCING..."
                    : siteListError
                      ? "OFFLINE"
                      : "SECURE"}
                </span>
                <span className="text-[9px] font-bold text-premium-400 uppercase leading-none">
                  Socket Connection
                </span>
              </div>
              <div className="text-5xl font-black tracking-tighter">
                {siteList.length}
              </div>
            </div>
            <button
              onClick={loadSiteList}
              disabled={siteListLoading}
              className={`w-full flex items-center justify-center gap-3 py-5 rounded-2xl font-black text-sm transition-all ${
                siteListLoading
                  ? "bg-premium-100 text-premium-300 cursor-wait animate-pulse"
                  : "bg-premium-950 dark:bg-white text-white dark:text-premium-950 shadow-2xl shadow-premium-900/20 active:scale-95"
              }`}
            >
              <RefreshCw
                size={20}
                className={siteListLoading ? "animate-spin" : ""}
              />
              {siteListLoading ? "Fetching Cloud..." : "Full Sync Now"}
            </button>
            {lastSyncTime && (
              <div className="mt-6 flex items-center justify-center gap-2 text-[10px] font-bold text-premium-400 uppercase tracking-widest">
                <Clock size={12} /> Sync: {lastSyncTime.toLocaleTimeString()}
              </div>
            )}
          </Card>

          <Card className="p-8 flex flex-col gap-8">
            <h5 className="text-[10px] font-black uppercase tracking-[0.2em] text-premium-400">
              Database Stream
            </h5>
            <div className="space-y-4 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
              {siteList.slice(0, 15).map((s, idx) => (
                <div
                  key={idx}
                  className="flex items-center justify-between p-4 bg-premium-50 dark:bg-white/5 rounded-2xl border border-premium-100 dark:border-premium-900 group transition-all hover:border-blue-200"
                >
                  <span className="font-mono text-xs font-black text-blue-600 group-hover:scale-110 transition-transform">
                    {s.shortCode}
                  </span>
                  <span className="text-[10px] font-bold text-premium-500 truncate max-w-[140px]">
                    {s.siteName}
                  </span>
                </div>
              ))}
            </div>
            {siteList.length > 15 && (
              <div className="text-center text-[10px] font-black text-premium-300 uppercase tracking-widest">
                + {siteList.length - 15} more records
              </div>
            )}
          </Card>
        </div>
      </div>

      {/* Appearance Settings */}
      <Card className="p-10">
        <div className="flex items-center gap-5 mb-10">
          <div className="w-14 h-14 bg-purple-500/10 text-purple-500 rounded-2xl flex items-center justify-center shadow-inner">
            <Palette size={28} />
          </div>
          <div>
            <h4 className="text-2xl font-black">Appearance</h4>
            <p className="text-sm font-medium text-premium-400">
              Customize your workspace aesthetics
            </p>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
          <div className="space-y-4">
            <h5 className="text-[10px] font-black uppercase tracking-[0.2em] text-premium-500">
              Display Mode
            </h5>
            <div className="flex items-center gap-4 bg-premium-50 dark:bg-white/5 p-2 rounded-2xl border border-premium-100 dark:border-premium-900">
              <button
                onClick={() => setNightMode(false)}
                className={`flex-1 flex items-center justify-center gap-2 py-4 rounded-xl font-black text-sm transition-all ${
                  !nightMode
                    ? "bg-white text-blue-600 shadow-xl scale-100"
                    : "text-premium-400 hover:text-premium-600 scale-95"
                }`}
              >
                <Sun size={18} /> Light Mode
              </button>
              <button
                onClick={() => setNightMode(true)}
                className={`flex-1 flex items-center justify-center gap-2 py-4 rounded-xl font-black text-sm transition-all ${
                  nightMode
                    ? "bg-premium-800 text-blue-400 shadow-xl shadow-black/20 scale-100"
                    : "text-premium-400 hover:text-premium-600 scale-95"
                }`}
              >
                <Moon size={18} /> Night Mode
              </button>
            </div>
          </div>

          <div className="space-y-4">
            <h5 className="text-[10px] font-black uppercase tracking-[0.2em] text-premium-500">
              Color Theme
            </h5>
            <div className="grid grid-cols-3 gap-4">
              {[
                { id: "default", label: "Midnight Blue", cls: "bg-slate-500" },
                { id: "orange", label: "Sunset Orange", cls: "bg-orange-500" },
                {
                  id: "emerald",
                  label: "Emerald Green",
                  cls: "bg-emerald-500",
                },
              ].map((t) => (
                <button
                  key={t.id}
                  onClick={() => setAppTheme(t.id)}
                  className={`flex flex-col items-center gap-3 p-4 rounded-2xl border-2 transition-all ${
                    appTheme === t.id
                      ? "border-blue-500 bg-blue-500/5 shadow-lg shadow-blue-500/10 scale-105"
                      : "border-premium-100 dark:border-premium-800 hover:border-premium-300 dark:hover:border-premium-600 scale-100"
                  }`}
                >
                  <div
                    className={`w-8 h-8 rounded-full shadow-inner ${t.cls}`}
                  />
                  <span className="text-[10px] font-black uppercase tracking-widest text-premium-600 dark:text-premium-300 text-center">
                    {t.label}
                  </span>
                </button>
              ))}
            </div>
          </div>
        </div>
      </Card>

      {/* Security & Licensing Settings */}
      <Card className="p-10">
        <div className="flex items-center gap-5 mb-10">
          <div className="w-14 h-14 bg-blue-500/10 text-blue-500 rounded-2xl flex items-center justify-center shadow-inner animate-pulse">
            <ShieldCheck size={28} />
          </div>
          <div>
            <h4 className="text-2xl font-black">Security & Access Lock</h4>
            <p className="text-sm font-medium text-premium-400">
              Manage software evaluation locks and trial key validations
            </p>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
          <div className="space-y-4">
            <h5 className="text-[10px] font-black uppercase tracking-[0.2em] text-premium-500">
              Web Client Lock Screen
            </h5>
            <div className="bg-premium-50 dark:bg-white/5 p-6 rounded-[2rem] border border-premium-100 dark:border-premium-900 space-y-4">
              <div className="flex items-center justify-between">
                <div className="flex flex-col pr-4">
                  <span className="text-xs font-black uppercase tracking-wider mb-1">
                    Require Trial Key Activation
                  </span>
                  <span className="text-[10px] font-bold text-premium-400 leading-normal">
                    When active, a secure trial screen will lock browser visitors until a valid trial key is entered. (Web version only)
                  </span>
                </div>
                <button
                  onClick={() => {
                    const newState = !trialLockEnabled;
                    setTrialLockEnabled(newState);
                    localStorage.setItem("energy_review_trial_lock_enabled", JSON.stringify(newState));
                    if (!newState) {
                      setIsActivated(true);
                      localStorage.setItem("energy_review_activated", JSON.stringify(true));
                    }
                  }}
                  className={`w-16 h-8 rounded-full p-1 transition-all duration-300 relative focus:outline-none shrink-0 ${
                    trialLockEnabled ? "bg-blue-600" : "bg-premium-300 dark:bg-premium-800"
                  }`}
                >
                  <div
                    className={`w-6 h-6 rounded-full bg-white shadow-md transform transition-all duration-300 ${
                      trialLockEnabled ? "translate-x-8" : "translate-x-0"
                    }`}
                  />
                </button>
              </div>
            </div>
          </div>

          <div className="space-y-4">
            <h5 className="text-[10px] font-black uppercase tracking-[0.2em] text-premium-500">
              License Status
            </h5>
            <div className="bg-premium-50 dark:bg-white/5 p-6 rounded-[2rem] border border-premium-100 dark:border-premium-900 space-y-4">
              <div className="flex items-center justify-between">
                <div className="flex flex-col">
                  <span className="text-xs font-black uppercase tracking-wider mb-1">
                    Active License Status
                  </span>
                  <span className="text-[10px] font-bold text-premium-400">
                    Current instance licensing parameters.
                  </span>
                </div>
                <span className={`px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest shrink-0 ${
                  isActivated ? "bg-emerald-500/10 text-emerald-500" : "bg-amber-500/10 text-amber-500"
                }`}>
                  {isActivated ? "Activated (Simulated)" : "Unactivated"}
                </span>
              </div>
              {isActivated && (
                <div className="pt-2">
                  <button
                    onClick={() => {
                      if (window.confirm("Are you sure you want to deactivate and lock the software trial?")) {
                        setIsActivated(false);
                        localStorage.removeItem("energy_review_activated");
                        alert("Software trial deactivated.");
                      }
                    }}
                    className="text-[10px] font-black text-rose-500 hover:text-rose-400 uppercase tracking-widest"
                  >
                    Reset & Lock Trial Again
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
      </Card>

      {/* BDT Rules Configuration Section */}
      <Card className="p-10">
        <div className="flex items-center justify-between mb-10">
          <div className="flex items-center gap-5">
            <div className="w-14 h-14 bg-indigo-500/10 text-indigo-500 rounded-2xl flex items-center justify-center shadow-inner">
              <ClipboardCheck size={28} />
            </div>
            <div>
              <h4 className="text-2xl font-black">BDT Rule Settings</h4>
              <p className="text-sm font-medium text-premium-400">
                Configure validation thresholds before starting any BDT audit
              </p>
            </div>
          </div>
          {bdtConfigSaved && (
            <span className="flex items-center gap-2 px-4 py-2 bg-emerald-500/10 text-emerald-500 rounded-xl text-[10px] font-black uppercase tracking-widest animate-fade-in">
              <CheckCircle2 size={14} /> Settings Saved
            </span>
          )}
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
          {/* Left Column: Photos & Data Fill */}
          <div className="space-y-5">
            <h5 className="text-[10px] font-black uppercase tracking-[0.2em] text-indigo-500 flex items-center gap-2">
              <Camera size={14} /> Photos & Data Fill
            </h5>
            <div className="space-y-3">
              {[
                {
                  key: "photo_min_pass",
                  label: "Min Photos (Pass)",
                  type: "number",
                  desc: "Minimum photos to mark section as passed",
                },
                {
                  key: "photo_min_warn",
                  label: "Min Photos (Warning)",
                  type: "number",
                  desc: "Below this triggers a warning",
                },
                {
                  key: "photo_min_fail",
                  label: "Min Photos (Fail Threshold)",
                  type: "number",
                  desc: "Below this triggers a failure",
                },
                {
                  key: "basic_data_min_cells",
                  label: "Basic Data Min Cells",
                  type: "number",
                  desc: "Minimum filled cells in basic data section",
                },
                {
                  key: "busbar_min_fill_rate",
                  label: "Busbar Fill Rate (%)",
                  type: "percent",
                  desc: "Minimum busbar section fill percentage",
                },
              ].map((field) => (
                <div
                  key={field.key}
                  className="flex items-center justify-between bg-premium-50 dark:bg-white/5 p-4 rounded-2xl border border-transparent hover:border-indigo-200 dark:hover:border-indigo-800 transition-all group"
                >
                  <div className="flex flex-col">
                    <span className="text-[11px] font-bold text-premium-600 dark:text-premium-300 uppercase tracking-tight">
                      {field.label}
                    </span>
                    <span className="text-[9px] font-medium text-premium-400 mt-0.5 opacity-0 group-hover:opacity-100 transition-opacity">
                      {field.desc}
                    </span>
                  </div>
                  <input
                    type="number"
                    step={field.type === "percent" ? "0.01" : "1"}
                    value={
                      field.type === "percent"
                        ? (bdtConfig[field.key] * 100).toFixed(0)
                        : bdtConfig[field.key]
                    }
                    onChange={(e) => {
                      const val = parseFloat(e.target.value);
                      setBdtConfig((prev) => ({
                        ...prev,
                        [field.key]: field.type === "percent" ? val / 100 : val,
                      }));
                    }}
                    className="bg-white dark:bg-premium-800 border-2 border-premium-100 dark:border-premium-700 rounded-xl px-4 py-2.5 text-sm font-black w-24 text-center focus:border-indigo-500 focus:ring-4 ring-indigo-500/10 outline-none transition-all"
                  />
                </div>
              ))}
            </div>
          </div>

          {/* Right Column: Validation Thresholds */}
          <div className="space-y-5">
            <h5 className="text-[10px] font-black uppercase tracking-[0.2em] text-emerald-500 flex items-center gap-2">
              <ShieldCheck size={14} /> Validation Thresholds
            </h5>
            <div className="space-y-3">
              {[
                {
                  key: "rule1_batt_amp_max",
                  label: "Rule 1: Batt Amp Max (A)",
                  desc: "Max battery ampere sum at start",
                },
                {
                  key: "rule1_min_volt_start",
                  label: "Rule 1: Min Volt Start (V)",
                  desc: "Minimum starting voltage",
                },
                {
                  key: "rule1_min_rect_amp",
                  label: "Rule 1: Min Rect Amp (A)",
                  desc: "Minimum rectifier ampere at start",
                },
                {
                  key: "rule2_seq_tolerance",
                  label: "Rule 2: Seq Tolerance (V/A)",
                  desc: "Tolerance for sequential readings",
                },
                {
                  key: "rule3_balance_max",
                  label: "Rule 3: Amp Balance Max (A)",
                  desc: "Max delta between batteries and rectifier",
                },
                {
                  key: "rule4_theoretical_tolerance",
                  label: "Rule 4: Theoretical Tolerance (min)",
                  desc: "Acceptable gap from theoretical backup",
                },
                {
                  key: "rule6_12v_dissipated_max",
                  label: "Rule 6: 12V Dissipated Max (V)",
                  desc: "Max voltage dissipation for 12V batteries",
                },
                {
                  key: "rule7_interval_mins",
                  label: "Rule 7: Interval Mins",
                  desc: "Required reading interval after 47V",
                },
                {
                  key: "rule8_duration_mismatch_tolerance",
                  label: "Rule 8: Duration Buffer (min)",
                  desc: "Tolerance for cells vs busbar duration",
                },
              ].map((field) => (
                <div
                  key={field.key}
                  className="flex items-center justify-between bg-premium-50 dark:bg-white/5 p-4 rounded-2xl border border-transparent hover:border-emerald-200 dark:hover:border-emerald-800 transition-all group"
                >
                  <div className="flex flex-col">
                    <span className="text-[11px] font-bold text-premium-600 dark:text-premium-300 uppercase tracking-tight">
                      {field.label}
                    </span>
                    <span className="text-[9px] font-medium text-premium-400 mt-0.5 opacity-0 group-hover:opacity-100 transition-opacity">
                      {field.desc}
                    </span>
                  </div>
                  <input
                    type="number"
                    step="0.1"
                    value={bdtConfig[field.key]}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value);
                      setBdtConfig((prev) => ({
                        ...prev,
                        [field.key]: val,
                      }));
                    }}
                    className="bg-white dark:bg-premium-800 border-2 border-premium-100 dark:border-premium-700 rounded-xl px-4 py-2.5 text-sm font-black w-24 text-center focus:border-emerald-500 focus:ring-4 ring-emerald-500/10 outline-none transition-all"
                  />
                </div>
              ))}
            </div>
          </div>
        </div>

        <div className="mt-10 flex items-center gap-6 pt-8 border-t border-premium-100 dark:border-premium-800">
          <button
            onClick={() => {
              localStorage.setItem(
                "bdt_review_config",
                JSON.stringify(bdtConfig),
              );
              setBdtConfigSaved(true);
              setTimeout(() => setBdtConfigSaved(false), 3000);
            }}
            className="flex items-center gap-3 bg-indigo-600 hover:bg-indigo-700 text-white px-8 py-4 rounded-2xl font-black text-sm shadow-xl shadow-indigo-500/20 active:scale-95 transition-all"
          >
            <Save size={20} /> Save & Apply
          </button>
          <button
            onClick={() => {
              if (
                window.confirm(
                  "Reset all BDT rules to factory defaults? Current audit results will not be affected.",
                )
              ) {
                setBdtConfig(DEFAULT_BDT_CONFIG);
                localStorage.setItem(
                  "bdt_review_config",
                  JSON.stringify(DEFAULT_BDT_CONFIG),
                );
              }
            }}
            className="flex items-center gap-2 text-rose-500 font-black text-xs uppercase tracking-widest hover:bg-rose-50 dark:hover:bg-rose-500/10 px-4 py-2 rounded-xl transition-colors"
          >
            <RotateCcw size={16} /> Reset to Defaults
          </button>
        </div>
      </Card>
    </div>
  );

  const renderMemory = () => (
    <div className="animate-slide-up max-w-6xl mx-auto space-y-12 pb-32">
      <header className="text-center">
        <div className="flex items-center justify-center gap-4 mb-4">
          <Database className="text-blue-500" size={32} />
          <h3 className="text-5xl font-black tracking-tighter">Alarm Memory</h3>
        </div>
        <p className="text-premium-500 font-bold uppercase text-[10px] tracking-[0.4em]">
          Persistent Storage & Data Management
        </p>
        {currentAlarms.length > 0 && (
          <button
            onClick={() => handleMemoryAnalysis()}
            className="mt-8 px-10 py-4 bg-premium-950 dark:bg-white text-white dark:text-premium-950 rounded-2xl font-black text-xs uppercase tracking-[0.3em] shadow-2xl hover:bg-blue-600 dark:hover:bg-blue-600 hover:text-white hover:scale-105 transition-all flex items-center gap-4 mx-auto"
          >
            <Activity size={20} />
            Run Analysis on Stored Data
          </button>
        )}
      </header>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
        <Card className="p-8 border-l-8 border-l-blue-500">
          <div className="text-[10px] font-black uppercase text-premium-400 mb-2">
            Stored Sessions
          </div>
          <div className="text-4xl font-black">{alarmDatabases.length}</div>
        </Card>
        <Card className="p-8 border-l-8 border-l-indigo-500">
          <div className="text-[10px] font-black uppercase text-premium-400 mb-2">
            Total Alarms
          </div>
          <div className="text-4xl font-black">
            {currentAlarms.length.toLocaleString()}
          </div>
        </Card>
        <Card className="p-8 border-l-8 border-l-emerald-500">
          <div className="text-[10px] font-black uppercase text-premium-400 mb-2">
            Unique Sites
          </div>
          <div className="text-4xl font-black">
            {new Set(currentAlarms.map((a) => a.siteCode)).size}
          </div>
        </Card>
      </div>

      <Card className="overflow-hidden">
        <div className="p-8 border-b border-premium-100 dark:border-premium-900 flex items-center justify-between bg-premium-50/30">
          <h4 className="text-lg font-black uppercase tracking-tight">
            Stored Databases
          </h4>
          <button
            onClick={() => {
              if (
                window.confirm(
                  "Purge entire alarm memory? This cannot be undone.",
                )
              ) {
                setAlarmDatabases([]);
                setCurrentAlarms([]);
                setHistory([]);
              }
            }}
            className="px-6 py-2 bg-rose-500 text-white rounded-xl text-[10px] font-black uppercase tracking-widest shadow-lg shadow-rose-500/20 hover:scale-105 transition-all"
          >
            Clear All Memory
          </button>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="text-[10px] font-black uppercase tracking-widest text-premium-400 border-b border-premium-100 dark:border-premium-900">
                <th className="p-6">Source File / Type</th>
                <th className="p-6">Upload Date</th>
                <th className="p-6 text-center">Records</th>
                <th className="p-6 text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-premium-100 dark:divide-premium-800">
              {alarmDatabases.map((db) => (
                <tr
                  key={db.id}
                  className="hover:bg-premium-50/30 transition-colors"
                >
                  <td className="p-6">
                    <div className="font-black text-premium-900 dark:text-white">
                      {db.fileName}
                    </div>
                    <div className="text-[9px] font-black text-blue-500 uppercase mt-1">
                      {db.type}
                    </div>
                  </td>
                  <td className="p-6 text-xs font-bold text-premium-500">
                    {new Date(db.timestamp).toLocaleString()}
                  </td>
                  <td className="p-6 text-center">
                    <span className="px-3 py-1 bg-premium-100 dark:bg-white/5 rounded-lg text-xs font-black">
                      {db.alarmCount}
                    </span>
                  </td>
                  <td className="p-6 text-right">
                    <div className="flex items-center justify-end gap-2">
                      <button
                        onClick={() => handleMemoryAnalysis(db.id)}
                        className="flex items-center gap-2 px-3 py-2 text-indigo-500 hover:bg-indigo-50 dark:hover:bg-indigo-500/10 rounded-xl text-[10px] font-black uppercase tracking-widest transition-colors"
                      >
                        <Activity size={16} /> Analyze
                      </button>
                      <button
                        onClick={() => {
                          if (
                            window.confirm(`Delete ${db.fileName} from memory?`)
                          ) {
                            setAlarmDatabases((prev) =>
                              prev.filter((d) => d.id !== db.id),
                            );
                            setCurrentAlarms((prev) =>
                              prev.filter((a) => a.dbId !== db.id),
                            );
                          }
                        }}
                        className="p-2 text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-500/10 rounded-xl transition-colors"
                      >
                        <XCircle size={20} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
              {alarmDatabases.length === 0 && (
                <tr>
                  <td
                    colSpan="4"
                    className="p-20 text-center text-premium-400 font-black uppercase tracking-[0.2em]"
                  >
                    No alarms stored in memory.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );

  const renderBDT = () => {
    const statusIcon = (status, size = 16) => {
      if (status === "pass")
        return <CheckCircle2 size={size} className="text-emerald-500" />;
      if (status === "warning")
        return <AlertTriangle size={size} className="text-amber-500" />;
      if (status === "fail")
        return <XCircle size={size} className="text-rose-500" />;
      return <AlertCircle size={size} className="text-premium-400" />;
    };

    const statusBg = (s) =>
      ({
        pass: "bg-emerald-500/10 border-emerald-500/30",
        warning: "bg-amber-500/10 border-amber-500/30",
        fail: "bg-rose-500/10 border-rose-500/30",
        unknown: "bg-premium-500/10 border-premium-500/30",
      })[s] || "bg-premium-200/10 border-premium-300/30";

    const fillBar = (rate, status) => (
      <div className="h-1.5 w-full bg-premium-100 dark:bg-premium-800 rounded-full overflow-hidden mt-2">
        <div
          className={`h-full rounded-full transition-all duration-700 ${
            status === "pass"
              ? "bg-emerald-500"
              : status === "warning"
                ? "bg-amber-500"
                : "bg-rose-500"
          }`}
          style={{ width: `${Math.round((rate || 0) * 100)}%` }}
        />
      </div>
    );

    if (bdtLoading) {
      return (
        <div className="flex flex-col items-center justify-center py-40">
          <div className="relative">
            <div className="w-20 h-20 border-[6px] border-indigo-600/20 border-t-indigo-600 rounded-full animate-spin" />
            <div className="absolute inset-0 flex items-center justify-center">
              <ClipboardCheck
                className="text-indigo-500 animate-pulse"
                size={28}
              />
            </div>
          </div>
          <h3 className="text-2xl font-black mt-8 tracking-tighter">
            Validating BDT Sheets
          </h3>
          <p className="text-premium-400 mt-2 font-bold uppercase tracking-widest text-[10px]">
            Analyzing structure and data integrity
          </p>
        </div>
      );
    }

    if (bdtResults.length === 0) {
      return (
        <div className="max-w-4xl mx-auto py-12 animate-slide-up">
          <header className="text-center mb-16">
            <div className="w-24 h-24 bg-gradient-to-br from-indigo-600 to-purple-600 text-white rounded-[2rem] flex items-center justify-center mx-auto shadow-2xl shadow-indigo-500/30 mb-8 hover:rotate-6 transition-transform duration-500">
              <ClipboardCheck size={48} />
            </div>
            <h2 className="text-5xl font-black tracking-tighter mb-4">
              BDT Validator
            </h2>
            <p className="text-premium-500 font-medium text-lg max-w-2xl mx-auto">
              Upload Battery Discharge Test workbooks to validate completeness,
              photo documentation, and measurement integrity.
            </p>
          </header>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
            {/* BDT Files Section */}
            <div
              className={`group relative border-2 border-dashed rounded-[3rem] p-12 transition-all duration-500 bg-white dark:bg-premium-800 shadow-2xl ${
                bdtDragActive
                  ? "border-indigo-500 bg-indigo-50/5 ring-4 ring-indigo-500/10"
                  : "border-premium-200 dark:border-premium-700 hover:border-indigo-400 hover:scale-[1.02]"
              }`}
              onDragOver={(e) => {
                e.preventDefault();
                setBdtDragActive(true);
              }}
              onDragLeave={() => setBdtDragActive(false)}
              onDrop={(e) => {
                e.preventDefault();
                setBdtDragActive(false);
                handleBdtUpload(e.dataTransfer.files);
              }}
            >
              <input
                type="file"
                accept=".xlsx,.xls"
                multiple
                onChange={(e) => handleBdtUpload(e.target.files)}
                className="absolute inset-0 opacity-0 cursor-pointer z-20"
              />
              <div className="flex flex-col items-center pointer-events-none group-hover:scale-105 transition-transform">
                <div className="w-16 h-16 bg-indigo-50 dark:bg-indigo-900/20 rounded-full flex items-center justify-center mb-6 text-indigo-600 shadow-inner">
                  <FileSpreadsheet size={32} />
                </div>
                <p className="text-xl font-black mb-2">BDT Workbooks</p>
                <p className="text-[9px] font-black text-premium-400 uppercase tracking-[0.2em] text-center">
                  Drop BDT files here <br /> (Required)
                </p>
              </div>
            </div>

            {/* Summary File Section */}
            <div
              className={`group relative border-2 border-dashed rounded-[3rem] p-12 transition-all duration-500 bg-white dark:bg-premium-800 shadow-2xl ${
                bdtSummaryDragActive
                  ? "border-emerald-500 bg-emerald-50/5 ring-4 ring-emerald-500/10"
                  : "border-premium-200 dark:border-premium-700 hover:border-emerald-400 hover:scale-[1.02]"
              }`}
              onDragOver={(e) => {
                e.preventDefault();
                setBdtSummaryDragActive(true);
              }}
              onDragLeave={() => setBdtSummaryDragActive(false)}
              onDrop={(e) => {
                e.preventDefault();
                setBdtSummaryDragActive(false);
                handleBdtSummaryUpload(e.dataTransfer.files);
              }}
            >
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => handleBdtSummaryUpload(e.target.files)}
                className="absolute inset-0 opacity-0 cursor-pointer z-20"
              />
              <div className="flex flex-col items-center pointer-events-none group-hover:scale-105 transition-transform">
                <div className="w-16 h-16 bg-emerald-50 dark:bg-emerald-900/20 rounded-full flex items-center justify-center mb-6 text-emerald-600 shadow-inner">
                  <Layers size={32} />
                </div>
                <p className="text-xl font-black mb-2">Current Summary</p>
                {bdtSummary ? (
                  <div className="text-center">
                    <p className="text-[10px] font-black text-emerald-600 uppercase mb-1">
                      {bdtSummary.fileName}
                    </p>
                    <p className="text-[8px] font-bold text-premium-400 uppercase tracking-widest">
                      {bdtSummary.rowCount} entries
                    </p>
                  </div>
                ) : (
                  <p className="text-[9px] font-black text-premium-400 uppercase tracking-[0.2em] text-center">
                    Drop Summary file here
                  </p>
                )}
              </div>
            </div>

            <div
              className={`group relative border-2 border-dashed rounded-[3rem] p-12 transition-all duration-500 bg-white dark:bg-premium-800 shadow-2xl ${
                oldSummaryDragActive
                  ? "border-amber-500 bg-amber-50/5 ring-4 ring-amber-500/10"
                  : "border-premium-200 dark:border-premium-700 hover:border-amber-400 hover:scale-[1.02]"
              }`}
              onDragOver={(e) => {
                e.preventDefault();
                setOldSummaryDragActive(true);
              }}
              onDragLeave={() => setOldSummaryDragActive(false)}
              onDrop={(e) => {
                e.preventDefault();
                setOldSummaryDragActive(false);
                handleOldSummaryUpload(e.dataTransfer.files);
              }}
            >
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => handleOldSummaryUpload(e.target.files)}
                className="absolute inset-0 opacity-0 cursor-pointer z-20"
              />
              <div className="flex flex-col items-center pointer-events-none group-hover:scale-105 transition-transform">
                <div className="w-16 h-16 bg-amber-50 dark:bg-amber-900/20 rounded-full flex items-center justify-center mb-6 text-amber-600 shadow-inner">
                  <History size={32} />
                </div>
                <p className="text-xl font-black mb-2">Old Summary</p>
                {oldSummary ? (
                  <div className="text-center">
                    <p className="text-[10px] font-black text-amber-600 uppercase mb-1">
                      {oldSummary.fileName}
                    </p>
                    <p className="text-[8px] font-bold text-premium-400 uppercase tracking-widest">
                      {oldSummary.rowCount} entries
                    </p>
                  </div>
                ) : (
                  <p className="text-[9px] font-black text-premium-400 uppercase tracking-[0.2em] text-center">
                    Drop Old Summary here
                  </p>
                )}
              </div>
            </div>
          </div>

          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-4 mt-10">
            {[
              {
                label: "General Data",
                // range: "A4:Q6",
                icon: <FileSpreadsheet size={18} />,
                cls: "text-blue-500 bg-blue-500/10",
              },
              {
                label: "Photos",
                // range: "M9:AB59",
                icon: <Camera size={18} />,
                cls: "text-purple-500 bg-purple-500/10",
              },
              {
                label: "Basic Data",
                // range: "I11:I71",
                icon: <Activity size={18} />,
                cls: "text-amber-500 bg-amber-500/10",
              },
              {
                label: "Busbar",
                // range: "A74:U112",
                icon: <Zap size={18} />,
                cls: "text-rose-500 bg-rose-500/10",
              },
              {
                label: "Cell Readings",
                // range: "A139:AM163",
                icon: <Battery size={18} />,
                cls: "text-emerald-500 bg-emerald-500/10",
              },
            ].map((s, i) => (
              <div
                key={i}
                className="p-5 bg-white dark:bg-premium-900 rounded-2xl border border-premium-100 dark:border-premium-800 text-center hover:shadow-lg transition-shadow"
              >
                <div
                  className={`w-10 h-10 rounded-xl flex items-center justify-center mx-auto mb-3 ${s.cls}`}
                >
                  {s.icon}
                </div>
                <div className="font-black text-sm mb-1">{s.label}</div>
                <div className="text-[9px] font-bold text-premium-400 font-mono">
                  {s.range}
                </div>
              </div>
            ))}
          </div>
        </div>
      );
    }

    const allSheets = bdtResults.flatMap((r) => r.sheets);
    const passCount = allSheets.filter(
      (s) => (s.manualOverride || s.overallStatus) === "pass",
    ).length;
    const warnCount = allSheets.filter(
      (s) => !s.manualOverride && s.overallStatus === "warning",
    ).length;
    const failCount = allSheets.filter(
      (s) => (s.manualOverride || s.overallStatus) === "fail",
    ).length;

    return (
      <div className="animate-slide-up space-y-10 pb-32">
        {/* Header */}
        <div className="flex flex-col md:flex-row md:items-end justify-between gap-8">
          <div>
            <div className="flex items-center gap-3 mb-4">
              <ClipboardCheck size={16} className="text-indigo-500" />
              <span className="text-[10px] font-black uppercase tracking-[0.3em] text-indigo-500">
                BDT Validation Report
              </span>
            </div>
            <h1 className="text-5xl font-black tracking-tight leading-none">
              Audit Results
            </h1>
            <p className="text-premium-400 mt-3 font-medium">
              {bdtResults.length} file{bdtResults.length > 1 ? "s" : ""} ·{" "}
              {allSheets.length} BDT sheet{allSheets.length !== 1 ? "s" : ""}{" "}
              analysed
              {reviewedBdtSheets.size > 0 && (
                <span className="ml-2 inline-flex items-center gap-1.5 px-3 py-0.5 bg-sky-500/10 text-sky-500 rounded-full text-xs font-black">
                  <Eye size={12} />
                  {reviewedBdtSheets.size}/{allSheets.length} reviewed
                </span>
              )}
            </p>
            <div className="flex gap-3 mt-6">
              <button
                onClick={() => setActiveTab("settings")}
                className="flex items-center gap-2 px-6 py-2.5 bg-indigo-500/10 text-indigo-600 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-indigo-500 hover:text-white transition-all"
              >
                <Settings size={14} /> Rules Settings
              </button>
              <button
                onClick={() => {
                  setBdtResults([]);
                  setExpandedBdtSheet(null);
                  setReviewedBdtSheets(new Set());
                }}
                className="flex items-center gap-2 px-6 py-2.5 bg-rose-500/10 text-rose-600 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-rose-500 hover:text-white transition-all"
              >
                <RotateCcw size={14} /> New Audit
              </button>
              <label className="relative cursor-pointer">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  multiple
                  className="absolute inset-0 opacity-0 cursor-pointer"
                  onChange={(e) => handleBdtUpload(e.target.files)}
                />
                <div className="flex items-center gap-2 px-6 py-2.5 bg-indigo-500/10 text-indigo-600 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-indigo-500 hover:text-white transition-all">
                  <Plus size={14} /> Add BDT Files
                </div>
              </label>
              <label className="relative cursor-pointer">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  className="absolute inset-0 opacity-0 cursor-pointer"
                  onChange={(e) => handleBdtSummaryUpload(e.target.files)}
                />
                <div className="flex items-center gap-2 px-6 py-2.5 bg-emerald-500/10 text-emerald-600 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-emerald-500 hover:text-white transition-all">
                  <Layers size={14} />{" "}
                  {bdtSummary ? "Change Summary" : "Add Summary"}
                </div>
              </label>
              <label className="relative cursor-pointer">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  className="absolute inset-0 opacity-0 cursor-pointer"
                  onChange={(e) => handleGenerateReport(e.target.files[0])}
                  disabled={bdtResults.length === 0 || reportLoading}
                />
                <div
                  className={`flex items-center gap-2 px-6 py-2.5 bg-emerald-500/10 text-emerald-600 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-emerald-500 hover:text-white transition-all ${bdtResults.length === 0 ? "opacity-50 cursor-not-allowed" : ""}`}
                >
                  {reportLoading ? (
                    <RefreshCw size={14} className="animate-spin" />
                  ) : (
                    <Download size={14} />
                  )}
                  Generate BDT Report
                </div>
              </label>
            </div>
          </div>

          <div className="flex gap-4">
            {[
              {
                label: "Passed",
                count: passCount,
                filterKey: "pass",
                activeCls:
                  "bg-emerald-500 border-emerald-500 text-white shadow-xl shadow-emerald-500/30 scale-105",
                inactiveCls:
                  "bg-emerald-500/10 border-emerald-500/30 text-emerald-600 dark:text-emerald-400 hover:scale-105 hover:bg-emerald-500/20",
              },
              {
                label: "Warnings",
                count: warnCount,
                filterKey: "warning",
                activeCls:
                  "bg-amber-500 border-amber-500 text-white shadow-xl shadow-amber-500/30 scale-105",
                inactiveCls:
                  "bg-amber-500/10 border-amber-500/30 text-amber-600 dark:text-amber-400 hover:scale-105 hover:bg-amber-500/20",
              },
              {
                label: "Failed",
                count: failCount,
                filterKey: "fail",
                activeCls:
                  "bg-rose-500 border-rose-500 text-white shadow-xl shadow-rose-500/30 scale-105",
                inactiveCls:
                  "bg-rose-500/10 border-rose-500/30 text-rose-600 dark:text-rose-400 hover:scale-105 hover:bg-rose-500/20",
              },
            ].map((kpi, i) => (
              <button
                key={i}
                onClick={() =>
                  setBdtStatusFilter(
                    bdtStatusFilter === kpi.filterKey ? null : kpi.filterKey,
                  )
                }
                className={`px-8 py-6 rounded-2xl border text-center min-w-[100px] transition-all duration-200 cursor-pointer ${
                  bdtStatusFilter === kpi.filterKey
                    ? kpi.activeCls
                    : kpi.inactiveCls
                }`}
              >
                <div className="text-3xl font-black">{kpi.count}</div>
                <div className="text-[10px] font-black uppercase tracking-widest mt-1 opacity-80">
                  {kpi.label}
                </div>
                {bdtStatusFilter === kpi.filterKey && (
                  <div className="text-[8px] font-black uppercase tracking-widest mt-1 opacity-70">
                    Click to clear
                  </div>
                )}
              </button>
            ))}
          </div>
        </div>

        {/* BDT Charts Section */}
        {allSheets.length > 0 && (
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 animate-fade-in">
            <Card className="p-6">
              <h3 className="text-sm font-black uppercase tracking-widest text-premium-400 mb-6 flex items-center gap-2">
                <CheckCircle2 size={16} /> Acceptance Rate
              </h3>
              <div className="h-[250px]">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={[
                        {
                          name: "Accepted",
                          value: passCount + warnCount,
                          color: "#10b981",
                        },
                        {
                          name: "Rejected",
                          value: failCount,
                          color: "#f43f5e",
                        },
                      ]}
                      dataKey="value"
                      nameKey="name"
                      cx="50%"
                      cy="50%"
                      innerRadius={60}
                      outerRadius={80}
                      paddingAngle={5}
                    >
                      {[
                        {
                          name: "Accepted",
                          value: passCount + warnCount,
                          color: "#10b981",
                        },
                        {
                          name: "Rejected",
                          value: failCount,
                          color: "#f43f5e",
                        },
                      ].map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={entry.color} />
                      ))}
                    </Pie>
                    <Tooltip
                      contentStyle={{
                        borderRadius: "12px",
                        border: "1px solid #334155",
                        backgroundColor: "#0f172a",
                        color: "#fff",
                      }}
                      itemStyle={{ color: "#fff", fontWeight: "bold" }}
                    />
                    <Legend
                      verticalAlign="bottom"
                      height={36}
                      wrapperStyle={{
                        fontSize: "12px",
                        fontWeight: "bold",
                        textTransform: "uppercase",
                        tracking: "0.1em",
                      }}
                    />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </Card>
            <Card className="p-6">
              <h3 className="text-sm font-black uppercase tracking-widest text-premium-400 mb-6 flex items-center gap-2">
                <BarChart3 size={16} /> Top Rejection Reasons
              </h3>
              <div className="h-[250px]">
                {(() => {
                  const issueCounts = {};
                  allSheets.forEach((sheet) => {
                    const status = sheet.manualOverride || sheet.overallStatus;
                    if (status === "fail") {
                      const uniqueIssues = Array.from(
                        new Set(sheet.allIssues.map(getShortIssueText)),
                      );
                      uniqueIssues.forEach((issue) => {
                        issueCounts[issue] = (issueCounts[issue] || 0) + 1;
                      });
                    }
                  });

                  const commonIssuesData = Object.keys(issueCounts)
                    .map((key) => ({ name: key, count: issueCounts[key] }))
                    .sort((a, b) => b.count - a.count)
                    .slice(0, 5);

                  if (commonIssuesData.length === 0) {
                    return (
                      <div className="flex items-center justify-center h-full text-premium-400 text-sm font-black uppercase tracking-widest">
                        No rejection reasons
                      </div>
                    );
                  }

                  return (
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart
                        data={commonIssuesData}
                        layout="vertical"
                        margin={{ top: 0, right: 20, left: 0, bottom: 0 }}
                      >
                        <XAxis type="number" hide />
                        <YAxis
                          type="category"
                          dataKey="name"
                          width={180}
                          tick={{
                            fontSize: 10,
                            fill: "#94a3b8",
                            fontWeight: "bold",
                          }}
                          tickLine={false}
                          axisLine={false}
                        />
                        <Tooltip
                          cursor={{ fill: "rgba(255,255,255,0.05)" }}
                          contentStyle={{
                            borderRadius: "12px",
                            border: "1px solid #334155",
                            backgroundColor: "#0f172a",
                            color: "#fff",
                          }}
                          itemStyle={{ color: "#fff", fontWeight: "bold" }}
                        />
                        <Bar
                          dataKey="count"
                          fill="#8b5cf6"
                          radius={[0, 4, 4, 0]}
                          barSize={20}
                        >
                          {commonIssuesData.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill="#8b5cf6" />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  );
                })()}
              </div>
            </Card>
          </div>
        )}

        {/* Active filter indicator */}
        {bdtStatusFilter && (
          <div className="flex items-center gap-3 animate-fade-in">
            <div
              className={`flex items-center gap-2 px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest border ${
                bdtStatusFilter === "pass"
                  ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-600"
                  : bdtStatusFilter === "warning"
                    ? "bg-amber-500/10 border-amber-500/30 text-amber-600"
                    : "bg-rose-500/10 border-rose-500/30 text-rose-600"
              }`}
            >
              <Filter size={12} />
              Showing:{" "}
              {bdtStatusFilter === "pass"
                ? "Passed"
                : bdtStatusFilter === "warning"
                  ? "Warnings"
                  : "Failed"}{" "}
              only
            </div>
            <button
              onClick={() => setBdtStatusFilter(null)}
              className="text-[10px] font-black text-premium-400 hover:text-premium-700 uppercase tracking-widest transition-colors"
            >
              Show all
            </button>
          </div>
        )}

        {/* Summary Sheet display (if uploaded separately) */}
        {bdtSummary && (
          <div className="flex items-center justify-between p-6 bg-emerald-500/5 border border-emerald-500/20 rounded-[2rem] animate-fade-in">
            <div className="flex items-center gap-6">
              <div className="w-14 h-14 bg-emerald-500 text-white rounded-2xl flex items-center justify-center shadow-lg shadow-emerald-500/20">
                <Layers size={28} />
              </div>
              <div>
                <div className="text-[10px] font-black text-emerald-600 uppercase tracking-widest mb-1">
                  Master Summary Protocol
                </div>
                <h3 className="font-black text-xl leading-none">
                  {bdtSummary.fileName}
                </h3>
                <div className="text-[10px] font-bold text-premium-400 mt-2 flex items-center gap-3">
                  <span className="bg-emerald-500/10 text-emerald-600 px-2 py-0.5 rounded uppercase">
                    {bdtSummary.rowCount} Records Detected
                  </span>
                  <span className="bg-blue-500/10 text-blue-600 px-2 py-0.5 rounded uppercase">
                    {bdtSummary.sheetName}
                  </span>
                </div>
              </div>
            </div>
            <button
              onClick={() => setBdtSummary(null)}
              className="px-4 py-2 text-[10px] font-black text-rose-500 uppercase hover:bg-rose-500/10 rounded-lg transition-colors"
            >
              Remove
            </button>
          </div>
        )}

        {/* One block per uploaded file */}
        {bdtResults.map((result, fileIdx) => {
          const filteredSheets = bdtStatusFilter
            ? result.sheets.filter((s) => {
                const effective = s.manualOverride || s.overallStatus;
                if (bdtStatusFilter === "pass") return effective === "pass";
                if (bdtStatusFilter === "warning")
                  return !s.manualOverride && s.overallStatus === "warning";
                if (bdtStatusFilter === "fail") return effective === "fail";
                return true;
              })
            : result.sheets;

          if (filteredSheets.length === 0) return null;

          return (
            <div key={fileIdx} className="space-y-4">
              <div className="flex items-center gap-4 pb-4 border-b border-premium-100 dark:border-premium-800">
                <FileSpreadsheet size={18} className="text-indigo-500" />
                <h3 className="font-black text-lg">{result.fileName}</h3>
                <span className="text-[10px] font-black text-premium-400 uppercase tracking-widest">
                  {filteredSheets.length}
                  {bdtStatusFilter ? ` of ${result.sheets.length}` : ""} sheet
                  {filteredSheets.length !== 1 ? "s" : ""}
                </span>
              </div>

              {filteredSheets.map((sheet, sheetIdx) => {
                const sheetKey = `${fileIdx}-${sheetIdx}`;
                const isExpanded = expandedBdtSheet === sheetKey;
                const sectionDefs = [
                  {
                    data: sheet.generalData,
                    icon: <FileSpreadsheet size={16} />,
                    iconCls: "text-blue-500 bg-blue-500/10",
                    stat: `${sheet.generalData.filledCells}/${sheet.generalData.totalCells} cells`,
                  },
                  {
                    data: sheet.photos,
                    icon: <Camera size={16} />,
                    iconCls: "text-purple-500 bg-purple-500/10",
                    stat: `~${sheet.photos.photoCount} photo${sheet.photos.photoCount !== 1 ? "s" : ""}${
                      sheet.photos.detectionMethod !== "direct" ? " (est.)" : ""
                    }`,
                  },
                  {
                    data: sheet.basicData,
                    icon: <Activity size={16} />,
                    iconCls: "text-amber-500 bg-amber-500/10",
                    stat: `${sheet.basicData.filledCells}/${sheet.basicData.totalCells} cells`,
                  },
                  {
                    data: sheet.busbarReadings,
                    icon: <Zap size={16} />,
                    iconCls: "text-rose-500 bg-rose-500/10",
                    stat: sheet.busbarReadings.finalDurationFormatted
                      ? `Final: ${sheet.busbarReadings.finalDurationFormatted}`
                      : `${sheet.busbarReadings.durationReadings} readings`,
                  },
                  {
                    data: sheet.cellsReadings,
                    icon: <Battery size={16} />,
                    iconCls: "text-emerald-500 bg-emerald-500/10",
                    stat: `${sheet.cellsReadings.filledCells}/${sheet.cellsReadings.totalCells} filled`,
                  },
                ];

                return (
                  <Card key={sheetIdx} className="overflow-hidden">
                    <button
                      className="w-full p-6 flex items-center justify-between hover:bg-premium-50/50 dark:hover:bg-white/5 transition-all text-left"
                      onClick={() => {
                        const nextExpanded = isExpanded ? null : sheetKey;
                        setExpandedBdtSheet(nextExpanded);
                        if (nextExpanded && !reviewedBdtSheets.has(sheetKey)) {
                          setReviewedBdtSheets(
                            (prev) => new Set([...prev, sheetKey]),
                          );
                        }
                      }}
                    >
                      <div className="flex items-center gap-5">
                        {statusIcon(
                          sheet.manualOverride || sheet.overallStatus,
                          24,
                        )}
                        <div>
                          <div className="font-black text-base leading-none flex items-center gap-2">
                            {sheet.sheetName}
                            {reviewedBdtSheets.has(sheetKey) ? (
                              <span
                                className="inline-flex items-center gap-1 px-2 py-0.5 bg-sky-500/10 text-sky-400 rounded-md text-[8px] font-black uppercase tracking-widest"
                                title="Reviewed"
                              >
                                <Eye size={10} /> Reviewed
                              </span>
                            ) : (
                              <span
                                className="inline-flex items-center gap-1 px-2 py-0.5 bg-orange-500/10 text-orange-400 rounded-md text-[8px] font-black uppercase tracking-widest animate-pulse"
                                title="Not yet reviewed"
                              >
                                <EyeOff size={10} /> Unreviewed
                              </span>
                            )}
                          </div>
                          {sheet.siteName && (
                            <div className="text-[10px] font-bold text-premium-400 uppercase mt-1">
                              {sheet.siteName}
                            </div>
                          )}
                          {sheet.testDate && (
                            <div className="text-[9px] font-bold text-premium-500 mt-0.5">
                              {sheet.testDate}{" "}
                              {sheet.startHour !== null &&
                                `(Started: ${sheet.startHour}:00)`}
                            </div>
                          )}
                        </div>
                        {sheet.issueCount > 0 && (
                          <span className="px-3 py-1 bg-rose-500/10 text-rose-500 rounded-lg text-[10px] font-black uppercase tracking-widest">
                            {sheet.issueCount} issue
                            {sheet.issueCount > 1 ? "s" : ""}
                          </span>
                        )}
                        {sheet.busbarReadings.testedBackup && (
                          <div className="flex flex-col items-center px-4 border-l border-premium-100 dark:border-premium-800 ml-4">
                            <span className="text-[9px] font-black text-premium-400 uppercase tracking-widest leading-none mb-1">
                              Tested Backup
                            </span>
                            <span className="text-xs font-black text-blue-600 dark:text-blue-400">
                              {sheet.busbarReadings.testedBackup}
                            </span>
                          </div>
                        )}
                        <div className="flex flex-col items-center px-4 border-l border-premium-100 dark:border-premium-800">
                          <span className="text-[9px] font-black text-premium-400 uppercase tracking-widest leading-none mb-1">
                            Status
                          </span>
                          <span
                            className={`text-[10px] font-black uppercase tracking-widest ${
                              (sheet.manualOverride || sheet.overallStatus) ===
                              "fail"
                                ? "text-rose-500"
                                : (sheet.manualOverride ||
                                      sheet.overallStatus) === "warning"
                                  ? "text-amber-500"
                                  : "text-emerald-500"
                            }`}
                          >
                            {sheet.manualOverride
                              ? sheet.manualOverride === "pass"
                                ? "Accepted (M)"
                                : "Rejected (M)"
                              : sheet.overallStatus === "pass"
                                ? "Accepted"
                                : sheet.overallStatus === "warning"
                                  ? "Accepted (W)"
                                  : "Rejected"}
                          </span>
                        </div>
                      </div>
                      <div className="flex items-center gap-6">
                        {/* Mini section status dots */}
                        <div className="hidden md:flex items-center gap-2">
                          {sectionDefs.map((sec, i) => (
                            <div
                              key={i}
                              title={sec.data.section}
                              className={`w-3 h-3 rounded-full ${
                                sec.data.status === "pass"
                                  ? "bg-emerald-500"
                                  : sec.data.status === "warning"
                                    ? "bg-amber-500"
                                    : sec.data.status === "fail"
                                      ? "bg-rose-500"
                                      : "bg-premium-400"
                              }`}
                            />
                          ))}
                        </div>
                        <ChevronDown
                          size={18}
                          className={`text-premium-400 transition-transform duration-300 ${isExpanded ? "rotate-180" : ""}`}
                        />
                      </div>
                    </button>

                    {isExpanded && (
                      <div className="border-t border-premium-100 dark:border-premium-800 p-8 space-y-8 animate-fade-in">
                        {/* BDT KPI Summary Row */}
                        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                          <div className="bg-blue-600/5 border border-blue-600/20 rounded-3xl p-6 flex items-center justify-between">
                            <div className="flex items-center gap-5">
                              <div className="w-12 h-12 bg-blue-600 text-white rounded-2xl flex items-center justify-center shadow-lg shadow-blue-500/20">
                                <History size={24} />
                              </div>
                              <div>
                                <p className="text-[10px] font-black text-blue-600/60 uppercase tracking-[0.2em] mb-1">
                                  Tested Backup Time
                                </p>
                                <h4 className="text-3xl font-black text-blue-600 dark:text-blue-400 leading-none">
                                  {sheet.busbarReadings.testedBackup || "N/A"}
                                </h4>
                              </div>
                            </div>
                          </div>

                          {sheet.busbarReadings.theoreticalBackup > 0 && (
                            <div className="bg-purple-600/5 border border-purple-600/20 rounded-3xl p-6 flex items-center justify-between">
                              <div className="flex items-center gap-5">
                                <div className="w-12 h-12 bg-purple-600 text-white rounded-2xl flex items-center justify-center shadow-lg shadow-purple-500/20">
                                  <Zap size={24} />
                                </div>
                                <div>
                                  <p className="text-[10px] font-black text-purple-600/60 uppercase tracking-[0.2em] mb-1">
                                    Theoretical Backup
                                  </p>
                                  <h4 className="text-3xl font-black text-purple-600 dark:text-purple-400 leading-none">
                                    {sheet.busbarReadings.theoreticalBackup} min
                                  </h4>
                                </div>
                              </div>
                            </div>
                          )}

                          <div
                            className={`${(sheet.manualOverride || sheet.overallStatus) !== "fail" ? "bg-emerald-500/5 border-emerald-500/20" : "bg-rose-500/5 border-rose-500/20"} rounded-3xl p-6 flex flex-col justify-center relative group`}
                          >
                            <div className="flex items-center justify-between">
                              <div className="flex items-center gap-5">
                                <div
                                  className={`w-12 h-12 ${(sheet.manualOverride || sheet.busbarReadings.status) !== "fail" ? "bg-emerald-500" : "bg-rose-500"} text-white rounded-2xl flex items-center justify-center shadow-lg`}
                                >
                                  {(sheet.manualOverride ||
                                    sheet.busbarReadings.status) !== "fail" ? (
                                    <ClipboardCheck size={24} />
                                  ) : (
                                    <XCircle size={24} />
                                  )}
                                </div>
                                <div>
                                  <p
                                    className={`text-[10px] font-black opacity-60 uppercase tracking-[0.2em] mb-1 ${(sheet.manualOverride || sheet.overallStatus) !== "fail" ? "text-emerald-500" : "text-rose-500"}`}
                                  >
                                    Audit Decision{" "}
                                    {sheet.manualOverride && "(Manual)"}
                                  </p>
                                  <h4
                                    className={`text-3xl font-black leading-none ${(sheet.manualOverride || sheet.overallStatus) !== "fail" ? "text-emerald-500" : "text-rose-500"}`}
                                  >
                                    {(sheet.manualOverride ||
                                      sheet.overallStatus) !== "fail"
                                      ? "ACCEPTED"
                                      : "REJECTED"}
                                  </h4>
                                </div>
                              </div>
                              <div className="flex flex-col gap-2 relative z-10">
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleManualOverride(
                                      fileIdx,
                                      sheetIdx,
                                      "pass",
                                    );
                                  }}
                                  className={`w-24 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-wider transition-all border ${sheet.manualOverride === "pass" ? "bg-emerald-500 text-white border-emerald-500" : "bg-transparent text-emerald-600 border-emerald-500/30 hover:bg-emerald-500/10"}`}
                                >
                                  Force Accept
                                </button>
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleManualOverride(
                                      fileIdx,
                                      sheetIdx,
                                      "fail",
                                    );
                                  }}
                                  className={`w-24 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-wider transition-all border ${sheet.manualOverride === "fail" ? "bg-rose-500 text-white border-rose-500" : "bg-transparent text-rose-600 border-rose-500/30 hover:bg-rose-500/10"}`}
                                >
                                  Force Reject
                                </button>
                              </div>
                            </div>
                          </div>
                        </div>

                        {/* 5 section cards */}
                        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-5 gap-4">
                          {sectionDefs.map((sec, i) => (
                            <div
                              key={i}
                              className={`p-5 rounded-2xl border ${statusBg(sec.data.status)}`}
                            >
                              <div className="flex items-center justify-between mb-4">
                                <div
                                  className={`w-8 h-8 rounded-xl flex items-center justify-center ${sec.iconCls}`}
                                >
                                  {sec.icon}
                                </div>
                                {statusIcon(sec.data.status, 14)}
                              </div>
                              <div className="font-black text-sm leading-tight mb-0.5">
                                {sec.data.section}
                              </div>
                              <div className="text-[9px] font-bold text-premium-400 font-mono mb-2">
                                {sec.data.range}
                              </div>
                              <div className="text-[10px] font-bold text-premium-500">
                                {sec.stat}
                              </div>
                              {"fillRate" in sec.data &&
                                sec.data.fillRate !== undefined &&
                                fillBar(sec.data.fillRate, sec.data.status)}
                            </div>
                          ))}
                        </div>

                        {/* Photo Inspector (OCR) */}
                        {sheet.photos?.images?.length > 0 && (
                          <BDTPhotoInspector images={sheet.photos.images} />
                        )}

                        {/* Issues list */}
                        {sheet.allIssues.length > 0 ? (
                          <div className="bg-rose-500/5 border border-rose-500/20 rounded-2xl p-6">
                            <h5 className="text-[10px] font-black uppercase tracking-widest text-rose-500 mb-4 flex items-center gap-2">
                              <AlertTriangle size={12} /> Issues Found (
                              {sheet.allIssues.length})
                            </h5>
                            <div className="space-y-2">
                              {sheet.allIssues.map((issue, i) => (
                                <div key={i} className="flex items-start gap-3">
                                  <span className="text-[9px] font-black bg-rose-500/10 text-rose-400 px-2 py-0.5 rounded uppercase tracking-wider mt-0.5 shrink-0 whitespace-nowrap">
                                    {issue.section}
                                  </span>
                                  <span className="text-xs font-medium text-premium-500">
                                    {issue.text}
                                  </span>
                                </div>
                              ))}
                            </div>
                            {sheet.photos.note && (
                              <div className="mt-4 flex items-center gap-2 text-[10px] font-bold text-amber-500 border-t border-amber-500/20 pt-4">
                                <AlertTriangle size={10} />
                                <span>{sheet.photos.note}</span>
                              </div>
                            )}
                          </div>
                        ) : (
                          <div className="flex items-center gap-3 p-5 bg-emerald-500/10 border border-emerald-500/20 rounded-2xl text-emerald-600 dark:text-emerald-400">
                            <CheckCircle2 size={18} />
                            <span className="text-sm font-black uppercase tracking-widest">
                              All sections validated successfully
                            </span>
                          </div>
                        )}
                      </div>
                    )}
                  </Card>
                );
              })}
            </div>
          );
        })}
      </div>
    );
  };

  // BDT Settings Modal removed — settings are now in the Console tab

  if (trialLockEnabled && !isActivated) {
    return (
      <TrialActivation
        onActivate={() => {
          setIsActivated(true);
          localStorage.setItem("energy_review_activated", JSON.stringify(true));
        }}
        trialLockEnabled={trialLockEnabled}
        setTrialLockEnabled={setTrialLockEnabled}
      />
    );
  }

  return (
    <div className="min-h-screen bg-premium-50 dark:bg-premium-950 text-premium-900 dark:text-premium-50 selection:bg-blue-500/30 overflow-x-hidden">
      <header className="sticky top-0 z-50 glass border-b border-premium-100 dark:border-premium-900">
        <div className="max-w-[1600px] mx-auto px-8 lg:px-12 h-24 flex items-center justify-between">
          <div className="flex items-center gap-6">
            <div className="w-14 h-14 bg-premium-950 dark:bg-white rounded-[1.2rem] flex items-center justify-center shadow-premium shadow-xl">
              <BarChart3
                className="text-white dark:text-premium-950"
                size={28}
              />
            </div>
            <div>
              <h2 className="text-2xl font-black tracking-tighter leading-none mb-1.5 flex items-center gap-2">
                NET <span className="text-blue-600">Energizer</span>
              </h2>
              <div className="flex items-center gap-2">
                <ShieldCheck size={12} className="text-emerald-500" />
                <span className="text-[9px] font-black tracking-[0.3em] text-premium-400 uppercase">
                  By Mobi Intelligence v1.0
                </span>
              </div>
            </div>
          </div>
          <nav className="hidden lg:flex items-center gap-2 p-1.5 bg-premium-100 dark:bg-premium-900/40 rounded-2xl backdrop-blur-sm">
            {[
              {
                id: "dashboard",
                icon: <LayoutDashboard size={18} />,
                label: "Analytics",
              },
              {
                id: "bdt",
                icon: <ClipboardCheck size={18} />,
                label: "BDT Audit",
              },
              { id: "memory", icon: <Database size={18} />, label: "Memory" },
              {
                id: "settings",
                icon: <Settings size={18} />,
                label: "Console",
              },
            ].map((tab) => (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id)}
                className={`flex items-center gap-3 px-6 py-3 rounded-xl text-xs font-black transition-all duration-300 ${activeTab === tab.id ? "bg-white dark:bg-premium-800 text-blue-600 shadow-xl" : "text-premium-400 hover:text-premium-600 hover:scale-105"}`}
              >
                {tab.icon}
                <span className="uppercase tracking-widest">{tab.label}</span>
              </button>
            ))}
          </nav>
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-4 px-6 py-3 bg-white dark:bg-premium-900 border border-premium-200 dark:border-premium-800 rounded-2xl shadow-sm">
              <div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></div>
              <div className="flex flex-col">
                <span className="text-[10px] font-black text-premium-400 uppercase leading-none mb-0.5">
                  Database Flow
                </span>
                <span className="text-xs font-black text-premium-900 dark:text-white leading-none">
                  {siteList.length} Connected
                </span>
              </div>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-[1600px] mx-auto px-8 lg:px-12 pt-12">
        {selectedSite && renderSiteSpotlight()}
        {loading ? (
          <div className="fixed inset-0 z-[100] bg-premium-950/80 backdrop-blur-xl flex flex-col items-center justify-center animate-fade-in">
            <div className="relative">
              <div className="w-24 h-24 border-[6px] border-blue-600/20 border-t-blue-600 rounded-full animate-spin"></div>
              <div className="absolute inset-0 flex items-center justify-center">
                <Activity className="text-blue-500 animate-pulse" size={32} />
              </div>
            </div>
            <p className="mt-8 text-[10px] font-black text-white uppercase tracking-[0.5em] animate-pulse">
              Running Deep Analytics Engine...
            </p>
          </div>
        ) : activeTab === "dashboard" ? (
          history.length > 0 ? (
            renderDashboard()
          ) : (
            <div className="max-w-6xl mx-auto py-12 animate-slide-up">
              <header className="text-center mb-16">
                <h3 className="text-5xl font-black mb-4 tracking-tighter">
                  Energy Data Payloads
                </h3>
                <p className="text-premium-500 max-w-2xl mx-auto font-medium text-lg leading-relaxed">
                  Inject your performance data to begin. Choose your ingestion
                  protocol below.
                </p>
                <div className="flex bg-premium-100 dark:bg-premium-900/50 p-1.5 rounded-2xl mt-12 w-fit mx-auto shadow-inner border border-white/10">
                  <button
                    onClick={() => setDataSourceMode("energy")}
                    className={`px-10 py-4 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all duration-300 ${dataSourceMode === "energy" ? "bg-white dark:bg-premium-700 text-blue-600 shadow-xl scale-105" : "text-premium-400 hover:text-black dark:hover:text-white"}`}
                  >
                    Energy Master Sheet
                  </button>
                  <button
                    onClick={() => setDataSourceMode("autin")}
                    className={`px-10 py-4 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all duration-300 ${dataSourceMode === "autin" ? "bg-white dark:bg-premium-700 text-blue-600 shadow-xl scale-105" : "text-premium-400 hover:text-black dark:hover:text-white"}`}
                  >
                    AUTIN (Multi-Source)
                  </button>
                </div>
                {currentAlarms.length > 0 && (
                  <div className="mt-8 animate-fade-in">
                    <button
                      onClick={handleMemoryAnalysis}
                      className="px-8 py-3 bg-blue-600 text-white rounded-xl text-[10px] font-black uppercase tracking-widest shadow-lg shadow-blue-500/20 hover:scale-105 transition-all flex items-center gap-3 mx-auto"
                    >
                      <RotateCcw size={14} />
                      Restore Last Analysis from Memory
                    </button>
                  </div>
                )}
              </header>
              <Card className="p-16 border-dashed border-4 border-premium-200 dark:border-premium-900 bg-gradient-to-b from-transparent to-premium-50/20 relative overflow-hidden">
                <div className="relative z-10">
                  {dataSourceMode === "energy" ? (
                    <div className="space-y-12">
                      <div className="w-28 h-28 bg-gradient-to-br from-blue-600 to-blue-400 text-white rounded-[2.5rem] flex items-center justify-center mx-auto shadow-2xl shadow-blue-500/40 transform rotate-3 hover:rotate-0 transition-transform duration-500">
                        <FileSpreadsheet size={50} />
                      </div>
                      <div
                        className={`group relative border-2 border-dashed rounded-[4rem] p-32 transition-all duration-500 bg-white dark:bg-premium-800 shadow-2xl ${dragActive ? "border-blue-500 bg-blue-50 ring-4 ring-blue-500/5" : "border-premium-200 hover:border-blue-400 hover:scale-[1.02]"}`}
                        onDragOver={(e) => {
                          e.preventDefault();
                          setDragActive(true);
                        }}
                        onDragLeave={() => setDragActive(false)}
                        onDrop={(e) => {
                          e.preventDefault();
                          setDragActive(false);
                          const file = e.dataTransfer.files[0];
                          if (file) handleFileUpload(file);
                        }}
                      >
                        <input
                          type="file"
                          accept=".xlsx,.xls"
                          onChange={(e) => handleFileUpload(e.target.files[0])}
                          className="absolute inset-0 opacity-0 cursor-pointer z-20"
                        />
                        <div className="flex flex-col items-center pointer-events-none group-hover:scale-105 transition-transform">
                          <div className="w-20 h-20 bg-blue-50 dark:bg-premium-700 rounded-full flex items-center justify-center mb-8 text-blue-600 dark:text-blue-400 shadow-inner">
                            <Plus size={40} />
                          </div>
                          <p className="text-2xl font-black text-premium-900 dark:text-white mb-3">
                            Initialize Analysis
                          </p>
                          <p className="text-[10px] font-black text-premium-400 uppercase tracking-[0.3em]">
                            Drop your weekly master file here
                          </p>
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="space-y-12">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
                        {[
                          "power",
                          "down",
                          "high_temp",
                          "generator",
                          "door",
                        ].map((cat) => (
                          <div key={cat} className="space-y-4">
                            <div className="flex items-center justify-between px-2">
                              <h5 className="text-[10px] font-black uppercase tracking-widest text-premium-400 flex items-center gap-2">
                                {cat === "power" && (
                                  <Zap size={14} className="text-amber-500" />
                                )}
                                {cat === "down" && (
                                  <AlertCircle
                                    size={14}
                                    className="text-rose-500"
                                  />
                                )}
                                {cat === "high_temp" && (
                                  <Activity
                                    size={14}
                                    className="text-orange-500"
                                  />
                                )}
                                {cat === "generator" && (
                                  <Settings
                                    size={14}
                                    className="text-purple-500"
                                  />
                                )}
                                {cat === "door" && (
                                  <DoorOpen
                                    size={14}
                                    className="text-teal-500"
                                  />
                                )}
                                {cat.replace("_", " ")} source
                              </h5>
                              <div className="relative">
                                <input
                                  type="file"
                                  multiple
                                  onChange={(e) => {
                                    const files = Array.from(e.target.files);
                                    setAutinFiles((prev) => ({
                                      ...prev,
                                      [cat]: [...prev[cat], ...files],
                                    }));
                                  }}
                                  className="absolute inset-0 opacity-0 cursor-pointer w-8 h-8"
                                />
                                <div className="w-8 h-8 rounded-full bg-premium-100 dark:bg-premium-800 flex items-center justify-center text-blue-600 hover:bg-blue-600 hover:text-white transition-all">
                                  <Plus size={16} />
                                </div>
                              </div>
                            </div>
                            <div className="bg-white dark:bg-premium-800 rounded-3xl p-6 border-2 border-premium-100 dark:border-premium-900 min-h-[140px] flex flex-col justify-center gap-3">
                              {autinFiles[cat].length > 0 ? (
                                autinFiles[cat].map((f, i) => (
                                  <div
                                    key={i}
                                    className="flex items-center justify-between bg-premium-50 dark:bg-white/5 py-2 px-4 rounded-xl"
                                  >
                                    <span className="text-[10px] font-bold truncate max-w-[200px]">
                                      {f.name}
                                    </span>
                                    <button
                                      onClick={() =>
                                        setAutinFiles((prev) => ({
                                          ...prev,
                                          [cat]: prev[cat].filter(
                                            (_, idx) => idx !== i,
                                          ),
                                        }))
                                      }
                                      className="text-rose-500 hover:scale-110 transition-transform"
                                    >
                                      <AlertTriangle size={12} />
                                    </button>
                                  </div>
                                ))
                              ) : (
                                <div className="text-center py-4">
                                  <span className="text-[9px] font-black text-rose-500/50 uppercase tracking-widest leading-relaxed">
                                    No data provided.
                                    <br />
                                    Analysis will skip this category.
                                  </span>
                                </div>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                      <button
                        onClick={handleAutinAnalysis}
                        disabled={Object.values(autinFiles).every(
                          (a) => a.length === 0,
                        )}
                        className={`w-full max-w-md mx-auto flex items-center justify-center gap-4 py-6 rounded-[2rem] font-black text-sm uppercase tracking-[0.3em] shadow-2xl transition-all active:scale-95 ${Object.values(autinFiles).every((a) => a.length === 0) ? "bg-premium-100 text-premium-300 cursor-not-allowed opacity-50" : "bg-premium-950 dark:bg-white text-white dark:text-premium-950 hover:bg-blue-600 dark:hover:bg-blue-600 hover:text-white"}`}
                      >
                        <Activity size={24} /> Launch Multi-Source Audit
                      </button>
                    </div>
                  )}
                </div>
                <div className="absolute -left-20 -bottom-20 opacity-5 dark:opacity-10 pointer-events-none">
                  <Database size={300} />
                </div>
              </Card>
              <div className="mt-20 grid grid-cols-1 md:grid-cols-3 gap-10">
                {[
                  {
                    icon: <ShieldCheck className="text-emerald-500" />,
                    title: "Secure Mismatch",
                    desc: "Logic identifies BDT values falling out of safe margins.",
                  },
                  {
                    icon: <Zap className="text-amber-500" />,
                    title: "Live Sync",
                    desc: "Matches against your Alexandria Database in real-time.",
                  },
                  {
                    icon: <AlertTriangle className="text-rose-500" />,
                    title: "Risk Priority",
                    desc: "Tags sites with (1+X) weighting for escalations.",
                  },
                ].map((feature, i) => (
                  <div
                    key={i}
                    className="flex gap-6 p-8 bg-white dark:bg-premium-900 rounded-[2.5rem] shadow-sm border border-premium-100 dark:border-premium-800"
                  >
                    <div className="w-14 h-14 rounded-2xl bg-premium-50 dark:bg-white/5 flex items-center justify-center flex-shrink-0 shadow-inner">
                      {feature.icon}
                    </div>
                    <div>
                      <h4 className="font-black text-sm mb-2 uppercase tracking-tight">
                        {feature.title}
                      </h4>
                      <p className="text-xs font-medium text-premium-500 leading-relaxed">
                        {feature.desc}
                      </p>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )
        ) : activeTab === "bdt" ? (
          renderBDT()
        ) : activeTab === "memory" ? (
          renderMemory()
        ) : (
          renderSettings()
        )}
      </main>

      <footer className="mt-32 pb-12 text-center">
        <div className="h-[2px] w-24 bg-gradient-to-r from-transparent via-premium-200 to-transparent mx-auto mb-8"></div>
        <p className="text-[10px] font-black text-premium-300 uppercase tracking-[0.4em]">
          OEG Energy Analysis tool - © Mohamed Elshenawy | Mobi-Egypt 2026
        </p>
      </footer>
    </div>
  );
};

export default App;
