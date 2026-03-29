"use client";

import React, { useEffect, useMemo, useRef, useState } from "react";
// @ts-ignore
const Papa: {
  parse: (...args: any[]) => any;
  unparse?: (...args: any[]) => any;
} = require("papaparse");
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import {
  ResponsiveContainer,
  ComposedChart,
  Bar,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
} from "recharts";

type CleanRow = {
  dateD: string;
  dateR: string;
  repere: string;
  origine: string;
  totalUniteRaw: string;
  totalUnite: number;
  partDepose: number;
  partRepose: number;
  weekDepose: string | null;
  weekRepose: string | null;
  monthDepose: string | null;
  monthRepose: string | null;
};

type FournitureRow = {
  date: string;
  repere: string;
  origine: string;
  totalRaw: string;
  total: number;
  mois: string | null;
};

type WeeklyRow = {
  semaine: string;
  depose: number;
  repose: number;
  totalPondere: number;
  totalBrut: number;
  heures: number;
  hAct: number;
  objectif: number;
  ecart: number;
  statut: "OK" | "Vigilance" | "Alerte" | "Critique";
};

type MonthlyBillingRow = {
  mois: string;
  depose: number;
  repose: number;
  totalRealise: number;
  montantDepose: number;
  montantRepose: number;
  montantTotal: number;
};

type MonthlyGlobalRow = {
  mois: string;
  depose: number;
  repose: number;
  totalActivites: number;
  montantActivites: number;
  totalFournitures: number;
  totalGeneral: number;
};

type OriginFilterMode = "locked" | "editable";
type MenuType =
  | "dashboard"
  | "weekly"
  | "detail"
  | "export"
  | "billing"
  | "supplies"
  | "debug"
  | "settings";

type SiteConfig = {
  key: string;
  name: string;
  objectif: number;
  coefDepose: number;
  coefRepose: number;
};

const SITE_CONFIGS: SiteConfig[] = [
  { key: "blayais", name: "Blayais", objectif: 8, coefDepose: 0.3, coefRepose: 0.7 },
  { key: "golfech", name: "Golfech", objectif: 8, coefDepose: 0.3, coefRepose: 0.7 },
  { key: "bugey", name: "Bugey", objectif: 8, coefDepose: 0.3, coefRepose: 0.7 },
  { key: "st-alban", name: "St Alban", objectif: 8, coefDepose: 0.3, coefRepose: 0.7 },
  { key: "tricastin", name: "Tricastin", objectif: 8, coefDepose: 0.3, coefRepose: 0.7 },
];

const EMPTY_ORIGIN_LABEL = "(vide)";
const LOCAL_STORAGE_KEY = "pilotage-chantier-local-v3";

function parseFrenchNumber(value: string | number | undefined | null): number {
  if (value === undefined || value === null || value === "") return 0;

  const cleaned = value
    .toString()
    .trim()
    .replace(/\u00A0/g, " ")
    .replace(/\u202F/g, " ")
    .replace(/\s/g, "")
    .replace(",", ".");

  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function round2(value: number): number {
  return Math.round(value * 100) / 100;
}

function parseFrDate(dateStr: string): Date | null {
  if (!dateStr) return null;

  const cleaned = String(dateStr)
    .trim()
    .replace(/\./g, "/")
    .replace(/-/g, "/")
    .replace(/\s+/g, "");

  const parts = cleaned.split("/");
  if (parts.length !== 3) return null;

  const day = Number(parts[0]);
  const month = Number(parts[1]);
  let year = Number(parts[2]);

  if (!Number.isFinite(day) || !Number.isFinite(month) || !Number.isFinite(year)) {
    return null;
  }

  if (year < 100) year += 2000;
  if (day < 1 || day > 31 || month < 1 || month > 12) return null;

  const date = new Date(year, month - 1, day);
  date.setHours(12, 0, 0, 0);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }

  return date;
}

function getWeek(dateStr: string): string | null {
  const date = parseFrDate(dateStr);
  if (!date) return null;

  const target = new Date(date);
  target.setHours(12, 0, 0, 0);

  const dayNr = (target.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);

  const isoYear = target.getFullYear();

  const firstThursday = new Date(isoYear, 0, 4);
  firstThursday.setHours(12, 0, 0, 0);

  const firstDayNr = (firstThursday.getDay() + 6) % 7;
  firstThursday.setDate(firstThursday.getDate() - firstDayNr + 3);

  const weekNumber =
    1 + Math.round((target.getTime() - firstThursday.getTime()) / 604800000);

  return `${isoYear}-S${String(weekNumber).padStart(2, "0")}`;
}

function normalizeWeekKey(value: string | null | undefined): string | null {
  if (!value) return null;

  const cleaned = String(value)
    .replace(/\u00A0/g, " ")
    .replace(/\u202F/g, " ")
    .trim()
    .replace(/\s+/g, "");

  const match = cleaned.match(/(\d{4})-?S(\d{1,2})/i);
  if (!match) return cleaned;

  return `${match[1]}-S${String(match[2]).padStart(2, "0")}`;
}

function getDateFromWeekKey(weekKey: string): Date | null {
  const normalized = normalizeWeekKey(weekKey);
  if (!normalized) return null;

  const match = normalized.match(/^(\d{4})-S(\d{2})$/);
  if (!match) return null;

  const year = Number(match[1]);
  const week = Number(match[2]);

  const jan4 = new Date(year, 0, 4);
  jan4.setHours(12, 0, 0, 0);

  const dayNr = (jan4.getDay() + 6) % 7;
  const mondayWeek1 = new Date(jan4);
  mondayWeek1.setDate(jan4.getDate() - dayNr);

  const monday = new Date(mondayWeek1);
  monday.setDate(mondayWeek1.getDate() + (week - 1) * 7);

  return monday;
}

function getMonthKey(dateStr: string): string | null {
  const date = parseFrDate(dateStr);
  if (!date) return null;
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

function formatMonthLabel(monthKey: string): string {
  const [year, month] = monthKey.split("-");
  const date = new Date(Number(year), Number(month) - 1, 1);
  return date.toLocaleDateString("fr-FR", {
    month: "long",
    year: "numeric",
  });
}

function formatNumber(value: number): string {
  return value.toLocaleString("fr-FR", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2,
  });
}

function formatCurrency(value: number): string {
  return `${formatNumber(value)} €`;
}

function formatPdfNumber(value: number): string {
  return value
    .toLocaleString("fr-FR", {
      minimumFractionDigits: 0,
      maximumFractionDigits: 2,
      useGrouping: false,
    })
    .replace(/\u202F/g, "")
    .replace(/\u00A0/g, "");
}

function formatDateTime(value: string | null): string {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "-";
  return date.toLocaleString("fr-FR");
}

function normalizeText(value: string): string {
  return value
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/['’]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function sanitizeFileName(value: string): string {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9_-]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function getDisplayOrigine(origine: string): string {
  const trimmed = (origine || "").trim();
  if (!trimmed) return EMPTY_ORIGIN_LABEL;

  const normalized = normalizeText(trimmed);

  if (normalized.includes("tch repose")) return "TCH Repose";
  if (normalized === "tch") return "TCH";
  if (normalized === "alt") return "ALT";
  if (normalized === "forfait") return "Forfait";
  if (normalized === "ald amiante") return "ALD AMIANTE";

  return trimmed;
}

function findHeaderRowIndex(rows: string[][]): number {
  for (let i = 0; i < Math.min(rows.length, 50); i++) {
    const row = rows[i] || [];
    const normalized = row.map((cell) => normalizeText(cell || ""));

    const hasDateD = normalized.some((c) => c.includes("date d"));
    const hasDateR = normalized.some((c) => c.includes("date r"));
    const hasRepere = normalized.some((c) => c.includes("repere"));
    const hasOrigine = normalized.some((c) => c.includes("origine"));

    if (hasDateD && hasDateR && hasRepere && hasOrigine) return i;
  }

  return -1;
}

function isEndOfFirstTable(row: string[]): boolean {
  const line = normalizeText(row.join(" "));

  return (
    line.includes("2- bpu dajustement") ||
    line.includes("2 - bpu dajustement") ||
    line.includes("2 bpu dajustement") ||
    line.includes("3- bpu fournitures") ||
    line.includes("3 - bpu fournitures") ||
    line.includes("3 bpu fournitures") ||
    line.includes("bpu dajustement") ||
    line.includes("bpu fournitures")
  );
}

function findBpuFournituresStart(rows: string[][]): number {
  for (let i = 0; i < rows.length; i++) {
    const joined = normalizeText((rows[i] || []).join(" "));
    if (joined.includes("bpu fournitures")) return i;
  }
  return -1;
}

function findHeaderRowFrom(
  rows: string[][],
  startIndex: number,
  searchTerms: string[]
): number {
  for (let i = startIndex; i < Math.min(rows.length, startIndex + 20); i++) {
    const row = rows[i] || [];
    const normalized = row.map((cell) => normalizeText(cell || ""));
    const hasAll = searchTerms.every((term) =>
      normalized.some((c) => c.includes(normalizeText(term)))
    );
    if (hasAll) return i;
  }
  return -1;
}

function extractPuActivite(rows: string[][]): number {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];

    for (let j = 0; j < row.length; j++) {
      const rawCell = String(row[j] || "").trim();
      const cell = normalizeText(rawCell);

      const isPuActivite =
        cell.includes("pu activit") ||
        cell.includes("puactivit") ||
        rawCell.toLowerCase().includes("pu activ");

      if (isPuActivite) {
        for (let k = j + 1; k < row.length; k++) {
          const value = parseFrenchNumber(row[k]);
          if (value > 100 && value < 2000) return value;
        }

        for (let r = Math.max(0, i - 1); r <= Math.min(rows.length - 1, i + 2); r++) {
          const testRow = rows[r] || [];
          for (
            let c = Math.max(0, j - 1);
            c <= Math.min(testRow.length - 1, j + 4);
            c++
          ) {
            const value = parseFrenchNumber(testRow[c]);
            if (value > 100 && value < 2000) return value;
          }
        }
      }
    }
  }

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const joined = normalizeText(row.join(" "));

    if (joined.includes("base entre")) {
      for (let r = i; r <= Math.min(rows.length - 1, i + 5); r++) {
        const testRow = rows[r] || [];
        for (let c = 0; c < testRow.length; c++) {
          const value = parseFrenchNumber(testRow[c]);
          if (value > 100 && value < 2000) return value;
        }
      }
    }
  }

  return 0;
}

function detectMainYearFromRows(rows: CleanRow[]): number {
  const yearsCount: Record<number, number> = {};

  rows.forEach((row) => {
    [row.dateD, row.dateR].forEach((dateStr) => {
      const date = parseFrDate(dateStr);
      if (!date) return;

      const year = date.getFullYear();
      yearsCount[year] = (yearsCount[year] || 0) + 1;
    });
  });

  const entries = Object.entries(yearsCount);
  if (entries.length === 0) return new Date().getFullYear();

  entries.sort((a, b) => Number(b[1]) - Number(a[1]));
  return Number(entries[0][0]);
}

function getSheetWeek(sheetName: string, fallbackYear = new Date().getFullYear()): string | null {
  const match = sheetName.match(/S(\d{1,2})/i);
  if (!match) return null;
  return `${fallbackYear}-S${match[1].padStart(2, "0")}`;
}

function extractHoursFromSheet(sheetData: any[][]): number {
  if (!sheetData || sheetData.length === 0) return 0;

  let totalColumnIndex = -1;

  for (let r = 0; r < Math.min(sheetData.length, 20); r++) {
    const row = sheetData[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] ?? "").toLowerCase().trim();
      if (
        cell.includes("total (en heures)") ||
        cell.includes("total(en heures)") ||
        cell.includes("total en heures") ||
        cell.includes("totalheures")
      ) {
        totalColumnIndex = c;
        break;
      }
    }
    if (totalColumnIndex !== -1) break;
  }

  if (totalColumnIndex !== -1) {
    let maxValue = 0;
    for (let r = 0; r < sheetData.length; r++) {
      const value = parseFrenchNumber(sheetData[r]?.[totalColumnIndex]);
      if (value > maxValue) maxValue = value;
    }
    if (maxValue > 0) return maxValue;
  }

  let globalMax = 0;
  for (let r = 0; r < sheetData.length; r++) {
    for (let c = 0; c < (sheetData[r] || []).length; c++) {
      const value = parseFrenchNumber(sheetData[r][c]);
      if (value > globalMax) globalMax = value;
    }
  }

  return globalMax;
}

function getStatut(hAct: number): "OK" | "Vigilance" | "Alerte" | "Critique" {
  if (hAct <= 8) return "OK";
  if (hAct <= 10) return "Vigilance";
  if (hAct <= 15) return "Alerte";
  return "Critique";
}

function getStatutStyle(statut: WeeklyRow["statut"]): React.CSSProperties {
  switch (statut) {
    case "OK":
      return {
        backgroundColor: "#dcfce7",
        color: "#166534",
        fontWeight: 700,
        textAlign: "center",
      };
    case "Vigilance":
      return {
        backgroundColor: "#fef3c7",
        color: "#92400e",
        fontWeight: 700,
        textAlign: "center",
      };
    case "Alerte":
      return {
        backgroundColor: "#fed7aa",
        color: "#9a3412",
        fontWeight: 700,
        textAlign: "center",
      };
    case "Critique":
      return {
        backgroundColor: "#fecaca",
        color: "#991b1b",
        fontWeight: 700,
        textAlign: "center",
      };
  }
}

function MenuButton({
  label,
  id,
  activeMenu,
  setActiveMenu,
}: {
  label: string;
  id: MenuType;
  activeMenu: MenuType;
  setActiveMenu: (menu: MenuType) => void;
}) {
  const active = activeMenu === id;

  return (
    <button
      onClick={() => setActiveMenu(id)}
      style={{
        width: "100%",
        textAlign: "left",
        border: "none",
        borderRadius: 10,
        padding: "12px 14px",
        marginBottom: 8,
        cursor: "pointer",
        fontWeight: 600,
        background: active ? "#2563eb" : "transparent",
        color: "#ffffff",
      }}
    >
      {label}
    </button>
  );
}

function StatusBadge({
  ok,
  label,
  detail,
}: {
  ok: boolean;
  label: string;
  detail: string;
}) {
  return (
    <div
      style={{
        ...cardStyle,
        padding: 16,
        display: "flex",
        alignItems: "center",
        gap: 12,
        borderColor: ok ? "#bbf7d0" : "#e5e7eb",
        background: ok ? "#f0fdf4" : "#ffffff",
      }}
    >
      <div
        style={{
          width: 14,
          height: 14,
          borderRadius: 999,
          background: ok ? "#22c55e" : "#d1d5db",
          flexShrink: 0,
        }}
      />
      <div style={{ minWidth: 0 }}>
        <div style={{ fontWeight: 700 }}>{label}</div>
        <div style={{ color: "#6b7280", fontSize: 13, marginTop: 2 }}>{detail}</div>
      </div>
    </div>
  );
}

export default function Home() {
  const factuInputRef = useRef<HTMLInputElement | null>(null);
  const heuresInputRef = useRef<HTMLInputElement | null>(null);

  const [activeMenu, setActiveMenu] = useState<MenuType>("dashboard");
  const [data, setData] = useState<CleanRow[]>([]);
  const [suppliesData, setSuppliesData] = useState<FournitureRow[]>([]);
  const [weeklyData, setWeeklyData] = useState<WeeklyRow[]>([]);
  const [heuresMap, setHeuresMap] = useState<Record<string, number>>({});
  const [debugHeures, setDebugHeures] = useState<string>("");
  const [debugCsv, setDebugCsv] = useState<string>("");
  const [statusFilter, setStatusFilter] = useState<"Tous" | WeeklyRow["statut"]>("Tous");
  const [mailCopied, setMailCopied] = useState(false);
  const [originFilterMode, setOriginFilterMode] = useState<OriginFilterMode>("editable");
  const [selectedOrigins, setSelectedOrigins] = useState<string[]>([]);
  const [billingMonthFilter, setBillingMonthFilter] = useState("Tous");
  const [suppliesMonthFilter, setSuppliesMonthFilter] = useState("Tous");

  const [selectedSiteKey, setSelectedSiteKey] = useState("blayais");
  const [customSiteName, setCustomSiteName] = useState("");
  const [customObjectifValue, setCustomObjectifValue] = useState(8);
  const [customCoefDepose, setCustomCoefDepose] = useState(0.3);
  const [customCoefRepose, setCustomCoefRepose] = useState(0.7);
  const [puActivite, setPuActivite] = useState(0);
  const [factuMainYear, setFactuMainYear] = useState<number>(new Date().getFullYear());

  const [lastFactuFileName, setLastFactuFileName] = useState("");
  const [lastHeuresFileName, setLastHeuresFileName] = useState("");
  const [lastFactuImportAt, setLastFactuImportAt] = useState<string | null>(null);
  const [lastHeuresImportAt, setLastHeuresImportAt] = useState<string | null>(null);

  const selectedSiteConfig =
    SITE_CONFIGS.find((site) => site.key === selectedSiteKey) || SITE_CONFIGS[0];

  const isCustomSite = selectedSiteKey === "custom";

  const siteName = isCustomSite ? customSiteName || "Personnalisé" : selectedSiteConfig.name;
  const objectifValue = isCustomSite ? customObjectifValue : selectedSiteConfig.objectif;
  const coefDepose = isCustomSite ? customCoefDepose : selectedSiteConfig.coefDepose;
  const coefRepose = isCustomSite ? customCoefRepose : selectedSiteConfig.coefRepose;

  useEffect(() => {
    try {
      const raw = localStorage.getItem(LOCAL_STORAGE_KEY);
      if (!raw) return;

      const saved = JSON.parse(raw);

      if (Array.isArray(saved.data)) setData(saved.data);
      if (Array.isArray(saved.suppliesData)) setSuppliesData(saved.suppliesData);
      if (saved.heuresMap && typeof saved.heuresMap === "object") setHeuresMap(saved.heuresMap);
      if (typeof saved.debugHeures === "string") setDebugHeures(saved.debugHeures);
      if (typeof saved.debugCsv === "string") setDebugCsv(saved.debugCsv);
      if (Array.isArray(saved.selectedOrigins)) setSelectedOrigins(saved.selectedOrigins);
      if (typeof saved.billingMonthFilter === "string") setBillingMonthFilter(saved.billingMonthFilter);
      if (typeof saved.suppliesMonthFilter === "string") setSuppliesMonthFilter(saved.suppliesMonthFilter);
      if (typeof saved.selectedSiteKey === "string") setSelectedSiteKey(saved.selectedSiteKey);
      if (typeof saved.customSiteName === "string") setCustomSiteName(saved.customSiteName);
      if (typeof saved.customObjectifValue === "number") setCustomObjectifValue(saved.customObjectifValue);
      if (typeof saved.customCoefDepose === "number") setCustomCoefDepose(saved.customCoefDepose);
      if (typeof saved.customCoefRepose === "number") setCustomCoefRepose(saved.customCoefRepose);
      if (typeof saved.puActivite === "number") setPuActivite(saved.puActivite);
      if (typeof saved.factuMainYear === "number") setFactuMainYear(saved.factuMainYear);
      if (typeof saved.lastFactuFileName === "string") setLastFactuFileName(saved.lastFactuFileName);
      if (typeof saved.lastHeuresFileName === "string") setLastHeuresFileName(saved.lastHeuresFileName);
      if (typeof saved.lastFactuImportAt === "string" || saved.lastFactuImportAt === null) {
        setLastFactuImportAt(saved.lastFactuImportAt);
      }
      if (typeof saved.lastHeuresImportAt === "string" || saved.lastHeuresImportAt === null) {
        setLastHeuresImportAt(saved.lastHeuresImportAt);
      }
      if (saved.originFilterMode === "locked" || saved.originFilterMode === "editable") {
        setOriginFilterMode(saved.originFilterMode);
      }
      if (
        saved.activeMenu === "dashboard" ||
        saved.activeMenu === "weekly" ||
        saved.activeMenu === "detail" ||
        saved.activeMenu === "export" ||
        saved.activeMenu === "billing" ||
        saved.activeMenu === "supplies" ||
        saved.activeMenu === "debug" ||
        saved.activeMenu === "settings"
      ) {
        setActiveMenu(saved.activeMenu);
      }
    } catch (error) {
      console.error("Erreur lecture sauvegarde locale :", error);
    }
  }, []);

  useEffect(() => {
    try {
      const payload = {
        data,
        suppliesData,
        heuresMap,
        debugHeures,
        debugCsv,
        selectedOrigins,
        billingMonthFilter,
        suppliesMonthFilter,
        selectedSiteKey,
        customSiteName,
        customObjectifValue,
        customCoefDepose,
        customCoefRepose,
        puActivite,
        factuMainYear,
        originFilterMode,
        activeMenu,
        lastFactuFileName,
        lastHeuresFileName,
        lastFactuImportAt,
        lastHeuresImportAt,
      };

      localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(payload));
    } catch (error) {
      console.error("Erreur sauvegarde locale :", error);
    }
  }, [
    data,
    suppliesData,
    heuresMap,
    debugHeures,
    debugCsv,
    selectedOrigins,
    billingMonthFilter,
    suppliesMonthFilter,
    selectedSiteKey,
    customSiteName,
    customObjectifValue,
    customCoefDepose,
    customCoefRepose,
    puActivite,
    factuMainYear,
    originFilterMode,
    activeMenu,
    lastFactuFileName,
    lastHeuresFileName,
    lastFactuImportAt,
    lastHeuresImportAt,
  ]);

  const availableOrigins = useMemo(() => {
    const allOrigins = [
      ...data.map((row) => getDisplayOrigine(row.origine)),
      ...suppliesData.map((row) => getDisplayOrigine(row.origine)),
    ];
    return Array.from(new Set(allOrigins)).sort((a, b) => a.localeCompare(b, "fr"));
  }, [data, suppliesData]);

  const filteredData = useMemo(() => {
    if (selectedOrigins.length === 0) return data;
    return data.filter((row) => selectedOrigins.includes(getDisplayOrigine(row.origine)));
  }, [data, selectedOrigins]);

  const filteredSuppliesData = useMemo(() => {
    if (selectedOrigins.length === 0) return suppliesData;
    return suppliesData.filter((row) => selectedOrigins.includes(getDisplayOrigine(row.origine)));
  }, [suppliesData, selectedOrigins]);

  const buildWeekly = (
    factuData: CleanRow[],
    importedHeures: Record<string, number>
  ) => {
    const grouped: Record<string, { depose: number; repose: number; totalBrut: number }> = {};

    factuData.forEach((row) => {
      const weekD = normalizeWeekKey(row.weekDepose);
      const weekR = normalizeWeekKey(row.weekRepose);

      if (row.partDepose > 0 && weekD) {
        if (!grouped[weekD]) grouped[weekD] = { depose: 0, repose: 0, totalBrut: 0 };
        grouped[weekD].depose += row.partDepose;
        grouped[weekD].totalBrut += row.totalUnite;
      }

      if (row.partRepose > 0 && weekR) {
        if (!grouped[weekR]) grouped[weekR] = { depose: 0, repose: 0, totalBrut: 0 };
        grouped[weekR].repose += row.partRepose;
        grouped[weekR].totalBrut += row.totalUnite;
      }
    });

    const normalizedHeuresMap: Record<string, number> = {};
    Object.entries(importedHeures).forEach(([key, value]) => {
      const normalizedKey = normalizeWeekKey(key);
      if (normalizedKey) normalizedHeuresMap[normalizedKey] = value;
    });

    const result: WeeklyRow[] = Object.entries(grouped)
      .map(([week, val]) => {
        const normalizedWeek = normalizeWeekKey(week) || week;
        const heures = normalizedHeuresMap[normalizedWeek] || 0;
        const totalPondere = round2(val.depose + val.repose);
        const hAct = totalPondere > 0 ? round2(heures / totalPondere) : 0;
        const objectif = objectifValue;
        const ecart = round2(hAct - objectif);
        const statut = getStatut(hAct);

        return {
          semaine: normalizedWeek,
          depose: round2(val.depose),
          repose: round2(val.repose),
          totalPondere,
          totalBrut: round2(val.totalBrut),
          heures,
          hAct,
          objectif,
          ecart,
          statut,
        };
      })
      .sort((a, b) => {
        const da = getDateFromWeekKey(a.semaine)?.getTime() || 0;
        const db = getDateFromWeekKey(b.semaine)?.getTime() || 0;
        return da - db;
      });

    setWeeklyData(result);
  };

  useEffect(() => {
    buildWeekly(filteredData, heuresMap);
  }, [filteredData, heuresMap, objectifValue]);

  const resetFactuOnly = () => {
    setData([]);
    setSuppliesData([]);
    setWeeklyData([]);
    setDebugCsv("");
    setSelectedOrigins([]);
    setBillingMonthFilter("Tous");
    setSuppliesMonthFilter("Tous");
    setPuActivite(0);
    setFactuMainYear(new Date().getFullYear());
    setLastFactuFileName("");
    setLastFactuImportAt(null);
  };

  const resetHeuresOnly = () => {
    setHeuresMap({});
    setDebugHeures("");
    setLastHeuresFileName("");
    setLastHeuresImportAt(null);
  };

  const clearLocalBackup = () => {
    const ok = window.confirm("Voulez-vous vraiment effacer toute la sauvegarde locale de cet appareil ?");
    if (!ok) return;

    localStorage.removeItem(LOCAL_STORAGE_KEY);

    setData([]);
    setSuppliesData([]);
    setWeeklyData([]);
    setHeuresMap({});
    setDebugHeures("");
    setDebugCsv("");
    setSelectedOrigins([]);
    setBillingMonthFilter("Tous");
    setSuppliesMonthFilter("Tous");
    setSelectedSiteKey("blayais");
    setCustomSiteName("");
    setCustomObjectifValue(8);
    setCustomCoefDepose(0.3);
    setCustomCoefRepose(0.7);
    setPuActivite(0);
    setFactuMainYear(new Date().getFullYear());
    setOriginFilterMode("editable");
    setActiveMenu("dashboard");
    setLastFactuFileName("");
    setLastHeuresFileName("");
    setLastFactuImportAt(null);
    setLastHeuresImportAt(null);
  };

  const processFactuFile = (file: File) => {
    Papa.parse(file, {
      header: false,
      delimiter: ";",
      skipEmptyLines: false,
      complete: (results: Papa.ParseResult<string[]>) => {
        const rows = (results.data as string[][]).map((row) =>
          (row || []).map((cell) => (cell ?? "").toString().trim())
        );

        const detectedPuActivite = extractPuActivite(rows);
        setPuActivite(detectedPuActivite);

        const headerRowIndex = findHeaderRowIndex(rows);

        if (headerRowIndex === -1) {
          alert("Impossible de trouver automatiquement la ligne DATE D / DATE R / REPERE / ORIGINE.");
          setData([]);
          setSuppliesData([]);
          setWeeklyData([]);
          setDebugCsv("");
          return;
        }

        const headerRow = rows[headerRowIndex] || [];
        const upperHeaderRow = headerRowIndex > 0 ? rows[headerRowIndex - 1] || [] : [];

        const rawDataRows = rows.slice(headerRowIndex + 1);

        const activityRows: string[][] = [];
        for (const row of rawDataRows) {
          if (isEndOfFirstTable(row)) break;
          activityRows.push(row);
        }

        const findIndexInRow = (row: string[], searchTerms: string[]) => {
          return row.findIndex((cell) => {
            const c = normalizeText(cell || "");
            return searchTerms.some((term) => c.includes(normalizeText(term)));
          });
        };

        const findIndexMerged = (searchTerms: string[]) => {
          const idxHeader = findIndexInRow(headerRow, searchTerms);
          if (idxHeader !== -1) return idxHeader;

          const idxUpper = findIndexInRow(upperHeaderRow, searchTerms);
          if (idxUpper !== -1) return idxUpper;

          return -1;
        };

        const IDX_DATE_D = findIndexMerged(["date d"]);
        const IDX_DATE_R = findIndexMerged(["date r"]);
        const IDX_REPERE = findIndexMerged(["repere"]);
        const IDX_ORIGINE = findIndexMerged(["origine"]);
        const IDX_TOTAL_UNITE = findIndexMerged(["total en unite", "total unite", "total"]);

        if (
          IDX_DATE_D === -1 ||
          IDX_DATE_R === -1 ||
          IDX_REPERE === -1 ||
          IDX_ORIGINE === -1 ||
          IDX_TOTAL_UNITE === -1
        ) {
          alert("Colonnes principales introuvables dans le premier tableau.");
          setData([]);
          setSuppliesData([]);
          setWeeklyData([]);
          setDebugCsv("");
          return;
        }

        const cleanRows: CleanRow[] = activityRows
          .map((row) => {
            const dateD = IDX_DATE_D >= 0 ? row[IDX_DATE_D] || "" : "";
            const dateR = IDX_DATE_R >= 0 ? row[IDX_DATE_R] || "" : "";
            const repere = IDX_REPERE >= 0 ? row[IDX_REPERE] || "" : "";
            const origine = IDX_ORIGINE >= 0 ? row[IDX_ORIGINE] || "" : "";
            const totalUniteRaw = IDX_TOTAL_UNITE >= 0 ? row[IDX_TOTAL_UNITE] || "" : "";

            const totalUnite = parseFrenchNumber(totalUniteRaw);
            const origineLower = normalizeText(origine);

            let partDepose = 0;
            let partRepose = 0;
            let weekDepose: string | null = null;
            let weekRepose: string | null = null;
            let monthDepose: string | null = null;
            let monthRepose: string | null = null;

            const hasDateD = dateD !== "";
            const hasDateR = dateR !== "";

            if (origineLower.includes("tch repose")) {
              if (hasDateR || hasDateD) {
                partRepose = round2(totalUnite * coefRepose);
                weekRepose = getWeek(dateR) || getWeek(dateD);
                monthRepose = getMonthKey(dateR) || getMonthKey(dateD);
              }
            } else {
              if (hasDateD) {
                partDepose = round2(totalUnite * coefDepose);
                weekDepose = getWeek(dateD);
                monthDepose = getMonthKey(dateD);
              }

              if (hasDateR) {
                partRepose = round2(totalUnite * coefRepose);
                weekRepose = getWeek(dateR);
                monthRepose = getMonthKey(dateR);
              }
            }

            return {
              dateD,
              dateR,
              repere,
              origine,
              totalUniteRaw,
              totalUnite,
              partDepose,
              partRepose,
              weekDepose,
              weekRepose,
              monthDepose,
              monthRepose,
            };
          })
          .filter((row) => {
            const repereNorm = normalizeText(row.repere);
            const origineNorm = normalizeText(row.origine);
            const totalRawNorm = normalizeText(row.totalUniteRaw);

            const hasRealRepere = repereNorm !== "";
            const hasDate = !!row.dateD || !!row.dateR;

            const isSummaryLine =
              repereNorm.includes("bpu") ||
              origineNorm.includes("bpu") ||
              repereNorm.includes("pourcentage") ||
              origineNorm.includes("pourcentage") ||
              repereNorm.includes("coefficient global") ||
              origineNorm.includes("coefficient global") ||
              repereNorm.includes("fortuit") ||
              origineNorm.includes("fortuit") ||
              repereNorm.includes("prix total") ||
              origineNorm.includes("prix total") ||
              repereNorm.includes("base entre") ||
              origineNorm.includes("base entre") ||
              totalRawNorm.includes("pourcentage") ||
              totalRawNorm.includes("fortuit");

            const isRealActivityLine = hasDate && hasRealRepere && row.totalUnite > 0;
            return isRealActivityLine && !isSummaryLine;
          });

        const detectedMainYear = detectMainYearFromRows(cleanRows);
        setFactuMainYear(detectedMainYear);

        const bpuStart = findBpuFournituresStart(rows);
        let fournituresRows: FournitureRow[] = [];
        let bpuHeaderIndex = -1;

        if (bpuStart !== -1) {
          bpuHeaderIndex = findHeaderRowFrom(rows, bpuStart, ["date", "repere", "origine", "total"]);
        }

        if (bpuHeaderIndex !== -1) {
          const bpuHeader = rows[bpuHeaderIndex] || [];

          const findBpuIdx = (terms: string[]) =>
            bpuHeader.findIndex((cell) => {
              const c = normalizeText(cell || "");
              return terms.some((t) => c.includes(normalizeText(t)));
            });

          const IDX_BPU_DATE = findBpuIdx(["date"]);
          const IDX_BPU_REPERE = findBpuIdx(["repere"]);
          const IDX_BPU_ORIGINE = findBpuIdx(["origine"]);
          const IDX_BPU_TOTAL = findBpuIdx(["total"]);

          const bpuDataRows = rows.slice(bpuHeaderIndex + 1);

          fournituresRows = bpuDataRows
            .map((row) => {
              const date = IDX_BPU_DATE >= 0 ? row[IDX_BPU_DATE] || "" : "";
              const repere = IDX_BPU_REPERE >= 0 ? row[IDX_BPU_REPERE] || "" : "";
              const origine = IDX_BPU_ORIGINE >= 0 ? row[IDX_BPU_ORIGINE] || "" : "";
              const totalRaw = IDX_BPU_TOTAL >= 0 ? row[IDX_BPU_TOTAL] || "" : "";
              const total = parseFrenchNumber(totalRaw);
              const mois = getMonthKey(date);

              return { date, repere, origine, totalRaw, total, mois };
            })
            .filter((row) => {
              const dateNorm = normalizeText(row.date);
              const repereNorm = normalizeText(row.repere);
              const origineNorm = normalizeText(row.origine);

              const looksLikeHeader =
                dateNorm === "date" || repereNorm === "repere" || origineNorm === "origine";

              const isUsefulRow = !!row.date && !!row.repere && row.total > 0;
              return isUsefulRow && !looksLikeHeader;
            });
        }

        const finalOrigins = Array.from(
          new Set([
            ...cleanRows.map((row) => getDisplayOrigine(row.origine)),
            ...fournituresRows.map((row) => getDisplayOrigine(row.origine)),
          ])
        ).sort((a, b) => a.localeCompare(b, "fr"));

        const excludedCount = activityRows.length - cleanRows.length;

        const debugInfo = {
          fichier: file.name,
          importedAt: new Date().toISOString(),
          headerRowIndex,
          detectedHeaderRow: headerRow,
          detectedUpperHeaderRow: upperHeaderRow,
          indexes: {
            DATE_D: IDX_DATE_D,
            DATE_R: IDX_DATE_R,
            REPERE: IDX_REPERE,
            ORIGINE: IDX_ORIGINE,
            TOTAL_UNITE: IDX_TOTAL_UNITE,
          },
          bpuFournitures: {
            startRowIndex: bpuStart,
            headerRowIndex: bpuHeaderIndex,
            rowsDetected: fournituresRows.length,
          },
          settingsUsed: {
            siteName,
            objectifValue,
            coefDepose,
            coefRepose,
            puActivite: detectedPuActivite,
            factuMainYear: detectedMainYear,
          },
          stats: {
            totalRowsInFile: rows.length,
            rawDataRowsBeforeFilter: activityRows.length,
            keptRows: cleanRows.length,
            excludedRows: excludedCount,
          },
          originsDetected: finalOrigins,
          sampleWeeks: cleanRows.slice(0, 20).map((row) => ({
            dateD: row.dateD,
            dateR: row.dateR,
            weekDepose: row.weekDepose,
            weekRepose: row.weekRepose,
          })),
        };

        setData(cleanRows);
        setSuppliesData(fournituresRows);
        setSelectedOrigins(finalOrigins);
        setBillingMonthFilter("Tous");
        setSuppliesMonthFilter("Tous");
        setDebugCsv(JSON.stringify(debugInfo, null, 2));
        setLastFactuFileName(file.name);
        setLastFactuImportAt(new Date().toISOString());
      },
      error: () => {
        alert("Erreur de lecture du CSV.");
        setDebugCsv("");
      },
    });
  };

  const processHeuresFile = async (file: File) => {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });

    const map: Record<string, number> = {};
    const debug: Record<string, { semaine: string; heures: number }> = {};

    workbook.SheetNames.forEach((sheetName) => {
      if (!/^S\d{1,2}$/i.test(sheetName.trim())) return;

      const week = getSheetWeek(sheetName, factuMainYear);
      if (!week) return;

      const sheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
      }) as any[][];

      const heures = extractHoursFromSheet(sheetData);
      const normalizedWeek = normalizeWeekKey(week);

      if (normalizedWeek) map[normalizedWeek] = heures;
      debug[sheetName] = { semaine: normalizedWeek || week, heures };
    });

    const factuWeeks = Array.from(
      new Set(
        data
          .flatMap((row) => [
            normalizeWeekKey(row.weekDepose),
            normalizeWeekKey(row.weekRepose),
          ])
          .filter(Boolean) as string[]
      )
    ).sort((a, b) => {
      const da = getDateFromWeekKey(a)?.getTime() || 0;
      const db = getDateFromWeekKey(b)?.getTime() || 0;
      return da - db;
    });

    const pointageWeeks = Object.keys(map).sort((a, b) => {
      const da = getDateFromWeekKey(a)?.getTime() || 0;
      const db = getDateFromWeekKey(b)?.getTime() || 0;
      return da - db;
    });

    const missingInPointage = factuWeeks.filter((w) => !pointageWeeks.includes(w));
    const missingInFactu = pointageWeeks.filter((w) => !factuWeeks.includes(w));

    setHeuresMap(map);
    setDebugHeures(
      JSON.stringify(
        {
          fichier: file.name,
          importedAt: new Date().toISOString(),
          factuMainYear,
          pointage: debug,
          comparaison: {
            semainesFactu: factuWeeks,
            semainesPointage: pointageWeeks,
            absentesDuPointage: missingInPointage,
            absentesDeLaFactu: missingInFactu,
          },
        },
        null,
        2
      )
    );
    setLastHeuresFileName(file.name);
    setLastHeuresImportAt(new Date().toISOString());
  };

  const handleFactuInputClick = () => {
    if (factuInputRef.current) factuInputRef.current.value = "";
    factuInputRef.current?.click();
  };

  const handleHeuresInputClick = () => {
    if (heuresInputRef.current) heuresInputRef.current.value = "";
    heuresInputRef.current?.click();
  };

  const loadFactu = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    processFactuFile(file);
    event.target.value = "";
  };

  const loadHeuresExcel = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    await processHeuresFile(file);
    event.target.value = "";
  };

  const handleOriginChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
    const values = Array.from(event.target.selectedOptions, (option) => option.value);
    setSelectedOrigins(values);
    setBillingMonthFilter("Tous");
    setSuppliesMonthFilter("Tous");
  };

  const selectAllOrigins = () => {
    setSelectedOrigins(availableOrigins);
    setBillingMonthFilter("Tous");
    setSuppliesMonthFilter("Tous");
  };

  const clearOrigins = () => {
    setSelectedOrigins([]);
    setBillingMonthFilter("Tous");
    setSuppliesMonthFilter("Tous");
  };

  const filteredWeeklyData = useMemo(() => {
    if (statusFilter === "Tous") return weeklyData;
    return weeklyData.filter((row) => row.statut === statusFilter);
  }, [weeklyData, statusFilter]);

  const totalDepose = weeklyData.reduce((sum, row) => sum + row.depose, 0);
  const totalRepose = weeklyData.reduce((sum, row) => sum + row.repose, 0);
  const totalPondere = weeklyData.reduce((sum, row) => sum + row.totalPondere, 0);
  const totalBrutExcel = filteredData.reduce((sum, row) => sum + row.totalUnite, 0);
  const totalHeures = weeklyData.reduce((sum, row) => sum + row.heures, 0);
  const hGlobal = totalPondere > 0 ? round2(totalHeures / totalPondere) : 0;
  const statutGlobal = getStatut(hGlobal);
  const ecartGlobal = round2(hGlobal - objectifValue);

  const montantTotalActivites = round2(totalBrutExcel * puActivite);
  const montantDepose = round2(totalDepose * puActivite);
  const montantRepose = round2(totalRepose * puActivite);
  const montantTotalRealise = round2(totalPondere * puActivite);

  const monthlyBillingData = useMemo(() => {
    const grouped: Record<string, MonthlyBillingRow> = {};

    filteredData.forEach((row) => {
      if (row.partDepose > 0 && row.monthDepose) {
        if (!grouped[row.monthDepose]) {
          grouped[row.monthDepose] = {
            mois: row.monthDepose,
            depose: 0,
            repose: 0,
            totalRealise: 0,
            montantDepose: 0,
            montantRepose: 0,
            montantTotal: 0,
          };
        }

        grouped[row.monthDepose].depose += row.partDepose;
        grouped[row.monthDepose].montantDepose += row.partDepose * puActivite;
      }

      if (row.partRepose > 0 && row.monthRepose) {
        if (!grouped[row.monthRepose]) {
          grouped[row.monthRepose] = {
            mois: row.monthRepose,
            depose: 0,
            repose: 0,
            totalRealise: 0,
            montantDepose: 0,
            montantRepose: 0,
            montantTotal: 0,
          };
        }

        grouped[row.monthRepose].repose += row.partRepose;
        grouped[row.monthRepose].montantRepose += row.partRepose * puActivite;
      }
    });

    return Object.values(grouped)
      .map((row) => {
        const depose = round2(row.depose);
        const repose = round2(row.repose);
        const totalRealise = round2(depose + repose);
        const montantDepose = round2(row.montantDepose);
        const montantRepose = round2(row.montantRepose);
        const montantTotal = round2(montantDepose + montantRepose);

        return {
          mois: row.mois,
          depose,
          repose,
          totalRealise,
          montantDepose,
          montantRepose,
          montantTotal,
        };
      })
      .sort((a, b) => a.mois.localeCompare(b.mois));
  }, [filteredData, puActivite]);

  const billingMonthOptions = useMemo(
    () => monthlyBillingData.map((row) => row.mois),
    [monthlyBillingData]
  );

  const displayedMonthlyBillingData = useMemo(() => {
    if (billingMonthFilter === "Tous") return monthlyBillingData;
    return monthlyBillingData.filter((row) => row.mois === billingMonthFilter);
  }, [monthlyBillingData, billingMonthFilter]);

  const monthlyBillingTotals = useMemo(() => {
    return displayedMonthlyBillingData.reduce(
      (acc, row) => {
        acc.depose += row.depose;
        acc.repose += row.repose;
        acc.totalRealise += row.totalRealise;
        acc.montantDepose += row.montantDepose;
        acc.montantRepose += row.montantRepose;
        acc.montantTotal += row.montantTotal;
        return acc;
      },
      {
        depose: 0,
        repose: 0,
        totalRealise: 0,
        montantDepose: 0,
        montantRepose: 0,
        montantTotal: 0,
      }
    );
  }, [displayedMonthlyBillingData]);

  const monthlySuppliesData = useMemo(() => {
    const grouped: Record<string, number> = {};

    filteredSuppliesData.forEach((row) => {
      if (!row.mois) return;
      if (!grouped[row.mois]) grouped[row.mois] = 0;
      grouped[row.mois] += row.total;
    });

    return Object.entries(grouped)
      .map(([mois, total]) => ({
        mois,
        total: round2(total),
      }))
      .sort((a, b) => a.mois.localeCompare(b.mois));
  }, [filteredSuppliesData]);

  const suppliesMonthOptions = useMemo(
    () => monthlySuppliesData.map((row) => row.mois),
    [monthlySuppliesData]
  );

  const displayedMonthlySuppliesData = useMemo(() => {
    if (suppliesMonthFilter === "Tous") return monthlySuppliesData;
    return monthlySuppliesData.filter((row) => row.mois === suppliesMonthFilter);
  }, [monthlySuppliesData, suppliesMonthFilter]);

  const totalSuppliesGlobal = useMemo(
    () => round2(filteredSuppliesData.reduce((sum, row) => sum + row.total, 0)),
    [filteredSuppliesData]
  );

  const totalSuppliesDisplayed = useMemo(
    () => round2(displayedMonthlySuppliesData.reduce((sum, row) => sum + row.total, 0)),
    [displayedMonthlySuppliesData]
  );

  const monthlyGlobalData = useMemo(() => {
    const grouped: Record<string, MonthlyGlobalRow> = {};

    monthlyBillingData.forEach((row) => {
      if (!grouped[row.mois]) {
        grouped[row.mois] = {
          mois: row.mois,
          depose: 0,
          repose: 0,
          totalActivites: 0,
          montantActivites: 0,
          totalFournitures: 0,
          totalGeneral: 0,
        };
      }

      grouped[row.mois].depose += row.depose;
      grouped[row.mois].repose += row.repose;
      grouped[row.mois].totalActivites += row.totalRealise;
      grouped[row.mois].montantActivites += row.montantTotal;
    });

    monthlySuppliesData.forEach((row) => {
      if (!grouped[row.mois]) {
        grouped[row.mois] = {
          mois: row.mois,
          depose: 0,
          repose: 0,
          totalActivites: 0,
          montantActivites: 0,
          totalFournitures: 0,
          totalGeneral: 0,
        };
      }

      grouped[row.mois].totalFournitures += row.total;
    });

    return Object.values(grouped)
      .map((row) => {
        const depose = round2(row.depose);
        const repose = round2(row.repose);
        const totalActivites = round2(row.totalActivites);
        const montantActivites = round2(row.montantActivites);
        const totalFournitures = round2(row.totalFournitures);
        const totalGeneral = round2(montantActivites + totalFournitures);

        return {
          mois: row.mois,
          depose,
          repose,
          totalActivites,
          montantActivites,
          totalFournitures,
          totalGeneral,
        };
      })
      .sort((a, b) => a.mois.localeCompare(b.mois));
  }, [monthlyBillingData, monthlySuppliesData]);

  const displayedMonthlyGlobalData = useMemo(() => {
    if (billingMonthFilter === "Tous") return monthlyGlobalData;
    return monthlyGlobalData.filter((row) => row.mois === billingMonthFilter);
  }, [monthlyGlobalData, billingMonthFilter]);

  const monthlyGlobalTotals = useMemo(() => {
    return displayedMonthlyGlobalData.reduce(
      (acc, row) => {
        acc.depose += row.depose;
        acc.repose += row.repose;
        acc.totalActivites += row.totalActivites;
        acc.montantActivites += row.montantActivites;
        acc.totalFournitures += row.totalFournitures;
        acc.totalGeneral += row.totalGeneral;
        return acc;
      },
      {
        depose: 0,
        repose: 0,
        totalActivites: 0,
        montantActivites: 0,
        totalFournitures: 0,
        totalGeneral: 0,
      }
    );
  }, [displayedMonthlyGlobalData]);

  const bestWeek = useMemo(() => {
    if (weeklyData.length === 0) return null;
    return [...weeklyData].sort((a, b) => a.hAct - b.hAct)[0];
  }, [weeklyData]);

  const worstWeek = useMemo(() => {
    if (weeklyData.length === 0) return null;
    return [...weeklyData].sort((a, b) => b.hAct - a.hAct)[0];
  }, [weeklyData]);

  const summaryText = useMemo(() => {
    if (weeklyData.length === 0) return "";

    const critiques = weeklyData.filter((r) => r.statut === "Critique").map((r) => r.semaine);
    const alertes = weeklyData.filter((r) => r.statut === "Alerte").map((r) => r.semaine);
    const vigilances = weeklyData.filter((r) => r.statut === "Vigilance").map((r) => r.semaine);

    return [
      `Site : ${siteName}.`,
      `Origines retenues : ${selectedOrigins.length > 0 ? selectedOrigins.join(", ") : "Toutes"}.`,
      `Objectif : ${formatNumber(objectifValue)} h/activités.`,
      `Total activités : ${formatNumber(totalBrutExcel)}.`,
      `Total réalisée : ${formatNumber(totalPondere)}.`,
      `Bilan global : ${formatNumber(totalHeures)} h pour ${formatNumber(totalPondere)} activités réalisées, soit ${formatNumber(hGlobal)} h/activités (${statutGlobal}).`,
      critiques.length ? `Semaines critiques : ${critiques.join(", ")}.` : "",
      alertes.length ? `Semaines en alerte : ${alertes.join(", ")}.` : "",
      vigilances.length ? `Semaines en vigilance : ${vigilances.join(", ")}.` : "",
    ]
      .filter(Boolean)
      .join(" ");
  }, [
    weeklyData,
    siteName,
    selectedOrigins,
    objectifValue,
    totalBrutExcel,
    totalPondere,
    totalHeures,
    hGlobal,
    statutGlobal,
  ]);

  const mailText = useMemo(() => {
    if (weeklyData.length === 0) return "";

    const critiques = weeklyData.filter((r) => r.statut === "Critique").map((r) => r.semaine);
    const alertes = weeklyData.filter((r) => r.statut === "Alerte").map((r) => r.semaine);

    return [
      "Bonjour,",
      "",
      `Voici le bilan hebdomadaire du Pilotage Chantier - ${siteName} :`,
      "",
      `- Origines retenues : ${selectedOrigins.length > 0 ? selectedOrigins.join(", ") : "Toutes"}`,
      `- Objectif : ${formatNumber(objectifValue)} h/activités`,
      `- Total activités : ${formatNumber(totalBrutExcel)}`,
      `- Dépose réalisée : ${formatNumber(totalDepose)}`,
      `- Repose réalisée : ${formatNumber(totalRepose)}`,
      `- Total réalisée : ${formatNumber(totalPondere)}`,
      `- Heures AT : ${formatNumber(totalHeures)}`,
      `- Rendement global : ${formatNumber(hGlobal)} h/activités`,
      `- Écart global : ${formatNumber(ecartGlobal)}`,
      `- Statut global : ${statutGlobal}`,
      "",
      critiques.length ? `Semaines critiques : ${critiques.join(", ")}` : "Aucune semaine critique",
      alertes.length ? `Semaines en alerte : ${alertes.join(", ")}` : "Aucune semaine en alerte",
      "",
      "Cordialement,",
    ].join("\n");
  }, [
    weeklyData,
    siteName,
    selectedOrigins,
    objectifValue,
    totalBrutExcel,
    totalDepose,
    totalRepose,
    totalPondere,
    totalHeures,
    hGlobal,
    ecartGlobal,
    statutGlobal,
  ]);

  const copyMail = async () => {
    if (!mailText) return;
    await navigator.clipboard.writeText(mailText);
    setMailCopied(true);
    setTimeout(() => setMailCopied(false), 2000);
  };

  const monthlyBillingText = useMemo(() => {
    if (displayedMonthlyBillingData.length === 0) return "";

    const title =
      billingMonthFilter === "Tous"
        ? "Facturation mensuelle"
        : `Facturation du mois ${formatMonthLabel(billingMonthFilter)}`;

    return [
      title,
      `Site : ${siteName}`,
      `PU Activités : ${formatNumber(puActivite)}`,
      "",
      ...displayedMonthlyBillingData.map(
        (row) =>
          `${formatMonthLabel(row.mois)} | Dépose: ${formatNumber(row.depose)} | Repose: ${formatNumber(row.repose)} | Total: ${formatNumber(row.totalRealise)} | Montant: ${formatCurrency(row.montantTotal)}`
      ),
      "",
      `TOTAL | Dépose: ${formatNumber(monthlyBillingTotals.depose)} | Repose: ${formatNumber(monthlyBillingTotals.repose)} | Total: ${formatNumber(monthlyBillingTotals.totalRealise)} | Montant: ${formatCurrency(monthlyBillingTotals.montantTotal)}`,
    ].join("\n");
  }, [
    displayedMonthlyBillingData,
    billingMonthFilter,
    siteName,
    puActivite,
    monthlyBillingTotals,
  ]);

  const copyMonthlyBilling = async () => {
    if (!monthlyBillingText) return;
    await navigator.clipboard.writeText(monthlyBillingText);
  };

  const suppliesSummaryText = useMemo(() => {
    if (displayedMonthlySuppliesData.length === 0) return "";

    const title =
      suppliesMonthFilter === "Tous"
        ? "Facturation fournitures"
        : `Facturation fournitures - ${formatMonthLabel(suppliesMonthFilter)}`;

    return [
      title,
      `Site : ${siteName}`,
      `Origines retenues : ${selectedOrigins.length > 0 ? selectedOrigins.join(", ") : "Toutes"}`,
      "",
      ...displayedMonthlySuppliesData.map(
        (row) => `${formatMonthLabel(row.mois)} | Total fournitures : ${formatCurrency(row.total)}`
      ),
      "",
      `TOTAL FOURNITURES : ${formatCurrency(totalSuppliesDisplayed)}`,
    ].join("\n");
  }, [displayedMonthlySuppliesData, suppliesMonthFilter, siteName, selectedOrigins, totalSuppliesDisplayed]);

  const copySuppliesSummary = async () => {
    if (!suppliesSummaryText) return;
    await navigator.clipboard.writeText(suppliesSummaryText);
  };

  const monthlyGlobalSummaryText = useMemo(() => {
    if (displayedMonthlyGlobalData.length === 0) return "";

    const title =
      billingMonthFilter === "Tous"
        ? "Synthèse mensuelle globale"
        : `Synthèse mensuelle globale - ${formatMonthLabel(billingMonthFilter)}`;

    return [
      title,
      `Site : ${siteName}`,
      `PU Activités : ${formatNumber(puActivite)}`,
      "",
      ...displayedMonthlyGlobalData.map(
        (row) =>
          `${formatMonthLabel(row.mois)} | Dépose: ${formatNumber(row.depose)} | Repose: ${formatNumber(row.repose)} | Activités: ${formatCurrency(row.montantActivites)} | Fournitures: ${formatCurrency(row.totalFournitures)} | Total général: ${formatCurrency(row.totalGeneral)}`
      ),
      "",
      `TOTAL | Activités: ${formatCurrency(monthlyGlobalTotals.montantActivites)} | Fournitures: ${formatCurrency(monthlyGlobalTotals.totalFournitures)} | Total général: ${formatCurrency(monthlyGlobalTotals.totalGeneral)}`,
    ].join("\n");
  }, [
    displayedMonthlyGlobalData,
    billingMonthFilter,
    siteName,
    puActivite,
    monthlyGlobalTotals,
  ]);

  const copyMonthlyGlobalSummary = async () => {
    if (!monthlyGlobalSummaryText) return;
    await navigator.clipboard.writeText(monthlyGlobalSummaryText);
  };

  const loadPdfLogo = async (): Promise<HTMLImageElement | null> => {
    try {
      const logo = new Image();
      logo.src = "/logo-opterm.png";
      await new Promise<void>((resolve, reject) => {
        logo.onload = () => resolve();
        logo.onerror = () => reject();
      });
      return logo;
    } catch {
      return null;
    }
  };

  const drawPdfHeader = async (
    doc: jsPDF,
    title: string,
    subtitle: string,
    metadata: Array<[string, string]>
  ) => {
    const pageWidth = doc.internal.pageSize.getWidth();

    doc.setFillColor(31, 41, 55);
    doc.rect(0, 0, pageWidth, 30, "F");

    const logo = await loadPdfLogo();
    if (logo) doc.addImage(logo, "PNG", 14, 6, 28, 14);

    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(17);
    doc.text(title, 50, 13);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text(subtitle, 50, 20);

    doc.setTextColor(0, 0, 0);
    doc.setDrawColor(220, 220, 220);
    doc.setFillColor(248, 250, 252);
    doc.roundedRect(14, 38, 182, 36, 3, 3, "FD");

    let y = 48;
    metadata.forEach(([label, value]) => {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(10);
      doc.text(`${label} :`, 18, y);
      doc.setFont("helvetica", "normal");
      doc.text(value || "-", 52, y);
      y += 8;
    });
  };

  const finalizePdf = (doc: jsPDF, filename: string) => {
    const pageHeight = doc.internal.pageSize.getHeight();
    doc.setFontSize(8);
    doc.setTextColor(120, 120, 120);
    doc.text("Document généré automatiquement - Pilotage Chantier", 14, pageHeight - 8);
    doc.save(filename);
  };

  const exportBillingPdf = async () => {
    if (displayedMonthlyBillingData.length === 0) return;

    const doc = new jsPDF("p", "mm", "a4");

    const selectedMonthLabel =
      billingMonthFilter === "Tous" ? "Tous les mois" : formatMonthLabel(billingMonthFilter);

    await drawPdfHeader(
      doc,
      `Facturation activités - ${siteName}`,
      "Synthèse mensuelle activités",
      [
        ["Site", siteName],
        ["Mois", selectedMonthLabel],
        ["Origines", selectedOrigins.length > 0 ? selectedOrigins.join(", ") : "Toutes"],
        ["PU Activités", formatCurrency(puActivite)],
      ]
    );

    autoTable(doc, {
      startY: 82,
      head: [[
        "Mois",
        "Dépose réalisée",
        "Repose réalisée",
        "Total réalisée",
        "Montant dépose",
        "Montant repose",
        "Montant total",
      ]],
      body: displayedMonthlyBillingData.map((row) => [
        formatMonthLabel(row.mois),
        formatPdfNumber(row.depose),
        formatPdfNumber(row.repose),
        formatPdfNumber(row.totalRealise),
        formatPdfNumber(row.montantDepose),
        formatPdfNumber(row.montantRepose),
        formatPdfNumber(row.montantTotal),
      ]),
      foot: [[
        "TOTAL",
        formatPdfNumber(monthlyBillingTotals.depose),
        formatPdfNumber(monthlyBillingTotals.repose),
        formatPdfNumber(monthlyBillingTotals.totalRealise),
        formatPdfNumber(monthlyBillingTotals.montantDepose),
        formatPdfNumber(monthlyBillingTotals.montantRepose),
        formatPdfNumber(monthlyBillingTotals.montantTotal),
      ]],
      theme: "grid",
      headStyles: {
        fillColor: [37, 99, 235],
        textColor: [255, 255, 255],
        fontStyle: "bold",
      },
      footStyles: {
        fillColor: [243, 244, 246],
        textColor: [17, 24, 39],
        fontStyle: "bold",
      },
      styles: {
        fontSize: 9,
        cellPadding: 3,
      },
    });

    const safeSite = sanitizeFileName(siteName || "Site");
    const safeMonth = sanitizeFileName(
      billingMonthFilter === "Tous" ? "tous_les_mois" : billingMonthFilter
    );
    finalizePdf(doc, `Facturation_Activites_${safeSite}_${safeMonth}.pdf`);
  };

  const exportSuppliesPdf = async () => {
    if (displayedMonthlySuppliesData.length === 0) return;

    const doc = new jsPDF("p", "mm", "a4");

    const selectedMonthLabel =
      suppliesMonthFilter === "Tous" ? "Tous les mois" : formatMonthLabel(suppliesMonthFilter);

    await drawPdfHeader(
      doc,
      `Facturation fournitures - ${siteName}`,
      "Synthèse mensuelle fournitures",
      [
        ["Site", siteName],
        ["Mois", selectedMonthLabel],
        ["Origines", selectedOrigins.length > 0 ? selectedOrigins.join(", ") : "Toutes"],
        ["Date édition", new Date().toLocaleDateString("fr-FR")],
      ]
    );

    autoTable(doc, {
      startY: 82,
      head: [["Mois", "Total fournitures"]],
      body: displayedMonthlySuppliesData.map((row) => [
        formatMonthLabel(row.mois),
        formatPdfNumber(row.total),
      ]),
      foot: [["TOTAL", formatPdfNumber(totalSuppliesDisplayed)]],
      theme: "grid",
      headStyles: {
        fillColor: [37, 99, 235],
        textColor: [255, 255, 255],
        fontStyle: "bold",
      },
      footStyles: {
        fillColor: [243, 244, 246],
        textColor: [17, 24, 39],
        fontStyle: "bold",
      },
      styles: {
        fontSize: 9,
        cellPadding: 3,
      },
    });

    const startY = (doc as any).lastAutoTable?.finalY || 120;

    autoTable(doc, {
      startY: startY + 8,
      head: [["Date", "Repère", "Origine", "Mois", "Total fournitures"]],
      body: filteredSuppliesData
        .filter((row) =>
          suppliesMonthFilter === "Tous" ? true : row.mois === suppliesMonthFilter
        )
        .map((row) => [
          row.date,
          row.repere,
          getDisplayOrigine(row.origine),
          row.mois ? formatMonthLabel(row.mois) : "",
          formatPdfNumber(row.total),
        ]),
      theme: "grid",
      headStyles: {
        fillColor: [59, 130, 246],
        textColor: [255, 255, 255],
        fontStyle: "bold",
      },
      styles: {
        fontSize: 8,
        cellPadding: 2.5,
      },
    });

    const safeSite = sanitizeFileName(siteName || "Site");
    const safeMonth = sanitizeFileName(
      suppliesMonthFilter === "Tous" ? "tous_les_mois" : suppliesMonthFilter
    );
    finalizePdf(doc, `Facturation_Fournitures_${safeSite}_${safeMonth}.pdf`);
  };

  const exportGlobalPdf = async () => {
    if (displayedMonthlyGlobalData.length === 0) return;

    const doc = new jsPDF("p", "mm", "a4");

    const selectedMonthLabel =
      billingMonthFilter === "Tous" ? "Tous les mois" : formatMonthLabel(billingMonthFilter);

    await drawPdfHeader(
      doc,
      `Facturation globale - ${siteName}`,
      "Synthèse mensuelle globale",
      [
        ["Site", siteName],
        ["Mois", selectedMonthLabel],
        ["Origines", selectedOrigins.length > 0 ? selectedOrigins.join(", ") : "Toutes"],
        ["PU Activités", formatCurrency(puActivite)],
      ]
    );

    autoTable(doc, {
      startY: 82,
      head: [[
        "Mois",
        "Dépose",
        "Repose",
        "Total activités",
        "Montant activités",
        "Fournitures",
        "Total général",
      ]],
      body: displayedMonthlyGlobalData.map((row) => [
        formatMonthLabel(row.mois),
        formatPdfNumber(row.depose),
        formatPdfNumber(row.repose),
        formatPdfNumber(row.totalActivites),
        formatPdfNumber(row.montantActivites),
        formatPdfNumber(row.totalFournitures),
        formatPdfNumber(row.totalGeneral),
      ]),
      foot: [[
        "TOTAL",
        formatPdfNumber(monthlyGlobalTotals.depose),
        formatPdfNumber(monthlyGlobalTotals.repose),
        formatPdfNumber(monthlyGlobalTotals.totalActivites),
        formatPdfNumber(monthlyGlobalTotals.montantActivites),
        formatPdfNumber(monthlyGlobalTotals.totalFournitures),
        formatPdfNumber(monthlyGlobalTotals.totalGeneral),
      ]],
      theme: "grid",
      headStyles: {
        fillColor: [37, 99, 235],
        textColor: [255, 255, 255],
        fontStyle: "bold",
      },
      footStyles: {
        fillColor: [243, 244, 246],
        textColor: [17, 24, 39],
        fontStyle: "bold",
      },
      styles: {
        fontSize: 9,
        cellPadding: 3,
      },
      didParseCell: (hookData) => {
        if (hookData.section === "body" && hookData.column.index === 6) {
          hookData.cell.styles.fontStyle = "bold";
        }
      },
    });

    const safeSite = sanitizeFileName(siteName || "Site");
    const safeMonth = sanitizeFileName(
      billingMonthFilter === "Tous" ? "tous_les_mois" : billingMonthFilter
    );
    finalizePdf(doc, `Facturation_Globale_${safeSite}_${safeMonth}.pdf`);
  };

  const generateProfessionalPDF = async () => {
    const doc = new jsPDF("p", "mm", "a4");
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();

    const periodeText =
      weeklyData.length > 0
        ? `${weeklyData[0].semaine} à ${weeklyData[weeklyData.length - 1].semaine}`
        : "-";

    const origineText = selectedOrigins.length > 0 ? selectedOrigins.join(", ") : "Toutes";

    doc.setFillColor(31, 41, 55);
    doc.rect(0, 0, pageWidth, 30, "F");

    try {
      const logo = new Image();
      logo.src = "/logo-opterm.png";

      await new Promise<void>((resolve, reject) => {
        logo.onload = () => resolve();
        logo.onerror = () => reject();
      });

      doc.addImage(logo, "PNG", 14, 6, 28, 14);
    } catch (e) {
      console.log("Logo non chargé", e);
    }

    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(18);
    doc.text(`Pilotage Chantier - ${siteName}`, 50, 13);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text("Bilan hebdomadaire de performance", 50, 20);

    doc.setTextColor(0, 0, 0);

    doc.setDrawColor(220, 220, 220);
    doc.setFillColor(248, 250, 252);
    doc.roundedRect(14, 38, 182, 40, 3, 3, "FD");

    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("Site :", 18, 48);
    doc.text("Origine(s) :", 18, 56);
    doc.text("Période :", 18, 64);
    doc.text("Date d'édition :", 18, 72);

    doc.setFont("helvetica", "normal");
    doc.text(siteName, 48, 48);
    doc.text(origineText, 48, 56);
    doc.text(periodeText, 48, 64);
    doc.text(new Date().toLocaleDateString("fr-FR"), 48, 72);

    doc.setFillColor(255, 255, 255);
    doc.roundedRect(14, 86, 182, 42, 3, 3, "FD");

    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("Total activités", 18, 96);
    doc.text("Dépose réalisée", 62, 96);
    doc.text("Repose réalisée", 112, 96);
    doc.text("Total réalisée", 160, 96);

    doc.setFontSize(13);
    doc.text(formatPdfNumber(totalBrutExcel), 18, 112);
    doc.text(formatPdfNumber(totalDepose), 62, 112);
    doc.text(formatPdfNumber(totalRepose), 112, 112);
    doc.text(formatPdfNumber(totalPondere), 160, 112);

    doc.setFillColor(248, 250, 252);
    doc.roundedRect(14, 134, 182, 24, 3, 3, "FD");

    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.text(`Rendement global : ${formatPdfNumber(hGlobal)} h/activités`, 18, 144);
    doc.text(
      `Objectif : ${formatPdfNumber(objectifValue)} | Écart : ${formatPdfNumber(ecartGlobal)}`,
      18,
      152
    );

    let statutColor: [number, number, number] = [220, 252, 231];
    let statutTextColor: [number, number, number] = [22, 101, 52];

    if (statutGlobal === "Vigilance") {
      statutColor = [254, 243, 199];
      statutTextColor = [146, 64, 14];
    } else if (statutGlobal === "Alerte") {
      statutColor = [254, 215, 170];
      statutTextColor = [154, 52, 18];
    } else if (statutGlobal === "Critique") {
      statutColor = [254, 202, 202];
      statutTextColor = [153, 27, 27];
    }

    doc.setFillColor(...statutColor);
    doc.roundedRect(150, 140, 35, 10, 2, 2, "F");
    doc.setTextColor(...statutTextColor);
    doc.text(statutGlobal || "N/A", 167.5, 146.5, { align: "center" });
    doc.setTextColor(0, 0, 0);

    autoTable(doc, {
      startY: 164,
      head: [[
        "Semaine",
        "Total activités",
        "Dépose réalisée",
        "Repose réalisée",
        "Total réalisée",
        "Heures AT",
        "h/activités",
        "Objectif",
        "Écart",
        "Statut",
      ]],
      body: weeklyData.map((row) => [
        row.semaine,
        formatPdfNumber(row.totalBrut),
        formatPdfNumber(row.depose),
        formatPdfNumber(row.repose),
        formatPdfNumber(row.totalPondere),
        formatPdfNumber(row.heures),
        formatPdfNumber(row.hAct),
        formatPdfNumber(row.objectif),
        formatPdfNumber(row.ecart),
        row.statut,
      ]),
      theme: "grid",
      headStyles: {
        fillColor: [37, 99, 235],
        textColor: [255, 255, 255],
        fontStyle: "bold",
        halign: "center",
      },
      styles: {
        fontSize: 8,
        cellPadding: 2.5,
      },
      didParseCell: (hookData) => {
        if (hookData.section === "body" && hookData.column.index === 9) {
          const statut = String(hookData.cell.raw);
          if (statut === "OK") hookData.cell.styles.fillColor = [220, 252, 231];
          if (statut === "Vigilance") hookData.cell.styles.fillColor = [254, 243, 199];
          if (statut === "Alerte") hookData.cell.styles.fillColor = [254, 215, 170];
          if (statut === "Critique") hookData.cell.styles.fillColor = [254, 202, 202];
        }
      },
    });

    doc.setFontSize(8);
    doc.setTextColor(120, 120, 120);
    doc.text("Document généré automatiquement - Pilotage Chantier", 14, pageHeight - 8);

    const safeSite = sanitizeFileName(siteName || "Site");
    doc.save(`Pilotage_Chantier_${safeSite}.pdf`);
  };

  const csvLoaded = data.length > 0 || suppliesData.length > 0;
  const heuresLoaded = Object.keys(heuresMap).length > 0;

  return (
    <main
      style={{
        display: "flex",
        minHeight: "100vh",
        fontFamily: "Arial, sans-serif",
        background: "#f5f7fb",
        color: "#1f2937",
      }}
    >
      <aside
        style={{
          width: 250,
          background: "#111827",
          color: "#ffffff",
          padding: 20,
          flexShrink: 0,
        }}
      >
        <h2 style={{ marginTop: 0, marginBottom: 24 }}>Pilotage Chantier</h2>

        <MenuButton label="📊 Dashboard" id="dashboard" activeMenu={activeMenu} setActiveMenu={setActiveMenu} />
        <MenuButton label="📅 Suivi semaine" id="weekly" activeMenu={activeMenu} setActiveMenu={setActiveMenu} />
        <MenuButton label="📄 Détail facturation" id="detail" activeMenu={activeMenu} setActiveMenu={setActiveMenu} />
        <MenuButton label="📤 Export" id="export" activeMenu={activeMenu} setActiveMenu={setActiveMenu} />
        <MenuButton label="💶 Facturation activités" id="billing" activeMenu={activeMenu} setActiveMenu={setActiveMenu} />
        <MenuButton label="🧰 Facturation fournitures" id="supplies" activeMenu={activeMenu} setActiveMenu={setActiveMenu} />
        <MenuButton label="⚙️ Paramètres" id="settings" activeMenu={activeMenu} setActiveMenu={setActiveMenu} />
        <MenuButton label="🛠 Debug" id="debug" activeMenu={activeMenu} setActiveMenu={setActiveMenu} />
      </aside>

      <section
        style={{
          flex: 1,
          padding: 32,
          overflow: "auto",
        }}
      >
        <div style={{ maxWidth: 1320, margin: "0 auto" }}>
          <h1 style={{ marginBottom: 8 }}>Pilotage Chantier</h1>
          <p style={{ marginTop: 0, color: "#6b7280" }}>
            Outil de pilotage des performances chantier
          </p>

          <input
            ref={factuInputRef}
            type="file"
            accept=".csv"
            onChange={loadFactu}
            style={{ display: "none" }}
          />

          <input
            ref={heuresInputRef}
            type="file"
            accept=".xlsx,.xls"
            onChange={loadHeuresExcel}
            style={{ display: "none" }}
          />

          <section style={{ ...cardStyle, marginBottom: 24, padding: 18 }}>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
                gap: 14,
                marginBottom: 14,
              }}
            >
              <StatusBadge
                ok={csvLoaded}
                label={csvLoaded ? "CSV chargé" : "CSV non chargé"}
                detail={
                  csvLoaded
                    ? `${lastFactuFileName || "-"} • ${formatDateTime(lastFactuImportAt)}`
                    : "Aucune facturation importée"
                }
              />

              <StatusBadge
                ok={heuresLoaded}
                label={heuresLoaded ? "Pointage chargé" : "Pointage non chargé"}
                detail={
                  heuresLoaded
                    ? `${lastHeuresFileName || "-"} • ${formatDateTime(lastHeuresImportAt)}`
                    : "Aucun pointage importé"
                }
              />
            </div>

            <div
              style={{
                display: "flex",
                gap: 10,
                flexWrap: "wrap",
                alignItems: "center",
              }}
            >
              <button onClick={handleFactuInputClick} style={buttonPrimaryStyle}>
                Importer / actualiser CSV
              </button>

              <button onClick={handleHeuresInputClick} style={buttonPrimaryStyle}>
                Importer / actualiser pointage
              </button>

              <button onClick={resetFactuOnly} style={buttonStyle}>
                Vider la factu
              </button>

              <button onClick={resetHeuresOnly} style={buttonStyle}>
                Vider le pointage
              </button>
            </div>
          </section>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
              gap: 16,
              marginTop: 20,
              marginBottom: 24,
            }}
          >
            <div style={cardStyle}>
              <p style={{ marginTop: 0, fontWeight: 700 }}>📄 Import facturation CSV</p>

              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button onClick={handleFactuInputClick} style={buttonPrimaryStyle}>
                  Choisir / réimporter le CSV
                </button>
              </div>

              <p style={{ marginTop: 12, marginBottom: 0 }}>
                <strong>Dernier fichier :</strong> {lastFactuFileName || "-"}
              </p>
              <p style={{ marginTop: 6, marginBottom: 0 }}>
                <strong>Dernier import :</strong> {formatDateTime(lastFactuImportAt)}
              </p>
              <p style={{ marginTop: 6, marginBottom: 0 }}>
                <strong>Lignes activités :</strong> {data.length}
              </p>
              <p style={{ marginTop: 6, marginBottom: 0 }}>
                <strong>Lignes fournitures :</strong> {suppliesData.length}
              </p>
            </div>

            <div style={cardStyle}>
              <p style={{ marginTop: 0, fontWeight: 700 }}>⏱ Import pointage Excel</p>

              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button onClick={handleHeuresInputClick} style={buttonPrimaryStyle}>
                  Choisir / réimporter le pointage
                </button>
              </div>

              <p style={{ marginTop: 12, marginBottom: 0 }}>
                <strong>Dernier fichier :</strong> {lastHeuresFileName || "-"}
              </p>
              <p style={{ marginTop: 6, marginBottom: 0 }}>
                <strong>Dernier import :</strong> {formatDateTime(lastHeuresImportAt)}
              </p>
              <p style={{ marginTop: 6, marginBottom: 0 }}>
                <strong>Semaines heures détectées :</strong> {Object.keys(heuresMap).length}
              </p>
            </div>
          </div>

          <section style={{ ...cardStyle, marginBottom: 24 }}>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
                gap: 16,
                alignItems: "start",
              }}
            >
              <div>
                <p style={{ marginTop: 0, fontWeight: 700 }}>⚙️ Mode filtre ORIGINE</p>

                <select
                  value={originFilterMode}
                  onChange={(e) => setOriginFilterMode(e.target.value as OriginFilterMode)}
                  style={{ ...selectStyle, width: "100%" }}
                >
                  <option value="locked">Origine verrouillée</option>
                  <option value="editable">Origine modifiable</option>
                </select>

                <p style={{ marginTop: 10, color: "#6b7280", fontSize: 14 }}>
                  {originFilterMode === "locked"
                    ? "Le filtre origine reste figé sur les origines importées."
                    : "Tu peux sélectionner une ou plusieurs origines."}
                </p>
              </div>

              <div>
                <p style={{ marginTop: 0, fontWeight: 700 }}>🏷️ Origines détectées</p>

                <select
                  multiple
                  value={selectedOrigins}
                  onChange={handleOriginChange}
                  disabled={originFilterMode === "locked"}
                  style={{
                    ...selectStyle,
                    width: "100%",
                    minHeight: 140,
                    background: originFilterMode === "locked" ? "#f3f4f6" : "#ffffff",
                    cursor: originFilterMode === "locked" ? "not-allowed" : "pointer",
                  }}
                >
                  {availableOrigins.map((origine) => (
                    <option key={origine} value={origine}>
                      {origine}
                    </option>
                  ))}
                </select>

                <div
                  style={{
                    display: "flex",
                    gap: 10,
                    marginTop: 10,
                    flexWrap: "wrap",
                  }}
                >
                  <button
                    onClick={selectAllOrigins}
                    style={buttonStyle}
                    disabled={originFilterMode === "locked"}
                  >
                    Tout sélectionner
                  </button>

                  <button
                    onClick={clearOrigins}
                    style={buttonStyle}
                    disabled={originFilterMode === "locked"}
                  >
                    Tout enlever
                  </button>
                </div>
              </div>
            </div>
          </section>

          {activeMenu === "dashboard" && (
            <>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(8, minmax(0, 1fr))",
                  gap: 16,
                  marginBottom: 28,
                }}
              >
                <div style={cardStyle}>
                  <div style={labelStyle}>Site</div>
                  <div style={{ ...valueStyle, fontSize: 22 }}>{siteName}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Objectif</div>
                  <div style={valueStyle}>{formatNumber(objectifValue)}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Total activités</div>
                  <div style={valueStyle}>{formatNumber(totalBrutExcel)}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Dépose réalisée</div>
                  <div style={valueStyle}>{formatNumber(totalDepose)}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Repose réalisée</div>
                  <div style={valueStyle}>{formatNumber(totalRepose)}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Total réalisée</div>
                  <div style={valueStyle}>{formatNumber(totalPondere)}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Heures AT</div>
                  <div style={valueStyle}>{formatNumber(totalHeures)}</div>
                </div>

                <div style={{ ...cardStyle, ...getStatutStyle(statutGlobal) }}>
                  <div style={{ fontSize: 14 }}>Statut global</div>
                  <div style={{ fontSize: 28, fontWeight: 700, marginTop: 8 }}>
                    {statutGlobal}
                  </div>
                  <div style={{ marginTop: 6 }}>
                    {formatNumber(hGlobal)} h/activités
                  </div>
                </div>
              </div>

              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(4, minmax(0, 1fr))",
                  gap: 16,
                  marginBottom: 28,
                }}
              >
                <div style={cardStyle}>
                  <div style={labelStyle}>Écart global</div>
                  <div style={valueStyle}>{formatNumber(ecartGlobal)}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Meilleure semaine</div>
                  <div style={{ ...valueStyle, fontSize: 22 }}>
                    {bestWeek ? bestWeek.semaine : "-"}
                  </div>
                  <div style={{ marginTop: 8, color: "#6b7280" }}>
                    {bestWeek ? `${formatNumber(bestWeek.hAct)} h/activités` : ""}
                  </div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Pire semaine</div>
                  <div style={{ ...valueStyle, fontSize: 22 }}>
                    {worstWeek ? worstWeek.semaine : "-"}
                  </div>
                  <div style={{ marginTop: 8, color: "#6b7280" }}>
                    {worstWeek ? `${formatNumber(worstWeek.hAct)} h/activités` : ""}
                  </div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Total fournitures</div>
                  <div style={valueStyle}>{formatCurrency(totalSuppliesGlobal)}</div>
                </div>
              </div>

              <section style={{ ...cardStyle, marginBottom: 28 }}>
                <h2 style={{ marginTop: 0, marginBottom: 6 }}>Synthèse chantier</h2>
                <p style={{ margin: 0, lineHeight: 1.5 }}>
                  {summaryText || "Importe les 2 fichiers pour générer la synthèse."}
                </p>
              </section>

              <section style={{ ...cardStyle, marginBottom: 28 }}>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    gap: 12,
                    flexWrap: "wrap",
                  }}
                >
                  <h2 style={{ margin: 0 }}>Graphique h/activités vs objectif</h2>

                  <select
                    value={statusFilter}
                    onChange={(e) => setStatusFilter(e.target.value as any)}
                    style={selectStyle}
                  >
                    <option value="Tous">Tous les statuts</option>
                    <option value="OK">OK</option>
                    <option value="Vigilance">Vigilance</option>
                    <option value="Alerte">Alerte</option>
                    <option value="Critique">Critique</option>
                  </select>
                </div>

                <div style={{ width: "100%", height: 360, marginTop: 20 }}>
                  <ResponsiveContainer>
                    <ComposedChart data={filteredWeeklyData}>
                      <CartesianGrid strokeDasharray="3 3" />
                      <XAxis dataKey="semaine" />
                      <YAxis />
                      <Tooltip />
                      <Legend />
                      <Bar
                        dataKey="hAct"
                        name="h/activités"
                        fill="#60a5fa"
                        radius={[4, 4, 0, 0]}
                      />
                      <Line
                        type="monotone"
                        dataKey="objectif"
                        name="Objectif"
                        stroke="#ef4444"
                        strokeWidth={2}
                        dot={false}
                      />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </section>
            </>
          )}

          {activeMenu === "weekly" && (
            <section style={{ ...cardStyle, marginBottom: 28 }}>
              <h2 style={{ marginTop: 0 }}>Suivi par semaine</h2>

              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 12 }}>
                  <thead>
                    <tr style={{ background: "#f9fafb" }}>
                      <th style={thStyle}>Semaine</th>
                      <th style={thStyle}>Total activités</th>
                      <th style={thStyle}>Dépose réalisée</th>
                      <th style={thStyle}>Repose réalisée</th>
                      <th style={thStyle}>Total réalisée</th>
                      <th style={thStyle}>Heures AT</th>
                      <th style={thStyle}>h/activités</th>
                      <th style={thStyle}>Objectif</th>
                      <th style={thStyle}>Écart</th>
                      <th style={thStyle}>Statut</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredWeeklyData.map((row, i) => (
                      <tr key={i}>
                        <td style={tdStyle}>{row.semaine}</td>
                        <td style={tdStyle}>{formatNumber(row.totalBrut)}</td>
                        <td style={tdStyle}>{formatNumber(row.depose)}</td>
                        <td style={tdStyle}>{formatNumber(row.repose)}</td>
                        <td style={tdStyle}>{formatNumber(row.totalPondere)}</td>
                        <td style={tdStyle}>{formatNumber(row.heures)}</td>
                        <td style={tdStyle}>{formatNumber(row.hAct)}</td>
                        <td style={tdStyle}>{formatNumber(row.objectif)}</td>
                        <td style={tdStyle}>{formatNumber(row.ecart)}</td>
                        <td style={{ ...tdStyle, ...getStatutStyle(row.statut) }}>
                          {row.statut}
                        </td>
                      </tr>
                    ))}

                    {statusFilter === "Tous" && weeklyData.length > 0 && (
                      <tr style={{ background: "#f9fafb", fontWeight: 700 }}>
                        <td style={tdStyle}>TOTAL GLOBAL</td>
                        <td style={tdStyle}>{formatNumber(totalBrutExcel)}</td>
                        <td style={tdStyle}>{formatNumber(totalDepose)}</td>
                        <td style={tdStyle}>{formatNumber(totalRepose)}</td>
                        <td style={tdStyle}>{formatNumber(totalPondere)}</td>
                        <td style={tdStyle}>{formatNumber(totalHeures)}</td>
                        <td style={tdStyle}>{formatNumber(hGlobal)}</td>
                        <td style={tdStyle}>{formatNumber(objectifValue)}</td>
                        <td style={tdStyle}>{formatNumber(ecartGlobal)}</td>
                        <td style={{ ...tdStyle, ...getStatutStyle(statutGlobal) }}>
                          {statutGlobal}
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {activeMenu === "detail" && (
            <section style={{ ...cardStyle, marginBottom: 28 }}>
              <h2 style={{ marginTop: 0 }}>Détail des lignes facturation activités</h2>

              <div style={{ overflowX: "auto", marginTop: 16 }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ background: "#f9fafb" }}>
                      <th style={thStyle}>DATE D</th>
                      <th style={thStyle}>DATE R</th>
                      <th style={thStyle}>REPERE</th>
                      <th style={thStyle}>ORIGINE</th>
                      <th style={thStyle}>Total activités</th>
                      <th style={thStyle}>Dépose réalisée</th>
                      <th style={thStyle}>Repose réalisée</th>
                      <th style={thStyle}>Sem. D</th>
                      <th style={thStyle}>Sem. R</th>
                      <th style={thStyle}>Mois D</th>
                      <th style={thStyle}>Mois R</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.map((row, index) => (
                      <tr key={index}>
                        <td style={tdStyle}>{row.dateD}</td>
                        <td style={tdStyle}>{row.dateR}</td>
                        <td style={tdStyle}>{row.repere}</td>
                        <td style={tdStyle}>{getDisplayOrigine(row.origine)}</td>
                        <td style={tdStyle}>{formatNumber(row.totalUnite)}</td>
                        <td style={tdStyle}>{formatNumber(row.partDepose)}</td>
                        <td style={tdStyle}>{formatNumber(row.partRepose)}</td>
                        <td style={tdStyle}>{row.weekDepose || ""}</td>
                        <td style={tdStyle}>{row.weekRepose || ""}</td>
                        <td style={tdStyle}>{row.monthDepose || ""}</td>
                        <td style={tdStyle}>{row.monthRepose || ""}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {activeMenu === "export" && (
            <>
              <section style={{ ...cardStyle, marginBottom: 24 }}>
                <h2 style={{ marginTop: 0 }}>Export</h2>
                <p style={{ color: "#6b7280" }}>
                  Génération du mail et export PDF client.
                </p>

                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
                  <button onClick={copyMail} style={buttonStyle}>
                    {mailCopied ? "Mail copié ✅" : "Copier le mail"}
                  </button>

                  <button onClick={generateProfessionalPDF} style={buttonPrimaryStyle}>
                    Exporter PDF
                  </button>
                </div>
              </section>

              <section style={{ ...cardStyle, marginBottom: 24 }}>
                <h3 style={{ marginTop: 0 }}>Aperçu du mail</h3>
                <pre style={{ whiteSpace: "pre-wrap", marginTop: 16 }}>
                  {mailText}
                </pre>
              </section>
            </>
          )}

          {activeMenu === "billing" && (
            <>
              <section style={{ ...cardStyle, marginBottom: 24 }}>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    gap: 12,
                    flexWrap: "wrap",
                  }}
                >
                  <div>
                    <h2 style={{ marginTop: 0, marginBottom: 6 }}>
                      Facturation mensuelle activités
                    </h2>
                    <p style={{ color: "#6b7280", marginTop: 0, marginBottom: 0 }}>
                      Dépose facturée au mois de DATE D et repose facturée au mois de DATE R.
                    </p>
                  </div>

                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                    <select
                      value={billingMonthFilter}
                      onChange={(e) => setBillingMonthFilter(e.target.value)}
                      style={selectStyle}
                    >
                      <option value="Tous">Tous les mois</option>
                      {billingMonthOptions.map((mois) => (
                        <option key={mois} value={mois}>
                          {formatMonthLabel(mois)}
                        </option>
                      ))}
                    </select>

                    <button onClick={copyMonthlyBilling} style={buttonStyle}>
                      Copier le résumé
                    </button>

                    <button onClick={exportBillingPdf} style={buttonPrimaryStyle}>
                      Exporter PDF activités
                    </button>
                  </div>
                </div>
              </section>

              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(5, minmax(0, 1fr))",
                  gap: 16,
                  marginBottom: 24,
                }}
              >
                <div style={cardStyle}>
                  <div style={labelStyle}>PU Activités</div>
                  <div style={valueStyle}>{formatNumber(puActivite)}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Montant activités</div>
                  <div style={valueStyle}>{formatCurrency(montantTotalActivites)}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Montant dépose</div>
                  <div style={valueStyle}>{formatCurrency(montantDepose)}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Montant repose</div>
                  <div style={valueStyle}>{formatCurrency(montantRepose)}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Montant total réalisé</div>
                  <div style={valueStyle}>{formatCurrency(montantTotalRealise)}</div>
                </div>
              </div>

              <section style={{ ...cardStyle, marginBottom: 24 }}>
                <h3 style={{ marginTop: 0 }}>Détail mensuel activités</h3>

                <div style={{ overflowX: "auto", marginTop: 12 }}>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ background: "#f9fafb" }}>
                        <th style={thStyle}>Mois</th>
                        <th style={thStyle}>Dépose réalisée</th>
                        <th style={thStyle}>Repose réalisée</th>
                        <th style={thStyle}>Total réalisée</th>
                        <th style={thStyle}>Montant dépose</th>
                        <th style={thStyle}>Montant repose</th>
                        <th style={thStyle}>Montant total</th>
                      </tr>
                    </thead>
                    <tbody>
                      {displayedMonthlyBillingData.map((row) => (
                        <tr key={row.mois}>
                          <td style={tdStyle}>{formatMonthLabel(row.mois)}</td>
                          <td style={tdStyle}>{formatNumber(row.depose)}</td>
                          <td style={tdStyle}>{formatNumber(row.repose)}</td>
                          <td style={tdStyle}>{formatNumber(row.totalRealise)}</td>
                          <td style={tdStyle}>{formatCurrency(row.montantDepose)}</td>
                          <td style={tdStyle}>{formatCurrency(row.montantRepose)}</td>
                          <td style={tdStyle}>{formatCurrency(row.montantTotal)}</td>
                        </tr>
                      ))}

                      {displayedMonthlyBillingData.length > 0 && (
                        <tr style={{ background: "#f9fafb", fontWeight: 700 }}>
                          <td style={tdStyle}>TOTAL</td>
                          <td style={tdStyle}>{formatNumber(monthlyBillingTotals.depose)}</td>
                          <td style={tdStyle}>{formatNumber(monthlyBillingTotals.repose)}</td>
                          <td style={tdStyle}>{formatNumber(monthlyBillingTotals.totalRealise)}</td>
                          <td style={tdStyle}>{formatCurrency(monthlyBillingTotals.montantDepose)}</td>
                          <td style={tdStyle}>{formatCurrency(monthlyBillingTotals.montantRepose)}</td>
                          <td style={tdStyle}>{formatCurrency(monthlyBillingTotals.montantTotal)}</td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </section>

              <section style={{ ...cardStyle, marginBottom: 24 }}>
                <h3 style={{ marginTop: 0 }}>Résumé facturation activités</h3>
                <pre style={{ whiteSpace: "pre-wrap", marginTop: 16 }}>
                  {monthlyBillingText}
                </pre>
              </section>

              <section style={{ ...cardStyle, marginBottom: 24 }}>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    gap: 12,
                    flexWrap: "wrap",
                  }}
                >
                  <div>
                    <h3 style={{ marginTop: 0, marginBottom: 6 }}>
                      Synthèse mensuelle globale
                    </h3>
                    <p style={{ color: "#6b7280", marginTop: 0, marginBottom: 0 }}>
                      Regroupement mensuel des activités et des fournitures.
                    </p>
                  </div>

                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                    <button onClick={copyMonthlyGlobalSummary} style={buttonStyle}>
                      Copier la synthèse globale
                    </button>

                    <button onClick={exportGlobalPdf} style={buttonPrimaryStyle}>
                      Exporter PDF global
                    </button>
                  </div>
                </div>

                <div style={{ overflowX: "auto", marginTop: 16 }}>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ background: "#f9fafb" }}>
                        <th style={thStyle}>Mois</th>
                        <th style={thStyle}>Dépose réalisée</th>
                        <th style={thStyle}>Repose réalisée</th>
                        <th style={thStyle}>Total activités</th>
                        <th style={thStyle}>Montant activités</th>
                        <th style={thStyle}>Total fournitures</th>
                        <th style={thStyle}>Total général</th>
                      </tr>
                    </thead>
                    <tbody>
                      {displayedMonthlyGlobalData.map((row) => (
                        <tr key={row.mois}>
                          <td style={tdStyle}>{formatMonthLabel(row.mois)}</td>
                          <td style={tdStyle}>{formatNumber(row.depose)}</td>
                          <td style={tdStyle}>{formatNumber(row.repose)}</td>
                          <td style={tdStyle}>{formatNumber(row.totalActivites)}</td>
                          <td style={tdStyle}>{formatCurrency(row.montantActivites)}</td>
                          <td style={tdStyle}>{formatCurrency(row.totalFournitures)}</td>
                          <td style={{ ...tdStyle, fontWeight: 700 }}>
                            {formatCurrency(row.totalGeneral)}
                          </td>
                        </tr>
                      ))}

                      {displayedMonthlyGlobalData.length > 0 && (
                        <tr style={{ background: "#f9fafb", fontWeight: 700 }}>
                          <td style={tdStyle}>TOTAL</td>
                          <td style={tdStyle}>{formatNumber(monthlyGlobalTotals.depose)}</td>
                          <td style={tdStyle}>{formatNumber(monthlyGlobalTotals.repose)}</td>
                          <td style={tdStyle}>{formatNumber(monthlyGlobalTotals.totalActivites)}</td>
                          <td style={tdStyle}>{formatCurrency(monthlyGlobalTotals.montantActivites)}</td>
                          <td style={tdStyle}>{formatCurrency(monthlyGlobalTotals.totalFournitures)}</td>
                          <td style={tdStyle}>{formatCurrency(monthlyGlobalTotals.totalGeneral)}</td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </section>

              <section style={cardStyle}>
                <h3 style={{ marginTop: 0 }}>Résumé global mensuel</h3>
                <pre style={{ whiteSpace: "pre-wrap", marginTop: 16 }}>
                  {monthlyGlobalSummaryText}
                </pre>
              </section>
            </>
          )}

          {activeMenu === "supplies" && (
            <>
              <section style={{ ...cardStyle, marginBottom: 24 }}>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    gap: 12,
                    flexWrap: "wrap",
                  }}
                >
                  <div>
                    <h2 style={{ marginTop: 0, marginBottom: 6 }}>
                      Facturation fournitures
                    </h2>
                    <p style={{ color: "#6b7280", marginTop: 0, marginBottom: 0 }}>
                      Regroupement mensuel du 2e tableau BPU FOURNITURES.
                    </p>
                  </div>

                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                    <select
                      value={suppliesMonthFilter}
                      onChange={(e) => setSuppliesMonthFilter(e.target.value)}
                      style={selectStyle}
                    >
                      <option value="Tous">Tous les mois</option>
                      {suppliesMonthOptions.map((mois) => (
                        <option key={mois} value={mois}>
                          {formatMonthLabel(mois)}
                        </option>
                      ))}
                    </select>

                    <button onClick={copySuppliesSummary} style={buttonStyle}>
                      Copier le résumé
                    </button>

                    <button onClick={exportSuppliesPdf} style={buttonPrimaryStyle}>
                      Exporter PDF fournitures
                    </button>
                  </div>
                </div>
              </section>

              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
                  gap: 16,
                  marginBottom: 24,
                }}
              >
                <div style={cardStyle}>
                  <div style={labelStyle}>Total fournitures global</div>
                  <div style={valueStyle}>{formatCurrency(totalSuppliesGlobal)}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Total mois sélectionné</div>
                  <div style={valueStyle}>{formatCurrency(totalSuppliesDisplayed)}</div>
                </div>

                <div style={cardStyle}>
                  <div style={labelStyle}>Lignes fournitures</div>
                  <div style={valueStyle}>{filteredSuppliesData.length}</div>
                </div>
              </div>

              <section style={{ ...cardStyle, marginBottom: 24 }}>
                <h3 style={{ marginTop: 0 }}>Détail mensuel fournitures</h3>

                <div style={{ overflowX: "auto", marginTop: 12 }}>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ background: "#f9fafb" }}>
                        <th style={thStyle}>Mois</th>
                        <th style={thStyle}>Total fournitures</th>
                      </tr>
                    </thead>
                    <tbody>
                      {displayedMonthlySuppliesData.map((row) => (
                        <tr key={row.mois}>
                          <td style={tdStyle}>{formatMonthLabel(row.mois)}</td>
                          <td style={tdStyle}>{formatCurrency(row.total)}</td>
                        </tr>
                      ))}

                      {displayedMonthlySuppliesData.length > 0 && (
                        <tr style={{ background: "#f9fafb", fontWeight: 700 }}>
                          <td style={tdStyle}>TOTAL</td>
                          <td style={tdStyle}>{formatCurrency(totalSuppliesDisplayed)}</td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </section>

              <section style={{ ...cardStyle, marginBottom: 24 }}>
                <h3 style={{ marginTop: 0 }}>Détail lignes fournitures</h3>

                <div style={{ overflowX: "auto", marginTop: 12 }}>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ background: "#f9fafb" }}>
                        <th style={thStyle}>Date</th>
                        <th style={thStyle}>Repère</th>
                        <th style={thStyle}>Origine</th>
                        <th style={thStyle}>Mois</th>
                        <th style={thStyle}>Total fournitures</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredSuppliesData
                        .filter((row) =>
                          suppliesMonthFilter === "Tous" ? true : row.mois === suppliesMonthFilter
                        )
                        .map((row, idx) => (
                          <tr key={idx}>
                            <td style={tdStyle}>{row.date}</td>
                            <td style={tdStyle}>{row.repere}</td>
                            <td style={tdStyle}>{getDisplayOrigine(row.origine)}</td>
                            <td style={tdStyle}>{row.mois ? formatMonthLabel(row.mois) : ""}</td>
                            <td style={tdStyle}>{formatCurrency(row.total)}</td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              </section>

              <section style={cardStyle}>
                <h3 style={{ marginTop: 0 }}>Résumé facturation fournitures</h3>
                <pre style={{ whiteSpace: "pre-wrap", marginTop: 16 }}>
                  {suppliesSummaryText}
                </pre>
              </section>
            </>
          )}

          {activeMenu === "settings" && (
            <section style={cardStyle}>
              <h2 style={{ marginTop: 0 }}>Paramètres</h2>

              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
                  gap: 16,
                  marginTop: 16,
                }}
              >
                <div>
                  <label style={settingsLabelStyle}>Site</label>
                  <select
                    value={selectedSiteKey}
                    onChange={(e) => setSelectedSiteKey(e.target.value)}
                    style={{ ...selectStyle, width: "100%", marginTop: 6 }}
                  >
                    {SITE_CONFIGS.map((site) => (
                      <option key={site.key} value={site.key}>
                        {site.name}
                      </option>
                    ))}
                    <option value="custom">Personnalisé</option>
                  </select>
                </div>

                <div>
                  <label style={settingsLabelStyle}>Objectif h/activités</label>
                  <input
                    type="number"
                    step="0.1"
                    value={objectifValue}
                    onChange={(e) => setCustomObjectifValue(parseFrenchNumber(e.target.value))}
                    style={inputStyle}
                    disabled={!isCustomSite}
                  />
                </div>

                <div>
                  <label style={settingsLabelStyle}>Coefficient dépose</label>
                  <input
                    type="number"
                    step="0.01"
                    value={coefDepose}
                    onChange={(e) => setCustomCoefDepose(parseFrenchNumber(e.target.value))}
                    style={inputStyle}
                    disabled={!isCustomSite}
                  />
                </div>

                <div>
                  <label style={settingsLabelStyle}>Coefficient repose</label>
                  <input
                    type="number"
                    step="0.01"
                    value={coefRepose}
                    onChange={(e) => setCustomCoefRepose(parseFrenchNumber(e.target.value))}
                    style={inputStyle}
                    disabled={!isCustomSite}
                  />
                </div>

                {isCustomSite && (
                  <div style={{ gridColumn: "1 / -1" }}>
                    <label style={settingsLabelStyle}>Nom du site personnalisé</label>
                    <input
                      value={customSiteName}
                      onChange={(e) => setCustomSiteName(e.target.value)}
                      style={inputStyle}
                    />
                  </div>
                )}
              </div>

              <div style={{ ...cardStyle, marginTop: 24, padding: 16 }}>
                <h3 style={{ marginTop: 0 }}>Aperçu des paramètres actifs</h3>
                <p style={{ marginBottom: 8 }}>
                  <strong>Site :</strong> {siteName}
                </p>
                <p style={{ marginBottom: 8 }}>
                  <strong>Objectif :</strong> {formatNumber(objectifValue)} h/activités
                </p>
                <p style={{ marginBottom: 8 }}>
                  <strong>Coefficient dépose :</strong> {formatNumber(coefDepose)}
                </p>
                <p style={{ marginBottom: 8 }}>
                  <strong>Coefficient repose :</strong> {formatNumber(coefRepose)}
                </p>
                <p style={{ marginBottom: 8 }}>
                  <strong>Année principale factu :</strong> {factuMainYear}
                </p>
                <p style={{ marginBottom: 8 }}>
                  <strong>Dernier CSV :</strong> {lastFactuFileName || "-"}
                </p>
                <p style={{ marginBottom: 8 }}>
                  <strong>Dernier pointage :</strong> {lastHeuresFileName || "-"}
                </p>
                <p style={{ marginBottom: 0 }}>
                  <strong>Dernière mise à jour pointage :</strong> {formatDateTime(lastHeuresImportAt)}
                </p>
              </div>

              <div style={{ marginTop: 20, display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button onClick={clearLocalBackup} style={buttonStyle}>
                  Effacer toute la sauvegarde locale
                </button>
              </div>
            </section>
          )}

          {activeMenu === "debug" && (
            <>
              <section style={{ ...cardStyle, marginBottom: 24 }}>
                <h2 style={{ marginTop: 0 }}>Debug CSV</h2>
                <pre style={{ whiteSpace: "pre-wrap", marginTop: 16 }}>
                  {debugCsv}
                </pre>
              </section>

              <section style={cardStyle}>
                <h2 style={{ marginTop: 0 }}>Debug pointage</h2>
                <pre style={{ whiteSpace: "pre-wrap", marginTop: 16 }}>
                  {debugHeures}
                </pre>
              </section>
            </>
          )}
        </div>
      </section>
    </main>
  );
}

const cardStyle: React.CSSProperties = {
  background: "#ffffff",
  border: "1px solid #e5e7eb",
  borderRadius: 12,
  padding: 20,
};

const labelStyle: React.CSSProperties = {
  fontSize: 14,
  color: "#6b7280",
};

const valueStyle: React.CSSProperties = {
  fontSize: 28,
  fontWeight: 700,
  marginTop: 8,
};

const buttonStyle: React.CSSProperties = {
  background: "#ffffff",
  border: "1px solid #d1d5db",
  color: "#111827",
  borderRadius: 8,
  padding: "10px 14px",
  cursor: "pointer",
  fontWeight: 600,
};

const buttonPrimaryStyle: React.CSSProperties = {
  background: "#2563eb",
  border: "1px solid #2563eb",
  color: "#ffffff",
  borderRadius: 8,
  padding: "10px 14px",
  cursor: "pointer",
  fontWeight: 600,
};

const selectStyle: React.CSSProperties = {
  padding: "10px 12px",
  borderRadius: 8,
  border: "1px solid #d1d5db",
  background: "#ffffff",
};

const thStyle: React.CSSProperties = {
  textAlign: "left",
  padding: "12px",
  borderBottom: "1px solid #e5e7eb",
  fontSize: 14,
};

const tdStyle: React.CSSProperties = {
  padding: "10px 12px",
  borderBottom: "1px solid #f0f0f0",
  fontSize: 14,
};

const inputStyle: React.CSSProperties = {
  width: "100%",
  padding: "10px 12px",
  borderRadius: 8,
  border: "1px solid #d1d5db",
  background: "#ffffff",
  marginTop: 6,
};

const settingsLabelStyle: React.CSSProperties = {
  display: "block",
  fontWeight: 600,
  color: "#374151",
};