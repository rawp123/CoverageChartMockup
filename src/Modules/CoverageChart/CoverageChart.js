// Modules/CoverageChart/CoverageChart.js
// Floating "coverage tower" renderer with 3 Views:
//   - carrier
//   - carrierGroup
//   - availability
//
// Key quota share behavior:
//  - A (Year, Attachment) is considered a quota share layer if there are >=2 distinct PolicyIDs at that (x, attach).
//  - In Carrier view, quota share layers are forced into a single dataset labeled "Quota share" (prevents gaps).
//  - Tooltip title: quota layers show "Quota share" instead of "(unknown group)".

let chart = null;
let currentView = "carrier";

let _cache = {
  allSlices: [],
  allXLabels: [],
  slices: [],
  xLabels: [],
  options: null,
  quotaKeySet: new Set(), // `${x}||${attach}`
  useYearAxis: true,
  xZoom: 1,
  dom: {
    canvas: null,
    viewport: null,
    surface: null
  },
  _wheelBound: false,
  showCoverageTotals: true,
  filters: {
    startYear: null,
    endYear: null,
    startDate: null,
    endDate: null,
    zoomMin: null,
    zoomMax: null,
    sirMode: "off",
    insuranceProgram: "",
    policyLimitType: "",
    carriers: [],
    carrierGroups: []
  }
};

/* ================================
   Utilities
================================ */

const num = (v) => {
  if (v === null || v === undefined || v === "") return 0;
  const cleaned = String(v).replace(/[^0-9.-]/g, "");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
};

const firstPresentNum = (...vals) => {
  for (const v of vals) {
    if (v === null || v === undefined) continue;
    const s = String(v).trim();
    if (s === "") continue;
    return num(s);
  }
  return 0;
};

const yearOf = (v) => {
  const s = String(v ?? "").trim();
  const m = s.match(/(\d{4})/);
  return m ? +m[1] : null;
};

const parseDateToUTC = (v) => {
  const s = String(v ?? "").trim();
  if (!s) return null;
  const d = new Date(s);
  if (Number.isNaN(d.getTime())) return null;
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
};

const startOfYearUTC = (y) => Date.UTC(y, 0, 1, 0, 0, 0, 0);
const endOfYearUTC = (y) => Date.UTC(y, 11, 31, 23, 59, 59, 999);

const money = (v) => `$${Number(v || 0).toLocaleString()}`;

function compactMoney(v) {
  const n = Number(v);
  if (!Number.isFinite(n)) return "$0";
  const abs = Math.abs(n);
  const sign = n < 0 ? "-" : "";
  const units = [
    { value: 1e12, suffix: "T" },
    { value: 1e9, suffix: "B" },
    { value: 1e6, suffix: "M" },
    { value: 1e3, suffix: "K" }
  ];

  for (const u of units) {
    if (abs >= u.value) {
      const scaled = abs / u.value;
      const digits = scaled >= 100 ? 0 : scaled >= 10 ? 1 : 2;
      const text = Number(scaled.toFixed(digits)).toString();
      return `$${sign}${text}${u.suffix}`;
    }
  }

  return `$${sign}${Math.round(abs).toLocaleString()}`;
}

const normKey = (k) =>
  String(k ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const clamp = (v, min, max) => Math.min(max, Math.max(min, v));
const legendTitleForView = (view) => {
  if (view === "carrierGroup") return "Carrier Group";
  if (view === "carrier") return "Carrier";
  if (view === "availability") return "Availability";
  return "Legend";
};
const summarizeFilterSelection = (items, singular, pluralAll) => {
  const list = normalizeStringList(items);
  if (!list.length) return pluralAll;
  if (list.length === 1) return list[0];
  return `${list.length} ${singular}s`;
};
function getLegendFilterLines() {
  const f = _cache?.filters || {};
  const yearsText =
    f.startDate || f.endDate
      ? `${String(f.startDate || "All")} to ${String(f.endDate || "All")}`
      : Number.isFinite(f.startYear) || Number.isFinite(f.endYear)
      ? `${Number.isFinite(f.startYear) ? f.startYear : "All"} to ${Number.isFinite(f.endYear) ? f.endYear : "All"}`
      : "All years";
  const zoomText =
    Number.isFinite(f.zoomMin) || Number.isFinite(f.zoomMax)
      ? `${Number.isFinite(f.zoomMin) ? compactMoney(f.zoomMin) : "Auto"} to ${Number.isFinite(f.zoomMax) ? compactMoney(f.zoomMax) : "Auto"}`
      : "Auto";
  return [
    `Program: ${String(f.insuranceProgram || "(none)")}`,
    `Policy Limit Type: ${String(f.policyLimitType || "(none)")}`,
    `Years: ${yearsText}`,
    `Carriers: ${summarizeFilterSelection(f.carriers, "carrier", "All carriers")}`,
    `Carrier Groups: ${summarizeFilterSelection(f.carrierGroups, "group", "All groups")}`,
    `Zoom: ${zoomText}`
  ];
}

function roundedRectPath(ctx, x, y, w, h, r) {
  const radius = Math.max(0, Math.min(r, w / 2, h / 2));
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.arcTo(x + w, y, x + w, y + h, radius);
  ctx.arcTo(x + w, y + h, x, y + h, radius);
  ctx.arcTo(x, y + h, x, y, radius);
  ctx.arcTo(x, y, x + w, y, radius);
  ctx.closePath();
}

const getThemeName = () =>
  document?.documentElement?.dataset?.theme === "light" ? "light" : "dark";

const getChartThemeColors = (themeName = getThemeName()) => {
  if (themeName === "light") {
    return {
      outline: "rgba(30, 41, 59, 0.32)",
      legendText: "rgba(15, 23, 42, 0.92)",
      xGrid: "rgba(15, 23, 42, 0.12)",
      yGrid: "rgba(15, 23, 42, 0.14)",
      axisTicks: "rgba(15, 23, 42, 0.92)",
      yTitle: "rgba(15, 23, 42, 0.95)"
    };
  }

  return {
    outline: "rgba(15,23,42,0.42)",
    legendText: "rgba(255,255,255,0.95)",
    xGrid: "rgba(255,255,255,0.06)",
    yGrid: "rgba(255,255,255,0.08)",
    axisTicks: "rgba(255,255,255,0.92)",
    yTitle: "rgba(255,255,255,0.95)"
  };
};

const getBy = (obj, ...candidates) => {
  if (!obj) return "";
  const map = {};
  for (const k of Object.keys(obj)) map[normKey(k)] = obj[k];
  for (const c of candidates) {
    const v = map[normKey(c)];
    if (v !== undefined && String(v).trim() !== "") return v;
  }
  return "";
};

function normalizeHeader(h) {
  return String(h ?? "").trim().replace(/\s+/g, " ");
}

const DISTINCT_COLOR_HUES = [
  210, 0, 120, 280, 32, 15, 330, 190,
  65, 265, 235, 142, 350, 45, 170, 255,
  95, 300, 20, 200, 110, 340, 52, 182,
  224, 10, 78, 160, 246, 132, 315, 26
];
const _colorSlotByKey = new Map();
const _usedPaletteSlots = new Set();

const hashString = (s) => {
  let hash = 2166136261 >>> 0;
  for (let i = 0; i < s.length; i++) {
    hash ^= s.charCodeAt(i);
    hash = Math.imul(hash, 16777619) >>> 0;
  }
  return hash >>> 0;
};

function colorFromString(str) {
  const s = String(str ?? "").trim().toLowerCase() || "(blank)";
  if (_colorSlotByKey.has(s)) {
    const slot = _colorSlotByKey.get(s);
    const hue = DISTINCT_COLOR_HUES[slot % DISTINCT_COLOR_HUES.length];
    const light = getThemeName() === "light" ? 40 : 58;
    const sat = getThemeName() === "light" ? 72 : 78;
    return `hsl(${hue}, ${sat}%, ${light}%)`;
  }

  const hash = hashString(s);
  const size = DISTINCT_COLOR_HUES.length;
  // Probe through a curated palette so adjacent keys don't end up visually close.
  const base = hash % size;
  const step = 11; // coprime with palette size
  for (let i = 0; i < size; i++) {
    const idx = (base + i * step) % size;
    if (_usedPaletteSlots.has(idx)) continue;
    _usedPaletteSlots.add(idx);
    _colorSlotByKey.set(s, idx);
    const hue = DISTINCT_COLOR_HUES[idx];
    const light = getThemeName() === "light" ? 40 : 58;
    const sat = getThemeName() === "light" ? 72 : 78;
    return `hsl(${hue}, ${sat}%, ${light}%)`;
  }

  // Fallback if palette is exhausted.
  const hue = hash % 360;
  const sat = getThemeName() === "light" ? 72 : 78;
  const light = getThemeName() === "light" ? 40 : 58;
  return `hsl(${hue}, ${sat}%, ${light}%)`;
}

/* ================================
   Outline plugin
================================ */

const outlineBarsPlugin = {
  id: "outlineBars",
  afterDatasetsDraw(chart, args, pluginOptions) {
    const { ctx, chartArea } = chart;
    const opts = pluginOptions || {};
    const lineWidth = Number(opts.lineWidth ?? 1);
    const strokeStyle = opts.color ?? "rgba(0,0,0,1)";
    if (!chartArea) return;

    ctx.save();
    // Keep custom outlines inside the plotting area so zoom/filter clipping
    // does not paint over the legend/title region.
    ctx.beginPath();
    ctx.rect(chartArea.left, chartArea.top, chartArea.right - chartArea.left, chartArea.bottom - chartArea.top);
    ctx.clip();
    ctx.lineWidth = lineWidth;
    ctx.strokeStyle = strokeStyle;

    chart.data.datasets.forEach((ds, di) => {
      const meta = chart.getDatasetMeta(di);
      if (meta.hidden) return;

      meta.data.forEach((bar) => {
        const props = bar.getProps(["x", "y", "base", "width"], true);
        if (![props.x, props.y, props.base, props.width].every(Number.isFinite)) return;
        const left = props.x - props.width / 2;
        const top = Math.min(props.y, props.base);
        const height = Math.abs(props.base - props.y);

        ctx.strokeRect(
          Math.round(left) + 0.5,
          Math.round(top) + 0.5,
          Math.round(props.width),
          Math.round(height)
        );
      });
    });

    ctx.restore();
  }
};

const quotaShareGuidesPlugin = {
  id: "quotaShareGuides",
  afterDatasetsDraw(chart) {
    const { ctx, scales, chartArea } = chart;
    const yScale = scales?.y;
    if (!yScale || !chartArea) return;

    ctx.save();
    ctx.beginPath();
    ctx.rect(chartArea.left, chartArea.top, chartArea.right - chartArea.left, chartArea.bottom - chartArea.top);
    ctx.clip();
    ctx.setLineDash([3, 2]);
    ctx.lineWidth = 1;
    ctx.strokeStyle = "rgba(0,0,0,0.95)";

    chart.data.datasets.forEach((ds, di) => {
      const meta = chart.getDatasetMeta(di);
      if (meta.hidden) return;

      meta.data.forEach((bar, pi) => {
        const raw = ds.data?.[pi];
        if (!raw || !raw.isQuotaShare) return;

        const parts = Array.isArray(raw.participants) ? raw.participants : [];
        if (parts.length < 2) return;

        const props = bar.getProps(["x", "width"], true);
        const left = props.x - props.width / 2;
        const right = props.x + props.width / 2;

        let cumulative = Number(raw.attach || 0);
        for (let i = 0; i < parts.length - 1; i++) {
          cumulative += Number(parts[i]?.sliceLimit || 0);
          const py = yScale.getPixelForValue(cumulative);
          ctx.beginPath();
          ctx.moveTo(left, py);
          ctx.lineTo(right, py);
          ctx.stroke();
        }
      });
    });

    ctx.restore();
  }
};

const legendTitleBadgePlugin = {
  id: "legendTitleBadge",
  afterDraw(chartInstance) {
    const legend = chartInstance?.legend;
    const titleOpts = chartInstance?.options?.plugins?.legend?.title;
    const text = String(titleOpts?.text || "").trim();
    if (!legend || !titleOpts?.display || !text) return;

    const ctx = chartInstance.ctx;
    const theme = getThemeName();
    const fill = theme === "light" ? "rgba(30, 64, 175, 0.12)" : "rgba(96, 165, 250, 0.17)";
    const stroke = theme === "light" ? "rgba(30, 64, 175, 0.45)" : "rgba(147, 197, 253, 0.45)";
    const textColor = theme === "light" ? "rgba(15, 23, 42, 0.95)" : "rgba(255,255,255,0.96)";

    const fontSize = 12;
    const fontWeight = "600";
    const fontFamily = "system-ui, -apple-system, Segoe UI, Roboto, sans-serif";
    const hPad = 10;
    const vPad = 4;
    const radius = 8;

    ctx.save();
    ctx.font = `${fontWeight} ${fontSize}px ${fontFamily}`;
    const textWidth = Math.ceil(ctx.measureText(text).width);
    const badgeW = textWidth + hPad * 2;
    const badgeH = fontSize + vPad * 2;
    const x = Math.round(legend.left + (legend.width - badgeW) / 2);
    const y = Math.round(legend.top + 3);

    roundedRectPath(ctx, x, y, badgeW, badgeH, radius);
    ctx.fillStyle = fill;
    ctx.strokeStyle = stroke;
    ctx.lineWidth = 1;
    ctx.fill();
    ctx.stroke();

    ctx.fillStyle = textColor;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(text, x + badgeW / 2, y + badgeH / 2 + 0.5);
    ctx.restore();
  }
};

const legendFiltersPanelPlugin = {
  id: "legendFiltersPanel",
  afterDraw(chartInstance) {
    const legend = chartInstance?.legend;
    if (!legend || legend.options?.display === false) return;
    const lines = getLegendFilterLines();
    if (!lines.length) return;

    const ctx = chartInstance.ctx;
    const theme = getThemeName();
    const boxFill = theme === "light" ? "rgba(239, 246, 255, 0.97)" : "rgba(10, 25, 47, 0.97)";
    const boxStroke = theme === "light" ? "rgba(30, 64, 175, 0.30)" : "rgba(148, 163, 184, 0.30)";
    const textColor = theme === "light" ? "rgba(15, 23, 42, 0.95)" : "rgba(241, 245, 249, 0.95)";

    const fontSize = 10.5;
    const lineHeight = 14;
    const hPad = 8;
    const vPad = 7;
    const radius = 10;

    ctx.save();
    ctx.font = `500 ${fontSize}px system-ui, -apple-system, Segoe UI, Roboto, sans-serif`;
    const contentW = Math.ceil(Math.max(...lines.map((l) => ctx.measureText(l).width)));
    const panelW = Math.max(legend.width - 4, contentW + hPad * 2);
    const panelH = lines.length * lineHeight + vPad * 2;
    const x = Math.round(legend.left + 2);
    const y = Math.round(Math.min(
      chartInstance.height - panelH - 8,
      legend.top + legend.height + 10
    ));

    roundedRectPath(ctx, x, y, panelW, panelH, radius);
    ctx.fillStyle = boxFill;
    ctx.strokeStyle = boxStroke;
    ctx.lineWidth = 1;
    ctx.fill();
    ctx.stroke();

    ctx.fillStyle = textColor;
    ctx.textBaseline = "top";
    for (let i = 0; i < lines.length; i++) {
      ctx.fillText(lines[i], x + hPad, y + vPad + i * lineHeight);
    }
    ctx.restore();
  }
};

const boxValueLabelsPlugin = {
  id: "boxValueLabels",
  afterDatasetsDraw(chartInstance) {
    const { ctx } = chartInstance;
    const xScale = chartInstance.scales?.x;
    const yScale = chartInstance.scales?.y;
    if (!xScale || !yScale) return;

    ctx.save();
    ctx.font = "600 10px system-ui, -apple-system, Segoe UI, Roboto, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillStyle = "rgba(255,255,255,0.94)";
    ctx.strokeStyle = "rgba(11,17,27,0.8)";
    ctx.lineWidth = 2;

    chartInstance.data.datasets.forEach((ds, di) => {
      if (!ds || ds.datasetId === "sirOverlay" || ds.type === "line") return;
      const meta = chartInstance.getDatasetMeta(di);
      if (!meta || meta.hidden) return;

      meta.data.forEach((bar, idx) => {
        const raw = ds.data?.[idx];
        if (!raw) return;
        const attach = Number(raw.attach || 0);
        const top = Number(raw.top || 0);
        const lim = Number(raw.sumLimit ?? Math.max(0, top - attach));
        if (!Number.isFinite(lim) || lim <= 0) return;

        const props = bar.getProps(["x", "width"], true);
        const yTop = yScale.getPixelForValue(top);
        const yBottom = yScale.getPixelForValue(attach);
        const boxHeight = Math.abs(yBottom - yTop);
        const boxWidth = Number(props?.width || 0);

        // Keep labels only where they are likely to remain legible.
        if (boxHeight < 15 || boxWidth < 34) return;

        const label = compactMoney(lim);
        const maxTextWidth = Math.max(0, boxWidth - 8);
        const textWidth = ctx.measureText(label).width;
        if (textWidth > maxTextWidth) return;

        const x = Number(props.x || 0);
        const y = (yTop + yBottom) / 2;
        ctx.strokeText(label, x, y);
        ctx.fillText(label, x, y);
      });
    });

    ctx.restore();
  }
};

const yearAvailableTotalsPlugin = {
  id: "yearAvailableTotals",
  afterDatasetsDraw(chartInstance) {
    // Totals are rendered in page-level UI above the chart (outside canvas).
    return;
    if (_cache.showCoverageTotals === false) return;
    const xScale = chartInstance?.scales?.x;
    const yScale = chartInstance?.scales?.y;
    const chartArea = chartInstance?.chartArea;
    if (!xScale || !yScale || !chartArea) return;

    const selection = getSelectionSets();
    const byX = new Map(); // x -> { total, top }

    for (const s of _cache.slices || []) {
      const x = String(s?.x ?? "");
      if (!x) continue;
      if (!sliceMatchesSelection(s, selection)) continue;
      if (String(s?.availability || "").toLowerCase().includes("unavail")) continue;

      const limit = Number(s?.sliceLimit || 0);
      const attach = Number(s?.attach || 0);
      const top = attach + limit;
      if (!Number.isFinite(limit) || limit <= 0) continue;
      if (!Number.isFinite(top)) continue;

      if (!byX.has(x)) byX.set(x, { total: 0, top: 0 });
      const e = byX.get(x);
      e.total += limit;
      e.top = Math.max(e.top, top);
    }

    if (!byX.size) return;

    const theme = getThemeName();
    const textColor = theme === "light" ? "rgba(15, 23, 42, 0.96)" : "rgba(241, 245, 249, 0.96)";
    const pillBg = theme === "light" ? "rgba(255,255,255,0.88)" : "rgba(15,23,42,0.72)";
    const pillBorder = theme === "light" ? "rgba(15, 23, 42, 0.2)" : "rgba(241,245,249,0.24)";

    const ctx = chartInstance.ctx;
    ctx.save();
    ctx.font = "700 10px system-ui, -apple-system, Segoe UI, Roboto, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    const entries = Array.from(byX.entries())
      .map(([x, stat]) => ({
        stat,
        px: xScale.getPixelForValue(x),
        pyTop: yScale.getPixelForValue(Number(stat?.top || 0))
      }))
      .filter((e) => Number.isFinite(e.px) && Number.isFinite(e.pyTop))
      .sort((a, b) => a.px - b.px);

    let lastRight = Number.NEGATIVE_INFINITY;
    for (const { stat, px, pyTop } of entries) {
      if (!Number.isFinite(stat.total) || stat.total <= 0) continue;
      const label = compactMoney(stat.total);
      const textW = ctx.measureText(label).width;
      const padX = 4;
      const pillW = textW + padX * 2;
      const pillH = 14;
      const yCenter = Math.max(chartArea.top + pillH / 2 + 2, pyTop - 8);
      const left = px - pillW / 2;
      const right = px + pillW / 2;

      // Skip cramped labels to avoid unreadable overlaps.
      if (left <= lastRight + 6) continue;
      lastRight = right;

      roundedRectPath(ctx, left, yCenter - pillH / 2, pillW, pillH, 4);
      ctx.fillStyle = pillBg;
      ctx.fill();
      ctx.strokeStyle = pillBorder;
      ctx.lineWidth = 1;
      ctx.stroke();

      ctx.fillStyle = textColor;
      ctx.fillText(label, px, yCenter);
    }

    ctx.restore();
  }
};

/* ================================
   CSV Parser (robust)
================================ */

function parseCSV(text) {
  const rows = [];
  let row = [];
  let cur = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"' && next === '"') {
      cur += '"';
      i++;
      continue;
    }
    if (char === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (char === "," && !inQuotes) {
      row.push(cur.trim());
      cur = "";
      continue;
    }
    if (char === "\n" && !inQuotes) {
      row.push(cur.trim());
      rows.push(row);
      row = [];
      cur = "";
      continue;
    }
    if (char !== "\r") cur += char;
  }

  if (cur.length || row.length) {
    row.push(cur.trim());
    rows.push(row);
  }

  const rawHeaders = rows.shift() || [];
  const headers = rawHeaders.map(normalizeHeader);

  return rows
    .filter((r) => r.some((cell) => String(cell ?? "").trim() !== ""))
    .map((r) => {
      const obj = {};
      headers.forEach((h, i) => (obj[h] = (r[i] ?? "").trim()));

      // Known typo support
      if (obj["Attatchment Point"] && !obj["Attachment Point"]) {
        obj["Attachment Point"] = obj["Attatchment Point"];
      }

      return obj;
    });
}

async function fetchCSV(url) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error("Failed to load CSV: " + url);
  return parseCSV(await res.text());
}

async function fetchInsuranceProgramRows(primaryUrl) {
  try {
    return await fetchCSV(primaryUrl);
  } catch (err) {
    const fallbackUrl = "/data/OriginalFiles/tblInsuranceprogramid.csv";
    if (String(primaryUrl) === fallbackUrl) throw err;
    return await fetchCSV(fallbackUrl);
  }
}

/* ================================
   Availability classifier (heuristic)
================================ */

function classifyAvailability(policyRow, carrierRow) {
  // 1) Prefer an explicit solvency field from the carrier table if present
  // Accept a few common field names, but DO NOT rely on generic "Status" unless it clearly looks like solvency.
  const rawSolvency = String(
    getBy(carrierRow, "CarrierSolvency", "Solvency", "SolvencyStatus", "FinancialStatus")
  ).trim();

  if (rawSolvency) {
    const s = rawSolvency.toLowerCase();

    // "solvent" / "insolvent"
    if (s.includes("insolv") || s.includes("bankrupt") || s.includes("liquidat")) return "Unavailable";
    if (s.includes("solvent") || s.includes("active") || s.includes("good")) return "Available";

    // Sometimes people encode as 0/1
    if (s === "0" || s === "false" || s === "no" || s === "n") return "Unavailable";
    if (s === "1" || s === "true" || s === "yes" || s === "y") return "Available";
  }

  // 2) Policy-level collectible flag overrides (if present)
  const collectible = String(getBy(policyRow, "Collectible", "IsCollectible", "bcollectible")).trim();
  if (collectible !== "") {
    const v = collectible.toLowerCase();
    if (v === "0" || v === "false" || v === "no" || v === "n") return "Unavailable";
    if (v === "1" || v === "true" || v === "yes" || v === "y") return "Available";
  }

  // 3) Policy-level insolvent flags
  const insolvent = String(getBy(policyRow, "Insolvent", "IsInsolvent", "binsolvent")).trim();
  if (insolvent !== "") {
    const v = insolvent.toLowerCase();
    if (v === "1" || v === "true" || v === "yes" || v === "y") return "Unavailable";
  }

  // 4) Free-text hints in policy row
  const status = String(getBy(policyRow, "Status", "Availability", "AvailStatus")).toLowerCase();
  if (status.includes("insolv") || status.includes("bankrupt") || status.includes("unavail"))
    return "Unavailable";

  // 5) Free-text hints in carrier row (only if they look like solvency-ish text)
  const cstatus = String(getBy(carrierRow, "Status", "Availability")).toLowerCase();
  if (cstatus.includes("insolv") || cstatus.includes("bankrupt") || cstatus.includes("unavail"))
    return "Unavailable";

  return "Available";
}

/* ================================
   Build base "slices"
================================ */

function buildSlices({
  limitsRows,
  datesRows,
  policyRows,
  carrierRows,
  carrierGroupRows,
  insuranceProgramRows,
  policyLimitTypeRows,
  useYearAxis
}) {
  const policyDateMap = {};
  let minYear = Infinity;
  let maxYear = -Infinity;

  for (const row of datesRows) {
    const pid = String(getBy(row, "PolicyID", "Policy Id", "ID")).trim();
    if (!pid) continue;

    const start = String(getBy(row, "PStartDate", "PolicyStartDate", "StartDate")).trim();
    const end = String(getBy(row, "PEndDate", "PolicyEndDate", "EndDate")).trim();
    policyDateMap[pid] = { start, end };

    const startDate = parseDateToUTC(start);
    const endDate = parseDateToUTC(end);
    const startYear = startDate ? startDate.getUTCFullYear() : yearOf(start);
    const endYear = endDate ? endDate.getUTCFullYear() : yearOf(end);
    if (Number.isFinite(startYear)) {
      minYear = Math.min(minYear, startYear);
      maxYear = Math.max(maxYear, startYear);
    }
    if (Number.isFinite(endYear)) {
      minYear = Math.min(minYear, endYear);
      maxYear = Math.max(maxYear, endYear);
    }
  }

  const carrierRowById = {};
  const carrierNameById = {};
  for (const r of carrierRows) {
    const id = String(getBy(r, "CarrierID", "Carrier Id", "ID")).trim();
    if (!id) continue;
    carrierRowById[id] = r;
    const nm = String(getBy(r, "Carrier", "CarrierName", "Name", "Insurer", "Company")).trim();
    if (nm) carrierNameById[id] = nm;
  }

  const carrierGroupNameById = {};
  for (const r of carrierGroupRows) {
    const id = String(getBy(r, "CarrierGroupID", "Carrier Group ID", "ID")).trim();
    if (!id) continue;
    const nm = String(getBy(r, "CarrierGroup", "CarrierGroupName", "Name", "Group")).trim();
    if (nm) carrierGroupNameById[id] = nm;
  }

  const insuranceProgramNameById = {};
  for (const r of insuranceProgramRows || []) {
    const id = String(getBy(r, "InsuranceProgramID", "Insurance Program ID", "ID")).trim();
    if (!id) continue;
    const nm = String(getBy(r, "InsuranceProgram", "Program", "Name")).trim();
    if (nm) insuranceProgramNameById[id] = nm;
  }

  const policyLimitTypeNameById = {};
  for (const r of policyLimitTypeRows || []) {
    const id = String(
      getBy(r, "PolicyLimitTypeID", "Policy Limit Type ID", "Policy Limit ID", "ID", "Policy ID")
    ).trim();
    if (!id) continue;
    const nm = String(
      getBy(
        r,
        "PolicyLimitTypeName",
        "PolicyLimitType",
        "Policy Limit Type",
        "Policy Limity Type",
        "Type",
        "Name"
      )
    ).trim();
    if (nm) policyLimitTypeNameById[id] = nm;
  }

  const policyInfoById = {};
  for (const r of policyRows) {
    const pid = String(getBy(r, "PolicyID", "Policy Id", "ID")).trim();
    if (!pid) continue;

        const carrierId = String(getBy(r, "CarrierID", "Carrier Id")).trim();
        let carrierGroupId = String(getBy(r, "CarrierGroupID", "Carrier Group ID")).trim();

        // If Policy row lacks CarrierGroupID, try to get it from Carrier row
        if (!carrierGroupId && carrierId && carrierRowById[carrierId]) {
          carrierGroupId = String(getBy(carrierRowById[carrierId], "CarrierGroupID", "Carrier Group ID")).trim();
        }

        const carrierName =
          String(getBy(r, "Carrier", "CarrierName", "Insurer", "Company")).trim() ||
          (carrierId ? carrierNameById[carrierId] || "" : "");

        let carrierGroupName =
          String(getBy(r, "CarrierGroup", "Carrier Group", "CarrierGroupName")).trim();
        if (!carrierGroupName && carrierGroupId) {
          carrierGroupName = carrierGroupNameById[carrierGroupId] || "";
        }

        const policyNo = String(
          getBy(r, "policy_no", "PolicyNo", "Policy Number", "PolicyNumber", "PolicyNum")
        ).trim();
        const insuranceProgramId = String(
          getBy(r, "InsuranceProgramID", "Insurance Program ID")
        ).trim();
        const namedInsuredId = String(
          getBy(r, "NamedInsuredID", "Named Insured ID")
        ).trim();
        const insuranceProgram =
          String(getBy(r, "InsuranceProgram", "Program", "ProgramName")).trim() ||
          (insuranceProgramId ? insuranceProgramNameById[insuranceProgramId] || "" : "");
        const sirPerOcc = num(getBy(r, "SIRPerOcc", "SIR Per Occ", "SIR"));
        const sirAggregate = num(getBy(r, "SIRAggregate", "SIR Aggregate"));

        const cRow = carrierId ? carrierRowById[carrierId] : null;
        const availability = classifyAvailability(r, cRow);

        policyInfoById[pid] = {
          policy_no: policyNo,
          carrier: carrierName || "(unknown carrier)",
          carrierGroup: carrierGroupName || "(unknown group)",
          insuranceProgramId,
          insuranceProgram: insuranceProgram || "(unknown program)",
          namedInsuredId,
          sirPerOcc,
          sirAggregate,
          availability,
        };
  }

  const slices = [];
  for (const r of limitsRows) {
    const pid = String(getBy(r, "PolicyID", "Policy Id", "ID")).trim();
    if (!pid) continue;

    const dates = policyDateMap[pid];
    if (!dates || !dates.start || !dates.end) continue;

    const startDate = parseDateToUTC(dates.start);
    const endDate = parseDateToUTC(dates.end);
    const startYear = startDate ? startDate.getUTCFullYear() : yearOf(dates.start);
    const endYear = endDate ? endDate.getUTCFullYear() : yearOf(dates.end);
    if (!Number.isFinite(startYear) || !Number.isFinite(endYear)) continue;
    const policyStartYear = Math.min(startYear, endYear);
    const policyEndYear = Math.max(startYear, endYear);
    const policyStartMs = startDate
      ? startDate.getTime()
      : startOfYearUTC(policyStartYear);
    const policyEndMs = endDate
      ? endDate.getTime() + (24 * 60 * 60 * 1000 - 1)
      : endOfYearUTC(policyEndYear);

    const x = useYearAxis ? String(policyStartYear) : `${dates.start} to ${dates.end}`;

    const attachRaw = firstPresentNum(
      getBy(r, "Attachment Point", "AttachmentPoint", "attach", "Attatchment Point")
    );

    const sliceLimitRaw = firstPresentNum(
      getBy(r, "LayerPerOccLimit", "Layer Per Occ Limit", "layerperocclim"),
      getBy(r, "PerOccLimit", "Per Occ Limit", "perocclim")
    );

    if (sliceLimitRaw <= 0) continue;

    const attach = Math.round(attachRaw);
    const sliceLimit = Math.round(sliceLimitRaw);
    const policyLimitTypeId = String(
      getBy(r, "PolicyLimitTypeID", "Policy Limit Type ID", "Policy Limit ID")
    ).trim();
    const policyLimitType = policyLimitTypeNameById[policyLimitTypeId] || policyLimitTypeId;

    const info = policyInfoById[pid] || {
      policy_no: "",
      carrier: "(unknown carrier)",
      carrierGroup: "(unknown group)",
      insuranceProgramId: "",
      insuranceProgram: "(unknown program)",
      namedInsuredId: "",
      sirPerOcc: 0,
      sirAggregate: 0,
      availability: "Available"
    };

    slices.push({
      x,
      year: policyStartYear,
      policyStartYear,
      policyEndYear,
      policyStartMs,
      policyEndMs,
      attach,
      sliceLimit,
      PolicyID: pid,
      policy_no: info.policy_no,
      carrier: info.carrier,
      carrierGroup: info.carrierGroup,
      insuranceProgramId: info.insuranceProgramId,
      insuranceProgram: info.insuranceProgram,
      namedInsuredId: info.namedInsuredId,
      sirPerOcc: Number(info.sirPerOcc || 0),
      sirAggregate: Number(info.sirAggregate || 0),
      policyLimitTypeId,
      policyLimitType,
      availability: info.availability
    });
  }

  let xLabels = [];
  if (useYearAxis && Number.isFinite(minYear) && Number.isFinite(maxYear) && minYear <= maxYear) {
    for (let y = minYear; y <= maxYear; y++) xLabels.push(String(y));
  } else {
    xLabels = [...new Set(slices.map((s) => s.x))].sort();
  }

  return { slices, xLabels };
}

/* ================================
   Quota-share detection
================================ */

function quotaGroupKey(s) {
  // Quota-share is scoped within the same program/year/layer/type (+ named insured).
  const program = String(s?.insuranceProgramId || s?.insuranceProgram || "").trim();
  const year = Number.isFinite(s?.year) ? String(s.year) : String(s?.x || "").trim();
  const attach = String(s?.attach ?? "").trim();
  const limitType = String(s?.policyLimitTypeId || "").trim();
  const namedInsured = String(s?.namedInsuredId || "").trim();
  return `${program}||${year}||${attach}||${limitType}||${namedInsured}`;
}

function buildQuotaKeySet(slices) {
  const byQuotaKey = new Map(); // quotaGroupKey -> Set(PolicyID)
  for (const s of slices) {
    const k = quotaGroupKey(s);
    if (!byQuotaKey.has(k)) byQuotaKey.set(k, new Set());
    byQuotaKey.get(k).add(String(s.PolicyID));
  }
  const quotaKeySet = new Set();
  for (const [k, set] of byQuotaKey.entries()) {
    if (set.size > 1) quotaKeySet.add(k);
  }
  return quotaKeySet;
}

function applyFiltersToCache() {
  const { allSlices, allXLabels, useYearAxis, filters } = _cache;
  const startYear = Number.isFinite(filters.startYear) ? filters.startYear : null;
  const endYear = Number.isFinite(filters.endYear) ? filters.endYear : null;
  const startDate = parseDateToUTC(filters.startDate);
  const endDate = parseDateToUTC(filters.endDate);
  const selectedProgram = String(filters.insuranceProgram || "").trim();
  const selectedPolicyLimitType = String(filters.policyLimitType || "").trim();

  let filteredSlices = allSlices.slice();
  if (useYearAxis && (startYear !== null || endYear !== null || startDate || endDate)) {
    const filterStartMs = startDate
      ? startDate.getTime()
      : startYear !== null
      ? startOfYearUTC(startYear)
      : Number.NEGATIVE_INFINITY;
    const filterEndMs = endDate
      ? endDate.getTime() + (24 * 60 * 60 * 1000 - 1)
      : endYear !== null
      ? endOfYearUTC(endYear)
      : Number.POSITIVE_INFINITY;
    filteredSlices = filteredSlices.filter((s) => {
      const policyStartMs = Number(s?.policyStartMs);
      const policyEndMs = Number(s?.policyEndMs);
      if (!Number.isFinite(policyStartMs) || !Number.isFinite(policyEndMs)) return false;
      return policyStartMs <= filterEndMs && policyEndMs >= filterStartMs;
    });
  }

  if (selectedProgram) {
    filteredSlices = filteredSlices.filter(
      (s) => String(s?.insuranceProgram || "").trim() === selectedProgram
    );
  }

  if (selectedPolicyLimitType) {
    filteredSlices = filteredSlices.filter(
      (s) => String(s?.policyLimitType || "").trim() === selectedPolicyLimitType
    );
  }

  let filteredXLabels = allXLabels.slice();
  if (useYearAxis && filteredSlices.length > 0) {
    const minSliceYear = Math.min(
      ...filteredSlices
        .map((s) => Number(s?.policyStartYear))
        .filter((y) => Number.isFinite(y))
    );
    const maxSliceYear = Math.max(
      ...filteredSlices
        .map((s) => Number(s?.policyEndYear))
        .filter((y) => Number.isFinite(y))
    );
    if (Number.isFinite(minSliceYear) && Number.isFinite(maxSliceYear) && minSliceYear <= maxSliceYear) {
      filteredXLabels = [];
      for (let y = minSliceYear; y <= maxSliceYear; y++) filteredXLabels.push(String(y));
    } else if (startYear !== null || endYear !== null) {
      filteredXLabels = [];
    }
  } else if (useYearAxis && (startYear !== null || endYear !== null || startDate || endDate)) {
    filteredXLabels = [];
  }

  _cache.slices = filteredSlices;
  _cache.xLabels = filteredXLabels;
}

function normalizeStringList(values) {
  if (!Array.isArray(values)) return [];
  return [...new Set(values.map((v) => String(v ?? "").trim()).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b)
  );
}

function getSelectionSets() {
  const carriers = new Set(normalizeStringList(_cache.filters?.carriers));
  const carrierGroups = new Set(normalizeStringList(_cache.filters?.carrierGroups));
  return { carriers, carrierGroups, active: carriers.size > 0 || carrierGroups.size > 0 };
}

function sliceMatchesSelection(slice, selection) {
  if (!selection?.active) return true;
  const carrierOk = selection.carriers.size === 0 || selection.carriers.has(String(slice?.carrier || ""));
  const groupOk =
    selection.carrierGroups.size === 0 || selection.carrierGroups.has(String(slice?.carrierGroup || ""));
  return carrierOk && groupOk;
}


function getChartSurfaceWidthPx() {
  const viewport = _cache.dom?.viewport;
  const labels = Array.isArray(_cache.xLabels) ? _cache.xLabels.length : 0;
  const baseViewportWidth = Math.max(320, viewport?.clientWidth || 0);
  const pxPerLabel = Math.max(20, 20 * (_cache.xZoom || 1));
  const target = Math.max(baseViewportWidth, labels * pxPerLabel + 80);
  return Math.round(target);
}

function syncChartViewportWidth({ anchorClientX } = {}) {
  const surface = _cache.dom?.surface;
  const viewport = _cache.dom?.viewport;
  if (!surface || !viewport) return;

  const prevWidth = surface.getBoundingClientRect().width || surface.clientWidth || 1;
  const nextWidth = getChartSurfaceWidthPx();
  const prevScrollLeft = viewport.scrollLeft;

  let anchorRatio = null;
  if (Number.isFinite(anchorClientX)) {
    const rect = viewport.getBoundingClientRect();
    const xInViewport = anchorClientX - rect.left;
    anchorRatio = (prevScrollLeft + xInViewport) / prevWidth;
  }

  surface.style.width = `${nextWidth}px`;

  if (anchorRatio !== null) {
    const rect = viewport.getBoundingClientRect();
    const xInViewport = anchorClientX - rect.left;
    const nextScroll = anchorRatio * nextWidth - xInViewport;
    viewport.scrollLeft = clamp(nextScroll, 0, Math.max(0, nextWidth - viewport.clientWidth));
  } else if (nextWidth <= viewport.clientWidth) {
    viewport.scrollLeft = 0;
  } else {
    viewport.scrollLeft = clamp(prevScrollLeft, 0, Math.max(0, nextWidth - viewport.clientWidth));
  }
}

function bindViewportInteractions() {
  const viewport = _cache.dom?.viewport;
  const canvas = _cache.dom?.canvas;
  if (!viewport || !canvas || _cache._wheelBound) return;

  const onWheel = (e) => {
    // Shift + wheel: horizontal scroll.
    if (e.shiftKey && Math.abs(e.deltaY) >= Math.abs(e.deltaX)) {
      e.preventDefault();
      viewport.scrollLeft += e.deltaY;
      return;
    }

    // Default wheel: zoom x-axis around cursor.
    e.preventDefault();
    const delta = Math.abs(e.deltaY) > 0 ? e.deltaY : e.deltaX;
    const factor = Math.exp(-delta * 0.0015);
    _cache.xZoom = clamp((_cache.xZoom || 1) * factor, 0.45, 4);
    syncChartViewportWidth({ anchorClientX: e.clientX });
  };

  // Bind to both container and canvas for browser consistency.
  viewport.addEventListener("wheel", onWheel, { passive: false });
  canvas.addEventListener("wheel", onWheel, { passive: false });

  _cache._wheelBound = true;
}

function rebuildChart() {
  if (!chart || !_cache.options) return;

  const { barThickness, categorySpacing } = _cache.options;

  chart.data.labels = _cache.xLabels;
  const barDatasets = buildDatasetsForView({
    slices: _cache.slices,
    xLabels: _cache.xLabels,
    view: currentView,
    barThickness,
    categorySpacing,
    quotaKeySet: _cache.quotaKeySet
  });
  const sirDataset = buildSirDataset({
    slices: _cache.slices,
    xLabels: _cache.xLabels,
    sirMode: _cache.filters?.sirMode
  });
  chart.data.datasets = sirDataset ? [...barDatasets, sirDataset] : barDatasets;

  const y = chart.options?.scales?.y;
  if (y) {
    y.min = Number.isFinite(_cache.filters.zoomMin) ? _cache.filters.zoomMin : undefined;
    y.max = Number.isFinite(_cache.filters.zoomMax) ? _cache.filters.zoomMax : undefined;
  }
  if (chart.options?.plugins?.legend?.title) {
    chart.options.plugins.legend.title.text = legendTitleForView(currentView);
  }

  chart.update();
  syncChartViewportWidth();
}

function applyChartTheme(themeName = getThemeName()) {
  if (!chart) return;
  const c = getChartThemeColors(themeName);
  chart.options.plugins.outlineBars.color = c.outline;
  chart.options.plugins.legend.labels.color = c.legendText;
  chart.options.plugins.legend.title.color = c.legendText;
  chart.options.scales.x.grid.color = c.xGrid;
  chart.options.scales.x.ticks.color = c.axisTicks;
  chart.options.scales.y.grid.color = c.yGrid;
  chart.options.scales.y.ticks.color = c.axisTicks;
  chart.options.scales.y.title.color = c.yTitle;
  const sirDs = chart.data.datasets?.find((ds) => ds?.datasetId === "sirOverlay");
  if (sirDs) {
    const sirColor = _cache.filters?.sirMode === "aggregate" ? "#f97316" : "#facc15";
    sirDs.borderColor = sirColor;
    sirDs.pointBackgroundColor = sirColor;
    sirDs.pointBorderColor = sirColor;
  }
  chart.update();
}

/* ================================
   View aggregation -> datasets
================================ */

function buildDatasetsForView({ slices, xLabels, view, barThickness, categorySpacing, quotaKeySet }) {
  const qaKey = (s) => quotaGroupKey(s);
  const isQuotaSlice = (s) => quotaKeySet && quotaKeySet.has(qaKey(s));
  const isUnavailable = (s) => String(s?.availability || "").toLowerCase().includes("unavail");
  const selection = getSelectionSets();

  const keyOf = (s) => {
    // Collapse all unavailable slices into a single dataset/legend item in every view.
    if (isUnavailable(s)) return "Unavailable";

    if (view === "carrier" || view === "carrierGroup") {
      // Prevent gaps: force quota layers into one dataset for both views.
      if (isQuotaSlice(s)) return "Quota share";
      return view === "carrier" ? (s.carrier || "(unknown carrier)") : (s.carrierGroup || "(unknown group)");
    }
    return s.availability || "Available";
  };

  // Carrier -> CarrierGroup mapping for coloring (only for non-quota datasets)
  const carrierToGroup = {};
  if (view === "carrier") {
    for (const s of slices) {
      const c = s.carrier || "(unknown carrier)";
      if (!carrierToGroup[c]) carrierToGroup[c] = s.carrierGroup || "(unknown group)";
    }
  }

  const layerMap = new Map(); // Non-quota: one layer per policy. Quota-share: grouped by quota key.
  const groups = new Set();

  for (const s of slices) {
    const group = keyOf(s);
    groups.add(group);

    const sliceQuotaKey = qaKey(s);
    const k = group === "Quota share"
      ? `${group}||${s.x}||${s.attach}||${sliceQuotaKey}`
      : `${group}||${s.x}||${s.attach}||${String(s.PolicyID)}`;
    if (!layerMap.has(k)) {
      layerMap.set(k, {
        group,
        x: s.x,
        attach: s.attach,
        quotaGroupKey: group === "Quota share" ? sliceQuotaKey : "",
        sumLimit: 0,
        participants: [],
        hasSelectionMatch: false
      });
    }

    const e = layerMap.get(k);
    e.sumLimit += s.sliceLimit;
    if (sliceMatchesSelection(s, selection)) e.hasSelectionMatch = true;
    e.participants.push({
      pid: s.PolicyID,
      carrier: s.carrier,
      carrierGroup: s.carrierGroup,
      availability: s.availability,
      policy_no: s.policy_no,
      sliceLimit: s.sliceLimit,
      sirPerOcc: Number(s.sirPerOcc || 0),
      sirAggregate: Number(s.sirAggregate || 0),
      quotaGroupKey: sliceQuotaKey
    });
  }

  const groupList = Array.from(groups);
  if (view === "availability") {
    groupList.sort((a, b) => {
      const order = (v) => (String(v).toLowerCase().includes("unavail") ? 1 : 0);
      return order(a) - order(b);
    });
  } else if (view === "carrier" || view === "carrierGroup") {
    groupList.sort((a, b) => {
      if (a === "Quota share" && b !== "Quota share") return -1;
      if (b === "Quota share" && a !== "Quota share") return 1;
      return String(a).localeCompare(String(b));
    });
  } else {
    groupList.sort((a, b) => String(a).localeCompare(String(b)));
  }

  const labelIndex = new Map(xLabels.map((lbl, i) => [lbl, i]));

  const requestedBarThickness = Number(barThickness);
  const useFlexBarWidth = !Number.isFinite(requestedBarThickness);
  // Chart.js can produce NaN bar widths for floating + non-grouped bars when using
  // `barThickness: "flex"` (especially after responsive width changes).
  // Use a stable numeric fallback instead and size it to the visible category slot.
  const liveXScaleWidth = Number(chart?.scales?.x?.width);
  const approxSurfaceWidth = getChartSurfaceWidthPx();
  const approxPlotWidth = Math.max(120, approxSurfaceWidth - 360);
  const slotWidth = xLabels.length
    ? (Number.isFinite(liveXScaleWidth) ? liveXScaleWidth / xLabels.length : approxPlotWidth / xLabels.length)
    : 20;
  // Fill the category slot so columns touch, while avoiding overflow from rounding.
  const autoBarThickness = Math.max(8, Math.floor(slotWidth));
  const barThicknessPx = useFlexBarWidth ? autoBarThickness : requestedBarThickness;
  const quotaGradientForCarrierView = (context, pointRaw) => {
    const parts = Array.isArray(pointRaw?.participants) ? pointRaw.participants : [];
    const weightedParts = parts
      .map((p) => ({
        limit: Number(p?.sliceLimit || 0),
        color: colorFromString(p?.carrier || "Quota share")
      }))
      .filter((p) => Number.isFinite(p.limit) && p.limit > 0);
    if (!weightedParts.length) return colorFromString("Quota share");

    const total = weightedParts.reduce((sum, p) => sum + p.limit, 0);
    if (!Number.isFinite(total) || total <= 0) return weightedParts[0].color;

    const chartInstance = context?.chart;
    const yScale = chartInstance?.scales?.y;
    const canvasCtx = chartInstance?.ctx;
    if (!yScale || !canvasCtx) return weightedParts[0].color;

    const top = Number(pointRaw?.top || 0);
    const attach = Number(pointRaw?.attach || 0);
    const yTop = yScale.getPixelForValue(top);
    const yBottom = yScale.getPixelForValue(attach);
    if (!Number.isFinite(yTop) || !Number.isFinite(yBottom)) return weightedParts[0].color;
    if (yTop === yBottom) return weightedParts[0].color;
    const gradient = canvasCtx.createLinearGradient(0, yBottom, 0, yTop);

    let running = 0;
    for (const p of weightedParts) {
      const start = clamp(running / total, 0, 1);
      running += p.limit;
      const end = clamp(running / total, 0, 1);
      gradient.addColorStop(start, p.color);
      gradient.addColorStop(end, p.color);
    }

    return gradient;
  };

  return groupList.map((group) => {
    const points = [];

    for (const e of layerMap.values()) {
      if (e.group !== group) continue;

      const top = e.attach + e.sumLimit;

      e.participants.sort((a, b) => {
        if (b.sliceLimit !== a.sliceLimit) return b.sliceLimit - a.sliceLimit;
        return String(a.carrier || "").localeCompare(String(b.carrier || ""));
      });

      const isQuotaShare = !!(e.quotaGroupKey && quotaKeySet && quotaKeySet.has(e.quotaGroupKey));

      points.push({
        x: e.x,
        y: [e.attach, top],
        attach: e.attach,
        top,
        sumLimit: e.sumLimit,
        participants: e.participants,
        group,
        isQuotaShare,
        quotaGroupKey: e.quotaGroupKey || "",
        isHighlighted: !selection.active || !!e.hasSelectionMatch
      });
    }

    points.sort((a, b) => {
      const ax = labelIndex.get(a.x) ?? 1e9;
      const bx = labelIndex.get(b.x) ?? 1e9;
      if (ax !== bx) return ax - bx;
      return a.attach - b.attach;
    });

    let bg;
    if (group === "Unavailable") {
      bg = "#94a3b8";
    } else if (view === "availability") {
      bg = String(group).toLowerCase().includes("unavail") ? "#888888" : "#22c55e";
    } else if (view === "carrier") {
      if (group === "Quota share") {
        bg = colorFromString("Quota share");
      } else {
        bg = colorFromString(group);
      }
    } else if (view === "carrierGroup") {
      bg = colorFromString(group);
    } else {
      bg = colorFromString(group);
    }

    const hasHighlightedPoint = points.some((p) => p.isHighlighted);
    const datasetInFocus = !selection.active || hasHighlightedPoint;
    const mutedFill = "rgba(255,255,255,0.92)";
    const mutedBorder = "rgba(0,0,0,0.9)";
    const isQuotaDataset = group === "Quota share";

    return {
      label: group,
      data: points,
      parsing: { xAxisKey: "x", yAxisKey: "y" },

      grouped: false,
      barThickness: barThicknessPx,
      maxBarThickness: barThicknessPx,

      categoryPercentage: 1.0,
      barPercentage: 1.0,
      inflateAmount: 0.6,

      backgroundColor: (context) => {
        if (!datasetInFocus) return mutedFill;
        const raw = context?.raw || null;
        if (view === "carrier" && isQuotaDataset && raw?.isQuotaShare) {
          return quotaGradientForCarrierView(context, raw);
        }
        return bg;
      },
      borderColor: datasetInFocus ? "rgba(11, 17, 27, 0.62)" : mutedBorder,
      borderWidth: 1,
      borderRadius: 2,
      borderSkipped: false
    };
  });
}

function getSirValueFromSlice(slice, sirMode) {
  if (sirMode === "aggregate") return Number(slice?.sirAggregate || 0);
  if (sirMode === "perOcc") return Number(slice?.sirPerOcc || 0);
  return 0;
}

function buildSirDataset({ slices, xLabels, sirMode }) {
  if (!sirMode || sirMode === "off") return null;
  const mode = sirMode === "aggregate" ? "aggregate" : "perOcc";
  const byX = new Map();
  for (const s of slices || []) {
    const x = String(s?.x ?? "");
    if (!x) continue;
    const v = getSirValueFromSlice(s, mode);
    if (!Number.isFinite(v) || v <= 0) continue;
    byX.set(x, Math.max(byX.get(x) || 0, v));
  }

  const color = mode === "aggregate" ? "#f97316" : "#facc15";
  const data = (xLabels || []).map((x) => {
    const v = byX.get(String(x));
    return { x: String(x), y: Number.isFinite(v) ? v : null };
  });

  return {
    datasetId: "sirOverlay",
    type: "line",
    label: mode === "aggregate" ? "SIR (Aggregate)" : "SIR (Per Occ)",
    data,
    parsing: { xAxisKey: "x", yAxisKey: "y" },
    spanGaps: true,
    borderColor: color,
    pointBackgroundColor: color,
    pointBorderColor: color,
    borderWidth: 2,
    pointRadius: 2,
    pointHoverRadius: 4,
    tension: 0.2,
    fill: false,
    yAxisID: "y",
    order: 1000
  };
}

/* ================================
   Export helpers
================================ */

function toDateStamp(d = new Date()) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}${m}${day}`;
}

function csvCell(v) {
  const s = String(v ?? "");
  if (!/[",\n]/.test(s)) return s;
  return `"${s.replace(/"/g, "\"\"")}"`;
}

function triggerDownload(href, filename) {
  const a = document.createElement("a");
  a.href = href;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

function triggerBlobDownload(content, filename, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  triggerDownload(url, filename);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function loadScriptOnce(src, isReady) {
  if (isReady()) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const existing = Array.from(document.querySelectorAll("script")).find((s) => s.src === src);
    if (existing) {
      existing.addEventListener("load", () => resolve(), { once: true });
      existing.addEventListener("error", () => reject(new Error(`Failed to load script: ${src}`)), { once: true });
      return;
    }

    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`Failed to load script: ${src}`));
    document.head.appendChild(script);
  });
}

async function ensurePdfLibs() {
  await loadScriptOnce(
    "https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js",
    () => typeof window.html2canvas === "function"
  );
  await loadScriptOnce(
    "https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js",
    () => !!window.jspdf?.jsPDF
  );
}

function getFilterMeta() {
  const f = _cache.filters || {};
  const yearRange =
    f.startDate || f.endDate
      ? `${String(f.startDate || "All")}-${String(f.endDate || "All")}`
      : Number.isFinite(f.startYear) || Number.isFinite(f.endYear)
      ? `${Number.isFinite(f.startYear) ? f.startYear : "All"}-${Number.isFinite(f.endYear) ? f.endYear : "All"}`
      : "All";
  const carriers = normalizeStringList(f.carriers);
  const carrierGroups = normalizeStringList(f.carrierGroups);
  const viewLabel = currentView === "carrierGroup"
    ? "Carrier Group"
    : currentView === "availability"
      ? "Availability"
      : "Carrier";

  return {
    viewLabel,
    view: currentView,
    insuranceProgram: String(f.insuranceProgram || "").trim() || "All",
    policyLimitType: String(f.policyLimitType || "").trim() || "All",
    yearRange,
    carriers: carriers.length ? carriers.join(", ") : "All",
    carrierGroups: carrierGroups.length ? carrierGroups.join(", ") : "All"
  };
}

function getExportFilterLines(meta = getFilterMeta()) {
  return [
    `View: ${meta.viewLabel} | Insurance Program: ${meta.insuranceProgram} | Policy Limit Type: ${meta.policyLimitType}`,
    `Year Range: ${meta.yearRange} | Carriers: ${meta.carriers} | Carrier Groups: ${meta.carrierGroups}`
  ];
}

function getFilteredSliceRows() {
  return (_cache.slices || []).map((s) => ({
    Year: Number.isFinite(s.year) ? s.year : s.x,
    InsuranceProgram: s.insuranceProgram || "",
    PolicyLimitType: s.policyLimitType || s.policyLimitTypeId || "",
    Carrier: s.carrier || "",
    CarrierGroup: s.carrierGroup || "",
    Availability: s.availability || "",
    Attachment: Number(s.attach || 0),
    LayerLimit: Number(s.sliceLimit || 0),
    PolicyNumber: s.policy_no || "",
    PolicyID: s.PolicyID || ""
  }));
}

function getAggregatedReportRows() {
  if (!chart?.data?.datasets) return [];
  const rows = [];
  for (const ds of chart.data.datasets) {
    if (ds?.datasetId === "sirOverlay" || ds?.type === "line") continue;
    const group = String(ds?.label || "");
    for (const p of ds?.data || []) {
      const limit = Number(p?.sumLimit ?? (Number(p?.top || 0) - Number(p?.attach || 0)));
      const participants = Array.isArray(p?.participants) ? p.participants.length : 0;
      rows.push({
        Year: p?.x ?? "",
        Group: group,
        Attachment: Number(p?.attach || 0),
        TotalLimit: Number.isFinite(limit) ? limit : 0,
        Participants: participants
      });
    }
  }

  rows.sort((a, b) => {
    const ay = Number(a.Year);
    const by = Number(b.Year);
    if (Number.isFinite(ay) && Number.isFinite(by) && ay !== by) return ay - by;
    const yCmp = String(a.Year).localeCompare(String(b.Year));
    if (yCmp !== 0) return yCmp;
    if (a.Attachment !== b.Attachment) return a.Attachment - b.Attachment;
    return String(a.Group).localeCompare(String(b.Group));
  });
  return rows;
}

/**
 * Export chart canvas as PNG using the current filtered/rendered state.
 */
export function exportChartAsPNG() {
  if (!chart) throw new Error("Chart is not initialized");
  const meta = getFilterMeta();
  const stamp = toDateStamp();
  const file = `CoverageTower_${stamp}_${currentView}.png`;
  const srcCanvas = chart.canvas;
  const theme = getThemeName();
  const bg = theme === "light" ? "#ffffff" : "#0f1720";
  const fg = theme === "light" ? "#0f172a" : "#f8fafc";
  const subFg = theme === "light" ? "rgba(15,23,42,0.88)" : "rgba(248,250,252,0.9)";

  const width = srcCanvas.width;
  const scale = Math.max(0.85, Math.min(1.6, width / 1200));
  const titleSize = Math.round(26 * scale);
  const textSize = Math.round(14 * scale);
  const sidePad = Math.round(24 * scale);
  const topPad = Math.round(20 * scale);
  const lineGap = Math.round(8 * scale);
  const metaLines = getExportFilterLines(meta);
  const headerHeight = Math.round(topPad + titleSize + 2 * (textSize + lineGap) + 20 * scale);
  const outCanvas = document.createElement("canvas");
  outCanvas.width = width;
  outCanvas.height = headerHeight + srcCanvas.height;
  const ctx = outCanvas.getContext("2d");

  ctx.fillStyle = bg;
  ctx.fillRect(0, 0, outCanvas.width, outCanvas.height);
  ctx.drawImage(srcCanvas, 0, headerHeight);

  let y = topPad + titleSize;
  ctx.fillStyle = fg;
  ctx.font = `700 ${titleSize}px system-ui, -apple-system, Segoe UI, Roboto, sans-serif`;
  ctx.textAlign = "left";
  ctx.textBaseline = "alphabetic";
  ctx.fillText("Insurance Program Coverage Tower", sidePad, y);

  ctx.fillStyle = subFg;
  ctx.font = `500 ${textSize}px system-ui, -apple-system, Segoe UI, Roboto, sans-serif`;
  y += lineGap + textSize;
  ctx.fillText(metaLines[0], sidePad, y);
  y += lineGap + textSize;
  ctx.fillText(metaLines[1], sidePad, y);

  const dataUrl = outCanvas.toDataURL("image/png", 1);
  triggerDownload(dataUrl, file);
  console.log(`[Export] PNG saved: ${file}`);
}

/**
 * Export currently filtered slices as CSV rows.
 */
export function exportFilteredCSV() {
  const rows = getFilteredSliceRows();
  const cols = [
    "Year",
    "InsuranceProgram",
    "PolicyLimitType",
    "Carrier",
    "CarrierGroup",
    "Availability",
    "Attachment",
    "LayerLimit",
    "PolicyNumber",
    "PolicyID"
  ];
  const lines = [cols.join(",")];
  for (const r of rows) lines.push(cols.map((c) => csvCell(r[c])).join(","));
  const file = `CoverageTower_FilteredData_${toDateStamp()}.csv`;
  triggerBlobDownload(lines.join("\n"), file, "text/csv;charset=utf-8");
  console.log(`[Export] CSV saved: ${file} (rows=${rows.length})`);
}

/**
 * Export a 2-page PDF report:
 * page 1 = metadata + chart image
 * page 2+ = aggregated layer table
 */
export async function exportReportPDF() {
  if (!chart || !_cache.dom?.canvas) throw new Error("Chart is not initialized");
  await ensurePdfLibs();

  const html2canvas = window.html2canvas;
  const jsPDF = window.jspdf.jsPDF;

  const meta = getFilterMeta();
  const stamp = toDateStamp();
  const filename = `CoverageTower_Report_${stamp}.pdf`;

  const renderCanvas = await html2canvas(_cache.dom.canvas, {
    scale: 2,
    backgroundColor: getThemeName() === "light" ? "#ffffff" : "#0f1720",
    useCORS: true
  });
  const chartImg = renderCanvas.toDataURL("image/png");

  const pdf = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
  const pageW = pdf.internal.pageSize.getWidth();
  const pageH = pdf.internal.pageSize.getHeight();
  const margin = 36;

  // Page 1: summary + image
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(18);
  pdf.text("Insurance Program Coverage Tower", margin, 42);

  pdf.setFont("helvetica", "normal");
  pdf.setFontSize(10);
  const subtitle = getExportFilterLines(meta).join("    ");
  const subtitleLines = pdf.splitTextToSize(subtitle, pageW - margin * 2);
  pdf.text(subtitleLines, margin, 62);

  const generated = `Generated: ${new Date().toLocaleString()}`;
  pdf.text(generated, margin, 78 + (subtitleLines.length - 1) * 12);

  const imgTop = 98 + (subtitleLines.length - 1) * 12;
  const imgMaxW = pageW - margin * 2;
  const imgMaxH = pageH - imgTop - 50;
  const imgRatio = renderCanvas.width / renderCanvas.height || 1;
  let imgW = imgMaxW;
  let imgH = imgW / imgRatio;
  if (imgH > imgMaxH) {
    imgH = imgMaxH;
    imgW = imgH * imgRatio;
  }
  pdf.addImage(chartImg, "PNG", margin, imgTop, imgW, imgH, undefined, "FAST");

  pdf.setFontSize(9);
  pdf.text("Generated by Coverage Dashboard", margin, pageH - 18);

  // Page 2: table summary
  const rows = getAggregatedReportRows();
  pdf.addPage("a4", "landscape");
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(14);
  pdf.text("Aggregated Layer Summary", margin, 38);

  const headers = ["Year", "Group", "Attachment", "Total Limit", "Number of Participants"];
  const widths = [70, 300, 110, 110, 130];
  const rowH = 16;
  let y = 58;

  const drawHeader = () => {
    let x = margin;
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(9);
    for (let i = 0; i < headers.length; i++) {
      pdf.text(headers[i], x + 2, y);
      x += widths[i];
    }
    pdf.setLineWidth(0.5);
    pdf.line(margin, y + 3, margin + widths.reduce((a, b) => a + b, 0), y + 3);
    y += rowH;
  };

  drawHeader();
  pdf.setFont("helvetica", "normal");
  pdf.setFontSize(8.5);

  for (const r of rows) {
    if (y > pageH - 30) {
      pdf.addPage("a4", "landscape");
      y = 34;
      drawHeader();
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(8.5);
    }

    const cells = [
      String(r.Year),
      String(r.Group),
      money(r.Attachment),
      money(r.TotalLimit),
      String(r.Participants)
    ];

    let x = margin;
    for (let i = 0; i < cells.length; i++) {
      const clipped = pdf.splitTextToSize(cells[i], widths[i] - 4)[0] || "";
      pdf.text(clipped, x + 2, y);
      x += widths[i];
    }
    y += rowH;
  }

  pdf.save(filename);
  console.log(`[Export] PDF saved: ${filename} (tableRows=${rows.length})`);
}

/* ================================
   Public API
================================ */

export function setView(view) {
  const v = String(view || "").trim();
  if (!["carrier", "carrierGroup", "availability"].includes(v)) return;

  currentView = v;
  if (!chart || !_cache.options || !_cache.slices.length) return;

  rebuildChart();
}

export function getView() {
  return currentView;
}

export function setChartTheme(themeName) {
  const theme = themeName === "light" ? "light" : "dark";
  // Rebuild so dataset colors are recalculated for the active theme.
  rebuildChart();
  applyChartTheme(theme);
}

export function setCoverageTotalsVisible(visible) {
  _cache.showCoverageTotals = visible !== false;
  // No-op for canvas totals; external UI handles visibility.
  if (chart) chart.update("none");
}

export function setSIRMode(mode) {
  const m = String(mode || "").trim();
  _cache.filters.sirMode = ["off", "perOcc", "aggregate"].includes(m) ? m : "off";
  rebuildChart();
}

export function getYearBounds() {
  const years = _cache.allXLabels
    .map((lbl) => Number(lbl))
    .filter((y) => Number.isFinite(y))
    .sort((a, b) => a - b);
  const filterStartDate = parseDateToUTC(_cache.filters.startDate);
  const filterEndDate = parseDateToUTC(_cache.filters.endDate);
  const selectedStartYear = filterStartDate
    ? filterStartDate.getUTCFullYear()
    : Number.isFinite(_cache.filters.startYear)
    ? _cache.filters.startYear
    : null;
  const selectedEndYear = filterEndDate
    ? filterEndDate.getUTCFullYear()
    : Number.isFinite(_cache.filters.endYear)
    ? _cache.filters.endYear
    : null;

  return {
    minYear: years.length ? years[0] : null,
    maxYear: years.length ? years[years.length - 1] : null,
    startYear: selectedStartYear,
    endYear: selectedEndYear
  };
}

export function getFilterOptions() {
  const carrierSet = new Set();
  const carrierGroupSet = new Set();
  const insuranceProgramSet = new Set();
  const policyLimitTypeSet = new Set();
  for (const s of _cache.allSlices || []) {
    const c = String(s?.carrier || "").trim();
    const g = String(s?.carrierGroup || "").trim();
    const p = String(s?.insuranceProgram || "").trim();
    const limitType = String(s?.policyLimitType || "").trim();
    if (c && c !== "(unknown carrier)") carrierSet.add(c);
    if (g && g !== "(unknown group)") carrierGroupSet.add(g);
    if (p && p !== "(unknown program)") insuranceProgramSet.add(p);
    if (limitType) policyLimitTypeSet.add(limitType);
  }

  return {
    insurancePrograms: Array.from(insuranceProgramSet).sort((a, b) => a.localeCompare(b)),
    policyLimitTypes: Array.from(policyLimitTypeSet).sort((a, b) => a.localeCompare(b)),
    carriers: Array.from(carrierSet).sort((a, b) => a.localeCompare(b)),
    carrierGroups: Array.from(carrierGroupSet).sort((a, b) => a.localeCompare(b)),
    selectedInsuranceProgram: String(_cache.filters?.insuranceProgram || "").trim(),
    selectedPolicyLimitType: String(_cache.filters?.policyLimitType || "").trim(),
    selectedCarriers: normalizeStringList(_cache.filters?.carriers),
    selectedCarrierGroups: normalizeStringList(_cache.filters?.carrierGroups)
  };
}

export function getFilteredSlices() {
  return Array.isArray(_cache.slices) ? _cache.slices.slice() : [];
}

export function getYearLabelAnchors() {
  const xScale = chart?.scales?.x;
  if (!xScale || !Array.isArray(_cache.xLabels)) return [];
  const anchors = [];
  for (const lbl of _cache.xLabels) {
    const key = String(lbl ?? "");
    const px = Number(xScale.getPixelForValue(key));
    if (!Number.isFinite(px)) continue;
    anchors.push({ x: key, px });
  }
  return anchors;
}

export function setInsuranceProgramFilter(insuranceProgram) {
  _cache.filters.insuranceProgram = String(insuranceProgram || "").trim();
  applyFiltersToCache();
  rebuildChart();
}

export function resetInsuranceProgramFilter() {
  _cache.filters.insuranceProgram = "";
  applyFiltersToCache();
  rebuildChart();
}

export function setPolicyLimitTypeFilter(policyLimitType) {
  _cache.filters.policyLimitType = String(policyLimitType || "").trim();
  applyFiltersToCache();
  rebuildChart();
}

export function resetPolicyLimitTypeFilter() {
  _cache.filters.policyLimitType = "";
  applyFiltersToCache();
  rebuildChart();
}

export function setEntityFilters({ carriers, carrierGroups } = {}) {
  _cache.filters.carriers = normalizeStringList(carriers);
  _cache.filters.carrierGroups = normalizeStringList(carrierGroups);
  applyFiltersToCache();
  rebuildChart();
}

export function resetEntityFilters() {
  _cache.filters.carriers = [];
  _cache.filters.carrierGroups = [];
  applyFiltersToCache();
  rebuildChart();
}

export function setYearRange(startYear, endYear) {
  const norm = (v) => {
    const n = Number(v);
    return Number.isFinite(n) ? Math.trunc(n) : null;
  };

  let start = norm(startYear);
  let end = norm(endYear);
  if (start !== null && end !== null && start > end) [start, end] = [end, start];

  _cache.filters.startYear = start;
  _cache.filters.endYear = end;
  _cache.filters.startDate = null;
  _cache.filters.endDate = null;
  applyFiltersToCache();
  rebuildChart();
}

export function resetYearRange() {
  _cache.filters.startYear = null;
  _cache.filters.endYear = null;
  _cache.filters.startDate = null;
  _cache.filters.endDate = null;
  applyFiltersToCache();
  rebuildChart();
}

export function setDateRange(startDate, endDate) {
  let start = parseDateToUTC(startDate);
  let end = parseDateToUTC(endDate);
  if (start && end && start.getTime() > end.getTime()) [start, end] = [end, start];

  _cache.filters.startDate = start ? start.toISOString().slice(0, 10) : null;
  _cache.filters.endDate = end ? end.toISOString().slice(0, 10) : null;
  _cache.filters.startYear = start ? start.getUTCFullYear() : null;
  _cache.filters.endYear = end ? end.getUTCFullYear() : null;
  applyFiltersToCache();
  rebuildChart();
}

export function resetDateRange() {
  _cache.filters.startDate = null;
  _cache.filters.endDate = null;
  _cache.filters.startYear = null;
  _cache.filters.endYear = null;
  applyFiltersToCache();
  rebuildChart();
}

export function setZoomRange(min, max) {
  const norm = (v) => {
    if (v === null || v === undefined || String(v).trim() === "") return null;
    const n = num(v);
    return Number.isFinite(n) ? n : null;
  };

  let zoomMin = norm(min);
  let zoomMax = norm(max);
  if (zoomMin !== null && zoomMax !== null && zoomMin > zoomMax) {
    [zoomMin, zoomMax] = [zoomMax, zoomMin];
  }

  _cache.filters.zoomMin = zoomMin;
  _cache.filters.zoomMax = zoomMax;
  rebuildChart();
}

export function resetZoomRange() {
  _cache.filters.zoomMin = null;
  _cache.filters.zoomMax = null;
  rebuildChart();
}

/* ================================
   Main render function
================================ */

export async function renderCoverageChart({
  canvasId,
  csvUrl,

  policyDatesUrl = "/data/OriginalFiles/tblPolicyDates.csv",
  policyUrl = "/data/OriginalFiles/tblPolicy.csv",
  carrierUrl = "/data/OriginalFiles/tblCarrier.csv",
  carrierGroupUrl = "/data/OriginalFiles/tblCarrierGroup.csv",
  insuranceProgramUrl = "/data/OriginalFiles/tblInsuranceProgram.csv",
  policyLimitTypeUrl = "/data/OriginalFiles/tblPolicyLimitType.csv",

  useYearAxis = true,

  barThickness = "flex",
  categorySpacing = 1.0,
  outlineWidth = 1,
  tooltipMaxParticipants = 25,

  initialView = "carrier"
}) {
  if (!window.Chart) throw new Error("Chart.js must be loaded before CoverageChart.js");

  const canvas = document.getElementById(canvasId);
  if (!canvas) throw new Error("Canvas element not found");

  currentView = ["carrier", "carrierGroup", "availability"].includes(initialView)
    ? initialView
    : "carrier";
  const themeColors = getChartThemeColors();

  const [limitsRows, datesRows, policyRows, carrierRows, carrierGroupRows, insuranceProgramRows, policyLimitTypeRows] =
    await Promise.all([
    fetchCSV(csvUrl),
    fetchCSV(policyDatesUrl),
    fetchCSV(policyUrl),
    fetchCSV(carrierUrl),
    fetchCSV(carrierGroupUrl),
    fetchInsuranceProgramRows(insuranceProgramUrl),
    fetchCSV(policyLimitTypeUrl)
  ]);

  const built = buildSlices({
    limitsRows,
    datesRows,
    policyRows,
    carrierRows,
    carrierGroupRows,
    insuranceProgramRows,
    policyLimitTypeRows,
    useYearAxis
  });

  const quotaKeySet = buildQuotaKeySet(built.slices);

  _cache = {
    allSlices: built.slices,
    allXLabels: built.xLabels,
    slices: built.slices,
    xLabels: built.xLabels,
    options: { barThickness, categorySpacing },
    quotaKeySet,
    useYearAxis,
    xZoom: _cache.xZoom || 1,
    dom: {
      canvas,
      viewport: canvas.closest(".chartViewport"),
      surface: canvas.parentElement
    },
    _wheelBound: _cache._wheelBound || false,
    showCoverageTotals: typeof _cache.showCoverageTotals === "boolean" ? _cache.showCoverageTotals : true,
    filters: {
      startYear: null,
      endYear: null,
      startDate: null,
      endDate: null,
      zoomMin: null,
      zoomMax: null,
      sirMode: _cache.filters?.sirMode || "off",
      insuranceProgram: "",
      policyLimitType: "",
      carriers: [],
      carrierGroups: []
    }
  };

  applyFiltersToCache();

  const datasets = buildDatasetsForView({
    slices: _cache.slices,
    xLabels: _cache.xLabels,
    view: currentView,
    barThickness,
    categorySpacing,
    quotaKeySet: _cache.quotaKeySet
  });
  const sirDataset = buildSirDataset({
    slices: _cache.slices,
    xLabels: _cache.xLabels,
    sirMode: _cache.filters?.sirMode
  });

  if (chart) chart.destroy();

  chart = new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: _cache.xLabels,
      datasets: sirDataset ? [...datasets, sirDataset] : datasets
    },
    plugins: [outlineBarsPlugin, quotaShareGuidesPlugin, boxValueLabelsPlugin, yearAvailableTotalsPlugin],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      interaction: {
        mode: "nearest",
        intersect: true
      },
      layout: {
        // Keep bars/outline away from the extreme plot edges without using x.offset,
        // which can break floating + non-grouped bar geometry in this chart.
        padding: { top: 8, left: 8, right: 24 }
      },

      plugins: {
        outlineBars: {
          lineWidth: outlineWidth,
          color: themeColors.outline
        },
        legend: {
          display: true,
          position: "right",
          align: "start",
          title: {
            display: true,
            text: legendTitleForView(currentView),
            color: themeColors.legendText,
            font: { size: 12, weight: "600" },
            padding: 8
          },
          labels: {
            color: themeColors.legendText,
            boxWidth: 9,
            boxHeight: 9,
            padding: 10,
            font: { size: 11 }
          }
        },
        tooltip: {
          displayColors: false,
          mode: "nearest",
          intersect: true,
          filter: (_item, index) => index === 0,
          callbacks: {
            title: (items) => {
              const raw = items?.[0]?.raw || {};
              const yr = String(raw.x ?? "");
              const ds = items?.[0]?.dataset || {};
              if (ds?.datasetId === "sirOverlay") return `${yr} — ${ds.label || "SIR"}`;

              if (raw.isQuotaShare) return `${yr} — Quota share`;

              const g = String(raw.group ?? "").trim();
              return g ? `${yr} — ${g}` : yr;
            },
            label: (ctx) => {
              if (ctx?.dataset?.datasetId === "sirOverlay") {
                const val = Number(ctx?.raw?.y ?? ctx?.parsed?.y ?? 0);
                return [`${ctx.dataset.label}: ${money(val)}`];
              }

              const r = ctx.raw || {};
              const attach = r.attach ?? 0;
              const top = r.top ?? 0;
              const lim = Math.max(0, top - attach);

              const lines = [];
              lines.push(`Attach: ${money(attach)}`);
              lines.push(`Limit: ${money(lim)}`);
              lines.push(`Top: ${money(top)}`);
              const isPrimaryLayer = Number(attach) <= 0;
              if (isPrimaryLayer) {
                const sirValsPerOcc = (Array.isArray(r.participants) ? r.participants : [])
                  .map((p) => Number(p?.sirPerOcc || 0))
                  .filter((v) => Number.isFinite(v) && v > 0);
                const sirValsAgg = (Array.isArray(r.participants) ? r.participants : [])
                  .map((p) => Number(p?.sirAggregate || 0))
                  .filter((v) => Number.isFinite(v) && v > 0);
                if (sirValsPerOcc.length) {
                  const min = Math.min(...sirValsPerOcc);
                  const max = Math.max(...sirValsPerOcc);
                  lines.push(`SIR (Per Occ): ${min === max ? money(min) : `${money(min)} - ${money(max)}`}`);
                }
                if (sirValsAgg.length) {
                  const min = Math.min(...sirValsAgg);
                  const max = Math.max(...sirValsAgg);
                  lines.push(`SIR (Aggregate): ${min === max ? money(min) : `${money(min)} - ${money(max)}`}`);
                }
              }

              const parts = Array.isArray(r.participants) ? r.participants : [];
              const quotaParts = r.isQuotaShare && r.quotaGroupKey
                ? parts.filter((p) => String(p?.quotaGroupKey || "") === String(r.quotaGroupKey))
                : parts;
              const isUnavailableGroup = String(r.group || "").toLowerCase() === "unavailable";
              const shouldShowParts = !!r.isQuotaShare && quotaParts.length > 1;

              if (isUnavailableGroup && quotaParts.length) {
                const uniqueCarriers = [...new Set(quotaParts.map((p) => p.carrier || "(unknown carrier)"))];
                const carrierLine =
                  uniqueCarriers.length === 1
                    ? `Carrier: ${uniqueCarriers[0]}`
                    : `Carriers: ${uniqueCarriers.join(", ")}`;
                lines.push(carrierLine);
              }

              if (shouldShowParts && quotaParts.length) {
                lines.push(`Quota share participants (${quotaParts.length}):`);

                const show = quotaParts.slice(0, tooltipMaxParticipants);
                for (const p of show) {
                  const carrier = p.carrier || "(unknown carrier)";
                  const polno = p.policy_no ? ` (${p.policy_no})` : "";
                  lines.push(`• ${carrier}${polno}: ${money(p.sliceLimit)}`);
                }

                if (quotaParts.length > tooltipMaxParticipants) {
                  lines.push(`… +${quotaParts.length - tooltipMaxParticipants} more`);
                }
              }

              return lines;
            }
          }
        }
      },

      scales: {
        x: {
          type: "category",
          offset: true,
          grid: {
            display: true,
            color: themeColors.xGrid,
            lineWidth: 1
          },
          ticks: {
            color: themeColors.axisTicks,
            autoSkip: false,
            maxRotation: 0,
            minRotation: 0
          },
          title: {
            display: true,
            text: "Policy Years",
            color: themeColors.yTitle
          }
        },
        y: {
          beginAtZero: true,
          grid: {
            display: true,
            color: themeColors.yGrid,
            lineWidth: 1
          },
          ticks: {
            color: themeColors.axisTicks,
            callback: (v) => compactMoney(v)
          },
          title: {
            display: true,
            text: "Coverage Limits",
            color: themeColors.yTitle
          }
        }
      }
    }
  });

  bindViewportInteractions();
  // Defer width sync so the viewport has final layout dimensions.
  requestAnimationFrame(() => {
    if (!chart) return;
    applyChartTheme();
    syncChartViewportWidth();
    // Rebuild once after first layout so auto bar width can use the real x-scale width
    // (prevents gaps/overlap when legend placement changes plot width).
    if (!Number.isFinite(Number(barThickness))) rebuildChart();
  });

  console.log("Rendered X Labels:", _cache.xLabels.length);
  console.log("Rendered Slices:", _cache.slices.length);
  console.log("Quota keys:", _cache.quotaKeySet.size);
  console.log("Current View:", currentView);
}
