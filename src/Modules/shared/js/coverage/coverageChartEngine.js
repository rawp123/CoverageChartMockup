// Modules/CoverageChart/coverageChartEngine.js
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
let _responsiveResizeBound = false;
let _responsiveResizeTimer = null;

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
  hiddenLegendCarriers: new Set(),
  hiddenLegendCarrierGroups: new Set(),
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
    annualized: false,
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
const startOfNextYearUTC = (y) => Date.UTC(y + 1, 0, 1, 0, 0, 0, 0);

function msToYearAxisValue(ms) {
  const t = Number(ms);
  if (!Number.isFinite(t)) return NaN;
  const d = new Date(t);
  const y = d.getUTCFullYear();
  const yearStart = startOfYearUTC(y);
  const nextYearStart = startOfNextYearUTC(y);
  const span = nextYearStart - yearStart;
  if (!(span > 0)) return y;
  const frac = clamp((t - yearStart) / span, 0, 1);
  // Center each calendar year on its integer tick:
  // Jan 1 -> (year - 0.5), Dec 31 -> (~year + 0.5).
  return y - 0.5 + frac;
}

function getYearAxisBounds(xLabels) {
  const years = (xLabels || [])
    .map((lbl) => Number(lbl))
    .filter((y) => Number.isFinite(y))
    .sort((a, b) => a - b);
  if (!years.length) return { min: 0, max: 1 };
  return {
    min: years[0] - 0.5,
    max: years[years.length - 1] + 0.5
  };
}

function yearLabelFromAxisValue(v) {
  const n = Number(v);
  if (!Number.isFinite(n)) return "";
  return String(Math.floor(n + 0.5));
}

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

function formatFullDateUTC(value) {
  const ms = Number(value);
  if (!Number.isFinite(ms)) return "";
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return "";
  return d.toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
    timeZone: "UTC"
  });
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
    `Annualized: ${f.annualized ? "On" : "Off"}`,
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

function getResponsiveMode() {
  const w = Number(typeof window !== "undefined" ? window.innerWidth : 1600);
  return Number.isFinite(w) && w <= 1200 ? "small" : "large";
}

function applyResponsiveChartSettings(updateMode = "none") {
  if (!chart) return;
  const mode = getResponsiveMode();
  const isSmall = mode === "small";
  const legend = chart.options?.plugins?.legend;
  const x = chart.options?.scales?.x;
  const y = chart.options?.scales?.y;

  if (legend) {
    legend.position = isSmall ? "bottom" : "right";
    legend.align = "start";
    legend.maxWidth = isSmall ? undefined : 300;
    legend.maxHeight = isSmall ? 140 : 560;
    legend.labels = {
      ...(legend.labels || {}),
      boxWidth: isSmall ? 8 : 9,
      boxHeight: isSmall ? 8 : 9,
      padding: isSmall ? 8 : 10,
      font: { size: isSmall ? 10 : 11 }
    };
    if (legend.title) {
      legend.title.font = { size: isSmall ? 11 : 12, weight: "600" };
      legend.title.padding = isSmall ? 6 : 8;
    }
  }

  if (x?.ticks) {
    x.ticks.font = { size: isSmall ? 10 : 12 };
    x.ticks.autoSkip = isSmall;
    x.ticks.maxTicksLimit = isSmall ? 10 : undefined;
    x.ticks.maxRotation = isSmall ? 28 : 0;
    x.ticks.minRotation = isSmall ? 18 : 0;
  }
  if (y?.ticks) {
    y.ticks.font = { size: isSmall ? 10 : 12 };
    y.ticks.autoSkip = true;
    y.ticks.maxTicksLimit = isSmall ? 6 : 9;
  }

  chart.resize();
  chart.update(updateMode);
}

function bindResponsiveResizeHandler() {
  if (_responsiveResizeBound || typeof window === "undefined") return;
  window.addEventListener(
    "resize",
    () => {
      if (_responsiveResizeTimer) window.clearTimeout(_responsiveResizeTimer);
      _responsiveResizeTimer = window.setTimeout(() => {
        applyResponsiveChartSettings("none");
      }, 100);
    },
    { passive: true }
  );
  _responsiveResizeBound = true;
}

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

const _colorSpecByKey = new Map();
const _distinctPaletteDark = [
  "#4E79A7", "#F28E2B", "#E15759", "#76B7B2", "#59A14F", "#EDC948",
  "#B07AA1", "#FF9DA7", "#9C755F", "#BAB0AC", "#2F4B7C", "#A05195",
  "#D45087", "#FF7C43", "#665191", "#003F5C", "#EF5675", "#FFA600",
  "#1F77B4", "#17BECF", "#9467BD", "#8C564B", "#BCBD22", "#E377C2",
  "#2CA02C", "#D62728", "#8DD3C7", "#FB8072", "#80B1D3", "#FDB462"
];
const _distinctPaletteLight = [
  "#2F5D8A", "#A95E16", "#B04747", "#2E7F7A", "#3D7B35", "#8D7420",
  "#7A5A86", "#B86473", "#6F533F", "#5F6368", "#233A60", "#6D3F73",
  "#A13E6A", "#B85A2C", "#4F3E7A", "#1D3852", "#A9466A", "#B06F00",
  "#215F97", "#187A94", "#6E55A2", "#6D4D3F", "#7A7B1E", "#9A4D94",
  "#2C7A2C", "#A83E3E", "#3A8A84", "#B05A56", "#4E6EA2", "#A87739"
];

const normalizeHue = (h) => {
  const n = Number(h);
  if (!Number.isFinite(n)) return 0;
  const m = n % 360;
  return m < 0 ? m + 360 : m;
};

const hueDistance = (a, b) => {
  const d = Math.abs(normalizeHue(a) - normalizeHue(b));
  return Math.min(d, 360 - d);
};

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
  if (!_colorSpecByKey.has(s)) {
    const hash = hashString(s);
    const paletteSize = _distinctPaletteDark.length;
    const usedSlots = new Set(Array.from(_colorSpecByKey.values()).map((v) => Number(v?.slot)));
    const seedSlot = hash % paletteSize;
    // Probe across the palette with a coprime step so active keys avoid
    // collisions as long as there are remaining free slots.
    const step = 7;
    let chosenSlot = seedSlot;
    for (let i = 0; i < paletteSize; i++) {
      const candidate = (seedSlot + i * step) % paletteSize;
      if (!usedSlots.has(candidate)) {
        chosenSlot = candidate;
        break;
      }
    }
    _colorSpecByKey.set(s, { slot: chosenSlot });
  }

  const spec = _colorSpecByKey.get(s) || { slot: 0 };
  const isLight = getThemeName() === "light";
  const slot = clamp(Number(spec.slot || 0), 0, _distinctPaletteDark.length - 1);
  return isLight ? _distinctPaletteLight[slot] : _distinctPaletteDark[slot];
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
        const props = bar.getProps(["x", "y", "base", "width"], false);
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

        const hideUnavailable = isUnavailableLegendHidden(chart);
        const parts = (Array.isArray(raw.participants) ? raw.participants : [])
          .filter((p) => !hideUnavailable || !String(p?.availability || "").toLowerCase().includes("unavail"));
        if (parts.length < 2) return;

        const props = bar.getProps(["x", "width"], false);
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

function isUnavailableLegendHidden(chartInstance) {
  if (!chartInstance || !Array.isArray(chartInstance?.data?.datasets)) return false;
  const dsIndex = chartInstance.data.datasets.findIndex((d) => String(d?.label || "") === "Unavailable");
  if (dsIndex < 0) return false;
  const meta = chartInstance.getDatasetMeta(dsIndex);
  return !!meta?.hidden;
}

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
    // Temporarily disabled: per-layer value labels add too much visual noise.
    return;
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

        const props = bar.getProps(["x", "width"], false);
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

function applyXRangeBarGeometry(chartInstance) {
  if (!chartInstance || !_cache.useYearAxis) return;
  const xScale = chartInstance?.scales?.x;
  if (!xScale) return;

  chartInstance.data.datasets.forEach((ds, di) => {
    if (!ds || ds.datasetId === "sirOverlay" || ds.type === "line") return;
    const meta = chartInstance.getDatasetMeta(di);
    if (!meta || meta.hidden) return;

    meta.data.forEach((bar, idx) => {
      const raw = ds.data?.[idx];
      const xStart = Number(raw?.xStart);
      const xEnd = Number(raw?.xEnd);
      if (!Number.isFinite(xStart) || !Number.isFinite(xEnd) || !(xEnd > xStart)) return;

      const pxStart = xScale.getPixelForValue(xStart);
      const pxEnd = xScale.getPixelForValue(xEnd);
      if (!Number.isFinite(pxStart) || !Number.isFinite(pxEnd)) return;

      const left = Math.min(pxStart, pxEnd);
      const right = Math.max(pxStart, pxEnd);
      const width = Math.max(1, right - left);
      const center = (left + right) / 2;

      // Mutate bar element geometry so draw + hitbox reflect the true day span.
      bar.x = center;
      bar.width = width;
    });
  });
}

const xRangeBarsPlugin = {
  id: "xRangeBars",
  afterDatasetsUpdate(chartInstance) {
    applyXRangeBarGeometry(chartInstance);
  },
  beforeDatasetsDraw(chartInstance) {
    applyXRangeBarGeometry(chartInstance);
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
    const rawStartMs = startDate
      ? startDate.getTime()
      : startOfYearUTC(policyStartYear);
    const rawEndMs = endDate
      ? endDate.getTime() + (24 * 60 * 60 * 1000 - 1)
      : endOfYearUTC(policyEndYear);
    const policyStartMs = Math.min(rawStartMs, rawEndMs);
    const policyEndMs = Math.max(rawStartMs, rawEndMs);

    const attachRaw = firstPresentNum(
      getBy(r, "Attachment Point", "AttachmentPoint", "attach", "Attatchment Point")
    );

    const sliceLimitRaw = firstPresentNum(
      getBy(r, "LayerPerOccLimit", "Layer Per Occ Limit", "layerperocclim"),
      getBy(r, "PerOccLimit", "Per Occ Limit", "perocclim")
    );

    if (sliceLimitRaw <= 0) continue;

    const attach = Math.round(attachRaw);
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

    const baseSlice = {
      policyStartYear,
      policyEndYear,
      policyStartMs,
      policyEndMs,
      attach,
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
    };

    if (useYearAxis) {
      // Emit one slice per covered calendar year.
      // Keep full layer limit in each overlapped year so attachment stacks remain
      // continuous (no artificial vertical gaps from limit-proration).
      for (let y = policyStartYear; y <= policyEndYear; y++) {
        const yStartMs = startOfYearUTC(y);
        const yEndMs = endOfYearUTC(y);
        const yNextStartMs = startOfNextYearUTC(y);
        const overlapStart = Math.max(policyStartMs, yStartMs);
        const overlapEnd = Math.min(policyEndMs, yEndMs);
        if (overlapEnd < overlapStart) continue;
        const overlapMs = overlapEnd - overlapStart + 1;
        const yearSpanMs = yEndMs - yStartMs + 1;
        const overlapEndExclusive = Math.min(overlapEnd + 1, yNextStartMs);
        const xStartValue = msToYearAxisValue(overlapStart);
        const xEndValue = msToYearAxisValue(overlapEndExclusive);
        const xMidValue = (xStartValue + xEndValue) / 2;
        if (!Number.isFinite(xStartValue) || !Number.isFinite(xEndValue) || !(xEndValue > xStartValue)) continue;

        slices.push({
          ...baseSlice,
          sliceLimit: sliceLimitRaw,
          yearOverlapStartMs: overlapStart,
          yearOverlapEndMs: overlapEnd,
          yearOverlapRatio: yearSpanMs > 0 ? overlapMs / yearSpanMs : 0,
          xStartValue,
          xEndValue,
          xMidValue,
          x: String(y),
          year: y
        });
      }
    } else {
      slices.push({
        ...baseSlice,
        sliceLimit: sliceLimitRaw,
        x: `${dates.start} to ${dates.end}`,
        year: policyStartYear
      });
    }
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
  // Quota-share is scoped within the same program/layer/type (+ named insured)
  // AND the same covered date span. This prevents sequential half-year policies
  // in one calendar year from being mislabeled as quota share.
  const program = String(s?.insuranceProgramId || s?.insuranceProgram || "").trim();
  const year = Number.isFinite(s?.year) ? String(s.year) : String(s?.x || "").trim();
  const attach = String(s?.attach ?? "").trim();
  const limitType = String(s?.policyLimitTypeId || "").trim();
  const namedInsured = String(s?.namedInsuredId || "").trim();
  const spanStart = Number.isFinite(Number(s?.yearOverlapStartMs))
    ? Number(s.yearOverlapStartMs)
    : Number(s?.policyStartMs || 0);
  const spanEnd = Number.isFinite(Number(s?.yearOverlapEndMs))
    ? Number(s.yearOverlapEndMs)
    : Number(s?.policyEndMs || 0);
  return `${program}||${year}||${attach}||${limitType}||${namedInsured}||${spanStart}||${spanEnd}`;
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

function hasExplicitQuotaShareEvidence({ limitsRows, policyRows }) {
  const rowSets = [limitsRows || [], policyRows || []];
  const keyHint = /(quota|share|coinsur|co-insur|participation|participat|percent|pct)/i;
  const valueHint = /(quota|share|co-?insur)/i;
  const explicitQuotaPhrase = /quota\s*share/i;

  const looksLikeQuotaPercent = (v) => {
    if (v === null || v === undefined) return false;
    const s = String(v).trim();
    if (!s) return false;
    const n = Number(s.replace(/[^0-9.-]/g, ""));
    if (!Number.isFinite(n)) return false;
    return (n > 0 && n <= 1) || (n > 0 && n <= 100);
  };

  for (const rows of rowSets) {
    if (!rows.length) continue;
    for (const r of rows) {
      if (!r || typeof r !== "object") continue;
      for (const [k, v] of Object.entries(r)) {
        const sv = String(v ?? "").trim();
        // Accept explicit quota-share phrases in values regardless of column name
        // (common when stored in fields like PolicyNotes).
        if (sv && explicitQuotaPhrase.test(sv)) return true;

        if (!keyHint.test(String(k || ""))) continue;
        if (!sv) continue;
        if (looksLikeQuotaPercent(sv) || valueHint.test(sv)) return true;
      }
    }
  }

  return false;
}

function applyFiltersToCache() {
  const { allSlices, allXLabels, useYearAxis, filters } = _cache;
  const startYear = Number.isFinite(filters.startYear) ? filters.startYear : null;
  const endYear = Number.isFinite(filters.endYear) ? filters.endYear : null;
  const startDate = parseDateToUTC(filters.startDate);
  const endDate = parseDateToUTC(filters.endDate);
  const selectedProgram = String(filters.insuranceProgram || "").trim();
  let selectedPolicyLimitType = String(filters.policyLimitType || "").trim();
  if (!selectedPolicyLimitType) {
    const availableTypes = Array.from(
      new Set((allSlices || []).map((s) => String(s?.policyLimitType || "").trim()).filter(Boolean))
    ).sort((a, b) => a.localeCompare(b));
    const bodilyInjury =
      availableTypes.find((t) => String(t || "").trim().toLowerCase() === "bodily injury") || "";
    selectedPolicyLimitType = bodilyInjury || String(availableTypes[0] || "").trim();
    filters.policyLimitType = selectedPolicyLimitType;
  }

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
  const x = chart.options?.scales?.x;
  if (x && _cache.useYearAxis && x.type === "linear") {
    const bounds = getYearAxisBounds(_cache.xLabels);
    x.min = bounds.min;
    x.max = bounds.max;
  }
  if (chart.options?.plugins?.legend?.title) {
    chart.options.plugins.legend.title.text = legendTitleForView(currentView);
  }

  applyResponsiveChartSettings("none");
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
  chart.options.scales.x.title.color = c.yTitle;
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
  applyResponsiveChartSettings("none");
}

/* ================================
   View aggregation -> datasets
================================ */

function buildDatasetsForView({ slices, xLabels, view, barThickness, categorySpacing, quotaKeySet }) {
  const qaKey = (s) => quotaGroupKey(s);
  const isQuotaSlice = (s) => quotaKeySet && quotaKeySet.has(qaKey(s));
  const isUnavailable = (s) => String(s?.availability || "").toLowerCase().includes("unavail");
  const selection = getSelectionSets();
  const usingYearAxis = !!_cache.useYearAxis;
  const annualized = !!_cache.filters?.annualized;
  const hiddenCarriers = _cache.hiddenLegendCarriers || new Set();
  const hiddenCarrierGroups = _cache.hiddenLegendCarrierGroups || new Set();
  const isHiddenBySyntheticLegend = (s) => {
    if (view === "carrier") {
      const c = String(s?.carrier || "").trim();
      return !!c && hiddenCarriers.has(c);
    }
    if (view === "carrierGroup") {
      const g = String(s?.carrierGroup || "").trim();
      return !!g && hiddenCarrierGroups.has(g);
    }
    return false;
  };

  const keyOf = (s) => {
    if (view === "carrier" || view === "carrierGroup") {
      // Keep quota-share rollup for both carrier and carrier-group views so
      // concurrent quota participants do not overdraw each other.
      if (isQuotaSlice(s)) return "Quota share";
      // Non-quota unavailable slices are still collapsed into one dataset.
      if (isUnavailable(s)) return "Unavailable";
      return view === "carrier" ? (s.carrier || "(unknown carrier)") : (s.carrierGroup || "(unknown group)");
    }
    // For non-carrier views, keep unavailable layers grouped.
    if (isUnavailable(s)) return "Unavailable";
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

  if (usingYearAxis) {
    const bucketMap = new Map();
    for (const s of slices) {
      if (isHiddenBySyntheticLegend(s)) continue;
      const group = keyOf(s);
      groups.add(group);

      const sliceQuotaKey = qaKey(s);
      const xStartValue = Number(s?.xStartValue);
      const xEndValue = Number(s?.xEndValue);
      if (!Number.isFinite(xStartValue) || !Number.isFinite(xEndValue) || !(xEndValue > xStartValue)) continue;
      const xMidValue = (xStartValue + xEndValue) / 2;
      const yearLabel = yearLabelFromAxisValue(xMidValue);
      const quotaGroupKey =
        (view === "carrier" || view === "carrierGroup") && isQuotaSlice(s)
          ? sliceQuotaKey
          : "";
      // Year-mode day-accurate bucket: (group, attachment, quotaGroupKey).
      const k = `${group}||${s.attach}||${quotaGroupKey}`;
      if (!bucketMap.has(k)) {
        bucketMap.set(k, {
          group,
          attach: Number(s?.attach || 0),
          quotaGroupKey,
          slices: []
        });
      }

      bucketMap.get(k).slices.push({
        source: s,
        sliceQuotaKey,
        limit: Number(s?.sliceLimit || 0),
        xStartValue,
        xEndValue
      });
    }

    for (const bucket of bucketMap.values()) {
      const bounds = Array.from(
        new Set(
          bucket.slices.flatMap((row) => [Number(row?.xStartValue), Number(row?.xEndValue)])
        )
      )
        .filter((v) => Number.isFinite(v))
        .sort((a, b) => a - b);

      const rawSegments = [];
      for (let i = 0; i < bounds.length - 1; i++) {
        const segStart = bounds[i];
        const segEnd = bounds[i + 1];
        if (!Number.isFinite(segStart) || !Number.isFinite(segEnd) || !(segEnd > segStart)) continue;

        const activeRows = bucket.slices.filter((row) => {
          const xs = Number(row?.xStartValue);
          const xe = Number(row?.xEndValue);
          return Number.isFinite(xs) && Number.isFinite(xe) && xs < segEnd && xe > segStart;
        });
        if (!activeRows.length) continue;

        const participants = [];
        let segLimit = 0;
        let hasSelectionMatch = false;

        for (const row of activeRows) {
          const limit = Number(row?.limit || 0);
          if (!Number.isFinite(limit) || limit <= 0) continue;
          segLimit += limit;
          const src = row?.source || {};
          if (sliceMatchesSelection(src, selection)) hasSelectionMatch = true;
          participants.push({
            pid: src?.PolicyID,
            carrier: src?.carrier,
            carrierGroup: src?.carrierGroup,
            insuranceProgram: src?.insuranceProgram,
            availability: src?.availability,
            policy_no: src?.policy_no,
            policyLimitType: String(src?.policyLimitType || src?.policyLimitTypeId || ""),
            policyStartMs: Number(src?.policyStartMs || 0),
            policyEndMs: Number(src?.policyEndMs || 0),
            segmentStartMs: Number(src?.yearOverlapStartMs || src?.policyStartMs || 0),
            segmentEndMs: Number(src?.yearOverlapEndMs || src?.policyEndMs || 0),
            sliceLimit: limit,
            xStartValue: segStart,
            xEndValue: segEnd,
            sirPerOcc: Number(src?.sirPerOcc || 0),
            sirAggregate: Number(src?.sirAggregate || 0),
            quotaGroupKey: bucket.quotaGroupKey ? row?.sliceQuotaKey : ""
          });
        }
        if (!(segLimit > 0)) continue;

        const signature = participants
          .map((p) => `${String(p?.pid || "")}||${Number(p?.sliceLimit || 0)}||${String(p?.quotaGroupKey || "")}`)
          .sort()
          .join("##");
        rawSegments.push({
          segStart,
          segEnd,
          segLimit,
          participants,
          hasSelectionMatch,
          signature
        });
      }

      const finalSegments = [];
      if (annualized) {
        for (const seg of rawSegments) finalSegments.push({ ...seg });
      } else {
        // Merge adjacent segments when the active participant set is unchanged.
        const eps = 1e-9;
        for (const seg of rawSegments) {
          const prev = finalSegments[finalSegments.length - 1];
          if (
            prev &&
            Math.abs(Number(prev.segEnd) - Number(seg.segStart)) < eps &&
            prev.signature === seg.signature &&
            Math.abs(Number(prev.segLimit) - Number(seg.segLimit)) < eps
          ) {
            prev.segEnd = seg.segEnd;
            prev.hasSelectionMatch = prev.hasSelectionMatch || seg.hasSelectionMatch;
          } else {
            finalSegments.push({ ...seg });
          }
        }
      }

      for (const seg of finalSegments) {
        const segMid = (Number(seg.segStart) + Number(seg.segEnd)) / 2;
        const yearLabel = yearLabelFromAxisValue(segMid);
        const layerKey = `${bucket.group}||${bucket.attach}||${Number(seg.segStart).toFixed(9)}||${Number(seg.segEnd).toFixed(9)}||${bucket.quotaGroupKey}`;
        layerMap.set(layerKey, {
          group: bucket.group,
          x: yearLabel,
          yearLabel,
          xStartValue: Number(seg.segStart),
          xEndValue: Number(seg.segEnd),
          xMidValue: segMid,
          attach: bucket.attach,
          quotaGroupKey: bucket.quotaGroupKey,
          sumLimit: Number(seg.segLimit),
          participants: seg.participants,
          hasSelectionMatch: !!seg.hasSelectionMatch
        });
      }
    }

    /*
      Legacy event-based aggregation retained for reference:
      The boundary-segmentation above provides exact day-level rendering by
      attachment and avoids year-bucket overcounting.
    */
    /*
    for (const bucket of bucketMap.values()) {
      const events = [];
      for (const row of bucket.slices) {
        const lim = Number(row.limit || 0);
        if (!Number.isFinite(lim) || lim <= 0) continue;
        const startMs = Number(row.overlapStartMs);
        const endMs = Number(row.overlapEndMs);
        if (!Number.isFinite(startMs) || !Number.isFinite(endMs)) continue;
        const lo = Math.min(startMs, endMs);
        const hi = Math.max(startMs, endMs);
        const endExclusive = hi + 1;
        const token = `${String(row.source?.PolicyID || "")}||${lo}||${hi}||${lim}`;
        events.push({ t: lo, delta: +lim, token, row });
        events.push({ t: endExclusive, delta: -lim, token, row });
      }
      if (!events.length) continue;

      // At equal timestamps process removals before adds so adjacent periods
      // do not count as concurrent overlap.
      events.sort((a, b) => (a.t - b.t) || (a.delta - b.delta));

      let running = 0;
      let maxRunning = 0;
      const active = new Map();
      let peakParticipants = [];
      let hasSelectionMatch = false;

      for (const ev of events) {
        if (ev.delta < 0) {
          const existing = active.get(ev.token);
          if (existing) {
            running = Math.max(0, running - Number(existing?.sliceLimit || 0));
            active.delete(ev.token);
          }
        } else {
          const p = {
            pid: ev.row.source?.PolicyID,
            carrier: ev.row.source?.carrier,
            carrierGroup: ev.row.source?.carrierGroup,
            availability: ev.row.source?.availability,
            policy_no: ev.row.source?.policy_no,
            policyStartMs: Number(ev.row.source?.policyStartMs || 0),
            policyEndMs: Number(ev.row.source?.policyEndMs || 0),
            sliceLimit: Number(ev.row.limit || 0),
            xStartValue: Number(bucket.xStartValue),
            xEndValue: Number(bucket.xEndValue),
            sirPerOcc: Number(ev.row.source?.sirPerOcc || 0),
            sirAggregate: Number(ev.row.source?.sirAggregate || 0),
            quotaGroupKey: ev.row.sliceQuotaKey
          };
          active.set(ev.token, p);
          running += p.sliceLimit;
        }

        if (running > maxRunning) {
          maxRunning = running;
          peakParticipants = Array.from(active.values());
          hasSelectionMatch = peakParticipants.some((p) =>
            sliceMatchesSelection(
              { carrier: p.carrier, carrierGroup: p.carrierGroup },
              selection
            )
          );
        }
      }

      if (!(maxRunning > 0)) continue;
      layerMap.set(`${bucket.group}||${bucket.yearLabel}||${bucket.attach}||${bucket.quotaGroupKey}`, {
        group: bucket.group,
        x: bucket.x,
        yearLabel: bucket.yearLabel,
        xStartValue: bucket.xStartValue,
        xEndValue: bucket.xEndValue,
        xMidValue: bucket.xMidValue,
        attach: bucket.attach,
        quotaGroupKey: bucket.quotaGroupKey,
        sumLimit: maxRunning,
        participants: peakParticipants,
        hasSelectionMatch
      });
    }
    */
  } else {
    for (const s of slices) {
      if (isHiddenBySyntheticLegend(s)) continue;
      const group = keyOf(s);
      groups.add(group);

      const sliceQuotaKey = qaKey(s);
      const quotaGroupKey =
        (view === "carrier" || view === "carrierGroup") && isQuotaSlice(s)
          ? sliceQuotaKey
          : "";
      const k = `${group}||${s.x}||${s.attach}||${quotaGroupKey}`;
      if (!layerMap.has(k)) {
        layerMap.set(k, {
          group,
          x: s.x,
          yearLabel: String(s?.x ?? ""),
          xStartValue: null,
          xEndValue: null,
          xMidValue: null,
          attach: s.attach,
          quotaGroupKey,
          sumLimit: 0,
          participants: [],
          hasSelectionMatch: false
        });
      }

      const e = layerMap.get(k);
      e.sumLimit += Number(s?.sliceLimit || 0);
      if (sliceMatchesSelection(s, selection)) e.hasSelectionMatch = true;
      e.participants.push({
        pid: s.PolicyID,
        carrier: s.carrier,
        carrierGroup: s.carrierGroup,
        insuranceProgram: s.insuranceProgram,
        availability: s.availability,
        policy_no: s.policy_no,
        policyLimitType: String(s?.policyLimitType || s?.policyLimitTypeId || ""),
        policyStartMs: Number(s.policyStartMs || 0),
        policyEndMs: Number(s.policyEndMs || 0),
        sliceLimit: Number(s?.sliceLimit || 0),
        xStartValue: null,
        xEndValue: null,
        sirPerOcc: Number(s.sirPerOcc || 0),
        sirAggregate: Number(s.sirAggregate || 0),
        quotaGroupKey: sliceQuotaKey
      });
    }
  }

  if (usingYearAxis) {
    const selectedCarriers = normalizeStringList(_cache.filters?.carriers);
    const debugGraniteOnly = selectedCarriers.length === 1 &&
      selectedCarriers[0].toLowerCase() === "granite state insurance company";
    if (debugGraniteOnly) {
      const rows = Array.from(layerMap.values())
        .filter((e) => String(e?.group || "").toLowerCase() === "granite state insurance company")
        .filter((e) => String(e?.yearLabel || "") === "1985")
        .sort((a, b) => Number(a.attach || 0) - Number(b.attach || 0));
      if (rows.length) {
        console.group("[CoverageChart Debug] Granite State 1985 bands");
        for (const e of rows) {
          const start = Number(e.attach || 0);
          const totalLayerLimit = Number(e.sumLimit || 0);
          const end = start + totalLayerLimit;
          console.log({
            AttachmentPoint: start,
            TotalLayerLimit: totalLayerLimit,
            CalculatedStart: start,
            CalculatedEnd: end
          });
        }
        console.groupEnd();
      }
    }
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
    const hideUnavailable = isUnavailableLegendHidden(context?.chart);
    const parts = (Array.isArray(pointRaw?.participants) ? pointRaw.participants : [])
      .filter((p) => !hideUnavailable || !String(p?.availability || "").toLowerCase().includes("unavail"));
    const weightedParts = parts
      .map((p) => ({
        limit: Number(p?.sliceLimit || 0),
        color: String(p?.availability || "").toLowerCase().includes("unavail")
          ? "#94a3b8"
          : colorFromString(p?.carrier || "Quota share")
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
  const quotaGradientForCarrierGroupView = (context, pointRaw) => {
    const hideUnavailable = isUnavailableLegendHidden(context?.chart);
    const parts = (Array.isArray(pointRaw?.participants) ? pointRaw.participants : [])
      .filter((p) => !hideUnavailable || !String(p?.availability || "").toLowerCase().includes("unavail"));
    const weightedParts = parts
      .map((p) => ({
        limit: Number(p?.sliceLimit || 0),
        color: String(p?.availability || "").toLowerCase().includes("unavail")
          ? "#94a3b8"
          : colorFromString(p?.carrierGroup || "Quota share")
      }))
      .filter((p) => Number.isFinite(p.limit) && p.limit > 0);
    if (!weightedParts.length) return colorFromString(String(pointRaw?.group || "Quota share"));

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
        x: usingYearAxis && Number.isFinite(Number(e.xMidValue)) ? Number(e.xMidValue) : e.x,
        xStart: usingYearAxis && Number.isFinite(Number(e.xStartValue)) ? Number(e.xStartValue) : null,
        xEnd: usingYearAxis && Number.isFinite(Number(e.xEndValue)) ? Number(e.xEndValue) : null,
        yearLabel: e.yearLabel || String(e.x ?? ""),
        y: [e.attach, top],
        attach: e.attach,
        top,
        sumLimit: e.sumLimit,
        participants: e.participants,
        group,
        annualized,
        isQuotaShare,
        quotaGroupKey: e.quotaGroupKey || "",
        isHighlighted: !selection.active || !!e.hasSelectionMatch
      });
    }

    points.sort((a, b) => {
      if (usingYearAxis) {
        const ax = Number(a?.xStart);
        const bx = Number(b?.xStart);
        if (Number.isFinite(ax) && Number.isFinite(bx) && ax !== bx) return ax - bx;
        const aend = Number(a?.xEnd);
        const bend = Number(b?.xEnd);
        if (Number.isFinite(aend) && Number.isFinite(bend) && aend !== bend) return aend - bend;
        return a.attach - b.attach;
      }
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

    const mutedFill = "rgba(255,255,255,0.92)";
    const mutedBorder = "rgba(0,0,0,0.9)";
    const isQuotaDataset = points.some((p) => !!p?.isQuotaShare);

    return {
      label: group,
      data: points,
      parsing: { xAxisKey: "x", yAxisKey: "y" },
      // Keep quota-share layers visually on top so split-color + dashed guides
      // are not hidden by same-attachment carrier bars.
      order: isQuotaDataset ? 100 : 10,

      grouped: false,
      barThickness: barThicknessPx,
      maxBarThickness: barThicknessPx,

      categoryPercentage: 1.0,
      barPercentage: 1.0,
      inflateAmount: 0.6,

      backgroundColor: (context) => {
        const raw = context?.raw || null;
        const pointInFocus = !selection.active || !!raw?.isHighlighted;
        if (!pointInFocus) return mutedFill;
        if (view === "carrier" && isQuotaDataset && raw?.isQuotaShare) {
          return quotaGradientForCarrierView(context, raw);
        }
        if (view === "carrierGroup" && isQuotaDataset && raw?.isQuotaShare) {
          return quotaGradientForCarrierGroupView(context, raw);
        }
        return bg;
      },
      borderColor: (context) => {
        const raw = context?.raw || null;
        const pointInFocus = !selection.active || !!raw?.isHighlighted;
        return pointInFocus ? "rgba(11, 17, 27, 0.62)" : mutedBorder;
      },
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
  const usingYearAxis = !!_cache.useYearAxis;
  const byX = new Map();
  for (const s of slices || []) {
    const x = usingYearAxis
      ? Number.isFinite(Number(s?.year))
        ? String(Number(s.year))
        : String(s?.x ?? "")
      : String(s?.x ?? "");
    if (!x) continue;
    const v = getSirValueFromSlice(s, mode);
    if (!Number.isFinite(v) || v <= 0) continue;
    byX.set(x, Math.max(byX.get(x) || 0, v));
  }

  const color = mode === "aggregate" ? "#f97316" : "#facc15";
  const data = (xLabels || [])
    .map((x) => {
      const yearLabel = String(x ?? "");
      const v = byX.get(yearLabel);
      if (usingYearAxis) {
        const yr = Number(yearLabel);
        if (!Number.isFinite(yr)) return null;
        return { x: yr, y: Number.isFinite(v) ? v : null, yearLabel };
      }
      return { x: yearLabel, y: Number.isFinite(v) ? v : null, yearLabel };
    })
    .filter(Boolean);

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
      ? `${String(f.startDate || "All")} to ${String(f.endDate || "All")}`
      : Number.isFinite(f.startYear) || Number.isFinite(f.endYear)
      ? `${Number.isFinite(f.startYear) ? f.startYear : "All"} to ${Number.isFinite(f.endYear) ? f.endYear : "All"}`
      : "All";
  const carriers = normalizeStringList(f.carriers);
  const carrierGroups = normalizeStringList(f.carrierGroups);
  const summarizeList = (list, allLabel) => {
    if (!list.length) return allLabel;
    if (list.length <= 6) return list.join(", ");
    return `${list.slice(0, 6).join(", ")} +${list.length - 6} more`;
  };
  const viewLabel = currentView === "carrierGroup"
    ? "Carrier Group"
    : currentView === "availability"
      ? "Availability"
      : "Carrier";
  const zoomRange =
    Number.isFinite(f.zoomMin) || Number.isFinite(f.zoomMax)
      ? `${Number.isFinite(f.zoomMin) ? money(f.zoomMin) : "Auto"} to ${Number.isFinite(f.zoomMax) ? money(f.zoomMax) : "Auto"}`
      : "Auto";

  return {
    viewLabel,
    view: currentView,
    insuranceProgram: String(f.insuranceProgram || "").trim() || "All",
    policyLimitType: String(f.policyLimitType || "").trim() || "All",
    yearRange,
    annualized: f.annualized ? "On" : "Off",
    zoomRange,
    carriers: summarizeList(carriers, "All"),
    carrierGroups: summarizeList(carrierGroups, "All")
  };
}

function getExportFilterLines(meta = getFilterMeta()) {
  return [
    `View: ${meta.viewLabel} | Annualized: ${meta.annualized} | Zoom Range: ${meta.zoomRange}`,
    `Insurance Program: ${meta.insuranceProgram} | Policy Limit Type: ${meta.policyLimitType}`,
    `Period: ${meta.yearRange}`,
    `Carriers: ${meta.carriers}`,
    `Carrier Groups: ${meta.carrierGroups}`
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
        Year: p?.yearLabel ?? p?.x ?? "",
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

function normalizeAvailabilityLabel(raw) {
  const s = String(raw || "").trim();
  if (!s) return "Available";
  if (/unavail/i.test(s)) return "Unavailable";
  if (/partial/i.test(s)) return "Partially Available";
  return s;
}

function getCoverageReportFacts() {
  const rows = getFilteredSliceRows();
  const policySet = new Set();
  const policyNumberSet = new Set();
  const carrierSet = new Set();
  const carrierGroupSet = new Set();
  const programSet = new Set();
  const yearMap = new Map();
  const carrierMap = new Map();
  const availabilityMap = new Map();
  const programMap = new Map();
  const limitTypeMap = new Map();

  let totalLayer = 0;
  let availableLayer = 0;
  let unavailableLayer = 0;
  let minAttachment = Number.POSITIVE_INFINITY;
  let maxAttachment = Number.NEGATIVE_INFINITY;
  let minLayer = Number.POSITIVE_INFINITY;
  let maxLayer = Number.NEGATIVE_INFINITY;

  for (const row of rows) {
    const policyKey = String(row.PolicyID || row.PolicyNumber || "").trim();
    if (policyKey) policySet.add(policyKey);
    const policyNumber = String(row.PolicyNumber || "").trim();
    if (policyNumber) policyNumberSet.add(policyNumber);

    const carrier = String(row.Carrier || "").trim() || "(unknown carrier)";
    const carrierGroup = String(row.CarrierGroup || "").trim() || "(unknown group)";
    const program = String(row.InsuranceProgram || "").trim() || "(unknown program)";
    const limitType = String(row.PolicyLimitType || "").trim() || "(unknown type)";
    carrierSet.add(carrier);
    carrierGroupSet.add(carrierGroup);
    programSet.add(program);

    const layer = Number(row.LayerLimit || 0);
    const attach = Number(row.Attachment || 0);
    totalLayer += layer;
    if (Number.isFinite(attach)) {
      minAttachment = Math.min(minAttachment, attach);
      maxAttachment = Math.max(maxAttachment, attach);
    }
    if (Number.isFinite(layer)) {
      minLayer = Math.min(minLayer, layer);
      maxLayer = Math.max(maxLayer, layer);
    }

    const availability = normalizeAvailabilityLabel(row.Availability);
    availabilityMap.set(availability, (availabilityMap.get(availability) || 0) + 1);
    if (/unavail/i.test(availability)) unavailableLayer += layer;
    else availableLayer += layer;

    const year = String(row.Year || "Unknown");
    if (!yearMap.has(year)) yearMap.set(year, { year, rows: 0, layer: 0, available: 0 });
    const yearAgg = yearMap.get(year);
    yearAgg.rows += 1;
    yearAgg.layer += layer;
    if (!/unavail/i.test(availability)) yearAgg.available += layer;

    if (!carrierMap.has(carrier)) carrierMap.set(carrier, { carrier, rows: 0, policies: new Set(), layer: 0, available: 0 });
    const carrierAgg = carrierMap.get(carrier);
    carrierAgg.rows += 1;
    carrierAgg.layer += layer;
    if (!/unavail/i.test(availability)) carrierAgg.available += layer;
    if (policyKey) carrierAgg.policies.add(policyKey);

    if (!programMap.has(program)) programMap.set(program, { program, rows: 0, policies: new Set(), layer: 0 });
    const programAgg = programMap.get(program);
    programAgg.rows += 1;
    programAgg.layer += layer;
    if (policyKey) programAgg.policies.add(policyKey);

    if (!limitTypeMap.has(limitType)) limitTypeMap.set(limitType, { limitType, rows: 0, policies: new Set(), layer: 0 });
    const limitTypeAgg = limitTypeMap.get(limitType);
    limitTypeAgg.rows += 1;
    limitTypeAgg.layer += layer;
    if (policyKey) limitTypeAgg.policies.add(policyKey);
  }

  const yearRows = Array.from(yearMap.values()).sort((a, b) => {
    const ay = Number(a.year);
    const by = Number(b.year);
    if (Number.isFinite(ay) && Number.isFinite(by)) return ay - by;
    return String(a.year).localeCompare(String(b.year));
  });

  const carrierRows = Array.from(carrierMap.values())
    .map((row) => ({
      carrier: row.carrier,
      rows: row.rows,
      policyCount: row.policies.size,
      layer: row.layer,
      available: row.available
    }))
    .sort((a, b) => b.layer - a.layer || a.carrier.localeCompare(b.carrier));

  const programRows = Array.from(programMap.values())
    .map((row) => ({
      program: row.program,
      rows: row.rows,
      policyCount: row.policies.size,
      layer: row.layer
    }))
    .sort((a, b) => b.layer - a.layer || a.program.localeCompare(b.program));

  const limitTypeRows = Array.from(limitTypeMap.values())
    .map((row) => ({
      limitType: row.limitType,
      rows: row.rows,
      policyCount: row.policies.size,
      layer: row.layer
    }))
    .sort((a, b) => b.layer - a.layer || a.limitType.localeCompare(b.limitType));

  const availabilityRows = Array.from(availabilityMap.entries())
    .map(([availability, count]) => ({ availability, count }))
    .sort((a, b) => b.count - a.count || a.availability.localeCompare(b.availability));

  return {
    rowsCount: rows.length,
    uniquePolicies: policySet.size,
    uniquePolicyNumbers: policyNumberSet.size,
    uniqueCarriers: carrierSet.size,
    uniqueCarrierGroups: carrierGroupSet.size,
    uniquePrograms: programSet.size,
    totalLayer,
    availableLayer,
    unavailableLayer,
    minAttachment: Number.isFinite(minAttachment) ? minAttachment : null,
    maxAttachment: Number.isFinite(maxAttachment) ? maxAttachment : null,
    minLayer: Number.isFinite(minLayer) ? minLayer : null,
    maxLayer: Number.isFinite(maxLayer) ? maxLayer : null,
    yearRows,
    carrierRows,
    programRows,
    limitTypeRows,
    availabilityRows
  };
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
  const headerHeight = Math.round(topPad + titleSize + Math.max(1, metaLines.length) * (textSize + lineGap) + 20 * scale);
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
  for (const line of metaLines) {
    y += lineGap + textSize;
    ctx.fillText(line, sidePad, y);
  }

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
 * Export a multi-page PDF report:
 * page 1 = report scope + key metrics + chart image
 * page 2+ = totals and filtered schedule tables
 */
export async function exportReportPDF() {
  if (!chart || !_cache.dom?.canvas) throw new Error("Chart is not initialized");
  await ensurePdfLibs();

  const html2canvas = window.html2canvas;
  const jsPDF = window.jspdf.jsPDF;

  const meta = getFilterMeta();
  const facts = getCoverageReportFacts();
  const filteredRows = getFilteredSliceRows();
  const aggregatedRows = getAggregatedReportRows();
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
  const bottomPad = 24;

  const ensureSpace = (state, needed = 18) => {
    if (state.y + needed <= pageH - margin - bottomPad) return false;
    pdf.addPage("a4", "landscape");
    state.y = margin;
    return true;
  };

  const addSectionTitle = (state, title) => {
    ensureSpace(state, 18);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(13);
    pdf.setTextColor(22, 34, 54);
    pdf.text(title, margin, state.y);
    state.y += 15;
  };

  const drawWrappedFactLine = (state, label, value) => {
    const full = `${label}: ${value}`;
    const lines = pdf.splitTextToSize(full, pageW - margin * 2);
    for (const line of lines) {
      ensureSpace(state, 12);
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(9.5);
      pdf.setTextColor(38, 52, 72);
      pdf.text(line, margin, state.y);
      state.y += 11;
    }
  };

  const drawGridTable = ({ state, title, columns, rows, rowToCells }) => {
    const tableW = columns.reduce((sum, col) => sum + col.width, 0);
    addSectionTitle(state, title);

    const drawHeader = () => {
      ensureSpace(state, 18);
      const top = state.y;
      let x = margin;
      pdf.setFillColor(231, 237, 247);
      pdf.setDrawColor(188, 198, 214);
      pdf.setLineWidth(0.4);
      pdf.rect(margin, top, tableW, 16, "FD");
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(8.5);
      pdf.setTextColor(30, 40, 58);
      for (const col of columns) {
        if (col.align === "right") {
          pdf.text(col.label, x + col.width - 3, top + 11, { align: "right" });
        } else {
          pdf.text(col.label, x + 3, top + 11);
        }
        x += col.width;
        if (x < margin + tableW) pdf.line(x, top, x, top + 16);
      }
      state.y += 16;
    };

    drawHeader();

    if (!rows.length) {
      ensureSpace(state, 16);
      pdf.setFont("helvetica", "italic");
      pdf.setFontSize(9);
      pdf.setTextColor(72, 84, 102);
      pdf.text("No rows for current filters.", margin + 3, state.y + 11);
      state.y += 18;
      return;
    }

    for (const row of rows) {
      if (ensureSpace(state, 16)) drawHeader();
      const top = state.y;
      let x = margin;
      const cells = rowToCells(row);
      pdf.setDrawColor(210, 219, 230);
      pdf.setLineWidth(0.3);
      pdf.rect(margin, top, tableW, 16);
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(8.3);
      pdf.setTextColor(36, 50, 70);
      for (let i = 0; i < columns.length; i++) {
        const col = columns[i];
        const clipped = pdf.splitTextToSize(String(cells[i] ?? ""), col.width - 6)[0] || "";
        if (col.align === "right") {
          pdf.text(clipped, x + col.width - 3, top + 11, { align: "right" });
        } else {
          pdf.text(clipped, x + 3, top + 11);
        }
        x += col.width;
        if (x < margin + tableW) pdf.line(x, top, x, top + 16);
      }
      state.y += 16;
    }

    state.y += 8;
  };

  // Page 1: report scope + key metrics + chart
  let pageState = { y: margin };
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(18);
  pdf.setTextColor(20, 32, 52);
  pdf.text("Insurance Program Coverage Tower Report", margin, pageState.y);
  pageState.y += 18;

  pdf.setFont("helvetica", "normal");
  pdf.setFontSize(10);
  pdf.setTextColor(60, 74, 96);
  pdf.text(`Generated: ${new Date().toLocaleString()}`, margin, pageState.y);
  pageState.y += 14;

  addSectionTitle(pageState, "Report Scope");
  for (const line of getExportFilterLines(meta)) {
    const wrapped = pdf.splitTextToSize(line, pageW - margin * 2);
    for (const subLine of wrapped) {
      ensureSpace(pageState, 12);
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(9.5);
      pdf.setTextColor(38, 52, 72);
      pdf.text(subLine, margin, pageState.y);
      pageState.y += 11;
    }
  }
  pageState.y += 3;

  const numericYears = facts.yearRows
    .map((row) => Number(row.year))
    .filter((year) => Number.isFinite(year));
  const yearSpan = numericYears.length
    ? `${Math.min(...numericYears)} to ${Math.max(...numericYears)}`
    : "N/A";
  const availabilitySummary = facts.availabilityRows.length
    ? facts.availabilityRows.map((row) => `${row.availability}: ${row.count}`).join(", ")
    : "N/A";
  const metricLines = [
    ["Filtered Slice Rows", String(facts.rowsCount)],
    ["Unique Policies", String(facts.uniquePolicies)],
    ["Distinct Policy Numbers", String(facts.uniquePolicyNumbers)],
    ["Unique Carriers", String(facts.uniqueCarriers)],
    ["Unique Carrier Groups", String(facts.uniqueCarrierGroups)],
    ["Unique Programs", String(facts.uniquePrograms)],
    ["Policy Year Span", yearSpan],
    ["Total Layer Limit", money(facts.totalLayer)],
    ["Available Layer Limit", money(facts.availableLayer)],
    ["Unavailable Layer Limit", money(facts.unavailableLayer)],
    [
      "Attachment Range",
      facts.minAttachment === null ? "N/A" : `${money(facts.minAttachment)} to ${money(facts.maxAttachment)}`
    ],
    [
      "Layer Range",
      facts.minLayer === null ? "N/A" : `${money(facts.minLayer)} to ${money(facts.maxLayer)}`
    ],
    ["Availability Buckets", availabilitySummary]
  ];
  addSectionTitle(pageState, "Key Facts");
  for (const [label, value] of metricLines) drawWrappedFactLine(pageState, label, value);

  if (pageState.y > pageH - margin - 160) {
    pdf.addPage("a4", "landscape");
    pageState = { y: margin };
  }
  addSectionTitle(pageState, "Chart Snapshot");
  const imgTop = pageState.y;
  const imgMaxW = pageW - margin * 2;
  const imgMaxH = Math.max(120, pageH - imgTop - 40);
  const imgRatio = renderCanvas.width / renderCanvas.height || 1;
  let imgW = imgMaxW;
  let imgH = imgW / imgRatio;
  if (imgH > imgMaxH) {
    imgH = imgMaxH;
    imgW = imgH * imgRatio;
  }
  pdf.addImage(chartImg, "PNG", margin, imgTop, imgW, imgH, undefined, "FAST");

  // Page 2+: totals by key groupings
  pdf.addPage("a4", "landscape");
  pageState = { y: margin };
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(16);
  pdf.setTextColor(20, 32, 52);
  pdf.text("Filtered Data Totals", margin, pageState.y);
  pageState.y += 18;

  drawGridTable({
    state: pageState,
    title: "Year Totals",
    columns: [
      { label: "Year", width: 92 },
      { label: "Rows", width: 72, align: "right" },
      { label: "Layer Limit", width: 130, align: "right" },
      { label: "Available Layer", width: 140, align: "right" }
    ],
    rows: facts.yearRows,
    rowToCells: (row) => [String(row.year), String(row.rows), money(row.layer), money(row.available)]
  });

  drawGridTable({
    state: pageState,
    title: "Program Totals",
    columns: [
      { label: "Program", width: 260 },
      { label: "Policies", width: 80, align: "right" },
      { label: "Rows", width: 70, align: "right" },
      { label: "Layer Limit", width: 130, align: "right" }
    ],
    rows: facts.programRows,
    rowToCells: (row) => [row.program, String(row.policyCount), String(row.rows), money(row.layer)]
  });

  drawGridTable({
    state: pageState,
    title: "Policy Limit Type Totals",
    columns: [
      { label: "Limit Type", width: 260 },
      { label: "Policies", width: 80, align: "right" },
      { label: "Rows", width: 70, align: "right" },
      { label: "Layer Limit", width: 130, align: "right" }
    ],
    rows: facts.limitTypeRows,
    rowToCells: (row) => [row.limitType, String(row.policyCount), String(row.rows), money(row.layer)]
  });

  drawGridTable({
    state: pageState,
    title: "Carrier Totals (Top 30 by Layer Limit)",
    columns: [
      { label: "Carrier", width: 260 },
      { label: "Policies", width: 80, align: "right" },
      { label: "Rows", width: 70, align: "right" },
      { label: "Layer Limit", width: 130, align: "right" },
      { label: "Available Layer", width: 140, align: "right" }
    ],
    rows: facts.carrierRows.slice(0, 30),
    rowToCells: (row) => [row.carrier, String(row.policyCount), String(row.rows), money(row.layer), money(row.available)]
  });

  // Aggregated layer summary page
  pdf.addPage("a4", "landscape");
  pageState = { y: margin };
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(16);
  pdf.setTextColor(20, 32, 52);
  pdf.text("Aggregated Layer Summary", margin, pageState.y);
  pageState.y += 18;

  drawGridTable({
    state: pageState,
    title: "Layer Stack by Year and Group",
    columns: [
      { label: "Year", width: 72 },
      { label: "Group", width: 320 },
      { label: "Attachment", width: 120, align: "right" },
      { label: "Total Limit", width: 120, align: "right" },
      { label: "Participants", width: 110, align: "right" }
    ],
    rows: aggregatedRows,
    rowToCells: (row) => [
      String(row.Year),
      String(row.Group),
      money(row.Attachment),
      money(row.TotalLimit),
      String(row.Participants)
    ]
  });

  // Filtered policy schedule page(s)
  pdf.addPage("a4", "landscape");
  pageState = { y: margin };
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(16);
  pdf.setTextColor(20, 32, 52);
  pdf.text("Filtered Policy Schedule", margin, pageState.y);
  pageState.y += 18;

  const scheduleRows = [...filteredRows].sort((a, b) => {
    const ay = Number(a.Year);
    const by = Number(b.Year);
    if (Number.isFinite(ay) && Number.isFinite(by) && ay !== by) return ay - by;
    const yearCmp = String(a.Year).localeCompare(String(b.Year));
    if (yearCmp !== 0) return yearCmp;
    if (Number(a.Attachment || 0) !== Number(b.Attachment || 0)) return Number(a.Attachment || 0) - Number(b.Attachment || 0);
    const carrierCmp = String(a.Carrier || "").localeCompare(String(b.Carrier || ""));
    if (carrierCmp !== 0) return carrierCmp;
    return String(a.PolicyNumber || "").localeCompare(String(b.PolicyNumber || ""));
  });

  drawGridTable({
    state: pageState,
    title: "Policy Rows Matching Current Filters",
    columns: [
      { label: "Year", width: 42 },
      { label: "Policy #", width: 82 },
      { label: "Carrier", width: 124 },
      { label: "Carrier Group", width: 104 },
      { label: "Program", width: 86 },
      { label: "Limit Type", width: 78 },
      { label: "Availability", width: 70 },
      { label: "Attachment", width: 86, align: "right" },
      { label: "Layer Limit", width: 86, align: "right" }
    ],
    rows: scheduleRows,
    rowToCells: (row) => [
      String(row.Year || ""),
      String(row.PolicyNumber || ""),
      String(row.Carrier || ""),
      String(row.CarrierGroup || ""),
      String(row.InsuranceProgram || ""),
      String(row.PolicyLimitType || ""),
      normalizeAvailabilityLabel(row.Availability),
      money(row.Attachment),
      money(row.LayerLimit)
    ]
  });

  const totalPages = pdf.getNumberOfPages();
  for (let p = 1; p <= totalPages; p++) {
    pdf.setPage(p);
    const w = pdf.internal.pageSize.getWidth();
    const h = pdf.internal.pageSize.getHeight();
    pdf.setFont("helvetica", "normal");
    pdf.setFontSize(9);
    pdf.setTextColor(112, 122, 138);
    pdf.text("Generated by Coverage Dashboard", margin, h - 12);
    pdf.text(`Page ${p} of ${totalPages}`, w - margin, h - 12, { align: "right" });
  }

  pdf.save(filename);
  console.log(`[Export] PDF saved: ${filename} (tableRows=${aggregatedRows.length})`);
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
  const isLinearYearAxis = _cache.useYearAxis && xScale.type === "linear";
  for (const lbl of _cache.xLabels) {
    const key = String(lbl ?? "");
    const value = isLinearYearAxis ? Number(key) : key;
    const px = Number(xScale.getPixelForValue(value));
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

export function setAnnualizedMode(enabled) {
  _cache.filters.annualized = !!enabled;
  applyFiltersToCache();
  rebuildChart();
}

export function getAnnualizedMode() {
  return !!_cache.filters?.annualized;
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

function pickPolicyParticipant(raw, datasetLabel) {
  const parts = Array.isArray(raw?.participants) ? raw.participants : [];
  if (!parts.length) return null;

  const label = String(datasetLabel || "").trim();
  if (raw?.isQuotaShare && label) {
    if (currentView === "carrier") {
      const match = parts.find((p) => String(p?.carrier || "").trim() === label);
      if (match) return match;
    } else if (currentView === "carrierGroup") {
      const match = parts.find((p) => String(p?.carrierGroup || "").trim() === label);
      if (match) return match;
    }
  }

  return parts
    .slice()
    .sort((a, b) => Number(b?.sliceLimit || 0) - Number(a?.sliceLimit || 0))[0] || parts[0];
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function policyInfoUrl(policyId, policyNumber = "", policyLimitType = "", insuranceProgram = "") {
  const params = new URLSearchParams();
  params.set("policyId", String(policyId || ""));
  const policyNo = String(policyNumber || "").trim();
  const limitType = String(policyLimitType || "").trim();
  const program = String(insuranceProgram || "").trim();
  if (policyNo) params.set("policyNumber", policyNo);
  if (limitType) params.set("policyLimitType", limitType);
  if (program) params.set("insuranceProgram", program);
  return `/Modules/PolicyInformation/index.html?${params.toString()}`;
}

function buildTooltipTitle(items) {
  const raw = items?.[0]?.raw || {};
  const yr = String(raw.yearLabel ?? raw.x ?? "");
  const ds = items?.[0]?.dataset || {};
  if (ds?.datasetId === "sirOverlay") return `${yr} — ${ds.label || "SIR"}`;
  if (raw.isQuotaShare) return `${yr} — Quota share`;
  const g = String(raw.group ?? "").trim();
  return g ? `${yr} — ${g}` : yr;
}

function buildTooltipLines(ctx, r) {
  if (ctx?.dataset?.datasetId === "sirOverlay") {
    const val = Number(ctx?.raw?.y ?? ctx?.parsed?.y ?? 0);
    return [`${ctx.dataset.label}: ${money(val)}`];
  }

  const attach = r.attach ?? 0;
  const top = r.top ?? 0;
  const lim = Math.max(0, top - attach);

  const lines = [];
  lines.push(`Attach: ${money(attach)}`);
  lines.push(`Limit: ${money(lim)}`);
  lines.push(`Top: ${money(top)}`);
  const parts = Array.isArray(r.participants) ? r.participants : [];
  if (r.annualized) {
    const segStartVals = parts
      .map((p) => Number(p?.segmentStartMs || p?.policyStartMs || 0))
      .filter((v) => Number.isFinite(v) && v > 0);
    const segEndVals = parts
      .map((p) => Number(p?.segmentEndMs || p?.policyEndMs || 0))
      .filter((v) => Number.isFinite(v) && v > 0);
    if (segStartVals.length && segEndVals.length) {
      const segStart = Math.min(...segStartVals);
      const segEnd = Math.max(...segEndVals);
      lines.push(`Policy Start: ${formatFullDateUTC(segStart)}`);
      lines.push(`Policy End: ${formatFullDateUTC(segEnd)}`);
      const isMultiYear = parts.some((p) => {
        const fullStart = Number(p?.policyStartMs || 0);
        const fullEnd = Number(p?.policyEndMs || 0);
        const pSegStart = Number(p?.segmentStartMs || fullStart || 0);
        const pSegEnd = Number(p?.segmentEndMs || fullEnd || 0);
        return Number.isFinite(fullStart) &&
          Number.isFinite(fullEnd) &&
          Number.isFinite(pSegStart) &&
          Number.isFinite(pSegEnd) &&
          (fullStart < pSegStart || fullEnd > pSegEnd);
      });
      if (isMultiYear) lines.push("Multi-year policy segment");
    }
  } else {
    const startVals = parts
      .map((p) => Number(p?.policyStartMs || 0))
      .filter((v) => Number.isFinite(v) && v > 0);
    const endVals = parts
      .map((p) => Number(p?.policyEndMs || 0))
      .filter((v) => Number.isFinite(v) && v > 0);
    if (startVals.length && endVals.length) {
      const minStart = Math.min(...startVals);
      const maxStart = Math.max(...startVals);
      const minEnd = Math.min(...endVals);
      const maxEnd = Math.max(...endVals);
      if (minStart === maxStart && minEnd === maxEnd) {
        lines.push(`Policy Start: ${formatFullDateUTC(minStart)}`);
        lines.push(`Policy End: ${formatFullDateUTC(maxEnd)}`);
      } else {
        lines.push(`Policy Start (earliest): ${formatFullDateUTC(minStart)}`);
        lines.push(`Policy End (latest): ${formatFullDateUTC(maxEnd)}`);
      }
    }
  }

  const isPrimaryLayer = Number(attach) <= 0;
  if (isPrimaryLayer) {
    const sirValsPerOcc = parts
      .map((p) => Number(p?.sirPerOcc || 0))
      .filter((v) => Number.isFinite(v) && v > 0);
    const sirValsAgg = parts
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

  const quotaParts = r.isQuotaShare && r.quotaGroupKey
    ? parts.filter((p) => String(p?.quotaGroupKey || "") === String(r.quotaGroupKey))
    : parts;
  const hideUnavailable = isUnavailableLegendHidden(ctx?.chart);
  const shownQuotaParts = hideUnavailable
    ? quotaParts.filter((p) => !String(p?.availability || "").toLowerCase().includes("unavail"))
    : quotaParts;
  const isUnavailableGroup = String(r.group || "").toLowerCase() === "unavailable";
  const shouldShowParts = !!r.isQuotaShare && shownQuotaParts.length > 1;

  if (isUnavailableGroup && shownQuotaParts.length) {
    const uniqueCarriers = [...new Set(shownQuotaParts.map((p) => p.carrier || "(unknown carrier)"))];
    const carrierLine =
      uniqueCarriers.length === 1
        ? `Carrier: ${uniqueCarriers[0]}`
        : `Carriers: ${uniqueCarriers.join(", ")}`;
    lines.push(carrierLine);
  }

  if (shouldShowParts && shownQuotaParts.length) {
    lines.push(`Quota share participants (${shownQuotaParts.length}):`);
    const tooltipMaxParticipants = Number(_cache.options?.tooltipMaxParticipants || 25);
    const show = shownQuotaParts.slice(0, tooltipMaxParticipants);
    for (const p of show) {
      const carrier = p.carrier || "(unknown carrier)";
      lines.push(`• ${carrier}: ${money(p.sliceLimit)}`);
    }
    if (shownQuotaParts.length > tooltipMaxParticipants) {
      lines.push(`… +${shownQuotaParts.length - tooltipMaxParticipants} more`);
    }
  }

  return lines;
}

function getPolicyLinksForTooltip(ctx, r) {
  const parts = Array.isArray(r?.participants) ? r.participants : [];
  if (!parts.length) return [];
  const hideUnavailable = isUnavailableLegendHidden(ctx?.chart);
  const quotaParts = r.isQuotaShare && r.quotaGroupKey
    ? parts.filter((p) => String(p?.quotaGroupKey || "") === String(r.quotaGroupKey))
    : parts;
  const shownParts = hideUnavailable
    ? quotaParts.filter((p) => !String(p?.availability || "").toLowerCase().includes("unavail"))
    : quotaParts;
  const source = shownParts.length ? shownParts : quotaParts;
  if (!source.length) return [];

  if (r.isQuotaShare) {
    const dedup = new Map();
    for (const p of source) {
      const pid = String(p?.pid || "").trim();
      if (!pid || dedup.has(pid)) continue;
      const carrier = String(p?.carrier || "(unknown carrier)").trim();
      const policyNo = String(p?.policy_no || "").trim();
      const policyLimitType = String(p?.policyLimitType || "").trim();
      const insuranceProgram = String(p?.insuranceProgram || "").trim();
      const label = policyNo ? `${carrier} (${policyNo})` : `${carrier} (Policy ${pid})`;
      dedup.set(pid, { policyId: pid, policyNumber: policyNo, policyLimitType, insuranceProgram, label });
    }
    return Array.from(dedup.values());
  }

  const dsLabel = String(ctx?.dataset?.label || r?.group || "").trim();
  const picked = pickPolicyParticipant(r, dsLabel);
  const pid = String(picked?.pid || "").trim();
  if (!pid) return [];
  const carrier = String(picked?.carrier || "(unknown carrier)").trim();
  const policyNo = String(picked?.policy_no || "").trim();
  const policyLimitType = String(picked?.policyLimitType || "").trim();
  const insuranceProgram = String(picked?.insuranceProgram || "").trim();
  const label = policyNo ? `${carrier} (${policyNo})` : `${carrier} (Policy ${pid})`;
  return [{ policyId: pid, policyNumber: policyNo, policyLimitType, insuranceProgram, label }];
}

function getOrCreateHtmlTooltip(chartInstance) {
  let el = document.body.querySelector(".coverageHtmlTooltip");
  if (el) return el;
  el = document.createElement("div");
  el.className = "coverageHtmlTooltip";
  Object.assign(el.style, {
    position: "fixed",
    transform: "translate(0, 0)",
    background: "rgba(2, 6, 23, 0.94)",
    color: "rgba(248, 250, 252, 0.96)",
    border: "1px solid rgba(148, 163, 184, 0.35)",
    borderRadius: "10px",
    padding: "10px 12px",
    font: "600 14px/1.35 ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif",
    boxShadow: "0 14px 30px rgba(2, 6, 23, 0.38)",
    pointerEvents: "auto",
    zIndex: "50",
    minWidth: "260px",
    maxWidth: "460px",
    whiteSpace: "normal",
    opacity: "0"
  });
  el.dataset.locked = "0";
  el.addEventListener("mouseenter", () => {
    el.dataset.locked = "1";
  });
  el.addEventListener("mouseleave", () => {
    el.dataset.locked = "0";
    el.style.opacity = "0";
  });
  document.body.appendChild(el);
  return el;
}

function externalCoverageTooltipHandler(context) {
  const { chart: chartInstance, tooltip } = context;
  const tooltipEl = getOrCreateHtmlTooltip(chartInstance);
  if (!tooltip || tooltip.opacity === 0 || !tooltip.dataPoints?.length) {
    if (tooltipEl.dataset.locked === "1") return;
    tooltipEl.style.opacity = "0";
    return;
  }
  const point = tooltip.dataPoints[0];
  const r = point?.raw || {};
  const title = buildTooltipTitle([point]);
  const lines = buildTooltipLines(point, r);
  const links = getPolicyLinksForTooltip(point, r);

  const linesHtml = lines.map((ln) => `<div>${escapeHtml(ln)}</div>`).join("");
  const linksHtml = links.length
    ? `<div style="margin-top:8px; border-top:1px solid rgba(148,163,184,0.25); padding-top:7px;">
         <div style="font-weight:700; margin-bottom:4px;">Policy link${links.length > 1 ? "s" : ""}:</div>
         ${links
           .map(
             (l) =>
               `<div><a href="${policyInfoUrl(l.policyId, l.policyNumber, l.policyLimitType, l.insuranceProgram)}" target="_blank" rel="noopener noreferrer" style="color:#93c5fd; text-decoration:underline;">${escapeHtml(l.label)}</a></div>`
           )
           .join("")}
       </div>`
    : "";

  tooltipEl.innerHTML = `
    <div style="font-size:13px; font-weight:700; margin-bottom:5px;">${escapeHtml(title)}</div>
    <div style="display:grid; gap:2px; font-size:12px; font-weight:500;">${linesHtml}</div>
    ${linksHtml}
  `;

  const rect = chartInstance.canvas.getBoundingClientRect();
  const offsetX = 18;
  const offsetY = 18;

  tooltipEl.style.left = "0px";
  tooltipEl.style.top = "0px";
  tooltipEl.style.opacity = "1";
  const tipW = tooltipEl.offsetWidth || 320;
  const tipH = tooltipEl.offsetHeight || 140;

  const minLeft = 8;
  const minTop = 8;
  const maxLeft = window.innerWidth - tipW - 8;
  const maxTop = window.innerHeight - tipH - 8;
  const rawLeft = rect.left + tooltip.caretX + offsetX;
  const rawTop = rect.top + tooltip.caretY + offsetY;
  const left = clamp(rawLeft, minLeft, Math.max(minLeft, maxLeft));
  const top = clamp(rawTop, minTop, Math.max(minTop, maxTop));

  tooltipEl.style.left = `${left}px`;
  tooltipEl.style.top = `${top}px`;
  tooltipEl.style.opacity = "1";
}

export function getPolicySelectionFromEvent(evt) {
  if (!chart || !evt) return null;
  const hits = chart.getElementsAtEventForMode(evt, "nearest", { intersect: true }, true);
  if (!Array.isArray(hits) || !hits.length) return null;
  const hit = hits[0];
  const ds = chart.data?.datasets?.[hit.datasetIndex];
  if (!ds || ds?.datasetId === "sirOverlay") return null;
  const raw = ds?.data?.[hit.index];
  if (!raw) return null;

  const participant = pickPolicyParticipant(raw, ds?.label || raw?.group || "");
  if (!participant) return null;

  const attach = Number(raw?.attach || 0);
  const top = Number(raw?.top || 0);
  const limit = Math.max(0, top - attach);

  return {
    policyId: String(participant?.pid || ""),
    policyNumber: String(participant?.policy_no || ""),
    policyLimitType: String(participant?.policyLimitType || ""),
    insuranceProgram: String(participant?.insuranceProgram || ""),
    carrier: String(participant?.carrier || ""),
    carrierGroup: String(participant?.carrierGroup || ""),
    availability: String(participant?.availability || ""),
    policyStartMs: Number(participant?.policyStartMs || 0),
    policyEndMs: Number(participant?.policyEndMs || 0),
    segmentStartMs: Number(participant?.segmentStartMs || 0),
    segmentEndMs: Number(participant?.segmentEndMs || 0),
    attach,
    limit,
    top,
    yearLabel: String(raw?.yearLabel || raw?.x || ""),
    isQuotaShare: !!raw?.isQuotaShare,
    view: currentView
  };
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
  if (!window.Chart) throw new Error("Chart.js must be loaded before coverageChartEngine.js");

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

  const quotaEnabled = hasExplicitQuotaShareEvidence({ limitsRows, policyRows });
  const quotaKeySet = quotaEnabled ? buildQuotaKeySet(built.slices) : new Set();

  _cache = {
    allSlices: built.slices,
    allXLabels: built.xLabels,
    slices: built.slices,
    xLabels: built.xLabels,
    options: { barThickness, categorySpacing, tooltipMaxParticipants },
    quotaKeySet,
    useYearAxis,
    xZoom: _cache.xZoom || 1,
    dom: {
      canvas,
      viewport: canvas.closest(".chartViewport"),
      surface: canvas.parentElement
    },
    _wheelBound: _cache._wheelBound || false,
    hiddenLegendCarriers: _cache.hiddenLegendCarriers || new Set(),
    hiddenLegendCarrierGroups: _cache.hiddenLegendCarrierGroups || new Set(),
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
      annualized: !!_cache.filters?.annualized,
      carriers: [],
      carrierGroups: []
    }
  };

  // Default limit-type filter on first render to avoid stacking multiple
  // limit types (for example BI + PD) at the same attachment.
  if (!String(_cache.filters.policyLimitType || "").trim()) {
    const limitTypeOptions = getFilterOptions().policyLimitTypes || [];
    const bodilyInjury =
      limitTypeOptions.find((t) => String(t || "").trim().toLowerCase() === "bodily injury") || "";
    _cache.filters.policyLimitType = bodilyInjury || String(limitTypeOptions[0] || "").trim();
  }

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
  const yearAxisBounds = getYearAxisBounds(_cache.xLabels);

  if (chart) chart.destroy();

  chart = new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: _cache.xLabels,
      datasets: sirDataset ? [...datasets, sirDataset] : datasets
    },
    plugins: [xRangeBarsPlugin, outlineBarsPlugin, quotaShareGuidesPlugin, boxValueLabelsPlugin, yearAvailableTotalsPlugin],
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
          onClick: (evt, legendItem, legend) => {
            if (legendItem?._syntheticCarrier) {
              const value = String(legendItem?.text || "").trim();
              if (!value) return;
              const hiddenSet = currentView === "carrier"
                ? (_cache.hiddenLegendCarriers || (_cache.hiddenLegendCarriers = new Set()))
                : currentView === "carrierGroup"
                ? (_cache.hiddenLegendCarrierGroups || (_cache.hiddenLegendCarrierGroups = new Set()))
                : null;
              if (!hiddenSet) return;
              if (hiddenSet.has(value)) hiddenSet.delete(value);
              else hiddenSet.add(value);
              rebuildChart();
              return;
            }
            const defaultClick = Chart?.defaults?.plugins?.legend?.onClick;
            if (typeof defaultClick === "function") {
              defaultClick(evt, legendItem, legend);
            }
          },
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
            font: { size: 11 },
            filter: (legendItem) => {
              const isQuota = String(legendItem?.text || "").trim().toLowerCase() === "quota share";
              if (!isQuota) return true;
              // Hide quota-share legend item in carrier & carrier-group views
              // (we render quota by participant colors, but legend should list
              // concrete carriers/groups instead of a quota bucket).
              return !(currentView === "carrier" || currentView === "carrierGroup");
            },
            generateLabels: (chartInstance) => {
              const base = Chart.defaults.plugins.legend.labels.generateLabels(chartInstance) || [];
              const labels = (currentView === "carrier" || currentView === "carrierGroup")
                ? base.filter((item) => String(item?.text || "").trim().toLowerCase() !== "quota share")
                : base.slice();
              if (!(currentView === "carrier" || currentView === "carrierGroup")) return labels;

              const valueKey = currentView === "carrier" ? "carrier" : "carrierGroup";
              const unknownValue = currentView === "carrier" ? "(unknown carrier)" : "(unknown group)";

              const existing = new Set(
                labels.map((item) => String(item?.text || "").trim().toLowerCase()).filter(Boolean)
              );
              const values = Array.from(
                new Set(
                  (_cache.slices || [])
                    .map((s) => String(s?.[valueKey] || "").trim())
                    .filter((v) => v && v !== unknownValue)
                )
              ).sort((a, b) => a.localeCompare(b));

              for (const value of values) {
                const key = value.toLowerCase();
                if (existing.has(key)) continue;
                const hiddenSet = currentView === "carrier"
                  ? (_cache.hiddenLegendCarriers || new Set())
                  : currentView === "carrierGroup"
                  ? (_cache.hiddenLegendCarrierGroups || new Set())
                  : new Set();
                labels.push({
                  text: value,
                  fillStyle: colorFromString(value),
                  strokeStyle: colorFromString(value),
                  lineWidth: 0,
                  hidden: hiddenSet.has(value),
                  datasetIndex: -1,
                  index: -1,
                  _syntheticCarrier: true
                });
              }
              return labels;
            }
          }
        },
        tooltip: {
          enabled: false,
          external: externalCoverageTooltipHandler,
          displayColors: false,
          mode: "nearest",
          intersect: true,
          filter: (_item, index) => index === 0,
          callbacks: {
            title: (items) => buildTooltipTitle(items),
            label: (ctx) => buildTooltipLines(ctx, ctx.raw || {})
          }
        }
      },

      scales: {
        x: useYearAxis
          ? {
              type: "linear",
              offset: false,
              min: yearAxisBounds.min,
              max: yearAxisBounds.max,
              grid: {
                display: true,
                color: themeColors.xGrid,
                lineWidth: 1
              },
              ticks: {
                color: themeColors.axisTicks,
                stepSize: 1,
                autoSkip: false,
                maxRotation: 0,
                minRotation: 0,
                callback: (v) => {
                  const n = Number(v);
                  if (!Number.isFinite(n)) return "";
                  const yr = Math.round(n);
                  const hasYear = (_cache.xLabels || []).some((lbl) => Number(lbl) === yr);
                  return hasYear ? String(yr) : "";
                }
              },
              title: {
                display: true,
                text: "Policy Years",
                color: themeColors.yTitle
              }
            }
          : {
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
  bindResponsiveResizeHandler();
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
