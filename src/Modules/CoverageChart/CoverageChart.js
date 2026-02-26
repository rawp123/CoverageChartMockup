// CoverageChart_TOWER.js
// Renders a "coverage tower" per year using floating bars (y=[start,end]) so layers stack by attachment.
//
// Supports (via exported setters; UI can live outside this module):
//   - View: "carrier" | "carrierGroup" | "availability"
//   - Year window (Start/End), inclusive
//   - Zoom range (Y-axis min/max)
//
// Notes:
// - Each input row is treated as a distinct policy slice unless it is part of a quota share layer.
// - Quota share (bquotashr==1): rows that share the same attachment + quotashrlim (+ term) are treated as
//   part of the same layer and are stacked *within that layer* so the overall layer height is preserved.
//
// Expected JSON fields (as in your sample):
//   Incept, term, attach, perocclim, bquotashr, quotashrlim, Carrier, CarrierGroup, policy_no, Consume, TotCost

let activeChart = null;
let cachedPolicies = [];
let currentView = "carrier";

let yearWindow = { start: null, end: null }; // inclusive
let zoomRange = null; // { min: number|null, max: number|null }

/* =========================
   Palette
========================= */
const AM_PALETTE = [
  "#500000", "#6B1F1F", "#7A2E2E", "#8A3B3B",
  "#1F2937", "#374151", "#4B5563", "#6B7280",
  "#0F766E", "#1D4ED8", "#7C3AED", "#B45309",
  "#A21CAF", "#0B7285",
];

const AVAILABLE_COLOR = "#2F855A";   // muted green
const UNAVAILABLE_COLOR = "#9B2C2C"; // muted maroon-red

/* ======================================================
   PUBLIC API
====================================================== */

export function renderCoverageChart({ elementId, policies }) {
  if (!window.Chart) throw new Error("Chart.js must be loaded via CDN before this module.");

  const canvas = typeof elementId === "string" ? document.getElementById(elementId) : elementId;
  if (!canvas) throw new Error("Canvas element not found.");

  cachedPolicies = Array.isArray(policies) ? policies : [];

  if (activeChart) {
    activeChart.destroy();
    activeChart = null;
  }

  initializeChart(canvas);
}

export function setView(view) {
  if (!view) return;
  currentView = String(view);
  rebuild();
}

export function setYearCaps(startYear, endYear) {
  const toYearOrNull = (v) => {
    if (v === null || v === undefined || v === "") return null;
    const n = parseInt(v, 10);
    return Number.isFinite(n) ? n : null;
  };

  let s = toYearOrNull(startYear);
  let e = toYearOrNull(endYear);

  // Normalize reversed inputs
  if (s !== null && e !== null && s > e) [s, e] = [e, s];

  yearWindow = { start: s, end: e };
  rebuild();
}

export function setZoomRange(min, max) {
  const mn = Number.isFinite(Number(min)) ? Number(min) : null;
  const mx = Number.isFinite(Number(max)) ? Number(max) : null;

  let zmin = mn, zmax = mx;
  if (zmin !== null && zmax !== null && zmin > zmax) [zmin, zmax] = [zmax, zmin];

  zoomRange = (zmin !== null || zmax !== null) ? { min: zmin, max: zmax } : null;

  // Update scale without rebuilding points
  if (activeChart) {
    activeChart.options.scales.y.min = zoomRange?.min ?? 0;
    activeChart.options.scales.y.max = zoomRange?.max ?? undefined;
    activeChart.update();
  }
}

/* ======================================================
   CHART BUILD
====================================================== */

function initializeChart(canvas) {
  const model = buildTowerModel();

  activeChart = new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: model.years.map(String),
      datasets: [
        {
          label: "Coverage",
          data: model.points, // floating bars: {x, y:[start,end], ...meta}
          borderWidth: 0,
          borderSkipped: false,
          borderRadius: 3,
          backgroundColor: (ctx) => {
            const raw = ctx.raw;
            if (!raw) return "rgba(255,255,255,0.3)";
            return raw.color || "rgba(255,255,255,0.3)";
          }
        }
      ]
    },
    options: buildChartOptions(model)
  });

  // Attach quota metadata for tooltip callouts
  activeChart._quotaDetailsByYear = model.quotaDetailsByYear || {};
  updateLegend(model.legendItems);
}

function buildChartOptions(model) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 160 },
    layout: { padding: { top: 8, right: 12, bottom: 8, left: 8 } },
    plugins: {
      legend: { display: false }, // we build our own legend in HTML
      tooltip: {
        callbacks: {
          title: (items) => `Year: ${items?.[0]?.label ?? ""}`,
          label: (ctx) => {
            const r = ctx.raw || {};
            const bucket = r.bucket || "Policy";
            const pol = r.policy_no ? ` • ${r.policy_no}` : "";
            const carrier = r.carrier ? ` • ${r.carrier}` : "";
            const rg = ` ${formatMoney(r.start)} – ${formatMoney(r.end)}`;
            return `${bucket}${pol}${carrier}${rg}`;
          },
          afterBody: (items) => {
            const first = items?.[0];
            const raw = first?.raw || {};

            // Only show quota share details when the hovered segment is quota share.
            if (!raw.isQuotaShare) return [];

            const yearStr = first?.label;
            const year = parseInt(yearStr, 10);
            if (!Number.isFinite(year)) return [];

            const chart = first?.chart;
            const layers = chart?._quotaDetailsByYear?.[year];
            if (!layers) return [];

            const qKey = raw.quotaLayerKey;
            const layer = qKey ? layers[qKey] : null;
            if (!layer) return [];

            const lines = [];
            lines.push("");
            lines.push("Quota Share Participants (this layer):");
            lines.push(`• ${layer.displayLabel}`);

            const MAX_PARTS = 12;
            const parts = layer.participants || [];
            parts.slice(0, MAX_PARTS).forEach((p) => {
              const who = p.carrier || p.bucketLabel || "Unknown";
              lines.push(`   - ${who}: ${formatMoney(p.sliceLimit || 0)}`);
            });

            if (parts.length > MAX_PARTS) lines.push(`   - (+${parts.length - MAX_PARTS} more)`);
            return lines;
          }
        }
      }
    },
    scales: {
      x: {
        type: "category",
        offset: true,
        grid: { display: false },
        ticks: { autoSkip: false, maxRotation: 0, minRotation: 0, font: { size: 11 } },
        title: { display: true, text: "Year", font: { size: 12 } }
      },
      y: {
        type: "linear",
        beginAtZero: true,
        min: zoomRange?.min ?? 0,
        max: zoomRange?.max ?? undefined,
        grid: { color: "rgba(255,255,255,0.10)" },
        ticks: { callback: (v) => formatMoney(v), font: { size: 11 } },
        title: { display: true, text: "Limits", font: { size: 12 } }
      }
    }
  };
}

/* ======================================================
   MODEL: turn rows into tower segments
====================================================== */

function buildTowerModel() {
  const byYear = new Map();

  for (const p of cachedPolicies) {
    const year = parseYear(p?.Incept);
    if (!Number.isFinite(year)) continue;

    if (!byYear.has(year)) byYear.set(year, []);
    byYear.get(year).push(p);
  }

  if (byYear.size === 0) return { years: [], points: [], legendItems: [], quotaDetailsByYear: {} };

  const dataYears = Array.from(byYear.keys()).sort((a, b) => a - b);
  const dataMin = dataYears[0];
  const dataMax = dataYears[dataYears.length - 1];

  const start = yearWindow.start ?? dataMin;
  const end = yearWindow.end ?? dataMax;

  const years = [];
  for (let y = start; y <= end; y++) years.push(y);

  const points = [];
  const legendBuckets = new Map(); // bucket -> color
  const quotaDetailsByYear = {};   // year -> layerKey -> {displayLabel, participants:[]}

  for (const year of years) {
    const rows = byYear.get(year) || [];
    if (!rows.length) continue;

    const slices = rows.map((r) => normalizeRow(r));

    // Sort by attachment, then quota-vs-nonquota, then by key
    slices.sort((a, b) => {
      if (a.attach !== b.attach) return a.attach - b.attach;
      if (a.isQuotaShare !== b.isQuotaShare) return a.isQuotaShare ? -1 : 1; // QS first within attach
      if (a.layerKey !== b.layerKey) return a.layerKey.localeCompare(b.layerKey);
      return (a.sortKey || "").localeCompare(b.sortKey || "");
    });

    // Group by attach + layerKey
    const groups = new Map(); // gid -> array slices
    for (const s of slices) {
      const gid = `${s.attach}||${s.layerKey}`;
      if (!groups.has(gid)) groups.set(gid, []);
      groups.get(gid).push(s);
    }

    const groupIds = Array.from(groups.keys()).sort((ga, gb) => {
      const [aa, ka] = ga.split("||");
      const [ab, kb] = gb.split("||");
      const da = parseFloat(aa), db = parseFloat(ab);
      if (da !== db) return da - db;
      return ka.localeCompare(kb);
    });

    quotaDetailsByYear[year] ??= {};

    for (const gid of groupIds) {
      const groupSlices = groups.get(gid) || [];
      if (!groupSlices.length) continue;

      const attach = groupSlices[0].attach;

      // IMPORTANT FIX:
      // - Non-quota rows should NOT be stacked within the same attachment point.
      //   Each distinct policy is its own layer, rendered from [attach, attach+limit] (overlap if same attach).
      // - Quota share rows ARE stacked within the quota layer so total layer height is preserved.
      const isQSLayer = groupSlices[0].isQuotaShare;

      // Within a QS layer, stack shares deterministically (largest first makes tooltip nicer)
      if (isQSLayer) groupSlices.sort((a, b) => (b.sliceLimit || 0) - (a.sliceLimit || 0));

      // Record QS participants for tooltip
      if (isQSLayer) {
        const qKey = groupSlices[0].quotaLayerKey;
        quotaDetailsByYear[year][qKey] ??= {
          displayLabel: groupSlices[0].quotaLayerLabel,
          participants: []
        };
        for (const s of groupSlices) {
          quotaDetailsByYear[year][qKey].participants.push({
            carrier: s.carrier,
            bucketLabel: s.bucket,
            sliceLimit: s.sliceLimit
          });
        }
      }

      // Build points
      if (isQSLayer) {
        let cursor = attach;

        for (const s of groupSlices) {
          const bucket = getBucketLabel(s);
          const color = getBucketColor(bucket, s);
          if (!legendBuckets.has(bucket)) legendBuckets.set(bucket, color);

          if (currentView === "availability") {
            const unavailable = Math.max(0, Math.min(s.sliceLimit, s.consumed));
            const available = Math.max(0, s.sliceLimit - unavailable);

            if (unavailable > 0) {
              points.push(makePoint(year, cursor, cursor + unavailable, "Unavailable", UNAVAILABLE_COLOR, s));
              cursor += unavailable;
            }
            if (available > 0) {
              points.push(makePoint(year, cursor, cursor + available, "Available", AVAILABLE_COLOR, s));
              cursor += available;
            }
          } else {
            points.push(makePoint(year, cursor, cursor + s.sliceLimit, bucket, color, s));
            cursor += s.sliceLimit;
          }
        }
      } else {
        // Non-QS: each slice is its own layer from [attach, attach+limit]
        for (const s of groupSlices) {
          const bucket = getBucketLabel(s);
          const color = getBucketColor(bucket, s);
          if (!legendBuckets.has(bucket)) legendBuckets.set(bucket, color);

          const startY = attach;
          const endY = attach + (s.sliceLimit || 0);

          if (currentView === "availability") {
            const unavailable = Math.max(0, Math.min(s.sliceLimit, s.consumed));
            const available = Math.max(0, s.sliceLimit - unavailable);

            if (unavailable > 0) points.push(makePoint(year, startY, startY + unavailable, "Unavailable", UNAVAILABLE_COLOR, s));
            if (available > 0) points.push(makePoint(year, startY + unavailable, endY, "Available", AVAILABLE_COLOR, s));
          } else {
            points.push(makePoint(year, startY, endY, bucket, color, s));
          }
        }
      }
    }

    // Sort quota participants within each layer (largest first)
    Object.keys(quotaDetailsByYear[year]).forEach((k) => {
      quotaDetailsByYear[year][k].participants.sort((a, b) => (b.sliceLimit || 0) - (a.sliceLimit || 0));
    });
  }

  const legendItems = Array.from(legendBuckets.entries())
    .map(([label, color]) => ({ label, color }))
    .sort((a, b) => a.label.localeCompare(b.label));

  return { years, points, legendItems, quotaDetailsByYear };
}

function normalizeRow(r) {
  const carrier = String(r?.Carrier || "Unknown Carrier").trim();
  const group = String(r?.CarrierGroup || "Other").trim();
  const policy_no = String(r?.policy_no || "").trim();

  const attach = parseMoney(r?.attach);
  const perocclim = parseMoney(r?.perocclim);
  const isQuotaShare = Number(r?.bquotashr) === 1;
  const quotashrlim = parseMoney(r?.quotashrlim);

  // Consumed / paid to date at the row level (if provided)
  const consumedRaw = (r?.Consume ?? r?.TotCost ?? 0);
  const consumed = parseMoney(consumedRaw);

  const term = String(r?.term || "").trim();

  const quotaLayerKey = buildQuotaLayerKey({ attach: r?.attach, quotashrlim: r?.quotashrlim, term });
  const quotaLayerLabel = buildQuotaLayerLabel({ attach: r?.attach, quotashrlim: r?.quotashrlim, term, policy_no });

  // Group key: for QS, use the quota layer; otherwise keep each policy distinct
  const layerKey = isQuotaShare ? `QS|${quotaLayerKey}` : `POL|${policy_no || carrier}|${term}`;

  const sortKey = `${carrier}||${policy_no}`;

  // For quota share, the overall layer limit is quotashrlim, but each row is a participant slice.
  // We draw slices by perocclim; typical data should have sum(perocclim)=quotashrlim.
  const sliceLimit = perocclim || 0;

  return {
    raw: r,
    carrier,
    group,
    policy_no,
    attach,
    sliceLimit,
    isQuotaShare,
    consumed,
    quotashrlim,
    quotaLayerKey,
    quotaLayerLabel,
    layerKey,
    sortKey,
    bucket: null
  };
}

function makePoint(year, start, end, bucket, color, s) {
  return {
    x: String(year),
    y: [start, end],
    start,
    end,
    bucket,
    color,
    carrier: s.carrier,
    group: s.group,
    policy_no: s.policy_no,
    attach: s.attach,
    sliceLimit: s.sliceLimit,
    consumed: s.consumed,
    isQuotaShare: s.isQuotaShare,
    quotaLayerKey: s.quotaLayerKey
  };
}

/* ======================================================
   LEGEND (HTML)
====================================================== */

function updateLegend(items) {
  const el = document.getElementById("chartLegend");
  if (!el) return;

  let legend = items;
  if (currentView === "availability") {
    legend = [
      { label: "Available", color: AVAILABLE_COLOR },
      { label: "Unavailable", color: UNAVAILABLE_COLOR },
    ];
  }

  const MAX = 50;
  const shown = legend.slice(0, MAX);
  const more = legend.length > MAX ? (legend.length - MAX) : 0;

  el.innerHTML = `
    <div class="legend-wrap">
      ${shown.map(i => `
        <div class="legend-item" title="${escapeHtml(i.label)}">
          <span class="swatch" style="background:${i.color}"></span>
          <span class="legend-label">${escapeHtml(i.label)}</span>
        </div>
      `).join("")}
      ${more ? `<div class="legend-more">(+${more} more)</div>` : ``}
    </div>
  `;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[c]));
}

/* ======================================================
   REBUILD
====================================================== */

function rebuild() {
  if (!activeChart) return;

  const model = buildTowerModel();

  activeChart.data.labels = model.years.map(String);
  activeChart.data.datasets[0].data = model.points;

  activeChart._quotaDetailsByYear = model.quotaDetailsByYear || {};
  updateLegend(model.legendItems);

  activeChart.options.scales.y.min = zoomRange?.min ?? 0;
  activeChart.options.scales.y.max = zoomRange?.max ?? undefined;

  activeChart.update();
}

/* ======================================================
   BUCKETS + COLORS
====================================================== */

function getBucketLabel(s) {
  if (currentView === "carrierGroup") return s.group || "Other";
  if (currentView === "carrier") return s.carrier || "Unknown Carrier";
  return "Policy";
}

function getBucketColor(bucket) {
  if (currentView === "availability") return AVAILABLE_COLOR;
  return getStableColor(bucket);
}

function getStableColor(str) {
  let hash = 0;
  const s = String(str);
  for (let i = 0; i < s.length; i++) hash = s.charCodeAt(i) + ((hash << 5) - hash);
  const idx = Math.abs(hash) % AM_PALETTE.length;
  return AM_PALETTE[idx];
}

/* ======================================================
   HELPERS
====================================================== */

function parseMoney(value) {
  if (value === undefined || value === null || value === "") return 0;
  return parseFloat(String(value).replace(/[$,]/g, "")) || 0;
}

function parseYear(dateStr) {
  if (!dateStr) return null;
  const s = String(dateStr);

  // Matches 15-Apr-79 or 15-Apr-1979
  const match = s.match(/\d{1,2}-[A-Za-z]{3}-(\d{2,4})/);
  if (match) {
    let year = match[1];
    if (year.length === 2) year = parseInt(year, 10) > 30 ? "19" + year : "20" + year;
    return parseInt(year, 10);
  }

  // Fallback: any 4-digit year
  const match2 = s.match(/(\d{4})/);
  if (match2) return parseInt(match2[1], 10);

  return null;
}

function formatMoney(n) {
  const num = Number(n || 0);
  return "$" + num.toLocaleString();
}

function normalizeMoneyStr(v) {
  if (v === undefined || v === null || v === "") return "";
  return String(v).trim();
}

function buildQuotaLayerKey({ attach, quotashrlim, term }) {
  return [normalizeMoneyStr(attach), normalizeMoneyStr(quotashrlim), String(term || "").trim()].join(" | ");
}

function buildQuotaLayerLabel({ attach, quotashrlim, term, policy_no }) {
  const a = normalizeMoneyStr(attach) || "Unknown Attach";
  const q = normalizeMoneyStr(quotashrlim) || "Unknown QS Limit";
  const t = String(term || "").trim();
  const p = String(policy_no || "").trim();

  const parts = [];
  parts.push(`${a} xs ${q}`);
  if (p) parts.push(p);
  if (t) parts.push(`Term ${t}`);
  return parts.join(" • ");
}
