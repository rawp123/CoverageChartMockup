// CoverageChart.js (polished + year floor fix + quota share hover details)

let activeChart = null;
let cachedPolicies = [];
let currentView = "carrier";

let isLogScale = false;
let zoomRange = null; // { min: number|null, max: number|null }
let zoomSmall = false; // kept for UI/state safety (used by toggleLogScale/syncControlsToState)

const ZOOM_SMALL_MAX = 50_000_000;
let controlsWired = false;

/* =========================
   YEAR RANGE (DISPLAY)
   - Show only MIN_YEAR and forward (optionally cap with MAX_YEAR)
========================= */
const MIN_YEAR = 1974;   // display 1974 and forward
const MAX_YEAR = null;   // set to a number (e.g., 1985) to cap the chart

/* =========================
   QUOTA SHARE BEHAVIOR
   - We treat each row as a "slice" (participant) of a quota share layer.
   - Hover tooltip will show *all participants* for each quota-share layer in that year.
========================= */
const ENABLE_QUOTA_SHARE_CALLOUT = true;

/* =========================
   A&M-inspired palette
   (maroon + neutrals + muted accents)
========================= */
const AM_PALETTE = [
  "#500000", // A&M maroon
  "#6B1F1F",
  "#7A2E2E",
  "#8A3B3B",
  "#1F2937", // slate
  "#374151",
  "#4B5563",
  "#6B7280",
  "#0F766E", // muted teal
  "#1D4ED8", // muted blue
  "#7C3AED", // muted purple
  "#B45309", // muted amber
  "#A21CAF", // muted magenta
  "#0B7285", // muted cyan
];

const AVAILABLE_COLOR = "#2F855A";   // muted green
const UNAVAILABLE_COLOR = "#9B2C2C"; // muted maroon-red

function formatMoney(n) {
  const num = Number(n || 0);
  return "$" + num.toLocaleString();
}

function formatMoneyTick(v) {
  const n = typeof v === "object" && v !== null && "value" in v ? v.value : v;
  return formatMoney(n);
}

/* ======================================================
   PUBLIC ENTRY
====================================================== */

export function renderCoverageChart({ elementId, policies }) {
  if (!window.Chart) {
    throw new Error("Chart.js must be loaded via CDN before this module.");
  }

  const canvas =
    typeof elementId === "string" ? document.getElementById(elementId) : elementId;

  if (!canvas) throw new Error("Canvas element not found.");

  cachedPolicies = Array.isArray(policies) ? policies : [];

  if (activeChart) {
    activeChart.destroy();
    activeChart = null;
  }

  initializeChart(canvas);
  wireControlsOnce();
  syncControlsToState();
}

export function setView(view) {
  currentView = view;
  syncControlsToState();
  updateChart();
}

export function toggleLogScale() {
  isLogScale = !isLogScale;

  // If log scale is on, disable zoom cap (cap doesn't make sense on log)
  if (isLogScale) zoomSmall = false;

  syncControlsToState();
  updateChart();
}

export function setZoomRange(min, max) {
  const parsedMin = min !== null ? parseFloat(min) : null;
  const parsedMax = max !== null ? parseFloat(max) : null;

  if (parsedMin !== null && parsedMax !== null && parsedMin >= parsedMax) {
    console.warn("Invalid zoom range: min must be less than max.");
    return;
  }

  if (isLogScale) {
    console.warn("Cannot set zoom range while log scale is active.");
    return;
  }

  zoomRange = { min: parsedMin, max: parsedMax };
  isLogScale = false; // Force linear scale
  updateChart();
}

export function resetZoomRange() {
  zoomRange = null;
  updateChart();
}

/* ======================================================
   INITIALIZE
====================================================== */

function initializeChart(canvas) {
  const { years, datasets, quotaDetailsByYear } = buildDatasets();

  activeChart = new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: { labels: years, datasets },
    options: buildChartOptions()
  });

  // attach quota metadata for tooltip callouts
  activeChart._quotaDetailsByYear = quotaDetailsByYear || {};

  applyScaleOptions();
  activeChart.update();
}

function buildChartOptions() {
  return {
    responsive: false,
    animation: { duration: 200 },
    layout: { padding: { top: 10, right: 18, bottom: 10, left: 10 } },
    plugins: {
      legend: {
        position: "bottom",
        labels: {
          boxWidth: 12,
          boxHeight: 12,
          usePointStyle: true,
          pointStyle: "rectRounded",
          padding: 14,
          font: { size: 11 }
        }
      },
      tooltip: {
        callbacks: {
          label: (ctx) => {
            const label = ctx.dataset.label || "";
            const val = ctx.parsed.y || 0;
            return `${label}: ${formatMoney(val)}`;
          },
          title: (items) => `Year: ${items?.[0]?.label ?? ""}`,
          afterBody: (items) => {
            if (!ENABLE_QUOTA_SHARE_CALLOUT) return [];
            const chart = items?.[0]?.chart;
            if (!chart || !chart._quotaDetailsByYear) return [];

            const yearStr = items?.[0]?.label;
            const year = parseInt(yearStr, 10);
            if (!Number.isFinite(year)) return [];

            const layers = chart._quotaDetailsByYear?.[year];
            if (!layers) return [];

            const layerKeys = Object.keys(layers);
            if (!layerKeys.length) return [];

            const lines = [];
            lines.push("");
            lines.push("Quota Share Participants (by layer):");

            // Keep this readable; if many layers/participants, show first few
            const MAX_LAYERS = 6;
            const MAX_PARTICIPANTS_PER_LAYER = 10;

            layerKeys.slice(0, MAX_LAYERS).forEach((layerKey) => {
              const layer = layers[layerKey];
              if (!layer) return;

              const header = layer.displayLabel || "Layer";
              lines.push(`• ${header}`);

              const participants = Array.isArray(layer.participants) ? layer.participants : [];
              participants.slice(0, MAX_PARTICIPANTS_PER_LAYER).forEach((p) => {
                const who = p.bucketLabel || p.carrier || "Unknown";
                const amt = formatMoney(p.sliceLimit || 0);
                lines.push(`   - ${who}: ${amt}`);
              });

              if (participants.length > MAX_PARTICIPANTS_PER_LAYER) {
                lines.push(`   - (+${participants.length - MAX_PARTICIPANTS_PER_LAYER} more)`);
              }
            });

            if (layerKeys.length > MAX_LAYERS) {
              lines.push(`• (+${layerKeys.length - MAX_LAYERS} more layers)`);
            }

            return lines;
          }
        }
      }
    },
    scales: {
      x: {
        stacked: true,
        grid: { display: false },
        ticks: { font: { size: 11 } },
        title: { display: true, text: "Year", font: { size: 12 } }
      },
      y: {
        stacked: true,
        beginAtZero: true,
        grid: { color: "rgba(31,41,55,0.10)" },
        ticks: { callback: formatMoneyTick, font: { size: 11 } },
        title: { display: true, text: "Coverage ($)", font: { size: 12 } }
      }
    }
  };
}

/* ======================================================
   DATASET SWITCHING
====================================================== */

function buildDatasets() {
  if (currentView === "availability") return buildAvailabilityView();
  if (currentView === "group") return buildGroupedView(true);
  return buildGroupedView(false);
}

/**
 * Filters an array of years to the display window.
 * - Always enforces MIN_YEAR floor
 * - Optionally enforces MAX_YEAR cap if set (number)
 */
function filterYears(years) {
  return years.filter((y) => y >= MIN_YEAR && (MAX_YEAR === null || y <= MAX_YEAR));
}

function updateChart() {
  if (!activeChart) return;

  const { years, datasets, quotaDetailsByYear } = buildDatasets();
  activeChart.data.labels = years;
  activeChart.data.datasets = datasets;

  // refresh quota metadata for tooltip callouts
  activeChart._quotaDetailsByYear = quotaDetailsByYear || {};

  applyScaleOptions();
  activeChart.update();
}

/* ======================================================
   GROUPED VIEW (Carrier / CarrierGroup)
====================================================== */

function buildGroupedView(useGroup) {
  const yearSet = new Set();
  const bucketMap = {};
  const quotaDetailsByYear = {}; // year -> layerKey -> { displayLabel, participants: [...] }

  cachedPolicies.forEach((policy) => {
    const year = parseYear(policy?.Incept);
    if (!year) return;

    yearSet.add(year);

    const carrier = policy?.Carrier || "Unknown Carrier";
    const group = policy?.CarrierGroup || "Other";
    const bucket = useGroup ? group : carrier;

    // Each row is a "slice" (participant) as you described.
    const sliceLimit = parseLimit(policy?.perocclim);

    bucketMap[bucket] ??= {};
    bucketMap[bucket][year] ??= 0;
    bucketMap[bucket][year] += sliceLimit;

    // Quota share callout details (only if flagged as QS)
    if (ENABLE_QUOTA_SHARE_CALLOUT && Number(policy?.bquotashr) === 1) {
      const layerKey = buildQuotaLayerKey(policy);
      const displayLabel = buildQuotaLayerLabel(policy);

      quotaDetailsByYear[year] ??= {};
      quotaDetailsByYear[year][layerKey] ??= { displayLabel, participants: [] };

      quotaDetailsByYear[year][layerKey].participants.push({
        carrier,
        bucketLabel: bucket,
        sliceLimit
      });
    }
  });

  // Filter the year list first, then build datasets off that list
  const years = filterYears(Array.from(yearSet).sort((a, b) => a - b));
  const buckets = Object.keys(bucketMap).sort();

  const datasets = buckets.map((bucket, idx) => ({
    label: bucket,
    data: years.map((y) => bucketMap[bucket][y] || 0),
    stack: "Stack",
    backgroundColor: getStableColor(bucket, idx),
    borderWidth: 0,
    borderRadius: 2
  }));

  // Sort participants within each layer (largest slice first) for nicer tooltips
  Object.keys(quotaDetailsByYear).forEach((y) => {
    Object.keys(quotaDetailsByYear[y]).forEach((k) => {
      quotaDetailsByYear[y][k].participants.sort((a, b) => (b.sliceLimit || 0) - (a.sliceLimit || 0));
    });
  });

  return { years, datasets, quotaDetailsByYear };
}

/* ======================================================
   AVAILABILITY VIEW (Available vs Unavailable)
====================================================== */

function buildAvailabilityView() {
  const yearSet = new Set();
  const yearMap = {};
  const quotaDetailsByYear = {}; // year -> layerKey -> { displayLabel, participants: [...] }

  cachedPolicies.forEach((policy) => {
    const year = parseYear(policy?.Incept);
    if (!year) return;

    yearSet.add(year);
    yearMap[year] ??= { available: 0, unavailable: 0 };

    const sliceLimit = parseLimit(policy?.perocclim);

    let consumedRaw = 0;
    if (policy?.Consume !== undefined && policy?.Consume !== null && policy?.Consume !== "") {
      consumedRaw = policy.Consume;
    } else if (policy?.TotCost !== undefined && policy?.TotCost !== null && policy?.TotCost !== "") {
      consumedRaw = policy.TotCost;
    }

    const consumed = parseLimit(consumedRaw);

    // IMPORTANT: In your data model, Consume/TotCost appear at the row level.
    // That means each slice carries its own consumed amount already (as provided).
    const unavailable = Math.max(0, Math.min(sliceLimit, consumed));
    const available = Math.max(0, sliceLimit - unavailable);

    yearMap[year].available += available;
    yearMap[year].unavailable += unavailable;

    // Quota share callout details
    if (ENABLE_QUOTA_SHARE_CALLOUT && Number(policy?.bquotashr) === 1) {
      const layerKey = buildQuotaLayerKey(policy);
      const displayLabel = buildQuotaLayerLabel(policy);

      quotaDetailsByYear[year] ??= {};
      quotaDetailsByYear[year][layerKey] ??= { displayLabel, participants: [] };

      quotaDetailsByYear[year][layerKey].participants.push({
        carrier: policy?.Carrier || "Unknown Carrier",
        bucketLabel: "Quota Share Slice",
        sliceLimit,
        consumed,
        available,
        unavailable
      });
    }
  });

  const years = filterYears(Array.from(yearSet).sort((a, b) => a - b));

  // Sort participants within each layer (largest slice first)
  Object.keys(quotaDetailsByYear).forEach((y) => {
    Object.keys(quotaDetailsByYear[y]).forEach((k) => {
      quotaDetailsByYear[y][k].participants.sort((a, b) => (b.sliceLimit || 0) - (a.sliceLimit || 0));
    });
  });

  return {
    years,
    quotaDetailsByYear,
    datasets: [
      {
        label: "Available",
        data: years.map((y) => yearMap[y]?.available ?? 0),
        backgroundColor: AVAILABLE_COLOR,
        stack: "Stack",
        borderRadius: 2
      },
      {
        label: "Unavailable",
        data: years.map((y) => yearMap[y]?.unavailable ?? 0),
        backgroundColor: UNAVAILABLE_COLOR,
        stack: "Stack",
        borderRadius: 2
      }
    ]
  };
}

/* ======================================================
   CONTROLS
====================================================== */

function wireControlsOnce() {
  if (controlsWired) return;
  controlsWired = true;

  const viewSelect = document.getElementById("viewSelect");
  if (viewSelect) {
    viewSelect.addEventListener("change", (e) => {
      currentView = e.target.value;
      syncControlsToState();
      updateChart();
    });
  }

  const scaleBtn = document.getElementById("scaleToggleBtn");
  if (scaleBtn) scaleBtn.addEventListener("click", toggleLogScale);

  const zoomApplyBtn = document.getElementById("zoomApplyBtn");
  if (zoomApplyBtn) {
    zoomApplyBtn.addEventListener("click", () => {
      const min = document.getElementById("zoomMin")?.value;
      const max = document.getElementById("zoomMax")?.value;
      setZoomRange(min, max);
    });
  }

  const zoomResetBtn = document.getElementById("zoomResetBtn");
  if (zoomResetBtn) {
    zoomResetBtn.addEventListener("click", resetZoomRange);
  }
}

function syncControlsToState() {
  const viewSelect = document.getElementById("viewSelect");
  if (viewSelect && viewSelect.value !== currentView) viewSelect.value = currentView;

  const scaleBtn = document.getElementById("scaleToggleBtn");
  if (scaleBtn) scaleBtn.textContent = isLogScale ? "Log Scale: ON" : "Log Scale: OFF";

  const zoomBtn = document.getElementById("zoomSmallBtn");
  if (zoomBtn) zoomBtn.textContent = zoomSmall ? "Zoom Small: ON" : "Zoom Small: OFF";

  // Disable zoom button in log mode (cleaner UX)
  if (zoomBtn) zoomBtn.disabled = isLogScale;
}

/* ======================================================
   SCALE STATE APPLICATION (SAFE)
====================================================== */

function applyScaleOptions() {
  if (!activeChart) return;

  const scales = activeChart.options.scales || (activeChart.options.scales = {});
  const prevY = scales.y || {};

  const baseY = {
    stacked: prevY.stacked ?? true,
    title: prevY.title ?? { display: true, text: "Coverage ($)" },
    grid: prevY.grid ?? { color: "rgba(31,41,55,0.10)" },
    ticks: { ...(prevY.ticks || {}), callback: formatMoneyTick, maxTicksLimit: 8 }
  };

  if (isLogScale) {
    scales.y = {
      ...baseY,
      type: "logarithmic",
      beginAtZero: false,
      min: 1,
      max: undefined
    };
  } else {
    scales.y = {
      ...baseY,
      type: "linear",
      beginAtZero: true,
      min: zoomRange?.min ?? 0,
      max: zoomRange?.max ?? undefined
    };
  }
}

/* ======================================================
   HELPERS
====================================================== */

function parseLimit(value) {
  if (value === undefined || value === null || value === "") return 0;
  return parseFloat(String(value).replace(/[$,]/g, "")) || 0;
}

function parseYear(dateStr) {
  if (!dateStr) return null;

  const s = String(dateStr);

  const match = s.match(/\d{2}-[A-Za-z]{3}-(\d{2,4})/);
  if (match) {
    let year = match[1];
    if (year.length === 2) year = parseInt(year, 10) > 30 ? "19" + year : "20" + year;
    return parseInt(year, 10);
  }

  const match2 = s.match(/(\d{4})/);
  if (match2) return parseInt(match2[1], 10);

  return null;
}

function normalizeMoneyStr(v) {
  if (v === undefined || v === null || v === "") return "";
  return String(v).trim();
}

/**
 * Layer key: group participants that are part of the same quota share "layer"
 * within a policy year. We use fields that appear stable in your data:
 * - attach (attachment point)
 * - quotashrlim (quota share layer limit)
 * - term (end date)
 * - policy_no (often shared across participants)
 */
function buildQuotaLayerKey(policy) {
  const attach = normalizeMoneyStr(policy?.attach);
  const qsl = normalizeMoneyStr(policy?.quotashrlim);
  const term = String(policy?.term || "").trim();
  const pol = String(policy?.policy_no || "").trim();
  // Keep it deterministic and "strict enough" to avoid accidental merges
  return [attach, qsl, term, pol].join(" | ");
}

function buildQuotaLayerLabel(policy) {
  const attach = normalizeMoneyStr(policy?.attach) || "Unknown Attach";
  const qsl = normalizeMoneyStr(policy?.quotashrlim) || "Unknown QS Limit";
  const term = String(policy?.term || "").trim();
  const pol = String(policy?.policy_no || "").trim();

  // Example: "$25,250,000 xs $25,000,000 • CE 343 7270 • Term 09-Oct-76"
  const parts = [];
  parts.push(`${attach} xs ${qsl}`);
  if (pol) parts.push(pol);
  if (term) parts.push(`Term ${term}`);
  return parts.join(" • ");
}

// Stable color: hash into A&M palette so carrier/group colors look intentional
function getStableColor(str, idx = 0) {
  let hash = 0;
  const s = String(str);
  for (let i = 0; i < s.length; i++) hash = s.charCodeAt(i) + ((hash << 5) - hash);
  const base = Math.abs(hash + idx) % AM_PALETTE.length;
  return AM_PALETTE[base];
}
