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
  filters: {
    startYear: null,
    endYear: null,
    zoomMin: null,
    zoomMax: null,
    insuranceProgram: "",
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

const money = (v) => `$${Number(v || 0).toLocaleString()}`;

const normKey = (k) =>
  String(k ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const clamp = (v, min, max) => Math.min(max, Math.max(min, v));

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

function colorFromString(str) {
  const s = String(str ?? "");
  let hash = 0;
  for (let i = 0; i < s.length; i++) hash = (hash * 31 + s.charCodeAt(i)) >>> 0;
  const hue = hash % 360;
  return `hsl(${hue}, 70%, 55%)`;
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

    const y = yearOf(start);
    if (Number.isFinite(y)) {
      minYear = Math.min(minYear, y);
      maxYear = Math.max(maxYear, y);
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
          getBy(r, "policy_no", "PolicyNo", "Policy Number", "PolicyNumber")
        ).trim();
        const insuranceProgramId = String(
          getBy(r, "InsuranceProgramID", "Insurance Program ID")
        ).trim();
        const insuranceProgram =
          String(getBy(r, "InsuranceProgram", "Program", "ProgramName")).trim() ||
          (insuranceProgramId ? insuranceProgramNameById[insuranceProgramId] || "" : "");

        const cRow = carrierId ? carrierRowById[carrierId] : null;
        const availability = classifyAvailability(r, cRow);

        policyInfoById[pid] = {
          policy_no: policyNo,
          carrier: carrierName || "(unknown carrier)",
          carrierGroup: carrierGroupName || "(unknown group)",
          insuranceProgram: insuranceProgram || "(unknown program)",
          availability,
        };
  }

  const slices = [];
  for (const r of limitsRows) {
    const pid = String(getBy(r, "PolicyID", "Policy Id", "ID")).trim();
    if (!pid) continue;

    const dates = policyDateMap[pid];
    if (!dates || !dates.start || !dates.end) continue;

    const yr = yearOf(dates.start);
    if (!Number.isFinite(yr)) continue;

    const x = useYearAxis ? String(yr) : `${dates.start} to ${dates.end}`;

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

    const info = policyInfoById[pid] || {
      policy_no: "",
      carrier: "(unknown carrier)",
      carrierGroup: "(unknown group)",
      availability: "Available"
    };

    slices.push({
      x,
      year: yr,
      attach,
      sliceLimit,
      PolicyID: pid,
      policy_no: info.policy_no,
      carrier: info.carrier,
      carrierGroup: info.carrierGroup,
      insuranceProgram: info.insuranceProgram,
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

function buildQuotaKeySet(slices) {
  const byXA = new Map(); // `${x}||${attach}` -> Set(PolicyID)
  for (const s of slices) {
    const k = `${s.x}||${s.attach}`;
    if (!byXA.has(k)) byXA.set(k, new Set());
    byXA.get(k).add(String(s.PolicyID));
  }
  const quotaKeySet = new Set();
  for (const [k, set] of byXA.entries()) {
    if (set.size > 1) quotaKeySet.add(k);
  }
  return quotaKeySet;
}

function applyFiltersToCache() {
  const { allSlices, allXLabels, useYearAxis, filters } = _cache;
  const startYear = Number.isFinite(filters.startYear) ? filters.startYear : null;
  const endYear = Number.isFinite(filters.endYear) ? filters.endYear : null;
  const selectedProgram = String(filters.insuranceProgram || "").trim();

  let filteredSlices = allSlices.slice();
  if (useYearAxis && (startYear !== null || endYear !== null)) {
    filteredSlices = filteredSlices.filter((s) => {
      if (!Number.isFinite(s.year)) return false;
      if (startYear !== null && s.year < startYear) return false;
      if (endYear !== null && s.year > endYear) return false;
      return true;
    });
  }

  if (selectedProgram) {
    filteredSlices = filteredSlices.filter(
      (s) => String(s?.insuranceProgram || "").trim() === selectedProgram
    );
  }

  let filteredXLabels = allXLabels.slice();
  if (useYearAxis && (startYear !== null || endYear !== null)) {
    filteredXLabels = filteredXLabels.filter((lbl) => {
      const y = Number(lbl);
      if (!Number.isFinite(y)) return false;
      if (startYear !== null && y < startYear) return false;
      if (endYear !== null && y > endYear) return false;
      return true;
    });
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
  if (!viewport || _cache._wheelBound) return;

  viewport.addEventListener(
    "wheel",
    (e) => {
      // Ctrl/Cmd + wheel: horizontal zoom around cursor.
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        const factor = Math.exp(-e.deltaY * 0.0015);
        _cache.xZoom = clamp((_cache.xZoom || 1) * factor, 0.45, 4);
        syncChartViewportWidth({ anchorClientX: e.clientX });
        return;
      }

      // Shift + wheel: smoother horizontal scrolling in dense views.
      if (e.shiftKey && Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
        e.preventDefault();
        viewport.scrollLeft += e.deltaY;
      }
    },
    { passive: false }
  );

  _cache._wheelBound = true;
}

function rebuildChart() {
  if (!chart || !_cache.options) return;

  const { barThickness, categorySpacing } = _cache.options;

  chart.data.labels = _cache.xLabels;
  chart.data.datasets = buildDatasetsForView({
    slices: _cache.slices,
    xLabels: _cache.xLabels,
    view: currentView,
    barThickness,
    categorySpacing,
    quotaKeySet: _cache.quotaKeySet
  });

  const y = chart.options?.scales?.y;
  if (y) {
    y.min = Number.isFinite(_cache.filters.zoomMin) ? _cache.filters.zoomMin : undefined;
    y.max = Number.isFinite(_cache.filters.zoomMax) ? _cache.filters.zoomMax : undefined;
  }

  chart.update();
  syncChartViewportWidth();
}

/* ================================
   View aggregation -> datasets
================================ */

function buildDatasetsForView({ slices, xLabels, view, barThickness, categorySpacing, quotaKeySet }) {
  const qaKey = (s) => `${s.x}||${s.attach}`;
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

  const layerMap = new Map(); // `${group}||${x}||${attach}`
  const groups = new Set();

  for (const s of slices) {
    const group = keyOf(s);
    groups.add(group);

    const k = `${group}||${s.x}||${s.attach}`;
    if (!layerMap.has(k)) {
      layerMap.set(k, {
        group,
        x: s.x,
        attach: s.attach,
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
      sliceLimit: s.sliceLimit
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

  return groupList.map((group) => {
    const points = [];

    for (const e of layerMap.values()) {
      if (e.group !== group) continue;

      const top = e.attach + e.sumLimit;

      e.participants.sort((a, b) => {
        if (b.sliceLimit !== a.sliceLimit) return b.sliceLimit - a.sliceLimit;
        return String(a.carrier || "").localeCompare(String(b.carrier || ""));
      });

      const isQuotaShare = quotaKeySet && quotaKeySet.has(`${e.x}||${e.attach}`);

      points.push({
        x: e.x,
        y: [e.attach, top],
        attach: e.attach,
        top,
        sumLimit: e.sumLimit,
        participants: e.participants,
        group,
        isQuotaShare,
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

      backgroundColor: datasetInFocus ? bg : mutedFill,
      borderColor: datasetInFocus ? "rgba(11, 17, 27, 0.62)" : mutedBorder,
      borderWidth: 1,
      borderRadius: 2,
      borderSkipped: false
    };
  });
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

export function getYearBounds() {
  const years = _cache.allXLabels
    .map((lbl) => Number(lbl))
    .filter((y) => Number.isFinite(y))
    .sort((a, b) => a - b);

  return {
    minYear: years.length ? years[0] : null,
    maxYear: years.length ? years[years.length - 1] : null,
    startYear: Number.isFinite(_cache.filters.startYear) ? _cache.filters.startYear : null,
    endYear: Number.isFinite(_cache.filters.endYear) ? _cache.filters.endYear : null
  };
}

export function getFilterOptions() {
  const carrierSet = new Set();
  const carrierGroupSet = new Set();
  const insuranceProgramSet = new Set();
  for (const s of _cache.allSlices || []) {
    const c = String(s?.carrier || "").trim();
    const g = String(s?.carrierGroup || "").trim();
    const p = String(s?.insuranceProgram || "").trim();
    if (c && c !== "(unknown carrier)") carrierSet.add(c);
    if (g && g !== "(unknown group)") carrierGroupSet.add(g);
    if (p && p !== "(unknown program)") insuranceProgramSet.add(p);
  }

  return {
    insurancePrograms: Array.from(insuranceProgramSet).sort((a, b) => a.localeCompare(b)),
    carriers: Array.from(carrierSet).sort((a, b) => a.localeCompare(b)),
    carrierGroups: Array.from(carrierGroupSet).sort((a, b) => a.localeCompare(b)),
    selectedInsuranceProgram: String(_cache.filters?.insuranceProgram || "").trim(),
    selectedCarriers: normalizeStringList(_cache.filters?.carriers),
    selectedCarrierGroups: normalizeStringList(_cache.filters?.carrierGroups)
  };
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
  applyFiltersToCache();
  rebuildChart();
}

export function resetYearRange() {
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

  const [limitsRows, datesRows, policyRows, carrierRows, carrierGroupRows, insuranceProgramRows] =
    await Promise.all([
    fetchCSV(csvUrl),
    fetchCSV(policyDatesUrl),
    fetchCSV(policyUrl),
    fetchCSV(carrierUrl),
    fetchCSV(carrierGroupUrl),
    fetchInsuranceProgramRows(insuranceProgramUrl)
  ]);

  const built = buildSlices({
    limitsRows,
    datesRows,
    policyRows,
    carrierRows,
    carrierGroupRows,
    insuranceProgramRows,
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
    filters: {
      startYear: null,
      endYear: null,
      zoomMin: null,
      zoomMax: null,
      insuranceProgram: "",
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

  if (chart) chart.destroy();

  chart = new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels: _cache.xLabels,
      datasets
    },
    plugins: [outlineBarsPlugin, quotaShareGuidesPlugin],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      layout: {
        // Keep bars/outline away from the extreme plot edges without using x.offset,
        // which can break floating + non-grouped bar geometry in this chart.
        padding: { left: 14, right: 24 }
      },

      plugins: {
        outlineBars: {
          lineWidth: outlineWidth,
          color: "rgba(15,23,42,0.42)"
        },
        legend: {
          display: true,
          position: "right",
          align: "start",
          labels: {
            color: "rgba(255,255,255,0.95)",
            boxWidth: 9,
            boxHeight: 9,
            padding: 10,
            font: { size: 11 }
          }
        },
        tooltip: {
          displayColors: false,
          callbacks: {
            title: (items) => {
              const raw = items?.[0]?.raw || {};
              const yr = String(raw.x ?? "");

              if (raw.isQuotaShare) return `${yr} — Quota share`;

              const g = String(raw.group ?? "").trim();
              return g ? `${yr} — ${g}` : yr;
            },
            label: (ctx) => {
              const r = ctx.raw || {};
              const attach = r.attach ?? 0;
              const top = r.top ?? 0;
              const lim = Math.max(0, top - attach);

              const lines = [];
              lines.push(`Attach: ${money(attach)}`);
              lines.push(`Limit: ${money(lim)}`);
              lines.push(`Top: ${money(top)}`);

              const parts = Array.isArray(r.participants) ? r.participants : [];
              const isUnavailableGroup = String(r.group || "").toLowerCase() === "unavailable";
              const shouldShowParts = parts.length > 1 || String(r.group || "") === "Quota share";

              if (isUnavailableGroup && parts.length) {
                const uniqueCarriers = [...new Set(parts.map((p) => p.carrier || "(unknown carrier)"))];
                const carrierLine =
                  uniqueCarriers.length === 1
                    ? `Carrier: ${uniqueCarriers[0]}`
                    : `Carriers: ${uniqueCarriers.join(", ")}`;
                lines.push(carrierLine);
              }

              if (r.isQuotaShare && parts.length) {
                const uniqueQuotaCarriers = [...new Set(parts.map((p) => p.carrier || "(unknown carrier)"))];
                const quotaCarrierLine =
                  uniqueQuotaCarriers.length === 1
                    ? `Carrier: ${uniqueQuotaCarriers[0]}`
                    : `Carriers: ${uniqueQuotaCarriers.join(", ")}`;
                lines.push(quotaCarrierLine);
              }

              if (shouldShowParts && parts.length) {
                lines.push(`Quota share participants (${parts.length}):`);

                const show = parts.slice(0, tooltipMaxParticipants);
                for (const p of show) {
                  const carrier = p.carrier || "(unknown carrier)";
                  const polno = p.policy_no ? ` (${p.policy_no})` : "";
                  lines.push(`• ${carrier}${polno}: ${money(p.sliceLimit)}`);
                }

                if (parts.length > tooltipMaxParticipants) {
                  lines.push(`… +${parts.length - tooltipMaxParticipants} more`);
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
            color: "rgba(255,255,255,0.06)",
            lineWidth: 1
          },
          ticks: {
            color: "rgba(255,255,255,0.92)",
            autoSkip: false,
            maxRotation: 0,
            minRotation: 0
          }
        },
        y: {
          beginAtZero: true,
          grid: {
            display: true,
            color: "rgba(255,255,255,0.08)",
            lineWidth: 1
          },
          ticks: {
            color: "rgba(255,255,255,0.95)",
            callback: (v) => "$" + Number(v).toLocaleString()
          },
          title: {
            display: true,
            text: "Coverage Limits",
            color: "rgba(255,255,255,0.95)"
          }
        }
      }
    }
  });

  bindViewportInteractions();
  // Defer width sync so the viewport has final layout dimensions.
  requestAnimationFrame(() => {
    if (!chart) return;
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
