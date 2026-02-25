// coverageChart.js
// Year on X-axis
// Dollars on Y-axis
// Toggle between:
//   1) Carrier
//   2) Carrier Group
//   3) Availability (Available vs Unavailable)

let activeChart = null;
let cachedPolicies = [];
let currentView = 'carrier';



/* ======================================================
   PUBLIC ENTRY
====================================================== */

export function renderCoverageChart({ elementId, policies }) {

    if (!window.Chart) {
        throw new Error("Chart.js must be loaded via CDN before this module.");
    }

    const canvas =
        typeof elementId === 'string'
            ? document.getElementById(elementId)
            : elementId;

    if (!canvas) {
        throw new Error("Canvas element not found.");
    }

    cachedPolicies = policies;

    if (activeChart) {
        activeChart.destroy();
        activeChart = null;
    }

    initializeChart(canvas);
    wireToggleButtons();
}


/* ======================================================
   INITIALIZE
====================================================== */

function initializeChart(canvas) {

    const { years, datasets } = buildDatasets();

    activeChart = new Chart(canvas.getContext("2d"), {
        type: "bar",
        data: {
            labels: years,
            datasets: datasets
        },
        options: {
            responsive: false,
            plugins: {
                legend: {
                    position: "bottom"
                }
            },
            scales: {
                x: {
                    stacked: true,
                    title: { display: true, text: "Year" }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    title: { display: true, text: "Coverage ($)" }
                }
            }
        }
    });
}


/* ======================================================
   DATASET SWITCHING
====================================================== */

function buildDatasets() {

    if (currentView === "availability") {
        return buildAvailabilityView();
    }

    if (currentView === "group") {
        return buildGroupedView(true);
    }

    return buildGroupedView(false);
}

function updateChart() {
    const { years, datasets } = buildDatasets();

    activeChart.data.labels = years;
    activeChart.data.datasets = datasets;
    activeChart.update();
}


/* ======================================================
   GROUPED VIEW
   If useGroup = false → group by Carrier
   If useGroup = true  → group by Carrier Group
====================================================== */

function buildGroupedView(useGroup) {

    const yearSet = new Set();
    const bucketMap = {};

    cachedPolicies.forEach(policy => {

        const year = parseYear(policy.Incept);
        if (!year) return;

        yearSet.add(year);

        const carrier = policy.Carrier || "Unknown Carrier";
        const bucket = useGroup
            ? (carrierGroupMap[carrier] || "Other")
            : carrier;

        const limit = parseLimit(policy.perocclim);

        if (!bucketMap[bucket]) {
            bucketMap[bucket] = {};
        }

        if (!bucketMap[bucket][year]) {
            bucketMap[bucket][year] = 0;
        }

        bucketMap[bucket][year] += limit;
    });

    const years = Array.from(yearSet).sort((a, b) => a - b);
    const buckets = Object.keys(bucketMap).sort();

    const datasets = buckets.map(bucket => ({
        label: bucket,
        data: years.map(y => bucketMap[bucket][y] || 0),
        stack: "Stack",
        backgroundColor: getStableColor(bucket)
    }));

    return { years, datasets };
}


/* ======================================================
   AVAILABILITY VIEW (Corrected + Safe)
====================================================== */

function buildAvailabilityView() {

    const yearSet = new Set();
    const yearMap = {};

    cachedPolicies.forEach(policy => {

        const year = parseYear(policy.Incept);
        if (!year) return;

        yearSet.add(year);

        if (!yearMap[year]) {
            yearMap[year] = {
                available: 0,
                unavailable: 0
            };
        }

        const limit = parseLimit(policy.perocclim);

        // Explicitly check both fields without || short-circuit issues
        let consumedRaw = 0;

        if (policy.Consume !== undefined && policy.Consume !== null && policy.Consume !== "") {
            consumedRaw = policy.Consume;
        } else if (policy.TotCost !== undefined && policy.TotCost !== null && policy.TotCost !== "") {
            consumedRaw = policy.TotCost;
        }

        const consumed = parseLimit(consumedRaw);

        // Guard against bad data
        const unavailable = Math.max(0, Math.min(limit, consumed));
        const available = Math.max(0, limit - unavailable);

        yearMap[year].available += available;
        yearMap[year].unavailable += unavailable;
    });

    const years = Array.from(yearSet).sort((a, b) => a - b);

    return {
        years,
        datasets: [
            {
                label: "Available Coverage",
                data: years.map(y => yearMap[y]?.available ?? 0),
                backgroundColor: "#2ca02c",
                stack: "Stack"
            },
            {
                label: "Unavailable Coverage",
                data: years.map(y => yearMap[y]?.unavailable ?? 0),
                backgroundColor: "#d62728",
                stack: "Stack"
            }
        ]
    };
}


/* ======================================================
   BUTTON WIRING
====================================================== */

function wireToggleButtons() {

    const viewSelect = document.getElementById("viewSelect");
    if (viewSelect) {
        viewSelect.addEventListener("change", (e) => {
            currentView = e.target.value;
            updateChart();
        });
    }
}


/* ======================================================
   HELPERS
====================================================== */

function parseLimit(value) {
    if (!value) return 0;
    return parseFloat(String(value).replace(/[$,]/g, "")) || 0;
}

function parseYear(dateStr) {
    if (!dateStr) return null;

    const match = dateStr.match(/\d{2}-[A-Za-z]{3}-(\d{2,4})/);
    if (match) {
        let year = match[1];
        if (year.length === 2) {
            year = parseInt(year) > 30 ? "19" + year : "20" + year;
        }
        return parseInt(year);
    }

    const match2 = dateStr.match(/(\d{4})/);
    if (match2) return parseInt(match2[1]);

    return null;
}
const carrierGroupMap = {

    /* =========================
       AIG FAMILY
    ========================= */
    "American Home Assurance Company": "AIG",
    "Insurance Company of the State of Pennsylvania": "AIG",
    "Lexington Insurance Company": "AIG",
    "National Union Fire Insurance Company of Pittsburgh, Pa.": "AIG",
    "Granite State Insurance Company": "AIG",
    "Commerce and Industry Insurance Company": "AIG",
    "New Hampshire Insurance Company": "AIG",
    "American International Underwriters": "AIG",

    /* =========================
       CHUBB
    ========================= */
    "Chubb Custom Market Incorporated": "Chubb",
    "Federal Insurance Company": "Chubb",
    "Pacific Indemnity Company": "Chubb",

    /* =========================
       ZURICH
    ========================= */
    "Zurich Insurance Company": "Zurich",
    "American Guarantee and Liability Insurance Company": "Zurich",
    "Maryland Casualty Company": "Zurich",
    "Midland Insurance Company": "Zurich",

    /* =========================
       BERKSHIRE HATHAWAY
    ========================= */
    "Admiral Insurance Company": "Berkshire Hathaway",
    "National Indemnity Company": "Berkshire Hathaway",
    "Columbia Casualty Company": "Berkshire Hathaway",
    "Continental Casualty Company": "Berkshire Hathaway",

    /* =========================
       CNA
    ========================= */
    "CNA Insurance Company": "CNA",
    "Continental Casualty Company": "CNA",
    "Transportation Insurance Company": "CNA",

    /* =========================
       LIBERTY MUTUAL
    ========================= */
    "Liberty Mutual Insurance Company": "Liberty Mutual",
    "Employers Insurance of Wausau": "Liberty Mutual",
    "First State Insurance Company": "Liberty Mutual",

    /* =========================
       TRAVELERS
    ========================= */
    "Travelers Insurance Company": "Travelers",
    "Aetna Cas. and Surety Co.": "Travelers",
    "Aetna Casualty and Surety Company": "Travelers",

    /* =========================
       XL / AXA XL
    ========================= */
    "International Surplus Lines Insurance Company": "AXA XL",
    "XL Insurance Company": "AXA XL",
    "Northbrook Excess & Surplus Insurance Company": "AXA XL",

    /* =========================
       ALLIANZ
    ========================= */
    "Allianz Underwriters Insurance Company": "Allianz",
    "Fireman's Fund Insurance Company": "Allianz",
    "National Surety Corporation (Fireman's Fund)": "Allianz",

    /* =========================
       OLD REPUBLIC
    ========================= */
    "Old Republic Insurance Company": "Old Republic",
    "Bituminous Casualty Corporation": "Old Republic",

    /* =========================
       HARTFORD
    ========================= */
    "Hartford Accident & Indemnity Company": "Hartford",
    "Hartford Insurance Company": "Hartford",

    /* =========================
       LLOYD'S / LONDON
    ========================= */
    "London Market": "Lloyd's",
    "Lloyd's of London": "Lloyd's",
    "Certain Underwriters at Lloyd's": "Lloyd's",

    /* =========================
       CRUM & FORSTER
    ========================= */
    "California Union Insurance Company": "Crum & Forster",
    "Crum & Forster Insurance Company": "Crum & Forster",

    /* =========================
       EVANSTON / MARKEL
    ========================= */
    "Evanston Insurance Company": "Markel",
    "Markel Insurance Company": "Markel",

    /* =========================
       RELIANCE (Historic)
    ========================= */
    "Reliance Insurance Company of Illinois": "Reliance (Historic)",
    "Reliance Insurance Company": "Reliance (Historic)",

    /* =========================
       ARGONAUT
    ========================= */
    "Argonaut Insurance Company": "Argonaut",

    /* =========================
       EMPLOYERS MUTUAL
    ========================= */
    "Employers Mutual Casualty Company": "EMC",

    /* =========================
       HUDSON
    ========================= */
    "Hudson Insurance Company": "OdysseyRe",

    /* =========================
       HARBOR
    ========================= */
    "Harbor Insurance Company": "Continental / CNA",

    /* =========================
       SIGNAL / STONEWALL
    ========================= */
    "Signal Insurance Company": "Reliance (Historic)",
    "Stonewall Insurance Company": "Reliance (Historic)",

    /* =========================
       MISSION (DEFUNCT)
    ========================= */
    "Mission Insurance Company": "Mission (Defunct)",

    /* =========================
       IDEAL MUTUAL
    ========================= */
    "Ideal Mutual Insurance Company": "Ideal Mutual (Defunct)",

    /* =========================
       GLACIER
    ========================= */
    "Glacier General Assurance Company": "Glacier (Historic)",

    /* =========================
       SOUTHERN AMERICAN
    ========================= */
    "Southern American Insurance Company": "Southern American",

    /* =========================
       TUREGUM
    ========================= */
    "Turegum Insurance Company": "Turegum",

    /* =========================
       DEFAULT
    ========================= */
};


/* ======================================================
   STABLE COLOR BY STRING HASH
   Ensures same carrier/group always same color
====================================================== */

function getStableColor(str) {
    const palette = [
        "#1f77b4","#ff7f0e","#2ca02c","#d62728",
        "#9467bd","#8c564b","#e377c2","#7f7f7f",
        "#bcbd22","#17becf"
    ];

    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = str.charCodeAt(i) + ((hash << 5) - hash);
    }

    const index = Math.abs(hash) % palette.length;
    return palette[index];
}
export function setView(view) {
    currentView = view;
    updateChart();
}