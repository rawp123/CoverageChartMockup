export function money(v) {
  return `$${Number(v || 0).toLocaleString()}`;
}

export function compactMoney(v) {
  const n = Number(v);
  if (!Number.isFinite(n)) return "$0";
  const abs = Math.abs(n);
  const sign = n < 0 ? "-" : "";
  if (abs >= 1e9) return `$${sign}${(abs / 1e9).toFixed(abs >= 1e10 ? 0 : 1).replace(/\.0$/, "")}B`;
  if (abs >= 1e6) return `$${sign}${(abs / 1e6).toFixed(abs >= 1e7 ? 0 : 1).replace(/\.0$/, "")}M`;
  if (abs >= 1e3) return `$${sign}${(abs / 1e3).toFixed(abs >= 1e4 ? 0 : 1).replace(/\.0$/, "")}K`;
  return `$${sign}${Math.round(abs).toLocaleString()}`;
}

export function shortLabel(text, max = 20) {
  const s = String(text || "").trim();
  if (s.length <= max) return s;
  return `${s.slice(0, max - 3)}...`;
}

export function formatDate(dateStr, { month = "2-digit" } = {}) {
  if (!dateStr) return "";
  const d = new Date(`${dateStr}T00:00:00Z`);
  if (Number.isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString("en-US", {
    year: "numeric",
    month,
    day: "2-digit",
    timeZone: "UTC"
  });
}

export function toDateStamp(d = new Date()) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}${m}${day}`;
}

export function sanitizeFilePart(value, fallback = "report") {
  const clean = String(value || "")
    .trim()
    .replace(/[^a-z0-9._-]+/gi, "_")
    .replace(/^_+|_+$/g, "");
  return clean || fallback;
}
