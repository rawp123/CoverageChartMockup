export function getPreferredTheme(storageKey = "coverageChartTheme") {
  const stored = localStorage.getItem(storageKey);
  if (stored === "light" || stored === "dark") return stored;
  return window.matchMedia("(prefers-color-scheme: light)").matches ? "light" : "dark";
}

export function applyThemeToPage(
  theme,
  {
    themeLabelEl = null,
    themeToggleBtn = null,
    lightLabel = "Light mode",
    darkLabel = "Dark mode"
  } = {}
) {
  document.documentElement.dataset.theme = theme;
  const isLight = theme === "light";
  if (themeToggleBtn) {
    themeToggleBtn.classList.toggle("isLight", isLight);
    themeToggleBtn.setAttribute("aria-pressed", String(isLight));
  }
  if (themeLabelEl) {
    themeLabelEl.textContent = isLight ? lightLabel : darkLabel;
  }
}
