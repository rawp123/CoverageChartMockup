    import {
      renderCoverageChart,
      setView,
      getFilterOptions,
      getYearBounds,
      setYearRange,
      resetYearRange,
      setInsuranceProgramFilter,
      resetInsuranceProgramFilter,
      setPolicyLimitTypeFilter,
      resetPolicyLimitTypeFilter,
      setChartTheme,
      exportChartAsPNG,
      exportFilteredCSV,
      exportReportPDF,
      setEntityFilters,
      resetEntityFilters,
      setZoomRange,
      resetZoomRange,
      getFilteredSlices,
      setCoverageTotalsVisible,
      getYearLabelAnchors,
      setAnnualizedMode,
      getPolicySelectionFromEvent
    } from "/Modules/shared/js/coverage/coverageChartEngine.js";
    import { money, compactMoney } from "../shared/js/core/format.js";
    import { getPreferredTheme, applyThemeToPage } from "../shared/js/core/theme.js";
    import {
      buildCheckboxMenu,
      selectedValuesFromCheckboxMenu,
      clearCheckboxMenu,
      resetCheckboxMenuSearch,
      updateDropdownLabel
    } from "../shared/js/ui/multiSelect.js";

    function showError(err) {
      const box = document.getElementById("errorBox");
      box.style.display = "block";
      box.textContent =
        "ERROR:\\n" + (err?.stack || err?.message || String(err)) +
        "\\n\\nQuick checks:" +
        "\\n- Chart.js loaded (CDN)" +
        "\\n- CSV URLs correct (PolicyLimits, PolicyDates, Policy, Carrier, CarrierGroup)" +
        "\\n- Server is serving /data and /Modules";
    }

    async function init() {
      try {
        const THEME_STORAGE_KEY = "coverageChartTheme";
        const COVERAGE_BADGE_STORAGE_KEY = "coverageChartShowAvailableTotal";
        const ANNUALIZED_STORAGE_KEY = "coverageChartAnnualizedMode";
        const getShowCoverageBadgePref = () => {
          const stored = localStorage.getItem(COVERAGE_BADGE_STORAGE_KEY);
          if (stored === "0") return false;
          if (stored === "1") return true;
          return true;
        };
        const getAnnualizedPref = () => {
          const stored = localStorage.getItem(ANNUALIZED_STORAGE_KEY);
          if (stored === "1") return true;
          if (stored === "0") return false;
          return false;
        };
        const initialTheme = getPreferredTheme(THEME_STORAGE_KEY);
        applyThemeToPage(initialTheme);

        await renderCoverageChart({
          canvasId: "coverageCanvas",
          csvUrl: "/data/OriginalFiles/tblPolicyLimits.csv",
          initialView: "carrier",
          barThickness: "flex",
          categorySpacing: 1.0,
          tooltipMaxParticipants: 50
        });

        const sel = document.getElementById("viewSelect");

        const startYearSelect = document.getElementById("startYearSelect");
        const endYearSelect = document.getElementById("endYearSelect");
        const insuranceProgramSelect = document.getElementById("insuranceProgramSelect");
        const policyLimitTypeSelect = document.getElementById("policyLimitTypeSelect");
        const themeToggleBtn = document.getElementById("themeToggleBtn");
        const themeLabel = document.getElementById("themeLabel");
        const carrierDropdownMenu = document.getElementById("carrierDropdownMenu");
        const carrierGroupDropdownMenu = document.getElementById("carrierGroupDropdownMenu");
        const carrierDropdownLabel = document.getElementById("carrierDropdownLabel");
        const carrierGroupDropdownLabel = document.getElementById("carrierGroupDropdownLabel");
        const activeFilterSummary = document.getElementById("activeFilterSummary");
        const yearTotalsStrip = document.getElementById("yearTotalsStrip");
        const chartViewport = document.querySelector(".chartViewport");
        const chartSurface = document.querySelector(".chartSurface");
        const zoomMinInput = document.getElementById("zoomMinInput");
        const zoomMaxInput = document.getElementById("zoomMaxInput");
        const resetAllBtn = document.getElementById("resetAll");
        const exportMenu = document.getElementById("exportMenu");
        const exportBtn = document.getElementById("exportBtn");
        const exportMenuPanel = document.getElementById("exportMenuPanel");
        const exportPngBtn = document.getElementById("exportPngBtn");
        const exportCsvBtn = document.getElementById("exportCsvBtn");
        const exportPdfBtn = document.getElementById("exportPdfBtn");
        const availableCoverageBadge = document.getElementById("availableCoverageBadge");
        const toggleCoverageBadgeBtn = document.getElementById("toggleCoverageBadgeBtn");
        const annualizeToggleBtn = document.getElementById("annualizeToggleBtn");
        const coverageCanvas = document.getElementById("coverageCanvas");
        let showCoverageBadge = true;
        let annualizedMode = getAnnualizedPref();
        setCoverageTotalsVisible(showCoverageBadge);
        setAnnualizedMode(annualizedMode);
        const syncAnnualizeToggleUI = () => {
          annualizeToggleBtn.textContent = annualizedMode ? "Annualized: On" : "Annualized: Off";
          annualizeToggleBtn.setAttribute("aria-pressed", String(annualizedMode));
        };
        sel.addEventListener("change", (e) => setView(e.target.value));
        syncAnnualizeToggleUI();

        coverageCanvas.addEventListener("click", (evt) => {
          const selected = getPolicySelectionFromEvent(evt);
          if (!selected || !selected.policyId) return;
          localStorage.setItem("coverageChartSelectedPolicy", JSON.stringify(selected));
          const params = new URLSearchParams({ policyId: String(selected.policyId) });
          const policyNumber = String(selected.policyNumber || "").trim();
          const policyLimitType = String(selected.policyLimitType || "").trim();
          const insuranceProgram = String(selected.insuranceProgram || "").trim();
          if (policyNumber) params.set("policyNumber", policyNumber);
          if (policyLimitType) params.set("policyLimitType", policyLimitType);
          if (insuranceProgram) params.set("insuranceProgram", insuranceProgram);
          window.location.href = `/Modules/PolicyInformation/index.html?${params.toString()}`;
        });

        function coverageLayerKey(slice) {
          return [
            String(slice?.PolicyLimitID || ""),
            String(slice?.PolicyID || ""),
            Number(slice?.attach || 0),
            Number(slice?.sliceLimit || 0),
            String(slice?.policyLimitTypeId || slice?.policyLimitType || ""),
            Number(slice?.policyStartMs || 0),
            Number(slice?.policyEndMs || 0)
          ].join("||");
        }

        function computeAvailableCoverage(slices, carriers, carrierGroups, annualized) {
          const carrierSet = new Set(carriers);
          const groupSet = new Set(carrierGroups);
          const rows = slices || [];
          if (annualized) {
            return rows.reduce((sum, s) => {
              if (carrierSet.size > 0 && !carrierSet.has(String(s?.carrier || ""))) return sum;
              if (groupSet.size > 0 && !groupSet.has(String(s?.carrierGroup || ""))) return sum;
              const unavailable = String(s?.availability || "").toLowerCase().includes("unavail");
              return sum + (unavailable ? 0 : Number(s?.sliceLimit || 0));
            }, 0);
          }

          const seen = new Set();
          let total = 0;
          for (const s of rows) {
            if (carrierSet.size > 0 && !carrierSet.has(String(s?.carrier || ""))) continue;
            if (groupSet.size > 0 && !groupSet.has(String(s?.carrierGroup || ""))) continue;
            const unavailable = String(s?.availability || "").toLowerCase().includes("unavail");
            if (unavailable) continue;
            const key = coverageLayerKey(s);
            if (seen.has(key)) continue;
            seen.add(key);
            total += Number(s?.sliceLimit || 0);
          }
          return total;
        }
        function computeAvailableCoverageByYear(slices, carriers, carrierGroups) {
          const carrierSet = new Set(carriers);
          const groupSet = new Set(carrierGroups);
          const byYearTop = new Map();
          for (const s of slices || []) {
            if (carrierSet.size > 0 && !carrierSet.has(String(s?.carrier || ""))) continue;
            if (groupSet.size > 0 && !groupSet.has(String(s?.carrierGroup || ""))) continue;
            if (String(s?.availability || "").toLowerCase().includes("unavail")) continue;
            const yearKey = String(s?.x ?? "").trim();
            if (!yearKey) continue;
            const attach = Number(s?.attach || 0);
            const limit = Number(s?.sliceLimit || 0);
            const top = attach + limit;
            if (!Number.isFinite(top) || !Number.isFinite(limit) || limit <= 0) continue;
            byYearTop.set(yearKey, Math.max(byYearTop.get(yearKey) || 0, top));
          }
          return Array.from(byYearTop.entries())
            .sort((a, b) => Number(a[0]) - Number(b[0]))
            .map(([year, total]) => ({ year, total }));
        }
        function formatTenthMillion(value) {
          const n = Number(value || 0);
          const rounded = Math.round((n / 1_000_000) * 10) / 10;
          return `$${rounded.toFixed(1)}M`;
        }

        function closeExportMenu() {
          exportMenuPanel.hidden = true;
          exportBtn.setAttribute("aria-expanded", "false");
        }

        function openExportMenu() {
          exportMenuPanel.hidden = false;
          exportBtn.setAttribute("aria-expanded", "true");
        }

        exportBtn.addEventListener("click", () => {
          if (exportMenuPanel.hidden) openExportMenu();
          else closeExportMenu();
        });

        document.addEventListener("click", (evt) => {
          if (!exportMenu.contains(evt.target)) closeExportMenu();
        });

        exportPngBtn.addEventListener("click", () => {
          exportChartAsPNG();
          closeExportMenu();
        });

        exportCsvBtn.addEventListener("click", () => {
          exportFilteredCSV();
          closeExportMenu();
        });

        exportPdfBtn.addEventListener("click", async () => {
          try {
            await exportReportPDF();
          } catch (err) {
            showError(err);
          } finally {
            closeExportMenu();
          }
        });

        function syncThemeUI(theme) {
          applyThemeToPage(theme, { themeLabelEl: themeLabel, themeToggleBtn });
          setChartTheme(theme);
        }

        syncThemeUI(initialTheme);
        themeToggleBtn.addEventListener("click", () => {
          const current = document.documentElement.dataset.theme === "light" ? "light" : "dark";
          const next = current === "light" ? "dark" : "light";
          localStorage.setItem(THEME_STORAGE_KEY, next);
          syncThemeUI(next);
        });

        const bounds = getYearBounds();
        if (bounds.minYear !== null && bounds.maxYear !== null) {
          for (let year = bounds.minYear; year <= bounds.maxYear; year++) {
            const a = document.createElement("option");
            a.value = String(year);
            a.textContent = String(year);
            startYearSelect.appendChild(a);

            const b = document.createElement("option");
            b.value = String(year);
            b.textContent = String(year);
            endYearSelect.appendChild(b);
          }
        }

        const filterOptions = getFilterOptions();
        buildCheckboxMenu(carrierDropdownMenu, filterOptions.carriers, "carrier_filter", "Search carriers...");
        buildCheckboxMenu(carrierGroupDropdownMenu, filterOptions.carrierGroups, "carrier_group_filter", "Search carrier groups...");
        const policyLimitTypeValues = filterOptions.policyLimitTypes || [];
        for (const policyLimitType of policyLimitTypeValues) {
          const opt = document.createElement("option");
          opt.value = policyLimitType;
          opt.textContent = policyLimitType;
          policyLimitTypeSelect.appendChild(opt);
        }
        if (policyLimitTypeValues.length > 0) {
          const defaultPolicyLimitType = policyLimitTypeValues.includes("Personal Injury")
            ? "Personal Injury"
            : policyLimitTypeValues[0];
          policyLimitTypeSelect.value = defaultPolicyLimitType;
          setPolicyLimitTypeFilter(defaultPolicyLimitType);
        }
        const insuranceProgramValues = filterOptions.insurancePrograms || [];
        for (const program of insuranceProgramValues) {
          const opt = document.createElement("option");
          opt.value = program;
          opt.textContent = program;
          insuranceProgramSelect.appendChild(opt);
        }
        if (insuranceProgramValues.length > 0) {
          const defaultInsuranceProgram = insuranceProgramValues.find(
            (v) => String(v).trim().toLowerCase() === "abc company"
          ) || insuranceProgramValues[0];
          insuranceProgramSelect.value = defaultInsuranceProgram;
          setInsuranceProgramFilter(defaultInsuranceProgram);
        }

        function applyYearFilters() {
          const toYear = (v) => {
            if (v === null || v === undefined || String(v).trim() === "") return null;
            const n = Number.parseInt(String(v), 10);
            return Number.isFinite(n) ? n : null;
          };
          const start = toYear(startYearSelect.value);
          const end = toYear(endYearSelect.value);
          const boundsNow = getYearBounds();
          const minBound = Number.isFinite(boundsNow.minYear) ? boundsNow.minYear : null;
          const maxBound = Number.isFinite(boundsNow.maxYear) ? boundsNow.maxYear : null;

          // Always pass concrete bounds to avoid any open-ended interpretation drift.
          const effectiveStart = start !== null ? start : minBound;
          const effectiveEnd = end !== null ? end : maxBound;

          if (effectiveStart === null && effectiveEnd === null) {
            resetYearRange();
            updateFilterSummary();
            return;
          }
          setYearRange(effectiveStart, effectiveEnd);
          updateFilterSummary();
        }

        function applyZoomFilters() {
          setZoomRange(zoomMinInput.value, zoomMaxInput.value);
          updateFilterSummary();
        }

        function applyInsuranceProgramFilter() {
          setInsuranceProgramFilter(insuranceProgramSelect.value);
          updateFilterSummary();
        }

        function applyPolicyLimitTypeFilter() {
          setPolicyLimitTypeFilter(policyLimitTypeSelect.value);
          updateFilterSummary();
        }
        function applyEntityFilters(source = "") {
          let carriers = selectedValuesFromCheckboxMenu(carrierDropdownMenu);
          let carrierGroups = selectedValuesFromCheckboxMenu(carrierGroupDropdownMenu);

          // Mutual exclusivity: user can filter by custom carriers OR carrier groups, not both.
          if (source === "carrier" && carriers.length > 0 && carrierGroups.length > 0) {
            clearCheckboxMenu(carrierGroupDropdownMenu);
            carrierGroups = [];
          } else if (source === "carrierGroup" && carrierGroups.length > 0 && carriers.length > 0) {
            clearCheckboxMenu(carrierDropdownMenu);
            carriers = [];
          } else if (!source && carriers.length > 0 && carrierGroups.length > 0) {
            // Fallback if both are somehow set simultaneously.
            clearCheckboxMenu(carrierGroupDropdownMenu);
            carrierGroups = [];
          }

          updateDropdownLabel(carrierDropdownLabel, carriers, "All carriers", "carrier");
          updateDropdownLabel(carrierGroupDropdownLabel, carrierGroups, "All carrier groups", "group");
          setEntityFilters({
            carriers,
            carrierGroups
          });
          updateFilterSummary();
        }

        function summarizeSelection(values, singularLabel) {
          if (!values.length) return `All ${singularLabel}s`;
          if (values.length === 1) return values[0];
          return `${values.length} ${singularLabel}s`;
        }

        function renderFilterChips(chipValues) {
          if (!activeFilterSummary) return;
          activeFilterSummary.innerHTML = "";
          const fragment = document.createDocumentFragment();
          chipValues
            .filter((value) => String(value || "").trim().length)
            .forEach((value) => {
              const chip = document.createElement("span");
              chip.className = "filterChip";
              chip.textContent = value;
              fragment.appendChild(chip);
            });
          activeFilterSummary.appendChild(fragment);
        }

        function updateFilterSummary() {
          const startYear = startYearSelect.value || "All";
          const endYear = endYearSelect.value || "All";
          const yearsText = startYear === "All" && endYear === "All"
            ? "All years"
            : `${startYear} to ${endYear}`;
          const yearsChipText = startYear !== "All" && endYear !== "All"
            ? `${startYear}\u2013${endYear}`
            : yearsText;

          const carriers = selectedValuesFromCheckboxMenu(carrierDropdownMenu);
          const carrierGroups = selectedValuesFromCheckboxMenu(carrierGroupDropdownMenu);
          const yearTotals = computeAvailableCoverageByYear(getFilteredSlices(), carriers, carrierGroups);
          const totalsByYear = new Map(yearTotals.map((t) => [String(t.year), Number(t.total || 0)]));
          const anchors = getYearLabelAnchors();
          const totalAvailable = computeAvailableCoverage(
            getFilteredSlices(),
            carriers,
            carrierGroups,
            annualizedMode
          );
          availableCoverageBadge.textContent = `Total Available Coverage: ${money(totalAvailable)}`;
          availableCoverageBadge.hidden = !showCoverageBadge;
          if (yearTotalsStrip) {
            yearTotalsStrip.hidden = !showCoverageBadge;
            const surfaceW = Math.max(chartSurface?.clientWidth || 0, 240);
            yearTotalsStrip.innerHTML = `<div class="yearTotalsSurface" style="width:${surfaceW}px;">` +
              `<span class="yearTotalsLabel">Available Limits:</span>` +
              anchors
                .map((a) => {
                  const total = totalsByYear.get(String(a.x)) || 0;
                  return `<span class="yearTotalValue" style="left:${a.px}px;">${formatTenthMillion(total)}</span>`;
                })
                .join("") +
              `</div>`;
            if (chartViewport) yearTotalsStrip.scrollLeft = chartViewport.scrollLeft;
          }
          if (toggleCoverageBadgeBtn) {
            toggleCoverageBadgeBtn.textContent = showCoverageBadge ? "Hide Total Avaiable Limits" : "Show Total Avaiable Limits";
            toggleCoverageBadgeBtn.setAttribute("aria-pressed", String(showCoverageBadge));
          }

          renderFilterChips([
            insuranceProgramSelect.value || "(none)",
            policyLimitTypeSelect.value || "(none)",
            yearsChipText,
            `Annualized: ${annualizedMode ? "On" : "Off"}`,
            summarizeSelection(carriers, "carrier"),
            summarizeSelection(carrierGroups, "group")
          ]);
        }

        startYearSelect.addEventListener("input", applyYearFilters);
        startYearSelect.addEventListener("change", applyYearFilters);
        startYearSelect.addEventListener("blur", applyYearFilters);
        endYearSelect.addEventListener("input", applyYearFilters);
        endYearSelect.addEventListener("change", applyYearFilters);
        endYearSelect.addEventListener("blur", applyYearFilters);
        insuranceProgramSelect.addEventListener("change", applyInsuranceProgramFilter);
        policyLimitTypeSelect.addEventListener("change", applyPolicyLimitTypeFilter);

        zoomMinInput.addEventListener("change", applyZoomFilters);
        zoomMaxInput.addEventListener("change", applyZoomFilters);
        zoomMinInput.addEventListener("blur", applyZoomFilters);
        zoomMaxInput.addEventListener("blur", applyZoomFilters);
        carrierDropdownMenu.addEventListener("change", () => applyEntityFilters("carrier"));
        carrierGroupDropdownMenu.addEventListener("change", () => applyEntityFilters("carrierGroup"));
        if (chartViewport && yearTotalsStrip) {
          chartViewport.addEventListener("scroll", () => {
            yearTotalsStrip.scrollLeft = chartViewport.scrollLeft;
          });
        }
        if (toggleCoverageBadgeBtn) {
          toggleCoverageBadgeBtn.addEventListener("click", () => {
            showCoverageBadge = !showCoverageBadge;
            localStorage.setItem(COVERAGE_BADGE_STORAGE_KEY, showCoverageBadge ? "1" : "0");
            setCoverageTotalsVisible(showCoverageBadge);
            updateFilterSummary();
          });
        }
        annualizeToggleBtn.addEventListener("click", () => {
          annualizedMode = !annualizedMode;
          localStorage.setItem(ANNUALIZED_STORAGE_KEY, annualizedMode ? "1" : "0");
          setAnnualizedMode(annualizedMode);
          syncAnnualizeToggleUI();
          updateFilterSummary();
        });

        function resetEntityFilterUI() {
          clearCheckboxMenu(carrierDropdownMenu);
          clearCheckboxMenu(carrierGroupDropdownMenu);
          resetCheckboxMenuSearch(carrierDropdownMenu);
          resetCheckboxMenuSearch(carrierGroupDropdownMenu);
          updateDropdownLabel(carrierDropdownLabel, [], "All carriers", "carrier");
          updateDropdownLabel(carrierGroupDropdownLabel, [], "All carrier groups", "group");
        }

        resetAllBtn.addEventListener("click", () => {
          startYearSelect.value = "";
          endYearSelect.value = "";
          const defaultInsuranceProgram = Array.from(insuranceProgramSelect.options)
            .map((o) => o.value)
            .find((v) => String(v).trim().toLowerCase() === "abc company")
            || insuranceProgramSelect.options[0]?.value
            || "";
          insuranceProgramSelect.value = defaultInsuranceProgram;
          const defaultPolicyLimitType = Array.from(policyLimitTypeSelect.options)
            .map((o) => o.value)
            .find((v) => v === "Personal Injury")
            || policyLimitTypeSelect.options[0]?.value
            || "";
          policyLimitTypeSelect.value = defaultPolicyLimitType;
          zoomMinInput.value = "";
          zoomMaxInput.value = "";
          resetEntityFilterUI();
          resetYearRange();
          resetInsuranceProgramFilter();
          setInsuranceProgramFilter(defaultInsuranceProgram);
          resetZoomRange();
          resetPolicyLimitTypeFilter();
          setPolicyLimitTypeFilter(defaultPolicyLimitType);
          resetEntityFilters();
          showCoverageBadge = true;
          localStorage.setItem(COVERAGE_BADGE_STORAGE_KEY, "1");
          setCoverageTotalsVisible(true);
          annualizedMode = false;
          localStorage.setItem(ANNUALIZED_STORAGE_KEY, "0");
          setAnnualizedMode(false);
          syncAnnualizeToggleUI();
          updateFilterSummary();
        });
        updateFilterSummary();
      } catch (e) {
        showError(e);
      }
    }

    init();
  
