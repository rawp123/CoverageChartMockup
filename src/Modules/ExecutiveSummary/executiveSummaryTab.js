    import {
      renderCoverageChart,
      getFilterOptions,
      getYearBounds,
      setInsuranceProgramFilter,
      setPolicyLimitTypeFilter,
      setChartTheme,
      getFilteredSlices
    } from "/Modules/shared/js/coverage/coverageChartEngine.js";
    import { money, compactMoney, shortLabel } from "../shared/js/core/format.js";
    import { getPreferredTheme, applyThemeToPage } from "../shared/js/core/theme.js";
    import {
      buildCheckboxMenu,
      selectedValuesFromCheckboxMenu,
      updateDropdownLabel
    } from "../shared/js/ui/multiSelect.js";

    let execCarrierGroupChart = null;
    let execYearTrendChart = null;

    const TOP_GROUPS_TO_SHOW = 7;
    const YEAR_FILTER_DEBOUNCE_MS = 140;
    const executiveFilterState = {
      programId: "",
      coverageType: "BI", // BI | PD
      yearFrom: null,
      yearTo: null,
      carrierIds: [],
      carrierGroupIds: []
    };
    const executiveFilterMeta = {
      coverageTypeMap: {
        BI: "",
        PD: ""
      }
    };

    function getExecTheme() {
      const light = document.documentElement.dataset.theme === "light";
      return light
        ? {
            text: "rgba(15, 23, 42, 0.92)",
            axis: "rgba(15, 23, 42, 0.9)",
            grid: "rgba(15, 23, 42, 0.12)",
            bar: "rgba(37, 99, 235, 0.62)",
            barTop: "rgba(29, 78, 216, 0.9)",
            barMuted: "rgba(15, 23, 42, 0.25)",
            line: "#0f766e"
          }
        : {
            text: "rgba(241, 245, 249, 0.95)",
            axis: "rgba(226, 232, 240, 0.92)",
            grid: "rgba(148, 163, 184, 0.2)",
            bar: "rgba(96, 165, 250, 0.62)",
            barTop: "rgba(147, 197, 253, 0.9)",
            barMuted: "rgba(148, 163, 184, 0.45)",
            line: "#34d399"
          };
    }

    function showError(err) {
      const box = document.getElementById("errorBox");
      box.style.display = "block";
      box.textContent =
        "ERROR:\n" + (err?.stack || err?.message || String(err)) +
        "\n\nQuick checks:" +
        "\n- Chart.js loaded (CDN)" +
        "\n- CSV URLs correct" +
        "\n- Server is serving /data and /Modules";
    }

    function getSliceYear(slice) {
      const yr = Number.isFinite(slice?.year) ? Number(slice.year) : Number(slice?.x);
      return Number.isFinite(yr) ? Math.trunc(yr) : null;
    }

    function isUnavailableSlice(slice) {
      return String(slice?.availability || "").toLowerCase().includes("unavail");
    }

    function buildYearPeakTotals(slices) {
      const byYearRows = new Map();
      for (const s of slices || []) {
        const year = getSliceYear(s);
        if (!Number.isFinite(year)) continue;
        const attach = Number(s?.attach || 0);
        const limit = Number(s?.sliceLimit || 0);
        if (!Number.isFinite(attach) || !Number.isFinite(limit) || limit <= 0) continue;
        const startMs = Number(s?.yearOverlapStartMs || s?.policyStartMs || 0);
        const endMs = Number(s?.yearOverlapEndMs || s?.policyEndMs || 0);
        if (!Number.isFinite(startMs) || !Number.isFinite(endMs) || endMs < startMs) continue;
        if (!byYearRows.has(year)) byYearRows.set(year, []);
        byYearRows.get(year).push({
          attach,
          limit,
          unavailable: isUnavailableSlice(s),
          startMs,
          endMsExclusive: endMs + 1
        });
      }

      const years = [];
      for (const [year, rows] of byYearRows.entries()) {
        const bounds = Array.from(
          new Set(rows.flatMap((r) => [Number(r.startMs), Number(r.endMsExclusive)]))
        )
          .filter((v) => Number.isFinite(v))
          .sort((a, b) => a - b);

        let yearGrossTop = 0;
        let yearAvailableTop = 0;

        for (let i = 0; i < bounds.length - 1; i++) {
          const segStart = bounds[i];
          const segEnd = bounds[i + 1];
          if (!(segEnd > segStart)) continue;

          const activeRows = rows.filter((r) => Number(r.startMs) < segEnd && Number(r.endMsExclusive) > segStart);
          if (!activeRows.length) continue;

          const grossByAttach = new Map();
          const availableByAttach = new Map();
          for (const r of activeRows) {
            grossByAttach.set(r.attach, (grossByAttach.get(r.attach) || 0) + Number(r.limit || 0));
            if (!r.unavailable) {
              availableByAttach.set(r.attach, (availableByAttach.get(r.attach) || 0) + Number(r.limit || 0));
            }
          }

          const segGrossTop = Array.from(grossByAttach.entries()).reduce(
            (max, [attach, total]) => Math.max(max, Number(attach || 0) + Number(total || 0)),
            0
          );
          const segAvailableTop = Array.from(availableByAttach.entries()).reduce(
            (max, [attach, total]) => Math.max(max, Number(attach || 0) + Number(total || 0)),
            0
          );

          yearGrossTop = Math.max(yearGrossTop, segGrossTop);
          yearAvailableTop = Math.max(yearAvailableTop, segAvailableTop);
        }

        years.push({ year, gross: yearGrossTop, available: yearAvailableTop });
      }

      return years.sort((a, b) => Number(a.year) - Number(b.year));
    }

    function filterSlicesForExecutive(data, state) {
      const slices = Array.isArray(data) ? data : [];
      const selectedProgram = String(state?.programId || "").trim();
      const coverageType = String(state?.coverageType || "BI").trim();
      const coverageTypeMap = executiveFilterMeta.coverageTypeMap || {};
      const targetPolicyLimitType = String(coverageTypeMap[coverageType] || "").trim();
      const yearFrom = Number.isFinite(state?.yearFrom) ? Number(state.yearFrom) : null;
      const yearTo = Number.isFinite(state?.yearTo) ? Number(state.yearTo) : null;
      const carrierSet = new Set((state?.carrierIds || []).map((v) => String(v).trim()).filter(Boolean));
      const carrierGroupSet = new Set((state?.carrierGroupIds || []).map((v) => String(v).trim()).filter(Boolean));

      return slices.filter((s) => {
        if (selectedProgram && String(s?.insuranceProgram || "").trim() !== selectedProgram) return false;
        if (targetPolicyLimitType) {
          const sliceType = String(s?.policyLimitType || s?.policyLimitTypeId || "").trim();
          if (sliceType !== targetPolicyLimitType) return false;
        }

        const yr = getSliceYear(s);
        if (yearFrom !== null && (!Number.isFinite(yr) || yr < yearFrom)) return false;
        if (yearTo !== null && (!Number.isFinite(yr) || yr > yearTo)) return false;

        if (carrierSet.size > 0) {
          const carrier = String(s?.carrier || "(unknown carrier)");
          if (!carrierSet.has(carrier)) return false;
        }

        if (carrierGroupSet.size > 0) {
          const group = String(s?.carrierGroup || "(unknown group)");
          if (!carrierGroupSet.has(group)) return false;
        }

        return true;
      });
    }

    function computeExecutiveSummary(data, state) {
      const filteredSlices = filterSlicesForExecutive(data, state);
      const totalGross = filteredSlices.reduce((sum, s) => sum + Number(s?.sliceLimit || 0), 0);
      const totalAvailable = filteredSlices.reduce((sum, s) => {
        return sum + (isUnavailableSlice(s) ? 0 : Number(s?.sliceLimit || 0));
      }, 0);
      const pctAvailable = totalGross > 0 ? totalAvailable / totalGross : 0;

      const byGroup = new Map();
      for (const s of filteredSlices) {
        const group = String(s?.carrierGroup || "(unknown group)");
        const limit = Number(s?.sliceLimit || 0);
        const unavailable = isUnavailableSlice(s);
        if (!byGroup.has(group)) byGroup.set(group, { name: group, gross: 0, available: 0 });
        const row = byGroup.get(group);
        row.gross += limit;
        if (!unavailable) row.available += limit;
      }

      const carrierGroups = Array.from(byGroup.values()).sort((a, b) => {
        if (b.gross !== a.gross) return b.gross - a.gross;
        return String(a.name).localeCompare(String(b.name));
      });

      const groupsByGross = carrierGroups
        .slice()
        .sort((a, b) => (b.gross - a.gross) || String(a.name).localeCompare(String(b.name)));
      const largestCarrierGroupName = groupsByGross[0]?.name || "(none)";
      const largestCarrierGroupGross = Number(groupsByGross[0]?.gross || 0);
      const largestCarrierGroupPctOfGross = totalGross > 0 ? largestCarrierGroupGross / totalGross : 0;
      const top3CarrierGroups = groupsByGross.slice(0, 3);
      const top3CarrierGroupsGross = top3CarrierGroups.reduce((sum, g) => sum + Number(g?.gross || 0), 0);
      const top3CarrierGroupsPctOfGross = totalGross > 0 ? top3CarrierGroupsGross / totalGross : 0;
      const top3CarrierGroupsBreakdown = top3CarrierGroups.map((g) => ({
        name: String(g?.name || "(unknown group)"),
        pctOfGross: totalGross > 0 ? Number(g?.gross || 0) / totalGross : 0
      }));

      const uniquePolicies = new Set(
        filteredSlices.map((s) => String(s?.PolicyID || "").trim()).filter(Boolean)
      ).size;

      const yearSeries = buildYearPeakTotals(filteredSlices);
      const coverageSpan = yearSeries.length
        ? `${yearSeries[0].year}-${yearSeries[yearSeries.length - 1].year}`
        : "--";

      return {
        totalGross,
        totalAvailable,
        pctAvailable,
        uniquePolicies,
        largestCarrierGroupName,
        largestCarrierGroupPctOfGross,
        top3CarrierGroupsPctOfGross,
        top3CarrierGroupsBreakdown,
        coverageSpan,
        carrierGroups,
        yearSeries,
        filteredSlices
      };
    }

    function renderKPISection(summary) {
      document.getElementById("execKpiGross").textContent = money(summary.totalGross);
      document.getElementById("execKpiAvailable").textContent = money(summary.totalAvailable);
      document.getElementById("execKpiAvailablePct").textContent = `${(summary.pctAvailable * 100).toFixed(1)}%`;
      document.getElementById("execKpiLargestGroupPct").textContent = `${(summary.largestCarrierGroupPctOfGross * 100).toFixed(1)}%`;
      document.getElementById("execKpiTop3Pct").textContent = `${(summary.top3CarrierGroupsPctOfGross * 100).toFixed(1)}%`;
      document.getElementById("execKpiPolicyCount").textContent = Number(summary.uniquePolicies || 0).toLocaleString();
    }

    function renderCarrierGroupChart(summary) {
      const theme = getExecTheme();
      const canvas = document.getElementById("execCarrierGroupCanvas");
      const subtitle = document.getElementById("execCarrierGroupSubtitle");

      const groups = summary.carrierGroups || [];
      const displayGroups = groups.slice(0, TOP_GROUPS_TO_SHOW);
      if (groups.length > TOP_GROUPS_TO_SHOW) {
        const remainderTotal = groups
          .slice(TOP_GROUPS_TO_SHOW)
          .reduce((sum, g) => sum + Number(g?.gross || 0), 0);
        displayGroups.push({ name: "All Other Groups", gross: remainderTotal, isRemainder: true });
      }

      subtitle.textContent =
        groups.length > TOP_GROUPS_TO_SHOW
          ? `Top ${TOP_GROUPS_TO_SHOW} shown; remainder grouped`
          : "All groups shown";

      const labels = displayGroups.map((g) => g.name);
      const values = displayGroups.map((g) => Number(g?.gross || 0));
      const colors = displayGroups.map((g, idx) => {
        if (g.isRemainder) return theme.barMuted;
        return idx < 3 ? theme.barTop : theme.bar;
      });

      if (execCarrierGroupChart) execCarrierGroupChart.destroy();
      execCarrierGroupChart = new Chart(canvas.getContext("2d"), {
        type: "bar",
        data: {
          labels,
          datasets: [{
            label: "Gross Limits",
            data: values,
            backgroundColor: colors,
            borderColor: colors,
            borderWidth: 1,
            borderRadius: 4
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          animation: false,
          layout: {
            padding: { top: 2, right: 2, bottom: 0, left: 0 }
          },
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                title: (items) => String(labels[items?.[0]?.dataIndex] || ""),
                label: (ctx) => `Gross Limits: ${money(ctx.parsed.y)}`
              }
            }
          },
          scales: {
            x: {
              ticks: {
                color: theme.axis,
                maxRotation: 25,
                autoSkip: false,
                callback: (value, index) => shortLabel(labels[index], 16)
              },
              grid: { color: theme.grid }
            },
            y: {
              beginAtZero: true,
              ticks: { color: theme.axis, callback: (v) => compactMoney(v), maxTicksLimit: 6 },
              grid: { color: theme.grid }
            }
          }
        }
      });
    }

    function renderYearTrendChart(summary) {
      const theme = getExecTheme();
      const canvas = document.getElementById("execYearTrendCanvas");
      const labels = (summary.yearSeries || []).map((y) => String(y.year));
      const gross = (summary.yearSeries || []).map((y) => Number(y.gross || 0));
      const available = (summary.yearSeries || []).map((y) => Number(y.available || 0));
      const datasets = [
        {
          type: "bar",
          label: "Total Limits",
          data: gross,
          backgroundColor: theme.bar,
          borderColor: theme.bar,
          borderRadius: 4
        },
        {
          type: "line",
          label: "Available Limits",
          data: available,
          borderColor: theme.line,
          pointBackgroundColor: theme.line,
          pointRadius: 2,
          pointHoverRadius: 4,
          borderWidth: 2,
          tension: 0.2
        }
      ];

      if (execYearTrendChart) execYearTrendChart.destroy();
      execYearTrendChart = new Chart(canvas.getContext("2d"), {
        data: {
          labels,
          datasets
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          animation: false,
          layout: {
            padding: { top: 0, right: 2, bottom: 0, left: 0 }
          },
          plugins: {
            legend: {
              display: true,
              position: "top",
              align: "start",
              labels: { color: theme.text, boxWidth: 20, boxHeight: 8, padding: 10 }
            },
            tooltip: {
              displayColors: false,
              callbacks: {
                title: (items) => String(labels[items?.[0]?.dataIndex] || ""),
                label: () => null,
                afterBody: (items) => {
                  const index = Number(items?.[0]?.dataIndex);
                  if (!Number.isFinite(index)) return [];
                  return [
                    `Total Limits: ${money(gross[index] || 0)}`,
                    `Available Limits: ${money(available[index] || 0)}`
                  ];
                }
              }
            }
          },
          interaction: {
            mode: "index",
            intersect: false
          },
          scales: {
            x: { ticks: { color: theme.axis, maxRotation: 0, autoSkip: true }, grid: { color: theme.grid } },
            y: { beginAtZero: true, ticks: { color: theme.axis, callback: (v) => compactMoney(v), maxTicksLimit: 6 }, grid: { color: theme.grid } }
          }
        }
      });
    }

    function renderSummaryFacts(summary) {
      const factsEl = document.getElementById("execSummaryFacts");
      const topGroups = summary.top3CarrierGroupsBreakdown || [];
      const topGroupSubBullets = topGroups.length
        ? topGroups
            .map((g) => {
              return `<li>${String(g?.name || "(unknown group)")}: ${(Number(g?.pctOfGross || 0) * 100).toFixed(1)}%</li>`;
            })
            .join("")
        : `<li>(none)</li>`;

      const facts = [
        `<li>Largest carrier group: ${summary.largestCarrierGroupName}</li>`,
        `<li>
          Top 3 carrier groups composition:
          <ul class="execFactsSubList">${topGroupSubBullets}</ul>
        </li>`
      ];
      factsEl.innerHTML = facts.join("");
      document.getElementById("execCoverageSpan").textContent = `Coverage Span: ${summary.coverageSpan}`;
    }

    function getDefaultProgram(programs) {
      return programs.find((v) => String(v).trim().toLowerCase() === "abc company") || programs[0] || "";
    }

    function getDefaultPolicyLimitType(types) {
      return (
        types.find((v) => String(v).trim().toLowerCase() === "bodily injury") ||
        types.find((v) => String(v).trim().toLowerCase() === "personal injury") ||
        types[0] ||
        ""
      );
    }

    async function init() {
      try {
        const THEME_STORAGE_KEY = "coverageChartTheme";

        const insuranceProgramSelect = document.getElementById("insuranceProgramSelect");
        const insuranceProgramLabel = document.getElementById("insuranceProgramLabel");
        const injuryTypeToggle = document.getElementById("injuryTypeToggle");
        const execYearFromSelect = document.getElementById("execYearFromSelect");
        const execYearToSelect = document.getElementById("execYearToSelect");
        const execMoreFiltersBtn = document.getElementById("execMoreFiltersBtn");
        const execMoreFiltersDrawer = document.getElementById("execMoreFiltersDrawer");
        const execCarrierDropdownMenu = document.getElementById("execCarrierDropdownMenu");
        const execCarrierDropdownLabel = document.getElementById("execCarrierDropdownLabel");
        const execCarrierGroupDropdownMenu = document.getElementById("execCarrierGroupDropdownMenu");
        const execCarrierGroupDropdownLabel = document.getElementById("execCarrierGroupDropdownLabel");
        const themeToggleBtn = document.getElementById("themeToggleBtn");
        const themeLabel = document.getElementById("themeLabel");

        const uiState = {
          moreFiltersOpen: false
        };
        let recomputeTimer = null;

        const parseYearValue = (value) => {
          const n = Number(value);
          return Number.isFinite(n) ? Math.trunc(n) : null;
        };
        const setMoreFiltersOpen = (open) => {
          uiState.moreFiltersOpen = !!open;
          if (execMoreFiltersDrawer) execMoreFiltersDrawer.hidden = !uiState.moreFiltersOpen;
          if (execMoreFiltersBtn) execMoreFiltersBtn.setAttribute("aria-expanded", String(uiState.moreFiltersOpen));
        };
        const updateExecutiveDashboard = () => {
          const summary = computeExecutiveSummary(getFilteredSlices(), executiveFilterState);
          renderKPISection(summary);
          renderCarrierGroupChart(summary);
          renderSummaryFacts(summary);
          renderYearTrendChart(summary);
        };
        const requestExecutiveRecompute = ({ debounceMs = 0 } = {}) => {
          if (recomputeTimer) {
            clearTimeout(recomputeTimer);
            recomputeTimer = null;
          }
          if (debounceMs > 0) {
            recomputeTimer = setTimeout(() => {
              recomputeTimer = null;
              updateExecutiveDashboard();
            }, debounceMs);
            return;
          }
          updateExecutiveDashboard();
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

        const filterOptions = getFilterOptions();
        const yearBounds = getYearBounds();

        const insuranceProgramValues = filterOptions.insurancePrograms || [];
        insuranceProgramSelect.innerHTML = "";
        const programValues = insuranceProgramValues.length ? insuranceProgramValues : ["(unknown program)"];
        for (const program of programValues) {
          const opt = document.createElement("option");
          opt.value = program;
          opt.textContent = program;
          insuranceProgramSelect.appendChild(opt);
        }
        executiveFilterState.programId = getDefaultProgram(programValues);
        insuranceProgramSelect.value = executiveFilterState.programId;
        setInsuranceProgramFilter(executiveFilterState.programId);
        insuranceProgramSelect.hidden = false;
        insuranceProgramLabel.hidden = false;

        const policyLimitTypeValues = filterOptions.policyLimitTypes || [];
        const injuryTypes = policyLimitTypeValues.filter((v) => /injury|damage/i.test(String(v || "")));
        const bodilyInjuryType = injuryTypes.find((v) => /bodily\s*injury/i.test(String(v || ""))) || "";
        const propertyDamageType = injuryTypes.find((v) => /property\s*damage/i.test(String(v || ""))) || "";
        const defaultPolicyLimitType = getDefaultPolicyLimitType(policyLimitTypeValues);
        const biType = bodilyInjuryType || defaultPolicyLimitType;
        const pdType =
          propertyDamageType ||
          policyLimitTypeValues.find((v) => String(v || "").trim() !== String(biType || "").trim()) ||
          biType;
        executiveFilterMeta.coverageTypeMap = {
          BI: String(biType || "").trim(),
          PD: String(pdType || "").trim()
        };

        const coverageTypeOptions = [
          { key: "BI", label: "Bodily Injury", value: executiveFilterMeta.coverageTypeMap.BI, disabled: !executiveFilterMeta.coverageTypeMap.BI },
          { key: "PD", label: "Property Damage", value: executiveFilterMeta.coverageTypeMap.PD, disabled: !executiveFilterMeta.coverageTypeMap.PD }
        ];
        let activeCoverage = coverageTypeOptions.find((opt) => opt.key === executiveFilterState.coverageType) || coverageTypeOptions[0];
        if (activeCoverage.disabled) {
          activeCoverage = coverageTypeOptions.find((opt) => !opt.disabled) || coverageTypeOptions[0];
        }
        executiveFilterState.coverageType = activeCoverage?.key || "BI";
        if (activeCoverage?.value) {
          setPolicyLimitTypeFilter(activeCoverage.value);
        }

        if (injuryTypeToggle) {
          injuryTypeToggle.innerHTML = coverageTypeOptions
            .map((option) => {
              const isActive = option.key === executiveFilterState.coverageType;
              return (
                `<button ` +
                `type="button" ` +
                `class="injuryTypeBtn${isActive ? " isActive" : ""}" ` +
                `data-key="${option.key}" ` +
                `data-value="${option.value || ""}" ` +
                `${option.disabled ? "disabled" : ""} ` +
                `title="${option.value || option.label}">${option.label}</button>`
              );
            })
            .join("");
          injuryTypeToggle.addEventListener("click", (evt) => {
            const btn = evt.target.closest(".injuryTypeBtn");
            if (!btn || btn.disabled) return;
            const nextKey = String(btn.dataset.key || "");
            const nextValue = String(btn.dataset.value || "");
            if (!nextKey || nextKey === executiveFilterState.coverageType) return;
            executiveFilterState.coverageType = nextKey;
            if (nextValue) {
              setPolicyLimitTypeFilter(nextValue);
            }
            for (const b of injuryTypeToggle.querySelectorAll(".injuryTypeBtn")) {
              b.classList.toggle("isActive", b === btn);
            }
            requestExecutiveRecompute();
          });
        }

        if (execYearFromSelect && execYearToSelect) {
          const minYear = Number(yearBounds?.minYear);
          const maxYear = Number(yearBounds?.maxYear);
          execYearFromSelect.innerHTML = `<option value="">From</option>`;
          execYearToSelect.innerHTML = `<option value="">To</option>`;
          if (Number.isFinite(minYear) && Number.isFinite(maxYear) && minYear <= maxYear) {
            for (let y = minYear; y <= maxYear; y++) {
              const fromOpt = document.createElement("option");
              fromOpt.value = String(y);
              fromOpt.textContent = String(y);
              execYearFromSelect.appendChild(fromOpt);
              const toOpt = document.createElement("option");
              toOpt.value = String(y);
              toOpt.textContent = String(y);
              execYearToSelect.appendChild(toOpt);
            }
          }
          const handleYearRangeChange = () => {
            let fromYear = parseYearValue(execYearFromSelect.value);
            let toYear = parseYearValue(execYearToSelect.value);
            if (fromYear !== null && toYear !== null && fromYear > toYear) {
              [fromYear, toYear] = [toYear, fromYear];
              execYearFromSelect.value = String(fromYear);
              execYearToSelect.value = String(toYear);
            }
            executiveFilterState.yearFrom = fromYear;
            executiveFilterState.yearTo = toYear;
            requestExecutiveRecompute({ debounceMs: YEAR_FILTER_DEBOUNCE_MS });
          };
          execYearFromSelect.addEventListener("change", handleYearRangeChange);
          execYearToSelect.addEventListener("change", handleYearRangeChange);
        }

        if (execCarrierDropdownMenu && execCarrierDropdownLabel) {
          const carriers = filterOptions.carriers || [];
          buildCheckboxMenu(execCarrierDropdownMenu, carriers, "execCarriers", "Search carriers");
          updateDropdownLabel(execCarrierDropdownLabel, executiveFilterState.carrierIds, "All Carriers", "carrier");
          execCarrierDropdownMenu.addEventListener("change", () => {
            const values = selectedValuesFromCheckboxMenu(execCarrierDropdownMenu);
            executiveFilterState.carrierIds = values;
            updateDropdownLabel(execCarrierDropdownLabel, values, "All Carriers", "carrier");
            requestExecutiveRecompute();
          });
        }

        if (execCarrierGroupDropdownMenu && execCarrierGroupDropdownLabel) {
          const carrierGroups = filterOptions.carrierGroups || [];
          buildCheckboxMenu(execCarrierGroupDropdownMenu, carrierGroups, "execCarrierGroups", "Search carrier groups");
          updateDropdownLabel(execCarrierGroupDropdownLabel, executiveFilterState.carrierGroupIds, "All Carrier Groups", "carrier group");
          execCarrierGroupDropdownMenu.addEventListener("change", () => {
            const values = selectedValuesFromCheckboxMenu(execCarrierGroupDropdownMenu);
            executiveFilterState.carrierGroupIds = values;
            updateDropdownLabel(execCarrierGroupDropdownLabel, values, "All Carrier Groups", "carrier group");
            requestExecutiveRecompute();
          });
        }

        if (execMoreFiltersBtn) {
          execMoreFiltersBtn.addEventListener("click", () => {
            setMoreFiltersOpen(!uiState.moreFiltersOpen);
          });
          setMoreFiltersOpen(false);
        }

        const syncThemeUI = (theme) => {
          applyThemeToPage(theme, { themeLabelEl: themeLabel, themeToggleBtn });
          setChartTheme(theme);
          requestExecutiveRecompute();
        };

        syncThemeUI(initialTheme);

        themeToggleBtn.addEventListener("click", () => {
          const current = document.documentElement.dataset.theme === "light" ? "light" : "dark";
          const next = current === "light" ? "dark" : "light";
          localStorage.setItem(THEME_STORAGE_KEY, next);
          syncThemeUI(next);
        });

        insuranceProgramSelect.addEventListener("change", () => {
          executiveFilterState.programId = insuranceProgramSelect.value;
          setInsuranceProgramFilter(insuranceProgramSelect.value);
          requestExecutiveRecompute();
        });

        requestExecutiveRecompute();
      } catch (e) {
        showError(e);
      }
    }

    init();
  
