    import {
      renderCoverageChart,
      getFilterOptions,
      setInsuranceProgramFilter,
      setPolicyLimitTypeFilter,
      setChartTheme,
      getFilteredSlices
    } from "/Modules/shared/js/coverage/coverageChartEngine.js";
    import { money, compactMoney, shortLabel } from "../shared/js/core/format.js";
    import { getPreferredTheme, applyThemeToPage } from "../shared/js/core/theme.js";

    let execCarrierGroupChart = null;
    let execYearTrendChart = null;

    const TOP_GROUPS_TO_SHOW = 7;

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

    function buildYearPeakTotals(slices) {
      const byYearRows = new Map();
      for (const s of slices || []) {
        const year = Number.isFinite(s?.year) ? Number(s.year) : Number(s?.x);
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
          unavailable: String(s?.availability || "").toLowerCase().includes("unavail"),
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

    function buildExecutiveMetrics() {
      const slices = getFilteredSlices() || [];

      const gross = slices.reduce((sum, s) => sum + Number(s?.sliceLimit || 0), 0);
      const available = slices.reduce((sum, s) => {
        const unavailable = String(s?.availability || "").toLowerCase().includes("unavail");
        return sum + (unavailable ? 0 : Number(s?.sliceLimit || 0));
      }, 0);

      const byGroup = new Map();
      const byLimitType = new Map();
      for (const s of slices) {
        const limit = Number(s?.sliceLimit || 0);
        const group = String(s?.carrierGroup || "(unknown group)");
        const limitType = String(s?.policyLimitType || s?.policyLimitTypeId || "(unknown type)");
        byGroup.set(group, (byGroup.get(group) || 0) + limit);
        byLimitType.set(limitType, (byLimitType.get(limitType) || 0) + limit);
      }

      const groups = Array.from(byGroup.entries())
        .map(([group, total]) => ({ group, total }))
        .sort((a, b) => b.total - a.total);
      const limitTypes = Array.from(byLimitType.entries())
        .map(([type, total]) => ({ type, total }))
        .sort((a, b) => b.total - a.total);
      const years = buildYearPeakTotals(slices);

      const largestGroup = groups.length ? groups[0].total : 0;
      const top3 = groups.slice(0, 3).reduce((sum, g) => sum + g.total, 0);
      const uniquePolicies = new Set(
        slices.map((s) => String(s?.PolicyID || "").trim()).filter(Boolean)
      ).size;

      return {
        gross,
        available,
        availablePct: gross > 0 ? available / gross : 0,
        largestGroupPct: gross > 0 ? largestGroup / gross : 0,
        top3Pct: gross > 0 ? top3 / gross : 0,
        largestGroupName: groups[0]?.group || "(none)",
        groups,
        limitTypes,
        years,
        uniquePolicies
      };
    }

    function renderKPISection(metrics) {
      document.getElementById("execKpiGross").textContent = money(metrics.gross);
      document.getElementById("execKpiAvailable").textContent = money(metrics.available);
      document.getElementById("execKpiAvailablePct").textContent = `${(metrics.availablePct * 100).toFixed(1)}%`;
      document.getElementById("execKpiLargestGroupPct").textContent = `${(metrics.largestGroupPct * 100).toFixed(1)}%`;
      document.getElementById("execKpiTop3Pct").textContent = `${(metrics.top3Pct * 100).toFixed(1)}%`;
      document.getElementById("execKpiPolicyCount").textContent = metrics.uniquePolicies.toLocaleString();
    }

    function renderCarrierGroupChart(metrics) {
      const theme = getExecTheme();
      const canvas = document.getElementById("execCarrierGroupCanvas");
      const subtitle = document.getElementById("execCarrierGroupSubtitle");

      const groups = metrics.groups || [];
      const displayGroups = groups.slice(0, TOP_GROUPS_TO_SHOW);
      if (groups.length > TOP_GROUPS_TO_SHOW) {
        const remainderTotal = groups
          .slice(TOP_GROUPS_TO_SHOW)
          .reduce((sum, g) => sum + Number(g?.total || 0), 0);
        displayGroups.push({ group: "All Other Groups", total: remainderTotal, isRemainder: true });
      }

      subtitle.textContent =
        groups.length > TOP_GROUPS_TO_SHOW
          ? `Top ${TOP_GROUPS_TO_SHOW} shown; remainder grouped`
          : "All groups shown";

      const labels = displayGroups.map((g) => g.group);
      const values = displayGroups.map((g) => g.total);
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
            label: "Total Limits",
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
                label: (ctx) => `Total: ${money(ctx.parsed.y)}`
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

    function renderYearTrendChart(metrics) {
      const theme = getExecTheme();
      const canvas = document.getElementById("execYearTrendCanvas");
      const labels = metrics.years.map((y) => String(y.year));
      const gross = metrics.years.map((y) => y.gross);
      const available = metrics.years.map((y) => y.available);

      if (execYearTrendChart) execYearTrendChart.destroy();
      execYearTrendChart = new Chart(canvas.getContext("2d"), {
        data: {
          labels,
          datasets: [
            {
              type: "bar",
              label: "Gross Limits",
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
          ]
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
              position: "top",
              align: "start",
              labels: { color: theme.text, boxWidth: 20, boxHeight: 8, padding: 10 }
            }
          },
          scales: {
            x: { ticks: { color: theme.axis, maxRotation: 0, autoSkip: true }, grid: { color: theme.grid } },
            y: { beginAtZero: true, ticks: { color: theme.axis, callback: (v) => compactMoney(v), maxTicksLimit: 6 }, grid: { color: theme.grid } }
          }
        }
      });
    }

    function renderSummaryFacts(metrics) {
      const factsEl = document.getElementById("execSummaryFacts");
      const spanStart = metrics.years.length ? metrics.years[0].year : "--";
      const spanEnd = metrics.years.length ? metrics.years[metrics.years.length - 1].year : "--";
      const topGroups = (metrics.groups || []).slice(0, 3);
      const topGroupSubBullets = topGroups.length
        ? topGroups
            .map((g) => {
              const pctOfGross = metrics.gross > 0 ? (Number(g?.total || 0) / Number(metrics.gross)) * 100 : 0;
              return `<li>${String(g?.group || "(unknown group)")}: ${pctOfGross.toFixed(1)}%</li>`;
            })
            .join("")
        : `<li>(none)</li>`;

      const facts = [
        `<li>Largest carrier group: ${metrics.largestGroupName}</li>`,
        `<li>
          Top 3 carrier groups composition:
          <ul class="execFactsSubList">${topGroupSubBullets}</ul>
        </li>`
      ];
      factsEl.innerHTML = facts.join("");
      document.getElementById("execCoverageSpan").textContent = `Coverage Span: ${spanStart}-${spanEnd}`;
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
        const themeToggleBtn = document.getElementById("themeToggleBtn");
        const themeLabel = document.getElementById("themeLabel");

        const updateExecutiveDashboard = () => {
          const metrics = buildExecutiveMetrics();
          renderKPISection(metrics);
          renderCarrierGroupChart(metrics);
          renderSummaryFacts(metrics);
          renderYearTrendChart(metrics);
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

        const policyLimitTypeValues = filterOptions.policyLimitTypes || [];
        const injuryTypes = policyLimitTypeValues.filter((v) => /injury|damage/i.test(String(v || "")));
        const bodilyInjuryType = injuryTypes.find((v) => /bodily\s*injury/i.test(String(v || ""))) || "";
        const propertyDamageType = injuryTypes.find((v) => /property\s*damage/i.test(String(v || ""))) || "";
        const toggleTypes = [
          ...(bodilyInjuryType ? [bodilyInjuryType] : []),
          ...(propertyDamageType && propertyDamageType !== bodilyInjuryType ? [propertyDamageType] : [])
        ];
        if (!toggleTypes.length) {
          if (injuryTypes[0]) toggleTypes.push(injuryTypes[0]);
          if (injuryTypes[1] && injuryTypes[1] !== injuryTypes[0]) toggleTypes.push(injuryTypes[1]);
        }
        const defaultPolicyLimitType =
          getDefaultPolicyLimitType(toggleTypes.length ? toggleTypes : policyLimitTypeValues);
        let activeInjuryType = defaultPolicyLimitType || toggleTypes[0] || "";
        if (activeInjuryType) setPolicyLimitTypeFilter(activeInjuryType);

        if (injuryTypeToggle) {
          injuryTypeToggle.innerHTML = toggleTypes
            .map((name) => {
              const isActive = name === activeInjuryType;
              return `<button type="button" class="injuryTypeBtn${isActive ? " isActive" : ""}" data-value="${name}">${name}</button>`;
            })
            .join("");
          injuryTypeToggle.addEventListener("click", (evt) => {
            const btn = evt.target.closest(".injuryTypeBtn");
            if (!btn) return;
            const value = String(btn.dataset.value || "");
            if (!value || value === activeInjuryType) return;
            activeInjuryType = value;
            setPolicyLimitTypeFilter(activeInjuryType);
            for (const b of injuryTypeToggle.querySelectorAll(".injuryTypeBtn")) {
              b.classList.toggle("isActive", b === btn);
            }
            updateExecutiveDashboard();
          });
        }

        const insuranceProgramValues = filterOptions.insurancePrograms || [];
        insuranceProgramSelect.innerHTML = "";
        const programValues = insuranceProgramValues.length ? insuranceProgramValues : ["(unknown program)"];
        for (const program of programValues) {
          const opt = document.createElement("option");
          opt.value = program;
          opt.textContent = program;
          insuranceProgramSelect.appendChild(opt);
        }
        const defaultProgram = getDefaultProgram(programValues);
        insuranceProgramSelect.value = defaultProgram;
        setInsuranceProgramFilter(defaultProgram);
        insuranceProgramSelect.hidden = false;
        insuranceProgramLabel.hidden = false;

        const syncThemeUI = (theme) => {
          applyThemeToPage(theme, { themeLabelEl: themeLabel, themeToggleBtn });
          setChartTheme(theme);
          updateExecutiveDashboard();
        };

        syncThemeUI(initialTheme);

        themeToggleBtn.addEventListener("click", () => {
          const current = document.documentElement.dataset.theme === "light" ? "light" : "dark";
          const next = current === "light" ? "dark" : "light";
          localStorage.setItem(THEME_STORAGE_KEY, next);
          syncThemeUI(next);
        });

        insuranceProgramSelect.addEventListener("change", () => {
          setInsuranceProgramFilter(insuranceProgramSelect.value);
          updateExecutiveDashboard();
        });

        updateExecutiveDashboard();
      } catch (e) {
        showError(e);
      }
    }

    init();
  
