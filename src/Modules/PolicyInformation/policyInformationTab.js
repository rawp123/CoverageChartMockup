import { fetchCSV, getBy, toNum } from "../shared/js/core/data.js";
import { money, formatDate as formatDateUS, toDateStamp, sanitizeFilePart } from "../shared/js/core/format.js";
import { ensureJSPdf } from "../shared/js/core/external.js";
import { getPreferredTheme, applyThemeToPage } from "../shared/js/core/theme.js";
import {
  buildCheckboxMenu,
  selectedValuesFromCheckboxMenu,
  clearCheckboxMenu,
  resetCheckboxMenuSearch,
  updateDropdownLabel
} from "../shared/js/ui/multiSelect.js";
import { setupSortableTable } from "../shared/js/ui/tableSort.js";

    const THEME_STORAGE_KEY = "coverageChartTheme";
    const SELECTED_POLICY_STORAGE_KEY = "coverageChartSelectedPolicy";

    const themeLabel = document.getElementById("themeLabel");
    const themeToggleBtn = document.getElementById("themeToggleBtn");

    const policyNumberSearch = document.getElementById("policyNumberSearch");
    const insuranceProgramSelect = document.getElementById("insuranceProgramSelect");
    const policyLimitTypeSelect = document.getElementById("policyLimitTypeSelect");
    const carrierDropdownMenu = document.getElementById("carrierDropdownMenu");
    const carrierGroupDropdownMenu = document.getElementById("carrierGroupDropdownMenu");
    const carrierDropdownLabel = document.getElementById("carrierDropdownLabel");
    const carrierGroupDropdownLabel = document.getElementById("carrierGroupDropdownLabel");
    const startDateFilter = document.getElementById("startDateFilter");
    const endDateFilter = document.getElementById("endDateFilter");
    const minLimitFilter = document.getElementById("minLimitFilter");
    const maxLimitFilter = document.getElementById("maxLimitFilter");
    const resetPolicyFiltersBtn = document.getElementById("resetPolicyFiltersBtn");

    const policyResultsCount = document.getElementById("policyResultsCount");
    const policyResultsList = document.getElementById("policyResultsList");

    const exportPolicyReportBtn = document.getElementById("exportPolicyReportBtn");
    const policyExportStatus = document.getElementById("policyExportStatus");

    const policyOverviewMeta = document.getElementById("policyOverviewMeta");
    const policyOverviewText = document.getElementById("policyOverviewText");
    const policyLimitsTable = document.querySelector(".policyTable");
    const policyLimitsBody = document.getElementById("policyLimitsBody");
    const policyTermsList = document.getElementById("policyTermsList");
    const policyConditionsList = document.getElementById("policyConditionsList");
    const policyExclusionsList = document.getElementById("policyExclusionsList");
    const policyEndorsementsList = document.getElementById("policyEndorsementsList");

    let allPolicies = [];
    let filteredPolicies = [];
    let selectedPolicyId = "";
    let initialUrlPolicyNumber = "";
    let initialUrlPolicyLimitType = "";
    let initialUrlInsuranceProgram = "";
    let policyLimitsSortState = { key: "attachment", direction: "asc" };
    const formatDate = (dateStr) => formatDateUS(dateStr, { month: "long" });

    function setPolicyExportStatus(message, isError = false) {
      policyExportStatus.textContent = message || "";
      policyExportStatus.classList.toggle("isError", Boolean(message) && isError);
    }

    function getSelectedPolicy() {
      if (!selectedPolicyId) return null;
      return allPolicies.find((p) => p.policyId === selectedPolicyId) || null;
    }

    function getURLSelectedPolicyId() {
      const params = new URLSearchParams(window.location.search);
      return String(params.get("policyId") || "").trim();
    }

    function getURLPolicyNumber() {
      const params = new URLSearchParams(window.location.search);
      return String(params.get("policyNumber") || "").trim();
    }

    function getURLPolicyLimitType() {
      const params = new URLSearchParams(window.location.search);
      return String(params.get("policyLimitType") || "").trim();
    }

    function getURLInsuranceProgram() {
      const params = new URLSearchParams(window.location.search);
      return String(params.get("insuranceProgram") || "").trim();
    }

    function getMatchingSelectOptionValue(selectEl, rawValue) {
      const target = String(rawValue || "").trim();
      if (!target) return "";
      const exact = Array.from(selectEl.options).find((opt) => String(opt.value || "").trim() === target);
      if (exact) return exact.value;
      const lowerTarget = target.toLowerCase();
      const caseInsensitive = Array.from(selectEl.options).find(
        (opt) => String(opt.value || "").trim().toLowerCase() === lowerTarget
      );
      return caseInsensitive ? caseInsensitive.value : "";
    }

    function getStoredSelectedPolicyId() {
      try {
        const raw = localStorage.getItem(SELECTED_POLICY_STORAGE_KEY);
        if (!raw) return "";
        const parsed = JSON.parse(raw);
        return String(parsed?.policyId || "").trim();
      } catch {
        return "";
      }
    }

    function buildPolicyData({ policyRows, dateRows, limitRows, carrierRows, carrierGroupRows, programRows, limitTypeRows }) {
      const dateByPolicyId = new Map(dateRows.map((r) => [String(getBy(r, "PolicyID")), r]));
      const carrierById = new Map(carrierRows.map((r) => [String(getBy(r, "CarrierID")), r]));
      const groupById = new Map(carrierGroupRows.map((r) => [String(getBy(r, "CarrierGroupID")), r]));
      const programById = new Map(programRows.map((r) => [String(getBy(r, "InsuranceProgramID")), r]));
      const limitTypeById = new Map(limitTypeRows.map((r) => [String(getBy(r, "PolicyLimitTypeID")), r]));

      const limitsByPolicyId = new Map();
      for (const row of limitRows) {
        const policyId = String(getBy(row, "PolicyID") || "").trim();
        if (!policyId) continue;
        if (!limitsByPolicyId.has(policyId)) limitsByPolicyId.set(policyId, []);
        const attach = toNum(getBy(row, "AttachmentPoint", "Attachment Point"));
        const layer = toNum(getBy(row, "LayerPerOccLimit", "Layer Per Occ Limit", "PerOccLimit", "Per Occ Limit"));
        const perOcc = toNum(getBy(row, "PerOccLimit", "Per Occ Limit"));
        const aggregate = toNum(getBy(row, "AggregateLimit", "Aggregate Limit"));
        const typeId = String(getBy(row, "PolicyLimitTypeID", "Policy Limit Type ID"));
        const typeRow = limitTypeById.get(typeId) || {};
        const typeName = String(getBy(typeRow, "PolicyLimitTypeName", "PolicyLimitType", "Policy Limit Type") || typeId).trim();
        limitsByPolicyId.get(policyId).push({
          typeName,
          attach,
          layer,
          perOcc,
          aggregate,
          top: attach + layer
        });
      }

      const policies = [];
      for (const p of policyRows) {
        const policyId = String(getBy(p, "PolicyID") || "").trim();
        if (!policyId) continue;

        const dateRow = dateByPolicyId.get(policyId) || {};
        const carrier = carrierById.get(String(getBy(p, "CarrierID") || "")) || {};
        const groupId = String(getBy(carrier, "CarrierGroupID", "Carrier Group ID", "CarrierGroupId") || getBy(p, "CarrierGroupID", "Carrier Group ID") || "").trim();
        const group = groupById.get(groupId) || {};
        const program = programById.get(String(getBy(p, "InsuranceProgramID") || "")) || {};

        const limits = (limitsByPolicyId.get(policyId) || []).sort((a, b) => a.attach - b.attach || a.layer - b.layer);
        const maxLayerLimit = limits.reduce((m, l) => Math.max(m, Number(l.layer || 0)), 0);

        const startDate = String(getBy(dateRow, "PStartDate", "PolicyStartDate", "StartDate") || "").trim();
        const endDate = String(getBy(dateRow, "PEndDate", "PolicyEndDate", "EndDate") || "").trim();
        const startMs = startDate ? new Date(`${startDate}T00:00:00Z`).getTime() : 0;
        const endMs = endDate ? new Date(`${endDate}T23:59:59Z`).getTime() : 0;

        policies.push({
          policyId,
          policyNumber: String(getBy(p, "PolicyNum", "PolicyNo", "policy_no", "Policy Number") || "").trim(),
          carrier: String(getBy(carrier, "CarrierName", "Carrier") || "(unknown carrier)").trim(),
          carrierGroup: String(getBy(group, "CarrierGroupName", "CarrierGroup", "Carrier Group") || "(unknown group)").trim(),
          insuranceProgram: String(getBy(program, "InsuranceProgram", "Program", "Name") || "(unknown program)").trim(),
          startDate,
          endDate,
          startMs,
          endMs,
          annualPeriod: String(getBy(dateRow, "AnnualPeriod") || "").trim(),
          layer: String(getBy(p, "Layer") || "").trim(),
          notes: String(getBy(p, "PolicyNotes") || "").trim(),
          sirPerOcc: toNum(getBy(p, "SIRPerOcc", "SIR Per Occ", "SIR")),
          sirAggregate: toNum(getBy(p, "SIRAggregate", "SIR Aggregate")),
          limits,
          maxLayerLimit
        });
      }

      return policies;
    }

    function populateFilterOptions(policies) {
      const programs = Array.from(new Set(policies.map((p) => p.insuranceProgram))).sort((a, b) => a.localeCompare(b));
      const limitTypes = Array.from(
        new Set(
          policies.flatMap((p) => (p.limits || []).map((l) => String(l.typeName || "").trim()).filter(Boolean))
        )
      ).sort((a, b) => a.localeCompare(b));
      const carriers = Array.from(new Set(policies.map((p) => p.carrier))).sort((a, b) => a.localeCompare(b));
      const groups = Array.from(new Set(policies.map((p) => p.carrierGroup))).sort((a, b) => a.localeCompare(b));

      for (const program of programs) {
        const opt = document.createElement("option");
        opt.value = program;
        opt.textContent = program;
        insuranceProgramSelect.appendChild(opt);
      }
      for (const limitType of limitTypes) {
        const opt = document.createElement("option");
        opt.value = limitType;
        opt.textContent = limitType;
        policyLimitTypeSelect.appendChild(opt);
      }
      buildCheckboxMenu(carrierDropdownMenu, carriers, "policy_carrier_filter", "Search carriers...");
      buildCheckboxMenu(carrierGroupDropdownMenu, groups, "policy_group_filter", "Search carrier groups...");
    }

    function getActiveLimitFilters() {
      const selectedLimitType = String(policyLimitTypeSelect.value || "").trim();
      const minLimit =
        String(minLimitFilter.value || "").trim() === ""
          ? Number.NEGATIVE_INFINITY
          : toNum(minLimitFilter.value);
      const maxLimit =
        String(maxLimitFilter.value || "").trim() === ""
          ? Number.POSITIVE_INFINITY
          : toNum(maxLimitFilter.value);

      return { selectedLimitType, minLimit, maxLimit };
    }

    function policyLimitMatchesFilters(limit, limitFilters) {
      const typeName = String(limit?.typeName || "").trim();
      const layer = Number(limit?.layer || 0);
      if (limitFilters.selectedLimitType && typeName !== limitFilters.selectedLimitType) return false;
      if (!Number.isFinite(layer)) return false;
      return layer >= limitFilters.minLimit && layer <= limitFilters.maxLimit;
    }

    function getVisiblePolicyLimits(policy, limitFilters = getActiveLimitFilters()) {
      return (policy?.limits || []).filter((limit) => policyLimitMatchesFilters(limit, limitFilters));
    }

    function normalizeSortText(value) {
      return String(value || "").trim().toLowerCase();
    }

    function getPolicyLimitSortValue(limit, key) {
      switch (key) {
        case "limitType":
          return normalizeSortText(limit?.typeName);
        case "attachment":
          return Number(limit?.attach || 0);
        case "layerLimit":
          return Number(limit?.layer || 0);
        case "topOfLayer":
          return Number(limit?.top || 0);
        case "perOcc":
          return Number(limit?.perOcc || 0);
        case "aggregate":
          return Number(limit?.aggregate || 0);
        default:
          return "";
      }
    }

    function comparePolicyLimits(a, b, sortState = policyLimitsSortState) {
      const direction = sortState?.direction === "desc" ? -1 : 1;
      const primaryA = getPolicyLimitSortValue(a, sortState?.key);
      const primaryB = getPolicyLimitSortValue(b, sortState?.key);
      if (typeof primaryA === "number" && typeof primaryB === "number") {
        if (primaryA !== primaryB) return (primaryA - primaryB) * direction;
      } else {
        const textCompare = String(primaryA).localeCompare(String(primaryB));
        if (textCompare !== 0) return textCompare * direction;
      }

      if (Number(a?.attach || 0) !== Number(b?.attach || 0)) {
        return Number(a?.attach || 0) - Number(b?.attach || 0);
      }
      if (Number(a?.layer || 0) !== Number(b?.layer || 0)) {
        return Number(a?.layer || 0) - Number(b?.layer || 0);
      }
      return normalizeSortText(a?.typeName).localeCompare(normalizeSortText(b?.typeName));
    }

    function sortPolicyLimits(limits, sortState = policyLimitsSortState) {
      return [...(limits || [])].sort((a, b) => comparePolicyLimits(a, b, sortState));
    }

    function applyPolicyFilters() {
      const policyText = String(policyNumberSearch.value || "").trim().toLowerCase();
      const exactPolicyNumber = String(initialUrlPolicyNumber || "").trim().toLowerCase();
      const useExactPolicyNumberMatch =
        !!policyText && !!exactPolicyNumber && policyText === exactPolicyNumber;
      const selectedProgram = String(insuranceProgramSelect.value || "").trim();
      const carriers = selectedValuesFromCheckboxMenu(carrierDropdownMenu);
      const groups = selectedValuesFromCheckboxMenu(carrierGroupDropdownMenu);
      const carrierSet = new Set(carriers);
      const groupSet = new Set(groups);
      updateDropdownLabel(carrierDropdownLabel, carriers, "All carriers", "carrier");
      updateDropdownLabel(carrierGroupDropdownLabel, groups, "All carrier groups", "group");

      const startDate = startDateFilter.value ? new Date(`${startDateFilter.value}T00:00:00Z`).getTime() : Number.NEGATIVE_INFINITY;
      const endDate = endDateFilter.value ? new Date(`${endDateFilter.value}T23:59:59Z`).getTime() : Number.POSITIVE_INFINITY;
      const limitFilters = getActiveLimitFilters();

      filteredPolicies = allPolicies.filter((p) => {
        if (policyText) {
          const policyNumber = String(p.policyNumber || "").trim().toLowerCase();
          if (useExactPolicyNumberMatch) {
            if (policyNumber !== policyText) return false;
          } else if (!policyNumber.includes(policyText)) {
            return false;
          }
        }
        if (selectedProgram && p.insuranceProgram !== selectedProgram) return false;
        if (carrierSet.size > 0 && !carrierSet.has(p.carrier)) return false;
        if (groupSet.size > 0 && !groupSet.has(p.carrierGroup)) return false;
        if (p.startMs > endDate || p.endMs < startDate) return false;
        if (!getVisiblePolicyLimits(p, limitFilters).length) return false;

        return true;
      });

      filteredPolicies.sort((a, b) => b.startMs - a.startMs || String(a.policyNumber).localeCompare(String(b.policyNumber)));
      renderPolicyResults();

      const stillExists = selectedPolicyId && filteredPolicies.some((p) => p.policyId === selectedPolicyId);
      if (stillExists) {
        renderPolicyDetails(filteredPolicies.find((p) => p.policyId === selectedPolicyId));
      } else {
        const first = filteredPolicies[0] || null;
        selectedPolicyId = first ? first.policyId : "";
        renderPolicyDetails(first);
      }
    }

    function renderPolicyResults() {
      policyResultsCount.textContent = `Policies: ${filteredPolicies.length}`;
      policyResultsList.innerHTML = "";

      for (const p of filteredPolicies) {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "policyResultItem";
        if (p.policyId === selectedPolicyId) btn.classList.add("isActive");
        btn.innerHTML = `
          <strong>${p.policyNumber || `Policy ${p.policyId}`}</strong>
          <span>${p.carrier} | ${p.carrierGroup}</span>
          <span>${formatDate(p.startDate)} - ${formatDate(p.endDate)}</span>
        `;
        btn.addEventListener("click", () => {
          selectedPolicyId = p.policyId;
          renderPolicyResults();
          renderPolicyDetails(p);
          const params = new URLSearchParams(window.location.search);
          params.set("policyId", p.policyId);
          history.replaceState(null, "", `${window.location.pathname}?${params.toString()}`);
        });
        policyResultsList.appendChild(btn);
      }
    }

    function fillList(el, items) {
      el.innerHTML = "";
      for (const item of items) {
        const li = document.createElement("li");
        li.textContent = item;
        el.appendChild(li);
      }
    }

    function getPolicyTerms(policy) {
      return [
        `${policy.policyNumber || "This policy"} is treated as occurrence-based coverage for placeholder analysis.`,
        "Coverage trigger and allocation language will be inserted when policy forms are loaded.",
        "Defense and indemnity obligations pending policy wording ingestion."
      ];
    }

    function getPolicyConditions() {
      return [
        "Notice and reporting obligations placeholder.",
        "Consent-to-settle and cooperation conditions placeholder.",
        "Other insurance and exhaustion conditions placeholder."
      ];
    }

    function getPolicyExclusions() {
      return [
        "Contractual liability exclusion placeholder.",
        "Known loss exclusion placeholder.",
        "Intentional acts exclusion placeholder."
      ];
    }

    function getPolicyEndorsements() {
      return [
        "Follow-form endorsement placeholder.",
        "Manuscript endorsement placeholder.",
        "Batch/occurrence wording endorsement placeholder."
      ];
    }

    function renderPolicyDetails(policy) {
      if (!policy) {
        exportPolicyReportBtn.disabled = true;
        setPolicyExportStatus("Select a policy to export a report.");
        policyOverviewMeta.innerHTML = "";
        policyOverviewText.textContent = "";
        policyOverviewText.hidden = true;
        policyLimitsBody.innerHTML = "";
        fillList(policyTermsList, []);
        fillList(policyConditionsList, []);
        fillList(policyExclusionsList, []);
        fillList(policyEndorsementsList, []);
        return;
      }

      exportPolicyReportBtn.disabled = false;
      if (policyExportStatus.textContent === "Select a policy to export a report.") {
        setPolicyExportStatus("");
      }

      policyOverviewMeta.innerHTML = `
        <div class="policyOverviewGrid" role="group" aria-label="Policy overview details">
          <section class="policyOverviewColumn" aria-label="Program context">
            <h3 class="policyOverviewSectionLabel">Program</h3>
            <div class="policyOverviewField">
              <span class="policyOverviewFieldLabel">Program:</span>
              <span class="policyOverviewFieldValue">${policy.insuranceProgram || "N/A"}</span>
            </div>
            <div class="policyOverviewField">
              <span class="policyOverviewFieldLabel">Policy Number:</span>
              <span class="policyOverviewFieldValue policyOverviewFieldValue--long">${policy.policyNumber || "N/A"}</span>
            </div>
            <div class="policyOverviewField">
              <span class="policyOverviewFieldLabel">Annual Period:</span>
              <span class="policyOverviewFieldValue">${policy.annualPeriod || "N/A"}</span>
            </div>
          </section>
          <section class="policyOverviewColumn" aria-label="Carrier information">
            <h3 class="policyOverviewSectionLabel">Carrier</h3>
            <div class="policyOverviewField">
              <span class="policyOverviewFieldLabel">Carrier:</span>
              <span class="policyOverviewFieldValue policyOverviewFieldValue--long">${policy.carrier || "N/A"}</span>
            </div>
            <div class="policyOverviewField">
              <span class="policyOverviewFieldLabel">Carrier Group:</span>
              <span class="policyOverviewFieldValue">${policy.carrierGroup || "N/A"}</span>
            </div>
            <div class="policyOverviewField">
              <span class="policyOverviewFieldLabel">Quota Share:</span>
              <span class="policyOverviewFieldValue">Yes</span>
            </div>
          </section>
          <section class="policyOverviewColumn" aria-label="Policy term">
            <h3 class="policyOverviewSectionLabel">Term</h3>
            <div class="policyOverviewField">
              <span class="policyOverviewFieldLabel">Policy Period:</span>
              <span class="policyOverviewFieldValue policyOverviewFieldValue--long">${formatDate(policy.startDate)} - ${formatDate(policy.endDate)}</span>
            </div>
          </section>
          <section class="policyOverviewColumn" aria-label="Financial structure">
            <h3 class="policyOverviewSectionLabel">Retention</h3>
            <div class="policyOverviewField">
              <span class="policyOverviewFieldLabel">SIR (Per Occurrence):</span>
              <span class="policyOverviewFieldValue">${money(policy.sirPerOcc)}</span>
            </div>
            <div class="policyOverviewField">
              <span class="policyOverviewFieldLabel">SIR (Aggregate):</span>
              <span class="policyOverviewFieldValue">${money(policy.sirAggregate)}</span>
            </div>
          </section>
        </div>
      `;

      if (String(policy.notes || "").trim()) {
        policyOverviewText.textContent = `Policy Notes: ${policy.notes}`;
        policyOverviewText.hidden = false;
      } else {
        policyOverviewText.textContent = "";
        policyOverviewText.hidden = true;
      }

      const visibleLimits = sortPolicyLimits(getVisiblePolicyLimits(policy));
      policyLimitsBody.innerHTML = "";
      if (!visibleLimits.length) {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td colspan="6">No policy limits match the current filters.</td>`;
        policyLimitsBody.appendChild(tr);
      } else {
        for (const l of visibleLimits) {
          const tr = document.createElement("tr");
          tr.innerHTML = `
            <td>${l.typeName || "N/A"}</td>
            <td>${money(l.attach)}</td>
            <td>${money(l.layer)}</td>
            <td>${money(l.top)}</td>
            <td>${money(l.perOcc)}</td>
            <td>${money(l.aggregate)}</td>
          `;
          policyLimitsBody.appendChild(tr);
        }
      }

      fillList(policyTermsList, getPolicyTerms(policy));
      fillList(policyConditionsList, getPolicyConditions());
      fillList(policyExclusionsList, getPolicyExclusions());
      fillList(policyEndorsementsList, getPolicyEndorsements());
    }

    async function exportPolicyReport() {
      const policy = getSelectedPolicy();
      if (!policy) {
        setPolicyExportStatus("Select a policy before exporting.", true);
        return;
      }

      const previousLabel = exportPolicyReportBtn.textContent;
      exportPolicyReportBtn.disabled = true;
      exportPolicyReportBtn.textContent = "Exporting...";
      setPolicyExportStatus("Building policy report PDF...");

      try {
        await ensureJSPdf();
        const jsPDF = window.jspdf?.jsPDF;
        if (!jsPDF) throw new Error("jsPDF is not available.");

        const pdf = new jsPDF({ orientation: "portrait", unit: "pt", format: "a4" });
        const pageW = pdf.internal.pageSize.getWidth();
        const pageH = pdf.internal.pageSize.getHeight();
        const margin = 40;
        const footerReserve = 24;
        const contentW = pageW - margin * 2;
        let y = margin;

        const ensurePageSpace = (needed = 20) => {
          if (y + needed <= pageH - margin - footerReserve) return false;
          pdf.addPage("a4", "portrait");
          y = margin;
          return true;
        };

        const drawSectionHeading = (title) => {
          ensurePageSpace(28);
          pdf.setFont("helvetica", "bold");
          pdf.setFontSize(13);
          pdf.setTextColor(18, 34, 58);
          pdf.text(title, margin, y);
          y += 8;
          pdf.setDrawColor(186, 196, 210);
          pdf.setLineWidth(0.6);
          pdf.line(margin, y, pageW - margin, y);
          y += 14;
        };

        const drawWrappedText = (text, { x = margin, width = contentW, fontStyle = "normal", fontSize = 10, lineHeight = 12, gapAfter = 6 } = {}) => {
          const lines = pdf.splitTextToSize(String(text || ""), width);
          ensurePageSpace(lines.length * lineHeight + gapAfter);
          pdf.setFont("helvetica", fontStyle);
          pdf.setFontSize(fontSize);
          pdf.setTextColor(36, 46, 66);
          lines.forEach((line, idx) => {
            pdf.text(line, x, y + idx * lineHeight);
          });
          y += lines.length * lineHeight + gapAfter;
        };

        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(20);
        pdf.setTextColor(16, 24, 39);
        pdf.text("Policy Report", margin, y);
        pdf.setFontSize(12);
        pdf.text(policy.policyNumber || `Policy ${policy.policyId}`, pageW - margin, y, { align: "right" });
        y += 20;

        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(10);
        pdf.setTextColor(90, 100, 120);
        pdf.text(`Generated: ${new Date().toLocaleString()}`, margin, y);
        y += 14;
        pdf.text(`Policy ID: ${policy.policyId || "N/A"}`, margin, y);
        y += 16;

        drawSectionHeading("Policy Overview");
        const overviewRows = [
          ["Policy Number", policy.policyNumber || "N/A"],
          ["Insurance Program", policy.insuranceProgram || "N/A"],
          ["Carrier", policy.carrier || "N/A"],
          ["Carrier Group", policy.carrierGroup || "N/A"],
          ["Policy Period", `${formatDate(policy.startDate)} - ${formatDate(policy.endDate)}`],
          ["Annual Period", policy.annualPeriod || "N/A"],
          ["SIR (Per Occ)", money(policy.sirPerOcc)],
          ["SIR (Aggregate)", money(policy.sirAggregate)]
        ];
        const columnGap = 18;
        const columnWidth = (contentW - columnGap) / 2;
        for (let i = 0; i < overviewRows.length; i += 2) {
          const left = overviewRows[i];
          const right = overviewRows[i + 1];
          const leftLines = pdf.splitTextToSize(`${left[0]}: ${left[1]}`, columnWidth);
          const rightLines = right ? pdf.splitTextToSize(`${right[0]}: ${right[1]}`, columnWidth) : [];
          const rowLines = Math.max(leftLines.length, rightLines.length || 1);
          const rowHeight = rowLines * 12 + 4;
          if (ensurePageSpace(rowHeight + 2)) drawSectionHeading("Policy Overview (cont.)");
          pdf.setFont("helvetica", "normal");
          pdf.setFontSize(10);
          pdf.setTextColor(36, 46, 66);
          leftLines.forEach((line, idx) => pdf.text(line, margin, y + idx * 12));
          rightLines.forEach((line, idx) => pdf.text(line, margin + columnWidth + columnGap, y + idx * 12));
          y += rowHeight;
        }
        y += 2;

        drawSectionHeading("Policy Notes");
        drawWrappedText(
          policy.notes
            ? policy.notes
            : "Policy notes/wording not provided in current dataset. Placeholder policy language shown below.",
          { lineHeight: 13, gapAfter: 8 }
        );

        const visibleLimits = getVisiblePolicyLimits(policy);
        drawSectionHeading("Policy Limits");
        if (!visibleLimits.length) {
          drawWrappedText("No policy limits are available for this policy with the current filters.");
        } else {
          const columns = [
            { label: "Limit Type", width: 158, align: "left" },
            { label: "Attachment", width: 72, align: "right" },
            { label: "Layer Limit", width: 72, align: "right" },
            { label: "Top of Layer", width: 72, align: "right" },
            { label: "Per Occ", width: 72, align: "right" },
            { label: "Aggregate", width: 72, align: "right" }
          ];
          const tableW = columns.reduce((sum, c) => sum + c.width, 0);

          const drawLimitsHeader = () => {
            ensurePageSpace(20);
            const top = y;
            let x = margin;
            pdf.setFillColor(232, 238, 248);
            pdf.setDrawColor(188, 198, 214);
            pdf.setLineWidth(0.4);
            pdf.rect(margin, top, tableW, 18, "FD");
            pdf.setFont("helvetica", "bold");
            pdf.setFontSize(9);
            pdf.setTextColor(25, 35, 52);
            for (const col of columns) {
              if (col.align === "right") {
                pdf.text(col.label, x + col.width - 4, top + 12, { align: "right" });
              } else {
                pdf.text(col.label, x + 4, top + 12);
              }
              x += col.width;
              if (x < margin + tableW) pdf.line(x, top, x, top + 18);
            }
            y = top + 18;
          };

          drawLimitsHeader();

          for (const limit of visibleLimits) {
            const cells = [
              String(limit.typeName || "N/A"),
              money(limit.attach),
              money(limit.layer),
              money(limit.top),
              money(limit.perOcc),
              money(limit.aggregate)
            ];
            const wrappedCells = columns.map((col, idx) => {
              const lines = pdf.splitTextToSize(cells[idx], col.width - 8);
              return lines.length ? lines : [""];
            });
            const lineCount = Math.max(...wrappedCells.map((lines) => lines.length));
            const rowHeight = lineCount * 10 + 6;

            if (y + rowHeight > pageH - margin - footerReserve) {
              pdf.addPage("a4", "portrait");
              y = margin;
              drawSectionHeading("Policy Limits (cont.)");
              drawLimitsHeader();
            }

            const top = y;
            let x = margin;
            pdf.setDrawColor(210, 218, 228);
            pdf.setLineWidth(0.3);
            pdf.rect(margin, top, tableW, rowHeight);
            pdf.setFont("helvetica", "normal");
            pdf.setFontSize(9);
            pdf.setTextColor(36, 46, 66);
            for (let i = 0; i < columns.length; i++) {
              const col = columns[i];
              const lines = wrappedCells[i];
              lines.forEach((line, idx) => {
                const lineY = top + 11 + idx * 10;
                if (col.align === "right") {
                  pdf.text(line, x + col.width - 4, lineY, { align: "right" });
                } else {
                  pdf.text(line, x + 4, lineY);
                }
              });
              x += col.width;
              if (x < margin + tableW) pdf.line(x, top, x, top + rowHeight);
            }
            y += rowHeight;
          }
          y += 8;
        }

        const headingBlockHeight = 28;
        const sectionBreakGap = 12;
        if (y + sectionBreakGap + headingBlockHeight > pageH - margin - footerReserve) {
          pdf.addPage("a4", "portrait");
          y = margin;
        } else {
          y += sectionBreakGap;
        }

        const sectionLists = [
          { title: "Terms", items: getPolicyTerms(policy) },
          { title: "Conditions", items: getPolicyConditions() },
          { title: "Exclusions", items: getPolicyExclusions() },
          { title: "Endorsements", items: getPolicyEndorsements() }
        ];
        for (const section of sectionLists) {
          drawSectionHeading(section.title);
          for (const item of section.items) {
            const lines = pdf.splitTextToSize(`- ${item}`, contentW - 6);
            if (y + lines.length * 12 + 2 > pageH - margin - footerReserve) {
              pdf.addPage("a4", "portrait");
              y = margin;
              drawSectionHeading(`${section.title} (cont.)`);
            }
            pdf.setFont("helvetica", "normal");
            pdf.setFontSize(10);
            pdf.setTextColor(36, 46, 66);
            lines.forEach((line, idx) => pdf.text(line, margin + 4, y + idx * 12));
            y += lines.length * 12 + 2;
          }
          y += 6;
        }

        const pageCount = pdf.getNumberOfPages();
        for (let pageNum = 1; pageNum <= pageCount; pageNum++) {
          pdf.setPage(pageNum);
          pdf.setFont("helvetica", "normal");
          pdf.setFontSize(9);
          pdf.setTextColor(116, 124, 139);
          pdf.text("Coverage Dashboard Policy Information", margin, pageH - 12);
          pdf.text(`Page ${pageNum} of ${pageCount}`, pageW - margin, pageH - 12, { align: "right" });
        }

        const fileName = `PolicyReport_${sanitizeFilePart(policy.policyNumber || policy.policyId)}_${toDateStamp()}.pdf`;
        pdf.save(fileName);
        setPolicyExportStatus(`Policy report exported: ${fileName}`);
      } catch (err) {
        console.error(err);
        setPolicyExportStatus(`Failed to export policy report: ${err?.message || err}`, true);
      } finally {
        exportPolicyReportBtn.textContent = previousLabel;
        exportPolicyReportBtn.disabled = !getSelectedPolicy();
      }
    }

    function resetFilters() {
      initialUrlPolicyNumber = "";
      initialUrlPolicyLimitType = "";
      initialUrlInsuranceProgram = "";
      policyNumberSearch.value = "";
      insuranceProgramSelect.value = "";
      policyLimitTypeSelect.value = "";
      clearCheckboxMenu(carrierDropdownMenu);
      clearCheckboxMenu(carrierGroupDropdownMenu);
      resetCheckboxMenuSearch(carrierDropdownMenu);
      resetCheckboxMenuSearch(carrierGroupDropdownMenu);
      updateDropdownLabel(carrierDropdownLabel, [], "All carriers", "carrier");
      updateDropdownLabel(carrierGroupDropdownLabel, [], "All carrier groups", "group");
      startDateFilter.value = "";
      endDateFilter.value = "";
      minLimitFilter.value = "";
      maxLimitFilter.value = "";
      applyPolicyFilters();
    }

    async function init() {
      applyThemeToPage(getPreferredTheme(THEME_STORAGE_KEY), { themeLabelEl: themeLabel, themeToggleBtn });
      themeToggleBtn.addEventListener("click", () => {
        const current = document.documentElement.dataset.theme === "light" ? "light" : "dark";
        const next = current === "light" ? "dark" : "light";
        localStorage.setItem(THEME_STORAGE_KEY, next);
        applyThemeToPage(next, { themeLabelEl: themeLabel, themeToggleBtn });
      });

      const [policyRows, dateRows, limitRows, carrierRows, carrierGroupRows, programRows, limitTypeRows] = await Promise.all([
        fetchCSV("/data/OriginalFiles/tblPolicy.csv"),
        fetchCSV("/data/OriginalFiles/tblPolicyDates.csv"),
        fetchCSV("/data/OriginalFiles/tblPolicyLimits.csv"),
        fetchCSV("/data/OriginalFiles/tblCarrier.csv"),
        fetchCSV("/data/OriginalFiles/tblCarrierGroup.csv"),
        fetchCSV("/data/OriginalFiles/tblInsuranceProgram.csv"),
        fetchCSV("/data/OriginalFiles/tblPolicyLimitType.csv")
      ]);

      allPolicies = buildPolicyData({
        policyRows,
        dateRows,
        limitRows,
        carrierRows,
        carrierGroupRows,
        programRows,
        limitTypeRows
      });

      populateFilterOptions(allPolicies);

      setupSortableTable({
        table: policyLimitsTable,
        columns: [
          { index: 0, key: "limitType", defaultDirection: "asc" },
          { index: 1, key: "attachment", defaultDirection: "asc" },
          { index: 2, key: "layerLimit", defaultDirection: "asc" },
          { index: 3, key: "topOfLayer", defaultDirection: "asc" },
          { index: 4, key: "perOcc", defaultDirection: "asc" },
          { index: 5, key: "aggregate", defaultDirection: "asc" }
        ],
        initialSort: policyLimitsSortState,
        onSortChange: (nextSortState) => {
          policyLimitsSortState = nextSortState;
          renderPolicyDetails(getSelectedPolicy());
        }
      });

      const queryPolicyId = getURLSelectedPolicyId();
      const queryPolicyNumber = getURLPolicyNumber();
      const queryPolicyLimitType = getURLPolicyLimitType();
      const queryInsuranceProgram = getURLInsuranceProgram();
      const storedPolicyId = getStoredSelectedPolicyId();
      selectedPolicyId = queryPolicyId || storedPolicyId || "";
      initialUrlPolicyNumber = queryPolicyNumber;
      initialUrlPolicyLimitType = queryPolicyLimitType;
      initialUrlInsuranceProgram = queryInsuranceProgram;
      if (!initialUrlPolicyNumber && selectedPolicyId) {
        const selectedPolicy = allPolicies.find((p) => p.policyId === selectedPolicyId);
        initialUrlPolicyNumber = String(selectedPolicy?.policyNumber || "").trim();
      }
      if (initialUrlPolicyNumber) {
        policyNumberSearch.value = initialUrlPolicyNumber;
      }
      if (initialUrlPolicyLimitType) {
        const matchingLimitType = getMatchingSelectOptionValue(policyLimitTypeSelect, initialUrlPolicyLimitType);
        if (matchingLimitType) policyLimitTypeSelect.value = matchingLimitType;
      }
      if (initialUrlInsuranceProgram) {
        const matchingProgram = getMatchingSelectOptionValue(insuranceProgramSelect, initialUrlInsuranceProgram);
        if (matchingProgram) insuranceProgramSelect.value = matchingProgram;
      }

      applyPolicyFilters();

      if (selectedPolicyId) {
        const exact = allPolicies.find((p) => p.policyId === selectedPolicyId);
        if (exact) {
          if (!filteredPolicies.some((p) => p.policyId === selectedPolicyId)) {
            filteredPolicies = [exact, ...filteredPolicies];
          }
          renderPolicyResults();
          renderPolicyDetails(exact);
        }
      }

      resetPolicyFiltersBtn.addEventListener("click", resetFilters);
      exportPolicyReportBtn.addEventListener("click", exportPolicyReport);

      for (const el of [policyNumberSearch, insuranceProgramSelect, policyLimitTypeSelect, startDateFilter, endDateFilter, minLimitFilter, maxLimitFilter]) {
        el.addEventListener("change", applyPolicyFilters);
      }
      carrierDropdownMenu.addEventListener("change", applyPolicyFilters);
      carrierGroupDropdownMenu.addEventListener("change", applyPolicyFilters);
      policyNumberSearch.addEventListener("input", () => {
        window.clearTimeout(policyNumberSearch._debounceTimer);
        policyNumberSearch._debounceTimer = window.setTimeout(applyPolicyFilters, 120);
      });
    }

    init().catch((err) => {
      console.error(err);
      policyOverviewText.textContent = `Failed to load policy data: ${err?.message || err}`;
    });
  
