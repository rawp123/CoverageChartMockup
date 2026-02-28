import { fetchCSV, getBy, normalizeISODate, toNum } from "../shared/js/core/data.js";
import { money, formatDate, toDateStamp, sanitizeFilePart } from "../shared/js/core/format.js";
import { ensureJSPdf } from "../shared/js/core/external.js";
import { getPreferredTheme, applyThemeToPage } from "../shared/js/core/theme.js";
import { setupSortableTable } from "../shared/js/ui/tableSort.js";

    const THEME_STORAGE_KEY = "coverageChartTheme";

    const themeLabel = document.getElementById("themeLabel");
    const themeToggleBtn = document.getElementById("themeToggleBtn");

    const languageTopicSelect = document.getElementById("languageTopicSelect");
    const languageProgramSelect = document.getElementById("languageProgramSelect");
    const languageStartDateFilter = document.getElementById("languageStartDateFilter");
    const languageEndDateFilter = document.getElementById("languageEndDateFilter");
    const languageMinLayerLimitFilter = document.getElementById("languageMinLayerLimitFilter");
    const languageMaxLayerLimitFilter = document.getElementById("languageMaxLayerLimitFilter");
    const exportLanguageReportBtn = document.getElementById("exportLanguageReportBtn");
    const resetLanguageFiltersBtn = document.getElementById("resetLanguageFiltersBtn");
    const languageExportStatus = document.getElementById("languageExportStatus");

    const languageTopicTitle = document.getElementById("languageTopicTitle");
    const languageExcerptCount = document.getElementById("languageExcerptCount");
    const languageSummaryStats = document.getElementById("languageSummaryStats");
    const languageScheduleCount = document.getElementById("languageScheduleCount");
    const languageScheduleTable = document.querySelector(".policyTable--language");
    const languageScheduleBody = document.getElementById("languageScheduleBody");
    const languageExcerptList = document.getElementById("languageExcerptList");
    const languageModal = document.getElementById("languageModal");
    const languageModalMeta = document.getElementById("languageModalMeta");
    const languageModalBody = document.getElementById("languageModalBody");
    const languageModalCloseBtn = document.getElementById("languageModalCloseBtn");
    const ALL_TOPICS_ID = "__all_policy_language__";
    const ALL_TOPICS_LABEL = "All Policy Language";
    const SUMMARY_ALL_TOPICS_LABEL = "Policy Language";

    function makeEntry(data) {
      const startMs = Date.parse(`${data.startDate}T00:00:00Z`);
      return { ...data, startMs: Number.isFinite(startMs) ? startMs : 0 };
    }

    let LANGUAGE_TOPICS = [];
    let CURRENT_SCHEDULE_ROWS = [];
    let languageScheduleSortState = { key: "carrier", direction: "asc" };

    const TOPIC_DEFINITIONS = [
      {
        id: "asbestos_exclusion",
        label: "Asbestos Exclusion",
        overview: "Placeholder asbestos exclusion wording sample.",
        minStartYear: 1985,
        templates: [
          "Policy {policyNumber} issued by {carrier} ({program}) states that coverage does not apply to loss, cost, or expense arising out of asbestos, asbestos fibers, or asbestos-containing materials.",
          "For policy {policyNumber}, asbestos-related bodily injury or property damage is excluded, including claims based on alleged inhalation, ingestion, or exposure.",
          "Policy {policyNumber} includes asbestos wording that references investigation, testing, monitoring, removal, treatment, and disposal activities."
        ]
      },
      {
        id: "pollution_exclusion",
        label: "Pollution Exclusion",
        overview: "Placeholder pollution exclusion wording sample.",
        templates: [
          "Policy {policyNumber} includes a pollution exclusion for actual, alleged, or threatened discharge, dispersal, release, or escape of pollutants.",
          "For policy {policyNumber}, pollution-related claims are excluded, including governmental direction to monitor, clean up, remove, contain, treat, or neutralize pollutants.",
          "Policy {policyNumber} applies pollution exclusion wording to bodily injury or property damage arising from pollutant conditions."
        ]
      },
      {
        id: "notice_reporting",
        label: "Notice and Reporting",
        overview: "Placeholder notice and reporting wording sample.",
        templates: [
          "Policy {policyNumber} contains notice and reporting language requiring written notice of occurrence and claim within stated timelines.",
          "For policy {policyNumber}, notice timing is stated as a condition to coverage and references reporting obligations during the policy period.",
          "Policy {policyNumber} includes reporting language tied to claim submission and compliance with notice conditions."
        ]
      },
      {
        id: "non_cumulation",
        label: "Non-Cumulation / Prior Insurance",
        overview: "Placeholder non-cumulation and prior-insurance wording sample.",
        templates: [
          "Policy {policyNumber} includes non-cumulation/prior-insurance wording that references reduction of limits by amounts paid under prior insurance for the same loss.",
          "For policy {policyNumber}, the non-cumulation clause addresses interaction between current limits and prior policy payments.",
          "Policy {policyNumber} includes prior-insurance language concerning continuous or repeated exposure and potential limit reduction."
        ]
      }
    ];

    function fillExcerptTemplate(template, policy) {
      return String(template || "")
        .replace(/\{policyNumber\}/g, String(policy.policyNumber || `Policy ${policy.policyId || ""}`))
        .replace(/\{carrier\}/g, String(policy.carrier || "Unknown Carrier"))
        .replace(/\{program\}/g, String(policy.insuranceProgram || "Unknown Program"));
    }

    function getPolicyStartYear(policy) {
      const startDate = String(policy?.startDate || "").trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(startDate)) return Number(startDate.slice(0, 4));
      const annual = Number(String(policy?.annualPeriod || "").trim());
      return Number.isFinite(annual) ? annual : NaN;
    }

    function pickRepresentativeLimit(limitRows) {
      if (!Array.isArray(limitRows) || !limitRows.length) {
        return { attachmentPoint: 0, perOccLimit: 0, aggregateLimit: 0, layerLimit: 0 };
      }
      return [...limitRows].sort((a, b) => {
        if (Number(b.layerLimit || 0) !== Number(a.layerLimit || 0)) return Number(b.layerLimit || 0) - Number(a.layerLimit || 0);
        if (Number(a.attachmentPoint || 0) !== Number(b.attachmentPoint || 0)) return Number(a.attachmentPoint || 0) - Number(b.attachmentPoint || 0);
        return Number(b.perOccLimit || 0) - Number(a.perOccLimit || 0);
      })[0];
    }

    function buildPolicyRowsFromDataset({ policyRows, dateRows, limitRows, carrierRows, carrierGroupRows, programRows }) {
      const dateByPolicyId = new Map(dateRows.map((row) => [String(getBy(row, "PolicyID")).trim(), row]));
      const carrierById = new Map(carrierRows.map((row) => [String(getBy(row, "CarrierID")).trim(), row]));
      const carrierGroupById = new Map(carrierGroupRows.map((row) => [String(getBy(row, "CarrierGroupID")).trim(), row]));
      const programById = new Map(programRows.map((row) => [String(getBy(row, "InsuranceProgramID")).trim(), row]));

      const limitsByPolicyId = new Map();
      for (const row of limitRows) {
        const policyId = String(getBy(row, "PolicyID")).trim();
        if (!policyId) continue;
        if (!limitsByPolicyId.has(policyId)) limitsByPolicyId.set(policyId, []);
        limitsByPolicyId.get(policyId).push({
          attachmentPoint: toNum(getBy(row, "AttachmentPoint", "Attachment Point")),
          perOccLimit: toNum(getBy(row, "PerOccLimit", "Per Occ Limit")),
          aggregateLimit: toNum(getBy(row, "AggregateLimit", "Aggregate Limit")),
          layerLimit: toNum(getBy(row, "LayerPerOccLimit", "Layer Per Occ Limit", "PerOccLimit", "Per Occ Limit"))
        });
      }

      const policies = [];
      for (const row of policyRows) {
        const policyId = String(getBy(row, "PolicyID")).trim();
        if (!policyId) continue;

        const dateRow = dateByPolicyId.get(policyId) || {};
        const carrierRow = carrierById.get(String(getBy(row, "CarrierID")).trim()) || {};
        const carrierGroupId = String(
          getBy(carrierRow, "CarrierGroupID", "Carrier Group ID", "CarrierGroupId") ||
          getBy(row, "CarrierGroupID", "Carrier Group ID")
        ).trim();
        const carrierGroupRow = carrierGroupById.get(carrierGroupId) || {};
        const programRow = programById.get(String(getBy(row, "InsuranceProgramID")).trim()) || {};
        const representativeLimit = pickRepresentativeLimit(limitsByPolicyId.get(policyId) || []);

        const startDate = normalizeISODate(getBy(dateRow, "PStartDate", "PolicyStartDate", "StartDate") || getBy(row, "MinPStartDate"));
        const endDate = normalizeISODate(getBy(dateRow, "PEndDate", "PolicyEndDate", "EndDate"));
        const annualPeriod = String(getBy(dateRow, "AnnualPeriod") || (startDate ? startDate.slice(0, 4) : "")).trim();

        policies.push({
          policyId,
          policyNumber: String(getBy(row, "PolicyNum", "PolicyNo", "policy_no", "Policy Number") || `Policy ${policyId}`).trim(),
          carrier: String(getBy(carrierRow, "CarrierName", "Carrier") || "(unknown carrier)").trim(),
          carrierGroup: String(getBy(carrierGroupRow, "CarrierGroupName", "CarrierGroup", "Carrier Group") || "(unknown group)").trim(),
          insuranceProgram: String(getBy(programRow, "InsuranceProgram", "Program", "Name") || "(unknown program)").trim(),
          startDate,
          endDate,
          annualPeriod,
          sirPerOcc: toNum(getBy(row, "SIRPerOcc", "SIR Per Occ", "SIR")),
          attachmentPoint: Number(representativeLimit.attachmentPoint || 0),
          perOccLimit: Number(representativeLimit.perOccLimit || 0),
          aggregateLimit: Number(representativeLimit.aggregateLimit || 0),
          layerLimit: Number(representativeLimit.layerLimit || 0)
        });
      }

      return policies.sort((a, b) =>
        String(a.startDate || "").localeCompare(String(b.startDate || "")) ||
        String(a.carrier || "").localeCompare(String(b.carrier || "")) ||
        String(a.policyNumber || "").localeCompare(String(b.policyNumber || ""))
      );
    }

    function buildLanguageTopicsFromPolicies(policies) {
      return TOPIC_DEFINITIONS.map((topic) => {
        const eligiblePolicies = policies.filter((policy) => {
          if (!Number.isFinite(topic.minStartYear)) return true;
          const startYear = getPolicyStartYear(policy);
          return Number.isFinite(startYear) && startYear >= topic.minStartYear;
        });

        const entries = eligiblePolicies.map((policy, index) =>
          makeEntry({
            ...policy,
            languageTopicId: topic.id,
            languageTopicLabel: topic.label,
            excerpt: fillExcerptTemplate(topic.templates[index % topic.templates.length], policy)
          })
        );
        return {
          id: topic.id,
          label: topic.label,
          overview: topic.overview,
          entries
        };
      });
    }

    async function loadLanguageTopicsFromDataset() {
      const [policyRows, dateRows, limitRows, carrierRows, carrierGroupRows, programRows] = await Promise.all([
        fetchCSV("/data/OriginalFiles/tblPolicy.csv"),
        fetchCSV("/data/OriginalFiles/tblPolicyDates.csv"),
        fetchCSV("/data/OriginalFiles/tblPolicyLimits.csv"),
        fetchCSV("/data/OriginalFiles/tblCarrier.csv"),
        fetchCSV("/data/OriginalFiles/tblCarrierGroup.csv"),
        fetchCSV("/data/OriginalFiles/tblInsuranceProgram.csv")
      ]);
      const policies = buildPolicyRowsFromDataset({
        policyRows,
        dateRows,
        limitRows,
        carrierRows,
        carrierGroupRows,
        programRows
      });
      return buildLanguageTopicsFromPolicies(policies);
    }

    function setLanguageExportStatus(message, isError = false) {
      languageExportStatus.textContent = message || "";
      languageExportStatus.classList.toggle("isError", Boolean(message) && isError);
    }

    function uniqueCount(values) {
      return new Set(values.map((v) => String(v || "").trim()).filter(Boolean)).size;
    }

    function getTopicById(topicId) {
      return (
        LANGUAGE_TOPICS.find((topic) => topic.id === topicId) ||
        LANGUAGE_TOPICS[0] ||
        { id: "", label: SUMMARY_ALL_TOPICS_LABEL, summaryLabel: SUMMARY_ALL_TOPICS_LABEL, entries: [] }
      );
    }

    function getTopicContext(topicId) {
      const selectedId = String(topicId || "").trim();
      if (selectedId === ALL_TOPICS_ID) {
        return {
          id: ALL_TOPICS_ID,
          label: ALL_TOPICS_LABEL,
          summaryLabel: SUMMARY_ALL_TOPICS_LABEL,
          entries: LANGUAGE_TOPICS.flatMap((topic) => topic.entries)
        };
      }
      const topic = getTopicById(selectedId);
      return {
        id: topic.id,
        label: topic.label,
        summaryLabel: topic.label,
        entries: topic.entries
      };
    }

    function fillTopicOptions() {
      languageTopicSelect.innerHTML = "";
      const allOpt = document.createElement("option");
      allOpt.value = ALL_TOPICS_ID;
      allOpt.textContent = ALL_TOPICS_LABEL;
      languageTopicSelect.appendChild(allOpt);
      for (const topic of LANGUAGE_TOPICS) {
        const opt = document.createElement("option");
        opt.value = topic.id;
        opt.textContent = topic.label;
        languageTopicSelect.appendChild(opt);
      }
    }

    function fillProgramOptions(topic, keepCurrent = false) {
      const prior = keepCurrent ? String(languageProgramSelect.value || "") : "";
      const programs = Array.from(
        new Set(topic.entries.map((entry) => String(entry.insuranceProgram || "").trim()).filter(Boolean))
      ).sort((a, b) => a.localeCompare(b));

      languageProgramSelect.innerHTML = "";
      const allOpt = document.createElement("option");
      allOpt.value = "";
      allOpt.textContent = "All programs";
      languageProgramSelect.appendChild(allOpt);

      for (const program of programs) {
        const opt = document.createElement("option");
        opt.value = program;
        opt.textContent = program;
        languageProgramSelect.appendChild(opt);
      }

      if (prior && programs.includes(prior)) {
        languageProgramSelect.value = prior;
      } else {
        languageProgramSelect.value = "";
      }
    }

    function dateInputToMs(value, endOfDay = false) {
      if (!value) return endOfDay ? Number.POSITIVE_INFINITY : Number.NEGATIVE_INFINITY;
      const stamp = Date.parse(`${value}T${endOfDay ? "23:59:59" : "00:00:00"}Z`);
      return Number.isFinite(stamp) ? stamp : endOfDay ? Number.POSITIVE_INFINITY : Number.NEGATIVE_INFINITY;
    }

    function getFilteredEntries(topic) {
      const selectedProgram = String(languageProgramSelect.value || "").trim();
      const startFromMs = dateInputToMs(languageStartDateFilter.value, false);
      const startToMs = dateInputToMs(languageEndDateFilter.value, true);
      const minLayer = String(languageMinLayerLimitFilter.value || "").trim() === ""
        ? Number.NEGATIVE_INFINITY
        : toNum(languageMinLayerLimitFilter.value);
      const maxLayer = String(languageMaxLayerLimitFilter.value || "").trim() === ""
        ? Number.POSITIVE_INFINITY
        : toNum(languageMaxLayerLimitFilter.value);

      return topic.entries.filter((entry) => {
        if (selectedProgram && entry.insuranceProgram !== selectedProgram) return false;
        if (entry.startMs < startFromMs || entry.startMs > startToMs) return false;
        if (Number(entry.layerLimit || 0) < minLayer || Number(entry.layerLimit || 0) > maxLayer) return false;
        return true;
      });
    }

    function sortEntries(entries) {
      return [...entries].sort((a, b) =>
        String(a.carrier).localeCompare(String(b.carrier)) ||
        String(a.startDate).localeCompare(String(b.startDate)) ||
        String(a.languageTopicLabel).localeCompare(String(b.languageTopicLabel)) ||
        String(a.insuranceProgram).localeCompare(String(b.insuranceProgram)) ||
        String(a.policyNumber).localeCompare(String(b.policyNumber))
      );
    }

    function summarizeList(values, maxVisible = 2) {
      const items = (values || []).filter(Boolean);
      if (!items.length) return "N/A";
      if (items.length <= maxVisible) return items.join(", ");
      return `${items.slice(0, maxVisible).join(", ")} +${items.length - maxVisible} more`;
    }

    function escapeHtml(value) {
      return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }

    function renderBulletedCell(items) {
      const values = (items || []).filter(Boolean);
      if (!values.length) return `<span class="languageListInline">N/A</span>`;
      if (values.length === 1) {
        return `<span class="languageListInline">${escapeHtml(values[0])}</span>`;
      }
      const listItems = values.map((value) => `<li>${escapeHtml(value)}</li>`).join("");
      return `<ul class="languageListBullets">${listItems}</ul>`;
    }

    function groupEntriesForSchedule(entries) {
      const byPolicy = new Map();
      for (const entry of sortEntries(entries)) {
        const policyId = String(entry?.policyId || "").trim();
        const fallbackKey = [
          String(entry?.policyNumber || "").trim(),
          String(entry?.startDate || "").trim(),
          String(entry?.endDate || "").trim(),
          String(entry?.carrier || "").trim()
        ].join("|");
        const key = policyId || fallbackKey;
        if (!key) continue;

        if (!byPolicy.has(key)) {
          byPolicy.set(key, {
            policyId,
            policyNumber: String(entry.policyNumber || "").trim(),
            carrier: String(entry.carrier || "").trim(),
            startDate: String(entry.startDate || "").trim(),
            endDate: String(entry.endDate || "").trim(),
            annualPeriod: String(entry.annualPeriod || "").trim(),
            sirPerOcc: Number(entry.sirPerOcc || 0),
            attachmentPoint: Number(entry.attachmentPoint || 0),
            perOccLimit: Number(entry.perOccLimit || 0),
            aggregateLimit: Number(entry.aggregateLimit || 0),
            layerLimit: Number(entry.layerLimit || 0),
            programs: new Set(),
            topics: new Set(),
            entries: []
          });
        }

        const group = byPolicy.get(key);
        group.sirPerOcc = Math.max(group.sirPerOcc, Number(entry.sirPerOcc || 0));
        group.attachmentPoint = Math.max(group.attachmentPoint, Number(entry.attachmentPoint || 0));
        group.perOccLimit = Math.max(group.perOccLimit, Number(entry.perOccLimit || 0));
        group.aggregateLimit = Math.max(group.aggregateLimit, Number(entry.aggregateLimit || 0));
        group.layerLimit = Math.max(group.layerLimit, Number(entry.layerLimit || 0));
        group.programs.add(String(entry.insuranceProgram || "").trim());
        group.topics.add(String(entry.languageTopicLabel || "").trim());
        group.entries.push(entry);
      }

      return Array.from(byPolicy.values())
        .map((group) => ({
          ...group,
          programList: Array.from(group.programs).filter(Boolean).sort((a, b) => a.localeCompare(b)),
          topicList: Array.from(group.topics).filter(Boolean).sort((a, b) => a.localeCompare(b))
        }))
        .sort((a, b) =>
          String(a.startDate).localeCompare(String(b.startDate)) ||
          String(a.carrier).localeCompare(String(b.carrier)) ||
          String(a.policyNumber).localeCompare(String(b.policyNumber))
        );
    }

    function normalizeSortText(value) {
      return String(value || "").trim().toLowerCase();
    }

    function sortNumber(value) {
      const num = Number(value);
      return Number.isFinite(num) ? num : 0;
    }

    function sortDateStamp(isoDate) {
      const stamp = Date.parse(`${String(isoDate || "").trim()}T00:00:00Z`);
      return Number.isFinite(stamp) ? stamp : 0;
    }

    function scheduleSortValue(row, key) {
      switch (key) {
        case "carrier":
          return normalizeSortText(row?.carrier);
        case "policyNumber":
          return normalizeSortText(row?.policyNumber);
        case "languageTopics":
          return normalizeSortText((row?.topicList || []).join(" | "));
        case "programs":
          return normalizeSortText((row?.programList || []).join(" | "));
        case "startDate":
          return sortDateStamp(row?.startDate);
        case "endDate":
          return sortDateStamp(row?.endDate);
        case "annual":
          return sortNumber(row?.annualPeriod);
        case "sir":
          return sortNumber(row?.sirPerOcc);
        case "attachment":
          return sortNumber(row?.attachmentPoint);
        case "perOcc":
          return sortNumber(row?.perOccLimit);
        case "aggregate":
          return sortNumber(row?.aggregateLimit);
        case "layerLimit":
          return sortNumber(row?.layerLimit);
        case "language":
          return sortNumber((row?.entries || []).length);
        default:
          return "";
      }
    }

    function sortScheduleRows(rows, sortState = languageScheduleSortState) {
      const direction = sortState?.direction === "desc" ? -1 : 1;
      return [...(rows || [])].sort((a, b) => {
        const primaryA = scheduleSortValue(a, sortState?.key);
        const primaryB = scheduleSortValue(b, sortState?.key);

        if (typeof primaryA === "number" && typeof primaryB === "number") {
          if (primaryA !== primaryB) return (primaryA - primaryB) * direction;
        } else {
          const textCompare = String(primaryA).localeCompare(String(primaryB));
          if (textCompare !== 0) return textCompare * direction;
        }

        return (
          sortDateStamp(a?.startDate) - sortDateStamp(b?.startDate) ||
          normalizeSortText(a?.carrier).localeCompare(normalizeSortText(b?.carrier)) ||
          normalizeSortText(a?.policyNumber).localeCompare(normalizeSortText(b?.policyNumber))
        );
      });
    }

    function renderSummary(topic, entries) {
      const policyCount = uniqueCount(entries.map((entry) => entry.policyNumber));
      const carrierCount = uniqueCount(entries.map((entry) => entry.carrier));
      const carrierGroupCount = uniqueCount(entries.map((entry) => entry.carrierGroup));

      const sortedByStart = [...entries].sort((a, b) => a.startMs - b.startMs);
      const firstStart = sortedByStart[0]?.startDate || "";
      const lastStart = sortedByStart[sortedByStart.length - 1]?.startDate || "";

      languageTopicTitle.textContent = topic.summaryLabel;
      languageExcerptCount.textContent = `Excerpts: ${entries.length}`;

      languageSummaryStats.innerHTML = `
        <article class="languageStatCard">
          <div class="languageStatLabel">Policies</div>
          <div class="languageStatValue">${policyCount}</div>
        </article>
        <article class="languageStatCard">
          <div class="languageStatLabel">Carriers</div>
          <div class="languageStatValue">${carrierCount}</div>
        </article>
        <article class="languageStatCard">
          <div class="languageStatLabel">Carrier Groups</div>
          <div class="languageStatValue">${carrierGroupCount}</div>
        </article>
        <article class="languageStatCard">
          <div class="languageStatLabel">First Policy Start</div>
          <div class="languageStatValue">${firstStart ? formatDate(firstStart) : "N/A"}</div>
        </article>
        <article class="languageStatCard">
          <div class="languageStatLabel">Last Policy Start</div>
          <div class="languageStatValue">${lastStart ? formatDate(lastStart) : "N/A"}</div>
        </article>
      `;
    }

    function renderSchedule(entries) {
      const rows = sortScheduleRows(groupEntriesForSchedule(entries));
      CURRENT_SCHEDULE_ROWS = rows;
      languageScheduleCount.textContent = `Rows: ${rows.length} (Grouped from ${entries.length} excerpts)`;
      languageScheduleBody.innerHTML = "";

      if (!rows.length) {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td colspan="13">No policies match the selected filters.</td>`;
        languageScheduleBody.appendChild(tr);
        return;
      }

      rows.forEach((row, index) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${row.carrier}</td>
          <td>${row.policyNumber}</td>
          <td>${renderBulletedCell(row.topicList)}</td>
          <td>${renderBulletedCell(row.programList)}</td>
          <td>${formatDate(row.startDate)}</td>
          <td>${formatDate(row.endDate)}</td>
          <td>${row.annualPeriod}</td>
          <td>${money(row.sirPerOcc)}</td>
          <td>${money(row.attachmentPoint)}</td>
          <td>${money(row.perOccLimit)}</td>
          <td>${money(row.aggregateLimit)}</td>
          <td>${money(row.layerLimit)}</td>
          <td><button type="button" class="languageViewBtn" data-language-index="${index}">View All</button></td>
        `;
        languageScheduleBody.appendChild(tr);
      });
    }

    function renderExcerpts(entries) {
      const rows = sortEntries(entries);
      languageExcerptList.innerHTML = "";

      if (!rows.length) {
        const empty = document.createElement("div");
        empty.className = "languageEmpty";
        empty.textContent = "No excerpts match the selected filters.";
        languageExcerptList.appendChild(empty);
        return;
      }

      for (const entry of rows) {
        const card = document.createElement("article");
        card.className = "languageExcerptCard";
        card.innerHTML = `
          <div class="languageExcerptMeta">
            <strong>${entry.policyNumber}</strong>
            <span>${entry.languageTopicLabel} | ${entry.carrier} | ${entry.insuranceProgram} | ${formatDate(entry.startDate)} - ${formatDate(entry.endDate)}</span>
          </div>
          <p class="languageQuote">${entry.excerpt}</p>
        `;
        languageExcerptList.appendChild(card);
      }
    }

    function getCurrentTopicAndEntries() {
      const topic = getTopicContext(languageTopicSelect.value);
      const entries = getFilteredEntries(topic);
      return { topic, entries };
    }

    function toPdfLanguageText(excerpt, policyNumber) {
      const policyId = String(policyNumber || "").trim();
      const escapedPolicyId = policyId.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      let text = String(excerpt || "").trim();
      if (!text) return "";

      if (escapedPolicyId) {
        text = text.replace(new RegExp(`^For\\s+policy\\s+${escapedPolicyId}\\s*,\\s*`, "i"), "");
        text = text.replace(new RegExp(`^Policy\\s+${escapedPolicyId}\\s+`, "i"), "");
      }

      // Fallback cleanup if policy number token varied.
      text = text.replace(/^For\s+policy\s+[^,]+,\s*/i, "");
      text = text.replace(/^Policy\s+[A-Z0-9-]+\s+/i, "");

      text = text.replace(/\s{2,}/g, " ").trim();
      if (!text) return String(excerpt || "").trim();
      return text.charAt(0).toUpperCase() + text.slice(1);
    }

    function renderAll() {
      setLanguageExportStatus("");
      const { topic, entries } = getCurrentTopicAndEntries();
      renderSummary(topic, entries);
      renderSchedule(entries);
      renderExcerpts(entries);
    }

    function openLanguageModal(row) {
      if (!row || !languageModal) return;
      languageModalMeta.textContent = `${row.policyNumber} | ${row.carrier} | ${formatDate(row.startDate)} - ${formatDate(row.endDate)} | Programs: ${summarizeList(row.programList, 10)}`;

      const lines = [];
      const sortedEntries = [...(row.entries || [])].sort((a, b) =>
        String(a.languageTopicLabel || "").localeCompare(String(b.languageTopicLabel || "")) ||
        String(a.insuranceProgram || "").localeCompare(String(b.insuranceProgram || ""))
      );

      for (const entry of sortedEntries) {
        lines.push(`${entry.languageTopicLabel} | ${entry.insuranceProgram}`);
        lines.push(String(entry.excerpt || ""));
        lines.push("");
      }

      languageModalBody.textContent = lines.length ? lines.join("\n") : "No language available for this row.";
      languageModal.hidden = false;
      document.body.classList.add("modalOpen");
      languageModalCloseBtn?.focus();
    }

    function closeLanguageModal() {
      if (!languageModal) return;
      languageModal.hidden = true;
      document.body.classList.remove("modalOpen");
    }

    function groupEntriesByPolicy(entries) {
      const rows = sortEntries(entries);
      const byPolicy = new Map();
      for (const entry of rows) {
        const key = String(entry?.policyId || entry?.policyNumber || "").trim();
        if (!key) continue;
        if (!byPolicy.has(key)) {
          byPolicy.set(key, {
            policyId: String(entry.policyId || "").trim(),
            policyNumber: String(entry.policyNumber || "").trim(),
            carrier: String(entry.carrier || "").trim(),
            insuranceProgram: String(entry.insuranceProgram || "").trim(),
            startDate: String(entry.startDate || "").trim(),
            endDate: String(entry.endDate || "").trim(),
            annualPeriod: String(entry.annualPeriod || "").trim(),
            sirPerOcc: Number(entry.sirPerOcc || 0),
            attachmentPoint: Number(entry.attachmentPoint || 0),
            perOccLimit: Number(entry.perOccLimit || 0),
            aggregateLimit: Number(entry.aggregateLimit || 0),
            layerLimit: Number(entry.layerLimit || 0),
            entries: []
          });
        }
        byPolicy.get(key).entries.push(entry);
      }

      return Array.from(byPolicy.values()).sort((a, b) =>
        String(a.startDate || "").localeCompare(String(b.startDate || "")) ||
        String(a.carrier || "").localeCompare(String(b.carrier || "")) ||
        String(a.policyNumber || "").localeCompare(String(b.policyNumber || ""))
      );
    }

    async function exportLanguageReport() {
      const { topic, entries } = getCurrentTopicAndEntries();
      if (!entries.length) {
        setLanguageExportStatus("No rows match current filters. Nothing to export.", true);
        return;
      }

      const previousLabel = exportLanguageReportBtn.textContent;
      exportLanguageReportBtn.disabled = true;
      exportLanguageReportBtn.textContent = "Exporting...";
      setLanguageExportStatus("Building policy language report PDF...");

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

        const filterProgram = String(languageProgramSelect.value || "").trim() || "All programs";
        const filterStartFrom = String(languageStartDateFilter.value || "").trim() || "All";
        const filterStartTo = String(languageEndDateFilter.value || "").trim() || "All";
        const filterMinLayer = String(languageMinLayerLimitFilter.value || "").trim() || "All";
        const filterMaxLayer = String(languageMaxLayerLimitFilter.value || "").trim() || "All";
        const groupedPolicies = groupEntriesByPolicy(entries);

        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(20);
        pdf.setTextColor(16, 24, 39);
        pdf.text("Policy Language Analysis Report", margin, y);
        y += 22;

        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(10);
        pdf.setTextColor(90, 100, 120);
        pdf.text(`Generated: ${new Date().toLocaleString()}`, margin, y);
        y += 14;

        const summaryRows = [
          `Topic filter: ${topic.summaryLabel}`,
          `Program filter: ${filterProgram}`,
          `Policy start-date filter: ${filterStartFrom} to ${filterStartTo}`,
          `Layer limit filter: ${filterMinLayer} to ${filterMaxLayer}`,
          `Rows exported: ${entries.length}`,
          `Policies exported: ${groupedPolicies.length}`
        ];
        summaryRows.forEach((line) => {
          pdf.text(line, margin, y);
          y += 12;
        });
        y += 4;

        for (const policy of groupedPolicies) {
          drawSectionHeading(`Policy ${policy.policyNumber || policy.policyId || "N/A"}`);

          const infoRows = [
            ["Policy ID", policy.policyId || "N/A"],
            ["Carrier", policy.carrier || "N/A"],
            ["Insurance Program", policy.insuranceProgram || "N/A"],
            ["Policy Period", `${formatDate(policy.startDate)} - ${formatDate(policy.endDate)}`],
            ["Annual Period", policy.annualPeriod || "N/A"],
            ["SIR", money(policy.sirPerOcc)],
            ["Attachment", money(policy.attachmentPoint)],
            ["Per Occ", money(policy.perOccLimit)],
            ["Aggregate", money(policy.aggregateLimit)],
            ["Layer Limit", money(policy.layerLimit)]
          ];

          const columnGap = 18;
          const columnWidth = (contentW - columnGap) / 2;
          for (let i = 0; i < infoRows.length; i += 2) {
            const left = infoRows[i];
            const right = infoRows[i + 1];
            const leftLines = pdf.splitTextToSize(`${left[0]}: ${left[1]}`, columnWidth);
            const rightLines = right ? pdf.splitTextToSize(`${right[0]}: ${right[1]}`, columnWidth) : [];
            const rowLines = Math.max(leftLines.length, rightLines.length || 1);
            const rowHeight = rowLines * 12 + 3;
            if (ensurePageSpace(rowHeight + 2)) drawSectionHeading(`Policy ${policy.policyNumber || policy.policyId || "N/A"} (cont.)`);
            pdf.setFont("helvetica", "normal");
            pdf.setFontSize(10);
            pdf.setTextColor(36, 46, 66);
            leftLines.forEach((line, idx) => pdf.text(line, margin, y + idx * 12));
            rightLines.forEach((line, idx) => pdf.text(line, margin + columnWidth + columnGap, y + idx * 12));
            y += rowHeight;
          }
          y += 4;

          drawWrappedText("Language Excerpts:", { fontStyle: "bold", fontSize: 11, lineHeight: 12, gapAfter: 4 });
          for (const entry of policy.entries) {
            const metaLine = `${entry.languageTopicLabel} | ${entry.carrier} | ${entry.insuranceProgram}`;
            drawWrappedText(metaLine, { fontStyle: "bold", fontSize: 10, lineHeight: 12, gapAfter: 2 });
            drawWrappedText(toPdfLanguageText(entry.excerpt, entry.policyNumber), { x: margin + 6, width: contentW - 6, fontSize: 10, lineHeight: 13, gapAfter: 6 });
          }
        }

        const pageCount = pdf.getNumberOfPages();
        for (let pageNum = 1; pageNum <= pageCount; pageNum++) {
          pdf.setPage(pageNum);
          pdf.setFont("helvetica", "normal");
          pdf.setFontSize(9);
          pdf.setTextColor(116, 124, 139);
          pdf.text("Coverage Dashboard Policy Language Analysis", margin, pageH - 12);
          pdf.text(`Page ${pageNum} of ${pageCount}`, pageW - margin, pageH - 12, { align: "right" });
        }

        const fileName = `PolicyLanguageAnalysis_${sanitizeFilePart(topic.summaryLabel, "all_topics")}_${toDateStamp()}.pdf`;
        pdf.save(fileName);
        setLanguageExportStatus(`Policy language report exported: ${fileName}`);
      } catch (err) {
        console.error(err);
        setLanguageExportStatus(`Failed to export report: ${err?.message || err}`, true);
      } finally {
        exportLanguageReportBtn.textContent = previousLabel;
        exportLanguageReportBtn.disabled = false;
      }
    }

    function resetFilters() {
      languageProgramSelect.value = "";
      languageStartDateFilter.value = "";
      languageEndDateFilter.value = "";
      languageMinLayerLimitFilter.value = "";
      languageMaxLayerLimitFilter.value = "";
      setLanguageExportStatus("");
      renderAll();
    }

    function applyQueryFiltersFromURL() {
      const params = new URLSearchParams(window.location.search);

      const topicId = String(params.get("topic") || "").trim();
      if (topicId === ALL_TOPICS_ID || (topicId && LANGUAGE_TOPICS.some((topic) => topic.id === topicId))) {
        languageTopicSelect.value = topicId;
      }

      fillProgramOptions(getTopicContext(languageTopicSelect.value), false);

      const program = String(params.get("program") || "").trim();
      if (program && Array.from(languageProgramSelect.options).some((opt) => opt.value === program)) {
        languageProgramSelect.value = program;
      }

      const startFrom = String(params.get("startFrom") || "").trim();
      const startTo = String(params.get("startTo") || "").trim();
      if (/^\\d{4}-\\d{2}-\\d{2}$/.test(startFrom)) languageStartDateFilter.value = startFrom;
      if (/^\\d{4}-\\d{2}-\\d{2}$/.test(startTo)) languageEndDateFilter.value = startTo;

      const minLayerLimit = String(params.get("minLayerLimit") || "").trim();
      const maxLayerLimit = String(params.get("maxLayerLimit") || "").trim();
      if (minLayerLimit) languageMinLayerLimitFilter.value = minLayerLimit;
      if (maxLayerLimit) languageMaxLayerLimitFilter.value = maxLayerLimit;

    }

    async function init() {
      applyThemeToPage(getPreferredTheme(THEME_STORAGE_KEY), { themeLabelEl: themeLabel, themeToggleBtn });
      themeToggleBtn.addEventListener("click", () => {
        const current = document.documentElement.dataset.theme === "light" ? "light" : "dark";
        const next = current === "light" ? "dark" : "light";
        localStorage.setItem(THEME_STORAGE_KEY, next);
        applyThemeToPage(next, { themeLabelEl: themeLabel, themeToggleBtn });
      });

      LANGUAGE_TOPICS = await loadLanguageTopicsFromDataset();
      fillTopicOptions();
      languageTopicSelect.value = ALL_TOPICS_ID;
      applyQueryFiltersFromURL();

      setupSortableTable({
        table: languageScheduleTable,
        columns: [
          { index: 0, key: "carrier", defaultDirection: "asc" },
          { index: 1, key: "policyNumber", defaultDirection: "asc" },
          { index: 2, key: "languageTopics", defaultDirection: "asc" },
          { index: 3, key: "programs", defaultDirection: "asc" },
          { index: 4, key: "startDate", defaultDirection: "asc" },
          { index: 5, key: "endDate", defaultDirection: "asc" },
          { index: 6, key: "annual", defaultDirection: "asc" },
          { index: 7, key: "sir", defaultDirection: "asc" },
          { index: 8, key: "attachment", defaultDirection: "asc" },
          { index: 9, key: "perOcc", defaultDirection: "asc" },
          { index: 10, key: "aggregate", defaultDirection: "asc" },
          { index: 11, key: "layerLimit", defaultDirection: "asc" },
          { index: 12, key: "language", defaultDirection: "asc" }
        ],
        initialSort: languageScheduleSortState,
        onSortChange: (nextSortState) => {
          languageScheduleSortState = nextSortState;
          const { entries } = getCurrentTopicAndEntries();
          renderSchedule(entries);
        }
      });

      renderAll();

      languageTopicSelect.addEventListener("change", () => {
        fillProgramOptions(getTopicContext(languageTopicSelect.value), false);
        renderAll();
      });
      languageProgramSelect.addEventListener("change", renderAll);
      languageStartDateFilter.addEventListener("change", renderAll);
      languageEndDateFilter.addEventListener("change", renderAll);
      languageMinLayerLimitFilter.addEventListener("change", renderAll);
      languageMaxLayerLimitFilter.addEventListener("change", renderAll);
      exportLanguageReportBtn.addEventListener("click", exportLanguageReport);
      resetLanguageFiltersBtn.addEventListener("click", resetFilters);
      languageScheduleBody.addEventListener("click", (event) => {
        const btn = event.target.closest(".languageViewBtn");
        if (!btn) return;
        const idx = Number(btn.dataset.languageIndex);
        if (!Number.isInteger(idx) || idx < 0 || idx >= CURRENT_SCHEDULE_ROWS.length) return;
        openLanguageModal(CURRENT_SCHEDULE_ROWS[idx]);
      });
      languageModalCloseBtn?.addEventListener("click", closeLanguageModal);
      languageModal?.addEventListener("click", (event) => {
        if (event.target?.hasAttribute("data-close-language-modal")) closeLanguageModal();
      });
      document.addEventListener("keydown", (event) => {
        if (event.key === "Escape" && languageModal && !languageModal.hidden) closeLanguageModal();
      });

      for (const el of [languageMinLayerLimitFilter, languageMaxLayerLimitFilter]) {
        el.addEventListener("input", () => {
          window.clearTimeout(el._debounceTimer);
          el._debounceTimer = window.setTimeout(renderAll, 120);
        });
      }
    }

    init().catch((err) => {
      console.error(err);
      languageTopicTitle.textContent = SUMMARY_ALL_TOPICS_LABEL;
      languageExcerptCount.textContent = "Excerpts: 0";
      languageSummaryStats.innerHTML = "";
      languageScheduleCount.textContent = "Rows: 0";
      languageScheduleBody.innerHTML = '<tr><td colspan="12">Failed to load dataset-backed policy records.</td></tr>';
      languageExcerptList.innerHTML = '<div class="languageEmpty">Failed to load dataset-backed policy records.</div>';
      setLanguageExportStatus("Failed to load data. Report export unavailable.", true);
    });
  
