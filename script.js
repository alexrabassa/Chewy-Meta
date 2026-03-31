const fallbackRows = [
  {
    launchDate: "2026-04-01",
    endDate: "Live",
    funnel: "UF",
    versionCampaign: "Core",
    campaign: "Spring Sale Prospecting",
    adSet: "US Broad 25-44",
    adName: "Problem Solution | 15",
    spend: 1240.52,
    purchases: 18,
    revenue: 3985.44,
    roas: 3.21,
    status: "Live",
    notes: "Main cold traffic concept pushing hero offer."
  },
  {
    launchDate: "2026-04-03",
    endDate: "2026-04-03",
    funnel: "UF",
    versionCampaign: "Core",
    campaign: "Spring Sale Prospecting",
    adSet: "Interest Stack Wellness",
    adName: "Founder Story | 30",
    spend: 0,
    purchases: 0,
    revenue: 0,
    roas: 0,
    status: "Planned",
    notes: "Second angle for broad awareness testing."
  },
  {
    launchDate: "2026-04-05",
    endDate: "Live",
    funnel: "MF",
    versionCampaign: "Core",
    campaign: "Retarget 14 Day Engagers",
    adSet: "Video View Retargeting",
    adName: "Testimonial Montage | 20",
    spend: 860.23,
    purchases: 11,
    revenue: 2410.35,
    roas: 2.8,
    status: "Live",
    notes: "Warmer audience follow-up with testimonial stack."
  },
  {
    launchDate: "2026-04-06",
    endDate: "2026-04-06",
    funnel: "LF",
    versionCampaign: "Core",
    campaign: "Cart And Checkout Recovery",
    adSet: "ATC 7 Day",
    adName: "Offer Reminder | Static",
    spend: 420.12,
    purchases: 7,
    revenue: 1455.22,
    roas: 3.46,
    status: "Refreshing",
    notes: "Updating offer callout and CTA placement."
  },
  {
    launchDate: "2026-04-08",
    endDate: "2026-04-08",
    funnel: "Retention",
    versionCampaign: "Retention",
    campaign: "Repeat Purchase Push",
    adSet: "Past Buyers 90 Day",
    adName: "Bundle Upsell Carousel 02",
    spend: 0,
    purchases: 0,
    revenue: 0,
    roas: 0,
    status: "Paused",
    notes: "Paused while inventory catches up."
  }
];

const storageKey = "metaCreativeCalendarRows.v5";
const tableBody = document.getElementById("calendarTableBody");
const summaryGrid = document.getElementById("summaryGrid");
const funnelFilter = document.getElementById("funnelFilter");
const campaignFilter = document.getElementById("campaignFilter");
const adSetFilter = document.getElementById("adSetFilter");
const statusFilter = document.getElementById("statusFilter");
const creativeForm = document.getElementById("creativeForm");
const exportCsvButton = document.getElementById("exportCsvButton");
const sortableHeaders = document.querySelectorAll("th[data-sort-key]");
const tableViewTab = document.getElementById("tableViewTab");
const boardViewTab = document.getElementById("boardViewTab");
const previewViewTab = document.getElementById("previewViewTab");
const calendarViewTab = document.getElementById("calendarViewTab");
const tableViewPanel = document.getElementById("tableViewPanel");
const boardViewPanel = document.getElementById("boardViewPanel");
const previewViewPanel = document.getElementById("previewViewPanel");
const calendarViewPanel = document.getElementById("calendarViewPanel");
const boardViewContent = document.getElementById("boardViewContent");
const previewViewContent = document.getElementById("previewViewContent");
const calendarViewContent = document.getElementById("calendarViewContent");
const calendarPlanSelect = document.getElementById("calendarPlanSelect");
const calendarFreeFieldSelect = document.getElementById("calendarFreeFieldSelect");
const calendarPlanDate = document.getElementById("calendarPlanDate");
const calendarPlanSave = document.getElementById("calendarPlanSave");
const calendarPlanClear = document.getElementById("calendarPlanClear");
const dataFreshness = document.getElementById("dataFreshness");
const previewTooltip = document.getElementById("previewTooltip");

let sortState = {
  key: "launchDate",
  direction: "asc"
};
let currentView = "table";
let activePreviewTarget = null;
const importedMeta = window.metaCreativeData?.meta || {};
const plannedEndDatesKey = "metaCreativeCalendar.plannedEndDates.v1";
let plannedEndDates = loadPlannedEndDates();

const fiscalPeriods = [
  { name: "P01", start: "2026-02-02", end: "2026-03-01" },
  { name: "P02", start: "2026-03-02", end: "2026-04-05" },
  { name: "P03", start: "2026-04-06", end: "2026-05-03" },
  { name: "P04", start: "2026-05-04", end: "2026-05-31" },
  { name: "P05", start: "2026-06-01", end: "2026-07-05" },
  { name: "P06", start: "2026-07-06", end: "2026-08-02" },
  { name: "P07", start: "2026-08-03", end: "2026-08-30" },
  { name: "P08", start: "2026-08-31", end: "2026-10-04" },
  { name: "P09", start: "2026-10-05", end: "2026-11-01" },
  { name: "P10", start: "2026-11-02", end: "2026-11-29" },
  { name: "P11", start: "2026-11-30", end: "2027-01-03" },
  { name: "P12", start: "2027-01-04", end: "2027-01-31" }
];
const visibleFiscalPeriods = fiscalPeriods.slice(0, 5);

let rows = loadRows();
let selectedCalendarPlannerVersion = "";
let selectedCalendarFreeField = "";

function isVisibleRow(row) {
  return !String(row.campaign || "").includes("Core BAU");
}

function loadRows() {
  try {
    if (Array.isArray(window.metaCreativeSeedRows) && window.metaCreativeSeedRows.length) {
      return [...window.metaCreativeSeedRows];
    }
    const savedRows = localStorage.getItem(storageKey);
    if (savedRows) {
      return JSON.parse(savedRows);
    }
    return [...fallbackRows];
  } catch (error) {
    if (Array.isArray(window.metaCreativeSeedRows) && window.metaCreativeSeedRows.length) {
      return [...window.metaCreativeSeedRows];
    }
    return [...fallbackRows];
  }
}

function loadPlannedEndDates() {
  try {
    const saved = localStorage.getItem(plannedEndDatesKey);
    return saved ? JSON.parse(saved) : {};
  } catch (error) {
    return {};
  }
}

function saveRows() {
  localStorage.setItem(storageKey, JSON.stringify(rows));
}

function savePlannedEndDates() {
  localStorage.setItem(plannedEndDatesKey, JSON.stringify(plannedEndDates));
}

function uniqueValues(key) {
  return [...new Set(rows.filter(isVisibleRow).map((row) => row[key]))].sort();
}

function getVisibleAdSets() {
  return [
    ...new Set(
      rows
        .filter(isVisibleRow)
        .filter((row) => funnelFilter.value === "All" || row.funnel === funnelFilter.value)
        .filter((row) => campaignFilter.value === "All" || row.campaign === campaignFilter.value)
        .map((row) => row.adSet)
    )
  ].sort();
}

function getVisibleCampaigns() {
  return [
    ...new Set(
      rows
        .filter(isVisibleRow)
        .filter((row) => funnelFilter.value === "All" || row.funnel === funnelFilter.value)
        .map((row) => row.campaign)
    )
  ].sort();
}

function populateFilter(selectElement, values, label) {
  const current = selectElement.value;
  selectElement.innerHTML = `<option value="All">All ${label}</option>`;

  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    selectElement.appendChild(option);
  });

  selectElement.value = values.includes(current) || current === "All" ? current : "All";
}

function getFilteredRows() {
  return rows.filter((row) => {
    const visibleMatch = isVisibleRow(row);
    const funnelMatch = funnelFilter.value === "All" || row.funnel === funnelFilter.value;
    const campaignMatch = campaignFilter.value === "All" || row.campaign === campaignFilter.value;
    const adSetMatch = adSetFilter.value === "All" || row.adSet === adSetFilter.value;
    const statusMatch = statusFilter.value === "All" || row.status === statusFilter.value;
    return visibleMatch && funnelMatch && campaignMatch && adSetMatch && statusMatch;
  });
}

function statusTagClass(status) {
  return {
    Live: "tag tag-live",
    Off: "tag tag-off",
    Planned: "tag tag-planned",
    Paused: "tag tag-paused",
    Refreshing: "tag tag-refreshing"
  }[status] || "tag";
}

function renderSummary(filteredRows) {
  const liveCount = filteredRows.filter((row) => row.status === "Live").length;
  const adSetCount = new Set(filteredRows.map((row) => row.adSet)).size;
  const uniqueAdsCount = new Set(filteredRows.map((row) => row.adName)).size;
  const totalSpend = filteredRows.reduce((sum, row) => sum + Number(row.spend || 0), 0);
  const startDate = filteredRows.length
    ? filteredRows.reduce(
        (earliest, row) => (!earliest || row.launchDate < earliest ? row.launchDate : earliest),
        ""
      )
    : null;

  const cards = [
    { label: "Live Ads", value: liveCount },
    { label: "Ad Sets In View", value: adSetCount },
    { label: "Unique Ads", value: formatNumber(uniqueAdsCount) },
    {
      label: startDate ? `Spend In View (Start Date ${startDate})` : "Spend In View",
      value: formatCurrency(totalSpend)
    }
  ];

  summaryGrid.innerHTML = "";

  cards.forEach((card) => {
    const element = document.createElement("article");
    element.className = "summary-card";
    element.innerHTML = `
      <p class="summary-label">${card.label}</p>
      <p class="summary-value">${card.value}</p>
    `;
    summaryGrid.appendChild(element);
  });
}

function renderTable() {
  const filteredRows = getFilteredRows();
  tableBody.innerHTML = "";

  filteredRows
    .sort(compareRows)
    .forEach((row) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${row.launchDate}</td>
        <td>${row.endDate || ""}</td>
        <td>${row.funnel}</td>
        <td>${row.campaign}</td>
        <td>${formatTableAdSet(row.adSet)}</td>
        <td>${renderAdName(row)}</td>
        <td>${renderPreviewLink(row)}</td>
        <td>${row.versionCampaign || ""}</td>
        <td>${formatCurrency(row.spend)}</td>
        <td>${formatNumber(row.purchases)}</td>
        <td>${formatRoas(row.roas ?? calculateRoas(row.revenue, row.spend))}</td>
        <td><span class="${statusTagClass(row.status)}">${row.status}</span></td>
        <td>${row.notes}</td>
      `;
      tableBody.appendChild(tr);
    });

  renderSummary(filteredRows);
  updateSortHeaders();
}

function renderBoard() {
  const filteredRows = getFilteredRows();
  const funnelOrder = ["UF", "MF", "LF", "Retention", "Other"];
  const funnelLabels = {
    UF: "Upper Funnel",
    MF: "Mid Funnel",
    LF: "Lower Funnel",
    Retention: "Retention",
    Other: "Other"
  };

  boardViewContent.innerHTML = "";

  funnelOrder
    .filter((funnel) => filteredRows.some((row) => row.funnel === funnel))
    .forEach((funnel) => {
      const section = document.createElement("section");
      section.className = "board-section";

      const header = document.createElement("div");
      header.className = "board-section-header";
      header.innerHTML = `<h3>${funnelLabels[funnel] || funnel}</h3>`;
      section.appendChild(header);

      const columnsWrap = document.createElement("div");
      columnsWrap.className = "board-columns";

      const adSets = [
        ...new Set(
          filteredRows
            .filter((row) => row.funnel === funnel)
            .map((row) => row.adSet)
        )
      ].sort();

      adSets.forEach((adSet) => {
        const column = document.createElement("article");
        column.className = "board-column";

        const columnHeader = document.createElement("div");
        columnHeader.className = "board-column-header";
        columnHeader.innerHTML = `<h4>${adSet}</h4>`;
        column.appendChild(columnHeader);

        const list = document.createElement("div");
        list.className = "board-list";

        filteredRows
          .filter((row) => row.funnel === funnel && row.adSet === adSet)
          .sort(compareRows)
          .forEach((row) => {
            const card = document.createElement("div");
            card.className = "board-card";
            card.innerHTML = `
              <div>
                <p class="board-card-name">${row.adName}</p>
                <p class="board-card-meta">${row.versionCampaign || ""}${row.endDate ? ` • ${row.endDate}` : ""}</p>
              </div>
              <div class="board-card-date">${row.launchDate}</div>
            `;
            list.appendChild(card);
          });

        column.appendChild(list);
        columnsWrap.appendChild(column);
      });

      section.appendChild(columnsWrap);
      boardViewContent.appendChild(section);
    });
}

function getAudienceLabel(campaign) {
  if (campaign.includes("Prospecting")) {
    return "Prospecting";
  }
  if (campaign.includes("Inactive")) {
    return "Inactive";
  }
  if (campaign.includes("Active")) {
    return "Active";
  }
  return "Other";
}

function renderPreviewView() {
  const filteredRows = getFilteredRows().filter((row) => row.previewUrl);
  const funnelOrder = ["UF", "MF", "LF", "Retention", "Other"];
  const funnelLabels = {
    UF: "Upper Funnel",
    MF: "Mid Funnel",
    LF: "Lower Funnel",
    Retention: "Retention",
    Other: "Other"
  };
  const audienceOrder = ["Inactive", "Prospecting", "Active", "Other"];

  previewViewContent.innerHTML = "";

  funnelOrder
    .filter((funnel) => filteredRows.some((row) => row.funnel === funnel))
    .forEach((funnel) => {
      const section = document.createElement("section");
      section.className = "preview-section";

      const header = document.createElement("div");
      header.className = "preview-section-header";
      header.innerHTML = `<h3>${funnelLabels[funnel] || funnel}</h3>`;
      section.appendChild(header);

      const audienceGrid = document.createElement("div");
      audienceGrid.className = "preview-audience-grid";

      audienceOrder
        .filter((audience) =>
          filteredRows.some((row) => row.funnel === funnel && getAudienceLabel(row.campaign) === audience)
        )
        .forEach((audience) => {
          const column = document.createElement("article");
          column.className = "preview-audience-column";

          const columnHeader = document.createElement("div");
          columnHeader.className = "preview-audience-header";
          columnHeader.innerHTML = `<h4>${audience}</h4>`;
          column.appendChild(columnHeader);

          const list = document.createElement("div");
          list.className = "preview-list";

          filteredRows
            .filter((row) => row.funnel === funnel && getAudienceLabel(row.campaign) === audience)
            .sort((a, b) => a.launchDate.localeCompare(b.launchDate) || a.adName.localeCompare(b.adName))
            .forEach((row) => {
              const card = document.createElement("div");
              card.className = "preview-card";
              card.innerHTML = `
                <p class="preview-card-title">${escapeHtml(row.adName)}</p>
                <div class="preview-card-meta">
                  <span class="preview-card-date">${row.launchDate}</span>
                  <a class="preview-card-link" href="${escapeHtml(row.previewUrl)}" target="_blank" rel="noreferrer">Open Preview</a>
                </div>
              `;
              list.appendChild(card);
            });

          column.appendChild(list);
          audienceGrid.appendChild(column);
        });

      section.appendChild(audienceGrid);
      previewViewContent.appendChild(section);
    });
}

function getRowKey(row) {
  return [row.funnel, row.campaign, row.adSet, row.adName, row.versionCampaign].join("||");
}

function getCalendarEndDate(row) {
  const planned = plannedEndDates[getRowKey(row)];
  if (planned) {
    return planned;
  }
  if (row.endDate && row.endDate !== "Live") {
    return row.endDate;
  }
  return importedMeta.lastDataDate || row.launchDate;
}

function formatShortDate(value) {
  if (!value) {
    return "";
  }
  const [year, month, day] = value.split("-");
  return `${Number(month)}/${Number(day)}/${year.slice(2)}`;
}

function getVersionCampaignBarClass(versionCampaign) {
  const value = String(versionCampaign || "").toLowerCase();

  if (value.includes("hg spring")) {
    return "calendar-bar-vc-hg-spring";
  }
  if (value.includes("chewpanions")) {
    return "calendar-bar-vc-chewpanions";
  }
  if (value.includes("influencer") || value === "inf") {
    return "calendar-bar-vc-influencer";
  }
  if (value.includes("education")) {
    return "calendar-bar-vc-education";
  }
  if (value.includes("food")) {
    return "calendar-bar-vc-food";
  }
  if (value.includes("rx")) {
    return "calendar-bar-vc-rx";
  }
  if (value.includes("imc")) {
    return "calendar-bar-vc-imc";
  }
  if (value.includes("vf")) {
    return "calendar-bar-vc-vf";
  }
  if (value.includes("crm")) {
    return "calendar-bar-vc-crm";
  }
  if (value.includes("brand")) {
    return "calendar-bar-vc-brand";
  }
  return "calendar-bar-vc-default";
}

function getBarClass(row) {
  return ["calendar-bar", getVersionCampaignBarClass(row.versionCampaign)].join(" ");
}

function getBarPosition(startDate, endDate) {
  const totalStart = visibleFiscalPeriods[0].start;
  const totalEnd = visibleFiscalPeriods[visibleFiscalPeriods.length - 1].end;
  const start = Math.max(new Date(startDate).getTime(), new Date(totalStart).getTime());
  const end = Math.min(new Date(endDate).getTime(), new Date(totalEnd).getTime());
  const totalSpan = new Date(totalEnd).getTime() - new Date(totalStart).getTime();
  const left = ((start - new Date(totalStart).getTime()) / totalSpan) * 100;
  const width = Math.max(((end - start) / totalSpan) * 100, 2.5);
  return { left, width };
}

function getFunnelSortIndex(funnel) {
  return {
    UF: 0,
    MF: 1,
    LF: 2,
    Retention: 3,
    Other: 4
  }[funnel] ?? 5;
}

function getCalendarPlannerMatches(filteredRows) {
  return filteredRows.filter((row) => {
    const plannerVersion = row.plannerVersionCampaign || row.versionCampaign || "";
    const adFreeField = row.adFreeField || row.adName || "";
    const versionMatch = !selectedCalendarPlannerVersion || plannerVersion === selectedCalendarPlannerVersion;
    const freeFieldMatch =
      !selectedCalendarFreeField ||
      selectedCalendarFreeField === "__ALL__" ||
      adFreeField === selectedCalendarFreeField;
    return versionMatch && freeFieldMatch;
  });
}

function syncCalendarPlanner(filteredRows) {
  const plannerVersions = [
    ...new Set(
      filteredRows
        .map((row) => row.plannerVersionCampaign || row.versionCampaign || "")
        .filter(Boolean)
    )
  ].sort();

  if (!plannerVersions.includes(selectedCalendarPlannerVersion)) {
    selectedCalendarPlannerVersion = "";
  }

  calendarPlanSelect.innerHTML = '<option value="">Select version/campaign</option>';
  plannerVersions.forEach((value) => {
    const element = document.createElement("option");
    element.value = value;
    element.textContent = value;
    if (value === selectedCalendarPlannerVersion) {
      element.selected = true;
    }
    calendarPlanSelect.appendChild(element);
  });

  const freeFieldOptions = [
    ...new Set(
      filteredRows
        .filter((row) => {
          const plannerVersion = row.plannerVersionCampaign || row.versionCampaign || "";
          return !selectedCalendarPlannerVersion || plannerVersion === selectedCalendarPlannerVersion;
        })
        .map((row) => row.adFreeField || row.adName || "")
        .filter(Boolean)
    )
  ].sort();

  if (selectedCalendarFreeField !== "__ALL__" && !freeFieldOptions.includes(selectedCalendarFreeField)) {
    selectedCalendarFreeField = "";
  }

  calendarFreeFieldSelect.innerHTML = [
    '<option value="">Select ad free field</option>',
    '<option value="__ALL__">All ad free fields</option>'
  ].join("");
  freeFieldOptions.forEach((value) => {
    const element = document.createElement("option");
    element.value = value;
    element.textContent = value;
    if (value === selectedCalendarFreeField) {
      element.selected = true;
    }
    calendarFreeFieldSelect.appendChild(element);
  });

  const matches = getCalendarPlannerMatches(filteredRows);
  const matchedKeys = matches.map((row) => getRowKey(row));
  const matchedValues = matchedKeys
    .map((key) => plannedEndDates[key])
    .filter(Boolean);
  const uniqueValues = [...new Set(matchedValues)];
  const resolvedValue = uniqueValues.length === 1 ? uniqueValues[0] : "";
  const hasSelection = Boolean(selectedCalendarPlannerVersion && selectedCalendarFreeField && matches.length);

  calendarPlanDate.disabled = !hasSelection;
  calendarPlanSave.disabled = !hasSelection;
  calendarPlanClear.disabled = !hasSelection;
  calendarPlanDate.value = resolvedValue;
}

function renderCalendarView() {
  const filteredRows = getFilteredRows();
  calendarViewContent.innerHTML = "";
  syncCalendarPlanner(filteredRows);

  const grid = document.createElement("div");
  grid.className = "calendar-grid";
  grid.style.setProperty("--calendar-period-count", String(visibleFiscalPeriods.length));

  const headerRow = document.createElement("div");
  headerRow.className = "calendar-header-row";
  headerRow.innerHTML = `
    <div class="calendar-header-label">Campaign / Ad Set / Ad</div>
    ${visibleFiscalPeriods
      .map(
        (period) => `
          <div class="calendar-header-cell">
            <span class="calendar-period-name">${period.name}</span>
            <span class="calendar-period-dates">${formatShortDate(period.start)} - ${formatShortDate(period.end)}</span>
          </div>
        `
      )
      .join("")}
  `;
  grid.appendChild(headerRow);

  const grouped = new Map();
  [...filteredRows]
    .sort(
      (a, b) =>
        getFunnelSortIndex(a.funnel) - getFunnelSortIndex(b.funnel) ||
        a.campaign.localeCompare(b.campaign) ||
        a.adSet.localeCompare(b.adSet) ||
        a.adName.localeCompare(b.adName)
    )
    .forEach((row) => {
    if (!grouped.has(row.campaign)) {
      grouped.set(row.campaign, new Map());
    }
    const adSets = grouped.get(row.campaign);
    if (!adSets.has(row.adSet)) {
      adSets.set(row.adSet, []);
    }
    adSets.get(row.adSet).push(row);
    });

  grouped.forEach((adSets, campaign) => {
    const campaignRow = document.createElement("div");
    campaignRow.className = "calendar-period-row";
    campaignRow.innerHTML = `
      <div class="calendar-row-label calendar-campaign-label">${escapeHtml(campaign)}</div>
      ${visibleFiscalPeriods.map(() => '<div class="calendar-period-cell"></div>').join("")}
    `;
    grid.appendChild(campaignRow);

    adSets.forEach((rowsForAdSet, adSet) => {
      const adSetRow = document.createElement("div");
      adSetRow.className = "calendar-period-row";
      adSetRow.innerHTML = `
        <div class="calendar-row-label calendar-adset-label">${escapeHtml(adSet)}</div>
        ${visibleFiscalPeriods.map(() => '<div class="calendar-period-cell"></div>').join("")}
      `;
      grid.appendChild(adSetRow);

      rowsForAdSet
        .sort((a, b) => a.launchDate.localeCompare(b.launchDate) || a.adName.localeCompare(b.adName))
        .forEach((row) => {
          const dataRow = document.createElement("div");
          dataRow.className = "calendar-data-row";

          const rowKey = getRowKey(row);
          const effectiveEnd = getCalendarEndDate(row);
          const position = getBarPosition(row.launchDate, effectiveEnd);

          const labelCell = document.createElement("div");
          labelCell.className = "calendar-row-label calendar-ad-label";
          labelCell.innerHTML = `<div class="calendar-ad-spacer"></div>`;

          const track = document.createElement("div");
          track.className = "calendar-track calendar-track-empty";
          track.innerHTML = `
            <div class="${getBarClass(row)}" style="left:${position.left}%; width:${position.width}%;">
              <span class="calendar-bar-text">${escapeHtml(row.adName)}</span>
            </div>
          `;

          dataRow.appendChild(labelCell);
          dataRow.appendChild(track);
          grid.appendChild(dataRow);
        });
    });
  });

  calendarViewContent.appendChild(grid);
}

function refreshFilters() {
  populateFilter(funnelFilter, uniqueValues("funnel"), "funnels");
  populateFilter(campaignFilter, getVisibleCampaigns(), "campaigns");
  populateFilter(adSetFilter, getVisibleAdSets(), "ad sets");
}

function renderAll() {
  refreshFilters();
  renderFreshness();
  renderTable();
  renderBoard();
  renderPreviewView();
  renderCalendarView();
  renderView();
}

function renderFreshness() {
  const label = importedMeta.lastDataDate
    ? `Data updated from Meta through ${formatDisplayDate(importedMeta.lastDataDate)}`
    : "Data updated from Meta";
  dataFreshness.textContent = label;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 2
  }).format(Number(value || 0));
}

function renderAdName(row) {
  if (!row.previewUrl) {
    return row.adName;
  }

  return `<span class="ad-name-link" data-preview-url="${escapeHtml(row.previewUrl)}" data-preview-name="${escapeHtml(row.adName)}">${escapeHtml(row.adName)}</span>`;
}

function renderPreviewLink(row) {
  if (!row.previewUrl) {
    return "";
  }

  return `<a href="${escapeHtml(row.previewUrl)}" target="_blank" rel="noreferrer">Open</a>`;
}

function formatTableAdSet(adSet) {
  if (!adSet) {
    return "";
  }

  if (adSet.includes("_DABA ") || adSet.includes("DABA ")) {
    return "DABA";
  }

  return adSet;
}

function formatNumber(value) {
  return new Intl.NumberFormat("en-US").format(Number(value || 0));
}

function formatDisplayDate(value) {
  if (!value) {
    return "";
  }
  const [year, month, day] = value.split("-");
  return `${Number(month)}/${Number(day)}/${year}`;
}

function formatRoas(value) {
  return `${Number(value || 0).toFixed(2)}x`;
}

function calculateRoas(revenue, spend) {
  if (!Number(spend)) {
    return 0;
  }
  return Number(revenue || 0) / Number(spend);
}

function compareRows(a, b) {
  const direction = sortState.direction === "asc" ? 1 : -1;
  const key = sortState.key;
  const numericKeys = new Set(["spend", "purchases", "revenue", "roas"]);
  const aValue =
    key === "roas" ? a.roas ?? calculateRoas(a.revenue, a.spend) : (a[key] ?? "");
  const bValue =
    key === "roas" ? b.roas ?? calculateRoas(b.revenue, b.spend) : (b[key] ?? "");

  if (numericKeys.has(key)) {
    return (Number(aValue) - Number(bValue)) * direction;
  }

  return String(aValue).localeCompare(String(bValue)) * direction;
}

function updateSortHeaders() {
  sortableHeaders.forEach((header) => {
    header.classList.remove("sorted-asc", "sorted-desc");
    if (header.dataset.sortKey === sortState.key) {
      header.classList.add(sortState.direction === "asc" ? "sorted-asc" : "sorted-desc");
    }
  });
}

function renderView() {
  tableViewPanel.classList.toggle("hidden", currentView !== "table");
  boardViewPanel.classList.toggle("hidden", currentView !== "board");
  previewViewPanel.classList.toggle("hidden", currentView !== "preview");
  calendarViewPanel.classList.toggle("hidden", currentView !== "calendar");
  tableViewTab.classList.toggle("is-active", currentView === "table");
  boardViewTab.classList.toggle("is-active", currentView === "board");
  previewViewTab.classList.toggle("is-active", currentView === "preview");
  calendarViewTab.classList.toggle("is-active", currentView === "calendar");
}

function escapeCsv(value) {
  return `"${String(value).replaceAll('"', '""')}"`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function exportCsv() {
  const filteredRows = getFilteredRows();
  const header = [
    "Launch Date",
    "End Date",
    "Funnel",
    "Campaign",
    "Ad Set",
    "Ad Name",
    "Preview Link",
    "Version/ Campaign",
    "Spend",
    "Purchases",
    "ROAS",
    "Status",
    "Notes"
  ];

  const lines = [
    header.join(","),
    ...filteredRows.map((row) =>
      [
        row.launchDate,
        row.endDate,
        row.funnel,
        row.campaign,
        row.adSet,
        row.adName,
        row.previewUrl,
        row.versionCampaign,
        row.spend,
        row.purchases,
        row.roas ?? calculateRoas(row.revenue, row.spend),
        row.status,
        row.notes
      ]
        .map(escapeCsv)
        .join(",")
    )
  ];

  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "meta-creative-calendar.csv";
  link.click();
  URL.revokeObjectURL(url);
}

creativeForm.addEventListener("submit", (event) => {
  event.preventDefault();

  const formData = new FormData(creativeForm);
  rows.push({
    launchDate: formData.get("launchDate"),
    endDate: formData.get("endDate").trim(),
    funnel: formData.get("funnel"),
    versionCampaign: formData.get("versionCampaign").trim(),
    campaign: formData.get("campaign").trim(),
    adSet: formData.get("adSet").trim(),
    adName: formData.get("adName").trim(),
    spend: Number(formData.get("spend") || 0),
    purchases: Number(formData.get("purchases") || 0),
    revenue: 0,
    roas: Number(formData.get("roas") || 0),
    status: formData.get("status"),
    notes: formData.get("notes").trim()
  });

  saveRows();
  creativeForm.reset();
  renderAll();
});

funnelFilter.addEventListener("change", () => {
  refreshFilters();
  renderTable();
  renderBoard();
  renderPreviewView();
  renderCalendarView();
});

campaignFilter.addEventListener("change", () => {
  refreshFilters();
  renderTable();
  renderBoard();
  renderPreviewView();
  renderCalendarView();
});

adSetFilter.addEventListener("change", renderTable);
adSetFilter.addEventListener("change", renderBoard);
adSetFilter.addEventListener("change", renderPreviewView);
adSetFilter.addEventListener("change", renderCalendarView);
statusFilter.addEventListener("change", renderTable);
statusFilter.addEventListener("change", renderBoard);
statusFilter.addEventListener("change", renderPreviewView);
statusFilter.addEventListener("change", renderCalendarView);

exportCsvButton.addEventListener("click", exportCsv);

tableViewTab.addEventListener("click", () => {
  currentView = "table";
  renderView();
});

boardViewTab.addEventListener("click", () => {
  currentView = "board";
  renderView();
});

previewViewTab.addEventListener("click", () => {
  currentView = "preview";
  renderView();
});

calendarViewTab.addEventListener("click", () => {
  currentView = "calendar";
  renderView();
});

calendarPlanSelect.addEventListener("change", () => {
  selectedCalendarPlannerVersion = calendarPlanSelect.value;
  renderCalendarView();
});

calendarFreeFieldSelect.addEventListener("change", () => {
  selectedCalendarFreeField = calendarFreeFieldSelect.value;
  renderCalendarView();
});

calendarPlanSave.addEventListener("click", () => {
  const matches = getCalendarPlannerMatches(getFilteredRows());
  if (!selectedCalendarPlannerVersion || !selectedCalendarFreeField || !calendarPlanDate.value || !matches.length) {
    return;
  }
  matches.forEach((row) => {
    plannedEndDates[getRowKey(row)] = calendarPlanDate.value;
  });
  savePlannedEndDates();
  renderCalendarView();
});

calendarPlanClear.addEventListener("click", () => {
  const matches = getCalendarPlannerMatches(getFilteredRows());
  if (!selectedCalendarPlannerVersion || !selectedCalendarFreeField || !matches.length) {
    return;
  }
  matches.forEach((row) => {
    delete plannedEndDates[getRowKey(row)];
  });
  calendarPlanDate.value = "";
  savePlannedEndDates();
  renderCalendarView();
});

sortableHeaders.forEach((header) => {
  header.addEventListener("click", () => {
    const key = header.dataset.sortKey;
    if (sortState.key === key) {
      sortState.direction = sortState.direction === "asc" ? "desc" : "asc";
    } else {
      sortState = { key, direction: key === "launchDate" ? "asc" : "desc" };
    }
    renderTable();
    renderBoard();
    renderPreviewView();
    renderCalendarView();
  });
});

function positionPreviewTooltip(x, y) {
  previewTooltip.style.left = `${Math.min(x + 16, window.innerWidth - 300)}px`;
  previewTooltip.style.top = `${Math.min(y + 16, window.innerHeight - 180)}px`;
}

function positionPreviewTooltipForTarget(target) {
  const rect = target.getBoundingClientRect();
  const x = rect.right + 12;
  const y = rect.top;
  previewTooltip.style.left = `${Math.min(x, window.innerWidth - 300)}px`;
  previewTooltip.style.top = `${Math.min(y, window.innerHeight - 180)}px`;
}

function showPreviewTooltip(target) {
  activePreviewTarget = target;
  previewTooltip.innerHTML = `
    <p class="preview-tooltip-title">${target.dataset.previewName}</p>
    <p class="preview-tooltip-copy">Meta preview links usually need to open in Facebook, so this hover shows the preview link instead of a direct image embed.</p>
    <a class="preview-tooltip-link" href="${target.dataset.previewUrl}" target="_blank" rel="noreferrer">Open Preview</a>
  `;
  previewTooltip.classList.remove("hidden");
  positionPreviewTooltipForTarget(target);
}

function hidePreviewTooltip() {
  activePreviewTarget = null;
  previewTooltip.classList.add("hidden");
}

document.addEventListener("mouseover", (event) => {
  const target = event.target.closest(".ad-name-link");
  if (target) {
    showPreviewTooltip(target);
  }
});

document.addEventListener("mouseout", (event) => {
  const fromLink = event.target.closest(".ad-name-link");
  if (!fromLink) {
    return;
  }

  const nextTarget = event.relatedTarget;
  if (nextTarget && (nextTarget.closest(".ad-name-link") || nextTarget.closest("#previewTooltip"))) {
    return;
  }

  hidePreviewTooltip();
});

previewTooltip.addEventListener("mouseleave", (event) => {
  const nextTarget = event.relatedTarget;
  if (nextTarget && activePreviewTarget && nextTarget.closest(".ad-name-link")) {
    return;
  }
  hidePreviewTooltip();
});

window.addEventListener("scroll", () => {
  if (!previewTooltip.classList.contains("hidden") && activePreviewTarget) {
    positionPreviewTooltipForTarget(activePreviewTarget);
  }
});

renderAll();
