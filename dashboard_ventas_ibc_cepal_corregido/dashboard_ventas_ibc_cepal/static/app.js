const $ = (id) => document.getElementById(id);
const CLASSIFICATIONS = ["Venta Salvada", "Venta Caída", "Otra Gestión", "Sin clasificación"];
const COLORS = ["#18a979", "#e26068", "#e8a43c", "#8d97a6"];
const FIELDS = [
  "FECHA GESTIÓN", "EMPRESA", "ÁREA", "VENDEDOR", "CLIENTE", "DNI", "TELÉFONO",
  "REGISTRO 1", "REGISTRO 2", "REGISTRO 3", "ESTADO",
  "CANTIDAD DE COMUNICACIONES DEL MISMO CLIENTE", "CONTRATOS",
  "CLIENTE ACEPTA PRODUCTO", "FECHA DE ACEPTACIÓN DE PRODUCTO",
  "COMUNICACIÓN DE VENTA", "CLASIFICACIÓN", "MOMENTO DÍAS",
  "FECHA PRIMER PAGO", "MOMENTO PRIMER PAGO",
];
let allRows = [];
let currentRows = [];
let numberFormat = new Intl.NumberFormat("es-PE");

const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({
  "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
}[char]));
const count = (value) => numberFormat.format(value || 0);
const isSaved = (row) => row["CLASIFICACIÓN"] === "Venta Salvada";
const statusOf = (row) => CLASSIFICATIONS.includes(row["CLASIFICACIÓN"])
  ? row["CLASIFICACIÓN"] : "Sin clasificación";
const percent = (value, total) => `${total ? (value / total * 100).toFixed(1) : "0.0"}%`;
const monthOf = (row) => (row["FECHA GESTIÓN"] || "").slice(0, 7);
const monthLabel = (month) => {
  if (!/^\d{4}-\d{2}$/.test(month)) return month || "Sin fecha";
  const date = new Date(`${month}-01T12:00:00`);
  return new Intl.DateTimeFormat("es-PE", { month: "short", year: "numeric" }).format(date);
};
const dateLabel = (value) => {
  const match = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? `${match[3]}/${match[2]}/${match[1]}` : (value || "No registrado en el Excel");
};
const sortedUnique = (rows, field) => [...new Set(rows.map((row) => row[field]).filter(Boolean))]
  .sort((a, b) => a.localeCompare(b, "es"));

async function request(url, options = {}) {
  const response = await fetch(url, options);
  if (!response.ok) {
    let message = `Error HTTP ${response.status}`;
    try {
      const body = await response.json();
      if (typeof body.detail === "string") message = body.detail;
    } catch { /* The response may not contain JSON. */ }
    throw new Error(message);
  }
  return response;
}

function showNotice(message, error = false) {
  $("notice").textContent = message;
  $("notice").classList.toggle("notice-error", error);
}

function updateOptions(id, rows, field, firstLabel) {
  const select = $(id);
  const selected = select.value;
  const values = sortedUnique(rows, field);
  select.innerHTML = `<option value="">${esc(firstLabel)}</option>${values.map((value) =>
    `<option value="${esc(value)}">${esc(value)}</option>`).join("")}`;
  if (values.includes(selected)) select.value = selected;
}

function setRows(payload) {
  for (const id of ["company", "area", "advisor", "classification", "month", "dateFrom", "dateTo", "search"]) {
    $(id).value = "";
  }
  $("momentMode").value = "all";
  allRows = payload.rows || [];
  $("sourceName").textContent = payload.filename || "Archivo cargado";
  $("periodLabel").textContent = payload.period?.from
    ? `Periodo disponible: ${dateLabel(payload.period.from)} – ${dateLabel(payload.period.to)}`
    : "Sin fechas de gestión";
  updateOptions("company", allRows, "EMPRESA", "Todas las empresas");
  updateOptions("area", allRows, "ÁREA", "Todas las áreas");
  updateOptions("advisor", allRows, "VENDEDOR", "Todos los asesores");
  const classifications = [...new Set(allRows.map(statusOf))].sort((a, b) => a.localeCompare(b, "es"));
  $("classification").innerHTML = `<option value="">Todas las clasificaciones</option>${classifications.map((value) =>
    `<option value="${esc(value)}">${esc(value || "Sin clasificación")}</option>`).join("")}`;
  const months = [...new Set(allRows.map(monthOf).filter(Boolean))].sort();
  $("month").innerHTML = `<option value="">Todos los meses</option>${months.map((month) =>
    `<option value="${esc(month)}">${esc(monthLabel(month))}</option>`).join("")}`;
  const warnings = payload.warnings || [];
  $("warning").hidden = warnings.length === 0;
  $("warning").textContent = warnings.length ? `Revisión de calidad: ${warnings.join(" · ")}.` : "";
  showNotice(`${count(allRows.length)} registros cargados desde ${payload.filename || "el Excel"}.`);
  render();
}

function filters() {
  return {
    company: $("company").value,
    area: $("area").value,
    advisor: $("advisor").value,
    classification: $("classification").value,
    month: $("month").value,
    dateFrom: $("dateFrom").value,
    dateTo: $("dateTo").value,
    search: $("search").value.trim().toLocaleLowerCase("es"),
  };
}

function selectRows({ ignoreArea = false, ignoreClassification = false } = {}) {
  const filter = filters();
  if (filter.dateFrom && filter.dateTo && filter.dateFrom > filter.dateTo) return [];
  return allRows.filter((row) => {
    const date = row["FECHA GESTIÓN"] || "";
    if (filter.company && row.EMPRESA !== filter.company) return false;
    if (!ignoreArea && filter.area && row["ÁREA"] !== filter.area) return false;
    if (filter.advisor && row.VENDEDOR !== filter.advisor) return false;
    if (!ignoreClassification && filter.classification && statusOf(row) !== filter.classification) return false;
    if (filter.month && !date.startsWith(filter.month)) return false;
    if (filter.dateFrom && (!date || date < filter.dateFrom)) return false;
    if (filter.dateTo && (!date || date > filter.dateTo)) return false;
    if (filter.search && !FIELDS.some((field) =>
      String(row[field] || "").toLocaleLowerCase("es").includes(filter.search))) return false;
    return true;
  });
}

function countCrossCases() {
  const rows = selectRows();
  const clients = new Map();
  for (const row of rows) {
    const dni = String(row.DNI || "").trim().replace(/\.0$/, "").replace(/^0+/, "");
    if (!/^\d+$/.test(dni)) continue;
    const company = String(row.EMPRESA || "").trim().toLocaleUpperCase("es");
    const key = `${company}\u0000${dni}`;
    if (!clients.has(key)) clients.set(key, new Set());
    if (row["ÁREA"] === "VENTAS") clients.get(key).add("VENTAS");
    if (row["ÁREA"] === "ATENCION AL CLIENTE") clients.get(key).add("ATC");
  }
  return [...clients.values()].filter((areas) => areas.has("VENTAS") && areas.has("ATC")).length;
}

function summarize(rows) {
  const counts = Object.fromEntries(CLASSIFICATIONS.map((label) => [label, 0]));
  rows.forEach((row) => { counts[statusOf(row)] += 1; });
  const saved = rows.filter(isSaved);
  const days = new Map();
  const payments = new Map();
  const dayCounts = {
    savedWithin: 0,
    savedOutside: 0,
    fallenWithin: 0,
    fallenOutside: 0,
    otherWithin: 0,
    otherOutside: 0,
  };
  const paymentCounts = {
    savedBefore: 0,
    savedAfter: 0,
    fallenBefore: 0,
    fallenAfter: 0,
    otherBefore: 0,
    otherAfter: 0,
  };
  saved.forEach((row) => {
    const day = row["MOMENTO DÍAS"] || "Sin dato";
    const payment = row["MOMENTO PRIMER PAGO"] || "Sin dato";
    days.set(day, (days.get(day) || 0) + 1);
    payments.set(payment, (payments.get(payment) || 0) + 1);
  });
  rows.forEach((row) => {
    const day = row["MOMENTO DÍAS"];
    if (day === "Dentro de 3 días" || day === "Fuera de 3 días") {
      if (row["CLASIFICACIÓN"] === "Venta Salvada") {
        if (day === "Dentro de 3 días") dayCounts.savedWithin += 1;
        else dayCounts.savedOutside += 1;
      } else if (row["CLASIFICACIÓN"] === "Venta Caída") {
        if (day === "Dentro de 3 días") dayCounts.fallenWithin += 1;
        else dayCounts.fallenOutside += 1;
      } else if (row["CLASIFICACIÓN"] === "Otra Gestión") {
        if (day === "Dentro de 3 días") dayCounts.otherWithin += 1;
        else dayCounts.otherOutside += 1;
      }
    }
    const payment = row["MOMENTO PRIMER PAGO"];
    if (payment !== "Antes del primer pago" && payment !== "Después del primer pago") return;
    if (row["CLASIFICACIÓN"] === "Venta Salvada") {
      if (payment === "Antes del primer pago") paymentCounts.savedBefore += 1;
      else paymentCounts.savedAfter += 1;
    } else if (row["CLASIFICACIÓN"] === "Venta Caída") {
      if (payment === "Antes del primer pago") paymentCounts.fallenBefore += 1;
      else paymentCounts.fallenAfter += 1;
    } else if (row["CLASIFICACIÓN"] === "Otra Gestión") {
      if (payment === "Antes del primer pago") paymentCounts.otherBefore += 1;
      else paymentCounts.otherAfter += 1;
    }
  });
  const recognized = counts["Venta Salvada"] + counts["Venta Caída"];
  return {
    total: rows.length,
    salesArea: rows.filter((row) => row["ÁREA"] === "VENTAS").length,
    atcArea: rows.filter((row) => row["ÁREA"] === "ATENCION AL CLIENTE").length,
    saved: counts["Venta Salvada"],
    fallen: counts["Venta Caída"],
    annulmentAttempts: recognized,
    other: counts["Otra Gestión"],
    unclassified: counts["Sin clasificación"],
    savedRate: percent(counts["Venta Salvada"], rows.length),
    fallenRate: percent(counts["Venta Caída"], rows.length),
    days,
    dayCounts,
    payments,
    paymentCounts,
    paymentDates: rows.filter((row) => row["FECHA PRIMER PAGO"]).length,
  };
}

function renderMetrics(rows) {
  const summary = summarize(rows);
  const crossCases = countCrossCases();
  $("totalMetric").textContent = count(summary.total);
  $("annulmentMetric").textContent = count(summary.annulmentAttempts);
  $("annulmentDescription").textContent = `${percent(summary.annulmentAttempts, summary.total)} del total · salvadas + caídas`;
  $("salesAreaMetric").textContent = count(summary.salesArea);
  $("salesAreaShare").textContent = `${percent(summary.salesArea, summary.total)} del total de casos`;
  $("atcAreaMetric").textContent = count(summary.atcArea);
  $("atcAreaShare").textContent = `${percent(summary.atcArea, summary.total)} del total de casos`;
  $("crossCaseMetric").textContent = count(crossCases);
  $("crossCaseDescription").textContent = `${percent(crossCases, summary.total)} del total · DNI en Ventas y ATC`;
  $("savedMetric").textContent = count(summary.saved);
  $("savedRate").textContent = `${summary.savedRate} del total de casos`;
  $("fallenMetric").textContent = count(summary.fallen);
  $("fallenShare").textContent = `Tasa de caídas: ${summary.fallenRate} del total de casos`;
  $("otherMetric").textContent = count(summary.other);
  $("otherShare").textContent = `${percent(summary.other, summary.total)} del total de casos`;
  const within = summary.dayCounts.savedWithin;
  const outside = summary.dayCounts.savedOutside;
  $("withinMetric").textContent = count(within);
  $("withinShare").textContent = `${percent(within, summary.total)} del total de casos`;
  $("outsideMetric").textContent = count(outside);
  $("outsideShare").textContent = `${percent(outside, summary.total)} del total de casos`;
  $("fallenWithinMetric").textContent = count(summary.dayCounts.fallenWithin);
  $("fallenWithinShare").textContent = `${percent(summary.dayCounts.fallenWithin, summary.total)} del total de casos`;
  $("fallenOutsideMetric").textContent = count(summary.dayCounts.fallenOutside);
  $("fallenOutsideShare").textContent = `${percent(summary.dayCounts.fallenOutside, summary.total)} del total de casos`;
  $("otherWithinMetric").textContent = count(summary.dayCounts.otherWithin);
  $("otherWithinShare").textContent = `${percent(summary.dayCounts.otherWithin, summary.total)} del total de casos`;
  $("otherOutsideMetric").textContent = count(summary.dayCounts.otherOutside);
  $("otherOutsideShare").textContent = `${percent(summary.dayCounts.otherOutside, summary.total)} del total de casos`;
  const before = summary.paymentCounts.savedBefore;
  const after = summary.paymentCounts.savedAfter;
  $("beforePaymentMetric").textContent = count(before);
  $("paymentCoverage").textContent = `${percent(before, summary.total)} del total de casos`;
  $("afterPaymentMetric").textContent = count(after);
  $("afterPaymentShare").textContent = `${percent(after, summary.total)} del total de casos`;
  $("fallenBeforePaymentMetric").textContent = count(summary.paymentCounts.fallenBefore);
  $("fallenBeforePaymentShare").textContent = `${percent(summary.paymentCounts.fallenBefore, summary.total)} del total de casos`;
  $("fallenAfterPaymentMetric").textContent = count(summary.paymentCounts.fallenAfter);
  $("fallenAfterPaymentShare").textContent = `${percent(summary.paymentCounts.fallenAfter, summary.total)} del total de casos`;
  $("otherBeforePaymentMetric").textContent = count(summary.paymentCounts.otherBefore);
  $("otherBeforePaymentShare").textContent = `${percent(summary.paymentCounts.otherBefore, summary.total)} del total de casos`;
  $("otherAfterPaymentMetric").textContent = count(summary.paymentCounts.otherAfter);
  $("otherAfterPaymentShare").textContent = `${percent(summary.paymentCounts.otherAfter, summary.total)} del total de casos`;
}

function renderDonut(rows) {
  const summary = summarize(rows);
  const labels = CLASSIFICATIONS.filter((label) => summary[label === "Venta Salvada" ? "saved"
    : label === "Venta Caída" ? "fallen"
      : label === "Otra Gestión" ? "other" : "unclassified"] > 0);
  const values = labels.map((label) => ({
    label,
    value: summary[label === "Venta Salvada" ? "saved" : label === "Venta Caída" ? "fallen"
      : label === "Otra Gestión" ? "other" : "unclassified"],
  }));
  let cursor = 0;
  const stops = values.map(({ label, value }) => {
    const start = cursor;
    cursor += summary.total ? value / summary.total * 100 : 0;
    return `${COLORS[CLASSIFICATIONS.indexOf(label)]} ${start}% ${cursor}%`;
  });
  $("donut").style.background = stops.length ? `conic-gradient(${stops.join(",")})` : "#e8edf2";
  $("donutTotal").textContent = count(summary.total);
  $("legend").innerHTML = values.length ? values.map(({ label, value }) =>
    `<div><span class="legend-dot" style="background:${COLORS[CLASSIFICATIONS.indexOf(label)]}"></span><span>${esc(label)}</span><b>${count(value)}</b><small>${percent(value, summary.total)}</small></div>`).join("")
    : '<p class="empty">No hay registros en el filtro seleccionado.</p>';
}

function areaMomentDefinitions(mode) {
  const days = [
    { group: "MOMENTO DÍAS", field: "MOMENTO DÍAS", label: "Dentro de 3 días" },
    { group: "MOMENTO DÍAS", field: "MOMENTO DÍAS", label: "Fuera de 3 días" },
  ];
  const payment = [
    { group: "MOMENTO PRIMER PAGO", field: "MOMENTO PRIMER PAGO", label: "Antes del primer pago" },
    { group: "MOMENTO PRIMER PAGO", field: "MOMENTO PRIMER PAGO", label: "Después del primer pago" },
  ];
  return mode === "days" ? days : mode === "payment" ? payment : [...days, ...payment];
}

function renderAreaMomentChart(hostId, rows, classification, moments) {
  const host = $(hostId);
  const companies = sortedUnique(rows, "EMPRESA");
  if (!companies.length) {
    host.innerHTML = '<p class="empty">No hay registros para comparar con los filtros actuales.</p>';
    return;
  }
  const areas = [
    { name: "VENTAS", label: "Ventas", className: "area-sales-bar" },
    { name: "ATENCION AL CLIENTE", label: "Atención al Cliente", className: "area-atc-bar" },
  ];
  host.innerHTML = companies.map((company) => {
    const values = moments.map((moment) => ({
      ...moment,
      areas: areas.map((area) => ({
        ...area,
        value: rows.filter((row) => row.EMPRESA === company
          && row["ÁREA"] === area.name
          && row["CLASIFICACIÓN"] === classification
          && row[moment.field] === moment.label).length,
      })),
    }));
    const max = Math.max(1, ...values.flatMap(({ areas: areaValues }) => areaValues.map(({ value }) => value)));
    const step = Math.max(1, Math.ceil(max / 4));
    const maxY = Math.ceil(max / step) * step;
    const width = 560, height = 265, left = 42, right = 12, top = 30, bottom = 68;
    const plotWidth = width - left - right, plotHeight = height - top - bottom;
    const y = (value) => top + plotHeight - value / maxY * plotHeight;
    const slotWidth = plotWidth / values.length;
    const barWidth = Math.min(28, slotWidth * 0.27);
    const barGap = Math.min(8, slotWidth * 0.08);
    const momentLabelLines = (label) => {
      if (label === "Dentro de 3 días") return ["Dentro de 3", "días"];
      if (label === "Fuera de 3 días") return ["Fuera de 3", "días"];
      if (label === "Antes del primer pago") return ["Antes del", "primer pago"];
      return ["Después del", "primer pago"];
    };
    let svg = `<svg class="area-bar-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="${esc(classification)} por momento en ${esc(company)}, comparando Ventas y Atención al Cliente">`;
    for (let value = 0; value <= maxY; value += step) {
      const yy = y(value);
      svg += `<line x1="${left}" y1="${yy}" x2="${width - right}" y2="${yy}" class="area-grid-line"/><text x="${left - 8}" y="${yy + 3}" text-anchor="end" class="area-axis-label">${value}</text>`;
    }
    const groups = [...new Set(values.map(({ group }) => group))];
    groups.forEach((group) => {
      const indices = values.map((moment, index) => moment.group === group ? index : -1).filter((index) => index >= 0);
      const groupCenter = left + slotWidth * (indices[0] + indices[indices.length - 1] + 1) / 2;
      svg += `<text x="${groupCenter}" y="14" text-anchor="middle" class="area-group-label">${esc(group)}</text>`;
    });
    values.forEach((moment, index) => {
      const center = left + slotWidth * (index + 0.5);
      const lines = momentLabelLines(moment.label);
      svg += `<text x="${center}" y="${height - bottom + 17}" text-anchor="middle" class="area-axis-label"><tspan x="${center}">${esc(lines[0])}</tspan><tspan x="${center}" dy="12">${esc(lines[1])}</tspan></text>`;
      moment.areas.forEach((area, areaIndex) => {
        const x = center + (areaIndex === 0 ? -barWidth - barGap / 2 : barGap / 2);
        const barY = y(area.value);
        const barHeight = top + plotHeight - barY;
        svg += `<rect x="${x}" y="${barY}" width="${barWidth}" height="${barHeight}" rx="3" class="${area.className}"><title>${esc(area.label)} · ${esc(moment.label)}: ${area.value}</title></rect>`;
        svg += `<text x="${x + barWidth / 2}" y="${Math.max(top - 3, barY - 5)}" text-anchor="middle" class="area-value-label">${area.value}</text>`;
      });
    });
    svg += `</svg><div class="area-chart-legend"><span><i class="area-sales-bar"></i>Ventas</span><span><i class="area-atc-bar"></i>Atención al Cliente</span></div>`;
    return `<section class="area-company-chart"><h4>${esc(company)}</h4>${svg}</section>`;
  }).join("");
}

function renderAreaMomentCharts(rows) {
  const moments = areaMomentDefinitions($("momentMode").value);
  renderAreaMomentChart("savedAreaMomentChart", rows, "Venta Salvada", moments);
  renderAreaMomentChart("fallenAreaMomentChart", rows, "Venta Caída", moments);
}

function renderCompanyOutcomeAreaChart(rows) {
  const host = $("companyOutcomeAreaChart");
  const areas = new Map([
    ["VENTAS", "Ventas"],
    ["ATENCION AL CLIENTE", "Atención al Cliente"],
  ]);
  const groups = groupedRows(rows.filter((row) => areas.has(row["ÁREA"])), ["EMPRESA", "ÁREA"])
    .map(({ values, rows: areaRows }) => ({
    company: values[0],
    area: values[1],
    label: `${values[0]} · ${areas.get(values[1])}`,
    saved: areaRows.filter(isSaved).length,
    fallen: areaRows.filter((row) => row["CLASIFICACIÓN"] === "Venta Caída").length,
  })).sort((a, b) => a.company.localeCompare(b.company, "es")
    || a.area.localeCompare(b.area, "es"));
  if (!groups.length) {
    host.innerHTML = '<p class="empty">No hay empresas con datos para comparar en los filtros actuales.</p>';
    return;
  }
  const series = [
    { key: "Salvadas", field: "saved", color: "#18a979", className: "company-saved-bar" },
    { key: "Caídas", field: "fallen", color: "#e26068", className: "company-fallen-bar" },
  ];
  const width = 760, rowHeight = 58, height = Math.max(190, groups.length * rowHeight + 48);
  const left = 160, right = 48, top = 24, bottom = 32;
  const plotWidth = width - left - right, plotHeight = height - top - bottom;
  const max = Math.max(1, ...groups.flatMap((group) => series.map(({ field }) => group[field])));
  const step = Math.max(1, Math.ceil(max / 5));
  const maxX = Math.ceil(max / step) * step;
  const x = (value) => left + value / maxX * plotWidth;
  let svg = `<svg class="company-outcome-bars-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Barras horizontales agrupadas con ventas salvadas y caídas por empresa y área">`;
  for (let value = 0; value <= maxX; value += step) {
    const xx = x(value);
    svg += `<line x1="${xx}" y1="${top}" x2="${xx}" y2="${top + plotHeight}" class="company-outcome-grid-line"/><text x="${xx}" y="${height - 8}" text-anchor="middle" class="axis-label">${count(value)}</text>`;
  }
  groups.forEach((group, index) => {
    const rowTop = top + index * rowHeight;
    const centerY = rowTop + rowHeight / 2;
    svg += `<text x="${left - 12}" y="${centerY + 4}" text-anchor="end" class="company-outcome-label">${esc(group.label)}</text>`;
    series.forEach(({ key, field, color, className }, seriesIndex) => {
      const barY = centerY + (seriesIndex === 0 ? -12 : 3);
      const barWidth = x(group[field]) - left;
      svg += `<rect x="${left}" y="${barY}" width="${barWidth}" height="10" rx="4" fill="${color}" class="${className}"><title>${esc(group.label)} · ${key}: ${count(group[field])}</title></rect>`;
      svg += `<text x="${Math.min(width - right + 26, x(group[field]) + 7)}" y="${barY + 9}" class="company-outcome-value">${count(group[field])}</text>`;
    });
  });
  svg += `</svg><div class="chart-legend">${series.map(({ key, color }) =>
    `<span><i style="background:${color}"></i>${key}</span>`).join("")}</div>`;
  host.innerHTML = svg;
}

function renderMonthlyChart(rows) {
  const months = [...new Set(rows.map(monthOf).filter(Boolean))].sort();
  const series = [
    { key: "Venta Salvada", color: COLORS[0] },
    { key: "Venta Caída", color: COLORS[1] },
    { key: "Otra Gestión", color: COLORS[2] },
  ];
  const host = $("monthlyChart");
  if (!months.length) {
    host.innerHTML = '<p class="empty">No hay fechas de gestión para graficar.</p>';
    return;
  }
  const width = 660, height = 265, left = 42, right = 18, top = 22, bottom = 54;
  const plotWidth = width - left - right, plotHeight = height - top - bottom;
  const values = series.map(({ key }) => months.map((month) =>
    rows.filter((row) => monthOf(row) === month && statusOf(row) === key).length));
  const max = Math.max(1, ...values.flat());
  const step = Math.max(1, Math.ceil(max / 4));
  const maxY = Math.ceil(max / step) * step;
  const x = (index) => months.length === 1 ? left + plotWidth / 2 : left + index * plotWidth / (months.length - 1);
  const y = (value) => top + plotHeight - value / maxY * plotHeight;
  let svg = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Clasificación por mes">`;
  for (let value = 0; value <= maxY; value += step) {
    const yy = y(value);
    svg += `<line x1="${left}" y1="${yy}" x2="${width - right}" y2="${yy}" stroke="#e9eef3"/><text x="${left - 9}" y="${yy + 4}" text-anchor="end" class="axis-label">${value}</text>`;
  }
  months.forEach((month, index) => {
    svg += `<text x="${x(index)}" y="${height - 17}" text-anchor="middle" class="axis-label">${esc(monthLabel(month))}</text>`;
  });
  series.forEach(({ key, color }, seriesIndex) => {
    const points = values[seriesIndex].map((value, index) => `${x(index)},${y(value)}`).join(" ");
    svg += `<polyline points="${points}" fill="none" stroke="${color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>`;
    values[seriesIndex].forEach((value, index) => {
      svg += `<circle cx="${x(index)}" cy="${y(value)}" r="4" fill="${color}"><title>${esc(key)} · ${esc(monthLabel(months[index]))}: ${value}</title></circle>`;
      svg += `<text x="${x(index)}" y="${y(value) - 8}" text-anchor="middle" class="point-label">${value}</text>`;
    });
  });
  svg += "</svg>";
  host.innerHTML = `${svg}<div class="chart-legend">${series.map(({ key, color }) =>
    `<span><i style="background:${color}"></i>${esc(key)}</span>`).join("")}</div>`;
}

function renderBreakdown(hostId, map, total, kind) {
  const host = $(hostId);
  const expected = kind === "days"
    ? ["Dentro de 3 días", "Fuera de 3 días"]
    : ["Antes del primer pago", "Después del primer pago"];
  const unknown = [...map.keys()].filter((label) => !expected.includes(label)).sort();
  const entries = [...expected, ...unknown].map((label) => [label, map.get(label) || 0]);
  const matches = [...map.values()].reduce((sum, value) => sum + value, 0);
  if (!matches) {
    host.innerHTML = '<p class="empty">No hay ventas salvadas en el filtro actual.</p>';
    return;
  }
  host.innerHTML = entries.map(([label, value]) => `
    <div class="breakdown-row">
      <div class="breakdown-heading"><span>${esc(label)}</span><b>${count(value)} <small>(${percent(value, total)} del total)</small></b></div>
      <div class="bar-track"><span class="${kind === "days" ? (label === "Dentro de 3 días" ? "bar-green" : "bar-amber") : (label === "Antes del primer pago" ? "bar-blue" : "bar-purple")}" style="width:${total ? value / total * 100 : 0}%"></span></div>
    </div>`).join("");
}

function groupedRows(rows, fields) {
  const groups = new Map();
  rows.forEach((row) => {
    const key = fields.map((field) => row[field] || "Sin dato").join("\u0000");
    if (!groups.has(key)) groups.set(key, { values: fields.map((field) => row[field] || "Sin dato"), rows: [] });
    groups.get(key).rows.push(row);
  });
  return [...groups.values()].map((group) => ({ ...group, summary: summarize(group.rows) }));
}

function statsCells(summary, includeOther = true) {
  const mode = $("momentMode").value;
  const cells = [
    `<td>${count(summary.total)}</td>`,
    `<td class="number-cell">${count(summary.saved)}</td>`,
    `<td class="number-cell">${count(summary.fallen)}</td>`,
  ];
  if (includeOther) cells.push(`<td class="number-cell">${count(summary.other)}</td>`);
  cells.push(`<td><span class="rate-pill">${summary.savedRate}</span></td>`);
  cells.push(`<td><span class="rate-pill rate-pill-fallen">${summary.fallenRate}</span></td>`);
  const dayMoments = [
    summary.dayCounts.savedWithin,
    summary.dayCounts.fallenWithin,
    summary.dayCounts.otherWithin,
    summary.dayCounts.savedOutside,
    summary.dayCounts.fallenOutside,
    summary.dayCounts.otherOutside,
  ];
  const paymentMoments = [
    summary.paymentCounts.savedBefore,
    summary.paymentCounts.fallenBefore,
    summary.paymentCounts.otherBefore,
    summary.paymentCounts.savedAfter,
    summary.paymentCounts.fallenAfter,
    summary.paymentCounts.otherAfter,
  ];
  const moments = mode === "payment" ? paymentMoments
    : mode === "all" ? [...dayMoments, ...paymentMoments]
      : dayMoments;
  cells.push(...moments.map((value) => `<td class="number-cell">${count(value)}</td>`));
  return cells.join("");
}

function momentHeaders(mode = $("momentMode").value) {
  return [
    "Registros",
    "Salvadas",
    "Caídas",
    "Otras gestiones",
    "Tasa de salvadas",
    "Tasa de caídas",
    ...(mode === "payment"
      ? [
        "Salvadas · Antes del primer pago",
        "Caídas · Antes del primer pago",
        "Otras gestiones · Antes del primer pago",
        "Salvadas · Después del primer pago",
        "Caídas · Después del primer pago",
        "Otras gestiones · Después del primer pago",
      ]
      : mode === "all"
        ? [
          "Salvadas · Dentro de 3 días",
          "Caídas · Dentro de 3 días",
          "Otras gestiones · Dentro de 3 días",
          "Salvadas · Fuera de 3 días",
          "Caídas · Fuera de 3 días",
          "Otras gestiones · Fuera de 3 días",
          "Salvadas · Antes del primer pago",
          "Caídas · Antes del primer pago",
          "Otras gestiones · Antes del primer pago",
          "Salvadas · Después del primer pago",
          "Caídas · Después del primer pago",
          "Otras gestiones · Después del primer pago",
        ]
        : [
        "Salvadas · Dentro de 3 días",
        "Caídas · Dentro de 3 días",
        "Otras gestiones · Dentro de 3 días",
        "Salvadas · Fuera de 3 días",
        "Caídas · Fuera de 3 días",
        "Otras gestiones · Fuera de 3 días",
        ]),
  ];
}

function momentHeaderMarkup(header) {
  const [label, moment] = header.split(" · ");
  return moment ? `${esc(label)}<br>${esc(moment)}` : esc(header);
}

function renderMomentView() {
  const mode = $("momentMode").value;
  const showDays = mode !== "payment";
  const showPayment = mode !== "days";
  $("withinMetricCard").hidden = !showDays;
  $("outsideMetricCard").hidden = !showDays;
  $("fallenWithinMetricCard").hidden = !showDays;
  $("fallenOutsideMetricCard").hidden = !showDays;
  $("otherWithinMetricCard").hidden = !showDays;
  $("otherOutsideMetricCard").hidden = !showDays;
  $("beforePaymentMetricCard").hidden = !showPayment;
  $("afterPaymentMetricCard").hidden = !showPayment;
  $("fallenBeforePaymentMetricCard").hidden = !showPayment;
  $("fallenAfterPaymentMetricCard").hidden = !showPayment;
  $("otherBeforePaymentMetricCard").hidden = !showPayment;
  $("otherAfterPaymentMetricCard").hidden = !showPayment;
  $("daysAnalysisPanel").hidden = !showDays;
  $("paymentAnalysisPanel").hidden = !showPayment;
}

function weeklyGroups(rows, month) {
  const groups = new Map();
  const isoDate = (date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
  const dateRows = new Map();
  rows.forEach((row) => {
    const date = row["FECHA GESTIÓN"] || "";
    const parsedDate = /^\d{4}-\d{2}-\d{2}$/.test(date) ? new Date(`${date}T12:00:00`) : null;
    const startDate = parsedDate ? new Date(parsedDate) : null;
    if (startDate) startDate.setDate(startDate.getDate() - (startDate.getDay() + 6) % 7);
    const weekStart = startDate ? isoDate(startDate) : "";
    const key = weekStart || "sin-fecha";
    if (!groups.has(key)) {
      groups.set(key, { weekStart: startDate, rows: [] });
    }
    groups.get(key).rows.push(row);
    if (date) {
      if (!dateRows.has(date)) dateRows.set(date, []);
      dateRows.get(date).push(row);
    }
  });
  const [year, monthNumber] = month.split("-").map(Number);
  const monthStart = monthNumber ? new Date(year, monthNumber - 1, 1, 12) : null;
  const monthEnd = monthNumber ? new Date(year, monthNumber, 0, 12) : null;
  if (!monthStart || !monthEnd) {
    return [...groups.values()].map((group) => ({
      label: "Sin fecha",
      ...group,
      summary: summarize(group.rows),
      dailyGroups: [],
    }));
  }
  const firstWeekStart = new Date(monthStart);
  firstWeekStart.setDate(firstWeekStart.getDate() - (firstWeekStart.getDay() + 6) % 7);
  const monthWeeks = [];
  for (let weekIndex = 0, weekStart = firstWeekStart; weekStart <= monthEnd; weekIndex += 1, weekStart.setDate(weekStart.getDate() + 7)) {
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekEnd.getDate() + 6);
    const visibleStart = weekStart < monthStart ? monthStart : weekStart;
    const visibleEnd = weekEnd > monthEnd ? monthEnd : weekEnd;
    const days = Math.floor((visibleEnd - visibleStart) / 86400000) + 1;
    const group = groups.get(isoDate(weekStart));
    const weekRows = group ? group.rows : [];
    const dailyGroups = [];
    for (let day = new Date(visibleStart); day <= visibleEnd; day.setDate(day.getDate() + 1)) {
      const date = isoDate(day);
      const dailyRows = dateRows.get(date) || [];
      dailyGroups.push({ date, rows: dailyRows, summary: summarize(dailyRows) });
    }
    monthWeeks.push({
      label: `Semana ${weekIndex + 1} (${dateLabel(isoDate(visibleStart))} – ${dateLabel(isoDate(visibleEnd))})`,
      rows: weekRows,
      summary: summarize(weekRows),
      dailyGroups,
    });
  }
  const undatedGroup = groups.get("sin-fecha");
  if (undatedGroup) {
    monthWeeks.push({ label: "Sin fecha", ...undatedGroup, summary: summarize(undatedGroup.rows) });
  }
  return monthWeeks;
}

function monthlyGroups(rows) {
  const groups = new Map();
  rows.forEach((row) => {
    const month = monthOf(row) || "Sin fecha";
    if (!groups.has(month)) groups.set(month, []);
    groups.get(month).push(row);
  });
  return [...groups.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([month, monthRows]) => ({ month, rows: monthRows, summary: summarize(monthRows) }));
}

function timelineTableMarkup(months) {
  const headers = momentHeaders();
  return `<div class="table-wrap"><table class="month-table">
    <thead><tr><th>Mes</th>${headers.map((header) => `<th>${momentHeaderMarkup(header)}</th>`).join("")}</tr></thead>
    <tbody>${months.map(({ month, summary, rows }) => {
      const weeks = weeklyGroups(rows, month);
      const weekRows = weeks.map(({ label, summary: weekSummary, dailyGroups }) => `
        <tr class="week-row" hidden><td><button class="week-toggle" type="button" aria-expanded="false" aria-label="Desplegar días de ${esc(label)}"><span class="month-chevron" aria-hidden="true">▸</span>${esc(label)}</button></td>${statsCells(weekSummary)}</tr>
        ${dailyGroups.map(({ date, summary: daySummary }) => `<tr class="day-row" hidden><td>${dateLabel(date)}</td>${statsCells(daySummary)}</tr>`).join("")}`).join("");
      return `<tr class="month-row"><td><button class="month-toggle" type="button" aria-expanded="false" aria-label="Desplegar semanas de ${esc(monthLabel(month))}"><span class="month-chevron" aria-hidden="true">▸</span>${esc(monthLabel(month))}</button></td>${statsCells(summary)}</tr>${weekRows}`;
    }).join("")}</tbody>
  </table></div>`;
}

function advisorTimelineRowsMarkup(rows) {
  return monthlyGroups(rows).map(({ month, summary, rows: monthRows }) => {
    const weeks = weeklyGroups(monthRows, month);
    const weekRows = weeks.map(({ label, summary: weekSummary, dailyGroups }) => `
      <tr class="week-row advisor-week-row" hidden><td><button class="week-toggle" type="button" aria-expanded="false" aria-label="Desplegar días de ${esc(label)}"><span class="month-chevron" aria-hidden="true">▸</span>${esc(label)}</button></td>${statsCells(weekSummary)}</tr>
      ${dailyGroups.map(({ date, summary: daySummary }) => `<tr class="day-row advisor-day-row" hidden><td>${dateLabel(date)}</td>${statsCells(daySummary)}</tr>`).join("")}`).join("");
    return `<tr class="month-row advisor-month-row" hidden><td><button class="month-toggle" type="button" aria-expanded="false" aria-label="Desplegar semanas de ${esc(monthLabel(month))}"><span class="month-chevron" aria-hidden="true">▸</span>${esc(monthLabel(month))}</button></td>${statsCells(summary)}</tr>${weekRows}`;
  }).join("");
}

function wireTimelineToggles(container) {
  container.querySelectorAll(".month-toggle").forEach((button) => {
    button.addEventListener("click", () => {
      let row = button.closest("tr").nextElementSibling;
      const expanded = button.getAttribute("aria-expanded") !== "true";
      while (row && !row.classList.contains("month-row") && !row.classList.contains("advisor-row")) {
        row.hidden = row.classList.contains("day-row") || !expanded;
        if (row.classList.contains("week-row")) {
          const weekButton = row.querySelector(".week-toggle");
          weekButton.setAttribute("aria-expanded", "false");
          weekButton.querySelector(".month-chevron").style.transform = "";
        }
        row = row.nextElementSibling;
      }
      button.setAttribute("aria-expanded", String(expanded));
      button.querySelector(".month-chevron").style.transform = expanded ? "rotate(90deg)" : "";
    });
  });
  container.querySelectorAll(".week-toggle").forEach((button) => {
    button.addEventListener("click", () => {
      let row = button.closest("tr").nextElementSibling;
      const expanded = button.getAttribute("aria-expanded") !== "true";
      while (row && row.classList.contains("day-row")) {
        row.hidden = !expanded;
        row = row.nextElementSibling;
      }
      button.setAttribute("aria-expanded", String(expanded));
      button.querySelector(".month-chevron").style.transform = expanded ? "rotate(90deg)" : "";
    });
  });
}

function renderComparison(rows) {
  const headers = momentHeaders();
  $("companyAreaHead").innerHTML = `<tr><th>Empresa</th><th>Área</th>${headers.map((header) => `<th>${momentHeaderMarkup(header)}</th>`).join("")}</tr>`;
  const areaGroups = groupedRows(rows, ["EMPRESA", "ÁREA"])
    .sort((a, b) => a.values.join(" ").localeCompare(b.values.join(" "), "es"));
  $("companyAreaTable").innerHTML = areaGroups.length
    ? areaGroups.map(({ values, summary }) => `<tr><td>${esc(values[0])}</td><td>${esc(values[1])}</td>${statsCells(summary)}</tr>`).join("")
    : `<tr><td colspan="${headers.length + 2}" class="empty-cell">Sin registros para comparar.</td></tr>`;

  const advisorGroups = new Map();
  groupedRows(rows, ["VENDEDOR", "EMPRESA", "ÁREA"]).forEach((advisor) => {
    const key = `${advisor.values[1]}\u0000${advisor.values[2]}`;
    if (!advisorGroups.has(key)) {
      advisorGroups.set(key, { company: advisor.values[1], area: advisor.values[2], advisors: [] });
    }
    advisorGroups.get(key).advisors.push(advisor);
  });
  const sortedAdvisorGroups = [...advisorGroups.values()]
    .sort((a, b) => a.company.localeCompare(b.company, "es") || a.area.localeCompare(b.area, "es"));
  $("advisorGroups").innerHTML = sortedAdvisorGroups.length
    ? sortedAdvisorGroups.map(({ company, area, advisors: groupAdvisors }) => `
      <section class="advisor-group">
        <h3>${esc(company)} · ${esc(area)}</h3>
        <div class="table-wrap">
          <table class="advisor-table">
            <thead><tr><th>Asesor</th>${headers.map((header) => `<th>${momentHeaderMarkup(header)}</th>`).join("")}</tr></thead>
            <tbody>${groupAdvisors
              .sort((a, b) => b.summary.total - a.summary.total || a.values[0].localeCompare(b.values[0], "es"))
              .map(({ values, summary, rows: advisorRows }) => `
                <tr class="advisor-row"><td class="advisor-name"><button class="advisor-toggle" type="button" aria-expanded="false" aria-label="Desplegar detalle mensual de ${esc(values[0])}"><span class="month-chevron" aria-hidden="true">▸</span>${esc(values[0])}</button></td>${statsCells(summary)}</tr>
                ${advisorTimelineRowsMarkup(advisorRows)}`).join("")}
            </tbody>
          </table>
        </div>
      </section>`).join("")
    : '<p class="empty">Sin asesores para comparar.</p>';
  $("advisorGroups").querySelectorAll(".advisor-toggle").forEach((button) => {
    button.addEventListener("click", () => {
      const expanded = button.getAttribute("aria-expanded") !== "true";
      let row = button.closest("tr").nextElementSibling;
      while (row && !row.classList.contains("advisor-row")) {
        row.hidden = row.classList.contains("advisor-month-row") ? !expanded : true;
        if (row.classList.contains("month-row") || row.classList.contains("week-row")) {
          const nestedButton = row.querySelector(".month-toggle, .week-toggle");
          nestedButton.setAttribute("aria-expanded", "false");
          nestedButton.querySelector(".month-chevron").style.transform = "";
        }
        row = row.nextElementSibling;
      }
      button.setAttribute("aria-expanded", String(expanded));
      button.querySelector(".month-chevron").style.transform = expanded ? "rotate(90deg)" : "";
    });
  });

  const byMonth = new Map();
  rows.forEach((row) => {
    const month = monthOf(row) || "Sin fecha";
    const key = `${month}\u0000${row.EMPRESA || "Sin dato"}\u0000${row["ÁREA"] || "Sin dato"}`;
    if (!byMonth.has(key)) byMonth.set(key, { month, values: [row.EMPRESA || "Sin dato", row["ÁREA"] || "Sin dato"], rows: [] });
    byMonth.get(key).rows.push(row);
  });
  const monthGroups = [...byMonth.values()].map((group) => ({ ...group, summary: summarize(group.rows) }))
    .sort((a, b) => a.values[0].localeCompare(b.values[0], "es")
      || a.values[1].localeCompare(b.values[1], "es")
      || a.month.localeCompare(b.month));
  const groupsByCompanyArea = new Map();
  monthGroups.forEach((group) => {
    const key = `${group.values[0]}\u0000${group.values[1]}`;
    if (!groupsByCompanyArea.has(key)) {
      groupsByCompanyArea.set(key, { company: group.values[0], area: group.values[1], months: [] });
    }
    groupsByCompanyArea.get(key).months.push(group);
  });
  $("monthDetails").innerHTML = groupsByCompanyArea.size
    ? [...groupsByCompanyArea.values()].map(({ company, area, months }) => `
      <section class="month-group">
        <h3>${esc(company)} · ${esc(area)}</h3>
        ${timelineTableMarkup(months)}
      </section>`).join("")
    : '<p class="empty">Sin periodos para mostrar.</p>';
  wireTimelineToggles($("monthDetails"));
  wireTimelineToggles($("advisorGroups"));
}

function renderDetails(rows) {
  const sorted = [...rows].sort((a, b) => (b["FECHA GESTIÓN"] || "").localeCompare(a["FECHA GESTIÓN"] || "")
    || (a["FILA EXCEL"] || 0) - (b["FILA EXCEL"] || 0));
  $("detailCount").textContent = `${count(rows.length)} registros`;
  $("detailTable").innerHTML = sorted.length ? sorted.map((row) => {
    const state = statusOf(row);
    const colorIndex = CLASSIFICATIONS.indexOf(state);
    const management = [row["REGISTRO 1"], row["REGISTRO 2"], row["REGISTRO 3"]].filter(Boolean).join(" · ") || "—";
    return `<tr>
      <td class="nowrap">${dateLabel(row["FECHA GESTIÓN"])}</td>
      <td>${esc(row.EMPRESA || "—")}</td>
      <td>${esc(row["ÁREA"] || "—")}</td>
      <td>${esc(row.VENDEDOR || "—")}</td>
      <td><b>${esc(row.CLIENTE || "No registrado en el Excel")}</b><small class="subcell">DNI ${esc(row.DNI || "No registrado")}</small></td>
      <td>${esc(row.TELÉFONO || "No registrado en el Excel")}</td>
      <td class="management-cell">${esc(management)}<small class="subcell">${esc(row["COMUNICACIÓN DE VENTA"] || "No registrada en el Excel")}</small></td>
      <td><span class="status-pill status-${colorIndex}">${esc(state)}</span></td>
      <td class="nowrap">${dateLabel(row["FECHA DE ACEPTACIÓN DE PRODUCTO"])}</td>
      <td>${esc(row["MOMENTO DÍAS"] || "No registrado en el Excel")}</td>
      <td>${esc(row["MOMENTO PRIMER PAGO"] || "No registrado en el Excel")}</td>
      <td class="nowrap">${dateLabel(row["FECHA PRIMER PAGO"])}</td>
    </tr>`;
  }).join("") : '<tr><td colspan="12" class="empty-cell">No hay registros que coincidan con los filtros.</td></tr>';
}

function render() {
  currentRows = selectRows();
  const summary = summarize(currentRows);
  renderMomentView();
  renderMetrics(currentRows);
  renderDonut(currentRows);
  renderMonthlyChart(currentRows);
  renderBreakdown("daysBreakdown", summary.days, summary.total, "days");
  renderBreakdown("paymentBreakdown", summary.payments, summary.total, "payment");
  renderAreaMomentCharts(selectRows({ ignoreArea: true, ignoreClassification: true }));
  renderCompanyOutcomeAreaChart(selectRows({ ignoreArea: true, ignoreClassification: true }));
  renderComparison(currentRows);
  renderDetails(currentRows);
}

async function loadDefault() {
  try {
    const response = await request("/api/data");
    setRows(await response.json());
  } catch (error) {
    showNotice(error.message, true);
  }
}

$("workbook").addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  const form = new FormData();
  form.append("file", file);
  showNotice("Procesando el Excel y preparando las comparaciones…");
  $("workbook").disabled = true;
  try {
    const response = await request("/api/analyze", { method: "POST", body: form });
    setRows(await response.json());
  } catch (error) {
    showNotice(error.message, true);
  } finally {
    $("workbook").disabled = false;
    event.target.value = "";
  }
});

for (const id of ["company", "area", "advisor", "classification", "momentMode", "month", "dateFrom", "dateTo", "search"]) {
  $(id).addEventListener("input", render);
  $(id).addEventListener("change", render);
}

$("clear").addEventListener("click", () => {
  for (const id of ["company", "area", "advisor", "classification", "month", "dateFrom", "dateTo", "search"]) {
    $(id).value = "";
  }
  $("momentMode").value = "all";
  render();
});

$("export").addEventListener("click", () => {
  const filter = filters();
  if (filter.dateFrom && filter.dateTo && filter.dateFrom > filter.dateTo) {
    showNotice("La fecha inicial no puede ser posterior a la fecha final.", true);
    return;
  }
  const params = new URLSearchParams({
    company: filter.company,
    area: filter.area,
    advisor: filter.advisor,
    classification: filter.classification,
    month: filter.month,
    date_from: filter.dateFrom,
    date_to: filter.dateTo,
    q: $("search").value.trim(),
  });
  window.location.href = `/api/export?${params}`;
});

loadDefault();
