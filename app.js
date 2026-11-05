(() => {
  "use strict";

  const $ = (selector) => document.querySelector(selector);
  const state = {
    fileName: "",
    rows: [],
    columns: [],
    types: {},
    results: [],
    chart: null,
    workbook: null,
    filter: null,
    sort: null,
    webR: null,
  };

  const numericTypes = new Set(["numeric", "ordinal"]);
  const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[char]);
  const fmt = (value, digits = 3) => Number.isFinite(value)
    ? new Intl.NumberFormat("tr-TR", { maximumFractionDigits: digits }).format(value)
    : "—";
  const fmtP = (value) => Number.isFinite(value) && value < 0.0001 ? "<0,0001" : fmt(value, 4);
  const isMissing = (value) => value === null || value === undefined || String(value).trim() === "";
  const parseNumber = (value) => {
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    if (isMissing(value)) return null;
    let text = String(value).trim().replace(/[\s\u00a0]/g, "");
    const comma = text.lastIndexOf(",");
    const dot = text.lastIndexOf(".");
    if (comma >= 0 && dot >= 0) {
      text = comma > dot ? text.replace(/\./g, "").replace(",", ".") : text.replace(/,/g, "");
    } else if (comma >= 0) {
      text = text.replace(",", ".");
    }
    const number = Number(text);
    return Number.isFinite(number) ? number : null;
  };
  const columnValues = (column) => activeRows().map((row) => row[column]);
  const numericColumns = () => state.columns.filter((column) => numericTypes.has(state.types[column]));
  const categoricalColumns = () => state.columns.filter((column) => !numericTypes.has(state.types[column]));
  const activeRows = () => {
    let rows = state.rows;
    if (state.filter) {
      const { column, operator, value } = state.filter;
      rows = rows.filter((row) => {
        const cell = row[column];
        const empty = isMissing(cell);
        const leftNumber = parseNumber(cell);
        const rightNumber = parseNumber(value);
        switch (operator) {
          case "equals": return String(cell ?? "") === value;
          case "notEquals": return String(cell ?? "") !== value;
          case "contains": return String(cell ?? "").toLocaleLowerCase("tr").includes(value.toLocaleLowerCase("tr"));
          case "gt": return leftNumber !== null && rightNumber !== null && leftNumber > rightNumber;
          case "gte": return leftNumber !== null && rightNumber !== null && leftNumber >= rightNumber;
          case "lt": return leftNumber !== null && rightNumber !== null && leftNumber < rightNumber;
          case "lte": return leftNumber !== null && rightNumber !== null && leftNumber <= rightNumber;
          case "empty": return empty;
          case "notEmpty": return !empty;
          default: return true;
        }
      });
    }
    if (state.sort) {
      const { column, direction } = state.sort;
      rows = rows.slice().sort((a, b) => {
        const an = parseNumber(a[column]); const bn = parseNumber(b[column]);
        const comparison = an !== null && bn !== null
          ? an - bn
          : String(a[column] ?? "").localeCompare(String(b[column] ?? ""), "tr");
        return direction === "desc" ? -comparison : comparison;
      });
    }
    return rows;
  };

  function showError(message) {
    const error = $("#uploadError");
    error.textContent = message;
    error.hidden = false;
  }

  function clearError() {
    $("#uploadError").hidden = true;
    $("#uploadError").textContent = "";
  }

  function uniqueHeaders(headers) {
    const seen = new Map();
    return headers.map((header, index) => {
      const base = String(header ?? "").trim() || `Değişken ${index + 1}`;
      const count = (seen.get(base) || 0) + 1;
      seen.set(base, count);
      return count === 1 ? base : `${base} (${count})`;
    });
  }

  function normalizeMatrix(matrix) {
    const usefulRows = matrix.filter((row) => row.some((value) => !isMissing(value)));
    if (usefulRows.length < 2) throw new Error("Dosyada başlık satırı ve en az bir veri satırı bulunmalıdır.");
    const headers = uniqueHeaders(usefulRows[0]);
    if (!headers.length) throw new Error("Dosyada sütun bulunamadı.");
    const rows = usefulRows.slice(1).map((cells) => {
      const row = {};
      headers.forEach((header, index) => { row[header] = cells[index] ?? ""; });
      return row;
    });
    return { headers, rows };
  }

  function inferType(column) {
    const values = columnValues(column).filter((value) => !isMissing(value));
    if (!values.length) return "nominal";
    const numericCount = values.filter((value) => parseNumber(value) !== null).length;
    return numericCount / values.length >= 0.9 ? "numeric" : "nominal";
  }

  function acceptDataset(fileName, data) {
    if (!data.headers.length || !data.rows.length) throw new Error("Analiz edilebilecek veri bulunamadı.");
    state.fileName = fileName;
    state.columns = data.headers;
    state.rows = data.rows;
    state.filter = null;
    state.sort = null;
    if (state.chart) {
      state.chart.destroy();
      state.chart = null;
    }
    state.types = {};
    state.columns.forEach((column) => { state.types[column] = inferType(column); });
    state.results = [];
    $("#downloadResults").disabled = true;
    $("#downloadWorkbook").disabled = true;
    $("#downloadReport").disabled = true;
    $("#downloadChart").disabled = true;
    $("#printReport").hidden = true;
    $("#runR").disabled = !state.webR;
    $("#rStatus").textContent = state.webR ? "R hazır · veri yüklenecek" : "R başlatılmadı";
    $("#datasetName").textContent = fileName;
    $("#workspace").hidden = false;
    $("#uploadPanel").hidden = true;
    document.body.classList.add("dataset-open");
    clearError();
    updateSummary();
    renderDataTools();
    renderDataTable();
    renderHistory();
    renderAnalysisFields();
    renderChartFields();
    renderRAnalysisFields();
    renderResultsEmpty();
    switchTab("data");
  }

  async function loadFile(file) {
    if (!file) return;
    const extension = file.name.split(".").pop().toLowerCase();
    clearError();
    try {
      if (extension === "csv") {
        state.workbook = null;
        $("#sheetControl").hidden = true;
        if (!window.Papa) throw new Error("CSV okuyucusu yüklenemedi. İnternet bağlantınızı kontrol edip sayfayı yenileyin.");
        const text = await file.text();
        const parsed = Papa.parse(text, { skipEmptyLines: "greedy", dynamicTyping: false });
        if (parsed.errors.length) throw new Error(`CSV okunamadı: ${parsed.errors[0].message}`);
        acceptDataset(file.name, normalizeMatrix(parsed.data));
      } else if (extension === "xlsx" || extension === "xls") {
        if (!window.XLSX) throw new Error("Excel okuyucusu yüklenemedi. İnternet bağlantınızı kontrol edip sayfayı yenileyin.");
        const bytes = await file.arrayBuffer();
        state.workbook = XLSX.read(bytes, { type: "array", cellDates: true });
        if (!state.workbook.SheetNames.length) throw new Error("Excel dosyasında çalışma sayfası bulunamadı.");
        $("#sheetControl").hidden = state.workbook.SheetNames.length < 2;
        $("#sheetSelect").innerHTML = state.workbook.SheetNames.map((name, index) =>
          `<option value="${index}">${escapeHtml(name)}</option>`).join("");
        readSheet(0, file.name);
      } else {
        throw new Error("Lütfen CSV, XLSX veya XLS biçiminde bir dosya seçin.");
      }
    } catch (error) {
      showError(error.message || "Dosya okunurken beklenmeyen bir hata oluştu.");
    }
  }

  function readSheet(index, fileName = state.fileName) {
    try {
      const sheetName = state.workbook.SheetNames[index];
      const matrix = XLSX.utils.sheet_to_json(state.workbook.Sheets[sheetName], {
        header: 1, defval: "", raw: false, blankrows: false,
      });
      acceptDataset(fileName, normalizeMatrix(matrix));
    } catch (error) {
      showError(error.message || "Çalışma sayfası okunamadı.");
    }
  }

  function updateSummary() {
    const rows = activeRows();
    const activeMissing = rows.reduce((total, row) =>
      total + state.columns.filter((column) => isMissing(row[column])).length, 0);
    $("#datasetMeta").textContent = `${rows.length.toLocaleString("tr-TR")} / ${state.rows.length.toLocaleString("tr-TR")} satır · ${state.columns.length} sütun`;
    $("#rowCount").textContent = rows.length.toLocaleString("tr-TR");
    $("#columnCount").textContent = state.columns.length.toLocaleString("tr-TR");
    $("#numericCount").textContent = numericColumns().length.toLocaleString("tr-TR");
    $("#missingCount").textContent = activeMissing.toLocaleString("tr-TR");
    $("#activeRowsLabel").textContent = `${rows.length.toLocaleString("tr-TR")} / ${state.rows.length.toLocaleString("tr-TR")} satır`;
    $("#overviewRows").textContent = rows.length.toLocaleString("tr-TR");
    $("#overviewColumns").textContent = state.columns.length.toLocaleString("tr-TR");
    $("#overviewMissing").textContent = activeMissing.toLocaleString("tr-TR");
  }

  function renderDataTools() {
    const options = state.columns.map((column) =>
      `<option value="${escapeHtml(column)}">${escapeHtml(column)}</option>`).join("");
    $("#filterColumn").innerHTML = options;
    $("#sortColumn").innerHTML = options;
    $("#filterValue").disabled = ["empty", "notEmpty"].includes($("#filterOperator").value);
  }

  function renderDataTable() {
    const table = $("#dataTable");
    const options = [["numeric", "Sayısal"], ["ordinal", "Sıralı"], ["nominal", "Nominal"], ["id", "Kimlik"]];
    const header = `<thead><tr class="data-header"><th class="row-number"></th>${state.columns.map((column) =>
      `<th data-type="${escapeHtml(state.types[column])}"><span class="column-name">${escapeHtml(column)}</span><select class="type-select" data-column="${escapeHtml(column)}" aria-label="${escapeHtml(column)} değişken türü">${options.map(([value, label]) =>
        `<option value="${value}"${state.types[column] === value ? " selected" : ""}>${label}</option>`).join("")}</select></th>`).join("")}</tr></thead>`;
    const rows = activeRows();
    const shown = rows.slice(0, 100);
    const body = `<tbody>${shown.map((row) => `<tr><td class="row-number">${state.rows.indexOf(row) + 1}</td>${state.columns.map((column) => {
      const value = row[column];
      const rowIndex = state.rows.indexOf(row);
      return `<td contenteditable="true" class="${isMissing(value) ? "missing-value" : ""}" data-row="${rowIndex}" data-column="${escapeHtml(column)}" aria-label="${escapeHtml(column)} verisi">${isMissing(value) ? "" : escapeHtml(value)}</td>`;
    }).join("")}</tr>`).join("")}</tbody>`;
    table.innerHTML = header + body;
    $("#previewLabel").textContent = `İlk ${shown.length} / ${rows.length} satır`;
    table.querySelectorAll(".type-select").forEach((select) => {
      select.addEventListener("change", () => {
        state.types[select.dataset.column] = select.value;
        updateSummary();
        renderAnalysisFields();
        renderChartFields();
      });
    });
    table.querySelectorAll("td[contenteditable=true]").forEach((cell) => {
      cell.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          cell.blur();
        }
      });
      cell.addEventListener("blur", () => {
        const row = state.rows[Number(cell.dataset.row)];
        if (!row) return;
        const value = cell.textContent.trim();
        row[cell.dataset.column] = value;
        cell.textContent = value;
        cell.classList.toggle("missing-value", !value);
        updateSummary();
        updateRIfEnabled();
      });
    });
  }

  function updateRIfEnabled() {
    if (state.webR && $("#autoUpdateR").checked) syncRDataset()
      .then(() => { $("#rStatus").textContent = "R verisi güncellendi"; })
      .catch((error) => { $("#rStatus").textContent = error.message || "R verisi güncellenemedi"; });
  }

  function parseExpression(expression) {
    expression = expression.trim();
    const tokens = [];
    const pattern = /\s*(?:\[([^\]]+)\]|((?:\d+(?:[.,]\d*)?|[.,]\d+)(?:e[+-]?\d+)?)|([()+\-*/]))/iy;
    let position = 0;
    while (position < expression.length) {
      pattern.lastIndex = position;
      const match = pattern.exec(expression);
      if (!match) throw new Error(`İfade çözümlenemedi: "${expression.slice(position).trim()}"`);
      position = pattern.lastIndex;
      if (match[1]) {
        const name = match[1].trim();
        if (!state.columns.includes(name)) throw new Error(`"${name}" adlı sütun bulunamadı.`);
        if (!numericTypes.has(state.types[name])) throw new Error(`"${name}" sayısal bir sütun değil.`);
        tokens.push({ type: "column", value: name });
      } else if (match[2]) tokens.push({ type: "number", value: Number(match[2].replace(",", ".")) });
      else tokens.push({ type: "operator", value: match[3] });
    }
    if (!tokens.length) throw new Error("Hesaplama ifadesi boş olamaz.");

    let cursor = 0;
    const primary = () => {
      const token = tokens[cursor++];
      if (!token) throw new Error("İfade beklenmedik biçimde sona erdi.");
      if (token.type === "number" || token.type === "column") return token;
      if (token.value === "+" || token.value === "-") return { type: "unary", operator: token.value, child: primary() };
      if (token.value === "(") {
        const node = addition();
        if (tokens[cursor]?.value !== ")") throw new Error("Parantezler eşleşmiyor.");
        cursor++;
        return node;
      }
      throw new Error("Sayı, sütun adı veya parantez bekleniyordu.");
    };
    const multiplication = () => {
      let node = primary();
      while (tokens[cursor]?.value === "*" || tokens[cursor]?.value === "/") {
        const operator = tokens[cursor++].value;
        node = { type: "binary", operator, left: node, right: primary() };
      }
      return node;
    };
    const addition = () => {
      let node = multiplication();
      while (tokens[cursor]?.value === "+" || tokens[cursor]?.value === "-") {
        const operator = tokens[cursor++].value;
        node = { type: "binary", operator, left: node, right: multiplication() };
      }
      return node;
    };
    const root = addition();
    if (cursor !== tokens.length) throw new Error("İfadede beklenmeyen bir öğe var.");
    return root;
  }

  function evaluateExpression(node, row) {
    if (node.type === "number") return node.value;
    if (node.type === "column") return parseNumber(row[node.value]);
    if (node.type === "unary") {
      const value = evaluateExpression(node.child, row);
      return value === null ? null : node.operator === "-" ? -value : value;
    }
    const left = evaluateExpression(node.left, row);
    const right = evaluateExpression(node.right, row);
    if (left === null || right === null || (node.operator === "/" && right === 0)) return null;
    const value = node.operator === "+" ? left + right
      : node.operator === "-" ? left - right
        : node.operator === "*" ? left * right : left / right;
    return Number.isFinite(value) ? value : null;
  }

  function addComputedColumn() {
    const error = $("#computedError");
    error.hidden = true;
    try {
      const name = $("#computedName").value.trim();
      if (!name) throw new Error("Yeni değişken için bir sütun adı girin.");
      if (state.columns.includes(name)) throw new Error("Bu sütun adı zaten kullanılıyor.");
      const tree = parseExpression($("#computedExpression").value);
      state.rows.forEach((row) => { row[name] = evaluateExpression(tree, row) ?? ""; });
      state.columns.push(name);
      state.types[name] = "numeric";
      $("#computedName").value = "";
      updateSummary();
      renderDataTools();
      renderDataTable();
      renderAnalysisFields();
      renderChartFields();
      updateRIfEnabled();
    } catch (caught) {
      error.textContent = caught.message || "Yeni değişken oluşturulamadı.";
      error.hidden = false;
    }
  }

  function addDataRow() {
    const row = {};
    state.columns.forEach((column) => { row[column] = ""; });
    state.rows.push(row);
    updateSummary();
    renderDataTable();
    updateRIfEnabled();
  }

  function downloadData(format) {
    const rows = state.filter ? activeRows() : state.rows;
    if (format === "excel") {
      if (!window.XLSX) throw new Error("Excel dışa aktarma kütüphanesi yüklenemedi.");
      const worksheet = XLSX.utils.json_to_sheet(rows, { header: state.columns });
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Veri");
      XLSX.writeFile(workbook, "biz-bize-veri-seti.xlsx", { compression: true });
      return;
    }
    const csvRows = rows.map((row) => state.columns.map((column) => row[column] ?? ""));
    const csv = Papa.unparse([state.columns, ...csvRows], { delimiter: ";", quotes: true });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" }));
    link.download = "biz-bize-veri-seti.csv";
    document.body.append(link);
    link.click();
    const url = link.href;
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function selectField(id, label, columns, selected = "", includeNone = false) {
    const options = columns.map((column) =>
      `<option value="${escapeHtml(column)}"${column === selected ? " selected" : ""}>${escapeHtml(column)}</option>`).join("");
    return `<div class="field-wrap"><label for="${id}">${label}</label><select id="${id}" class="select-control">${includeNone ? '<option value="">Seçin</option>' : ""}${options}</select></div>`;
  }

  function groupField(id, selected = "") {
    const options = state.columns.map((column) => {
      const levels = new Set(activeRows().map((row) => row[column]).filter((value) => !isMissing(value)).map(String)).size;
      return `<option value="${escapeHtml(column)}"${column === selected ? " selected" : ""}>${escapeHtml(column)} · ${levels} kategori</option>`;
    }).join("");
    return `<div class="field-wrap"><label for="${id}">GRUP DEĞİŞKENİ (SAYISAL KODLAR DAHİL)</label><select id="${id}" class="select-control"><option value="">Seçin</option>${options}</select></div>`;
  }

  function checklistField(id, label, columns, selected = [], hint = "") {
    const items = columns.map((column) =>
      `<label class="choice-item"><input type="checkbox" name="${id}" value="${escapeHtml(column)}"${selected.includes(column) ? " checked" : ""}><span>${escapeHtml(column)}</span></label>`).join("");
    return `<fieldset class="field-wrap checklist-field"><legend>${label}</legend><div class="checklist-heading"><button class="text-button" type="button" data-select-all="${id}">Tümünü seç</button><button class="text-button" type="button" data-clear-all="${id}">Temizle</button></div>${hint ? `<span class="field-hint">${hint}</span>` : ""}<div class="choice-list">${items || '<span class="field-hint">Uygun sütun bulunamadı.</span>'}</div></fieldset>`;
  }

  function checkedValues(name) {
    return [...document.querySelectorAll(`input[name="${name}"]:checked`)].map((input) => input.value);
  }

  function renderAnalysisFields() {
    const type = $("#analysisType").value;
    const numeric = numericColumns();
    let markup = "";
    if (type === "descriptive") {
      markup = checklistField("variable", "DEĞİŞKENLER", state.columns, state.columns.slice(0, 1), "Birden fazla sütun seçebilirsiniz.");
    } else if (["ttest", "anova", "mannwhitney", "kruskal"].includes(type)) {
      markup = selectField("outcome", "SONUÇ DEĞİŞKENİ (SAYISAL)", numeric, numeric[0] || "", true);
      markup += groupField("group");
      markup += '<span class="field-hint">Sonuç değişkeninden farklı bir grup seçin. Sayısal kodlu grup sütunları da kullanılabilir.</span>';
      if (type === "ttest" || type === "mannwhitney") markup += '<div id="groupLevelsField"></div>';
      if (type === "anova") {
        markup += '<label class="checkbox-field"><input id="tukeyOption" type="checkbox"> Tukey HSD ikili karşılaştırmaları</label>';
      }
    } else if (type === "regression") {
      markup = selectField("dependent", "BAĞIMLI DEĞİŞKEN", numeric, numeric[0] || "", true);
      markup += checklistField("predictors", "YORDAYICILAR", state.columns.filter((column) => column !== numeric[0]), numeric.slice(1, 2), "Sayısal sütunlar sürekli; nominal sütunlar referans kategoriye göre kukla değişken olarak modele eklenir.");
      markup += '<label class="checkbox-field"><input id="regressionDependentCheck" type="checkbox" checked> Sabit terim ekle</label>';
    } else if (type === "chisquare") {
      markup = selectField("categoryX", "SATIR DEĞİŞKENİ", state.columns, "", true);
      markup += selectField("categoryY", "SÜTUN DEĞİŞKENİ", state.columns, "", true);
      markup += '<span class="field-hint">Sayısal kodlu kategoriler kullanılabilir. Sürekli ölçümlerde önce kategorik gruplama yapın.</span>';
    } else if (type === "correlation") {
      markup = checklistField("correlationVariables", "SAYISAL DEĞİŞKENLER", numeric, numeric.slice(0, 2), "En az iki değişken seçin; tüm çiftler için Pearson korelasyon matrisi hesaplanır.");
    } else if (type === "spearman") {
      markup = checklistField("correlationVariables", "SIRALI / SAYISAL DEĞİŞKENLER", numeric, numeric.slice(0, 2), "En az iki değişken seçin; Spearman sıra korelasyonu hesaplanır.");
    } else if (type === "pairedT" || type === "wilcoxon") {
      markup = selectField("variableX", "1. ÖLÇÜM", numeric, numeric[0] || "", true);
      markup += selectField("variableY", "2. ÖLÇÜM", numeric, numeric[1] || "", true);
    } else if (type === "oneSampleT") {
      markup = selectField("outcome", "SAYISAL DEĞİŞKEN", numeric, numeric[0] || "", true);
      markup += '<div class="field-wrap"><label for="testValue">KARŞILAŞTIRMA DEĞERİ</label><input id="testValue" class="text-control" type="number" step="any" value="0"></div>';
    } else if (type === "repeatedAnova") {
      markup = checklistField("repeatedVariables", "TEKRARLI ÖLÇÜM SÜTUNLARI (SIRALI)", numeric, numeric.slice(0, 3), "Her satır bir katılımcı; seçilen sütunlar ölçüm zamanlarıdır.");
    } else if (type === "logistic") {
      markup = selectField("binaryOutcome", "İKİ KATEGORİLİ SONUÇ", state.columns, "", true);
      markup += checklistField("logitPredictors", "YORDAYICILAR", state.columns, numeric.slice(0, 1), "Sayısal ve kategorik değişkenler desteklenir.");
    } else if (type === "reliability") {
      markup = checklistField("reliabilityVariables", "ÖLÇEK MADDELERİ", numeric, numeric.slice(0, 5), "En az iki sayısal madde seçin.");
    } else if (type === "pca") {
      markup = checklistField("pcaVariables", "SAYISAL DEĞİŞKENLER", numeric, numeric.slice(0, Math.min(5, numeric.length)), "En az iki değişken seçin.");
    } else {
      markup = selectField("variableX", "1. SAYISAL DEĞİŞKEN", numeric, numeric[0] || "", true);
      markup += selectField("variableY", "2. SAYISAL DEĞİŞKEN", numeric, numeric[1] || numeric[0] || "", true);
    }
    $("#analysisFields").innerHTML = markup || '<p class="analysis-hint">Bu analiz için uygun sütun bulunamadı. Veri tablosundan sütun türünü değiştirin.</p>';
    const dependent = $("#dependent");
    if (dependent) {
      dependent.addEventListener("change", () => {
        document.querySelectorAll('input[name="predictors"]').forEach((input) => {
          input.disabled = input.value === dependent.value;
          if (input.disabled) input.checked = false;
        });
      });
      dependent.dispatchEvent(new Event("change"));
    }
    const outcome = $("#outcome");
    const group = $("#group");
    if (outcome && group) {
      const renderGroupLevelChoices = (preserve = true) => {
        const container = $("#groupLevelsField");
        if (!container) return;
        const selected = preserve ? checkedValues("groupLevels") : [];
        const validRows = activeRows().filter((row) => parseNumber(row[outcome.value]) !== null && !isMissing(row[group.value]));
        const counts = new Map();
        validRows.forEach((row) => {
          const value = String(row[group.value]);
          counts.set(value, (counts.get(value) || 0) + 1);
        });
        const levels = [...counts.entries()];
        const checkedCount = selected.filter((value) => counts.has(value)).length;
        container.innerHTML = `<fieldset class="field-wrap checklist-field group-levels"><legend>İKİ GRUP SEÇİN <span id="groupLevelCount" class="selected-count">${checkedCount ? `${checkedCount} seçili` : "henüz seçilmedi"}</span></legend><label class="group-level-search"><span aria-hidden="true">⌕</span><input id="groupLevelSearch" class="text-control" type="search" placeholder="Kategori ara…" aria-label="Grup kategorilerinde ara"></label><p class="field-hint">t-testi ve Mann–Whitney için tam iki kategori seçin. Seçim yapmazsanız yalnızca sütunda tam iki kategori varsa otomatik kullanılır.</p><div class="choice-list group-level-list">${levels.map(([level, count]) => `<label class="choice-item"><input type="checkbox" name="groupLevels" value="${escapeHtml(level)}"${selected.includes(level) ? " checked" : ""}><span>${escapeHtml(level)}</span><small>${count} gözlem</small></label>`).join("") || '<span class="field-hint">Sonuç değişkeninde geçerli gözlemi olan kategori bulunamadı.</span>'}</div></fieldset>`;
        const updateSelectionCount = () => {
          const current = checkedValues("groupLevels").length;
          $("#groupLevelCount").textContent = current ? `${current} seçili` : "henüz seçilmedi";
          $("#groupLevelCount").classList.toggle("selection-invalid", current > 0 && current !== 2);
        };
        container.querySelectorAll('input[name="groupLevels"]').forEach((input) => input.addEventListener("change", updateSelectionCount));
        $("#groupLevelSearch")?.addEventListener("input", (event) => {
          const query = event.target.value.trim().toLocaleLowerCase("tr");
          container.querySelectorAll(".group-level-list .choice-item").forEach((item) => {
            item.hidden = !item.querySelector("span").textContent.toLocaleLowerCase("tr").includes(query);
          });
        });
        updateSelectionCount();
      };
      outcome.addEventListener("change", () => {
        [...group.options].forEach((option) => {
          option.disabled = option.value !== "" && option.value === outcome.value;
          if (option.disabled && option.selected) group.value = "";
        });
        renderGroupLevelChoices(true);
      });
      group.addEventListener("change", () => {
        renderGroupLevelChoices(false);
      });
      outcome.dispatchEvent(new Event("change"));
      renderGroupLevelChoices(true);
    }
    const binaryOutcome = $("#binaryOutcome");
    if (binaryOutcome) {
      binaryOutcome.addEventListener("change", () => {
        document.querySelectorAll('input[name="logitPredictors"]').forEach((input) => {
          input.disabled = input.value === binaryOutcome.value;
          if (input.disabled) input.checked = false;
        });
      });
      binaryOutcome.dispatchEvent(new Event("change"));
    }
    document.querySelectorAll("[data-select-all], [data-clear-all]").forEach((button) => {
      button.addEventListener("click", () => {
        const name = button.dataset.selectAll || button.dataset.clearAll;
        const check = button.hasAttribute("data-select-all");
        document.querySelectorAll(`input[name="${name}"]`).forEach((input) => {
          if (!input.disabled) input.checked = check;
        });
      });
    });
  }

  function summarize(values) {
    const sorted = values.slice().sort((a, b) => a - b);
    const n = values.length;
    const mean = values.reduce((sum, value) => sum + value, 0) / n;
    const median = n % 2 ? sorted[(n - 1) / 2] : (sorted[n / 2 - 1] + sorted[n / 2]) / 2;
    const variance = n > 1 ? values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (n - 1) : NaN;
    const counts = new Map();
    values.forEach((value) => counts.set(value, (counts.get(value) || 0) + 1));
    const highestFrequency = Math.max(...counts.values());
    const modes = [...counts.entries()].filter(([, count]) => count === highestFrequency && count > 1).map(([value]) => value);
    const quantile = (probability) => {
      const position = (n - 1) * probability;
      const lower = Math.floor(position);
      const upper = Math.ceil(position);
      return sorted[lower] + (sorted[upper] - sorted[lower]) * (position - lower);
    };
    return {
      n, mean, median, mode: modes.length ? modes.map((value) => fmt(value)).join(", ") : "Yok",
      variance, sd: Math.sqrt(variance), min: sorted[0], max: sorted[n - 1],
      q1: quantile(0.25), q3: quantile(0.75),
    };
  }

  function descriptiveAnalysis(column) {
    if (numericTypes.has(state.types[column])) {
      const values = columnValues(column).map(parseNumber).filter((value) => value !== null);
      if (!values.length) throw new Error("Seçilen sütunda sayısal gözlem yok.");
      const stats = summarize(values);
      const rows = [
        ["N", stats.n], ["Ortalama", fmt(stats.mean)], ["Medyan", fmt(stats.median)], ["Mod", stats.mode],
        ["Standart sapma", fmt(stats.sd)], ["Varyans", fmt(stats.variance)],
        ["Birinci çeyrek (Q1)", fmt(stats.q1)], ["Üçüncü çeyrek (Q3)", fmt(stats.q3)],
        ["Minimum", fmt(stats.min)], ["Maksimum", fmt(stats.max)],
      ];
      return { title: `Tanımlayıcı istatistikler · ${column}`, headers: ["İstatistik", "Değer"], rows, note: "Standart sapma ve varyans örneklem ölçüleridir (n − 1). Çeyrekler doğrusal enterpolasyonla hesaplanır." };
    }
    const counts = new Map();
    columnValues(column).forEach((value) => {
      if (!isMissing(value)) counts.set(String(value), (counts.get(String(value)) || 0) + 1);
    });
    if (!counts.size) throw new Error("Seçilen sütunda gözlem yok.");
    const rows = [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 50)
      .map(([category, count]) => [category, count, fmt(count / [...counts.values()].reduce((sum, n) => sum + n, 0) * 100, 1)]);
    return { title: `Frekanslar · ${column}`, headers: ["Kategori", "Frekans", "Yüzde (%)"], rows, note: "En sık görülen ilk 50 kategori gösterilir." };
  }

  function runDescriptiveAnalysis(columns) {
    if (!columns.length) throw new Error("En az bir değişken seçin.");
    if (columns.length === 1) return descriptiveAnalysis(columns[0]);
    const analyses = columns.map(descriptiveAnalysis);
    return {
      title: `Tanımlayıcı istatistikler · ${columns.length} değişken`,
      headers: ["Değişken", "Tür", "Geçerli N", "Özet"],
      rows: analyses.map((analysis, index) => {
        const numeric = numericTypes.has(state.types[columns[index]]);
        const n = numeric
          ? columnValues(columns[index]).filter((value) => parseNumber(value) !== null).length
          : columnValues(columns[index]).filter((value) => !isMissing(value)).length;
        const summary = numeric
          ? analysis.rows.map(([label, value]) => `${label}: ${value}`).join(" · ")
          : `${analysis.rows.length} kategori`;
        return [columns[index], numeric ? "Sayısal" : "Kategorik", n, summary];
      }),
      extraTables: analyses.flatMap((analysis) => analysis.headers[0] === "Kategori"
        ? [{ title: analysis.title, headers: analysis.headers, rows: analysis.rows }]
        : []),
      note: "Sayısal özetler geçerli değerler üzerinden, kategorik frekanslar ilk 50 kategori üzerinden hesaplanır.",
    };
  }

  function pairedGroups(outcome, group, selectedLevels = []) {
    if (!outcome || !group) throw new Error("Sonuç ve grup değişkenlerini seçin.");
    if (outcome === group) throw new Error("Sonuç ve grup değişkeni aynı olamaz.");
    const values = new Map();
    activeRows().forEach((row) => {
      const value = parseNumber(row[outcome]);
      const category = row[group];
      if (value !== null && !isMissing(category) && (!selectedLevels.length || selectedLevels.includes(String(category)))) {
        const key = String(category);
        if (!values.has(key)) values.set(key, []);
        values.get(key).push(value);
      }
    });
    return [...values.entries()];
  }

  function runTTest(outcome, group, selectedLevels = []) {
    if (selectedLevels.length && selectedLevels.length !== 2) throw new Error("Bağımsız örneklem t-testi için tam olarak 2 grup kategorisi seçin.");
    const groups = pairedGroups(outcome, group, selectedLevels);
    if (groups.length !== 2) {
      if (!selectedLevels.length && groups.length > 2) throw new Error(`"${group}" değişkeninde ${groups.length} geçerli kategori var. t-testi için aşağıdan tam 2 kategori seçin.`);
      throw new Error(`t-testi için tam olarak 2 geçerli grup gerekir (bulunan: ${groups.length}).`);
    }
    const [aName, a] = groups[0];
    const [bName, b] = groups[1];
    if (a.length < 2 || b.length < 2) throw new Error("Her grupta en az 2 sayısal gözlem olmalıdır.");
    const sa = summarize(a); const sb = summarize(b);
    const va = sa.sd ** 2; const vb = sb.sd ** 2;
    const se2 = va / sa.n + vb / sb.n;
    if (!Number.isFinite(se2) || se2 === 0) throw new Error("Gruplardaki değişkenlik t-testi hesaplamak için yetersiz.");
    const t = (sa.mean - sb.mean) / Math.sqrt(se2);
    const df = se2 ** 2 / ((va / sa.n) ** 2 / (sa.n - 1) + (vb / sb.n) ** 2 / (sb.n - 1));
    const p = 2 * (1 - jStat.studentt.cdf(Math.abs(t), df));
    return {
      title: `Bağımsız örneklem t-testi · ${outcome} / ${group}`,
      headers: ["Grup", "N", "Ortalama", "Standart sapma"],
      rows: [[aName, sa.n, fmt(sa.mean), fmt(sa.sd)], [bName, sb.n, fmt(sb.mean), fmt(sb.sd)],
        ["Welch t", "", fmt(t), ""], ["Serbestlik derecesi", "", fmt(df), ""], ["p-değeri (iki yönlü)", "", fmtP(p), ""]],
      note: "Welch bağımsız örneklem t-testi kullanıldı; eşit varyans varsayılmaz.",
      p,
    };
  }

  function runAnova(outcome, group, doTukey = false) {
    const groups = pairedGroups(outcome, group).filter(([, values]) => values.length > 0);
    if (groups.length < 2) {
      throw new Error(`ANOVA için en az 2 geçerli grup gerekir. "${group}" değişkeninde geçerli sonuçla eşleşen grup sayısı ${groups.length}; grup sütununu ve etkin filtreyi kontrol edin.`);
    }
    const smallGroups = groups.filter(([, values]) => values.length < 2).map(([name]) => name);
    if (smallGroups.length) throw new Error(`Her ANOVA grubunda en az 2 sayısal gözlem olmalıdır. Yetersiz gözlemi olan gruplar: ${smallGroups.join(", ")}.`);
    const n = groups.reduce((sum, [, values]) => sum + values.length, 0);
    const grandMean = groups.reduce((sum, [, values]) => sum + values.reduce((s, value) => s + value, 0), 0) / n;
    const ssBetween = groups.reduce((sum, [, values]) => {
      const mean = values.reduce((s, value) => s + value, 0) / values.length;
      return sum + values.length * (mean - grandMean) ** 2;
    }, 0);
    const ssWithin = groups.reduce((sum, [, values]) => {
      const mean = values.reduce((s, value) => s + value, 0) / values.length;
      return sum + values.reduce((s, value) => s + (value - mean) ** 2, 0);
    }, 0);
    const dfBetween = groups.length - 1;
    const dfWithin = n - groups.length;
    const msBetween = ssBetween / dfBetween;
    const msWithin = ssWithin / dfWithin;
    if (!Number.isFinite(msWithin) || msWithin === 0) throw new Error("Gruplar içinde değişkenlik olmadığı için ANOVA hesaplanamıyor.");
    const f = msBetween / msWithin;
    const p = 1 - jStat.centralF.cdf(f, dfBetween, dfWithin);
    return {
      title: `Tek yönlü ANOVA · ${outcome} / ${group}`,
      headers: ["Kaynak", "Kareler toplamı", "sd", "Kareler ortalaması", "F", "p"],
      rows: [
        ["Gruplar arası", fmt(ssBetween), dfBetween, fmt(msBetween), fmt(f), fmtP(p)],
        ["Grup içi", fmt(ssWithin), dfWithin, fmt(msWithin), "—", "—"],
        ["Toplam", fmt(ssBetween + ssWithin), n - 1, "—", "—", "—"],
      ],
      note: doTukey
        ? "Tek yönlü klasik ANOVA ve Tukey–Kramer ikili karşılaştırmaları; varyans homojenliği varsayımı değerlendirilmelidir."
        : "Tek yönlü klasik ANOVA; parametrik varsayımları sonuçları yorumlarken ayrıca değerlendirin.",
      p,
      extraTables: doTukey ? [{
        title: "Tukey–Kramer ikili karşılaştırmaları",
        headers: ["Karşılaştırma", "Ortalama farkı", "q", "Düzeltilmiş p"],
        rows: tukeyComparisons(groups, msWithin, dfWithin),
      }] : [],
    };
  }

  function runCorrelation(columnX, columnY, method = "pearson") {
    if (columnX === columnY) throw new Error("Korelasyon için iki farklı değişken seçin.");
    let pairs = activeRows().map((row) => [parseNumber(row[columnX]), parseNumber(row[columnY])])
      .filter(([x, y]) => x !== null && y !== null);
    if (pairs.length < 3) throw new Error("Korelasyon için en az 3 eşleşmiş sayısal gözlem gerekir.");
    if (method === "spearman") {
      const ranksX = rankValues(pairs.map(([x]) => x)).ranks;
      const ranksY = rankValues(pairs.map(([, y]) => y)).ranks;
      pairs = pairs.map((_, index) => [ranksX[index], ranksY[index]]);
    }
    const xs = pairs.map(([x]) => x); const ys = pairs.map(([, y]) => y);
    const mx = xs.reduce((sum, value) => sum + value, 0) / xs.length;
    const my = ys.reduce((sum, value) => sum + value, 0) / ys.length;
    const cov = pairs.reduce((sum, [x, y]) => sum + (x - mx) * (y - my), 0);
    const sx = Math.sqrt(xs.reduce((sum, x) => sum + (x - mx) ** 2, 0));
    const sy = Math.sqrt(ys.reduce((sum, y) => sum + (y - my) ** 2, 0));
    if (sx === 0 || sy === 0) throw new Error("Sabit değerli değişkende korelasyon hesaplanamaz.");
    const r = Math.max(-1, Math.min(1, cov / (sx * sy)));
    const df = pairs.length - 2;
    const t = r * Math.sqrt(df / (1 - r * r));
    const p = Math.abs(r) === 1 ? 0 : 2 * (1 - jStat.studentt.cdf(Math.abs(t), df));
    return {
      title: `${method === "spearman" ? "Spearman" : "Pearson"} korelasyonu · ${columnX} / ${columnY}`,
      headers: ["Değişkenler", "N", method === "spearman" ? "Spearman rho" : "Pearson r", "t", "sd", "p-değeri"],
      rows: [[`${columnX} – ${columnY}`, pairs.length, fmt(r), fmt(t), df, fmtP(p)]],
      note: method === "spearman"
        ? "Spearman sıra korelasyonu; eşit sıralarda ortalama sıra ve iki yönlü asimptotik p-değeri kullanılır."
        : "Pearson doğrusal korelasyonu; p-değeri iki yönlüdür.",
      p,
    };
  }

  function runCorrelationMatrix(columns, method = "pearson") {
    if (columns.length < 2) throw new Error("Korelasyon için en az iki sayısal değişken seçin.");
    const pairs = [];
    for (let i = 0; i < columns.length; i++) {
      for (let j = i + 1; j < columns.length; j++) {
        const result = runCorrelation(columns[i], columns[j], method);
        pairs.push(result.rows[0]);
      }
    }
    return {
      title: `${method === "spearman" ? "Spearman" : "Pearson"} korelasyonları · ${columns.length} değişken`,
      headers: ["Değişken çifti", "N", method === "spearman" ? "Spearman rho" : "Pearson r", "t", "sd", "p-değeri"],
      rows: pairs,
      note: `Her değişken çifti için listwise geçerli gözlemler kullanılır; p-değerleri iki yönlüdür. Yöntem: ${method === "spearman" ? "Spearman" : "Pearson"}.`,
    };
  }

  function runPairedT(columnX, columnY) {
    if (!columnX || !columnY || columnX === columnY) throw new Error("İki farklı ölçüm değişkeni seçin.");
    const pairs = activeRows().map((row) => [parseNumber(row[columnX]), parseNumber(row[columnY])])
      .filter(([x, y]) => x !== null && y !== null);
    if (pairs.length < 2) throw new Error("Bağımlı örneklem t-testi için en az 2 eşleşmiş çift gerekir.");
    const differences = pairs.map(([x, y]) => x - y);
    const stats = summarize(differences);
    if (!stats.sd) throw new Error("Fark puanlarında değişkenlik yok.");
    const t = stats.mean / (stats.sd / Math.sqrt(stats.n));
    const df = stats.n - 1;
    const p = 2 * (1 - jStat.studentt.cdf(Math.abs(t), df));
    return {
      title: `Bağımlı örneklem t-testi · ${columnX} − ${columnY}`,
      headers: ["N", "Fark ortalaması", "Fark SS", "t", "sd", "p-değeri"],
      rows: [[stats.n, fmt(stats.mean), fmt(stats.sd), fmt(t), df, fmtP(p)]],
      note: "Eşleşmiş gözlemlerin fark puanlarına tek örneklem t-testi uygulanır; p-değeri iki yönlüdür.",
      p,
    };
  }

  function runOneSampleT(column, testValue) {
    if (!column || !Number.isFinite(testValue)) throw new Error("Sayısal değişken ve karşılaştırma değeri girin.");
    const values = columnValues(column).map(parseNumber).filter((value) => value !== null);
    if (values.length < 2) throw new Error("Tek örneklem t-testi için en az 2 gözlem gerekir.");
    const stats = summarize(values);
    if (!stats.sd) throw new Error("Değişkende varyans yok.");
    const t = (stats.mean - testValue) / (stats.sd / Math.sqrt(stats.n));
    const df = stats.n - 1;
    const p = 2 * (1 - jStat.studentt.cdf(Math.abs(t), df));
    const critical = jStat.studentt.inv(0.975, df);
    const margin = critical * stats.sd / Math.sqrt(stats.n);
    return {
      title: `Tek örneklem t-testi · ${column} (μ₀ = ${fmt(testValue)})`,
      headers: ["N", "Ortalama", "μ₀", "t", "sd", "p-değeri", "%95 GA alt", "%95 GA üst"],
      rows: [[stats.n, fmt(stats.mean), fmt(testValue), fmt(t), df, fmtP(p), fmt(stats.mean - margin), fmt(stats.mean + margin)]],
      note: "p-değeri iki yönlüdür; güven aralığı ortalama için %95 t aralığıdır.",
      p,
    };
  }

  function runWilcoxon(columnX, columnY) {
    if (!columnX || !columnY || columnX === columnY) throw new Error("İki farklı eşleştirilmiş ölçüm seçin.");
    const differences = activeRows().map((row) => [parseNumber(row[columnX]), parseNumber(row[columnY])])
      .filter(([x, y]) => x !== null && y !== null).map(([x, y]) => x - y).filter((difference) => difference !== 0);
    if (differences.length < 5) throw new Error("Wilcoxon yaklaşımı için sıfır olmayan en az 5 fark gerekir.");
    const ranked = rankValues(differences.map(Math.abs));
    const wPositive = differences.reduce((sum, difference, index) => sum + (difference > 0 ? ranked.ranks[index] : 0), 0);
    const n = differences.length;
    const mean = n * (n + 1) / 4;
    const variance = n * (n + 1) * (2 * n + 1) / 24 - ranked.tieSum / 48;
    if (variance <= 0) throw new Error("Sıra varyansı hesaplanamadı.");
    const z = (wPositive - mean - 0.5 * Math.sign(wPositive - mean)) / Math.sqrt(variance);
    const p = Math.min(1, 2 * jStat.normal.cdf(-Math.abs(z), 0, 1));
    return {
      title: `Wilcoxon işaretli sıralar testi · ${columnX} / ${columnY}`,
      headers: ["Eşleşmiş N", "Sıfır farkı çıkarılan", "W⁺", "z", "p-değeri"],
      rows: [[n, "Evet", fmt(wPositive), fmt(z), fmtP(p)]],
      note: "Sıfır farkları dışarıda bırakılır; bağlar için varyans düzeltmeli iki yönlü normal yaklaşımı kullanılır.",
      p,
    };
  }

  function runRepeatedAnova(columns) {
    if (columns.length < 2) throw new Error("Tekrarlı ölçümler ANOVA için en az 2 zaman/koşul seçin.");
    const observations = activeRows().map((row) => columns.map((column) => parseNumber(row[column])))
      .filter((values) => values.every((value) => value !== null));
    const n = observations.length; const k = columns.length;
    if (n < 2) throw new Error("En az 2 tam gözlemli katılımcı gerekir.");
    const grand = observations.flat().reduce((sum, value) => sum + value, 0) / (n * k);
    const conditionMeans = columns.map((_, j) => observations.reduce((sum, row) => sum + row[j], 0) / n);
    const subjectMeans = observations.map((row) => row.reduce((sum, value) => sum + value, 0) / k);
    const ssCondition = n * conditionMeans.reduce((sum, mean) => sum + (mean - grand) ** 2, 0);
    const ssSubject = k * subjectMeans.reduce((sum, mean) => sum + (mean - grand) ** 2, 0);
    const ssTotal = observations.flat().reduce((sum, value) => sum + (value - grand) ** 2, 0);
    const ssError = ssTotal - ssCondition - ssSubject;
    const dfCondition = k - 1; const dfError = (n - 1) * (k - 1);
    const msCondition = ssCondition / dfCondition; const msError = ssError / dfError;
    if (!(msError > 0)) throw new Error("Tekrarlı ölçümlerde hata varyansı sıfır.");
    const f = msCondition / msError;
    const p = 1 - jStat.centralF.cdf(f, dfCondition, dfError);
    return {
      title: "Tekrarlı ölçümler ANOVA",
      headers: ["Etki", "Kareler toplamı", "sd", "Kareler ortalaması", "F", "p"],
      rows: [
        ["Koşul / zaman", fmt(ssCondition), dfCondition, fmt(msCondition), fmt(f), fmtP(p)],
        ["Katılımcı", fmt(ssSubject), n - 1, "—", "—", "—"],
        ["Hata", fmt(ssError), dfError, fmt(msError), "—", "—"],
      ],
      extraTables: [{ title: "Ölçüm ortalamaları", headers: ["Ölçüm", "Ortalama"], rows: columns.map((column, j) => [column, fmt(conditionMeans[j])]) }],
      note: `${n} tam gözlemli katılımcı kullanıldı. Tek yönlü tekrarlı ölçümler modeli; sphericity düzeltmesi uygulanmamıştır.`,
      p,
    };
  }

  function symmetricEigen(matrix) {
    const n = matrix.length;
    const a = matrix.map((row) => row.slice());
    const vectors = Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => i === j ? 1 : 0));
    for (let iteration = 0; iteration < 50 * n * n; iteration++) {
      let p = 0; let q = 1;
      for (let i = 0; i < n; i++) for (let j = i + 1; j < n; j++) {
        if (Math.abs(a[i][j]) > Math.abs(a[p][q])) { p = i; q = j; }
      }
      if (Math.abs(a[p][q]) < 1e-10) break;
      const angle = 0.5 * Math.atan2(2 * a[p][q], a[q][q] - a[p][p]);
      const c = Math.cos(angle); const s = Math.sin(angle);
      const app = c * c * a[p][p] - 2 * s * c * a[p][q] + s * s * a[q][q];
      const aqq = s * s * a[p][p] + 2 * s * c * a[p][q] + c * c * a[q][q];
      for (let k = 0; k < n; k++) {
        if (k === p || k === q) continue;
        const akp = c * a[k][p] - s * a[k][q];
        const akq = s * a[k][p] + c * a[k][q];
        a[k][p] = a[p][k] = akp;
        a[k][q] = a[q][k] = akq;
      }
      a[p][p] = app; a[q][q] = aqq; a[p][q] = a[q][p] = 0;
      for (let k = 0; k < n; k++) {
        const vkp = c * vectors[k][p] - s * vectors[k][q];
        const vkq = s * vectors[k][p] + c * vectors[k][q];
        vectors[k][p] = vkp; vectors[k][q] = vkq;
      }
    }
    return Array.from({ length: n }, (_, i) => ({ value: a[i][i], vector: vectors.map((row) => row[i]) }))
      .sort((a, b) => b.value - a.value);
  }

  function standardizedMatrix(columns) {
    if (columns.length < 2) throw new Error("En az iki sayısal değişken seçin.");
    const rows = activeRows().map((row) => columns.map((column) => parseNumber(row[column])))
      .filter((values) => values.every((value) => value !== null));
    if (rows.length <= columns.length) throw new Error("Seçilen sütun sayısından fazla tam gözlem gerekir.");
    const means = columns.map((_, j) => rows.reduce((sum, row) => sum + row[j], 0) / rows.length);
    const sds = columns.map((_, j) => Math.sqrt(rows.reduce((sum, row) => sum + (row[j] - means[j]) ** 2, 0) / (rows.length - 1)));
    if (sds.some((sd) => !sd)) throw new Error("Sabit değerli değişkende analiz yapılamaz.");
    const matrix = columns.map((_, i) => columns.map((_, j) =>
      rows.reduce((sum, row) => sum + ((row[i] - means[i]) / sds[i]) * ((row[j] - means[j]) / sds[j]), 0) / (rows.length - 1)));
    return { rows, matrix };
  }

  function runPca(columns) {
    const { rows, matrix } = standardizedMatrix(columns);
    const eigen = symmetricEigen(matrix);
    const total = eigen.reduce((sum, item) => sum + Math.max(0, item.value), 0);
    return {
      title: `Temel bileşenler analizi · ${columns.length} değişken`,
      headers: ["Bileşen", "Özdeğer", "Açıklanan varyans (%)", "Kümülatif (%)", ...columns.map((column) => `Yük: ${column}`)],
      rows: eigen.map((item, index) => {
        const variance = Math.max(0, item.value) / total * 100;
        const cumulative = eigen.slice(0, index + 1).reduce((sum, value) => sum + Math.max(0, value.value) / total * 100, 0);
        return [`PC${index + 1}`, fmt(item.value), fmt(variance, 1), fmt(cumulative, 1), ...item.vector.map((value) => fmt(value * Math.sqrt(Math.max(0, item.value))))];
      }),
      note: `Korelasyon matrisi üzerinden PCA; ${rows.length} tam gözlem kullanıldı. Bileşen işaretleri matematiksel olarak keyfidir.`,
    };
  }

  function runReliability(columns) {
    if (columns.length < 2) throw new Error("Güvenirlik için en az iki madde seçin.");
    const { rows, matrix } = standardizedMatrix(columns);
    const k = columns.length;
    const meanInterItem = matrix.reduce((sum, row, i) =>
      sum + row.reduce((rowSum, value, j) => rowSum + (i === j ? 0 : value), 0) / (k - 1), 0) / k;
    const alpha = k * meanInterItem / (1 + (k - 1) * meanInterItem);
    const eigen = symmetricEigen(matrix);
    const loadings = eigen[0].vector.map((value) => value * Math.sqrt(Math.max(0, eigen[0].value)));
    const omega = (loadings.reduce((sum, value) => sum + value, 0)) ** 2 /
      ((loadings.reduce((sum, value) => sum + value, 0)) ** 2 + loadings.reduce((sum, value) => sum + Math.max(0, 1 - value ** 2), 0));
    return {
      title: `Ölçek güvenirliği · ${k} madde`,
      headers: ["Ölçüt", "Değer"],
      rows: [["Cronbach alfa (standartlaştırılmış)", fmt(alpha)], ["McDonald omega (tek bileşen yaklaşımı)", fmt(omega)], ["Madde sayısı", k], ["Tam gözlem", rows.length]],
      extraTables: [{ title: "Madde korelasyonları", headers: ["Değişken", ...columns], rows: columns.map((column, i) => [column, ...matrix[i].map(fmt)]) }],
      note: "Omega tek bileşenli PCA yüklerinden yaklaşık hesaplanmıştır; çok boyutlu ölçeklerde uzman faktör analiziyle doğrulayın.",
    };
  }

  function runLogistic(outcome, predictors) {
    if (!outcome || !predictors.length) throw new Error("İkili sonuç değişkeni ve en az bir yordayıcı seçin.");
    if (predictors.includes(outcome)) throw new Error("Sonuç değişkeni yordayıcı olarak seçilemez.");
    const observations = activeRows().filter((row) => !isMissing(row[outcome]) &&
      predictors.every((column) => !isMissing(row[column]) &&
        (!numericTypes.has(state.types[column]) || parseNumber(row[column]) !== null)));
    const levels = [...new Set(observations.map((row) => String(row[outcome])))];
    if (levels.length !== 2) throw new Error(`Lojistik regresyon için sonuç tam iki kategorili olmalı (bulunan: ${levels.length}).`);
    const design = predictors.flatMap((column) => {
      if (numericTypes.has(state.types[column])) return [{ column, level: null, label: column }];
      const categories = [...new Set(observations.map((row) => String(row[column])))];
      return categories.slice(1).map((level) => ({ column, level, label: `${column} = ${level}` }));
    });
    const columns = ["(Sabit)", ...design.map((term) => term.label)];
    if (observations.length <= columns.length) throw new Error("Lojistik model için daha fazla tam gözlem gerekir.");
    const y = observations.map((row) => String(row[outcome]) === levels[1] ? 1 : 0);
    const x = observations.map((row) => [1, ...design.map((term) => term.level === null
      ? parseNumber(row[term.column]) : Number(String(row[term.column]) === term.level))]);
    let beta = Array(columns.length).fill(0);
    let inverse;
    for (let iteration = 0; iteration < 80; iteration++) {
      const probabilities = x.map((row) => {
        const eta = row.reduce((sum, value, i) => sum + value * beta[i], 0);
        return 1 / (1 + Math.exp(-Math.max(-30, Math.min(30, eta))));
      });
      const weights = probabilities.map((p) => Math.max(1e-8, p * (1 - p)));
      const hessian = columns.map((_, i) => columns.map((_, j) =>
        x.reduce((sum, row, r) => sum + row[i] * row[j] * weights[r], 0)));
      inverse = invertMatrix(hessian);
      const gradient = columns.map((_, i) =>
        x.reduce((sum, row, r) => sum + row[i] * (y[r] - probabilities[r]), 0));
      const step = inverse.map((row) => row.reduce((sum, value, j) => sum + value * gradient[j], 0));
      beta = beta.map((value, i) => value + step[i]);
      if (Math.max(...step.map(Math.abs)) < 1e-8) break;
      if (iteration === 79) throw new Error("Lojistik regresyon yakınsamadı; modeli veya yordayıcıları kontrol edin.");
    }
    const rows = columns.map((column, i) => {
      const se = Math.sqrt(Math.max(0, inverse[i][i]));
      const z = beta[i] / se;
      const p = 2 * (1 - jStat.normal.cdf(Math.abs(z), 0, 1));
      return [column, fmt(beta[i]), fmt(se), fmt(z), fmtP(p), fmt(Math.exp(Math.max(-30, Math.min(30, beta[i]))))];
    });
    const probabilities = x.map((row) => {
      const linearPredictor = row.reduce((sum, value, i) => sum + value * beta[i], 0);
      const boundedPredictor = Math.max(-30, Math.min(30, linearPredictor));
      return 1 / (1 + Math.exp(-boundedPredictor));
    });
    const logLikelihood = probabilities.reduce((sum, p, i) => sum + (y[i] ? Math.log(Math.max(p, 1e-15)) : Math.log(Math.max(1 - p, 1e-15))), 0);
    return {
      title: `İkili lojistik regresyon · ${outcome}`,
      headers: ["Terim", "Katsayı", "SH", "Wald z", "p-değeri", "Odds oranı"],
      rows,
      extraTables: [{ title: "Model uyum ölçütleri", headers: ["N", "AIC", "Log-olabilirlik", "Sonuç kategorileri"],
        rows: [[observations.length, fmt(-2 * logLikelihood + 2 * columns.length), fmt(logLikelihood), `${levels[0]} (0) / ${levels[1]} (1)`]] }],
      note: "Wald z testleri ve odds oranları raporlanır. Küçük örneklem veya tam ayrışma durumlarında katsayılar kararsız olabilir.",
    };
  }

  function runChiSquare(columnX, columnY) {
    if (!columnX || !columnY) throw new Error("Çapraz tablo için iki değişken seçin.");
    if (columnX === columnY) throw new Error("Çapraz tablo için iki farklı kategorik değişken seçin.");
    const paired = activeRows().filter((row) => !isMissing(row[columnX]) && !isMissing(row[columnY]));
    if (!paired.length) throw new Error("Çapraz tabloda kullanılabilecek eşleşmiş gözlem yok.");
    const rowLevels = [...new Set(paired.map((row) => String(row[columnX])))];
    const columnLevels = [...new Set(paired.map((row) => String(row[columnY])))];
    if (rowLevels.length < 2 || columnLevels.length < 2) throw new Error("Ki-kare testi için her iki değişkende de en az iki kategori gerekir.");
    if (rowLevels.length * columnLevels.length > 1000) throw new Error("Çapraz tablo çok büyük (en fazla 1.000 hücre desteklenir).");
    const counts = rowLevels.map(() => columnLevels.map(() => 0));
    paired.forEach((row) => {
      const i = rowLevels.indexOf(String(row[columnX]));
      const j = columnLevels.indexOf(String(row[columnY]));
      counts[i][j]++;
    });
    const rowTotals = counts.map((row) => row.reduce((sum, value) => sum + value, 0));
    const colTotals = columnLevels.map((_, j) => counts.reduce((sum, row) => sum + row[j], 0));
    const n = rowTotals.reduce((sum, value) => sum + value, 0);
    if (!n) throw new Error("Seçilen değişkenlerde eşleşen gözlem yok.");
    const rows = rowLevels.map((level, i) => [
      level, ...counts[i], rowTotals[i],
    ]);
    rows.push(["Toplam", ...colTotals, n]);
    let statistic = 0; let smallExpected = 0;
    counts.forEach((row, i) => row.forEach((observed, j) => {
      const expected = rowTotals[i] * colTotals[j] / n;
      if (expected < 5) smallExpected++;
      statistic += (observed - expected) ** 2 / expected;
    }));
    const df = (rowLevels.length - 1) * (columnLevels.length - 1);
    const p = 1 - jStat.chisquare.cdf(statistic, df);
    return {
      title: `Ki-kare bağımsızlık testi · ${columnX} × ${columnY}`,
      headers: [columnX, ...columnLevels, "Toplam"],
      rows,
      extraTables: [{
        title: "Ki-kare testi",
        headers: ["İstatistik", "Değer"],
        rows: [["Pearson χ²", fmt(statistic)], ["Serbestlik derecesi", df], ["p-değeri", fmtP(p)], ["N", n]],
      }],
      note: smallExpected
        ? `${smallExpected} hücrenin beklenen frekansı 5'in altında; ki-kare yaklaşımı dikkatle yorumlanmalıdır.`
        : "Pearson ki-kare bağımsızlık testi; beklenen frekanslar hesaplanmıştır.",
      p,
    };
  }

  function rankValues(values) {
    const sorted = values.map((value, index) => ({ value, index })).sort((a, b) => a.value - b.value);
    const ranks = Array(values.length);
    let tieSum = 0;
    for (let start = 0; start < sorted.length;) {
      let end = start + 1;
      while (end < sorted.length && sorted[end].value === sorted[start].value) end++;
      const count = end - start;
      const rank = (start + 1 + end) / 2;
      for (let index = start; index < end; index++) ranks[sorted[index].index] = rank;
      if (count > 1) tieSum += count ** 3 - count;
      start = end;
    }
    return { ranks, tieSum };
  }

  function rankGroups(outcome, group, selectedLevels = []) {
    const groups = pairedGroups(outcome, group, selectedLevels).filter(([, values]) => values.length);
    if (groups.length < 2) throw new Error("Sıra temelli test için en az 2 geçerli grup gerekir.");
    const flat = [];
    groups.forEach(([name, values], groupIndex) => values.forEach((value) => flat.push({ value, groupIndex, name })));
    const { ranks, tieSum } = rankValues(flat.map((item) => item.value));
    return { groups, flat, ranks, tieSum };
  }

  function runMannWhitney(outcome, group, selectedLevels = []) {
    if (selectedLevels.length && selectedLevels.length !== 2) throw new Error("Mann–Whitney testi için tam olarak 2 grup kategorisi seçin.");
    const ranked = rankGroups(outcome, group, selectedLevels);
    if (ranked.groups.length !== 2) {
      if (!selectedLevels.length && ranked.groups.length > 2) throw new Error(`"${group}" değişkeninde ${ranked.groups.length} geçerli kategori var. Mann–Whitney için yukarıdan tam 2 kategori seçin.`);
      throw new Error(`Mann–Whitney testi için tam olarak 2 geçerli grup gerekir (bulunan: ${ranked.groups.length}).`);
    }
    const n1 = ranked.groups[0][1].length; const n2 = ranked.groups[1][1].length;
    if (n1 < 2 || n2 < 2) throw new Error("Her grupta en az 2 gözlem olmalıdır.");
    const rankSum = ranked.flat.reduce((sum, item, index) => sum + (item.groupIndex === 0 ? ranked.ranks[index] : 0), 0);
    const u1 = rankSum - n1 * (n1 + 1) / 2;
    const u2 = n1 * n2 - u1;
    const u = Math.min(u1, u2);
    const n = n1 + n2;
    const variance = n1 * n2 / 12 * ((n + 1) - ranked.tieSum / (n * (n - 1)));
    if (variance <= 0) throw new Error("Sıra değişkenliği hesaplanamadı; değerler sabit olabilir.");
    const z = (u - n1 * n2 / 2) / Math.sqrt(variance);
    const p = Math.min(1, 2 * jStat.normal.cdf(z, 0, 1));
    return {
      title: `Mann–Whitney U testi · ${outcome} / ${group}`,
      headers: ["Grup", "N", "Medyan", "Sıra toplamı"],
      rows: ranked.groups.map(([name, values], groupIndex) => {
        const sum = ranked.flat.reduce((total, item, index) => total + (item.groupIndex === groupIndex ? ranked.ranks[index] : 0), 0);
        return [name, values.length, fmt(summarize(values).median), fmt(sum)];
      }).concat([["U", "", fmt(u), ""], ["z", "", fmt(z), ""], ["p-değeri", "", fmtP(p), ""]]),
      note: "İki yönlü asimptotik normal yaklaşım; bağlar için varyans düzeltmesi uygulanmıştır. Küçük örneklemlerde kesin test tercih edilmelidir.",
      p,
    };
  }

  function runKruskal(outcome, group) {
    const ranked = rankGroups(outcome, group);
    if (ranked.groups.some(([, values]) => values.length < 2)) throw new Error("Her grupta en az 2 gözlem olmalıdır.");
    const n = ranked.flat.length;
    let h = 0;
    const rows = ranked.groups.map(([name, values], groupIndex) => {
      const rankSum = ranked.flat.reduce((sum, item, index) => sum + (item.groupIndex === groupIndex ? ranked.ranks[index] : 0), 0);
      h += rankSum ** 2 / values.length;
      return [name, values.length, fmt(summarize(values).median), fmt(rankSum)];
    });
    h = 12 * h / (n * (n + 1)) - 3 * (n + 1);
    const correction = 1 - ranked.tieSum / (n ** 3 - n);
    if (correction <= 0) throw new Error("Sıra değişkenliği hesaplanamadı; değerler sabit olabilir.");
    h /= correction;
    const df = ranked.groups.length - 1;
    const p = 1 - jStat.chisquare.cdf(h, df);
    rows.push(["Kruskal–Wallis H", "", fmt(h), ""]);
    rows.push(["Serbestlik derecesi", "", df, ""]);
    rows.push(["p-değeri", "", fmtP(p), ""]);
    return {
      title: `Kruskal–Wallis testi · ${outcome} / ${group}`,
      headers: ["Grup", "N", "Medyan", "Sıra toplamı"],
      rows,
      note: "Ki-kare asimptotik p-değeri; bağlar için düzeltme uygulanmıştır. Anlamlı sonuç sonrası ikili testler ayrıca değerlendirilmelidir.",
      p,
    };
  }

  function invertMatrix(matrix) {
    const size = matrix.length;
    const augmented = matrix.map((row, i) => [...row, ...Array.from({ length: size }, (_, j) => i === j ? 1 : 0)]);
    for (let column = 0; column < size; column++) {
      let pivot = column;
      for (let row = column + 1; row < size; row++) {
        if (Math.abs(augmented[row][column]) > Math.abs(augmented[pivot][column])) pivot = row;
      }
      if (Math.abs(augmented[pivot][column]) < 1e-12) throw new Error("Yordayıcılar doğrusal bağlantılı; model hesaplanamıyor.");
      [augmented[column], augmented[pivot]] = [augmented[pivot], augmented[column]];
      const divisor = augmented[column][column];
      augmented[column] = augmented[column].map((value) => value / divisor);
      for (let row = 0; row < size; row++) {
        if (row === column) continue;
        const factor = augmented[row][column];
        augmented[row] = augmented[row].map((value, index) => value - factor * augmented[column][index]);
      }
    }
    return augmented.map((row) => row.slice(size));
  }

  function runRegression(dependent, predictors, includeIntercept) {
    if (!dependent) throw new Error("Bağımlı değişken seçin.");
    if (!predictors.length) throw new Error("En az bir yordayıcı seçin.");
    if (predictors.includes(dependent)) throw new Error("Bağımlı değişken yordayıcı olarak seçilemez.");
    const sourceRows = activeRows();
    const completeRows = sourceRows.filter((row) => {
      if (parseNumber(row[dependent]) === null) return false;
      return predictors.every((column) => {
        const value = row[column];
        return isMissing(value) ? false : !numericTypes.has(state.types[column]) || parseNumber(value) !== null;
      });
    });
    if (!completeRows.length) throw new Error("Seçilen değişkenlerde eksiksiz gözlem bulunamadı.");
    const design = predictors.flatMap((column) => {
      if (numericTypes.has(state.types[column])) return [{ column, level: null, label: column }];
      const levels = [...new Set(completeRows.map((row) => String(row[column])))];
      if (levels.length < 2) throw new Error(`"${column}" yordayıcısında en az iki geçerli kategori bulunmalıdır.`);
      return levels.slice(1).map((level) => ({ column, level, label: `${column} = ${level}` }));
    });
    const columns = [...(includeIntercept ? ["(Sabit)"] : []), ...design.map((term) => term.label)];
    const observations = completeRows.map((row) => {
      const y = parseNumber(row[dependent]);
      const values = design.map((term) => {
        const value = row[term.column];
        if (isMissing(value)) return null;
        if (term.level !== null) return String(value) === term.level ? 1 : 0;
        return parseNumber(value);
      });
      return y !== null && values.every((value) => value !== null)
        ? { y, x: includeIntercept ? [1, ...values] : values } : null;
    }).filter(Boolean);
    const n = observations.length; const pCount = columns.length;
    if (n <= pCount) throw new Error(`Model için gözlem sayısı yordayıcı sayısından fazla olmalı (N=${n}, parametre=${pCount}).`);
    const xtx = columns.map((_, i) => columns.map((_, j) =>
      observations.reduce((sum, observation) => sum + observation.x[i] * observation.x[j], 0)));
    const xty = columns.map((_, i) =>
      observations.reduce((sum, observation) => sum + observation.x[i] * observation.y, 0));
    const inverse = invertMatrix(xtx);
    const coefficients = inverse.map((row) => row.reduce((sum, value, j) => sum + value * xty[j], 0));
    const meanY = observations.reduce((sum, observation) => sum + observation.y, 0) / n;
    const sse = observations.reduce((sum, observation) => {
      const predicted = observation.x.reduce((total, value, i) => total + value * coefficients[i], 0);
      return sum + (observation.y - predicted) ** 2;
    }, 0);
    const sst = observations.reduce((sum, observation) =>
      sum + (observation.y - (includeIntercept ? meanY : 0)) ** 2, 0);
    const r2 = sst === 0 ? NaN : 1 - sse / sst;
    const residualDf = n - pCount;
    const mse = sse / residualDf;
    const rows = columns.map((column, i) => {
      const se = Math.sqrt(Math.max(0, mse * inverse[i][i]));
      const t = se ? coefficients[i] / se : NaN;
      const pValue = Number.isFinite(t) ? 2 * (1 - jStat.studentt.cdf(Math.abs(t), residualDf)) : NaN;
      return [column, fmt(coefficients[i]), fmt(se), fmt(t), fmtP(pValue)];
    });
    const modelDf = pCount - (includeIntercept ? 1 : 0);
    const modelSs = sst - sse;
    const f = modelDf > 0 && mse > 0 ? (modelSs / modelDf) / mse : NaN;
    const modelP = Number.isFinite(f) ? 1 - jStat.centralF.cdf(f, modelDf, residualDf) : NaN;
    return {
      title: `Doğrusal regresyon · ${dependent}`,
      headers: ["Terim", "Katsayı", "Standart hata", "t", "p-değeri"],
      rows,
      extraTables: [{
        title: "Model özeti",
        headers: ["N", "R²", "F", "sd model", "sd hata", "p-değeri"],
        rows: [[n, fmt(r2), fmt(f), modelDf, residualDf, fmtP(modelP)]],
      }],
      note: `${includeIntercept ? "Sabit terimli" : "Sabit terimsiz"} en küçük kareler modeli. Nominal yordayıcılar ilk gözlenen kategori referans alınarak kukla değişkenlere dönüştürüldü. Eksik gözlemler listwise olarak çıkarıldı.`,
      p: modelP,
    };
  }

  const tukeyCache = new Map();
  function tukeyCdfInfinite(q, groups) {
    const intervals = 120;
    const lower = -8; const upper = 8; const step = (upper - lower) / intervals;
    let sum = 0;
    for (let i = 0; i <= intervals; i++) {
      const x = lower + i * step;
      const difference = jStat.normal.cdf(x + q, 0, 1) - jStat.normal.cdf(x, 0, 1);
      const value = jStat.normal.pdf(x, 0, 1) * Math.max(0, difference) ** (groups - 1);
      sum += (i === 0 || i === intervals ? 1 : i % 2 ? 4 : 2) * value;
    }
    return Math.max(0, Math.min(1, groups * sum * step / 3));
  }

  function tukeyPValue(q, groups, df) {
    const key = `${q.toFixed(5)}:${groups}:${df}`;
    if (tukeyCache.has(key)) return tukeyCache.get(key);
    const intervals = 120;
    const upper = df + 10 * Math.sqrt(2 * df) + 60;
    const step = upper / intervals;
    let integral = 0;
    for (let i = 0; i <= intervals; i++) {
      const chi = i * step;
      const density = chi === 0 ? 0 : jStat.gamma.pdf(chi / 2, df / 2, 1) / 2;
      const cdf = chi === 0 ? 0 : tukeyCdfInfinite(q * Math.sqrt(chi / df), groups);
      const value = density * cdf;
      integral += (i === 0 || i === intervals ? 1 : i % 2 ? 4 : 2) * value;
    }
    const p = Math.max(0, Math.min(1, 1 - integral * step / 3));
    tukeyCache.set(key, p);
    return p;
  }

  function tukeyComparisons(groups, mse, df) {
    const comparisons = [];
    for (let i = 0; i < groups.length; i++) {
      for (let j = i + 1; j < groups.length; j++) {
        const [nameA, valuesA] = groups[i]; const [nameB, valuesB] = groups[j];
        const meanA = summarize(valuesA).mean; const meanB = summarize(valuesB).mean;
        const standardError = Math.sqrt(mse / 2 * (1 / valuesA.length + 1 / valuesB.length));
        const q = Math.abs(meanA - meanB) / standardError;
        comparisons.push([`${nameA} – ${nameB}`, fmt(meanA - meanB), fmt(q), fmtP(tukeyPValue(q, groups.length, df))]);
      }
    }
    return comparisons;
  }

  function resultTable(result) {
    const renderTable = (headers, rows, title = "") => `${title ? `<h4 class="sub-result-title">${escapeHtml(title)}</h4>` : ""}<div class="table-wrap results-table"><table><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${rows.map((row) =>
      `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`;
    return `<div class="result-banner"><strong>${escapeHtml(result.title)}</strong></div>${renderTable(result.headers, result.rows)}${(result.extraTables || []).map((table) => renderTable(table.headers, table.rows, table.title)).join("")}<p class="result-note">${escapeHtml(result.note)}</p>`;
  }

  function renderResultsEmpty() {
    $("#resultsContent").className = "results-empty";
    $("#resultsContent").innerHTML = '<span class="empty-chart">∑</span><p>Sonuçları görmek için bir analiz çalıştırın.</p>';
  }

  function runAnalysis() {
    try {
      if (!window.jStat) throw new Error("İstatistik kütüphanesi yüklenemedi. İnternet bağlantınızı kontrol edin.");
      if (!activeRows().length) throw new Error("Etkin filtreye uyan gözlem yok.");
      const type = $("#analysisType").value;
      let result;
      if (type === "descriptive") result = runDescriptiveAnalysis(checkedValues("variable"));
      else if (type === "ttest") result = runTTest($("#outcome").value, $("#group").value, checkedValues("groupLevels"));
      else if (type === "anova") result = runAnova($("#outcome").value, $("#group").value, $("#tukeyOption").checked);
      else if (type === "correlation" || type === "spearman") result = runCorrelationMatrix(checkedValues("correlationVariables"), type);
      else if (type === "pairedT") result = runPairedT($("#variableX").value, $("#variableY").value);
      else if (type === "wilcoxon") result = runWilcoxon($("#variableX").value, $("#variableY").value);
      else if (type === "oneSampleT") result = runOneSampleT($("#outcome").value, Number($("#testValue").value));
      else if (type === "repeatedAnova") result = runRepeatedAnova(checkedValues("repeatedVariables"));
      else if (type === "logistic") result = runLogistic($("#binaryOutcome").value, checkedValues("logitPredictors"));
      else if (type === "reliability") result = runReliability(checkedValues("reliabilityVariables"));
      else if (type === "pca") result = runPca(checkedValues("pcaVariables"));
      else if (type === "regression") {
        const predictors = checkedValues("predictors");
        result = runRegression($("#dependent").value, predictors, $("#regressionDependentCheck").checked);
      } else if (type === "chisquare") result = runChiSquare($("#categoryX").value, $("#categoryY").value);
      else if (type === "mannwhitney") result = runMannWhitney($("#outcome").value, $("#group").value, checkedValues("groupLevels"));
      else result = runKruskal($("#outcome").value, $("#group").value);
      $("#resultsContent").className = "";
      $("#resultsContent").innerHTML = resultTable(result);
      state.results.unshift({ ...result, date: new Date().toLocaleString("tr-TR") });
      renderHistory();
      $("#downloadResults").disabled = false;
      $("#downloadWorkbook").disabled = false;
      $("#downloadReport").disabled = false;
    } catch (error) {
      $("#resultsContent").className = "";
      $("#resultsContent").innerHTML = `<p class="error-message" role="alert">${escapeHtml(error.message || "Analiz çalıştırılamadı.")}</p>`;
    }
  }

  function renderHistory() {
    const table = $("#historyTable");
    $("#historyCount").textContent = `${state.results.length} analiz`;
    if (!state.results.length) {
      table.innerHTML = '<tbody><tr><td class="history-empty">Henüz analiz çalıştırılmadı.</td></tr></tbody>';
      return;
    }
    table.innerHTML = `<thead><tr><th>Analiz</th><th>Tarih</th><th>p-değeri</th></tr></thead><tbody>${state.results.map((result) =>
      `<tr><td>${escapeHtml(result.title)}</td><td>${escapeHtml(result.date)}</td><td>${Number.isFinite(result.p) ? fmtP(result.p) : "—"}</td></tr>`).join("")}</tbody>`;
  }

  function downloadResults() {
    if (!state.results.length) return;
    const data = [
      ["VERILAB", "ANALİZ SONUÇ RAPORU"],
      ["Veri seti", state.fileName],
      ["Rapor tarihi", new Date().toLocaleString("tr-TR")],
      ["Kullanılan gözlem", `${activeRows().length} / ${state.rows.length}`],
      [],
    ];
    state.results.forEach((result) => {
      data.push(["ANALİZ", result.title], ["Tarih", result.date], []);
      [{ title: "", headers: result.headers, rows: result.rows }, ...(result.extraTables || [])].forEach((table) => {
        if (table.title) data.push([table.title]);
        data.push(table.headers, ...table.rows, []);
      });
      data.push(["Yöntem notu", result.note], [], []);
    });
    const csvSafe = data.map((row) => row.map((cell) => {
      const text = String(cell ?? "");
      return /^[\t\r ]*[=+\-@]/.test(text) ? `'${text}` : text;
    }));
    const csv = Papa.unparse(csvSafe, { delimiter: ";", quotes: true });
    const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "biz-bize-analiz-sonuclari.csv";
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function downloadWorkbook() {
    if (!state.results.length) return;
    if (!window.XLSX) throw new Error("Excel dışa aktarma kütüphanesi yüklenemedi. İnternet bağlantınızı kontrol edin.");
    const workbook = XLSX.utils.book_new();
    const summary = [
      ["VERILAB", "ANALİZ RAPORU"],
      ["Veri seti", state.fileName],
      ["Rapor tarihi", new Date().toLocaleString("tr-TR")],
      ["Kullanılan gözlem", `${activeRows().length} / ${state.rows.length}`],
      [],
      ["Analiz", "Tarih", "p-değeri", "Yöntem notu"],
      ...state.results.map((result) => [result.title, result.date, Number.isFinite(result.p) ? result.p : "", result.note]),
    ];
    const summarySheet = XLSX.utils.aoa_to_sheet(summary);
    summarySheet["!cols"] = [{ wch: 38 }, { wch: 24 }, { wch: 14 }, { wch: 78 }];
    summarySheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
    XLSX.utils.book_append_sheet(workbook, summarySheet, "Özet");
    const usedNames = new Set(["Özet"]);
    state.results.forEach((result, index) => {
      const sheetData = [
        ["VERILAB | ANALİZ SONUÇLARI"],
        ["Analiz", result.title],
        ["Tarih", result.date],
        ["Veri seti", state.fileName],
        ["Kullanılan gözlem", `${activeRows().length} / ${state.rows.length}`],
        [],
      ];
      const tables = [{ headers: result.headers, rows: result.rows }, ...(result.extraTables || [])];
      tables.forEach((table, tableIndex) => {
        if (tableIndex > 0) sheetData.push([result.extraTables[tableIndex - 1].title]);
        sheetData.push(table.headers, ...table.rows, []);
      });
      sheetData.push(["Yöntem notu", result.note]);
      const sheet = XLSX.utils.aoa_to_sheet(sheetData);
      const widthCount = Math.max(...sheetData.map((row) => row.length));
      sheet["!cols"] = Array.from({ length: widthCount }, (_, column) => {
        const contentLength = Math.max(12, ...sheetData.map((row) => String(row[column] ?? "").length));
        return { wch: Math.min(55, contentLength + 2) };
      });
      const base = (result.title.split("·")[0].trim() || `Analiz ${index + 1}`)
        .replace(/[\\/?*[\]:]/g, " ").slice(0, 25).trim() || `Analiz ${index + 1}`;
      let sheetName = base; let suffix = 2;
      while (usedNames.has(sheetName)) {
        const ending = ` ${suffix++}`;
        sheetName = `${base.slice(0, 31 - ending.length)}${ending}`;
      }
      usedNames.add(sheetName);
      XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
    });
    XLSX.writeFile(workbook, "biz-bize-analiz-raporu.xlsx", { compression: true });
  }

  function downloadReport() {
    if (!state.results.length) return;
    const report = $("#printReport");
    const makeTable = (headers, rows) => `<table><thead><tr>${headers.map((cell) => `<th>${escapeHtml(cell)}</th>`).join("")}</tr></thead><tbody>${rows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
    report.className = "print-report";
    const chartImage = state.chart ? `<figure class="report-chart"><img src="${state.chart.toBase64Image("image/png", 1)}" alt="Oluşturulan analiz grafiği"><figcaption>Oluşturulan grafik</figcaption></figure>` : "";
    report.innerHTML = `<header class="report-header"><div class="report-brand"><span class="report-mark">V</span><span>VERILAB <small>İSTATİSTİKSEL ANALİZ RAPORU</small></span></div><span class="report-date">${escapeHtml(new Date().toLocaleDateString("tr-TR"))}</span></header><h1>Analiz sonuçları</h1><p class="report-lead">Veri setinizin istatistiksel analiz özeti.</p><div class="report-metadata"><div><span>VERİ SETİ</span><strong>${escapeHtml(state.fileName)}</strong></div><div><span>ETKİN GÖZLEM</span><strong>${activeRows().length.toLocaleString("tr-TR")} / ${state.rows.length.toLocaleString("tr-TR")}</strong></div><div><span>DEĞİŞKEN</span><strong>${state.columns.length}</strong></div><div><span>RAPOR TARİHİ</span><strong>${escapeHtml(new Date().toLocaleString("tr-TR"))}</strong></div></div>${chartImage}<main>${state.results.map((result, index) =>
      `<article class="report-analysis"><div class="report-section-heading"><span class="report-index">${String(index + 1).padStart(2, "0")}</span><h2>${escapeHtml(result.title)}</h2></div><p class="report-note">${escapeHtml(result.note)}</p>${makeTable(result.headers, result.rows)}${(result.extraTables || []).map((table) => `<h3>${escapeHtml(table.title)}</h3>${makeTable(table.headers, table.rows)}`).join("")}</article>`).join("")}</main><footer class="report-footer"><span>Biz Bize · Tarayıcınızda işlenen analiz raporu</span><span>Bu rapor otomatik olarak oluşturulmuştur.</span></footer>`;
    report.hidden = false;
    window.print();
  }

  const rPackages = {
    efa: ["psych"], omega: ["psych"], cfa: ["lavaan"], sem: ["lavaan"],
    mediation: ["mediation"], meta: ["metafor"], bayes: ["BayesFactor"],
    survival: ["survival"], mixed: ["lme4"],
  };

  function rString(value) {
    return JSON.stringify(String(value)).replace(/\u2028|\u2029/g, (char) => char === "\u2028" ? "\\u2028" : "\\u2029");
  }

  function rColumn(column) {
    return `data[[${rString(column)}]]`;
  }

  function renderRAnalysisFields() {
    const type = $("#advancedAnalysis").value;
    const numeric = numericColumns();
    let markup = "";
    if (["efa", "omega"].includes(type)) {
      markup = checklistField("rItems", "ÖLÇEK / FAKTÖR MADDELERİ", numeric, numeric.slice(0, Math.min(5, numeric.length)), "Birden fazla sayısal sütun seçin.");
    } else if (type === "cfa" || type === "sem") {
      markup = checklistField("rItems", "MODEL GÖSTERGELERİ", numeric, numeric.slice(0, Math.min(4, numeric.length)), "Seçili değişkenlerle basit bir tek faktör modeli üretir; R kodunu düzenleyebilirsiniz.");
    } else if (type === "mediation") {
      markup = selectField("rX", "YORDAYICI (X)", numeric, numeric[0] || "", true);
      markup += selectField("rMediator", "ARACI (M)", numeric, numeric[1] || "", true);
      markup += selectField("rY", "SONUÇ (Y)", numeric, numeric[2] || "", true);
    } else if (type === "moderation") {
      markup = selectField("rY", "SONUÇ (Y)", numeric, numeric[0] || "", true);
      markup += selectField("rX", "YORDAYICI (X)", numeric, numeric[1] || "", true);
      markup += selectField("rModerator", "DÜZENLEYİCİ (W)", numeric, numeric[2] || "", true);
    } else if (type === "meta") {
      markup = selectField("rEffect", "ETKİ BÜYÜKLÜĞÜ (yi)", numeric, numeric[0] || "", true);
      markup += selectField("rVariance", "ETKİ VARYANSI (vi)", numeric, numeric[1] || "", true);
    } else if (type === "bayes") {
      markup = selectField("rOutcome", "SAYISAL DEĞİŞKEN", numeric, numeric[0] || "", true);
      markup += '<div class="field-wrap"><label for="rMu">KARŞILAŞTIRMA ORTALAMASI</label><input id="rMu" class="text-control" type="number" step="any" value="0"></div>';
    } else if (type === "survival") {
      markup = selectField("rTime", "SÜRE DEĞİŞKENİ", numeric, numeric[0] || "", true);
      markup += selectField("rStatus", "OLAY / SANSÜR DEĞİŞKENİ (0/1)", state.columns, "", true);
      markup += selectField("rGroup", "İSTEĞE BAĞLI GRUP", state.columns, "", true);
    } else if (type === "mixed") {
      markup = selectField("rY", "SAYISAL SONUÇ", numeric, numeric[0] || "", true);
      markup += selectField("rX", "SAYISAL YORDAYICI", numeric, numeric[1] || "", true);
      markup += selectField("rGroup", "KÜME / DENEK KİMLİĞİ", state.columns, "", true);
    }
    $("#rAnalysisFields").innerHTML = markup;
    document.querySelectorAll("#rAnalysisFields [data-select-all], #rAnalysisFields [data-clear-all]").forEach((button) => {
      button.addEventListener("click", () => {
        const name = button.dataset.selectAll || button.dataset.clearAll;
        const check = button.hasAttribute("data-select-all");
        document.querySelectorAll(`#rAnalysisFields input[name="${name}"]`).forEach((input) => { input.checked = check; });
      });
    });
    document.querySelectorAll("#rAnalysisFields select").forEach((select) => select.addEventListener("change", updateRTemplate));
    document.querySelectorAll("#rAnalysisFields input[type=checkbox]").forEach((input) => input.addEventListener("change", updateRTemplate));
    updateRTemplate();
  }

  function updateRTemplate() {
    if ($("#advancedAnalysis").value === "custom") return;
    const type = $("#advancedAnalysis").value;
    const selected = (name) => checkedValues(name).map(rString).join(", ");
    const x = $("#rX")?.value;
    const y = $("#rY")?.value;
    const itemNames = checkedValues("rItems");
    const safeItemNames = itemNames.map((_, index) => `v${index + 1}`);
    const factorModel = `factor1 =~ ${safeItemNames.join(" + ")}`;
    const templates = {
      efa: `# Keiser-Meyer-Olkin / Bartlett ve EFA\nitems <- data[c(${selected("rItems")})]\nprint(psych::KMO(items))\nprint(psych::cortest.bartlett(cor(items, use = "pairwise.complete.obs"), n = nrow(items)))\nprint(psych::fa(items, nfactors = 2, rotate = "oblimin", fm = "minres"))`,
      omega: `# McDonald omega ve madde istatistikleri\nitems <- data[c(${selected("rItems")})]\nprint(psych::omega(items, nfactors = 1, plot = FALSE))`,
      cfa: `# Confirmatory factor analysis — lavaan paketi\nmodel_data <- data[c(${selected("rItems")})]\nnames(model_data) <- c(${safeItemNames.map(rString).join(", ")})\nvariable_map <- data.frame(model_variable = names(model_data), original_variable = c(${selected("rItems")}))\nmodel <- ${rString(factorModel)}\nfit <- lavaan::cfa(model, data = model_data, std.lv = TRUE)\nprint(variable_map)\nprint(lavaan::fitMeasures(fit, c("chisq", "df", "pvalue", "cfi", "tli", "rmsea", "srmr", "aic", "bic")))\nprint(lavaan::standardizedSolution(fit))`,
      sem: `# Structural equation model — lavaan; yapısal yolları model satırlarına ekleyin\nmodel_data <- data[c(${selected("rItems")})]\nnames(model_data) <- c(${safeItemNames.map(rString).join(", ")})\nvariable_map <- data.frame(model_variable = names(model_data), original_variable = c(${selected("rItems")}))\nmodel <- ${rString(factorModel)}\nfit <- lavaan::sem(model, data = model_data)\nprint(variable_map)\nprint(lavaan::fitMeasures(fit, c("chisq", "df", "pvalue", "cfi", "tli", "rmsea", "srmr", "aic", "bic")))\nprint(lavaan::standardizedSolution(fit))`,
      mediation: `# Basit aracılık; bootstrap güven aralığı\nanalysis_data <- data.frame(x = ${rColumn(x || "")}, m = ${rColumn($("#rMediator")?.value || "")}, y = ${rColumn(y || "")})\nmed.fit <- lm(m ~ x, data = analysis_data)\nout.fit <- lm(y ~ x + m, data = analysis_data)\nmediation::mediate(med.fit, out.fit, treat = "x", mediator = "m", boot = TRUE, sims = 1000)`,
      moderation: `# Etkileşimli doğrusal model\nanalysis_data <- data.frame(y = ${rColumn(y || "")}, x = ${rColumn(x || "")}, moderator = ${rColumn($("#rModerator")?.value || "")})\nfit <- lm(y ~ x * moderator, data = analysis_data)\nprint(summary(fit))\nprint(confint(fit))`,
      meta: `# Random-effects meta-analizi (REML)\nprint(metafor::rma(yi = ${rColumn($("#rEffect")?.value || "")}, vi = ${rColumn($("#rVariance")?.value || "")}, data = data, method = "REML"))`,
      bayes: `# Bayes faktörlü tek örneklem t-testi\nprint(BayesFactor::ttestBF(x = ${rColumn($("#rOutcome")?.value || "")}, mu = ${$("#rMu")?.value || 0}))`,
      survival: `# Kaplan–Meier sağkalım eğrisi ve log-rank karşılaştırması\nsurvival_data <- data.frame(time = ${rColumn($("#rTime")?.value || "")}, status = ${rColumn($("#rStatus")?.value || "")}${$("#rGroup")?.value ? `, group = ${rColumn($("#rGroup").value)}` : ""})\nsurv <- survival::Surv(time, status)\nfit <- survival::survfit(surv ~ ${$("#rGroup")?.value ? "group" : "1"}, data = survival_data)\nprint(summary(fit))\n${$("#rGroup")?.value ? "survival::survdiff(surv ~ group, data = survival_data)" : "plot(fit, xlab = 'Zaman', ylab = 'Sağkalım olasılığı')"}`
      ,
      mixed: `# Random intercept mixed-effects model\nanalysis_data <- data.frame(y = ${rColumn(y || "")}, x = ${rColumn(x || "")}, group = ${rColumn($("#rGroup")?.value || "")})\nfit <- lme4::lmer(y ~ x + (1 | group), data = analysis_data, REML = TRUE)\nprint(summary(fit))\nprint(lme4::confint.merMod(fit, method = "Wald"))`,
    };
    $("#rCode").value = templates[type] || "";
  }

  async function syncRDataset() {
    if (!state.webR) return;
    const rows = activeRows();
    if (!rows.length) throw new Error("Etkin filtreye uyan gözlem yok.");
    const csv = Papa.unparse(rows, { quotes: true });
    const bytes = new TextEncoder().encode(csv);
    await state.webR.FS.writeFile("/tmp/verilab-data.csv", bytes);
    await state.webR.evalRVoid('data <- read.csv("/tmp/verilab-data.csv", check.names = FALSE, na.strings = c("", "NA"), stringsAsFactors = FALSE)');
    state.rDataVersion++;
  }

  async function startR() {
    const button = $("#startR");
    button.disabled = true;
    $("#rStatus").textContent = "R motoru indiriliyor…";
    $("#rProgress").textContent = "İlk başlatma biraz sürebilir.";
    try {
      const { WebR } = await import("https://webr.r-wasm.org/latest/webr.mjs");
      const webR = new WebR();
      await webR.init();
      state.webR = webR;
      await syncRDataset();
      $("#rStatus").textContent = "R hazır · veri tarayıcıda";
      $("#runR").disabled = false;
      $("#installRPackages").disabled = false;
      $("#rProgress").textContent = "R ve aktif veri seti hazır.";
      $("#rProgress").classList.remove("error-message");
    } catch (error) {
      $("#rStatus").textContent = "R başlatılamadı";
      $("#rProgress").textContent = error.message || "WebAssembly R yüklenemedi.";
      $("#rProgress").classList.add("error-message");
      button.disabled = false;
    }
  }

  async function installRPackages() {
    if (!state.webR) throw new Error("Önce R motorunu başlatın.");
    const packages = requiredRPackages($("#advancedAnalysis").value, $("#rCode").value);
    if (!packages.length) throw new Error("Bu analiz için ek R paketi gerekmiyor.");
    $("#rProgress").classList.remove("error-message");
    $("#rProgress").textContent = `İndiriliyor: ${packages.join(", ")}…`;
    await ensureRPackages(packages);
    $("#rProgress").textContent = `Paketler hazır: ${packages.join(", ")}`;
  }

  function requiredRPackages(type, code) {
    const packages = new Set(rPackages[type] || []);
    const knownPackages = new Set(Object.values(rPackages).flat());
    knownPackages.forEach((name) => {
      const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const usage = new RegExp(`(?:\\b${escapedName}::|\\b(?:library|require)\\s*\\(\\s*['"]?${escapedName}\\b)`);
      if (usage.test(code)) packages.add(name);
    });
    return [...packages];
  }

  async function ensureRPackages(packages) {
    if (!packages.length) return;
    await state.webR.installPackages(packages);
    const packageNames = packages.map(rString).join(", ");
    await state.webR.evalRVoid(`.verilab_missing_packages <- c(${packageNames})[!vapply(c(${packageNames}), requireNamespace, logical(1), quietly = TRUE)]; if (length(.verilab_missing_packages)) stop(paste("R paketleri yüklenemedi:", paste(.verilab_missing_packages, collapse = ", ")))`);
  }

  async function runRCode() {
    if (!state.webR) throw new Error("Önce R motorunu başlatın.");
    const code = $("#rCode").value.trim();
    if (!code) throw new Error("Çalıştırmak için R kodu girin.");
    $("#runR").disabled = true;
    let packageSetup = false;
    $("#rProgress").textContent = "R paketleri kontrol ediliyor…";
    $("#rProgress").classList.remove("error-message");
    try {
      const packages = requiredRPackages($("#advancedAnalysis").value, code);
      if (packages.length) {
        packageSetup = true;
        $("#rProgress").textContent = `R paketleri kontrol ediliyor / yükleniyor: ${packages.join(", ")}…`;
        await ensureRPackages(packages);
        packageSetup = false;
      }
      $("#rProgress").textContent = "Veri tarayıcıdaki R ortamına aktarılıyor…";
      await syncRDataset();
      $("#rProgress").textContent = "R kodu çalışıyor…";
      await state.webR.evalRVoid("verilab_output <- capture.output({ .verilab_expressions <- parse(text = " +
        rString(code) + `); for (.verilab_expression in .verilab_expressions) { .verilab_assignment <- is.call(.verilab_expression) && as.character(.verilab_expression[[1]]) %in% c("<-", "=", "<<-"); .verilab_evaluated <- withVisible(eval(.verilab_expression, envir = .GlobalEnv)); if (!.verilab_assignment && .verilab_evaluated$visible && !is.null(.verilab_evaluated$value)) print(.verilab_evaluated$value) } })`);
      const result = await state.webR.evalR("paste(verilab_output, collapse = '\\n')");
      const output = await result.toJs();
      const outputText = typeof output === "string" ? output
        : Array.isArray(output) ? output.join("\n")
          : output && Array.isArray(output.values) ? output.values.join("\n")
            : JSON.stringify(output, null, 2);
      $("#rOutput").textContent = outputText || "(R kodu çıktı üretmedi.)";
      $("#rProgress").textContent = "R kodu tamamlandı.";
    } catch (error) {
      $("#rOutput").textContent = error.message || String(error);
      $("#rProgress").textContent = packageSetup ? "R paketleri yüklenemedi." : "R kodu çalıştırılırken hata oluştu.";
      $("#rProgress").classList.add("error-message");
    } finally {
      $("#runR").disabled = false;
    }
  }

  function renderChartFields() {
    const type = $("#chartType").value;
    const numeric = numericColumns();
    let markup = "";
    if (["histogram", "qq", "density"].includes(type)) markup = selectField("chartVariable", "SAYISAL DEĞİŞKEN", numeric, numeric[0] || "", true);
    else if (["box", "violin"].includes(type)) {
      markup = selectField("chartVariable", "SAYISAL DEĞİŞKEN", numeric, numeric[0] || "", true);
      markup += selectField("chartGroup", "İSTEĞE BAĞLI GRUP", state.columns, "", true);
    }
    else if (type === "scatter") {
      markup = selectField("chartX", "X EKSENİ", numeric, numeric[0] || "", true);
      markup += selectField("chartY", "Y EKSENİ", numeric, numeric[1] || numeric[0] || "", true);
    } else markup = selectField("chartVariable", "KATEGORİK DEĞİŞKEN (SAYISAL KODLAR DAHİL)", state.columns, state.columns[0] || "", true);
    $("#chartFields").innerHTML = markup || '<span class="analysis-hint">Bu grafik türü için uygun değişken bulunamadı.</span>';
  }

  function drawChart() {
    if (!window.Chart) {
      return setChartError("Grafik kütüphanesi yüklenemedi. İnternet bağlantınızı kontrol edin.");
    }
    const type = $("#chartType").value;
    if (!activeRows().length) return setChartError("Etkin filtreye uyan gözlem yok.");
    let chartType; let labels = []; let datasets = []; let title = "";
    const palette = ["#5d7ee4", "#45a982", "#e7a84c", "#9a79dc", "#e27070", "#54a8b6"];
    if (type === "histogram") {
      const column = $("#chartVariable").value;
      const values = columnValues(column).map(parseNumber).filter((value) => value !== null);
      if (!values.length) return setChartError("Seçilen değişkende sayısal gözlem yok.");
      const min = Math.min(...values); const max = Math.max(...values);
      const binCount = Math.min(20, Math.max(5, Math.ceil(Math.sqrt(values.length))));
      const width = (max - min) / binCount || 1;
      const counts = Array(binCount).fill(0);
      values.forEach((value) => counts[Math.min(binCount - 1, Math.floor((value - min) / width))]++);
      labels = counts.map((_, index) => `${fmt(min + index * width, 2)} – ${fmt(min + (index + 1) * width, 2)}`);
      datasets = [{ label: "Frekans", data: counts, backgroundColor: "#6d89e8cc", borderColor: "#5d7ee4", borderWidth: 1, borderRadius: 3 }];
      chartType = "bar"; title = `Histogram · ${column}`;
    } else if (type === "qq") {
      const column = $("#chartVariable").value;
      const values = columnValues(column).map(parseNumber).filter((value) => value !== null).sort((a, b) => a - b);
      if (values.length < 3) return setChartError("Q–Q grafiği için en az 3 sayısal gözlem gerekir.");
      const stats = summarize(values);
      datasets = [{ label: "Gözlem / normal beklenen", data: values.map((value, index) => ({
        x: jStat.normal.inv((index + 0.5) / values.length, stats.mean, stats.sd), y: value,
      })), backgroundColor: "#5d7ee4bb", pointRadius: 4 }];
      chartType = "scatter"; title = `Normal Q–Q grafiği · ${column}`;
    } else if (type === "density") {
      const column = $("#chartVariable").value;
      const values = columnValues(column).map(parseNumber).filter((value) => value !== null);
      if (values.length < 2) return setChartError("Yoğunluk grafiği için en az 2 sayısal gözlem gerekir.");
      const stats = summarize(values);
      const bandwidth = Math.max(1e-8, 1.06 * stats.sd * values.length ** -0.2);
      const min = stats.min - 3 * bandwidth; const max = stats.max + 3 * bandwidth;
      const points = Array.from({ length: 80 }, (_, index) => {
        const x = min + (max - min) * index / 79;
        const density = values.reduce((sum, value) => sum + jStat.normal.pdf((x - value) / bandwidth, 0, 1), 0) / (values.length * bandwidth);
        return { x, y: density };
      });
      datasets = [{ label: "Yoğunluk", data: points, borderColor: "#5d7ee4", backgroundColor: "#5d7ee422", fill: true, pointRadius: 0, tension: 0.25 }];
      chartType = "line"; title = `Kernel yoğunluk tahmini · ${column}`;
    } else if (type === "scatter") {
      const x = $("#chartX").value; const y = $("#chartY").value;
      const points = activeRows().map((row) => ({ x: parseNumber(row[x]), y: parseNumber(row[y]) }))
        .filter((point) => point.x !== null && point.y !== null);
      if (!points.length) return setChartError("Seçilen değişkenlerde eşleşen gözlem yok.");
      datasets = [{ label: `${x} / ${y}`, data: points, backgroundColor: "#5d7ee4bb", pointRadius: 4 }];
      chartType = "scatter"; title = `Dağılım grafiği · ${x} / ${y}`;
    } else if (type === "box" || type === "violin") {
      const column = $("#chartVariable").value;
      const group = $("#chartGroup").value;
      const grouped = new Map();
      activeRows().forEach((row) => {
        const value = parseNumber(row[column]);
        const category = group ? String(row[group] ?? "") : column;
        if (value !== null && category) {
          if (!grouped.has(category)) grouped.set(category, []);
          grouped.get(category).push(value);
        }
      });
      const groups = [...grouped.entries()].filter(([, values]) => values.length);
      if (!groups.length) return setChartError("Grafik için sayısal gözlem bulunamadı.");
      if (type === "box") {
        const boxData = groups.map(([name, values]) => {
          const stats = summarize(values);
          const sorted = values.slice().sort((a, b) => a - b);
          const spread = stats.q3 - stats.q1;
          const lowerFence = stats.q1 - 1.5 * spread;
          const upperFence = stats.q3 + 1.5 * spread;
          return {
            name, ...stats,
            min: sorted.find((value) => value >= lowerFence),
            max: sorted.slice().reverse().find((value) => value <= upperFence),
          };
        });
        datasets = groups.map(([name, values], index) => ({
          label: name, data: values.map((value) => ({ x: index + (Math.random() - 0.5) * 0.12, y: value })),
          backgroundColor: palette[index % palette.length], pointRadius: 2.5,
        }));
        chartType = "scatter"; labels = groups.map(([name]) => name);
        title = `Kutu grafiği · ${column}`;
        state.chartOptions = { boxData };
      } else {
        const allValues = groups.flatMap(([, values]) => values);
        const min = Math.min(...allValues); const max = Math.max(...allValues);
        const range = max - min || 1;
        const points = 48;
        datasets = groups.map(([name, values], index) => {
          const stats = summarize(values);
          const bandwidth = Math.max(1e-8, 1.06 * stats.sd * values.length ** -0.2);
          const data = Array.from({ length: points * 2 }, (_, point) => {
            const reverse = point >= points;
            const valueIndex = reverse ? points * 2 - point - 1 : point;
            const y = min + range * valueIndex / (points - 1);
            const density = values.reduce((sum, value) => sum + jStat.normal.pdf((y - value) / bandwidth, 0, 1), 0) / (values.length * bandwidth);
            const maxDensity = 1 / bandwidth / Math.sqrt(2 * Math.PI);
            return { x: index + (reverse ? -1 : 1) * 0.35 * density / maxDensity, y };
          });
          return { label: name, data, borderColor: palette[index % palette.length], backgroundColor: `${palette[index % palette.length]}33`, fill: true, pointRadius: 0, showLine: true };
        });
        chartType = "scatter"; title = `Violin grafiği · ${column}`;
      }
    } else {
      const column = $("#chartVariable").value;
      const counts = new Map();
      columnValues(column).forEach((value) => {
        if (!isMissing(value)) counts.set(String(value), (counts.get(String(value)) || 0) + 1);
      });
      const top = [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 25);
      if (!top.length) return setChartError("Seçilen değişkende kategori bulunamadı.");
      labels = top.map(([label]) => label);
      datasets = [{ label: "Frekans", data: top.map(([, count]) => count), backgroundColor: labels.map((_, index) => palette[index % palette.length]), borderRadius: 3 }];
      chartType = "bar"; title = `Kategori grafiği · ${column}`;
    }
    if (state.chart) state.chart.destroy();
    $("#chartEmpty").hidden = true;
    const boxPlugin = type === "box" ? {
      id: "verilabBox",
      afterDatasetsDraw(chart) {
        const boxes = state.chartOptions?.boxData || [];
        const { ctx, scales } = chart;
        const xScale = scales.x; const yScale = scales.y;
        boxes.forEach((box, index) => {
          const x = xScale.getPixelForValue(index);
          const left = xScale.getPixelForValue(index - 0.22);
          const right = xScale.getPixelForValue(index + 0.22);
          const color = palette[index % palette.length];
          ctx.save();
          ctx.strokeStyle = color; ctx.fillStyle = `${color}33`; ctx.lineWidth = 2;
          ctx.beginPath(); ctx.moveTo(x, yScale.getPixelForValue(box.min)); ctx.lineTo(x, yScale.getPixelForValue(box.max)); ctx.moveTo(x - 7, yScale.getPixelForValue(box.min)); ctx.lineTo(x + 7, yScale.getPixelForValue(box.min)); ctx.moveTo(x - 7, yScale.getPixelForValue(box.max)); ctx.lineTo(x + 7, yScale.getPixelForValue(box.max)); ctx.stroke();
          ctx.fillRect(left, yScale.getPixelForValue(box.q3), right - left, yScale.getPixelForValue(box.q1) - yScale.getPixelForValue(box.q3));
          ctx.strokeRect(left, yScale.getPixelForValue(box.q3), right - left, yScale.getPixelForValue(box.q1) - yScale.getPixelForValue(box.q3));
          ctx.beginPath(); ctx.moveTo(left, yScale.getPixelForValue(box.median)); ctx.lineTo(right, yScale.getPixelForValue(box.median)); ctx.stroke();
          ctx.restore();
        });
      },
    } : null;
    state.chartOptions = state.chartOptions || {};
    state.chart = new Chart($("#chartCanvas"), {
      type: chartType,
      data: { labels, datasets },
      plugins: boxPlugin ? [boxPlugin] : [],
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { title: { display: true, text: title, align: "start", color: "#35425b", font: { family: "DM Sans", size: 13, weight: "600" }, padding: { bottom: 20 } }, legend: { display: type !== "bar", labels: { color: "#788397", boxWidth: 10 } } },
        scales: type === "scatter" || type === "qq" || type === "density" || type === "box" || type === "violin" ? {
          x: { type: "linear", title: { display: true, text: type === "qq" ? "Normal beklenen kuantiller" : type === "box" || type === "violin" ? "Grup (jitter / yoğunluk)" : $("#chartX")?.value || "X" }, grid: { color: "#f0f2f6" } },
          y: { title: { display: true, text: type === "qq" ? "Gözlenen kuantiller" : type === "density" ? "Yoğunluk" : $("#chartY")?.value || "Y" }, grid: { color: "#f0f2f6" } },
        } : { x: { grid: { display: false }, ticks: { color: "#8993a4", maxRotation: 45, minRotation: 0 } }, y: { beginAtZero: true, grid: { color: "#f0f2f6" }, ticks: { color: "#8993a4" } } },
      },
    });
    $("#downloadChart").disabled = false;
  }

  function downloadChart() {
    if (!state.chart) return;
    const link = document.createElement("a");
    link.href = state.chart.toBase64Image("image/png", 1);
    link.download = "biz-bize-grafik.png";
    document.body.append(link);
    link.click();
    link.remove();
  }

  function setChartError(message) {
    if (state.chart) {
      state.chart.destroy();
      state.chart = null;
    }
    $("#downloadChart").disabled = true;
    $("#chartEmpty").textContent = message;
    $("#chartEmpty").hidden = false;
  }

  function switchTab(tabName) {
    document.querySelectorAll(".tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.tab === tabName));
    $("#dataTab").hidden = false;
    $("#dataOverview").hidden = tabName !== "data";
    $("#analysisTab").hidden = tabName !== "analysis";
    $("#chartsTab").hidden = tabName !== "charts";
    $("#rTab").hidden = tabName !== "r";
    if (tabName === "charts" && state.chart) state.chart.resize();
  }

  $("#browseButton").addEventListener("click", () => $("#fileInput").click());
  $("#replaceFileButton").addEventListener("click", () => $("#fileInput").click());
  $("#fileInput").addEventListener("change", (event) => {
    loadFile(event.target.files[0]);
    event.target.value = "";
  });
  $("#dropZone").addEventListener("dragover", (event) => { event.preventDefault(); $("#dropZone").classList.add("drag-over"); });
  $("#dropZone").addEventListener("dragleave", () => $("#dropZone").classList.remove("drag-over"));
  $("#dropZone").addEventListener("drop", (event) => {
    event.preventDefault();
    $("#dropZone").classList.remove("drag-over");
    loadFile(event.dataTransfer.files[0]);
  });
  $("#sheetSelect").addEventListener("change", (event) => readSheet(Number(event.target.value)));
  $("#filterOperator").addEventListener("change", () => {
    $("#filterValue").disabled = ["empty", "notEmpty"].includes($("#filterOperator").value);
  });
  $("#applyFilter").addEventListener("click", () => {
    const operator = $("#filterOperator").value;
    const value = $("#filterValue").value.trim();
    const error = $("#filterError");
    if (!["empty", "notEmpty"].includes(operator) && !value) {
      error.textContent = "Filtre koşulu için bir değer girin.";
      error.hidden = false;
      return;
    }
    error.hidden = true;
    state.filter = { column: $("#filterColumn").value, operator, value };
    updateSummary();
    renderDataTable();
  });
  $("#clearFilter").addEventListener("click", () => {
    state.filter = null;
    $("#filterValue").value = "";
    $("#filterError").hidden = true;
    updateSummary();
    renderDataTable();
  });
  $("#applySort").addEventListener("click", () => {
    state.sort = { column: $("#sortColumn").value, direction: $("#sortDirection").value };
    renderDataTable();
  });
  $("#clearSort").addEventListener("click", () => {
    state.sort = null;
    renderDataTable();
  });
  $("#addComputedColumn").addEventListener("click", addComputedColumn);
  $("#addRow").addEventListener("click", addDataRow);
  $("#openDataTools").addEventListener("click", () => {
    $("#dataTools").open = !$("#dataTools").open;
  });
  $("#exportDataCsv").addEventListener("click", () => {
    try { downloadData("csv"); } catch (error) { showError(error.message || "CSV dışa aktarılamadı."); }
  });
  $("#exportDataExcel").addEventListener("click", () => {
    try { downloadData("excel"); } catch (error) { showError(error.message || "Excel dışa aktarılamadı."); }
  });
  $("#analysisType").addEventListener("change", renderAnalysisFields);
  $("#advancedAnalysis").addEventListener("change", renderRAnalysisFields);
  $("#startR").addEventListener("click", startR);
  $("#installRPackages").addEventListener("click", () => {
    $("#installRPackages").disabled = true;
    installRPackages().catch((error) => {
      $("#rProgress").textContent = error.message || "R paketleri yüklenemedi.";
      $("#rProgress").classList.add("error-message");
    }).finally(() => { $("#installRPackages").disabled = !state.webR; });
  });
  $("#runR").addEventListener("click", runRCode);
  $("#clearR").addEventListener("click", () => { $("#rCode").value = ""; $("#rOutput").textContent = "R çıktısı burada görüntülenir."; });
  $("#runAnalysis").addEventListener("click", runAnalysis);
  $("#downloadResults").addEventListener("click", () => {
    try {
      downloadResults();
    } catch (error) {
      $("#resultsContent").insertAdjacentHTML("afterbegin", `<p class="error-message" role="alert">${escapeHtml(error.message || "CSV dosyası oluşturulamadı.")}</p>`);
    }
  });
  $("#downloadWorkbook").addEventListener("click", () => {
    try {
      downloadWorkbook();
    } catch (error) {
      $("#resultsContent").insertAdjacentHTML("afterbegin", `<p class="error-message" role="alert">${escapeHtml(error.message || "Excel dosyası oluşturulamadı.")}</p>`);
    }
  });
  $("#downloadReport").addEventListener("click", downloadReport);
  $("#chartType").addEventListener("change", renderChartFields);
  $("#drawChart").addEventListener("click", drawChart);
  $("#downloadChart").addEventListener("click", downloadChart);
  window.addEventListener("afterprint", () => { $("#printReport").hidden = true; });
  document.querySelectorAll(".tab").forEach((tab) => tab.addEventListener("click", () => switchTab(tab.dataset.tab)));
})();
