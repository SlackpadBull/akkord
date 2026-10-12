
"use strict";

// Номера колонок с нуля: A = 0, B = 1, C = 2...
const COL = {
  name: 0,      // A — название товара
  warehouse: 1, // B — склад
  barcode: 3,   // D — штрихкод
  stock: 6      // G — остаток
};

const CENTRAL = "Центральный склад";
const STORE = "Склад точки Аккорд";

const $ = id => document.getElementById(id);

let workbook = null;
let sourceFile = "";
let activeTab = "missing";
let results = {
  missing: [],
  low: [],
  warnings: []
};

function normalize(value) {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toLocaleLowerCase("ru")
    .replace(/ё/g, "е");
}

function normalizeBarcode(value) {
  return String(value ?? "").trim().replace(/\s+/g, "");
}

function quantity(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : NaN;
  }

  const text = String(value ?? "")
    .trim()
    .replace(/\u00a0/g, "")
    .replace(/\s/g, "")
    .replace(",", ".");

  if (!text) return NaN;

  const result = Number(text);
  return Number.isFinite(result) ? result : NaN;
}

function warehouseType(value) {
  const name = normalize(value);

  if (name === normalize(CENTRAL)) return "central";
  if (name === normalize(STORE)) return "store";

  return null;
}

function escapeHTML(value) {
  return String(value ?? "").replace(/[&<>"']/g, char => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[char]);
}

function showStatus(message, type = "") {
  $("status").textContent = message;
  $("status").className = type;
}

function resetResults() {
  results = { missing: [], low: [], warnings: [] };
  $("results").hidden = true;
  $("exportAll").disabled = true;
  $("exportCurrent").disabled = true;
}

function loadFile(file) {
  if (!file) return;

  if (!/\.(xlsx|xls|xlsm|csv)$/i.test(file.name)) {
    showStatus("Выбери файл Excel или CSV.", "error");
    return;
  }

  if (typeof XLSX === "undefined") {
    showStatus(
      "Не загрузилась библиотека Excel. Проверь подключение к интернету.",
      "error"
    );
    return;
  }

  resetResults();
  workbook = null;
  sourceFile = file.name;

  $("fileLabel").textContent = file.name;
  $("analyze").disabled = true;
  $("sheetSelect").disabled = true;

  showStatus("Читаю файл...", "info");

  const reader = new FileReader();

  reader.onerror = () => {
    showStatus("Не удалось прочитать файл.", "error");
  };

  reader.onload = event => {
    try {
      workbook = XLSX.read(event.target.result, {
        type: "array",
        cellText: true,
        cellDates: false
      });

      if (!workbook.SheetNames.length) {
        throw new Error("В файле нет листов.");
      }

      const select = $("sheetSelect");
      select.innerHTML = "";

      workbook.SheetNames.forEach(name => {
        const option = document.createElement("option");
        option.value = name;
        option.textContent = name;
        select.appendChild(option);
      });

      select.disabled = false;
      $("analyze").disabled = false;

      showStatus(
        `Файл загружен. Листов: ${workbook.SheetNames.length}.`,
        "success"
      );
    } catch (error) {
      workbook = null;
      $("analyze").disabled = true;
      showStatus(`Ошибка чтения файла: ${error.message}`, "error");
    }
  };

  reader.readAsArrayBuffer(file);
}

function analyze() {
  if (!workbook) {
    showStatus("Сначала загрузи Excel-файл.", "error");
    return;
  }

  const threshold = Number($("threshold").value);

  if (!Number.isInteger(threshold) || threshold < 2) {
    showStatus("Порог должен быть целым числом не меньше 2.", "error");
    return;
  }

  const sheet = workbook.Sheets[$("sheetSelect").value];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false
  });

  const products = new Map();
  const warnings = [];

  let recognizedRows = 0;
  let ignoredWarehouses = 0;

  rows.forEach((row, index) => {
    const warehouse = row[COL.warehouse];
    const barcode = normalizeBarcode(row[COL.barcode]);
    const name = String(row[COL.name] ?? "").trim();
    const rawStock = row[COL.stock];

    // Пропускаем обычную строку заголовков.
    if (
      index === 0 &&
      normalize(warehouse).includes("склад") &&
      normalize(row[COL.barcode]).includes("штрих")
    ) {
      return;
    }

    if (
      !String(warehouse ?? "").trim() &&
      !barcode &&
      !String(rawStock ?? "").trim()
    ) {
      return;
    }

    const type = warehouseType(warehouse);

    if (!type) {
      if (String(warehouse ?? "").trim()) ignoredWarehouses++;
      return;
    }

    recognizedRows++;

    if (!barcode) {
      warnings.push({
        type: "Нет штрихкода",
        name: name || "Без названия",
        barcode: "—",
        warehouse: String(warehouse),
        stock: String(rawStock ?? "") || "—",
        row: index + 1,
        details: "Нельзя сопоставить товар без штрихкода."
      });
      return;
    }

    const stock = quantity(rawStock);

    if (!Number.isFinite(stock) || stock < 0) {
      warnings.push({
        type: "Некорректный остаток",
        name: name || "Без названия",
        barcode,
        warehouse: String(warehouse),
        stock: String(rawStock ?? "") || "—",
        row: index + 1,
        details: "Остаток пустой, нечисловой или отрицательный."
      });
      return;
    }

    if (!products.has(barcode)) {
      products.set(barcode, {
        barcode,
        name: name || barcode,
        names: new Set(),
        central: 0,
        store: 0,
        centralRows: 0,
        storeRows: 0
      });
    }

    const product = products.get(barcode);

    if (name) product.names.add(name);

    if (type === "central") {
      product.central += stock;
      product.centralRows++;
    } else {
      product.store += stock;
      product.storeRows++;
    }
  });

  const missing = [];
  const low = [];

  for (const product of products.values()) {
    if (product.names.size > 1) {
      warnings.push({
        type: "Разные названия",
        name: [...product.names].join(" / "),
        barcode: product.barcode,
        warehouse: "Несколько",
        stock: "—",
        row: "—",
        details: "У одного штрихкода обнаружены разные названия."
      });
    }

    if (product.centralRows > 1 || product.storeRows > 1) {
      warnings.push({
        type: "Повторные строки",
        name: product.name,
        barcode: product.barcode,
        warehouse: [
          product.centralRows > 1
            ? `${CENTRAL}: ${product.centralRows} строк`
            : "",
          product.storeRows > 1
            ? `${STORE}: ${product.storeRows} строк`
            : ""
        ].filter(Boolean).join("; "),
        stock: "Суммированы",
        row: "—",
        details: "Остатки повторяющихся строк сложены. Проверь, что это не дубли."
      });
    }

    const centralStock = product.centralRows > 0
      ? product.central
      : 0;

    // Если записи точки нет, её остаток считается нулевым.
    const storeStock = product.storeRows > 0
      ? product.store
      : 0;

    // Товар должен физически быть доступен на центральном складе.
    if (centralStock <= 0) continue;

    const item = {
      name: product.name,
      barcode: product.barcode,
      central: centralStock,
      store: storeStock,

      // Рекомендация: довести остаток на точке до порога.
      recommended: Math.max(0, threshold - storeStock)
    };

    // СПИСОК 1:
    // товара нет на точке, но он есть на центральном складе.
    if (storeStock === 0) {
      missing.push(item);
    }

    // СПИСОК 2:
    // остаток на точке строго больше 0, но меньше порога.
    // При пороге 3 это означает только 1 или 2 штуки.
    // Товары с нулём сюда НЕ попадают.
    if (storeStock > 0 && storeStock < threshold) {
      low.push(item);
    }
  }

  const sortByName = (a, b) =>
    a.name.localeCompare(b.name, "ru");

  results.missing = missing.sort(sortByName);
  results.low = low.sort(sortByName);
  results.warnings = warnings;

  $("missingCount").textContent = missing.length;
  $("lowCount").textContent = low.length;
  $("warningCount").textContent = warnings.length;

  $("missingTabCount").textContent = missing.length;
  $("lowTabCount").textContent = low.length;
  $("warningTabCount").textContent = warnings.length;

  $("summary").textContent =
    `Обработано строк: ${recognizedRows}. ` +
    `Уникальных штрихкодов: ${products.size}.` +
    (ignoredWarehouses
      ? ` Строк с другими названиями складов пропущено: ${ignoredWarehouses}.`
      : "");

  $("results").hidden = false;
  $("exportAll").disabled = false;

  $("search").value = "";
  $("sort").value = "name";

  setActiveTab("missing");
  renderTable();

  showStatus(
    `Готово. Нет на точке: ${missing.length}. ` +
    `Остаток от 1 до ${threshold - 1}: ${low.length}.` +
    (warnings.length ? ` Предупреждений: ${warnings.length}.` : ""),
    "success"
  );
}

function setActiveTab(tabName) {
  activeTab = tabName;

  document.querySelectorAll(".tab").forEach(button => {
    const active = button.dataset.tab === tabName;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", String(active));
  });
}

function currentItems() {
  return results[activeTab] || [];
}

function renderTable() {
  const query = normalize($("search").value);
  const sort = $("sort").value;

  let items = currentItems().filter(item => {
    const text = normalize([
      item.name,
      item.barcode,
      item.warehouse,
      item.type,
      item.details
    ].join(" "));

    return text.includes(query);
  });

  if (activeTab !== "warnings") {
    items.sort((a, b) => {
      if (sort === "stock-desc") {
        return b.central - a.central || a.name.localeCompare(b.name, "ru");
      }

      if (sort === "stock-asc") {
        return a.central - b.central || a.name.localeCompare(b.name, "ru");
      }

      if (sort === "barcode") {
        return a.barcode.localeCompare(b.barcode, "ru");
      }

      return a.name.localeCompare(b.name, "ru");
    });
  }

  const head = $("tableHead");
  const body = $("tableBody");

  if (activeTab === "warnings") {
    head.innerHTML = `
      <tr>
        <th>Тип</th>
        <th>Товар</th>
        <th>Штрихкод</th>
        <th>Склад</th>
        <th>Строка Excel</th>
        <th>Комментарий</th>
      </tr>`;

    body.innerHTML = items.map(item => `
      <tr>
        <td><span class="pill low">${escapeHTML(item.type)}</span></td>
        <td class="product-name">${escapeHTML(item.name)}</td>
        <td class="barcode">${escapeHTML(item.barcode)}</td>
        <td>${escapeHTML(item.warehouse)}</td>
        <td>${escapeHTML(item.row)}</td>
        <td>${escapeHTML(item.details)}</td>
      </tr>
    `).join("");
  } else {
    head.innerHTML = `
      <tr>
        <th>Название товара</th>
        <th>Штрихкод</th>
        <th>Центральный склад</th>
        <th>Склад точки Аккорд</th>
        <th>Рекомендуется переместить</th>
      </tr>`;

    body.innerHTML = items.map(item => {
      const storeClass = item.store === 0
        ? "zero"
        : item.store < Number($("threshold").value)
          ? "low"
          : "available";

      return `
        <tr>
          <td class="product-name" title="${escapeHTML(item.name)}">
            ${escapeHTML(item.name)}
          </td>
          <td class="barcode">${escapeHTML(item.barcode)}</td>
          <td><span class="pill available">${item.central}</span></td>
          <td><span class="pill ${storeClass}">${item.store}</span></td>
          <td>${item.recommended}</td>
        </tr>
      `;
    }).join("");
  }

  $("empty").hidden = items.length > 0;
  $("tableFooter").textContent =
    `Показано ${items.length} из ${currentItems().length} товаров.`;

  $("exportCurrent").disabled = items.length === 0;
}

function exportWorkbook(sheets) {
  const output = XLSX.utils.book_new();

  for (const entry of sheets) {
    let data;

    if (entry.warnings) {
      data = entry.items.map(item => ({
        "Тип проблемы": item.type,
        "Название": item.name,
        "Штрихкод": item.barcode,
        "Склад": item.warehouse,
        "Остаток": item.stock,
        "Строка Excel": item.row,
        "Комментарий": item.details
      }));
    } else {
      data = entry.items.map(item => ({
        "Название товара": item.name,
        "Штрихкод": item.barcode,
        "Центральный склад": item.central,
        "Склад точки Аккорд": item.store,
        "Рекомендуется переместить": item.recommended
      }));
    }

    const worksheet = XLSX.utils.json_to_sheet(data);

    XLSX.utils.book_append_sheet(
      output,
      worksheet,
      entry.name.slice(0, 31)
    );
  }

  const baseName = sourceFile.replace(/\.[^.]+$/, "")
    .replace(/[\\/:*?"<>|]+/g, "_");

  XLSX.writeFile(
    output,
    `Перемещение_${baseName || "товары"}.xlsx`
  );
}

// Загрузка файла через кнопку.
$("fileInput").addEventListener("change", event => {
  loadFile(event.target.files[0]);
});

// Перетаскивание файла.
const dropzone = $("dropzone");

["dragenter", "dragover"].forEach(eventName => {
  dropzone.addEventListener(eventName, event => {
    event.preventDefault();
    dropzone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach(eventName => {
  dropzone.addEventListener(eventName, event => {
    event.preventDefault();
    dropzone.classList.remove("dragover");
  });
});

dropzone.addEventListener("drop", event => {
  const file = event.dataTransfer.files[0];
  if (file) loadFile(file);
});

$("analyze").addEventListener("click", analyze);
$("search").addEventListener("input", renderTable);
$("sort").addEventListener("change", renderTable);

document.querySelectorAll(".tab").forEach(button => {
  button.addEventListener("click", () => {
    setActiveTab(button.dataset.tab);
    renderTable();
  });
});

$("exportCurrent").addEventListener("click", () => {
  const names = {
    missing: "Нет на точке",
    low: "Меньше порога",
    warnings: "Предупреждения"
  };

  exportWorkbook([{
    name: names[activeTab],
    items: currentItems(),
    warnings: activeTab === "warnings"
  }]);
});

$("exportAll").addEventListener("click", () => {
  exportWorkbook([
    { name: "Нет на точке", items: results.missing },
    { name: "Меньше порога", items: results.low },
    { name: "Предупреждения", items: results.warnings, warnings: true }
  ]);
});
