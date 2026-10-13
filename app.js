
const CENTRAL_WAREHOUSE = "Центральный склад";
const STORE_WAREHOUSE = "Склад точки Аккорд";

// Нумерация колонок начинается с нуля:
// A = 0, B = 1, C = 2 и так далее
const COLUMNS = {
  warehouse: 1, // B — склад
  name: 2,      // C — название товара
  barcode: 3,   // D — штрихкод
  price: 4,     // E — цена
  stock: 6      // G — количество
};

const RESTOCK_EXCLUSIONS = [
  "Акустическая система",
  "Гитара",
  "Домбра",
  "Комбоусилитель",
  "Комплект",
  "Коробка",
  "Микрофон",
  "MIDI-контроллер",
  "Синтезатор",
  "Скрипка",
  "Усилитель",
  "Ударная установка",
  "Укулеле",
  "Флейта",
  "Электрогитара",
  "Электроскрипка",
  "Звуковая карта",
  "6358",
  "Аудио интерфейс",
];

const RESTOCK_ALLOWLIST = [
  "Домбра 48р",
  "Домбра 46 простая",
  "Кнопка домбра",
  "Домбра 44"
];

const fileInput = document.getElementById("excelFile");
const fileName = document.getElementById("fileName");
const exportBtn = document.getElementById("exportBtn");
const statusEl = document.getElementById("status");
const resultsSection = document.getElementById("results");
const resultsBody = document.getElementById("resultsBody");
const emptyMessage = document.getElementById("emptyMessage");
const resultsTitle = document.getElementById("resultsTitle");
const stockHeader = document.getElementById("stockHeader");
const centralStockHeader = document.getElementById("centralStockHeader");
const centralTab = document.getElementById("centralTab");
const restockTab = document.getElementById("restockTab");
const restockViews = document.getElementById("restockViews");
const restockProductsTab = document.getElementById("restockProductsTab");
const restockExceptionsTab = document.getElementById("restockExceptionsTab");
const resultsPanel = document.getElementById("resultsPanel");
const exceptionsInfo = document.getElementById("exceptionsInfo");
const exceptionsList = document.getElementById("exceptionsList");
const exceptionsAllowed = document.getElementById("exceptionsAllowed");

let workbook = null;
let results = [];
let restockResults = [];
let excludedRestockResults = [];
let activeList = "central";
let selectedFileName = "";

fileInput.addEventListener("change", async () => {
  const file = fileInput.files[0];

  workbook = null;
  results = [];
  restockResults = [];
  excludedRestockResults = [];
  activeList = "central";
  centralTab.classList.add("is-active");
  centralTab.setAttribute("aria-selected", "true");
  restockTab.classList.remove("is-active");
  restockTab.setAttribute("aria-selected", "false");
  restockProductsTab.classList.add("is-active");
  restockProductsTab.setAttribute("aria-selected", "true");
  restockExceptionsTab.classList.remove("is-active");
  restockExceptionsTab.setAttribute("aria-selected", "false");
  resultsPanel.setAttribute("aria-labelledby", "centralTab");
  exportBtn.disabled = true;
  resultsSection.hidden = true;
  document.body.classList.remove("has-results");
  statusEl.textContent = "";
  statusEl.className = "status";

  if (!file) {
    selectedFileName = "";
    fileName.textContent = "Выберите Excel-файл";
    return;
  }

  selectedFileName = file.name;
  fileName.textContent = file.name;
  statusEl.textContent = "Читаем файл…";

  try {
    const buffer = await file.arrayBuffer();

    workbook = XLSX.read(buffer, {
      type: "array",
      cellText: true,
      cellDates: false
    });

    if (!workbook.SheetNames.length) {
      throw new Error("В файле нет листов.");
    }

    analyzeStock();
  } catch (error) {
    workbook = null;
    statusEl.textContent = `Не удалось прочитать файл: ${error.message}`;
    statusEl.className = "status error";
  }
});

exportBtn.addEventListener("click", exportResults);
centralTab.addEventListener("click", () => selectResultsTab("central"));
restockTab.addEventListener("click", () => selectResultsTab("restock"));
restockProductsTab.addEventListener("click", () => selectRestockView("restock"));
restockExceptionsTab.addEventListener("click", () => selectRestockView("exceptions"));
exceptionsList.textContent = RESTOCK_EXCLUSIONS.join(", ");
exceptionsAllowed.textContent = `Разрешено: ${RESTOCK_ALLOWLIST.join(", ")}`;

function normalizeText(value) {
  return String(value ?? "").trim();
}

function normalizeBarcode(value) {
  if (value === null || value === undefined) return "";

  // Штрихкод не должен зависеть от пробелов по краям.
  // Ведущие нули сохраняются, если Excel хранит их в ячейке как текст.
  return String(value).trim();
}

function parseStock(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  const normalized = String(value ?? "")
    .trim()
    .replace(/\s/g, "")
    .replace(",", ".");

  if (!normalized) return null;

  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function isZeroPrice(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) && value === 0;
  }

  const normalized = String(value ?? "")
    .trim()
    .replace(/[\s\u00a0]/g, "")
    .replace(/[₽$€£]/g, "")
    .replace(/(?:руб\.?|р\.?)/giu, "")
    .replace(",", ".");

  if (!normalized) return false;

  const price = Number(normalized);
  return Number.isFinite(price) && price === 0;
}

function hasRestockExclusion(name) {
  const normalizedName = normalizeText(name).toLocaleLowerCase("ru");
  return RESTOCK_EXCLUSIONS.some(exclusion =>
    normalizedName.includes(exclusion.toLocaleLowerCase("ru"))
  );
}

function isRestockAllowed(name) {
  const normalizedName = normalizeText(name).toLocaleLowerCase("ru");
  return RESTOCK_ALLOWLIST.some(
    allowedName => normalizedName === allowedName.toLocaleLowerCase("ru")
  );
}

function analyzeStock() {
  if (!workbook) {
    setError("Сначала выберите Excel-файл.");
    return;
  }

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false
  });

  const products = new Map();
  let matchedRows = 0;
  let skippedRows = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    const warehouse = normalizeText(row[COLUMNS.warehouse]);
    const barcode = normalizeBarcode(row[COLUMNS.barcode]);
    const name = normalizeText(row[COLUMNS.name]);
    const price = normalizeText(row[COLUMNS.price]);
    const stock = parseStock(row[COLUMNS.stock]);

    // Нас интересуют только два указанных склада.
    if (
      warehouse !== CENTRAL_WAREHOUSE &&
      warehouse !== STORE_WAREHOUSE
    ) {
      continue;
    }

    // Пропускаем заголовки и строки без штрихкода.
    if (!barcode || !name || stock === null) {
      skippedRows++;
      continue;
    }

    matchedRows++;

    const productKey = JSON.stringify([barcode, name]);

    if (!products.has(productKey)) {
      products.set(productKey, {
        barcode,
        name,
        price,
        centralStock: 0,
        storeStock: 0,
        hasCentralRow: false,
        hasStoreRow: false
      });
    }

    const product = products.get(productKey);

    if (!product.name && name) {
      product.name = name;
    }
    if (!product.price && price) {
      product.price = price;
    }

    if (warehouse === CENTRAL_WAREHOUSE) {
      product.hasCentralRow = true;
      product.centralStock += stock;
    }

    if (warehouse === STORE_WAREHOUSE) {
      product.hasStoreRow = true;
      product.storeStock += stock;
    }
  }

  results = [];
  restockResults = [];
  excludedRestockResults = [];

  for (const product of products.values()) {
    // Требуется именно наличие строки на центральном складе
    // и отсутствие строки с тем же штрихкодом на складе точки.
    if (
      product.hasCentralRow &&
      !product.hasStoreRow &&
      !isZeroPrice(product.price)
    ) {
      results.push({
        name: product.name,
        barcode: product.barcode,
        stock: product.centralStock,
        price: product.price
      });
    }

    if (
      product.storeStock > 0 &&
      product.storeStock < 2 &&
      product.centralStock > 0
    ) {
      const restockProduct = {
        name: product.name,
        barcode: product.barcode,
        storeStock: product.storeStock,
        centralStock: product.centralStock,
        price: product.price
      };

      if (
        hasRestockExclusion(product.name) &&
        !isRestockAllowed(product.name)
      ) {
        excludedRestockResults.push(restockProduct);
      } else {
        restockResults.push(restockProduct);
      }
    }
  }

  const sortByName = (a, b) =>
    a.name.localeCompare(b.name, "ru", { sensitivity: "base" });
  results.sort(sortByName);
  restockResults.sort(sortByName);
  excludedRestockResults.sort(sortByName);

  renderResults();

  statusEl.className = "status";
  statusEl.textContent =
    `Проверено: ${matchedRows}, пропущено: ${skippedRows}`;
}

function renderResults() {
  const isCentralList = activeList === "central";
  const isExceptionsList = activeList === "exceptions";
  const isRestockList = !isCentralList;
  const activeResults = isCentralList
    ? results
    : isExceptionsList
      ? excludedRestockResults
      : restockResults;

  restockViews.hidden = isCentralList;
  restockProductsTab.classList.toggle("is-active", activeList === "restock");
  restockProductsTab.setAttribute("aria-selected", String(activeList === "restock"));
  restockExceptionsTab.classList.toggle("is-active", isExceptionsList);
  restockExceptionsTab.setAttribute("aria-selected", String(isExceptionsList));
  exceptionsInfo.hidden = !isExceptionsList;
  resultsBody.replaceChildren();

  for (const product of activeResults) {
    const tr = document.createElement("tr");

    addCell(tr, product.name);
    addCell(tr, product.barcode);
    addCell(
      tr,
      formatStock(isRestockList ? product.storeStock : product.stock)
    );

    if (isRestockList) {
      addCell(tr, formatStock(product.centralStock));
    }
    addCell(tr, product.price);

    resultsBody.appendChild(tr);
  }

  resultsTitle.textContent = isCentralList
    ? "Товары только на центральном складе"
    : isExceptionsList
      ? "Товары, исключённые из пополнения"
      : "Товары, которые нужно пополнить";
  stockHeader.textContent = isRestockList
    ? "Остаток в Аккорде"
    : "Остаток ЦС";
  centralStockHeader.textContent = "Остаток ЦС";
  centralStockHeader.hidden = !isRestockList;
  centralTab.textContent = `Только на центральном ${results.length}`;
  restockTab.textContent = `Нужно пополнить ${restockResults.length}`;
  fileName.textContent = selectedFileName;
  emptyMessage.textContent = isCentralList
    ? "Товаров, которые есть только на центральном складе, не найдено."
    : isExceptionsList
      ? "Исключённых товаров для пополнения не найдено."
      : "Товаров для пополнения не найдено.";
  emptyMessage.hidden = activeResults.length !== 0;
  resultsSection.hidden = false;
  document.body.classList.add("has-results");
  exportBtn.disabled = activeResults.length === 0;
}

function selectResultsTab(tab) {
  activeList = tab;
  const isCentralList = tab === "central";
  const isRestockList = tab === "restock";

  centralTab.classList.toggle("is-active", isCentralList);
  centralTab.setAttribute("aria-selected", String(isCentralList));
  restockTab.classList.toggle("is-active", isRestockList);
  restockTab.setAttribute("aria-selected", String(isRestockList));
  resultsPanel.setAttribute(
    "aria-labelledby",
    isCentralList ? "centralTab" : "restockTab"
  );

  renderResults();
}

function selectRestockView(view) {
  activeList = view;
  renderResults();
}

function addCell(row, value) {
  const td = document.createElement("td");
  td.textContent = String(value ?? "");
  row.appendChild(td);
}

function formatStock(value) {
  return Number.isInteger(value)
    ? String(value)
    : String(Number(value.toFixed(3)));
}

function exportResults() {
  const isCentralList = activeList === "central";
  const isRestockList = !isCentralList;
  const isExceptionsList = activeList === "exceptions";
  const activeResults = isCentralList
    ? results
    : isExceptionsList
      ? excludedRestockResults
      : restockResults;
  if (!activeResults.length) return;

  const data = [
    isRestockList
      ? ["Товар", "Штрихкод", "Остаток в Аккорде", "Остаток ЦС", "Цена"]
      : ["Товар", "Штрихкод", "Остаток ЦС", "Цена"],
    ...activeResults.map(product => isRestockList
      ? [product.name, product.barcode, product.storeStock, product.centralStock, product.price]
      : [product.name, product.barcode, product.stock, product.price]
    )
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(data);
  worksheet["!cols"] = [
    { wch: 40 },
    { wch: 24 },
    { wch: 24 },
    ...(isRestockList ? [{ wch: 30 }] : []),
    { wch: 16 }
  ];

  const exportWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    exportWorkbook,
    worksheet,
    isCentralList
      ? "Только центральный склад"
      : isExceptionsList
        ? "Исключения для пополнения"
        : "Нужно пополнить"
  );

  XLSX.writeFile(
    exportWorkbook,
    isCentralList
      ? "Товары_только_на_центральном.xlsx"
      : isExceptionsList
        ? "Исключения_для_пополнения.xlsx"
        : "Нужно_пополнить.xlsx"
  );
}

function setError(message) {
  statusEl.textContent = message;
  statusEl.className = "status error";
}
