
const CENTRAL_WAREHOUSE = "Центральный склад";
const STORE_WAREHOUSE = "Склад точки Аккорд";

// Нумерация колонок начинается с нуля:
// A = 0, B = 1, C = 2 и так далее
const COLUMNS = {
  warehouse: 1, // B — склад
  name: 2,      // C — название товара
  barcode: 3,   // D — штрихкод
  stock: 6      // G — количество
};

const fileInput = document.getElementById("excelFile");
const fileName = document.getElementById("fileName");
const analyzeBtn = document.getElementById("analyzeBtn");
const exportBtn = document.getElementById("exportBtn");
const statusEl = document.getElementById("status");
const resultsSection = document.getElementById("results");
const resultsBody = document.getElementById("resultsBody");
const resultCount = document.getElementById("resultCount");
const emptyMessage = document.getElementById("emptyMessage");

let workbook = null;
let results = [];

fileInput.addEventListener("change", async () => {
  const file = fileInput.files[0];

  workbook = null;
  results = [];
  analyzeBtn.disabled = true;
  exportBtn.disabled = true;
  resultsSection.hidden = true;
  statusEl.textContent = "";
  statusEl.className = "status";

  if (!file) {
    fileName.textContent = "Выберите Excel-файл";
    return;
  }

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

    analyzeBtn.disabled = false;
    statusEl.textContent =
      `Файл загружен. Листов: ${workbook.SheetNames.length}.`;
  } catch (error) {
    workbook = null;
    statusEl.textContent = `Не удалось прочитать файл: ${error.message}`;
    statusEl.className = "status error";
  }
});

analyzeBtn.addEventListener("click", analyzeStock);

exportBtn.addEventListener("click", exportResults);

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

  for (const product of products.values()) {
    // Требуется именно наличие строки на центральном складе
    // и отсутствие строки с тем же штрихкодом на складе точки.
    if (product.hasCentralRow && !product.hasStoreRow) {
      results.push({
        name: product.name,
        barcode: product.barcode,
        stock: product.centralStock
      });
    }
  }

  results.sort((a, b) =>
    a.name.localeCompare(b.name, "ru", { sensitivity: "base" })
  );

  renderResults();

  statusEl.className = "status";
  statusEl.textContent =
    `Анализ завершён. Проверено строк складов: ${matchedRows}. ` +
    `Пропущено неполных строк: ${skippedRows}.`;
}

function renderResults() {
  resultsBody.replaceChildren();

  for (const product of results) {
    const tr = document.createElement("tr");

    addCell(tr, product.name);
    addCell(tr, product.barcode);
    addCell(tr, formatStock(product.stock));

    resultsBody.appendChild(tr);
  }

  resultCount.textContent = `Найдено: ${results.length}`;
  emptyMessage.hidden = results.length !== 0;
  resultsSection.hidden = false;
  exportBtn.disabled = results.length === 0;
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
  if (!results.length) return;

  const data = [
    ["Товар", "Штрихкод", "Остаток на центральном складе"],
    ...results.map(product => [
      product.name,
      product.barcode,
      product.stock
    ])
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(data);
  worksheet["!cols"] = [
    { wch: 40 },
    { wch: 24 },
    { wch: 30 }
  ];

  const exportWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    exportWorkbook,
    worksheet,
    "Только центральный склад"
  );

  XLSX.writeFile(exportWorkbook, "Товары_только_на_центральном.xlsx");
}

function setError(message) {
  statusEl.textContent = message;
  statusEl.className = "status error";
}
