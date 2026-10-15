const CENTRAL_WAREHOUSE = 'Центральный склад'
const STORE_WAREHOUSE = 'Склад точки Аккорд'
const REPAIR_WAREHOUSE = 'Склад точки На ремонт'

// Нумерация колонок начинается с нуля:
// A = 0, B = 1, C = 2 и так далее
const COLUMNS = {
    warehouse: 1, // B — склад
    name: 2, // C — название товара
    barcode: 3, // D — штрихкод
    salePrice: 4, // E — цена продажи
    purchasePrice: 5, // F — цена закупки
    stock: 6, // G — количество
    supplier: 9, // J — поставщик
}

const RESTOCK_EXCLUSIONS = [
    // Исключения «Нужно пополнить»
    'Акустическая система',
    'Гитара',
    'Домбра',
    'Комбоусилитель',
    'Комплект',
    'Коробка',
    'Микрофон',
    'MIDI-контроллер',
    'Синтезатор',
    'Скрипка',
    'Усилитель',
    'Ударная установка',
    'Укулеле',
    'Флейта',
    'Электрогитара',
    'Электроскрипка',
    'Звуковая карта',
    '6358',
    'Аудио интерфейс',
    'Кларнет',
    'Банкетка Rowell',
]

const RESTOCK_ALLOWLIST = [
    // Разрешено «Нужно пополнить»
    'Домбра 48р',
    'Домбра 46 простая',
    'Кнопка домбра',
    'Домбра 44',
    'береза лиственница бук',
]

const CENTRAL_EXCLUSIONS = [
    // Исключения «Только на центральном»
    'Электрогитара G110',
    'Чехол простой 39',
    'на комплекты',
    'для комплектов',
    'Fender 097-0113-022',
    'Fender 097-4220-500',
    'Гитара акустическая 6454',
    'Блок питания 9.5 В 1 А',
    'Струны цветные',
]

const fileInput = document.getElementById('excelFile')
const fileName = document.getElementById('fileName')
const exportBtn = document.getElementById('exportBtn')
const statusEl = document.getElementById('status')
const resultsSection = document.getElementById('results')
const resultsBody = document.getElementById('resultsBody')
const resultsHeadings = document.getElementById('resultsHeadings')
const emptyMessage = document.getElementById('emptyMessage')
const resultsTitle = document.getElementById('resultsTitle')
const exceptionsAllowedInfo = document.getElementById('exceptionsAllowedInfo')
const centralTab = document.getElementById('centralTab')
const restockTab = document.getElementById('restockTab')
const allStockTab = document.getElementById('allStockTab')
const transferTab = document.getElementById('transferTab')
const exceptionsLink = document.getElementById('exceptionsLink')
const restockBackLink = document.getElementById('restockBackLink')
const resultsPanel = document.getElementById('resultsPanel')
const transferPanel = document.getElementById('transferPanel')
const transferBody = document.getElementById('transferBody')
const transferEmptyMessage = document.getElementById('transferEmptyMessage')
const transferExportBtn = document.getElementById('transferExportBtn')
const transferPrintBtn = document.getElementById('transferPrintBtn')
const transferExcludedLink = document.getElementById('transferExcludedLink')
const transferBackLink = document.getElementById('transferBackLink')
const exceptionsInfo = document.getElementById('exceptionsInfo')
const exceptionsList = document.getElementById('exceptionsList')
const exceptionsAllowed = document.getElementById('exceptionsAllowed')

let workbook = null
let results = []
let excludedCentralResults = []
let restockResults = []
let excludedRestockResults = []
let allStockResults = []
let activeList = 'central'
let transferView = 'items'
let selectedFileName = ''

fileInput.addEventListener('change', async () => {
    const file = fileInput.files[0]

    workbook = null
    results = []
    excludedCentralResults = []
    restockResults = []
    excludedRestockResults = []
    allStockResults = []
    activeList = 'central'
    transferView = 'items'
    centralTab.classList.add('is-active')
    centralTab.setAttribute('aria-selected', 'true')
    restockTab.classList.remove('is-active')
    restockTab.setAttribute('aria-selected', 'false')
    allStockTab.classList.remove('is-active')
    allStockTab.setAttribute('aria-selected', 'false')
    transferTab.classList.remove('is-active')
    transferTab.setAttribute('aria-selected', 'false')
    resultsPanel.setAttribute('aria-labelledby', 'centralTab')
    exportBtn.disabled = true
    resultsSection.hidden = true
    document.body.classList.remove('has-results')
    statusEl.textContent = ''
    statusEl.className = 'status'

    if (!file) {
        selectedFileName = ''
        fileName.textContent = 'Выберите Excel-файл'
        return
    }

    selectedFileName = file.name
    fileName.textContent = file.name
    statusEl.textContent = 'Читаем файл…'

    try {
        const buffer = await file.arrayBuffer()

        workbook = XLSX.read(buffer, {
            type: 'array',
            cellText: true,
            cellDates: false,
        })

        if (!workbook.SheetNames.length) {
            throw new Error('В файле нет листов.')
        }

        analyzeStock()
    } catch (error) {
        workbook = null
        statusEl.textContent = `Не удалось прочитать файл: ${error.message}`
        statusEl.className = 'status error'
    }
})

exportBtn.addEventListener('click', exportResults)
transferExportBtn.addEventListener('click', exportTransferResults)
transferPrintBtn.addEventListener('click', printTransferResults)
centralTab.addEventListener('click', () => selectResultsTab('central'))
restockTab.addEventListener('click', () => selectResultsTab('restock'))
allStockTab.addEventListener('click', () => selectResultsTab('all-stock'))
transferTab.addEventListener('click', () => {
    transferView = 'items'
    selectResultsTab('transfer')
})
transferExcludedLink.addEventListener('click', () => {
    transferView = 'excluded'
    renderResults()
})
transferBackLink.addEventListener('click', () => {
    transferView = 'items'
    renderResults()
})
exceptionsLink.addEventListener('click', () =>
    selectRestockView(activeList === 'central' ? 'central-exceptions' : 'exceptions'),
)
restockBackLink.addEventListener('click', () =>
    selectRestockView(activeList === 'central-exceptions' ? 'central' : 'restock'),
)
exceptionsList.textContent = RESTOCK_EXCLUSIONS.join(', ')
exceptionsAllowed.textContent = RESTOCK_ALLOWLIST.join(', ')

function normalizeText(value) {
    return String(value ?? '').trim()
}

function getProductKey(product) {
    return JSON.stringify([normalizeBarcode(product?.barcode), normalizeText(product?.name)])
}

function syncTransferQuantityAcrossLists(productKey, quantity) {
    const normalized = String(quantity ?? '')

    for (const list of [results, excludedCentralResults, restockResults, excludedRestockResults, allStockResults]) {
        for (const item of list) {
            if (getProductKey(item) === productKey) {
                item.transferQuantity = normalized
            }
        }
    }
}

function normalizeBarcode(value) {
    if (value === null || value === undefined) return ''

    // Штрихкод не должен зависеть от пробелов по краям.
    // Ведущие нули сохраняются, если Excel хранит их в ячейке как текст.
    return String(value).trim()
}

function parseStock(value) {
    if (typeof value === 'number') {
        return Number.isFinite(value) ? value : null
    }

    const normalized = String(value ?? '')
        .trim()
        .replace(/\s/g, '')
        .replace(',', '.')

    if (!normalized) return null

    const number = Number(normalized)
    return Number.isFinite(number) ? number : null
}

function isZeroPrice(value) {
    if (typeof value === 'number') {
        return Number.isFinite(value) && value === 0
    }

    const normalized = String(value ?? '')
        .trim()
        .replace(/[\s\u00a0]/g, '')
        .replace(/[₽$€£]/g, '')
        .replace(/(?:руб\.?|р\.?)/giu, '')
        .replace(',', '.')

    if (!normalized) return false

    const price = Number(normalized)
    return Number.isFinite(price) && price === 0
}

function hasRestockExclusion(name) {
    const normalizedName = normalizeText(name).toLocaleLowerCase('ru')
    return RESTOCK_EXCLUSIONS.some(exclusion => normalizedName.includes(exclusion.toLocaleLowerCase('ru')))
}

function isRestockAllowed(name) {
    const normalizedName = normalizeText(name).toLocaleLowerCase('ru')
    return RESTOCK_ALLOWLIST.some(allowedName => normalizedName === allowedName.toLocaleLowerCase('ru'))
}

function hasCentralExclusion(name) {
    const normalizedName = normalizeText(name).toLocaleLowerCase('ru')
    return CENTRAL_EXCLUSIONS.some(exclusion => normalizedName.includes(exclusion.toLocaleLowerCase('ru')))
}

function analyzeStock() {
    if (!workbook) {
        setError('Сначала выберите Excel-файл.')
        return
    }

    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]

    const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: '',
        raw: false,
    })

    const products = new Map()
    const allProducts = new Map()
    let matchedRows = 0
    let skippedRows = 0

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i]

        const warehouse = normalizeText(row[COLUMNS.warehouse])
        const barcode = normalizeBarcode(row[COLUMNS.barcode])
        const name = normalizeText(row[COLUMNS.name])
        const salePrice = normalizeText(row[COLUMNS.salePrice])
        const purchasePrice = normalizeText(row[COLUMNS.purchasePrice])
        const supplier = normalizeText(row[COLUMNS.supplier])
        const stock = parseStock(row[COLUMNS.stock])
        const normalizedName = name.toLocaleLowerCase('ru')
        const isNameHeader = [
            'название',
            'название товара',
            'наименование',
            'наименование товара',
            'номенклатура',
            'товар',
        ].includes(normalizedName)

        if (name && !isNameHeader) {
            const allProductKey = JSON.stringify([barcode, name])
            if (!allProducts.has(allProductKey)) {
                allProducts.set(allProductKey, {
                    name,
                    barcode,
                    storeStock: 0,
                    centralStock: 0,
                    repairStock: 0,
                    salePrice,
                    purchasePrice,
                    supplier,
                })
            }

            const allProduct = allProducts.get(allProductKey)
            if (!allProduct.salePrice && salePrice) allProduct.salePrice = salePrice
            if (!allProduct.purchasePrice && purchasePrice) allProduct.purchasePrice = purchasePrice
            if (!allProduct.supplier && supplier) allProduct.supplier = supplier
            if (stock !== null) {
                if (warehouse === STORE_WAREHOUSE) allProduct.storeStock += stock
                if (warehouse === CENTRAL_WAREHOUSE) allProduct.centralStock += stock
                if (warehouse === REPAIR_WAREHOUSE) allProduct.repairStock += stock
            }
        }

        // Считаем строки всех складов, участвующих в сводке остатков.
        if (warehouse !== CENTRAL_WAREHOUSE && warehouse !== STORE_WAREHOUSE && warehouse !== REPAIR_WAREHOUSE) {
            continue
        }

        // Пропускаем заголовки и строки без штрихкода.
        if (!barcode || !name || stock === null) {
            skippedRows++
            continue
        }

        matchedRows++
        // Склад на ремонт участвует в общей сводке, но не в списках перемещения и пополнения.
        if (warehouse === REPAIR_WAREHOUSE) continue

        const price = salePrice

        const productKey = JSON.stringify([barcode, name])

        if (!products.has(productKey)) {
            products.set(productKey, {
                barcode,
                name,
                price,
                centralStock: 0,
                storeStock: 0,
                hasCentralRow: false,
                hasStoreRow: false,
            })
        }

        const product = products.get(productKey)

        if (!product.name && name) {
            product.name = name
        }
        if (!product.price && price) {
            product.price = price
        }

        if (warehouse === CENTRAL_WAREHOUSE) {
            product.hasCentralRow = true
            product.centralStock += stock
        }

        if (warehouse === STORE_WAREHOUSE) {
            product.hasStoreRow = true
            product.storeStock += stock
        }
    }

    results = []
    excludedCentralResults = []
    restockResults = []
    excludedRestockResults = []
    allStockResults = Array.from(allProducts.values()).map(product => ({
        ...product,
        price: product.salePrice ?? product.purchasePrice ?? '',
        transferQuantity: '',
    }))

    for (const product of products.values()) {
        // Требуется именно наличие строки на центральном складе
        // и отсутствие строки с тем же штрихкодом на складе точки.
        if (product.hasCentralRow && !product.hasStoreRow && !isZeroPrice(product.price)) {
            const centralProduct = {
                name: product.name,
                barcode: product.barcode,
                stock: product.centralStock,
                price: product.price,
                transferQuantity: '1',
            }

            if (hasCentralExclusion(product.name)) {
                excludedCentralResults.push(centralProduct)
            } else {
                results.push(centralProduct)
            }
        }

        if (product.storeStock > 0 && product.storeStock < 3 && product.centralStock > 0) {
            const restockProduct = {
                name: product.name,
                barcode: product.barcode,
                storeStock: product.storeStock,
                centralStock: product.centralStock,
                price: product.price,
                transferQuantity: String(Math.min(product.centralStock, 3 - product.storeStock)),
            }

            if (hasRestockExclusion(product.name) && !isRestockAllowed(product.name)) {
                excludedRestockResults.push(restockProduct)
            } else {
                restockResults.push(restockProduct)
            }
        }
    }

    const sortByName = (a, b) => a.name.localeCompare(b.name, 'ru', { sensitivity: 'base' })
    results.sort(sortByName)
    excludedCentralResults.sort(sortByName)
    restockResults.sort(sortByName)
    excludedRestockResults.sort(sortByName)
    allStockResults.sort(sortByName)

    for (const product of [...results, ...restockResults]) {
        const matchedAllStock = allStockResults.find(item => getProductKey(item) === getProductKey(product))
        if (matchedAllStock) {
            matchedAllStock.transferQuantity = product.transferQuantity
            syncTransferQuantityAcrossLists(getProductKey(product), product.transferQuantity)
        }
    }

    for (const product of results) {
        syncTransferQuantityAcrossLists(getProductKey(product), product.transferQuantity)
    }

    renderResults()

    statusEl.className = 'status'
    statusEl.textContent = `Проверено: ${matchedRows}, пропущено: ${skippedRows}`
}

function renderResults() {
    if (activeList === 'transfer') {
        const transferItems = getTransferItems()
        const excludedTransferItems = getNotInTransferItems()
        resultsPanel.hidden = true
        transferPanel.hidden = false
        transferTab.textContent = `Перемещение ${transferItems.length}`
        transferExcludedLink.hidden = transferView === 'excluded' || excludedTransferItems.length === 0
        transferExcludedLink.textContent = `Не вошло в перемещение · ${excludedTransferItems.length}`
        transferBackLink.hidden = transferView !== 'excluded'
        transferTab.classList.add('is-active')
        transferTab.setAttribute('aria-selected', 'true')
        centralTab.classList.remove('is-active')
        centralTab.setAttribute('aria-selected', 'false')
        restockTab.classList.remove('is-active')
        restockTab.setAttribute('aria-selected', 'false')
        allStockTab.classList.remove('is-active')
        allStockTab.setAttribute('aria-selected', 'false')
        resultsPanel.setAttribute('aria-labelledby', 'centralTab')
        renderTransferTable(
            transferView === 'excluded' ? excludedTransferItems : transferItems,
            transferView === 'excluded',
        )
        fileName.textContent = selectedFileName
        return
    }

    resultsPanel.hidden = false
    transferPanel.hidden = true
    updateTransferCount()
    transferTab.classList.remove('is-active')
    transferTab.setAttribute('aria-selected', 'false')

    const isAllStockList = activeList === 'all-stock'
    const isCentralList = activeList === 'central'
    const isCentralExceptionsList = activeList === 'central-exceptions'
    const isRestockList = activeList === 'restock' || activeList === 'exceptions'
    const isMoveEditable = isCentralList || activeList === 'restock' || isAllStockList
    const isExceptionsList = isCentralExceptionsList || activeList === 'exceptions'
    const activeResults = isAllStockList
        ? allStockResults
        : isCentralList
          ? results
          : isCentralExceptionsList
            ? excludedCentralResults
            : activeList === 'exceptions'
              ? excludedRestockResults
              : restockResults

    exceptionsLink.hidden = isExceptionsList || isAllStockList
    exceptionsLink.textContent = isRestockList
        ? `Исключения · ${excludedRestockResults.length}`
        : `Исключения · ${excludedCentralResults.length}`
    restockBackLink.hidden = !isExceptionsList
    restockBackLink.textContent = isCentralExceptionsList
        ? '← К товарам на центральном складе'
        : '← К товарам к пополнению'
    exceptionsInfo.hidden = !isExceptionsList
    exceptionsList.textContent = isCentralExceptionsList ? CENTRAL_EXCLUSIONS.join(', ') : RESTOCK_EXCLUSIONS.join(', ')
    exceptionsAllowedInfo.hidden = isCentralExceptionsList
    resultsBody.replaceChildren()
    resultsHeadings.replaceChildren()

    const headings = isAllStockList
        ? [
              'Название',
              'Ак',
              'ЦС',
              'Рем',
              'Продажа',
              'Закупка',
              'Штрих код',
              'Поставщик',
              ...(isMoveEditable ? ['Перемещение'] : []),
          ]
        : isRestockList
          ? ['Товар', 'Штрихкод', 'Аккорд', 'ЦС', 'Цена', ...(isMoveEditable ? ['Перемещение'] : [])]
          : ['Товар', 'Штрихкод', 'ЦС', 'Цена', ...(isMoveEditable ? ['Перемещение'] : [])]
    for (const heading of headings) {
        const th = document.createElement('th')
        th.textContent = heading
        if (['Аккорд', 'ЦС'].includes(heading)) {
            th.classList.add('count-column')
        } else if (heading === 'Перемещение') {
            th.classList.add('same-width-column')
        }
        resultsHeadings.appendChild(th)
    }

    for (const product of activeResults) {
        const tr = document.createElement('tr')

        if (isAllStockList) {
            addCell(tr, product.name)
            addCell(tr, formatStock(product.storeStock))
            addCell(tr, formatStock(product.centralStock))
            addCell(tr, formatStock(product.repairStock))
            addCell(tr, formatPrice(product.salePrice))
            addCell(tr, formatPrice(product.purchasePrice))
            addCell(tr, product.barcode)
            addCell(tr, product.supplier)
            if (isMoveEditable) {
                addTransferQuantityCell(tr, product)
            }
        } else {
            addCell(tr, product.name)
            addCell(tr, product.barcode)
            addCell(tr, formatStock(isRestockList ? product.storeStock : product.stock))

            if (isRestockList) {
                addCell(tr, formatStock(product.centralStock))
            }
            addCell(tr, formatPrice(product.price))
            if (isMoveEditable) {
                addTransferQuantityCell(tr, product)
            }
        }

        for (const [index, heading] of headings.entries()) {
            if (['Аккорд', 'ЦС'].includes(heading)) {
                tr.cells[index].classList.add('count-column')
            } else if (heading === 'Перемещение') {
                tr.cells[index].classList.add('same-width-column')
            }
        }

        resultsBody.appendChild(tr)
    }

    resultsTitle.textContent = isAllStockList
        ? 'Все остатки'
        : isCentralList
          ? 'Товары только на центральном складе'
          : isCentralExceptionsList
            ? 'Товары, исключённые из списка центрального склада'
            : activeList === 'exceptions'
              ? 'Товары, исключённые из пополнения'
              : 'Товары, которые нужно пополнить'
    centralTab.textContent = `Только на центральном ${results.length}`
    restockTab.textContent = `Нужно пополнить ${restockResults.length}`
    allStockTab.textContent = `Все остатки ${allStockResults.length}`
    const isCentralActive = isCentralList || isCentralExceptionsList
    const isRestockActive = activeList === 'restock' || activeList === 'exceptions'
    centralTab.classList.toggle('is-active', isCentralActive)
    centralTab.setAttribute('aria-selected', String(isCentralActive))
    restockTab.classList.toggle('is-active', isRestockActive)
    restockTab.setAttribute('aria-selected', String(isRestockActive))
    allStockTab.classList.toggle('is-active', isAllStockList)
    allStockTab.setAttribute('aria-selected', String(isAllStockList))
    resultsPanel.setAttribute(
        'aria-labelledby',
        isAllStockList ? 'allStockTab' : isRestockActive ? 'restockTab' : 'centralTab',
    )
    fileName.textContent = selectedFileName
    emptyMessage.textContent = isAllStockList
        ? 'Товары в файле не найдены.'
        : isCentralList
          ? 'Товаров, которые есть только на центральном складе, не найдено.'
          : isCentralExceptionsList
            ? 'Исключённых товаров на центральном складе не найдено.'
            : activeList === 'exceptions'
              ? 'Исключённых товаров для пополнения не найдено.'
              : 'Товаров для пополнения не найдено.'
    emptyMessage.hidden = activeResults.length !== 0
    resultsSection.hidden = false
    document.body.classList.add('has-results')
    exportBtn.disabled = activeResults.length === 0
}

function getTransferItems() {
    const seen = new Set()
    return [...results, ...restockResults, ...allStockResults]
        .filter(product => {
            const key = getProductKey(product)
            if (seen.has(key)) return false
            seen.add(key)
            return parseStock(product.transferQuantity) > 0
        })
        .sort((a, b) => a.name.localeCompare(b.name, 'ru', { sensitivity: 'base' }))
}

function getNotInTransferItems() {
    const seen = new Set()
    return [...results, ...restockResults]
        .filter(product => {
            const key = getProductKey(product)
            if (seen.has(key)) return false
            seen.add(key)
            return parseStock(product.transferQuantity) <= 0
        })
        .sort((a, b) => a.name.localeCompare(b.name, 'ru', { sensitivity: 'base' }))
}

function updateTransferCount() {
    transferTab.textContent = `Перемещение ${getTransferItems().length}`
}

function renderTransferTable(products, isExcludedList = false) {
    transferBody.replaceChildren()
    for (const product of products) {
        const row = document.createElement('tr')
        addCell(row, product.name)
        addCell(row, formatPrice(product.price ?? product.salePrice ?? product.purchasePrice ?? ''))
        if (isExcludedList) {
            addCell(row, '')
        } else {
            addTransferQuantityCell(row, product, true)
        }
        transferBody.appendChild(row)
    }
    transferEmptyMessage.textContent = isExcludedList
        ? 'Товаров, не вошедших в перемещение, не найдено.'
        : 'В таблице на перемещение пока нет товаров.'
    transferEmptyMessage.hidden = products.length !== 0
    transferExportBtn.disabled = products.length === 0
    transferPrintBtn.disabled = products.length === 0
}

function addTransferQuantityCell(row, product, isTransferTable = false) {
    const cell = document.createElement('td')
    cell.className = 'transfer-quantity-cell same-width-column'
    const controls = document.createElement('div')
    controls.className = 'transfer-quantity-controls'
    const quantityInput = document.createElement('input')
    quantityInput.className = 'transfer-quantity'
    quantityInput.type = 'number'
    quantityInput.min = '0'
    quantityInput.step = 'any'
    const maxQuantity = Math.max(0, parseStock(product.centralStock ?? product.stock ?? product.storeStock) ?? 0)
    const currentQuantity = parseStock(product.transferQuantity) ?? 0
    const initialQuantity = Math.max(0, currentQuantity)
    quantityInput.max = String(maxQuantity)
    quantityInput.value = String(initialQuantity)
    product.transferQuantity = quantityInput.value
    quantityInput.setAttribute('aria-label', `Количество перемещения: ${product.name}`)

    const updateQuantity = () => {
        const enteredQuantity = parseStock(quantityInput.value)
        if (enteredQuantity !== null) {
            quantityInput.value = String(Math.min(maxQuantity, Math.max(0, enteredQuantity)))
        }
        product.transferQuantity = quantityInput.value
        syncTransferQuantityAcrossLists(getProductKey(product), product.transferQuantity)
        updateTransferCount()
        if (activeList === 'transfer') {
            const hasTransferItems = getTransferItems().length > 0
            transferExportBtn.disabled = !hasTransferItems
            transferPrintBtn.disabled = !hasTransferItems
        }
    }

    quantityInput.addEventListener('input', updateQuantity)
    if (isTransferTable) {
        quantityInput.addEventListener('blur', () => {
            if (parseStock(product.transferQuantity) <= 0) {
                renderResults()
            }
        })
    }

    const decrementButton = document.createElement('button')
    decrementButton.type = 'button'
    decrementButton.className = 'transfer-quantity-button'
    decrementButton.textContent = '−'
    decrementButton.setAttribute('aria-label', `Уменьшить количество для ${product.name}`)

    const incrementButton = document.createElement('button')
    incrementButton.type = 'button'
    incrementButton.className = 'transfer-quantity-button'
    incrementButton.textContent = '+'
    incrementButton.setAttribute('aria-label', `Увеличить количество для ${product.name}`)

    const adjustQuantity = delta => {
        const currentQuantity = parseStock(quantityInput.value) ?? 0
        const nextQuantity = currentQuantity + delta
        if (nextQuantity < 0 || nextQuantity > maxQuantity) return

        quantityInput.value = String(nextQuantity)
        updateQuantity()
        if (isTransferTable && parseStock(product.transferQuantity) <= 0) {
            renderResults()
        }
    }

    decrementButton.addEventListener('click', () => adjustQuantity(-1))
    incrementButton.addEventListener('click', () => adjustQuantity(1))

    controls.append(quantityInput, decrementButton, incrementButton)
    cell.appendChild(controls)
    row.appendChild(cell)
}

function selectResultsTab(tab) {
    activeList = tab
    const isCentralList = tab === 'central'
    const isRestockList = tab === 'restock'
    const isAllStockList = tab === 'all-stock'

    centralTab.classList.toggle('is-active', isCentralList)
    centralTab.setAttribute('aria-selected', String(isCentralList))
    restockTab.classList.toggle('is-active', isRestockList)
    restockTab.setAttribute('aria-selected', String(isRestockList))
    allStockTab.classList.toggle('is-active', isAllStockList)
    allStockTab.setAttribute('aria-selected', String(isAllStockList))
    resultsPanel.setAttribute(
        'aria-labelledby',
        isCentralList ? 'centralTab' : isRestockList ? 'restockTab' : 'allStockTab',
    )

    renderResults()
}

function selectRestockView(view) {
    activeList = view
    renderResults()
}

function addCell(row, value) {
    const td = document.createElement('td')
    td.textContent = String(value ?? '')
    row.appendChild(td)
}

function formatStock(value) {
    return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(3)))
}

function formatPrice(value) {
    return String(value ?? '').replace(/[,.]00(?=\s*(?:₽|руб\.?|р\.?)?\s*$)/iu, '')
}

function formatExportDate() {
    return new Intl.DateTimeFormat('ru-RU', {
        day: '2-digit',
        month: '2-digit',
        year: '2-digit',
    }).format(new Date())
}

function getTransferExportData() {
    const transferItems = getTransferItems()
    return [
        ['Название', 'Цена ЦС', 'Кол-во', 'Цена Ак'],
        ...transferItems.map(product => [
            product.name,
            formatPrice(product.price),
            Number(product.transferQuantity),
            '',
        ]),
    ]
}

function exportTransferResults() {
    const transferItems = getTransferItems()
    if (!transferItems.length) return

    const worksheet = XLSX.utils.aoa_to_sheet(getTransferExportData())
    worksheet['!cols'] = [{ wch: 40 }, { wch: 12 }, { wch: 10 }, { wch: 12 }]

    const exportWorkbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(exportWorkbook, worksheet, 'Накладная на перемещение')
    XLSX.writeFile(exportWorkbook, `Накладная на перемещение ${formatExportDate()}.xlsx`)
}

function printTransferResults() {
    const transferItems = getTransferItems()
    if (!transferItems.length) return

    const printWindow = window.open('', '_blank')
    if (!printWindow) {
        setError('Не удалось открыть окно печати. Разрешите всплывающие окна и попробуйте снова.')
        return
    }

    const totalQuantity = transferItems.reduce((sum, product) => sum + (parseStock(product.transferQuantity) ?? 0), 0)
    const totalAmount = transferItems.reduce(
        (sum, product) => sum + (parseStock(product.transferQuantity) ?? 0) * (parseStock(product.price) ?? 0),
        0,
    )

    const itemRows = transferItems
        .map((product, index) => {
            const quantity = parseStock(product.transferQuantity) ?? 0
            const barcodeValue = String(product.barcode ?? '').trim()
            const barcodeId = `barcode-${index}`
            const barcodeFormat =
                barcodeValue && /^\d+$/.test(barcodeValue)
                    ? barcodeValue.length === 13
                        ? 'EAN13'
                        : 'CODE128'
                    : 'CODE128'
            const barcodeMarkup = barcodeValue
                ? `<svg id="${barcodeId}" class="print-barcode" data-barcode-code="${escapeHtml(barcodeValue)}" aria-label="Штрихкод товара ${escapeHtml(product.name)}"></svg>`
                : '<div class="print-barcode empty-barcode"></div>'

            return `
              <tr>
                <td class="product-name-cell">
                  <div class="print-product-name-wrapper">
                    <span class="print-product-name">${escapeHtml(product.name)}</span>
                    <span class="print-product-barcode">${barcodeMarkup}</span>
                  </div>
                </td>
                <td class="price-cell">${escapeHtml(formatPrice(product.price))}</td>
                <td class="qty-cell">${escapeHtml(String(quantity))}</td>
                <td class="empty-price-cell">&nbsp;</td>
              </tr>
            `
        })
        .join('')

    const blankRows = Array.from(
        { length: 12 },
        () => `
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    `,
    ).join('')

    const printDate = new Intl.DateTimeFormat('ru-RU', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
    }).format(new Date())

    printWindow.document.write(`<!DOCTYPE html>
    <html lang="ru">
      <head>
        <meta charset="UTF-8">
        <title>Накладная на перемещение ${printDate}</title>
        <style>
          @page { margin: 0; size: A4 portrait; }
          body {
            margin: 0;
            padding: 16mm;
            font-family: Arial, Helvetica, sans-serif;
            color: #202124;
          }
          .print-header {
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            margin-bottom: 14px;
          }
          h1 {
            margin: 0;
            font-size: 18px;
            font-weight: 700;
          }
          .print-date {
            font-size: 14px;
            font-weight: 400;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            text-align: left;
          }
          th:nth-child(1), td:nth-child(1) { width: 78%; }
          th:nth-child(2), td:nth-child(2) { width: 8%; }
          th:nth-child(3), td:nth-child(3) { width: 4%; }
          th:nth-child(4), td:nth-child(4) { width: 10%; }
          th, td {
            padding: 5px 7px;
            border: 1px solid #999;
            font-size: 12px;
            height: 18px;
            vertical-align: middle;
          }
          th {
            background: #f2f2f2;
            font-weight: 700;
          }
          .summary-row td {
            font-weight: 700;
            background: #fafafa;
          }
          .product-name-cell {
            padding: 0;
            border-left: 1px solid #999;
          }
          .print-product-name-wrapper {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 8px;
            min-height: 28px;
            padding: 4px 6px;
          }
          .print-product-name {
            font-weight: 400;
            line-height: 1.3;
            flex: 1 1 auto;
            white-space: normal;
            min-width: 0;
          }
          .print-product-barcode {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            flex: 0 0 auto;
            min-width: 80px;
            width: 80px;
            padding: 0;
            border: none;
            background: transparent;
          }
          .print-barcode {
            display: block;
            width: 100%;
            max-width: 80px;
            height: 22px;
            background: #fff;
          }
          .empty-barcode {
            width: 80px;
            height: 22px;
            border: none;
          }
          .price-cell,
          .qty-cell,
          .empty-price-cell {
            text-align: left;
          }
          @media print {
            body { margin: 0; padding: 16mm; }
            th { background: #f2f2f2 !important; print-color-adjust: exact; }
          }
        </style>
      </head>
      <body>
        <header class="print-header">
          <h1>Накладная на перемещение</h1>
          <span class="print-date">${printDate}</span>
        </header>
        <table>
          <thead>
            <tr>
              <th>Название</th>
              <th>Цена ЦС</th>
              <th>Кол.</th>
              <th>Цена Ак</th>
            </tr>
          </thead>
          <tbody>
            ${itemRows}
            <tr class="summary-row">
              <td>Итого</td>
              <td>${escapeHtml(formatPrice(totalAmount))}</td>
              <td>${escapeHtml(String(totalQuantity))}</td>
              <td>&nbsp;</td>
            </tr>
            ${blankRows}
          </tbody>
        </table>
        <script src="https://cdn.jsdelivr.net/npm/jsbarcode@3.11.6/dist/JsBarcode.all.min.js"></script>
        <script>
          window.onload = () => {
            const barcodes = Array.from(document.querySelectorAll('.print-barcode'));
            barcodes.forEach(svg => {
              const code = svg.dataset.barcodeCode || '';
              if (code) {
                const format = /^\d+$/.test(code) ? (code.length === 13 ? 'EAN13' : 'CODE128') : 'CODE128';
                JsBarcode(svg, code, {
                  format,
                  displayValue: false,
                  width: 1.2,
                  height: 32,
                  margin: 0,
                  background: '#fff',
                  lineColor: '#000'
                });
              }
            });
            window.focus();
            window.print();
          };
        </script>
      </body>
    </html>`)
    printWindow.document.close()
}

function escapeHtml(value) {
    return String(value).replace(
        /[&<>"']/g,
        character =>
            ({
                '&': '&amp;',
                '<': '&lt;',
                '>': '&gt;',
                '"': '&quot;',
                "'": '&#39;',
            })[character],
    )
}

function exportResults() {
    const isAllStockList = activeList === 'all-stock'
    const isCentralList = activeList === 'central' || activeList === 'central-exceptions'
    const isCentralExceptionsList = activeList === 'central-exceptions'
    const isRestockList = activeList === 'restock' || activeList === 'exceptions'
    const isExceptionsList = activeList === 'exceptions'
    const activeResults = isAllStockList
        ? allStockResults
        : isCentralList
          ? isCentralExceptionsList
              ? excludedCentralResults
              : results
          : isExceptionsList
            ? excludedRestockResults
            : restockResults
    if (!activeResults.length) return

    const data = isAllStockList
        ? [
              ['Название', 'Аккорд', 'ЦС', 'Ремонт', 'Продажа', 'Закупка', 'Штрих код', 'Поставщик'],
              ...activeResults.map(product => [
                  product.name,
                  formatStock(product.storeStock),
                  formatStock(product.centralStock),
                  formatStock(product.repairStock),
                  formatPrice(product.salePrice),
                  formatPrice(product.purchasePrice),
                  product.barcode,
                  product.supplier,
              ]),
          ]
        : [
              isRestockList ? ['Товар', 'Штрихкод', 'Аккорд', 'ЦС', 'Цена'] : ['Товар', 'Штрихкод', 'ЦС', 'Цена'],
              ...activeResults.map(product =>
                  isRestockList
                      ? [
                            product.name,
                            product.barcode,
                            product.storeStock,
                            product.centralStock,
                            formatPrice(product.price),
                        ]
                      : [product.name, product.barcode, product.stock, formatPrice(product.price)],
              ),
          ]

    const worksheet = XLSX.utils.aoa_to_sheet(data)
    worksheet['!cols'] = isAllStockList
        ? [{ wch: 40 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 24 }, { wch: 30 }]
        : [{ wch: 40 }, { wch: 24 }, { wch: 24 }, ...(isRestockList ? [{ wch: 30 }] : []), { wch: 16 }]

    const exportWorkbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(
        exportWorkbook,
        worksheet,
        isAllStockList
            ? 'Все остатки'
            : isCentralList
              ? isCentralExceptionsList
                  ? 'Исключения центрального склада'
                  : 'Только центральный склад'
              : isExceptionsList
                ? 'Исключения для пополнения'
                : 'Нужно пополнить',
    )

    const exportName = isAllStockList
        ? 'Все_остатки'
        : isCentralList
          ? isCentralExceptionsList
              ? 'Исключения_центрального_склада'
              : 'Товары_только_на_центральном'
          : isExceptionsList
            ? 'Исключения_для_пополнения'
            : 'Нужно_пополнить'

    XLSX.writeFile(exportWorkbook, `${exportName} ${formatExportDate()}.xlsx`)
}

function setError(message) {
    statusEl.textContent = message
    statusEl.className = 'status error'
}
