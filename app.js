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
    'Чехол для бас гитары черный B12-EG-BK',
    'Педаль сустейна',
    'Саксофон',
    'Банкетка',
    'Цифровое пианино',
]

const RESTOCK_ALLOWLIST = [
    // Разрешено «Нужно пополнить»
    'Домбра 48р',
    'Домбра 46 простая',
    'Кнопка домбра',
    'Домбра 44',
    'береза лиственница бук',
    'Футляр домбра тёплый мягкий',
    'Футляр домбра тёплый',
]

const RESTOCK_TO_TEN_ITEMS = [
    // Пополняются до 10, если в магазине меньше 5
    'Чехол для домбры',
    'Чехол домбра детский',
    'Увлажнитель китай 6359',
    'Тюнер китай 6357 арома',
    'Каподастр',
    'Каподастр Китай 6411',
    'Каподастр металл',
    'Чехол простой 38',
    'Чехол простой 41',
    'Чехол теплый 41 ВК',
    // 'Чехол 41 толстый',
    'Чехол утепленный Deviser 39',
    'Чехол утепленный Deviser 41',
    'Чехол утеплённый RockStel Delux DGB110A black',
    'Чехол утеплённый RockStel Element EGB110A Gray',
    'Чехол утеплённый RockStel Element EGB110C RED',    
    'Чехол утеплённый RockStel Element EGB110EB BK',
    // 'Ремень Fender',
    // 'Ремень Fender цветной',
    // 'Ремень Rizo',
]
const RESTOCK_TO_TEN_MINIMUM = 5
const RESTOCK_TO_TEN_TARGET = 10

const STOP_ITEMS = [
    // На стопе «Только на центральном»
    'Комплект Ronnie wood с вырезом Bls',
    'Кабель Standard Cable 3m STR/ANG',
    'Гитара Adagio 7000 CE с подключением',
    'Тюнер Китай 7444 Арома',
    'Чехол для бас гитары черный B12-EG-BK',
    'Электрогитара G110 OPBC',
    'Электрогитара G110 OPBK',
]

const CENTRAL_EXCLUSIONS = [
    // Исключения «Только на центральном»
    'Чехол простой 39',
    'на комплекты',
    'для комплектов',
    'Струны цветные',
]

const fileInput = document.getElementById('excelFile')
const fileName = document.getElementById('fileName')
const exportBtn = document.getElementById('exportBtn')
const statusEl = document.getElementById('status')
const amountInput = document.getElementById('amountInput')
const amountResult = document.getElementById('amountResult')
const amountMaxFeedback = document.getElementById('amountMaxFeedback')
const resultsSection = document.getElementById('results')
const resultsBody = document.getElementById('resultsBody')
const resultsHeadings = document.getElementById('resultsHeadings')
const allStockTableWrapper = document.getElementById('allStockTableWrapper')
const allStockSelectiveExportLink = document.getElementById('allStockSelectiveExportLink')
const cancelSelectiveExportBtn = document.getElementById('cancelSelectiveExportBtn')
const selectiveExportBtn = document.getElementById('selectiveExportBtn')
const stopSection = document.getElementById('stopSection')
const stopResultsHeadings = document.getElementById('stopResultsHeadings')
const stopResultsBody = document.getElementById('stopResultsBody')
const emptyMessage = document.getElementById('emptyMessage')
const resultsTitle = document.getElementById('resultsTitle')
const exceptionsAllowedInfo = document.getElementById('exceptionsAllowedInfo')
const centralTab = document.getElementById('centralTab')
const restockTab = document.getElementById('restockTab')
const allStockTab = document.getElementById('allStockTab')
const transferTab = document.getElementById('transferTab')
const revisionTab = document.getElementById('revisionTab')
const exceptionsLink = document.getElementById('exceptionsLink')
const restockBackLink = document.getElementById('restockBackLink')
const resultsPanel = document.getElementById('resultsPanel')
const transferPanel = document.getElementById('transferPanel')
const revisionPanel = document.getElementById('revisionPanel')
const transferBody = document.getElementById('transferBody')
const revisionBody = document.getElementById('revisionBody')
const revisionSelectionBody = document.getElementById('revisionSelectionBody')
const revisionSelectionView = document.getElementById('revisionSelectionView')
const revisionSelectionTableWrapper = document.getElementById('revisionSelectionTableWrapper')
const revisionWorkView = document.getElementById('revisionWorkView')
const revisionSelectionActions = document.getElementById('revisionSelectionActions')
const revisionWorkActions = document.getElementById('revisionWorkActions')
const revisionTitle = document.getElementById('revisionTitle')
const revisionStartBtn = document.getElementById('revisionStartBtn')
const revisionChooseBtn = document.getElementById('revisionChooseBtn')
const revisionSelectionEmptyMessage = document.getElementById('revisionSelectionEmptyMessage')
const revisionTableWrapper = document.getElementById('revisionTableWrapper')
const revisionMessage = document.getElementById('revisionMessage')
const revisionEmptyMessage = document.getElementById('revisionEmptyMessage')
const revisionExportBtn = document.getElementById('revisionExportBtn')
const revisionSort = document.getElementById('revisionSort')
const scanNotFoundToast = document.getElementById('scanNotFoundToast')
const transferEmptyMessage = document.getElementById('transferEmptyMessage')
const transferExportBtn = document.getElementById('transferExportBtn')
const transferPrintBtn = document.getElementById('transferPrintBtn')
const transferPrintDialog = document.getElementById('transferPrintDialog')
const transferPrintDate = document.getElementById('transferPrintDate')
const transferPrintDateOptions = Array.from(transferPrintDate.querySelectorAll('[data-print-date]'))
const pageScrollbar = document.getElementById('pageScrollbar')
const pageScrollbarThumb = pageScrollbar.querySelector('.page-scrollbar-thumb')
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
let selectedAllStockBarcodes = new Set()
let allStockSelectionAnchor = ''
let isSelectingAllStockForExport = false
let revisionResults = []
let selectedRevisionBarcodes = new Set()
let revisionSelectionAnchor = ''
let revisionMode = 'select'
let selectedRevisionBarcode = ''
let revisionSelectionMatchBarcode = ''
let revisionRepeatHintBarcode = ''
let revisionSortMode = 'default'
let stopResults = []
let activeList = 'central'
let transferView = 'items'
let selectedFileName = ''
let scrollbarDragStart = null
let scrollbarUpdateScheduled = false
let amountMaxFeedbackTimeout
let revisionBarcodeBuffer = ''
let revisionBarcodeBufferTimeout
let allStockBarcodeBuffer = ''
let allStockBarcodeBufferTimeout
let revisionSelectionBarcodeBuffer = ''
let revisionSelectionBarcodeBufferTimeout
let scanNotFoundToastTimeout

function formatGroupedDigits(digits) {
    return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ' ')
}

function showAmountMaxFeedback() {
    amountMaxFeedback.hidden = false
    amountMaxFeedback.classList.remove('is-visible')
    void amountMaxFeedback.offsetWidth
    amountMaxFeedback.classList.add('is-visible')

    clearTimeout(amountMaxFeedbackTimeout)
    amountMaxFeedbackTimeout = setTimeout(() => {
        amountMaxFeedback.classList.remove('is-visible')
        amountMaxFeedback.hidden = true
    }, 2000)
}

amountInput.addEventListener('input', () => {
    const rawValue = amountInput.value
    const selectionStart = amountInput.selectionStart ?? rawValue.length
    const digitsBeforeCaret = rawValue.slice(0, selectionStart).replace(/\D/g, '').length
    const enteredDigits = rawValue.replace(/\D/g, '')
    const digits = enteredDigits.slice(0, 7)
    const formattedValue = formatGroupedDigits(digits)

    if (enteredDigits.length > 7) {
        showAmountMaxFeedback()
    }

    amountInput.value = formattedValue
    let caretPosition = 0
    let digitsAtCaret = 0
    while (caretPosition < formattedValue.length && digitsAtCaret < digitsBeforeCaret) {
        if (/\d/.test(formattedValue[caretPosition])) digitsAtCaret += 1
        caretPosition += 1
    }
    amountInput.setSelectionRange(caretPosition, caretPosition)

    amountResult.textContent = digits
        ? formatGroupedDigits(String(Math.round(Number(digits) * 0.03 + 2500)))
        : ''
})

function updatePageScrollbar() {
    scrollbarUpdateScheduled = false
    const viewportHeight = document.documentElement.clientHeight
    const pageHeight = document.documentElement.scrollHeight
    const maxScroll = Math.max(0, pageHeight - viewportHeight)

    pageScrollbar.hidden = maxScroll === 0
    document.body.classList.toggle('has-page-scrollbar', maxScroll > 0)
    if (maxScroll === 0) return

    const trackHeight = pageScrollbar.clientHeight
    const thumbHeight = Math.min(trackHeight, Math.max(64, (viewportHeight / pageHeight) * trackHeight))
    const maxThumbTop = trackHeight - thumbHeight
    const scrollTop = Math.min(maxScroll, Math.max(0, document.documentElement.scrollTop))
    const thumbTop = maxScroll === 0 ? 0 : (scrollTop / maxScroll) * maxThumbTop

    pageScrollbarThumb.style.height = `${thumbHeight}px`
    pageScrollbarThumb.style.transform = `translateY(${thumbTop}px)`
    pageScrollbar.setAttribute('aria-valuenow', String(Math.round((scrollTop / maxScroll) * 100)))
}

function schedulePageScrollbarUpdate() {
    if (scrollbarUpdateScheduled) return
    scrollbarUpdateScheduled = true
    requestAnimationFrame(updatePageScrollbar)
}

function scrollToScrollbarPosition(clientY) {
    const trackRect = pageScrollbar.getBoundingClientRect()
    const maxScroll = document.documentElement.scrollHeight - document.documentElement.clientHeight
    const thumbHeight = pageScrollbarThumb.getBoundingClientRect().height
    const maxThumbTop = trackRect.height - thumbHeight
    const thumbTop = Math.min(maxThumbTop, Math.max(0, clientY - trackRect.top - thumbHeight / 2))
    window.scrollTo(0, maxThumbTop > 0 ? (thumbTop / maxThumbTop) * maxScroll : 0)
}

pageScrollbar.addEventListener('pointerdown', event => {
    if (event.button !== 0) return
    event.preventDefault()
    pageScrollbar.focus({ preventScroll: true })

    if (event.target === pageScrollbarThumb) {
        scrollbarDragStart = {
            pointerY: event.clientY,
            thumbTop: pageScrollbarThumb.getBoundingClientRect().top - pageScrollbar.getBoundingClientRect().top,
        }
        pageScrollbar.setPointerCapture(event.pointerId)
        return
    }

    scrollToScrollbarPosition(event.clientY)
})

pageScrollbar.addEventListener('pointermove', event => {
    if (!scrollbarDragStart) return

    const trackHeight = pageScrollbar.clientHeight
    const thumbHeight = pageScrollbarThumb.getBoundingClientRect().height
    const maxThumbTop = trackHeight - thumbHeight
    const maxScroll = document.documentElement.scrollHeight - document.documentElement.clientHeight
    const thumbTop = Math.min(
        maxThumbTop,
        Math.max(0, scrollbarDragStart.thumbTop + event.clientY - scrollbarDragStart.pointerY),
    )
    window.scrollTo(0, maxThumbTop > 0 ? (thumbTop / maxThumbTop) * maxScroll : 0)
})

const stopScrollbarDrag = () => {
    scrollbarDragStart = null
}

pageScrollbar.addEventListener('pointerup', stopScrollbarDrag)
pageScrollbar.addEventListener('pointercancel', stopScrollbarDrag)
pageScrollbar.addEventListener('keydown', event => {
    const viewportHeight = document.documentElement.clientHeight
    const keyScroll = {
        ArrowDown: 40,
        ArrowUp: -40,
        PageDown: viewportHeight * 0.9,
        PageUp: -viewportHeight * 0.9,
        Home: -document.documentElement.scrollTop,
        End: document.documentElement.scrollHeight,
    }[event.key]

    if (keyScroll === undefined) return
    event.preventDefault()
    window.scrollBy(0, keyScroll)
})

window.addEventListener('scroll', schedulePageScrollbarUpdate, { passive: true })
window.addEventListener('resize', schedulePageScrollbarUpdate)
new ResizeObserver(schedulePageScrollbarUpdate).observe(document.querySelector('.container'))
new MutationObserver(schedulePageScrollbarUpdate).observe(document.body, { childList: true, subtree: true })
schedulePageScrollbarUpdate()

fileInput.addEventListener('change', async () => {
    const file = fileInput.files[0]

    workbook = null
    results = []
    excludedCentralResults = []
    restockResults = []
    excludedRestockResults = []
    allStockResults = []
    selectedAllStockBarcodes = new Set()
    allStockSelectionAnchor = ''
    isSelectingAllStockForExport = false
    allStockSelectiveExportLink.textContent = 'Выборочный экспорт'
    allStockBarcodeBuffer = ''
    clearTimeout(allStockBarcodeBufferTimeout)
    revisionResults = []
    selectedRevisionBarcodes = new Set()
    revisionSelectionAnchor = ''
    revisionSelectionMatchBarcode = ''
    revisionSelectionBarcodeBuffer = ''
    clearTimeout(revisionSelectionBarcodeBufferTimeout)
    revisionMode = 'select'
    selectedRevisionBarcode = ''
    revisionRepeatHintBarcode = ''
    revisionSortMode = 'default'
    revisionSort.value = revisionSortMode
    stopResults = []
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
    revisionTab.classList.remove('is-active')
    revisionTab.setAttribute('aria-selected', 'false')
    revisionPanel.hidden = true
    revisionExportBtn.disabled = true
    revisionEmptyMessage.hidden = false
    revisionMessage.textContent = ''
    revisionMessage.classList.remove('is-error', 'is-warning')
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
allStockSelectiveExportLink.addEventListener('click', () => {
    isSelectingAllStockForExport = !isSelectingAllStockForExport
    selectedAllStockBarcodes.clear()
    allStockSelectionAnchor = ''
    renderResults()
    allStockTableWrapper.focus({ preventScroll: true })
})
cancelSelectiveExportBtn.addEventListener('click', () => {
    isSelectingAllStockForExport = false
    selectedAllStockBarcodes.clear()
    allStockSelectionAnchor = ''
    allStockSelectiveExportLink.textContent = 'Выборочный экспорт'
    renderResults()
    allStockTableWrapper.focus({ preventScroll: true })
})
selectiveExportBtn.addEventListener('click', () => {
    if (selectedAllStockBarcodes.size === 0) return
    exportResults()
})
const allStockDragSelection = attachBarcodeDragSelection(
    resultsBody,
    () => selectedAllStockBarcodes,
    () => activeList === 'all-stock' && isSelectingAllStockForExport,
    () => {
        renderResults()
        allStockTableWrapper.focus({ preventScroll: true })
    },
    barcode => {
        allStockSelectionAnchor = barcode
    },
)
resultsBody.addEventListener('click', event => {
    if (activeList !== 'all-stock' || !isSelectingAllStockForExport) return
    if (allStockDragSelection.consumeClick()) return
    if (event.target.closest('input, button')) return

    const row = event.target.closest('tr[data-barcode]')
    if (!row) return

    allStockSelectionAnchor = updateBarcodeSelection(
        allStockResults,
        selectedAllStockBarcodes,
        allStockSelectionAnchor,
        row.dataset.barcode,
        event.shiftKey,
    )
    renderResults()
    allStockTableWrapper.focus({ preventScroll: true })
})
document.addEventListener('keydown', event => {
    if (activeList !== 'all-stock') return
    if (event.target instanceof HTMLElement && event.target.matches('input, textarea, select, [contenteditable="true"]')) {
        return
    }

    if (event.key === 'Enter') {
        if (!allStockBarcodeBuffer) return
        event.preventDefault()
        allStockBarcodeBuffer = ''
        clearTimeout(allStockBarcodeBufferTimeout)
        return
    }

    if (!/^\d$/.test(event.key) || event.ctrlKey || event.altKey || event.metaKey) return

    event.preventDefault()
    allStockBarcodeBuffer += event.key
    clearTimeout(allStockBarcodeBufferTimeout)
    allStockBarcodeBufferTimeout = setTimeout(() => {
        allStockBarcodeBuffer = ''
    }, 1200)

    if (allStockBarcodeBuffer.length === 13) {
        scrollToAllStockMatch(allStockBarcodeBuffer)
        allStockBarcodeBuffer = ''
        clearTimeout(allStockBarcodeBufferTimeout)
    } else if (allStockBarcodeBuffer.length > 13) {
        allStockBarcodeBuffer = ''
    }
})
document.addEventListener('keydown', event => {
    if (activeList !== 'revision' || revisionMode !== 'select') return
    if (event.target instanceof HTMLElement && event.target.matches('input, textarea, select, [contenteditable="true"]')) {
        return
    }

    if (event.key === 'Enter') {
        if (!revisionSelectionBarcodeBuffer) return
        event.preventDefault()
        revisionSelectionBarcodeBuffer = ''
        clearTimeout(revisionSelectionBarcodeBufferTimeout)
        return
    }

    if (!/^\d$/.test(event.key) || event.ctrlKey || event.altKey || event.metaKey) return

    event.preventDefault()
    revisionSelectionBarcodeBuffer += event.key
    clearTimeout(revisionSelectionBarcodeBufferTimeout)
    revisionSelectionBarcodeBufferTimeout = setTimeout(() => {
        revisionSelectionBarcodeBuffer = ''
    }, 1200)

    if (revisionSelectionBarcodeBuffer.length === 13) {
        scrollToRevisionSelectionMatch(revisionSelectionBarcodeBuffer)
        revisionSelectionBarcodeBuffer = ''
        clearTimeout(revisionSelectionBarcodeBufferTimeout)
    } else if (revisionSelectionBarcodeBuffer.length > 13) {
        revisionSelectionBarcodeBuffer = ''
    }
})
document.addEventListener('click', () => {
    resultsBody.querySelectorAll('.all-stock-match').forEach(row => row.classList.remove('all-stock-match'))
    revisionBody.querySelectorAll('.all-stock-match').forEach(row => row.classList.remove('all-stock-match'))
    revisionSelectionBody
        .querySelectorAll('.all-stock-match')
        .forEach(row => row.classList.remove('all-stock-match'))
    selectedRevisionBarcode = ''
    revisionSelectionMatchBarcode = ''
    allStockBarcodeBuffer = ''
    clearTimeout(allStockBarcodeBufferTimeout)
})
resultsPanel.addEventListener('focusout', event => {
    if (!event.target.matches('.transfer-quantity:not(.revision-fact), .transfer-quantity-button')) return

    window.setTimeout(() => {
        if (activeList === 'all-stock') allStockTableWrapper.focus({ preventScroll: true })
    }, 0)
})
revisionExportBtn.addEventListener('click', exportRevisionResults)
transferExportBtn.addEventListener('click', exportTransferResults)
transferPrintBtn.addEventListener('click', requestTransferPrint)
transferPrintDialog.querySelectorAll('[data-print-position]').forEach(button => {
    button.addEventListener('click', () => {
        submitTransferPrint(button.dataset.printPosition)
    })
})
transferPrintDialog.addEventListener('keydown', event => {
    if (event.key !== 'Escape') return

    event.preventDefault()
    closeTransferPrintDialog()
})
transferPrintDialog.addEventListener('click', event => {
    if (event.target === transferPrintDialog) {
        closeTransferPrintDialog()
    }
})
transferPrintDateOptions.forEach((button, index) => {
    button.addEventListener('click', () => {
        if (button.getAttribute('aria-checked') === 'true') return
        selectTransferPrintDate(button)
    })
    button.addEventListener('keydown', event => {
        if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return

        event.preventDefault()
        const nextIndex =
            event.key === 'Home'
                ? 0
                : event.key === 'End'
                  ? transferPrintDateOptions.length - 1
                  : (index + (event.key === 'ArrowRight' ? 1 : -1) + transferPrintDateOptions.length) %
                    transferPrintDateOptions.length
        const nextOption = transferPrintDateOptions[nextIndex]
        selectTransferPrintDate(nextOption)
        nextOption.focus()
    })
})
centralTab.addEventListener('click', () => selectResultsTab('central'))
restockTab.addEventListener('click', () => selectResultsTab('restock'))
allStockTab.addEventListener('click', () => selectResultsTab('all-stock'))
revisionTab.addEventListener('click', () => selectResultsTab('revision'))
revisionStartBtn.addEventListener('click', () => {
    if (selectedRevisionBarcodes.size === 0) return
    revisionMode = 'work'
    revisionTitle.textContent = 'Ревизия'
    revisionSelectionView.hidden = true
    revisionWorkView.hidden = false
    revisionSelectionActions.hidden = true
    revisionWorkActions.hidden = false
    revisionChooseBtn.hidden = false
    revisionMessage.textContent = ''
    renderRevisionTable()
    revisionTableWrapper.focus({ preventScroll: true })
})
revisionChooseBtn.addEventListener('click', () => {
    revisionMode = 'select'
    revisionTitle.textContent = 'Выбор товаров на ревизию'
    revisionSelectionView.hidden = false
    revisionWorkView.hidden = true
    revisionSelectionActions.hidden = false
    revisionWorkActions.hidden = true
    revisionChooseBtn.hidden = true
    revisionRepeatHintBarcode = ''
    renderRevisionSelectionTable()
    revisionSelectionTableWrapper.focus({ preventScroll: true })
})
revisionSort.addEventListener('change', () => {
    revisionSortMode = revisionSort.value
    renderRevisionTable()
})
const revisionDragSelection = attachBarcodeDragSelection(
    revisionSelectionBody,
    () => selectedRevisionBarcodes,
    () => activeList === 'revision' && revisionMode === 'select',
    renderRevisionSelectionTable,
    barcode => {
        revisionSelectionAnchor = barcode
    },
)
revisionSelectionBody.addEventListener('click', event => {
    if (revisionDragSelection.consumeClick()) return
    const row = event.target.closest('tr[data-barcode]')
    if (!row) return

    const barcode = row.dataset.barcode
    revisionSelectionAnchor = updateBarcodeSelection(
        revisionResults,
        selectedRevisionBarcodes,
        revisionSelectionAnchor,
        barcode,
        event.shiftKey,
    )
    renderRevisionSelectionTable()
})
revisionTableWrapper.addEventListener('keydown', event => {
    if (event.target !== revisionTableWrapper) return

    if (event.key === 'Enter') {
        event.preventDefault()
        scanRevisionBarcode(revisionBarcodeBuffer)
        revisionBarcodeBuffer = ''
        clearTimeout(revisionBarcodeBufferTimeout)
        return
    }

    if (event.key.length !== 1 || event.ctrlKey || event.altKey || event.metaKey) return

    event.preventDefault()
    revisionBarcodeBuffer += event.key
    clearTimeout(revisionBarcodeBufferTimeout)
    revisionBarcodeBufferTimeout = setTimeout(() => {
        revisionBarcodeBuffer = ''
    }, 1500)
})
revisionTableWrapper.addEventListener('click', event => {
    if (!event.target.closest('input, button')) revisionTableWrapper.focus()
})
revisionBody.addEventListener('click', event => {
    if (event.target.closest('input, button')) return

    const row = event.target.closest('tr[data-barcode]')
    if (!row) return

    const product = revisionResults.find(item => normalizeBarcode(item.barcode) === row.dataset.barcode)
    selectedRevisionBarcode = product?.isScanned ? row.dataset.barcode : ''
    revisionRepeatHintBarcode = ''
    renderRevisionTable()
})
revisionPanel.addEventListener('focusout', () => {
    window.setTimeout(() => {
        const activeElement = document.activeElement
        if (
            activeList === 'revision' &&
            revisionMode === 'work' &&
            !activeElement?.matches('.revision-fact, .transfer-quantity-button, .revision-sort, #revisionExportBtn')
        ) {
            revisionTableWrapper.focus({ preventScroll: true })
        }
    }, 0)
})
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
    return JSON.stringify([
        normalizeBarcode(product?.barcode),
        normalizeText(product?.name),
        normalizeText(product?.price ?? product?.salePrice),
    ])
}

function compareProductsByName(a, b) {
    const nameComparison = a.name.localeCompare(b.name, 'ru', { sensitivity: 'base' })
    if (nameComparison) return nameComparison

    return Number((parseStock(b.storeStock) ?? 0) > 0) - Number((parseStock(a.storeStock) ?? 0) > 0)
}

function syncTransferQuantityAcrossLists(productKey, quantity) {
    const normalized = String(quantity ?? '')

    for (const list of [
        results,
        stopResults,
        excludedCentralResults,
        restockResults,
        excludedRestockResults,
        allStockResults,
    ]) {
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
    return parsePrice(value) === 0
}

function parsePrice(value) {
    if (typeof value === 'number') {
        return Number.isFinite(value) ? value : null
    }

    const normalized = String(value ?? '')
        .trim()
        .replace(/[\s\u00a0]/g, '')
        .replace(/[₽$€£]/g, '')
        .replace(/(?:руб\.?|р\.?)/giu, '')
        .replace(',', '.')

    if (!normalized) return null

    const price = Number(normalized)
    return Number.isFinite(price) ? price : null
}

function hasRestockExclusion(name) {
    const normalizedName = normalizeText(name).toLocaleLowerCase('ru')
    return RESTOCK_EXCLUSIONS.some(exclusion => normalizedName.includes(exclusion.toLocaleLowerCase('ru')))
}

function isRestockAllowed(name) {
    const normalizedName = normalizeText(name).toLocaleLowerCase('ru')
    return RESTOCK_ALLOWLIST.some(allowedName => normalizedName === allowedName.toLocaleLowerCase('ru'))
}

function isRestockToTenItem(name) {
    const normalizedName = normalizeText(name).toLocaleLowerCase('ru')
    return RESTOCK_TO_TEN_ITEMS.some(item => normalizedName.includes(item.toLocaleLowerCase('ru')))
}

function hasCentralExclusion(name) {
    const normalizedName = normalizeText(name).toLocaleLowerCase('ru')
    return CENTRAL_EXCLUSIONS.some(exclusion => normalizedName.includes(exclusion.toLocaleLowerCase('ru')))
}

function isStopItem(name) {
    const normalizedName = normalizeText(name).trim()
    return STOP_ITEMS.some(stopItem => stopItem.localeCompare(normalizedName, 'ru', { sensitivity: 'base' }) === 0)
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
            const allProductKey = JSON.stringify([barcode, name, salePrice])
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

        const productKey = JSON.stringify([barcode, name, price])

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
    stopResults = []
    allStockResults = Array.from(allProducts.values()).map(product => ({
        ...product,
        price: product.salePrice ?? product.purchasePrice ?? '',
        transferQuantity: '',
    }))

    const storeBarcodes = new Set(
        Array.from(products.values())
            .filter(product => product.hasStoreRow)
            .map(product => product.barcode),
    )

    for (const product of products.values()) {
        // Наличие товара в Аккорде определяем по штрихкоду, даже если цена отличается.
        if (product.hasCentralRow && !storeBarcodes.has(product.barcode) && !isZeroPrice(product.price)) {
            const centralProduct = {
                name: product.name,
                barcode: product.barcode,
                storeStock: product.storeStock,
                stock: product.centralStock,
                price: product.price,
                transferQuantity: isStopItem(product.name) ? '' : '1',
                isStopItem: isStopItem(product.name),
            }

            if (hasCentralExclusion(product.name)) {
                excludedCentralResults.push(centralProduct)
            } else if (centralProduct.isStopItem) {
                stopResults.push(centralProduct)
            } else {
                results.push(centralProduct)
            }
        }

        const useRestockToTenRule = isRestockToTenItem(product.name)
        const hasStoreStockAtDifferentPrice = Array.from(products.values()).some(
            candidate =>
                candidate.barcode === product.barcode &&
                candidate.storeStock > 0 &&
                candidate.price !== product.price,
        )
        const shouldRestock = useRestockToTenRule
            ? product.storeStock < RESTOCK_TO_TEN_MINIMUM
            : product.storeStock > 0 && product.storeStock < 3

        if (
            shouldRestock &&
            product.centralStock > 0 &&
            (product.storeStock > 0 || !hasStoreStockAtDifferentPrice)
        ) {
            const targetStoreStock = useRestockToTenRule ? RESTOCK_TO_TEN_TARGET : 3
            const restockProduct = {
                name: product.name,
                barcode: product.barcode,
                storeStock: product.storeStock,
                centralStock: product.centralStock,
                price: product.price,
                transferQuantity: String(Math.min(product.centralStock, targetStoreStock - product.storeStock)),
            }

            if (!useRestockToTenRule && hasRestockExclusion(product.name) && !isRestockAllowed(product.name)) {
                excludedRestockResults.push(restockProduct)
            } else {
                restockResults.push(restockProduct)
            }
        }
    }

    results.sort(compareProductsByName)
    excludedCentralResults.sort(compareProductsByName)
    restockResults.sort(compareProductsByName)
    excludedRestockResults.sort(compareProductsByName)
    stopResults.sort(compareProductsByName)
    allStockResults.sort(compareProductsByName)
    revisionResults = allStockResults
        .filter(product => (parseStock(product.storeStock) ?? 0) > 0)
        .map(product => ({
            ...product,
            factQuantity: '',
            isScanned: false,
            zeroFactDecrementCount: 0,
        }))
    selectedRevisionBarcodes = new Set()
    revisionSelectionAnchor = ''
    revisionMode = 'select'
    selectedRevisionBarcode = ''
    revisionRepeatHintBarcode = ''
    revisionSortMode = 'default'
    revisionSort.value = revisionSortMode

    for (const product of [...results, ...stopResults, ...restockResults]) {
        const matchedAllStock = allStockResults.find(item => getProductKey(item) === getProductKey(product))
        if (matchedAllStock) {
            matchedAllStock.transferQuantity = product.transferQuantity
            syncTransferQuantityAcrossLists(getProductKey(product), product.transferQuantity)
        }
    }

    for (const product of [...results, ...stopResults]) {
        syncTransferQuantityAcrossLists(getProductKey(product), product.transferQuantity)
    }

    renderResults()

    statusEl.className = 'status'
    statusEl.replaceChildren()

    const statusCounts = [
        { icon: '✓', label: 'Проверено', count: matchedRows },
        { icon: '✗', label: 'Пропущено', count: skippedRows },
    ]
    statusCounts.forEach(({ icon, label, count }, index) => {
        if (index > 0) statusEl.append(document.createTextNode(', '))

        const group = document.createElement('span')
        group.className = 'status-count'
        group.setAttribute('aria-label', `${label}: ${count}`)

        const statusIcon = document.createElement('span')
        statusIcon.className = 'status-icon'
        statusIcon.setAttribute('aria-hidden', 'true')
        statusIcon.textContent = icon

        const statusCount = document.createElement('span')
        statusCount.textContent = String(count)
        if (index === 1 && count > 0) {
            statusCount.classList.add('status-count-error')
        }
        group.append(statusIcon, statusCount)
        statusEl.append(group)
    })
}

function renderResults() {
    if (activeList === 'revision') {
        resultsPanel.hidden = true
        transferPanel.hidden = true
        revisionPanel.hidden = false
        transferTab.classList.remove('is-active')
        transferTab.setAttribute('aria-selected', 'false')
        revisionTab.classList.add('is-active')
        revisionTab.setAttribute('aria-selected', 'true')
        centralTab.classList.remove('is-active')
        centralTab.setAttribute('aria-selected', 'false')
        restockTab.classList.remove('is-active')
        restockTab.setAttribute('aria-selected', 'false')
        allStockTab.classList.remove('is-active')
        allStockTab.setAttribute('aria-selected', 'false')
        revisionTab.textContent = 'Ревизия'
        const isSelectingRevisionItems = revisionMode === 'select'
        revisionTitle.textContent = isSelectingRevisionItems ? 'Выбор товаров на ревизию' : 'Ревизия'
        revisionSelectionView.hidden = !isSelectingRevisionItems
        revisionWorkView.hidden = isSelectingRevisionItems
        revisionSelectionActions.hidden = !isSelectingRevisionItems
        revisionWorkActions.hidden = isSelectingRevisionItems
        revisionChooseBtn.hidden = isSelectingRevisionItems
        if (isSelectingRevisionItems) {
            renderRevisionSelectionTable()
        } else {
            renderRevisionTable()
        }
        return
    }

    if (activeList === 'transfer') {
        const transferItems = getTransferItems()
        const excludedTransferItems = getNotInTransferItems()
        resultsPanel.hidden = true
        transferPanel.hidden = false
        revisionPanel.hidden = true
        revisionTab.classList.remove('is-active')
        revisionTab.setAttribute('aria-selected', 'false')
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
    revisionPanel.hidden = true
    revisionTab.classList.remove('is-active')
    revisionTab.setAttribute('aria-selected', 'false')
    updateTransferCount()
    transferTab.classList.remove('is-active')
    transferTab.setAttribute('aria-selected', 'false')

    const isAllStockList = activeList === 'all-stock'
    const isCentralList = activeList === 'central'
    const isCentralExceptionsList = activeList === 'central-exceptions'
    const isRestockList = activeList === 'restock' || activeList === 'exceptions'
    const isMoveEditable = isCentralList || activeList === 'restock' || isAllStockList
    const isExceptionsList = isCentralExceptionsList || activeList === 'exceptions'
    resultsBody.classList.toggle('range-selection-body', isAllStockList && isSelectingAllStockForExport)
    resultsBody.closest('table').classList.toggle('selection-table', isAllStockList && isSelectingAllStockForExport)
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
    stopResultsHeadings.replaceChildren()
    stopResultsBody.replaceChildren()

    const headings = isAllStockList
        ? ['Название', 'Ак', 'ЦС', 'Рем', 'Продажа', 'Штрих код', ...(isMoveEditable ? ['Перемещение'] : [])]
        : isRestockList
          ? ['Товар', 'Ак', 'ЦС', 'Цена', 'Штрихкод', ...(isMoveEditable ? ['Перемещение'] : [])]
          : ['Товар', 'ЦС', 'Цена', 'Штрихкод', ...(isMoveEditable ? ['Перемещение'] : [])]
    const renderTableHeadings = tableHeadings => {
        for (const heading of tableHeadings) {
            const th = document.createElement('th')
            th.textContent = heading
            if (heading === 'ЦС') {
                th.classList.add('central-stock-column')
            } else if (['Ак', 'Рем'].includes(heading)) {
                th.classList.add('count-column')
            } else if (['Цена', 'Продажа'].includes(heading)) {
                th.classList.add('price-column')
            } else if (['Штрихкод', 'Штрих код'].includes(heading)) {
                th.classList.add('barcode-column')
            } else if (heading === 'Перемещение') {
                th.classList.add('same-width-column')
            }
            tableHeadings === headings ? resultsHeadings.appendChild(th) : stopResultsHeadings.appendChild(th)
        }
    }

    renderTableHeadings(headings)

    for (const [index, product] of activeResults.entries()) {
        const tr = document.createElement('tr')
        if (isAllStockList) {
            const barcode = normalizeBarcode(product.barcode)
            const isSelected = isSelectingAllStockForExport && selectedAllStockBarcodes.has(barcode)
            const previousIsSelected =
                index > 0 &&
                isSelectingAllStockForExport &&
                selectedAllStockBarcodes.has(normalizeBarcode(activeResults[index - 1].barcode))
            const nextIsSelected =
                index < activeResults.length - 1 &&
                isSelectingAllStockForExport &&
                selectedAllStockBarcodes.has(normalizeBarcode(activeResults[index + 1].barcode))
            tr.dataset.barcode = barcode
            tr.classList.toggle('revision-selection-selected', isSelected)
            tr.classList.toggle('revision-selection-start', isSelected && !previousIsSelected)
            tr.classList.toggle('revision-selection-end', isSelected && !nextIsSelected)
        }

        if (isAllStockList) {
            const nameCell = addCell(tr, '')
            if (isSelectingAllStockForExport && selectedAllStockBarcodes.has(normalizeBarcode(product.barcode))) {
                const selectedIcon = document.createElement('span')
                selectedIcon.className = 'revision-selection-check'
                selectedIcon.textContent = '✓'
                selectedIcon.setAttribute('aria-label', 'Выбран для экспорта')
                nameCell.append(selectedIcon, document.createTextNode(' '))
            }
            nameCell.appendChild(document.createTextNode(product.name))
            addCell(tr, formatStock(product.storeStock))
            addCell(tr, formatStock(product.centralStock))
            addCell(tr, formatStock(product.repairStock))
            addCell(tr, formatPrice(product.salePrice))
            addCell(tr, product.barcode)
            if (isMoveEditable) {
                addTransferQuantityCell(tr, product)
            }
        } else {
            addCell(tr, product.name)

            if (isRestockList) {
                addCell(tr, formatStock(product.storeStock))
                addCell(tr, formatStock(product.centralStock))
            } else {
                addCell(tr, formatStock(product.stock))
            }
            addCell(tr, formatPrice(product.price))
            addCell(tr, product.barcode)
            if (isMoveEditable) {
                addTransferQuantityCell(tr, product)
            }
        }

        for (const [index, heading] of headings.entries()) {
            if (heading === 'ЦС') {
                tr.cells[index].classList.add('central-stock-column')
            } else if (['Ак', 'Рем'].includes(heading)) {
                tr.cells[index].classList.add('count-column')
            } else if (['Цена', 'Продажа'].includes(heading)) {
                tr.cells[index].classList.add('price-column')
            } else if (['Штрихкод', 'Штрих код'].includes(heading)) {
                tr.cells[index].classList.add('barcode-column')
            } else if (heading === 'Перемещение') {
                tr.cells[index].classList.add('same-width-column')
            }
        }

        resultsBody.appendChild(tr)
    }

    if (isCentralList && stopResults.length > 0) {
        stopSection.hidden = false
        const stopHeadings = ['Название', 'ЦС', 'Цена', 'Штрих код', ...(isMoveEditable ? ['Перемещение'] : [])]
        stopResultsHeadings.replaceChildren()
        for (const heading of stopHeadings) {
            const th = document.createElement('th')
            th.textContent = heading
            if (heading === 'ЦС') {
                th.classList.add('central-stock-column')
            } else if (['Цена', 'Продажа'].includes(heading)) {
                th.classList.add('price-column')
            } else if (['Штрихкод', 'Штрих код'].includes(heading)) {
                th.classList.add('barcode-column')
            } else if (heading === 'Перемещение') {
                th.classList.add('same-width-column')
            }
            stopResultsHeadings.appendChild(th)
        }

        for (const product of stopResults) {
            const tr = document.createElement('tr')
            addCell(tr, product.name)
            const stockCell = addCell(tr, formatStock(product.stock))
            stockCell.classList.add('central-stock-column')
            addCell(tr, formatPrice(product.price))
            const barcodeCell = addCell(tr, product.barcode)
            barcodeCell.classList.add('barcode-column')
            if (isMoveEditable) {
                addTransferQuantityCell(tr, product)
            }
            stopResultsBody.appendChild(tr)
        }
    } else {
        stopSection.hidden = true
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
    revisionTab.textContent = 'Ревизия'
    const isCentralActive = isCentralList || isCentralExceptionsList
    const isRestockActive = activeList === 'restock' || activeList === 'exceptions'
    centralTab.classList.toggle('is-active', isCentralActive)
    centralTab.setAttribute('aria-selected', String(isCentralActive))
    restockTab.classList.toggle('is-active', isRestockActive)
    restockTab.setAttribute('aria-selected', String(isRestockActive))
    allStockTab.classList.toggle('is-active', isAllStockList)
    allStockTab.setAttribute('aria-selected', String(isAllStockList))
    revisionTab.setAttribute('aria-selected', 'false')
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
    exportBtn.hidden = !isAllStockList || isSelectingAllStockForExport
    allStockSelectiveExportLink.hidden = !isAllStockList
    allStockSelectiveExportLink.textContent = isSelectingAllStockForExport
        ? 'Все товары'
        : 'Выборочный экспорт'
    cancelSelectiveExportBtn.hidden = !isAllStockList || !isSelectingAllStockForExport
    selectiveExportBtn.hidden = !isAllStockList || !isSelectingAllStockForExport
    selectiveExportBtn.disabled = selectedAllStockBarcodes.size === 0
    resultsPanel.classList.toggle('is-selecting-export', isAllStockList && isSelectingAllStockForExport)
}

function updateBarcodeSelection(products, selection, anchorBarcode, selectedBarcode, extendRange) {
    const sortedProducts = [...products].sort(compareProductsByName)
    const selectedIndex = sortedProducts.findIndex(
        product => normalizeBarcode(product.barcode) === selectedBarcode,
    )
    if (selectedIndex === -1) return anchorBarcode

    if (extendRange && anchorBarcode) {
        const anchorIndex = sortedProducts.findIndex(
            product => normalizeBarcode(product.barcode) === anchorBarcode,
        )
        if (anchorIndex !== -1) {
            const start = Math.min(anchorIndex, selectedIndex)
            const end = Math.max(anchorIndex, selectedIndex)
            const range = sortedProducts
                .slice(start, end + 1)
                .map(product => normalizeBarcode(product.barcode))
            const shouldDeselect = range.every(barcode => selection.has(barcode))
            for (let index = start; index <= end; index++) {
                const barcode = normalizeBarcode(sortedProducts[index].barcode)
                if (shouldDeselect) {
                    selection.delete(barcode)
                } else {
                    selection.add(barcode)
                }
            }
            return anchorBarcode
        }
    }

    if (selection.has(selectedBarcode)) {
        selection.delete(selectedBarcode)
    } else {
        selection.add(selectedBarcode)
    }
    return selectedBarcode
}

function attachBarcodeDragSelection(body, getSelection, isEnabled, onSelectionChange, setAnchor) {
    let startBarcode = ''
    let currentBarcode = ''
    let suppressClick = false
    let shouldSelect = false
    let didDrag = false

    body.addEventListener(
        'click',
        event => {
            if (!suppressClick) return
            suppressClick = false
            event.preventDefault()
            event.stopImmediatePropagation()
        },
        true,
    )

    body.addEventListener('pointerdown', event => {
        suppressClick = false
        startBarcode = ''
        currentBarcode = ''
        didDrag = false
        if (event.pointerType !== 'mouse' || event.button !== 0 || !isEnabled()) return
        if (event.target.closest('input, button, select, textarea, a')) return
        const row = event.target.closest('tr[data-barcode]')
        if (!row) return
        startBarcode = row.dataset.barcode
        currentBarcode = startBarcode
        shouldSelect = !getSelection().has(startBarcode)
    })

    body.addEventListener('pointerover', event => {
        if (!startBarcode || !(event.buttons & 1) || !isEnabled()) return
        const row = event.target.closest('tr[data-barcode]')
        if (!row || row.dataset.barcode === currentBarcode) return
        didDrag = true
        currentBarcode = row.dataset.barcode
        const selection = getSelection()
        const orderedBarcodes = Array.from(body.querySelectorAll('tr[data-barcode]'), item => item.dataset.barcode)
        const startIndex = orderedBarcodes.indexOf(startBarcode)
        const currentIndex = orderedBarcodes.indexOf(currentBarcode)
        for (const barcode of orderedBarcodes.slice(
            Math.min(startIndex, currentIndex),
            Math.max(startIndex, currentIndex) + 1,
        )) {
            if (shouldSelect) {
                selection.add(barcode)
            } else {
                selection.delete(barcode)
            }
        }
        renderBarcodeSelectionState(body, selection)
    })

    document.addEventListener('pointerup', () => {
        if (!startBarcode) return
        if (didDrag && isEnabled()) {
            setAnchor(startBarcode)
            onSelectionChange()
            suppressClick = true
        }
        startBarcode = ''
        currentBarcode = ''
        didDrag = false
    })

    return {
        consumeClick() {
            if (!suppressClick) return false
            suppressClick = false
            return true
        },
    }
}

function renderBarcodeSelectionState(body, selection) {
    const rows = Array.from(body.querySelectorAll('tr[data-barcode]'))
    rows.forEach((row, index) => {
        const barcode = row.dataset.barcode
        const isSelected = selection.has(barcode)
        const previousIsSelected = index > 0 && selection.has(rows[index - 1].dataset.barcode)
        const nextIsSelected = index < rows.length - 1 && selection.has(rows[index + 1].dataset.barcode)
        row.classList.toggle('revision-selection-selected', isSelected)
        row.classList.toggle('revision-selection-start', isSelected && !previousIsSelected)
        row.classList.toggle('revision-selection-end', isSelected && !nextIsSelected)

        const nameCell = row.cells[0]
        const check = nameCell.querySelector('.revision-selection-check')
        if (isSelected && !check) {
            const icon = document.createElement('span')
            icon.className = 'revision-selection-check'
            icon.textContent = '✓'
            icon.setAttribute('aria-label', body === resultsBody ? 'Выбран для экспорта' : 'Выбран на ревизию')
            nameCell.prepend(icon, document.createTextNode(' '))
        } else if (!isSelected && check) {
            check.nextSibling?.textContent === ' ' && check.nextSibling.remove()
            check.remove()
        }
    })
}

function scrollToAllStockMatch(barcode) {
    resultsBody.querySelectorAll('.all-stock-match').forEach(row => row.classList.remove('all-stock-match'))
    const match = Array.from(resultsBody.rows).find(row => row.dataset.barcode === barcode)
    if (!match) {
        showScanNotFoundToast()
        return false
    }

    hideScanNotFoundToast()
    match.classList.add('all-stock-match')
    match.scrollIntoView({ block: 'center', inline: 'nearest' })
    return true
}

function scrollToRevisionSelectionMatch(barcode) {
    revisionSelectionBody.querySelectorAll('.all-stock-match').forEach(row => row.classList.remove('all-stock-match'))
    const match = Array.from(revisionSelectionBody.rows).find(row => row.dataset.barcode === normalizeBarcode(barcode))
    if (!match) {
        revisionSelectionMatchBarcode = ''
        renderRevisionSelectionTable()
        showScanNotFoundToast()
        return false
    }

    hideScanNotFoundToast()
    revisionSelectionMatchBarcode = match.dataset.barcode
    renderRevisionSelectionTable()
    const selectedRow = Array.from(revisionSelectionBody.rows).find(
        row => row.dataset.barcode === revisionSelectionMatchBarcode,
    )
    selectedRow?.scrollIntoView({ block: 'center', inline: 'nearest' })
    return true
}

function showScanNotFoundToast() {
    scanNotFoundToast.hidden = false
    scanNotFoundToast.classList.remove('is-visible')
    void scanNotFoundToast.offsetWidth
    scanNotFoundToast.classList.add('is-visible')
    clearTimeout(scanNotFoundToastTimeout)
    scanNotFoundToastTimeout = setTimeout(hideScanNotFoundToast, 2500)
}

function hideScanNotFoundToast() {
    clearTimeout(scanNotFoundToastTimeout)
    scanNotFoundToast.classList.remove('is-visible')
    scanNotFoundToast.hidden = true
}

function getTransferItems() {
    const seen = new Set()
    return [...results, ...stopResults, ...restockResults, ...allStockResults]
        .filter(product => {
            if (parseStock(product.transferQuantity) <= 0) return false

            const key = getProductKey(product)
            if (seen.has(key)) return false
            seen.add(key)
            return true
        })
        .sort(compareProductsByName)
}

function getNotInTransferItems() {
    const seen = new Set()
    return [...results, ...stopResults, ...restockResults]
        .filter(product => {
            const key = getProductKey(product)
            if (seen.has(key)) return false
            seen.add(key)
            return parseStock(product.transferQuantity) <= 0
        })
        .sort(compareProductsByName)
}

function updateTransferCount() {
    transferTab.textContent = `Перемещение ${getTransferItems().length}`
}

function scanRevisionBarcode(value) {
    const barcode = normalizeBarcode(value)
    if (!barcode) return

    const product = getSelectedRevisionResults().find(item => normalizeBarcode(item.barcode) === barcode)
    if (!product) {
        selectedRevisionBarcode = ''
        revisionRepeatHintBarcode = ''
        const knownProduct = allStockResults.find(item => normalizeBarcode(item.barcode) === barcode)
        revisionMessage.textContent = knownProduct
            ? `Товара «${knownProduct.name}» нет в списке ревизии`
            : `Товар со штрихкодом ${barcode} не найден`
        revisionMessage.classList.remove('is-warning')
        revisionMessage.classList.add('is-error')
        renderRevisionTable()
        return
    }

    selectedRevisionBarcode = barcode
    revisionRepeatHintBarcode = product.isScanned ? barcode : ''
    if (!product.isScanned) {
        product.factQuantity = String(parseStock(product.storeStock) ?? 0)
    }
    product.isScanned = true
    product.zeroFactDecrementCount = 0
    revisionMessage.textContent = ''
    revisionMessage.classList.remove('is-error', 'is-warning')
    renderRevisionTable()
    const selectedRow = revisionBody.querySelector('.all-stock-match')
    selectedRow?.scrollIntoView({ block: 'center', inline: 'nearest' })
}

function renderRevisionTable() {
    revisionBody.replaceChildren()
    const selectedProducts = getSelectedRevisionResults()
    for (const product of getSortedRevisionResults()) {
        const row = document.createElement('tr')
        row.dataset.barcode = normalizeBarcode(product.barcode)
        row.classList.toggle('revision-scanned', product.isScanned)
        row.classList.toggle('revision-unscanned', !product.isScanned)
        if (normalizeBarcode(product.barcode) === selectedRevisionBarcode) {
            row.classList.add('all-stock-match')
        }

        const nameCell = addCell(row, '')
        if (product.isScanned) {
            const isDifferent =
                (parseStock(product.factQuantity) ?? 0) !== (parseStock(product.storeStock) ?? 0)
            const statusIcon = document.createElement('span')
            statusIcon.className = `revision-status-icon ${isDifferent ? 'is-different' : 'is-checked'}`
            statusIcon.textContent = isDifferent ? '!' : '✓'
            statusIcon.setAttribute('aria-label', isDifferent ? 'Факт отличается от Ак' : 'Проверено')
            nameCell.append(statusIcon, document.createTextNode(' '))
        }
        nameCell.appendChild(document.createTextNode(product.name))
        if (normalizeBarcode(product.barcode) === revisionRepeatHintBarcode) {
            const hint = document.createElement('span')
            hint.className = 'revision-row-hint'
            hint.textContent = 'Уже отсканирован'
            nameCell.appendChild(hint)
        }
        const storeStockCell = addCell(row, formatStock(product.storeStock ?? 0))
        storeStockCell.classList.add('count-column')
        const centralStockCell = addCell(row, formatStock(product.centralStock ?? 0))
        centralStockCell.classList.add('central-stock-column')
        const priceCell = addCell(row, formatPrice(product.salePrice ?? product.price ?? ''))
        priceCell.classList.add('price-column')
        const barcodeCell = addCell(row, product.barcode)
        barcodeCell.classList.add('barcode-column')
        addRevisionFactCell(row, product)
        revisionBody.appendChild(row)
    }

    revisionEmptyMessage.textContent = selectedProducts.length ? 'Просканируйте товар.' : 'Выберите товары для ревизии.'
    revisionEmptyMessage.hidden = selectedProducts.length !== 0
    revisionExportBtn.disabled = selectedProducts.length === 0
    revisionTab.textContent = 'Ревизия'
}

function getSelectedRevisionResults() {
    return revisionResults.filter(product => selectedRevisionBarcodes.has(normalizeBarcode(product.barcode)))
}

function renderRevisionSelectionTable() {
    const products = [...revisionResults].sort(compareProductsByName)
    revisionSelectionBody.replaceChildren()
    for (const [index, product] of products.entries()) {
        const barcode = normalizeBarcode(product.barcode)
        const isSelected = selectedRevisionBarcodes.has(barcode)
        const previousIsSelected =
            index > 0 && selectedRevisionBarcodes.has(normalizeBarcode(products[index - 1].barcode))
        const nextIsSelected =
            index < products.length - 1 &&
            selectedRevisionBarcodes.has(normalizeBarcode(products[index + 1].barcode))
        const row = document.createElement('tr')
        row.dataset.barcode = barcode
        row.classList.toggle('all-stock-match', barcode === revisionSelectionMatchBarcode)
        row.classList.toggle('revision-selection-selected', isSelected)
        row.classList.toggle('revision-selection-start', isSelected && !previousIsSelected)
        row.classList.toggle('revision-selection-end', isSelected && !nextIsSelected)
        const nameCell = addCell(row, '')
        if (selectedRevisionBarcodes.has(barcode)) {
            const selectedIcon = document.createElement('span')
            selectedIcon.className = 'revision-selection-check'
            selectedIcon.textContent = '✓'
            selectedIcon.setAttribute('aria-label', 'Выбран на ревизию')
            nameCell.append(selectedIcon, document.createTextNode(' '))
        }
        nameCell.appendChild(document.createTextNode(product.name))
        const storeStockCell = addCell(row, formatStock(product.storeStock ?? 0))
        storeStockCell.classList.add('count-column')
        const centralStockCell = addCell(row, formatStock(product.centralStock ?? 0))
        centralStockCell.classList.add('central-stock-column')
        const priceCell = addCell(row, formatPrice(product.salePrice ?? product.price ?? ''))
        priceCell.classList.add('price-column')
        const barcodeCell = addCell(row, product.barcode)
        barcodeCell.classList.add('barcode-column')
        revisionSelectionBody.appendChild(row)
    }

    revisionStartBtn.disabled = selectedRevisionBarcodes.size === 0
    revisionSelectionEmptyMessage.hidden = products.length !== 0
    revisionTab.textContent = 'Ревизия'
}

function getSortedRevisionResults() {
    return getSelectedRevisionResults().sort((a, b) => {
        const nameOrder = compareProductsByName(a, b)
        if (revisionSortMode === 'unscanned') {
            return Number(a.isScanned) - Number(b.isScanned) || nameOrder
        }
        if (revisionSortMode === 'different') {
            const aDiffers = a.isScanned && (parseStock(a.factQuantity) ?? 0) !== (parseStock(a.storeStock) ?? 0)
            const bDiffers = b.isScanned && (parseStock(b.factQuantity) ?? 0) !== (parseStock(b.storeStock) ?? 0)
            return Number(bDiffers) - Number(aDiffers) || nameOrder
        }
        return nameOrder
    })
}

function addRevisionFactCell(row, product) {
    const cell = document.createElement('td')
    cell.className = 'transfer-quantity-cell same-width-column'
    const controls = document.createElement('div')
    controls.className = 'transfer-quantity-controls'
    const factInput = document.createElement('input')
    factInput.className = 'transfer-quantity revision-fact'
    factInput.type = 'number'
    factInput.min = '0'
    factInput.step = 'any'
    factInput.value = product.isScanned ? String(parseStock(product.factQuantity) ?? 0) : ''
    factInput.disabled = !product.isScanned
    factInput.setAttribute('aria-label', `Фактическое количество: ${product.name}`)

    const updateFact = () => {
        const quantity = parseStock(factInput.value)
        if (quantity !== null) {
            const boundedQuantity = Math.max(0, quantity)
            factInput.value = String(boundedQuantity)
            product.factQuantity = String(boundedQuantity)
            if (boundedQuantity > 0) {
                product.zeroFactDecrementCount = 0
                revisionMessage.textContent = ''
            }
        } else if (!factInput.value.trim()) {
            product.factQuantity = ''
        }
        const isDifferent =
            product.isScanned && (parseStock(product.factQuantity) ?? 0) !== (parseStock(product.storeStock) ?? 0)
        factInput.classList.toggle('is-different', isDifferent)
        const statusIcon = row.querySelector('.revision-status-icon')
        if (statusIcon) {
            statusIcon.textContent = isDifferent ? '!' : '✓'
            statusIcon.classList.toggle('is-different', isDifferent)
            statusIcon.classList.toggle('is-checked', !isDifferent)
            statusIcon.setAttribute('aria-label', isDifferent ? 'Факт отличается от Ак' : 'Проверено')
        }
    }
    factInput.addEventListener('input', updateFact)
    updateFact()

    const createButton = (label, delta, accessibleLabel) => {
        const button = document.createElement('button')
        button.type = 'button'
        button.className = 'transfer-quantity-button'
        button.textContent = label
        button.disabled = !product.isScanned
        button.setAttribute('aria-label', `${accessibleLabel}: ${product.name}`)
        button.addEventListener('click', () => {
            const currentQuantity = parseStock(factInput.value) ?? 0
            if (delta < 0 && currentQuantity === 0) {
                product.zeroFactDecrementCount = (product.zeroFactDecrementCount ?? 0) + 1
                if (product.zeroFactDecrementCount >= 5) {
                    product.isScanned = false
                    product.factQuantity = ''
                    product.zeroFactDecrementCount = 0
                    if (selectedRevisionBarcode === normalizeBarcode(product.barcode)) {
                        selectedRevisionBarcode = ''
                        revisionRepeatHintBarcode = ''
                    }
                    renderRevisionTable()
                    revisionTableWrapper.focus({ preventScroll: true })
                    return
                }
                revisionMessage.textContent = `Для удаления ${product.name} нажмите «−» ещё ${5 - product.zeroFactDecrementCount} раз`
            } else if (delta > 0 || currentQuantity > 0) {
                product.zeroFactDecrementCount = 0
                revisionMessage.textContent = ''
            }
            factInput.value = String(Math.max(0, currentQuantity + delta))
            updateFact()
            revisionTableWrapper.focus({ preventScroll: true })
        })
        return button
    }

    controls.append(
        factInput,
        createButton('−', -1, 'Уменьшить факт'),
        createButton('+', 1, 'Увеличить факт'),
    )
    cell.appendChild(controls)
    row.appendChild(cell)
}

function renderTransferTable(products, isExcludedList = false) {
    transferBody.replaceChildren()
    for (const product of products) {
        const row = document.createElement('tr')
        addCell(row, product.name)
        const storeStockCell = addCell(row, formatStock(product.storeStock ?? 0))
        storeStockCell.classList.add('count-column')
        const centralStock = parseStock(product.centralStock ?? product.stock) ?? 0
        const centralStockCell = addCell(row, formatStock(centralStock))
        centralStockCell.classList.add('central-stock-column')
        const priceCell = addCell(row, formatPrice(product.price ?? product.salePrice ?? product.purchasePrice ?? ''))
        priceCell.classList.add('price-column')
        addTransferQuantityCell(row, product, true, isExcludedList)
        transferBody.appendChild(row)
    }
    transferEmptyMessage.textContent = isExcludedList
        ? 'Товаров, не вошедших в перемещение, не найдено.'
        : 'В таблице на перемещение пока нет товаров.'
    transferEmptyMessage.hidden = products.length !== 0
    transferExportBtn.disabled = products.length === 0
    transferPrintBtn.disabled = false
}

function addTransferQuantityCell(row, product, isTransferTable = false, isExcludedTable = false) {
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

    const quantityInputWrap = document.createElement('div')
    quantityInputWrap.className = 'transfer-quantity-input-wrap'
    const maxFeedback = document.createElement('span')
    maxFeedback.className = 'transfer-max-feedback'
    maxFeedback.textContent = 'Max'
    maxFeedback.setAttribute('role', 'status')
    maxFeedback.setAttribute('aria-live', 'polite')
    maxFeedback.hidden = true
    quantityInputWrap.append(quantityInput, maxFeedback)

    let feedbackTimeout
    const showLimitFeedback = () => {
        const centralStockCell = row.querySelector('.central-stock-column')
        if (centralStockCell) {
            centralStockCell.classList.remove('is-limit-warning')
            void centralStockCell.offsetWidth
            centralStockCell.classList.add('is-limit-warning')
        }

        maxFeedback.hidden = false
        maxFeedback.classList.remove('is-visible')
        void maxFeedback.offsetWidth
        maxFeedback.classList.add('is-visible')

        clearTimeout(feedbackTimeout)
        feedbackTimeout = setTimeout(() => {
            centralStockCell?.classList.remove('is-limit-warning')
            maxFeedback.classList.remove('is-visible')
            maxFeedback.hidden = true
        }, 2000)
    }

    const updateQuantity = () => {
        const enteredQuantity = parseStock(quantityInput.value)
        if (enteredQuantity !== null) {
            if (enteredQuantity > maxQuantity) showLimitFeedback()
            quantityInput.value = String(Math.min(maxQuantity, Math.max(0, enteredQuantity)))
        }
        product.transferQuantity = quantityInput.value
        syncTransferQuantityAcrossLists(getProductKey(product), product.transferQuantity)
        updateTransferCount()
        if (activeList === 'transfer') {
            const hasTransferItems = getTransferItems().length > 0
            transferExportBtn.disabled = !hasTransferItems
        }
    }
    const shouldLeaveCurrentTransferView = () => {
        const quantity = parseStock(product.transferQuantity)
        return isExcludedTable ? quantity > 0 : quantity <= 0
    }

    quantityInput.addEventListener('input', updateQuantity)
    if (isTransferTable) {
        quantityInput.addEventListener('blur', () => {
            if (shouldLeaveCurrentTransferView()) {
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
        if (nextQuantity > maxQuantity) {
            showLimitFeedback()
            return
        }
        if (nextQuantity < 0) return

        quantityInput.value = String(nextQuantity)
        updateQuantity()
        if (isTransferTable && shouldLeaveCurrentTransferView()) {
            renderResults()
        }
    }

    decrementButton.addEventListener('click', () => adjustQuantity(-1))
    incrementButton.addEventListener('click', () => adjustQuantity(1))

    controls.append(quantityInputWrap, decrementButton, incrementButton)
    cell.appendChild(controls)
    row.appendChild(cell)
}

function selectResultsTab(tab) {
    window.scrollTo(0, 0)
    activeList = tab
    const isCentralList = tab === 'central'
    const isRestockList = tab === 'restock'
    const isAllStockList = tab === 'all-stock'
    const isRevisionList = tab === 'revision'

    centralTab.classList.toggle('is-active', isCentralList)
    centralTab.setAttribute('aria-selected', String(isCentralList))
    restockTab.classList.toggle('is-active', isRestockList)
    restockTab.setAttribute('aria-selected', String(isRestockList))
    allStockTab.classList.toggle('is-active', isAllStockList)
    allStockTab.setAttribute('aria-selected', String(isAllStockList))
    revisionTab.classList.toggle('is-active', isRevisionList)
    revisionTab.setAttribute('aria-selected', String(isRevisionList))
    resultsPanel.setAttribute(
        'aria-labelledby',
        isRevisionList
            ? 'revisionTab'
            : isCentralList
              ? 'centralTab'
              : isRestockList
                ? 'restockTab'
                : 'allStockTab',
    )

    renderResults()
    if (isRevisionList && revisionMode === 'select') revisionSelectionTableWrapper.focus({ preventScroll: true })
    if (isRevisionList && revisionMode === 'work') revisionTableWrapper.focus({ preventScroll: true })
    if (isAllStockList) allStockTableWrapper.focus({ preventScroll: true })
}

function selectRestockView(view) {
    activeList = view
    renderResults()
}

function addCell(row, value) {
    const td = document.createElement('td')
    td.textContent = String(value ?? '')
    row.appendChild(td)
    return td
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

function getTransferPrintDateValue(selection) {
    const date = new Date()
    if (selection === 'tomorrow') date.setDate(date.getDate() + 1)
    const year = date.getFullYear()
    const month = String(date.getMonth() + 1).padStart(2, '0')
    const day = String(date.getDate()).padStart(2, '0')
    return `${year}-${month}-${day}`
}

function formatTransferPrintDate(value) {
    const [year, month, day] = value.split('-').map(Number)
    return new Intl.DateTimeFormat('ru-RU', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
    }).format(new Date(year, month - 1, day))
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

function requestTransferPrint() {
    const selection = transferPrintDate.querySelector('[aria-checked="true"]').dataset.printDate
    printTransferResults(null, getTransferPrintDateValue(selection), true)
}

function selectTransferPrintDate(selectedOption) {
    for (const option of transferPrintDateOptions) {
        const isSelected = option === selectedOption
        option.classList.toggle('is-selected', isSelected)
        option.setAttribute('aria-checked', String(isSelected))
        option.setAttribute('aria-disabled', String(isSelected))
        option.tabIndex = isSelected ? 0 : -1
    }
}

function submitTransferPrint(position) {
    transferPrintDialog.close()
    const selection = transferPrintDate.querySelector('[aria-checked="true"]').dataset.printDate
    printTransferResults(position, getTransferPrintDateValue(selection))
}

function closeTransferPrintDialog() {
    transferPrintDialog.close()
    requestAnimationFrame(() => {
        if (document.activeElement === transferPrintBtn) {
            transferPrintBtn.blur()
        }
    })
}

function printTransferResults(position, date, isPreview = false) {
    const transferItems = getTransferItems()

    const printFrame = document.createElement('iframe')
    printFrame.setAttribute('title', 'Печать накладной')
    printFrame.setAttribute('aria-hidden', 'true')
    printFrame.style.position = 'fixed'
    printFrame.style.left = '-10000px'
    printFrame.style.top = '0'
    printFrame.style.width = '210mm'
    printFrame.style.height = '297mm'
    printFrame.style.border = '0'
    printFrame.style.opacity = '0'
    printFrame.style.pointerEvents = 'none'
    document.body.appendChild(printFrame)

    const printWindow = printFrame.contentWindow
    if (!printWindow) {
        document.body.removeChild(printFrame)
        setError('Не удалось открыть окно печати. Разрешите всплывающие окна и попробуйте снова.')
        return
    }

    const totalQuantity = transferItems.reduce((sum, product) => sum + (parseStock(product.transferQuantity) ?? 0), 0)
    const totalAmount = transferItems.reduce(
        (sum, product) => sum + (parseStock(product.transferQuantity) ?? 0) * (parsePrice(product.price) ?? 0),
        0,
    )

    const itemRows = transferItems.map((product, index) => {
        const quantity = parseStock(product.transferQuantity) ?? 0
        const barcodeValue = String(product.barcode ?? '').trim()
        const barcodeId = `barcode-${index}`
        const barcodeFormat =
            barcodeValue && /^\d+$/.test(barcodeValue) ? (barcodeValue.length === 13 ? 'EAN13' : 'CODE128') : 'CODE128'
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
                <td class="price-cell">
                  <span>${escapeHtml(formatPrice(product.price))}</span>
                </td>
                <td class="qty-cell">${escapeHtml(String(quantity))}</td>
                <td class="accord-price-cell">
                  ${
                      quantity > 1
                          ? `<span class="print-item-total">${escapeHtml(
                                formatPrice(String(quantity * (parsePrice(product.price) ?? 0))),
                            )}</span>`
                          : ''
                  }
                </td>
              </tr>
            `
    })

    const printDate = formatTransferPrintDate(date)

    const summaryRow = `
                <tr class="summary-row">
                    <td class="summary-values-cell">
                        <span class="summary-values">
                            <span class="summary-value">${escapeHtml(formatPrice(totalAmount))}</span>
                            <span class="summary-value">${escapeHtml(String(totalQuantity))}</span>
                        </span>
                    </td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                </tr>
        `

    const pageMarkup = `
            <main class="print-page">
                <div class="print-offset" aria-hidden="true"></div>
                <section class="print-content">
                    <header class="print-header">
                        <div class="print-header-left">
                            <h1>Накладная на перемещение</h1>
                            <p class="print-note">Если цена отличается от ЦС, выделить маркером и отправить в реализацию</p>
                        </div>
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
                        <tbody>${itemRows.join('')}${summaryRow}</tbody>
                    </table>
                </section>
            </main>
        `

    printWindow.document.write(`<!DOCTYPE html>
    <html lang="ru">
      <head>
        <meta charset="UTF-8">
        <title>Накладная на перемещение ${printDate}</title>
        <style>
          @page { margin: 0; size: A4 portrait; }
                    body {
            margin: 0;
                    padding: 16mm 16mm 11mm;
            font-family: Arial, Helvetica, sans-serif;
            color: #202124;
          }
                                        .print-page {
                                            height: 265mm;
                                            break-after: page;
                                            page-break-after: always;
                                        }
                                        .print-page:last-of-type {
                                            break-after: auto;
                                            page-break-after: auto;
                                        }
                    .print-offset { height: 0; }
          .print-header {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 12px;
            margin-bottom: 12px;
          }
                    .print-header-continuation { padding-top: 14mm; }
          .print-header-left {
            display: flex;
            flex-direction: column;
            gap: 4px;
          }
          h1 {
            margin: 0;
            font-size: 18px;
            font-weight: 700;
          }
          .print-note {
            margin: 0;
            font-size: 9px;
            line-height: 1.3;
            color: #5f6368;
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
          th:nth-child(2), td:nth-child(2) { width: 10%; }
          th:nth-child(3), td:nth-child(3) { width: 4%; }
          th:nth-child(4), td:nth-child(4) { width: 8%; }
          th, td {
                        padding: 1.5px 4px;
            border: 1px solid #999;
                        font-size: 10px;
                        line-height: 1.15;
                        height: 16px;
            vertical-align: middle;
          }
                    thead { display: table-header-group; }
                    tr { break-inside: avoid; }
          th {
            background: #f2f2f2;
            font-weight: 700;
          }
          .summary-row td {
                        font-weight: 400;
            background: #fafafa;
          }
                    .summary-value,
                    .print-item-total {
                        color: #bdbdbd;
                        font-family: inherit;
                        font-size: 8px;
                        font-weight: 400;
                        line-height: 1.15;
                    }
                    .summary-values-cell { text-align: right; }
                    .summary-values { display: inline-flex; gap: 10px; }
          .product-name-cell {
            padding: 0;
            border-left: 1px solid #999;
          }
          .print-product-name-wrapper {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 8px;
                        min-height: 24px;
            padding: 0 4px;
          }
          .print-product-name {
                        font-size: 14px;
            font-weight: 400;
                        line-height: 1.2;
            flex: 1 1 auto;
            white-space: normal;
            min-width: 0;
          }
          .print-product-barcode {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            flex: 0 0 auto;
            min-width: 0;
            width: auto;
            padding: 0;
            border: none;
            background: transparent;
          }
          .print-barcode {
            display: block;
                        width: auto;
                        max-width: none;
                                                height: 16.8px;
            background: #fff;
          }
          .empty-barcode {
                        width: 60px;
                        height: 16px;
            border: none;
          }
          .price-cell,
          .qty-cell {
            text-align: left;
          }
                    .accord-price-cell { text-align: right; }
                    .print-item-total {
                        display: block;
                    }
                    .price-cell,
                    .qty-cell {
                        font-size: 14px;
                    }
          @media print {
                        body { margin: 0; padding: 16mm 16mm 11mm; }
            th { background: #f2f2f2 !important; print-color-adjust: exact; }
          }
        </style>
      </head>
                        <body>${pageMarkup}
        <script src="https://cdn.jsdelivr.net/npm/jsbarcode@3.11.6/dist/JsBarcode.all.min.js"></script>
        <script>
          window.onload = () => {
            const barcodes = Array.from(document.querySelectorAll('.print-barcode'));
            barcodes.forEach(svg => {
              const code = svg.dataset.barcodeCode || '';
                if (code && typeof JsBarcode === 'function') {
                const format = /^\d+$/.test(code) ? (code.length === 13 ? 'EAN13' : 'CODE128') : 'CODE128';
                JsBarcode(svg, code, {
                  format,
                  displayValue: false,
                                    width: 0.6435,
                  height: 16.8,
                  margin: 0,
                  background: '#fff',
                  lineColor: '#000'
                });
              }
            });
                                                const firstPage = document.querySelector('.print-page');
                                                const firstContent = firstPage.querySelector('.print-content');
                                                const pageHeight = firstPage.clientHeight;
                                                const usablePageHeight = pageHeight;
                                                const halfPageHeight = pageHeight / 2;
                                                const desiredBottomGap = (8 * 96) / 25.4;
                                                const firstBody = firstPage.querySelector('tbody');
                                                const productRows = Array.from(firstBody.querySelectorAll('tr')).filter(
                                                    row => !row.classList.contains('summary-row')
                                                );
                                                const summaryRow = firstBody.querySelector('.summary-row');
                                                const lowerHalfCapacity = usablePageHeight - halfPageHeight;
                                                const initialContentHeight = firstContent.getBoundingClientRect().height;
                                                const fitsHalfPage = initialContentHeight <= Math.min(halfPageHeight, lowerHalfCapacity);
                                                const isPrintPreview = ${isPreview};

                                                if (isPrintPreview) {
                                                    const dialog = window.parent.document.getElementById('transferPrintDialog');
                                                    dialog.querySelector('[data-print-position="top"]').hidden = !fitsHalfPage;
                                                    dialog.querySelector('[data-print-position="bottom"]').hidden = !fitsHalfPage;
                                                    dialog.querySelector('[data-print-position="full"]').hidden = fitsHalfPage;
                                                    dialog.showModal();
                                                    window.frameElement.remove();
                                                    return;
                                                }

                                                const printPages = [firstPage];
                                                const pageTemplate = firstPage.cloneNode(true);
                                                pageTemplate.querySelector('tbody').replaceChildren();
                                                firstBody.replaceChildren();
                                                let currentPage = firstPage;
                                                let currentBody = firstBody;

                                                const createContinuationPage = () => {
                                                    const nextPage = pageTemplate.cloneNode(true);
                                                    const pageNumber = printPages.length + 1;
                                                    nextPage.querySelector('h1').textContent = 'Накладная на перемещение №' + pageNumber;
                                                    nextPage.querySelector('.print-header').classList.add('print-header-continuation');
                                                    document.body.appendChild(nextPage);
                                                    printPages.push(nextPage);
                                                    currentPage = nextPage;
                                                    currentBody = nextPage.querySelector('tbody');
                                                };

                                                if (!fitsHalfPage) {
                                                    for (const row of productRows) {
                                                        currentBody.appendChild(row);
                                                        if (currentPage.querySelector('.print-content').getBoundingClientRect().height > usablePageHeight) {
                                                            currentBody.removeChild(row);
                                                            if (currentBody.children.length > 0) createContinuationPage();
                                                            currentBody.appendChild(row);
                                                        }
                                                    }

                                                    currentBody.appendChild(summaryRow);
                                                    if (currentPage.querySelector('.print-content').getBoundingClientRect().height > usablePageHeight) {
                                                        currentBody.removeChild(summaryRow);
                                                        createContinuationPage();
                                                        currentBody.appendChild(summaryRow);
                                                    }
                                                } else {
                                                    currentBody.append(...productRows, summaryRow);
                                                }

                                                let bottomOffset = 0;
                                                if (fitsHalfPage && '${position}' === 'bottom') {
                                                    const availableGap = Math.max(0, lowerHalfCapacity - initialContentHeight);
                                                    bottomOffset = halfPageHeight + Math.min(desiredBottomGap, availableGap);
                                                    firstPage.querySelector('.print-offset').style.height = bottomOffset + 'px';
                                                }

                                                const lastPage = printPages[printPages.length - 1];
                                                const lastContent = lastPage.querySelector('.print-content');
                                                let lastTargetHeight = lastContent.getBoundingClientRect().height <= halfPageHeight
                                                    ? halfPageHeight
                                                    : usablePageHeight;
                                                if (printPages.length === 1 && fitsHalfPage && '${position}' === 'bottom') {
                                                    lastTargetHeight = usablePageHeight - bottomOffset;
                                                }
                                                const lastBody = lastPage.querySelector('tbody');

                                                while (lastContent.getBoundingClientRect().height < lastTargetHeight) {
                                                    const blankRow = document.createElement('tr');
                                                    blankRow.innerHTML = '<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>';
                                                    lastBody.appendChild(blankRow);
                                                    if (lastContent.getBoundingClientRect().height > lastTargetHeight) {
                                                        blankRow.remove();
                                                        break;
                                                    }
                                                }

                                                lastPage.style.breakAfter = 'auto';
                                                lastPage.style.pageBreakAfter = 'auto';

            window.focus();
            window.print();
            setTimeout(() => {
                            if (window.frameElement) window.frameElement.remove();
            }, 1000);
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

async function exportRevisionResults() {
    if (getSelectedRevisionResults().length === 0) return

    const data = [
        ['Название', 'Ак', 'Разница', 'Цена', 'Штрихкод', 'Факт'],
        ...getSortedRevisionResults().map(product => {
            const storeStock = parseStock(product.storeStock) ?? 0
            if (!product.isScanned) {
                return [
                    product.name,
                    storeStock,
                    '',
                    formatPrice(product.salePrice ?? product.price ?? ''),
                    product.barcode,
                    '',
                ]
            }

            const factQuantity = parseStock(product.factQuantity) ?? 0
            const difference = factQuantity - storeStock
            return [
                product.name,
                storeStock,
                difference === 0 ? '' : `${difference > 0 ? '+' : ''}${formatStock(difference)}`,
                formatPrice(product.salePrice ?? product.price ?? ''),
                product.barcode,
                factQuantity,
            ]
        }),
    ]

    const worksheet = XLSX.utils.aoa_to_sheet(data)
    getSortedRevisionResults().forEach((product, index) => {
        const storeStock = parseStock(product.storeStock) ?? 0
        const factQuantity = parseStock(product.factQuantity) ?? 0
        const isDifferent = product.isScanned && factQuantity !== storeStock
        const fillColor = !product.isScanned ? 'E7E7E7' : isDifferent ? 'F4CCCC' : ''
        if (!fillColor) return

        for (let column = 0; column < data[0].length; column++) {
            const cell = worksheet[XLSX.utils.encode_cell({ r: index + 1, c: column })]
            if (!cell) continue
            cell.s = {
                fill: {
                    patternType: 'solid',
                    fgColor: { rgb: fillColor },
                },
            }
        }
    })
    worksheet['!cols'] = [{ wch: 40 }, { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 24 }, { wch: 12 }]
    const exportWorkbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(exportWorkbook, worksheet, 'Ревизия')
    const fileName = `Ревизия ${formatExportDate()}.xlsx`

    if (typeof window.showSaveFilePicker !== 'function') {
        XLSX.writeFile(exportWorkbook, fileName)
        return
    }

    try {
        const fileHandle = await window.showSaveFilePicker({
            suggestedName: fileName.replace(/\.xlsx$/i, ''),
            types: [
                {
                    description: 'Книга Excel',
                    accept: {
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                    },
                },
            ],
        })
        const fileContents = XLSX.write(exportWorkbook, { bookType: 'xlsx', type: 'array' })
        const writable = await fileHandle.createWritable()
        await writable.write(fileContents)
        await writable.close()
    } catch (error) {
        if (error.name === 'AbortError') return
        setError(`Не удалось сохранить файл ревизии: ${error.message}`)
    }
}

function exportResults() {
    const isAllStockList = activeList === 'all-stock'
    const isSelectiveAllStockExport = isAllStockList && isSelectingAllStockForExport
    const isCentralList = activeList === 'central' || activeList === 'central-exceptions'
    const isCentralExceptionsList = activeList === 'central-exceptions'
    const isRestockList = activeList === 'restock' || activeList === 'exceptions'
    const isExceptionsList = activeList === 'exceptions'
    const currentResults = isAllStockList
        ? allStockResults
        : isCentralList
          ? isCentralExceptionsList
              ? excludedCentralResults
              : results
          : isExceptionsList
            ? excludedRestockResults
            : restockResults
    const activeResults = isSelectiveAllStockExport
        ? currentResults.filter(product => selectedAllStockBarcodes.has(normalizeBarcode(product.barcode)))
        : currentResults
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
        ? isSelectiveAllStockExport
            ? 'Выборочные_остатки'
            : 'Все_остатки'
        : isCentralList
          ? isCentralExceptionsList
              ? 'Исключения_центрального_склада'
              : 'Товары_только_на_центральном'
          : isExceptionsList
            ? 'Исключения_для_пополнения'
            : 'Нужно_пополнить'

    XLSX.writeFile(exportWorkbook, `${exportName} ${formatExportDate()}.xlsx`)
    if (isSelectiveAllStockExport) {
        isSelectingAllStockForExport = false
        selectedAllStockBarcodes.clear()
        allStockSelectionAnchor = ''
        allStockSelectiveExportLink.textContent = 'Выборочный экспорт'
        renderResults()
        allStockTableWrapper.focus({ preventScroll: true })
    }
}

function setError(message) {
    statusEl.textContent = message
    statusEl.className = 'status error'
}
