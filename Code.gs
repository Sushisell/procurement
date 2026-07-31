const CONFIG = {
  masterSpreadsheetId: '12bLLftMzcRzUZBA_BDeDNUZ7M7jywTrL3d2rvhukhqI',
  sheets: {
    prices: 'Цены',
    template: 'Шаблон',
    suppliers: 'О поставщике',
    employeesCandidates: ['Сотрудники', 'Инфо'],
  },
  branchInfo: {
    employeeBranchCol: 1,
    employeeNameCol: 2,
    branchNameCol: 4,
    branchUrlCol: 5,
  },
  branchSheet: {
    productCol: 3,
    priceCol: 4,
    unitCol: 5,
    firstDateCol: 6,
    initiatorRow: 3,
    totalRow: 4,
    firstProductRow: 5,
    supplierMarkerColor: '#ffffff',
    obsoleteRowColor: '#d9d9d9',
  },
};

function doGet(event) {
  const action = clean_(event && event.parameter && event.parameter.action);
  if (action) return handleApiRequest_(action, null);

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Заявка на продукты')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(event) {
  try {
    const action = clean_(event && event.parameter && event.parameter.action) || 'submit';
    const body = event && event.postData && event.postData.contents;
    const payload = body ? JSON.parse(body) : {};
    return handleApiRequest_(action, payload);
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message || String(error) });
  }
}

function handleApiRequest_(action, payload) {
  try {
    if (action === 'bootstrap' || action === 'getBootstrapData') {
      return jsonResponse_({ ok: true, data: getBootstrapData() });
    }

    if (action === 'submit' || action === 'submitOrder') {
      return jsonResponse_({ ok: true, data: submitOrder(payload) });
    }

    throw new Error('Неизвестное действие API: ' + action);
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message || String(error) });
  }
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getBootstrapData() {
  const master = getMasterSpreadsheet_();
  const employeesSheet = getFirstExistingSheet_(master, CONFIG.sheets.employeesCandidates);

  return {
    branches: readBranches_(employeesSheet),
    employeesByBranch: readEmployees_(employeesSheet),
    suppliers: readSupplierInfo_(master.getSheetByName(CONFIG.sheets.suppliers)),
    catalog: readCatalog_(master),
    today: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
  };
}

function submitOrder(payload) {
  validatePayload_(payload);

  const master = getMasterSpreadsheet_();
  const employeesSheet = getFirstExistingSheet_(master, CONFIG.sheets.employeesCandidates);
  const branchMap = readBranches_(employeesSheet).reduce((acc, branch) => {
    acc[normalizeKey_(branch.name)] = branch;
    return acc;
  }, {});
  const branch = branchMap[normalizeKey_(payload.branch)];

  if (!branch) throw new Error('Не найден филиал: ' + payload.branch);

  const date = parseDate_(payload.date);
  const branchSpreadsheet = branch.url ? SpreadsheetApp.openByUrl(branch.url) : master;
  const monthSheet = getOrCreateMonthSheet_(branchSpreadsheet, master, date);
  const templateProducts = readTemplateProducts_(master.getSheetByName(CONFIG.sheets.template), getKnownSuppliers_(master));

  syncBranchSheetWithTemplate_(monthSheet, templateProducts);

  const dateColumn = ensureDateColumn_(monthSheet, date);
  clearPreviousOrder_(monthSheet, dateColumn);

  const rowByProduct = buildProductRowIndex_(monthSheet);
  let total = 0;
  const warnings = [];

  payload.items.forEach((item) => {
    const quantity = toNumber_(item.quantity);
    if (!quantity) return;

    const row = rowByProduct[normalizeKey_(item.product)];
    const price = toNumber_(item.price);
    const lineTotal = quantity * price;
    total += lineTotal;

    if (!row) {
      warnings.push('Товар не найден в таблице филиала и не был записан: ' + item.product);
      return;
    }

    monthSheet.getRange(row, dateColumn).setValue(quantity);
  });

  monthSheet.getRange(CONFIG.branchSheet.initiatorRow, dateColumn).setValue(payload.employee);
  monthSheet.getRange(CONFIG.branchSheet.totalRow, dateColumn).setValue(total);

  return {
    ok: true,
    total,
    warnings,
    spreadsheetUrl: branch.url || master.getUrl(),
    sheetName: monthSheet.getName(),
  };
}

function getMasterSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID')
    || CONFIG.masterSpreadsheetId;
  return SpreadsheetApp.openById(id);
}

function getFirstExistingSheet_(spreadsheet, names) {
  for (const name of names) {
    const sheet = spreadsheet.getSheetByName(name);
    if (sheet) return sheet;
  }
  throw new Error('Не найден лист: ' + names.join(' или '));
}

function readBranches_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const branches = [];
  const seen = {};

  values.slice(1).forEach((row) => {
    const name = clean_(row[CONFIG.branchInfo.branchNameCol - 1])
      || clean_(row[CONFIG.branchInfo.employeeBranchCol - 1]);
    const url = clean_(row[CONFIG.branchInfo.branchUrlCol - 1]);
    if (!name || seen[normalizeKey_(name)]) return;
    seen[normalizeKey_(name)] = true;
    branches.push({ name, url });
  });

  return branches.sort((a, b) => a.name.localeCompare(b.name, 'ru'));
}

function readEmployees_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const result = {};

  values.slice(1).forEach((row) => {
    const branch = clean_(row[CONFIG.branchInfo.employeeBranchCol - 1]);
    const name = clean_(row[CONFIG.branchInfo.employeeNameCol - 1]);
    if (!branch || !name) return;
    if (!result[branch]) result[branch] = [];
    result[branch].push(name);
  });

  Object.keys(result).forEach((branch) => {
    result[branch] = Array.from(new Set(result[branch])).sort((a, b) => a.localeCompare(b, 'ru'));
  });

  return result;
}

function readSupplierInfo_(sheet) {
  if (!sheet) return {};

  const values = sheet.getDataRange().getDisplayValues();
  const result = {};
  values.slice(1).forEach((row) => {
    const name = clean_(row[0]);
    if (!name) return;
    result[name] = {
      name,
      contact: clean_(row[1]),
      minimum: clean_(row[2]),
      unit: clean_(row[3]),
      orderInfo: clean_(row[4] || row[3]),
    };
  });
  return result;
}

function readCatalog_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.sheets.prices);
  if (!sheet) return [];

  const values = sheet.getDataRange().getDisplayValues();
  const products = {};
  values.slice(1).forEach((row) => {
    const supplier = clean_(row[0]);
    const product = clean_(row[1]);
    if (!supplier || !product) return;
    const productKey = normalizeKey_(supplier + '|' + product);
    if (!products[productKey]) {
      products[productKey] = {
        supplier,
        product,
        photo: clean_(row[2]),
        comment: clean_(row[3]),
        unit: clean_(row[5]),
        prices: [],
      };
    }
    products[productKey].prices.push({
      price: toNumber_(row[4]),
      effectiveDate: normalizeSheetDate_(row[6]),
    });
  });

  return groupProductsBySupplier_(Object.keys(products).map((key) => {
    const item = products[key];
    item.prices.sort((a, b) => a.effectiveDate.localeCompare(b.effectiveDate));
    item.price = item.prices[item.prices.length - 1]?.price || 0;
    return item;
  }));
}

function normalizeSheetDate_(value) {
  const text = clean_(value);
  if (!text) return '';
  const match = text.match(/^(\d{1,2})[.\/-](\d{1,2})(?:[.\/-](\d{2,4}))?$/);
  if (!match) return text;
  let year = Number(match[3] || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy'));
  if (year < 100) year += 2000;
  return [year, String(Number(match[2])).padStart(2, '0'), String(Number(match[1])).padStart(2, '0')].join('-');
}

function getKnownSuppliers_(spreadsheet) {
  const result = {};

  const supplierInfoSheet = spreadsheet.getSheetByName(CONFIG.sheets.suppliers);
  if (supplierInfoSheet) {
    supplierInfoSheet.getDataRange().getDisplayValues().slice(1).forEach((row) => {
      const name = clean_(row[0]);
      if (name) result[normalizeKey_(name)] = name;
    });
  }

  const pricesSheet = spreadsheet.getSheetByName(CONFIG.sheets.prices);
  if (pricesSheet) {
    pricesSheet.getDataRange().getDisplayValues().slice(1).forEach((row) => {
      const name = clean_(row[0]);
      if (name) result[normalizeKey_(name)] = name;
    });
  }

  return result;
}

function readTemplateProducts_(sheet, knownSuppliers) {
  if (!sheet) return [];
  knownSuppliers = knownSuppliers || {};

  const values = sheet.getDataRange().getDisplayValues();
  const products = [];
  let currentSupplier = '';

  values.forEach((row, index) => {
    const cells = row.map(clean_);
    const nonEmptyCells = cells.filter(Boolean);
    if (!nonEmptyCells.length) return;

    const knownSupplierCell = nonEmptyCells.find((cell) => Boolean(knownSuppliers[normalizeKey_(cell)]));
    const firstCell = cells[0] || nonEmptyCells[0];
    const knownSupplier = knownSupplierCell ? knownSuppliers[normalizeKey_(knownSupplierCell)] : '';
    const isHeader = Boolean(knownSupplier) || (index === 0 && nonEmptyCells.length === 1);
    if (isHeader) {
      currentSupplier = knownSupplier || firstCell;
      return;
    }

    if (!currentSupplier) return;

    const productCellIndex = getTemplateProductCellIndex_(row);
    const product = clean_(row[productCellIndex]);
    if (!product) return;

    products.push({
      supplier: currentSupplier,
      product,
      price: getTemplatePrice_(row, productCellIndex),
      unit: getTemplateUnit_(row, productCellIndex),
    });
  });

  return products;
}


function getTemplateProductCellIndex_(row) {
  const preferredIndexes = [0, CONFIG.branchSheet.productCol - 1];
  for (const index of preferredIndexes) {
    if (clean_(row[index])) return index;
  }

  return row.findIndex((cell) => Boolean(clean_(cell)));
}

function getTemplatePrice_(row, productCellIndex) {
  const preferredIndexes = [productCellIndex + 1, CONFIG.branchSheet.priceCol - 1, 1];
  for (const index of preferredIndexes) {
    const price = toNumber_(row[index]);
    if (price) return price;
  }

  for (let index = productCellIndex + 1; index < row.length; index += 1) {
    const price = toNumber_(row[index]);
    if (price) return price;
  }

  return 0;
}

function getTemplateUnit_(row, productCellIndex) {
  const preferredIndexes = [productCellIndex + 2, CONFIG.branchSheet.unitCol - 1, 2];
  for (const index of preferredIndexes) {
    const value = clean_(row[index]);
    if (value) return value;
  }

  return '';
}

function groupProductsBySupplier_(rows) {
  const map = {};
  rows.forEach((row) => {
    if (!map[row.supplier]) {
      map[row.supplier] = { supplier: row.supplier, products: [] };
    }
    map[row.supplier].products.push({
      product: row.product,
      price: row.price,
      unit: row.unit,
      photo: row.photo || '',
      comment: row.comment || '',
      prices: row.prices || [],
    });
  });
  return Object.keys(map).map((supplier) => map[supplier]);
}

function getOrCreateMonthSheet_(branchSpreadsheet, masterSpreadsheet, date) {
  const name = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM.yy');
  let sheet = branchSpreadsheet.getSheetByName(name);
  if (sheet) return sheet;

  sheet = branchSpreadsheet.insertSheet(name);
  const templateProducts = readTemplateProducts_(masterSpreadsheet.getSheetByName(CONFIG.sheets.template), getKnownSuppliers_(masterSpreadsheet));
  buildMonthSheet_(sheet, templateProducts, date);
  return sheet;
}

function buildMonthSheet_(sheet, templateProducts, date) {
  sheet.clear();
  sheet.getRange(3, CONFIG.branchSheet.productCol).setValue('Инициатор заявки').setFontWeight('bold');
  sheet.getRange(4, CONFIG.branchSheet.productCol).setValue('Сумма заявки').setFontWeight('bold');
  sheet.getRange(1, CONFIG.branchSheet.priceCol).setValue('Цена на сегодня');
  sheet.getRange(1, CONFIG.branchSheet.unitCol).setValue('Кратность. Цена за упаковку');

  const start = new Date(date.getFullYear(), date.getMonth(), 1);
  const days = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
  for (let i = 0; i < days; i += 1) {
    const current = new Date(start.getFullYear(), start.getMonth(), i + 1);
    const column = CONFIG.branchSheet.firstDateCol + i;
    sheet.getRange(1, column).setValue(dayName_(current));
    sheet.getRange(2, column).setValue(Utilities.formatDate(current, Session.getScriptTimeZone(), 'dd.MM'));
  }

  writeTemplateRows_(sheet, templateProducts);
  sheet.setFrozenColumns(CONFIG.branchSheet.unitCol);
}

function syncBranchSheetWithTemplate_(sheet, templateProducts) {
  const currentRows = buildProductRowIndex_(sheet);
  const templateKeys = {};

  templateProducts.forEach((item) => {
    templateKeys[normalizeKey_(item.product)] = true;
    const row = currentRows[normalizeKey_(item.product)];
    if (row) {
      sheet.getRange(row, CONFIG.branchSheet.priceCol).setValue(item.price || '');
      sheet.getRange(row, CONFIG.branchSheet.unitCol).setValue(item.unit || '');
      sheet.getRange(row, 1, 1, sheet.getMaxColumns()).setBackground(null);
    }
  });

  Object.keys(currentRows).forEach((key) => {
    if (!templateKeys[key]) {
      sheet.getRange(currentRows[key], 1, 1, sheet.getMaxColumns())
        .setBackground(CONFIG.branchSheet.obsoleteRowColor);
    }
  });

  const missing = templateProducts.filter((item) => !currentRows[normalizeKey_(item.product)]);
  if (missing.length) {
    appendMissingTemplateRows_(sheet, missing);
  }
}

function writeTemplateRows_(sheet, templateProducts) {
  let row = CONFIG.branchSheet.firstProductRow;
  let currentSupplier = '';

  templateProducts.forEach((item) => {
    if (item.supplier !== currentSupplier) {
      currentSupplier = item.supplier;
      sheet.getRange(row, CONFIG.branchSheet.productCol).setValue(currentSupplier).setFontWeight('bold').setFontSize(18);
      row += 1;
    }

    sheet.getRange(row, CONFIG.branchSheet.productCol).setValue(item.product);
    sheet.getRange(row, CONFIG.branchSheet.priceCol).setValue(item.price || '');
    sheet.getRange(row, CONFIG.branchSheet.unitCol).setValue(item.unit || '');
    row += 1;
  });
}

function appendMissingTemplateRows_(sheet, products) {
  let row = Math.max(sheet.getLastRow() + 1, CONFIG.branchSheet.firstProductRow);
  let currentSupplier = '';

  products.forEach((item) => {
    if (item.supplier !== currentSupplier) {
      currentSupplier = item.supplier;
      sheet.getRange(row, CONFIG.branchSheet.productCol).setValue(currentSupplier).setFontWeight('bold').setFontSize(18);
      row += 1;
    }
    sheet.getRange(row, CONFIG.branchSheet.productCol).setValue(item.product);
    sheet.getRange(row, CONFIG.branchSheet.priceCol).setValue(item.price || '');
    sheet.getRange(row, CONFIG.branchSheet.unitCol).setValue(item.unit || '');
    row += 1;
  });
}

function ensureDateColumn_(sheet, date) {
  const wanted = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM');
  const lastColumn = Math.max(sheet.getLastColumn(), CONFIG.branchSheet.firstDateCol);
  const dates = sheet.getRange(2, CONFIG.branchSheet.firstDateCol, 1, lastColumn - CONFIG.branchSheet.firstDateCol + 1).getDisplayValues()[0];
  const index = dates.findIndex((value) => clean_(value) === wanted);

  if (index >= 0) return CONFIG.branchSheet.firstDateCol + index;

  const column = lastColumn + 1;
  sheet.getRange(1, column).setValue(dayName_(date));
  sheet.getRange(2, column).setValue(wanted);
  return column;
}

function clearPreviousOrder_(sheet, dateColumn) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.branchSheet.initiatorRow) return;
  sheet.getRange(CONFIG.branchSheet.initiatorRow, dateColumn, lastRow - CONFIG.branchSheet.initiatorRow + 1, 1).clearContent();
}

function buildProductRowIndex_(sheet) {
  const lastRow = sheet.getLastRow();
  const result = {};
  if (lastRow < CONFIG.branchSheet.firstProductRow) return result;

  const height = lastRow - CONFIG.branchSheet.firstProductRow + 1;
  const products = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.productCol, height, 1).getDisplayValues();
  const prices = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.priceCol, height, 1).getDisplayValues();
  const units = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.unitCol, height, 1).getDisplayValues();

  products.forEach((row, index) => {
    const product = clean_(row[0]);
    if (!product) return;
    const isSupplierHeader = !clean_(prices[index][0]) && !clean_(units[index][0]);
    if (isSupplierHeader) return;
    result[normalizeKey_(product)] = CONFIG.branchSheet.firstProductRow + index;
  });

  return result;
}

function validatePayload_(payload) {
  if (!payload) throw new Error('Пустая заявка.');
  if (!payload.branch) throw new Error('Выберите филиал.');
  if (!payload.employee) throw new Error('Выберите ФИО.');
  if (!payload.date) throw new Error('Выберите дату.');
  if (!payload.items || !payload.items.length) throw new Error('Добавьте хотя бы одну позицию.');
}

function parseDate_(value) {
  const parts = String(value).split('-').map(Number);
  if (parts.length !== 3) throw new Error('Некорректная дата: ' + value);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function dayName_(date) {
  return ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'][date.getDay()];
}

function clean_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function normalizeKey_(value) {
  return clean_(value).toLowerCase();
}

function toNumber_(value) {
  if (typeof value === 'number') return value;
  const normalized = String(value || '').replace(/\s/g, '').replace(',', '.').replace(/[^0-9.-]/g, '');
  const number = Number(normalized);
  return Number.isFinite(number) ? number : 0;
}
