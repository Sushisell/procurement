const CONFIG = {
  timeZone: 'Asia/Krasnoyarsk',
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
    receivedColor: '#b7e1cd',
    partiallyReceivedColor: '#ffe599',
    notReceivedColor: '#f4cccc',
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

    if (action === 'latestOrder' || action === 'getLatestOrder') {
      return jsonResponse_({ ok: true, data: getLatestOrder(payload) });
    }

    if (action === 'orders' || action === 'getOrders') {
      return jsonResponse_({ ok: true, data: getOrders(payload) });
    }

    if (action === 'receive' || action === 'receiveOrder') {
      return jsonResponse_({ ok: true, data: receiveOrder(payload) });
    }

    if (action === 'updateOrder' || action === 'updateSubmittedOrder') {
      return jsonResponse_({ ok: true, data: updateSubmittedOrder(payload) });
    }

    if (action === 'markOrdered' || action === 'markOrderOrdered') {
      return jsonResponse_({ ok: true, data: markOrderOrdered(payload) });
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
  const suppliers = readSupplierInfo_(master.getSheetByName(CONFIG.sheets.suppliers));

  return {
    branches: readBranches_(employeesSheet),
    employeesByBranch: readEmployees_(employeesSheet),
    suppliers,
    catalog: readCatalog_(master),
    today: Utilities.formatDate(new Date(), CONFIG.timeZone, 'yyyy-MM-dd'),
  };
}

function submitOrder(payload) {
  validatePayload_(payload);

  const master = getMasterSpreadsheet_();
  const suppliers = readSupplierInfo_(master.getSheetByName(CONFIG.sheets.suppliers));
  const supplier = suppliers[payload.supplier];
  if (!supplier) throw new Error('Не найден поставщик: ' + payload.supplier);
  const availability = getDeliveryAvailability_(supplier.deliveryRules, payload.date);
  if (supplier.deliveryRules.length && !availability.available) {
    throw new Error('Поставщик «' + payload.supplier + '» не принимает сейчас заявку на '
      + formatRussianDate_(payload.date) + '. Выберите доступную дату поставки.');
  }
  if (!supplier.deliveryRules.length && payload.date < supplier.earliestDeliveryDate) {
    throw new Error('Поставщик «' + payload.supplier + '» уже не принимает заявку на текущий день.');
  }
  validateSupplierMinimum_(payload, supplier);

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

  const dateColumn = ensureSubmissionColumn_(monthSheet, date);

  const rowByProduct = buildProductRowIndex_(monthSheet);
  // Повторная отправка поставщику является редактированием его части заявки:
  // удаляем старые количества этого поставщика, не затрагивая остальных.
  getSupplierProductRows_(monthSheet, payload.supplier).forEach((row) => {
    monthSheet.getRange(row, dateColumn).clearContent().clearNote().setBackground(null);
  });
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

    monthSheet.getRange(row, dateColumn)
      .setValue(quantity)
      .setBackground(CONFIG.branchSheet.notReceivedColor)
      .setNote('Получение не отмечено.');
  });

  monthSheet.getRange(CONFIG.branchSheet.initiatorRow, dateColumn).setValue(payload.employee);
  // Общую сумму по колонке намеренно не заполняем: заявка отправляется
  // отдельно поставщику, поэтому сумма записывается напротив его заголовка.
  monthSheet.getRange(CONFIG.branchSheet.totalRow, dateColumn).clearContent();
  const supplierRow = findSupplierRow_(monthSheet, payload.supplier);
  if (supplierRow) monthSheet.getRange(supplierRow, dateColumn).setValue(total);

  return {
    ok: true,
    total,
    warnings,
    spreadsheetUrl: branch.url || master.getUrl(),
    sheetName: monthSheet.getName(),
  };
}

function getLatestOrder(payload) {
  const orders = getOrders(payload);
  return orders.length ? orders[0] : null;
}

function getOrders(payload) {
  const context = getBranchContext_(payload && payload.branch);
  const orders = findOrders_(context.spreadsheet, payload && payload.date);
  orders.forEach((order) => {
    order.receipt = readReceiptFromOrderCells_(context.spreadsheet, order);
    if (!order.receipt && order.status === 'submitted') markOrderAsNotReceived_(context.spreadsheet, order);
  });
  return orders;
}

function receiveOrder(payload) {
  if (!payload || !payload.branch) throw new Error('Выберите филиал.');
  if (!payload.employee) throw new Error('Выберите ФИО принявшего товар.');
  if (!payload.orderId) throw new Error('Не найдена заявка для приёмки.');
  if (!Array.isArray(payload.items) || !payload.items.length) {
    throw new Error('Укажите фактически полученные товары.');
  }

  const context = getBranchContext_(payload.branch);
  const employees = readEmployees_(context.employeesSheet)[payload.branch] || [];
  if (!employees.some((name) => normalizeKey_(name) === normalizeKey_(payload.employee))) {
    throw new Error('Сотрудник не найден в выбранном филиале.');
  }

  const order = findOrderById_(context.spreadsheet, payload.orderId);
  if (!order) throw new Error('Заявка больше не найдена. Обновите данные.');
  if (order.status !== 'ordered') throw new Error('Сначала отметьте, что заявка заказана.');

  const expected = {};
  order.items.forEach((item) => { expected[normalizeKey_(item.product)] = item; });
  payload.items.forEach((item) => {
    const orderedProduct = clean_(item.orderedProduct);
    const ordered = expected[normalizeKey_(orderedProduct)];
    if (!ordered) throw new Error('Товар не относится к выбранной заявке: ' + orderedProduct);
    const actualProduct = clean_(item.actualProduct);
    const actualQuantity = toNumber_(item.actualQuantity);
    if (!actualProduct && actualQuantity) throw new Error('Укажите фактически приехавший товар.');
    if (actualQuantity < 0) throw new Error('Фактическое количество не может быть отрицательным.');
  });

  applyReceiptFormatting_(context.spreadsheet, order, payload.items, payload.employee, new Date());

  return { ok: true, orderId: order.id, receivedAt: Utilities.formatDate(new Date(), CONFIG.timeZone, "yyyy-MM-dd'T'HH:mm:ss") };
}

function updateSubmittedOrder(payload) {
  const context = getBranchContext_(payload && payload.branch);
  const order = findOrderById_(context.spreadsheet, payload && payload.orderId);
  if (!order) throw new Error('Заявка больше не найдена.');
  if (order.status === 'ordered') throw new Error('Заказанную заявку уже нельзя корректировать.');
  if (!Array.isArray(payload.items)) throw new Error('Не переданы позиции заявки.');
  const sheetAndColumn = getOrderSheetAndColumn_(context.spreadsheet, order.id);
  const rowByProduct = buildProductRowIndex_(sheetAndColumn.sheet);
  const allowed = {};
  order.items.forEach((item) => { allowed[normalizeKey_(item.product)] = true; });
  if (!payload.items.some((item) => toNumber_(item.quantity) > 0)) {
    throw new Error('В заявке должна остаться хотя бы одна позиция.');
  }
  payload.items.forEach((item) => {
    if (!allowed[normalizeKey_(item.product)]) throw new Error('Нельзя добавить товар из другой заявки.');
    const row = rowByProduct[normalizeKey_(item.product)];
    if (!row) return;
    const quantity = Math.max(0, toNumber_(item.quantity));
    sheetAndColumn.sheet.getRange(row, sheetAndColumn.column)
      .setValue(quantity || '')
      .setBackground(CONFIG.branchSheet.notReceivedColor)
      .setNote('Статус: отправлено. Получение не отмечено.');
  });
  const newTotal = payload.items.reduce((sum, item) => {
    const row = rowByProduct[normalizeKey_(item.product)];
    const price = row ? toNumber_(sheetAndColumn.sheet.getRange(row, CONFIG.branchSheet.priceCol).getDisplayValue()) : 0;
    return sum + toNumber_(item.quantity) * price;
  }, 0);
  const supplierRow = findSupplierRow_(sheetAndColumn.sheet, order.supplier);
  if (supplierRow) sheetAndColumn.sheet.getRange(supplierRow, sheetAndColumn.column).setValue(newTotal);
  return findOrderById_(context.spreadsheet, order.id);
}

function markOrderOrdered(payload) {
  const context = getBranchContext_(payload && payload.branch);
  const order = findOrderById_(context.spreadsheet, payload && payload.orderId);
  if (!order) throw new Error('Заявка больше не найдена.');
  const sheetAndColumn = getOrderSheetAndColumn_(context.spreadsheet, order.id);
  const rowByProduct = buildProductRowIndex_(sheetAndColumn.sheet);
  order.items.forEach((item) => {
    const row = rowByProduct[normalizeKey_(item.product)];
    if (!row) return;
    sheetAndColumn.sheet.getRange(row, sheetAndColumn.column)
      .setBackground(CONFIG.branchSheet.notReceivedColor)
      .setNote('Статус: заказано. Получение не отмечено.');
  });
  return findOrderById_(context.spreadsheet, order.id);
}

function getBranchContext_(branchName) {
  if (!branchName) throw new Error('Выберите филиал.');
  const master = getMasterSpreadsheet_();
  const employeesSheet = getFirstExistingSheet_(master, CONFIG.sheets.employeesCandidates);
  const branch = readBranches_(employeesSheet).find((item) => normalizeKey_(item.name) === normalizeKey_(branchName));
  if (!branch) throw new Error('Не найден филиал: ' + branchName);
  return {
    master,
    employeesSheet,
    branch,
    spreadsheet: branch.url ? SpreadsheetApp.openByUrl(branch.url) : master,
  };
}

function findLatestOrder_(spreadsheet, deliveryDate) {
  const orders = findOrders_(spreadsheet, deliveryDate);
  return orders.length ? orders[0] : null;
}

function findOrders_(spreadsheet, deliveryDate) {
  const candidates = [];
  spreadsheet.getSheets().forEach((sheet) => {
    if (!/^\d{2}\.\d{2}$/.test(sheet.getName()) || sheet.getLastColumn() < CONFIG.branchSheet.firstDateCol) return;
    const year = 2000 + Number(sheet.getName().slice(3));
    const month = Number(sheet.getName().slice(0, 2));
    const lastColumn = sheet.getLastColumn();
    const dates = sheet.getRange(2, CONFIG.branchSheet.firstDateCol, 1,
      lastColumn - CONFIG.branchSheet.firstDateCol + 1).getDisplayValues()[0];
    const initiators = sheet.getRange(CONFIG.branchSheet.initiatorRow, CONFIG.branchSheet.firstDateCol, 1,
      dates.length).getDisplayValues()[0];
    dates.forEach((date, index) => {
      if (!clean_(date) || !clean_(initiators[index])) return;
      const parts = clean_(date).split('.').map(Number);
      const timestamp = Date.UTC(year, month - 1, parts[0]);
      const isoDate = '20' + sheet.getName().slice(3) + '-' + sheet.getName().slice(0, 2)
        + '-' + String(parts[0]).padStart(2, '0');
      if (!deliveryDate || deliveryDate === isoDate) {
        candidates.push({ sheet, column: CONFIG.branchSheet.firstDateCol + index, timestamp });
      }
    });
  });
  candidates.sort((a, b) => b.timestamp - a.timestamp || b.column - a.column);
  const orders = [];
  candidates.forEach((candidate) => {
    readOrdersFromColumn_(candidate.sheet, candidate.column).forEach((order) => orders.push(order));
  });
  return orders;
}

function findOrderById_(spreadsheet, orderId) {
  const match = String(orderId).match(/^(.+):(\d+)(?::(.+))?$/);
  if (!match) return null;
  const sheet = spreadsheet.getSheetByName(match[1]);
  const column = Number(match[2]);
  if (!sheet || column < CONFIG.branchSheet.firstDateCol || column > sheet.getLastColumn()) return null;
  const orders = readOrdersFromColumn_(sheet, column);
  if (!match[3]) return orders.length ? orders[0] : null;
  const supplier = decodeURIComponent(match[3]);
  return orders.find((order) => normalizeKey_(order.supplier) === normalizeKey_(supplier)) || null;
}

function readOrderFromColumn_(sheet, column) {
  const orders = readOrdersFromColumn_(sheet, column);
  return orders.length ? orders[0] : null;
}

function readOrdersFromColumn_(sheet, column) {
  const date = clean_(sheet.getRange(2, column).getDisplayValue());
  const initiator = clean_(sheet.getRange(CONFIG.branchSheet.initiatorRow, column).getDisplayValue());
  if (!date || !initiator) return [];
  const lastRow = sheet.getLastRow();
  const height = lastRow - CONFIG.branchSheet.firstProductRow + 1;
  const names = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.productCol, height, 1).getDisplayValues();
  const quantities = sheet.getRange(CONFIG.branchSheet.firstProductRow, column, height, 1).getDisplayValues();
  const prices = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.priceCol, height, 1).getDisplayValues();
  const units = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.unitCol, height, 1).getDisplayValues();
  const fontSizes = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.productCol, height, 1).getFontSizes();
  let currentSupplier = '';
  const itemsBySupplier = {};
  names.forEach((row, index) => {
    const name = clean_(row[0]);
    const isHeader = name && !clean_(prices[index][0]) && !clean_(units[index][0]) && Number(fontSizes[index][0]) >= 16;
    if (isHeader) {
      currentSupplier = name;
      if (!itemsBySupplier[currentSupplier]) itemsBySupplier[currentSupplier] = [];
      return;
    }
    const quantity = toNumber_(quantities[index][0]);
    if (name && quantity) {
      if (!itemsBySupplier[currentSupplier]) itemsBySupplier[currentSupplier] = [];
      itemsBySupplier[currentSupplier].push({ product: name, quantity, unit: clean_(units[index][0]) });
    }
  });
  const dateParts = date.split('.');
  return Object.keys(itemsBySupplier).filter((supplier) => supplier && itemsBySupplier[supplier].length).map((supplier) => {
    const result = {
    id: sheet.getName() + ':' + column + ':' + encodeURIComponent(supplier),
    date: '20' + sheet.getName().slice(3) + '-' + sheet.getName().slice(0, 2) + '-' + dateParts[0].padStart(2, '0'),
    supplier,
    initiator,
    items: itemsBySupplier[supplier],
    };
    result.status = readOrderStatus_(sheet, column, result.items);
    return result;
  });
}

function readOrderStatus_(sheet, column, items) {
  const rowByProduct = buildProductRowIndex_(sheet);
  return items.some((item) => {
    const row = rowByProduct[normalizeKey_(item.product)];
    const note = row ? clean_(sheet.getRange(row, column).getNote()) : '';
    return /Статус:\s*заказано/i.test(note) || /^Получено/i.test(note);
  }) ? 'ordered' : 'submitted';
}

function readReceiptFromOrderCells_(spreadsheet, order) {
  const sheetAndColumn = getOrderSheetAndColumn_(spreadsheet, order.id);
  if (!sheetAndColumn) return null;
  const rowByProduct = buildProductRowIndex_(sheetAndColumn.sheet);
  const items = [];
  let receivedAt = '';
  let employee = '';
  let hasSavedReceipt = false;

  order.items.forEach((ordered) => {
    const row = rowByProduct[normalizeKey_(ordered.product)];
    if (!row) return;
    const range = sheetAndColumn.sheet.getRange(row, sheetAndColumn.column);
    const note = String(range.getNote() || '');
    const parsed = parseReceiptNote_(note);
    if (parsed.events.length) hasSavedReceipt = true;
    receivedAt = parsed.events.length ? parsed.events[parsed.events.length - 1].date : receivedAt;
    employee = parsed.events.length ? parsed.events[parsed.events.length - 1].employee : employee;
    items.push({
      orderedProduct: ordered.product,
      actualProduct: parsed.actualProduct || ordered.product,
      actualQuantity: parsed.total,
      remainingQuantity: Math.max(0, ordered.quantity - parsed.total),
    });
  });

  return hasSavedReceipt ? { receivedAt, employee, items } : null;
}

function markOrderAsNotReceived_(spreadsheet, order) {
  const sheetAndColumn = getOrderSheetAndColumn_(spreadsheet, order.id);
  if (!sheetAndColumn) return;
  const rowByProduct = buildProductRowIndex_(sheetAndColumn.sheet);
  order.items.forEach((item) => {
    const row = rowByProduct[normalizeKey_(item.product)];
    if (!row) return;
    sheetAndColumn.sheet.getRange(row, sheetAndColumn.column)
      .setBackground(CONFIG.branchSheet.notReceivedColor)
      .setNote('Получение не отмечено.');
  });
}

function applyReceiptFormatting_(spreadsheet, order, receivedItems, employee, receivedAt) {
  const sheetAndColumn = getOrderSheetAndColumn_(spreadsheet, order.id);
  if (!sheetAndColumn) return;
  const rowByProduct = buildProductRowIndex_(sheetAndColumn.sheet);
  const receivedByProduct = {};
  receivedItems.forEach((item) => {
    receivedByProduct[normalizeKey_(item.orderedProduct)] = item;
  });
  const dateText = Utilities.formatDate(receivedAt, CONFIG.timeZone, 'dd.MM.yyyy');

  order.items.forEach((ordered) => {
    const row = rowByProduct[normalizeKey_(ordered.product)];
    if (!row) return;
    const received = receivedByProduct[normalizeKey_(ordered.product)] || {};
    const addedQuantity = toNumber_(received.actualQuantity);
    const actualProduct = clean_(received.actualProduct) || ordered.product;
    const range = sheetAndColumn.sheet.getRange(row, sheetAndColumn.column);
    const previous = parseReceiptNote_(range.getNote());
    const actualQuantity = previous.total + addedQuantity;
    let color = CONFIG.branchSheet.notReceivedColor;
    const event = addedQuantity > 0
      ? 'Получено ' + formatReceiptQuantity_(addedQuantity, ordered.unit)
        + ' от ' + dateText + ' сотрудником ' + employee
      : '';
    const previousLines = previous.events.map((item) => item.text);
    if (previous.actualProduct) previousLines.push('Фактически приехал товар: ' + previous.actualProduct + '.');
    let note = ['Статус: заказано.'].concat(previousLines).concat(event || []).join('\n');
    if (actualQuantity >= ordered.quantity) {
      color = CONFIG.branchSheet.receivedColor;
    } else if (actualQuantity > 0) {
      color = CONFIG.branchSheet.partiallyReceivedColor;
    } else {
      note += '\nПолучение не отмечено.';
    }
    if (addedQuantity > 0 && normalizeKey_(actualProduct) !== normalizeKey_(ordered.product)) {
      note += '\nФактически приехал товар: ' + actualProduct + '.';
    }
    range.setBackground(color).setNote(note);
  });
}

function parseReceiptNote_(value) {
  const note = String(value || '');
  const events = [];
  note.split(/\r?\n/).forEach((line) => {
    const match = clean_(line).match(/^Получено(?: полностью:)?\s*([\d.,]+)(?:\s+[^\s]+)?(?:\s+из\s+[\d.,]+(?:\s+[^\s]+)?)?\s+от\s+(\d{2}\.\d{2}\.\d{4})\s+сотрудником\s+(.+)$/i);
    if (match) events.push({ quantity: toNumber_(match[1]), date: match[2], employee: clean_(match[3]), text: clean_(line) });
  });
  const actualProducts = Array.from(note.matchAll(/Фактически приехал товар:\s*(.+?)\./gi));
  return {
    events,
    total: events.reduce((sum, item) => sum + item.quantity, 0),
    actualProduct: actualProducts.length ? clean_(actualProducts[actualProducts.length - 1][1]) : '',
  };
}

function getOrderSheetAndColumn_(spreadsheet, orderId) {
  const match = String(orderId).match(/^(.+):(\d+)(?::.+)?$/);
  if (!match) return null;
  const sheet = spreadsheet.getSheetByName(match[1]);
  const column = Number(match[2]);
  return sheet && column >= CONFIG.branchSheet.firstDateCol && column <= sheet.getLastColumn()
    ? { sheet, column }
    : null;
}

function formatReceiptQuantity_(quantity, unit) {
  const value = Number(quantity);
  const formatted = Number.isInteger(value) ? String(value) : String(value).replace('.', ',');
  return formatted + (clean_(unit) ? ' ' + clean_(unit) : ' шт');
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
    if (!result[name]) {
      result[name] = { name, contact: '', minimum: '', unit: '', orderInfo: '', cutoffTime: '', deliveryRules: [] };
    }
    const supplier = result[name];
    supplier.contact = supplier.contact || clean_(row[1]);
    supplier.minimum = supplier.minimum || clean_(row[2]);
    supplier.unit = supplier.unit || clean_(row[3]);
    supplier.orderInfo = supplier.orderInfo || clean_(row[4] || row[3]);

    const rule = parseDeliveryRule_(row[5], row[6], row[7]);
    if (rule) supplier.deliveryRules.push(rule);
    supplier.cutoffTime = supplier.cutoffTime || normalizeCutoffTime_(row[5] || row[4]);
  });

  Object.keys(result).forEach((name) => {
    const supplier = result[name];
    supplier.deliveryRules = deduplicateDeliveryRules_(supplier.deliveryRules);
    supplier.availableDeliveryDates = supplier.deliveryRules.length
      ? getAvailableDeliveryDates_(supplier.deliveryRules, 370)
      : [];
    supplier.availabilityThrough = addDaysToIsoDate_(
      Utilities.formatDate(new Date(), CONFIG.timeZone, 'yyyy-MM-dd'), 369);
    supplier.earliestDeliveryDate = supplier.availableDeliveryDates[0]
      || getEarliestDeliveryDate_(supplier.cutoffTime);
  });
  return result;
}

function parseDeliveryRule_(orderDay, cutoffValue, deliveryValue) {
  const weekday = parseRussianWeekday_(orderDay);
  const cutoffTime = normalizeCutoffTime_(cutoffValue);
  const delivery = clean_(deliveryValue).toLowerCase().replace(/ё/g, 'е');
  if (weekday === null || !cutoffTime || !delivery) return null;

  if (/тот\s*же\s*день/.test(delivery)) {
    return { orderWeekday: weekday, cutoffTime, deliveryType: 'В тот же день', deliveryOffset: 0 };
  }
  if (/следующ[а-я]*\s*день/.test(delivery)) {
    return { orderWeekday: weekday, cutoffTime, deliveryType: 'Следующий день', deliveryOffset: 1 };
  }

  const deliveryWeekday = parseRussianWeekday_(delivery.replace(/^(?:в|\s)+/i, ''));
  if (deliveryWeekday !== null) {
    return { orderWeekday: weekday, cutoffTime, deliveryType: 'В день недели', deliveryWeekday };
  }
  return null;
}

function parseRussianWeekday_(value) {
  const text = clean_(value).toLowerCase().replace(/ё/g, 'е').replace(/[^а-я]/g, '');
  const aliases = {
    'пн': 1, 'понедельник': 1, 'понедельника': 1,
    'вт': 2, 'вторник': 2, 'вторника': 2,
    'ср': 3, 'среда': 3, 'среду': 3, 'среды': 3,
    'чт': 4, 'четверг': 4, 'четверга': 4,
    'пт': 5, 'пятница': 5, 'пятницу': 5, 'пятницы': 5,
    'сб': 6, 'суббота': 6, 'субботу': 6, 'субботы': 6,
    'вс': 0, 'воскресенье': 0, 'воскресенья': 0,
  };
  return Object.prototype.hasOwnProperty.call(aliases, text) ? aliases[text] : null;
}

function deduplicateDeliveryRules_(rules) {
  const seen = {};
  return rules.filter((rule) => {
    const key = JSON.stringify(rule);
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function getAvailableDeliveryDates_(rules, numberOfDays, currentTime) {
  const now = currentTime || new Date();
  const today = Utilities.formatDate(now, CONFIG.timeZone, 'yyyy-MM-dd');
  const result = [];
  for (let offset = 0; offset < numberOfDays; offset += 1) {
    const date = addDaysToIsoDate_(today, offset);
    if (getDeliveryAvailability_(rules, date, now).available) result.push(date);
  }
  return result;
}

function getDeliveryAvailability_(rules, deliveryDate, currentTime) {
  const now = currentTime || new Date();
  const today = Utilities.formatDate(now, CONFIG.timeZone, 'yyyy-MM-dd');
  const currentMinutes = Number(Utilities.formatDate(now, CONFIG.timeZone, 'H')) * 60
    + Number(Utilities.formatDate(now, CONFIG.timeZone, 'm'));

  for (let daysBefore = 0; daysBefore <= 7; daysBefore += 1) {
    const orderDate = addDaysToIsoDate_(deliveryDate, -daysBefore);
    const orderWeekday = isoWeekday_(orderDate);
    for (const rule of rules) {
      if (rule.orderWeekday !== orderWeekday || calculateDeliveryDate_(orderDate, rule) !== deliveryDate) continue;
      if (orderDate > today) return { available: true, rule, orderDate };
      if (orderDate === today && currentMinutes < timeToMinutes_(rule.cutoffTime)) {
        return { available: true, rule, orderDate };
      }
    }
  }
  return { available: false };
}

function calculateDeliveryDate_(orderDate, rule) {
  if (rule.deliveryType === 'В тот же день') return orderDate;
  if (rule.deliveryType === 'Следующий день') return addDaysToIsoDate_(orderDate, rule.deliveryOffset || 1);
  const daysAhead = (rule.deliveryWeekday - isoWeekday_(orderDate) + 7) % 7 || 7;
  return addDaysToIsoDate_(orderDate, daysAhead);
}

function addDaysToIsoDate_(isoDate, days) {
  const parts = isoDate.split('-').map(Number);
  const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + days));
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
}

function isoWeekday_(isoDate) {
  const parts = isoDate.split('-').map(Number);
  return new Date(Date.UTC(parts[0], parts[1] - 1, parts[2])).getUTCDay();
}

function timeToMinutes_(value) {
  const parts = value.split(':').map(Number);
  return parts[0] * 60 + parts[1];
}

function formatRussianDate_(isoDate) {
  const parts = isoDate.split('-');
  return parts[2] + '.' + parts[1] + '.' + parts[0];
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
      // Не очищаем цвета в колонках заявок: там хранится статус приёмки.
      sheet.getRange(row, 1, 1, CONFIG.branchSheet.unitCol).setBackground(null);
    }
  });

  Object.keys(currentRows).forEach((key) => {
    if (!templateKeys[key]) {
      sheet.getRange(currentRows[key], 1, 1, CONFIG.branchSheet.unitCol)
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

function ensureSubmissionColumn_(sheet, date) {
  const wanted = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM');
  const lastColumn = Math.max(sheet.getLastColumn(), CONFIG.branchSheet.firstDateCol);
  const dates = sheet.getRange(2, CONFIG.branchSheet.firstDateCol, 1, lastColumn - CONFIG.branchSheet.firstDateCol + 1).getDisplayValues()[0];
  const indexes = [];
  dates.forEach((value, index) => {
    if (clean_(value) === wanted) indexes.push(index);
  });

  // Все заявки на одну дату хранятся в одной колонке. Повторная отправка
  // дополняет/редактирует позиции нужного поставщика, не создавая дубль даты.
  if (indexes.length) return CONFIG.branchSheet.firstDateCol + indexes[0];

  const column = lastColumn + 1;
  sheet.getRange(1, column).setValue(dayName_(date));
  sheet.getRange(2, column).setValue(wanted);
  return column;
}

function findSupplierRow_(sheet, supplier) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.branchSheet.firstProductRow) return 0;
  const values = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.productCol,
    lastRow - CONFIG.branchSheet.firstProductRow + 1, 1).getDisplayValues();
  const fontSizes = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.productCol,
    values.length, 1).getFontSizes();
  const wanted = normalizeKey_(supplier);
  const index = values.findIndex((row, offset) => normalizeKey_(row[0]) === wanted && Number(fontSizes[offset][0]) >= 16);
  return index < 0 ? 0 : CONFIG.branchSheet.firstProductRow + index;
}

function getSupplierProductRows_(sheet, supplier) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.branchSheet.firstProductRow) return [];
  const height = lastRow - CONFIG.branchSheet.firstProductRow + 1;
  const names = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.productCol, height, 1).getDisplayValues();
  const prices = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.priceCol, height, 1).getDisplayValues();
  const units = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.unitCol, height, 1).getDisplayValues();
  const fontSizes = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.productCol, height, 1).getFontSizes();
  const wanted = normalizeKey_(supplier);
  let active = false;
  const rows = [];
  names.forEach((value, index) => {
    const name = clean_(value[0]);
    const header = name && !clean_(prices[index][0]) && !clean_(units[index][0]) && Number(fontSizes[index][0]) >= 16;
    if (header) {
      active = normalizeKey_(name) === wanted;
    } else if (active && name) {
      rows.push(CONFIG.branchSheet.firstProductRow + index);
    }
  });
  return rows;
}

function validateSupplierMinimum_(payload, supplier) {
  const minimum = toNumber_(supplier.minimum);
  if (!minimum) return;
  const unit = clean_(supplier.unit).toLowerCase();
  let current = 0;
  if (/кг/.test(unit)) {
    current = payload.items.reduce((sum, item) => sum + toNumber_(item.quantity) * (toNumber_(item.unit) || 1), 0);
  } else if (/(?:шт|штук)/.test(unit)) {
    current = payload.items.reduce((sum, item) => sum + toNumber_(item.quantity), 0);
  } else {
    current = payload.items.reduce((sum, item) => sum + toNumber_(item.quantity) * toNumber_(item.price), 0);
  }
  if (current < minimum) {
    throw new Error('Минимальный заказ у поставщика — ' + supplier.minimum + ' ' + (supplier.unit || '₽') + '.');
  }
}

function buildProductRowIndex_(sheet) {
  const lastRow = sheet.getLastRow();
  const result = {};
  if (lastRow < CONFIG.branchSheet.firstProductRow) return result;

  const height = lastRow - CONFIG.branchSheet.firstProductRow + 1;
  const products = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.productCol, height, 1).getDisplayValues();
  const productFontSizes = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.productCol, height, 1).getFontSizes();
  const prices = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.priceCol, height, 1).getDisplayValues();
  const units = sheet.getRange(CONFIG.branchSheet.firstProductRow, CONFIG.branchSheet.unitCol, height, 1).getDisplayValues();

  products.forEach((row, index) => {
    const product = clean_(row[0]);
    if (!product) return;
    // Supplier rows are deliberately formatted as large headings. Do not mistake a
    // real product with temporarily empty price/multiplicity cells for a heading:
    // otherwise its ordered quantity would never reach the branch sheet.
    const isSupplierHeader = !clean_(prices[index][0])
      && !clean_(units[index][0])
      && Number(productFontSizes[index][0]) >= 16;
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
  if (!payload.supplier) throw new Error('Выберите поставщика.');
  if (!payload.items || !payload.items.length) throw new Error('Добавьте хотя бы одну позицию.');
}

function getEarliestDeliveryDate_(cutoffTime, currentTime) {
  const now = currentTime || new Date();
  const today = Utilities.formatDate(now, CONFIG.timeZone, 'yyyy-MM-dd');
  const cutoff = cutoffTime && cutoffTime.match(/^(\d{2}):(\d{2})$/);
  if (!cutoff) return today;

  const currentMinutes = Number(Utilities.formatDate(now, CONFIG.timeZone, 'H')) * 60
    + Number(Utilities.formatDate(now, CONFIG.timeZone, 'm'));
  const cutoffMinutes = Number(cutoff[1]) * 60 + Number(cutoff[2]);
  if (currentMinutes < cutoffMinutes) return today;

  const tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
  return Utilities.formatDate(tomorrow, CONFIG.timeZone, 'yyyy-MM-dd');
}

function normalizeCutoffTime_(value) {
  const text = clean_(value).toLowerCase();
  if (!text) return '';
  const match = text.match(/(?:до\s*)?(?:[01]?\d|2[0-3])(?:[:.]\d{2})?/);
  if (!match || (!/^\d{1,2}(?::|\.)?\d{0,2}$/.test(text) && !/до\s*\d/.test(text))) return '';
  const parts = match[0].replace(/до\s*/, '').replace('.', ':').split(':');
  const hour = Number(parts[0]);
  const minute = Number(parts[1] || 0);
  if (hour > 23 || minute > 59) return '';
  return String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
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
