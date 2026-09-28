const CONFIG = {
  SPREADSHEET_ID: '1kpkk1UdmMExtew7s504aPZTDwKQvWjhxt83sdaijis',
  DRIVE_FOLDER_ID: '1rqHbMgyFSKYRDbxv-UHFpby5S-9-Wlzj'
};

const SHEETS = {
  PUMP: 'Pump Trucks',
  ASSET: 'Asset Counts',
  TYPES: 'Asset Types',
  SETTINGS: 'Settings'
};

const HEADERS = {
  [SHEETS.PUMP]: ['Pump Truck ID', 'Date/Time', 'Pump Truck Type', 'Condition', 'Area', 'Tags', 'Notes', 'Image File ID', 'Image URL', 'Recorded By'],
  [SHEETS.ASSET]: ['Record ID', 'Date/Time', 'Asset Type', 'Quantity', 'Area', 'Tags', 'Notes', 'Image File ID', 'Image URL', 'Recorded By'],
  [SHEETS.TYPES]: ['Asset Type', 'Active'],
  [SHEETS.SETTINGS]: ['Key', 'Value']
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Warehouse Asset Register')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupProject() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  Object.keys(HEADERS).forEach(name => ensureSheet_(ss, name, HEADERS[name]));

  const typeSheet = ss.getSheetByName(SHEETS.TYPES);
  if (typeSheet.getLastRow() < 2) {
    const defaults = [
      ['Magnum', true],
      ['Wood Pallet', true],
      ['Plastic Pallet', true],
      ['Car Bag', true],
      ['Red Sleeve', true],
      ['Blue Sleeve', true],
      ['Pump Truck', true]
    ];
    typeSheet.getRange(2, 1, defaults.length, 2).setValues(defaults);
  }

  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  const rows = settings.getLastRow();
  if (rows < 2) {
    settings.getRange(2, 1, 3, 2).setValues([
      ['Project', 'Warehouse Asset Register'],
      ['PumpTruckPrefix', 'PT-'],
      ['AssetRecordPrefix', 'AC-']
    ]);
  }

  return { ok: true, message: 'Sheet structure is ready.' };
}

function getInitialData() {
  setupProject();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return {
    pumpTrucks: readSheet_(ss.getSheetByName(SHEETS.PUMP)),
    assetCounts: readSheet_(ss.getSheetByName(SHEETS.ASSET)),
    assetTypes: readSheet_(ss.getSheetByName(SHEETS.TYPES)).filter(r => String(r[1]).toLowerCase() === 'true').map(r => r[0]).filter(Boolean)
  };
}

function getData() {
  return getInitialData();
}

function addAssetType(type) {
  const clean = String(type || '').trim();
  if (!clean) throw new Error('Enter an asset type.');
  setupProject();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEETS.TYPES);
  const data = sh.getDataRange().getValues();
  const existing = data.slice(1).find(r => String(r[0]).trim().toLowerCase() === clean.toLowerCase());
  if (existing) return { ok: true, message: 'Asset type already exists.' };
  sh.appendRow([clean, true]);
  return { ok: true, message: 'Asset type added.' };
}

function savePumpTruck(record) {
  setupProject();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEETS.PUMP);
    const id = nextPumpTruckId_(sh);
    const image = saveImage_(record.imageData, record.imageName, 'Pump Trucks');
    const now = new Date();
    const row = [
      id,
      now,
      clean_(record.type),
      clean_(record.condition),
      clean_(record.area),
      clean_(record.tags),
      clean_(record.notes),
      image ? image.id : '',
      image ? image.url : '',
      getUser_()
    ];
    sh.appendRow(row);
    return { ok: true, record: rowToObject_(HEADERS[SHEETS.PUMP], row) };
  } finally {
    lock.releaseLock();
  }
}

function saveAssetCount(record) {
  setupProject();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEETS.ASSET);
  const image = saveImage_(record.imageData, record.imageName, 'Asset Counter');
  const id = nextRecordId_(sh, 'AC-');
  const now = new Date();
  const qty = Math.max(1, Number(record.quantity || 1));
  const row = [
    id,
    now,
    clean_(record.type),
    qty,
    clean_(record.area),
    clean_(record.tags),
    clean_(record.notes),
    image ? image.id : '',
    image ? image.url : '',
    getUser_()
  ];
  sh.appendRow(row);
  return { ok: true, record: rowToObject_(HEADERS[SHEETS.ASSET], row) };
}

function updatePumpTruck(id, record) {
  setupProject();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEETS.PUMP);
  const values = sh.getDataRange().getValues();
  const idx = values.findIndex((r, i) => i > 0 && String(r[0]) === String(id));
  if (idx < 1) throw new Error('Pump truck not found.');

  const rowNum = idx + 1;
  let imageId = values[idx][7] || '';
  let imageUrl = values[idx][8] || '';
  if (record.removeImage) {
    if (imageId) trashFile_(imageId);
    imageId = '';
    imageUrl = '';
  }
  if (record.imageData) {
    if (imageId) trashFile_(imageId);
    const image = saveImage_(record.imageData, record.imageName, 'Pump Trucks');
    imageId = image.id;
    imageUrl = image.url;
  }

  const row = [
    id,
    values[idx][1] || new Date(),
    clean_(record.type),
    clean_(record.condition),
    clean_(record.area),
    clean_(record.tags),
    clean_(record.notes),
    imageId,
    imageUrl,
    values[idx][9] || getUser_()
  ];
  sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
  return { ok: true, record: rowToObject_(HEADERS[SHEETS.PUMP], row) };
}

function deletePumpTruck(id) {
  setupProject();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEETS.PUMP);
  const values = sh.getDataRange().getValues();
  const idx = values.findIndex((r, i) => i > 0 && String(r[0]) === String(id));
  if (idx < 1) throw new Error('Pump truck not found.');
  const imageId = values[idx][7];
  if (imageId) trashFile_(imageId);
  sh.deleteRow(idx + 1);
  return { ok: true };
}

function deleteAssetRecord(id) {
  setupProject();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEETS.ASSET);
  const values = sh.getDataRange().getValues();
  const idx = values.findIndex((r, i) => i > 0 && String(r[0]) === String(id));
  if (idx < 1) throw new Error('Asset record not found.');
  const imageId = values[idx][7];
  if (imageId) trashFile_(imageId);
  sh.deleteRow(idx + 1);
  return { ok: true };
}

function saveImage_(dataUrl, originalName, subfolderName) {
  if (!dataUrl) return null;
  const match = String(dataUrl).match(/^data:(.+?);base64,(.*)$/);
  if (!match) throw new Error('Invalid image data.');
  const contentType = match[1];
  const bytes = Utilities.base64Decode(match[2]);
  const root = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const folder = getOrCreateFolder_(root, subfolderName);
  const safeName = sanitizeFileName_(originalName || 'warehouse-photo.jpg');
  const file = folder.createFile(Utilities.newBlob(bytes, contentType, safeName));
  return {
    id: file.getId(),
    url: 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(file.getId()) + '&sz=w1200'
  };
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function trashFile_(id) {
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
}

function nextPumpTruckId_(sh) {
  const values = sh.getDataRange().getValues();
  let max = 0;
  values.slice(1).forEach(r => {
    const m = String(r[0] || '').match(/^PT-(\d+)$/i);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return 'PT-' + String(max + 1).padStart(4, '0');
}

function nextRecordId_(sh, prefix) {
  const values = sh.getDataRange().getValues();
  let max = 0;
  values.slice(1).forEach(r => {
    const m = String(r[0] || '').match(new RegExp('^' + prefix.replace('-', '\\-') + '(\\d+)$', 'i'));
    if (m) max = Math.max(max, Number(m[1]));
  });
  return prefix + String(max + 1).padStart(5, '0');
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  else sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  const header = sh.getRange(1, 1, 1, headers.length);
  header.setFontWeight('bold');
  header.setBackground('#17324d');
  header.setFontColor('#ffffff');
  header.setWrap(true);
  sh.setFrozenRows(1);
  if (name === SHEETS.PUMP || name === SHEETS.ASSET) {
    if (sh.getMaxRows() > 1) sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  }
  sh.autoResizeColumns(1, headers.length);
}

function readSheet_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).filter(r => r.some(v => v !== '')).map(r => rowToObject_(headers, r));
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    let value = row[i];
    if (value instanceof Date) value = value.toISOString();
    obj[h] = value;
  });
  return obj;
}

function clean_(value) {
  return String(value == null ? '' : value).trim();
}

function sanitizeFileName_(name) {
  return String(name).replace(/[^a-zA-Z0-9._-]/g, '_').slice(0, 100);
}

function getUser_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}
