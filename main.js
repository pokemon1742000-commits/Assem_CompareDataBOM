const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron');
const path = require('node:path');
const fs = require('node:fs');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const { autoUpdater } = require('electron-updater');

let mainWindow;
const TRIAL_DAYS = 7;

function createWindow() {
  const win = new BrowserWindow({
    width: 1400,
    height: 860,
    minWidth: 1100,
    minHeight: 680,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  if (app.isPackaged) {
    win.setMenu(null);
    win.setMenuBarVisibility(false);
  }

  mainWindow = win;
  win.loadFile('index.html');
}

ipcMain.handle('app:version', async () => app.getVersion());

ipcMain.handle('app:openGithub', async () => {
  await shell.openExternal('https://github.com/pokemon1742000-commits/Assem_CompareDataBOM');
  return true;
});

ipcMain.handle('app:licenseStatus', async () => getLicenseStatus());

ipcMain.handle('app:activateLicense', async (_event, code) => activateLicense(code));

ipcMain.handle('app:quit', async () => {
  app.quit();
  return true;
});

ipcMain.handle('excel:open', async () => {
  const result = await dialog.showOpenDialog({
    title: 'Chon file Excel',
    filters: [{ name: 'Excel', extensions: ['xlsx', 'xls', 'xlsm', 'csv'] }],
    properties: ['openFile', 'multiSelections']
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }

  return result.filePaths.map(readExcelInfo);
});

ipcMain.handle('excel:readSheets', async (_event, selections) => {
  return selections.map((selection) => readExcelFile(selection.filePath, selection.sheetName));
});

ipcMain.handle('update:check', async () => {
  if (!app.isPackaged) {
    return { message: 'Chức năng update chỉ hoạt động trên bản đã build exe.' };
  }

  try {
    sendUpdateStatus('Đang kiểm tra phiên bản mới...');
    await autoUpdater.checkForUpdates();
    return { message: 'Đang kiểm tra phiên bản mới...' };
  } catch (error) {
    const message = `Không thể kiểm tra update: ${error.message}`;
    sendUpdateStatus(message);
    return { message };
  }
});

ipcMain.handle('recent:load', async () => {
  const recent = readRecentState();
  return {
    khoFiles: loadRecentFiles(recent.khoSources || recent.khoPaths || []),
    bomFiles: loadRecentFiles(recent.bomSources || recent.bomPaths || [])
  };
});

ipcMain.handle('recent:save', async (_event, payload) => {
  writeRecentState({
    khoSources: payload.khoSources || [],
    bomSources: payload.bomSources || [],
    khoPaths: payload.khoPaths || [],
    bomPaths: payload.bomPaths || [],
    updatedAt: new Date().toISOString()
  });
  return true;
});

ipcMain.handle('recent:clear', async () => {
  writeRecentState({ khoSources: [], bomSources: [], khoPaths: [], bomPaths: [], updatedAt: new Date().toISOString() });
  return true;
});

ipcMain.handle('excel:export', async (_event, payload) => {
  const result = await dialog.showSaveDialog({
    title: 'Luu bao cao so sanh',
    defaultPath: `BaoCao_SoSanh_${getTimestamp()}.xlsx`,
    filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
  });

  if (result.canceled || !result.filePath) {
    return null;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Inventory Compare App';
  workbook.created = new Date();

  addSheet(workbook, 'Du Lieu Ma', [
    ['STT', 'Ma ban ve', 'Ten hang', 'Don vi tinh']
  ], payload.khoRows.map((row, index) => [
    index + 1,
    row.drawingCode,
    row.itemName,
    row.unit
  ]));

  addSheet(workbook, 'Bomlist Thiet Ke', [
    ['STT', 'Ten mat hang', 'Ma ban ve', 'Nha san xuat', 'So luong/may', 'Don vi tinh']
  ], payload.bomRows.map((row, index) => [
    index + 1,
    row.itemName,
    row.drawingCode,
    row.manufacturer,
    row.quantity,
    row.unit
  ]));

  addCompareSheetWithConfirm(workbook, payload.compareRows || [], payload.confirmRows || []);
  addDiscrepancySheet(workbook, payload.discrepancyRows || []);

  await workbook.xlsx.writeFile(result.filePath);
  return result.filePath;
});

ipcMain.handle('excel:exportCompare', async (_event, payload) => {
  const result = await dialog.showSaveDialog({
    title: 'Luu bang So Sanh',
    defaultPath: `SoSanh_ThieuThuaDu_${getTimestamp()}.xlsx`,
    filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
  });

  if (result.canceled || !result.filePath) {
    return null;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Inventory Compare App';
  workbook.created = new Date();
  addCompareSheetWithConfirm(workbook, payload.compareRows || [], payload.confirmRows || []);

  await workbook.xlsx.writeFile(result.filePath);
  return result.filePath;
});

ipcMain.handle('excel:exportDiscrepancy', async (_event, payload) => {
  const result = await dialog.showSaveDialog({
    title: 'Luu bang Ma Moi',
    defaultPath: `MaMoi_${getTimestamp()}.xlsx`,
    filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
  });

  if (result.canceled || !result.filePath) {
    return null;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Inventory Compare App';
  workbook.created = new Date();
  addDiscrepancySheet(workbook, payload.discrepancyRows || []);

  await workbook.xlsx.writeFile(result.filePath);
  return result.filePath;
});

function addCompareSheet(workbook, compareRows) {
  const compareSheet = addSheet(workbook, 'So Sanh', [
    ['Bang Giong Nhau'],
    ['STT', 'Ma da dat hang', 'Ma thiet ke', 'Ten ma dat hang', 'Don vi tinh', 'Ghi chu']
  ], compareRows.map((row, index) => [
    index + 1,
    row.orderDrawingCode,
    row.designDrawingCode,
    row.orderItemName,
    row.unit,
    row.note
  ]));

  compareRows.forEach((row, index) => {
    if (!row.corrected) return;
    compareSheet.getRow(index + 3).getCell(3).font = { name: 'Times New Roman', strike: true };
  });

  return compareSheet;
}

function addCompareSheetWithConfirm(workbook, compareRows, confirmRows) {
  const sheet = workbook.addWorksheet('So Sanh');

  sheet.addRow(['Bang Xac Nhan']);
  sheet.addRow(['STT', 'Ma thiet ke', 'Ma da dat hang de xuat', 'Ten mat hang', 'Don vi tinh', 'Do tuong dong']);
  (confirmRows || []).forEach((row, index) => {
    sheet.addRow([
      index + 1,
      row.designDrawingCode,
      row.orderDrawingCode,
      row.itemName,
      row.unit,
      row.similarity
    ]);
  });

  sheet.addRow([]);
  sheet.addRow(['Bang Giong Nhau']);
  sheet.addRow(['STT', 'Ma da dat hang', 'Ma thiet ke', 'Ten ma dat hang', 'Don vi tinh', 'Ghi chu']);
  const compareStartRow = sheet.rowCount + 1;
  (compareRows || []).forEach((row, index) => {
    sheet.addRow([
      index + 1,
      row.orderDrawingCode,
      row.designDrawingCode,
      row.orderItemName,
      row.unit,
      row.note
    ]);
    if (row.corrected) {
      sheet.getRow(compareStartRow + index).getCell(3).font = { name: 'Times New Roman', strike: true };
    }
  });

  formatSheet(sheet);
  return sheet;
}

function readExcelInfo(filePath) {
  const workbook = XLSX.readFile(filePath, { bookSheets: true });
  return {
    filePath,
    fileName: path.basename(filePath),
    sheets: workbook.SheetNames
  };
}

function readExcelFile(filePath, selectedSheetName) {
  const workbook = XLSX.readFile(filePath, { cellDates: false, raw: false });
  const sheetName = workbook.SheetNames.includes(selectedSheetName) ? selectedSheetName : workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    blankrows: false
  });

  return {
    filePath,
    fileName: path.basename(filePath),
    sheetName,
    rows
  };
}

function loadRecentFiles(sources) {
  return sources
    .map(normalizeRecentSource)
    .filter((source) => {
      try {
        return fs.existsSync(source.filePath);
      } catch {
        return false;
      }
    })
    .map((source) => {
      try {
        return readExcelFile(source.filePath, source.sheetName);
      } catch {
        return null;
      }
    })
    .filter(Boolean);
}

function getRecentStatePath() {
  return path.join(app.getPath('temp'), 'inventory-compare-recent-files.json');
}

function readRecentState() {
  try {
    const raw = fs.readFileSync(getRecentStatePath(), 'utf8');
    return JSON.parse(raw);
  } catch {
    return { khoPaths: [], bomPaths: [] };
  }
}

function normalizeRecentSource(source) {
  if (typeof source === 'string') {
    return { filePath: source, sheetName: '' };
  }

  return {
    filePath: source.filePath,
    sheetName: source.sheetName || ''
  };
}

function writeRecentState(payload) {
  fs.writeFileSync(getRecentStatePath(), JSON.stringify(payload, null, 2), 'utf8');
}

function addDiscrepancySheetLegacy(workbook, discrepancyRows) {
  const sheet = addSheet(workbook, 'Thieu Thua', [
    ['STT', 'Nguon', 'Ma BOM', 'Ma Kho', 'Ten mat hang', 'Nha san xuat', 'So luong BOM', 'So luong Kho', 'Chenh lech', 'Trang thai', 'Ghi chu']
  ], discrepancyRows.map((row, index) => [
    index + 1,
    row.source,
    row.bomDrawingCode,
    row.khoDrawingCode,
    row.itemName,
    row.manufacturer,
    row.bomQuantity,
    row.khoQuantity,
    row.difference,
    row.status,
    row.note
  ]));

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const status = row.getCell(10).value;
    if (status === 'Thiếu') {
      emphasizeRow(row);
    }
  });

  return sheet;
}

function addConfirmSheet(workbook, confirmRows) {
  if (!confirmRows.length) return null;
  return addSheet(workbook, 'Can Xac Nhan', [
    ['STT', 'Ma Kho', 'So luong Kho', 'Ma BOM de xuat', 'Ten mat hang', 'So luong BOM', 'Do tuong dong']
  ], confirmRows.map((row, index) => [
    index + 1,
    row.khoDrawingCode,
    row.khoQuantity,
    row.bomDrawingCode,
    row.itemName,
    row.bomQuantity,
    row.similarity
  ]));
}

function addDiscrepancySheet(workbook, discrepancyRows) {
  return addSheet(workbook, 'Ma Moi', [
    ['STT', 'Ma ban ve']
  ], discrepancyRows.map((row, index) => [
    index + 1,
    row.designDrawingCode
  ]));
}

function addSheet(workbook, name, headerRows, dataRows) {
  const sheet = workbook.addWorksheet(name);
  headerRows.concat(dataRows).forEach((row) => sheet.addRow(row));

  sheet.getRow(1).eachCell((cell) => {
    cell.font = { name: 'Times New Roman', bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DBEAFE' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  sheet.eachRow((row) => {
    row.eachCell((cell) => {
      cell.font = { name: 'Times New Roman', bold: row.number === 1 };
      cell.border = {
        top: { style: 'thin', color: { argb: 'D9E0EA' } },
        left: { style: 'thin', color: { argb: 'D9E0EA' } },
        bottom: { style: 'thin', color: { argb: 'D9E0EA' } },
        right: { style: 'thin', color: { argb: 'D9E0EA' } }
      };
    });
  });

  sheet.columns.forEach((column) => {
    let maxLength = 12;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const value = cell.value == null ? '' : String(cell.value);
      maxLength = Math.max(maxLength, value.length + 2);
    });
    column.width = Math.min(maxLength, 36);
  });

  sheet.views = [{ state: 'frozen', ySplit: 1 }];
  return sheet;
}

function formatSheet(sheet) {
  sheet.eachRow((row) => {
    const values = row.values.filter((value) => value !== null && value !== undefined && value !== '');
    const isHeader = row.number === 1 || row.number === 2 || values.length === 1 || values[0] === 'STT';
    row.eachCell((cell) => {
      cell.font = { name: 'Times New Roman', bold: isHeader, strike: cell.font?.strike };
      cell.border = {
        top: { style: 'thin', color: { argb: 'D9E0EA' } },
        left: { style: 'thin', color: { argb: 'D9E0EA' } },
        bottom: { style: 'thin', color: { argb: 'D9E0EA' } },
        right: { style: 'thin', color: { argb: 'D9E0EA' } }
      };
      if (isHeader) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DBEAFE' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    });
  });

  sheet.columns.forEach((column) => {
    let maxLength = 12;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const value = cell.value == null ? '' : String(cell.value);
      maxLength = Math.max(maxLength, value.length + 2);
    });
    column.width = Math.min(maxLength, 42);
  });

  sheet.views = [{ state: 'frozen', ySplit: 1 }];
}

function emphasizeRow(row) {
  row.eachCell((cell) => {
    cell.font = { name: 'Times New Roman', bold: true };
  });
}

function getTimestamp() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, '0');
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}`;
}

function getLicenseStatePath() {
  return path.join(app.getPath('userData'), 'license-state.json');
}

function readLicenseState() {
  const defaultState = {
    installedAt: new Date().toISOString(),
    licensed: false,
    activatedAt: '',
    licenseCode: ''
  };

  try {
    if (!fs.existsSync(getLicenseStatePath())) {
      writeLicenseState(defaultState);
      return defaultState;
    }

    const state = JSON.parse(fs.readFileSync(getLicenseStatePath(), 'utf8'));
    return { ...defaultState, ...state };
  } catch {
    writeLicenseState(defaultState);
    return defaultState;
  }
}

function writeLicenseState(state) {
  fs.mkdirSync(path.dirname(getLicenseStatePath()), { recursive: true });
  fs.writeFileSync(getLicenseStatePath(), JSON.stringify(state, null, 2), 'utf8');
}

function getLicenseStatus() {
  return {
    appVersion: app.getVersion(),
    licensed: true,
    trialDays: TRIAL_DAYS,
    trialEndsAt: '',
    remainingMs: 0,
    trialExpired: false
  };
}

function activateLicense(code) {
  return { ok: true, message: 'Phần mềm đang ở bản dùng vĩnh viễn.', status: getLicenseStatus() };
}

function findLicenseFile() {
  const candidatePaths = [
    path.join(app.getPath('userData'), 'licenses.json'),
    process.env.APPDATA ? path.join(process.env.APPDATA, 'Inventory Compare', 'licenses.json') : '',
    process.env.LOCALAPPDATA ? path.join(process.env.LOCALAPPDATA, 'Inventory Compare', 'licenses.json') : '',
    path.join(process.cwd(), 'licenses.json'),
    path.join(process.cwd(), 'release', 'licenses.json'),
    path.join(__dirname, 'licenses.json'),
    path.join(__dirname, 'release', 'licenses.json')
  ];

  if (app.isPackaged) {
    candidatePaths.unshift(path.join(path.dirname(app.getPath('exe')), 'licenses.json'));
  }

  return candidatePaths.filter(Boolean).find((filePath) => fs.existsSync(filePath));
}

function readLicensePool(filePath) {
  const raw = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  if (Array.isArray(raw)) return raw;
  if (Array.isArray(raw.licenses)) return raw.licenses;
  return [];
}

function writeLicensePool(filePath, licenses) {
  fs.writeFileSync(filePath, JSON.stringify({ licenses }, null, 2), 'utf8');
}

function normalizeLicenseCode(code) {
  return String(code || '').trim().toUpperCase();
}

function isLicenseFormat(code) {
  return /^[A-Z0-9]{3}-[A-Z0-9]{3}-[A-Z0-9]{3}-[A-Z0-9]{4}$/.test(code);
}

function sendUpdateStatus(message) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('update:status', message);
  }
}

function setupAutoUpdater() {
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;

  autoUpdater.on('checking-for-update', () => {
    sendUpdateStatus('Đang kiểm tra phiên bản mới...');
  });

  autoUpdater.on('update-available', (info) => {
    sendUpdateStatus(`Có phiên bản mới ${info.version}. Đang tải về...`);
  });

  autoUpdater.on('update-not-available', () => {
    sendUpdateStatus('Bạn đang dùng phiên bản mới nhất.');
  });

  autoUpdater.on('download-progress', (progress) => {
    sendUpdateStatus(`Đang tải update: ${Math.round(progress.percent)}%`);
  });

  autoUpdater.on('update-downloaded', async (info) => {
    sendUpdateStatus(`Đã tải xong phiên bản ${info.version}.`);
    const result = await dialog.showMessageBox(mainWindow, {
      type: 'question',
      buttons: ['Khởi động lại để cập nhật', 'Để sau'],
      defaultId: 0,
      cancelId: 1,
      title: 'Cập nhật phiên bản mới',
      message: `Đã tải xong phiên bản ${info.version}. Khởi động lại app để cập nhật ngay?`
    });

    if (result.response === 0) {
      autoUpdater.quitAndInstall();
    }
  });

  autoUpdater.on('error', (error) => {
    sendUpdateStatus(`Lỗi update: ${error.message}`);
  });
}

app.whenReady().then(() => {
  createWindow();
  setupAutoUpdater();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
