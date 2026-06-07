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

  addSheet(workbook, 'Kho Quet Ma', [
    ['STT', 'Ten ma du an', 'Ma ban ve', 'So luong/may', 'Nha san xuat', 'Ngay nhap kho', 'N/A']
  ], payload.khoRows.map((row, index) => [
    index + 1,
    row.projectCode,
    row.drawingCode,
    row.quantity,
    row.manufacturer,
    row.importDate,
    row.note
  ]));

  addSheet(workbook, 'Bomlist Thiet Ke', [
    ['STT', 'Ten mat hang', 'Ma ban ve', 'Nha san xuat', 'So luong/may']
  ], payload.bomRows.map((row, index) => [
    index + 1,
    row.itemName,
    row.drawingCode,
    row.manufacturer,
    row.quantity
  ]));

  addCompareSheet(workbook, payload.compareRows);
  addDiscrepancySheet(workbook, payload.discrepancyRows || []);

  addConfirmSheet(workbook, payload.confirmRows || []);

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
  addCompareSheet(workbook, payload.compareRows || []);
  addConfirmSheet(workbook, payload.confirmRows || []);

  await workbook.xlsx.writeFile(result.filePath);
  return result.filePath;
});

ipcMain.handle('excel:exportDiscrepancy', async (_event, payload) => {
  const result = await dialog.showSaveDialog({
    title: 'Luu bang Thieu Thua',
    defaultPath: `ThieuThua_${getTimestamp()}.xlsx`,
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
    ['STT', 'Ma BOM', 'Ma Kho', 'Ten mat hang', 'Nha san xuat', 'So luong BOM', 'So luong Kho', 'Chenh lech', 'Trang thai', 'Do tuong dong', 'Ghi chu']
  ], compareRows.map((row, index) => [
    index + 1,
    row.bomDrawingCode,
    row.khoDrawingCode,
    row.itemName,
    row.manufacturer,
    row.bomQuantity,
    row.khoQuantity,
    row.difference,
    row.status,
    row.similarity,
    row.mergeNote
  ]));

  return compareSheet;
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

function addDiscrepancySheet(workbook, discrepancyRows) {
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
  const state = readLicenseState();
  const installedAt = new Date(state.installedAt);
  const trialEndsAt = new Date(installedAt.getTime() + TRIAL_DAYS * 24 * 60 * 60 * 1000);
  const now = new Date();
  const remainingMs = Math.max(0, trialEndsAt.getTime() - now.getTime());

  return {
    appVersion: app.getVersion(),
    licensed: Boolean(state.licensed),
    trialDays: TRIAL_DAYS,
    trialEndsAt: trialEndsAt.toISOString(),
    remainingMs,
    trialExpired: !state.licensed && remainingMs <= 0
  };
}

function activateLicense(code) {
  const normalizedCode = normalizeLicenseCode(code);
  if (!isLicenseFormat(normalizedCode)) {
    return { ok: false, message: 'Mã license không đúng định dạng XXX-XXX-XXX-XXXX.', status: getLicenseStatus() };
  }

  const licenseFile = findLicenseFile();
  if (!licenseFile) {
    return { ok: false, message: 'Không tìm thấy file licenses.json để xác thực license.', status: getLicenseStatus() };
  }

  const licenses = readLicensePool(licenseFile);
  const licenseIndex = licenses.findIndex((license) => normalizeLicenseCode(license) === normalizedCode);
  if (licenseIndex < 0) {
    return { ok: false, message: 'License không hợp lệ hoặc đã được sử dụng.', status: getLicenseStatus() };
  }

  licenses.splice(licenseIndex, 1);
  try {
    writeLicensePool(licenseFile, licenses);
  } catch (error) {
    return {
      ok: false,
      message: `Không thể xóa license đã dùng khỏi file: ${error.message}`,
      status: getLicenseStatus()
    };
  }

  const state = readLicenseState();
  writeLicenseState({
    ...state,
    licensed: true,
    activatedAt: new Date().toISOString(),
    licenseCode: normalizedCode
  });

  return { ok: true, message: 'Kích hoạt license thành công.', status: getLicenseStatus() };
}

function findLicenseFile() {
  const candidatePaths = [
    path.join(app.getPath('userData'), 'licenses.json'),
    path.join(process.cwd(), 'licenses.json')
  ];

  if (app.isPackaged) {
    candidatePaths.unshift(path.join(path.dirname(app.getPath('exe')), 'licenses.json'));
  } else {
    candidatePaths.unshift(path.join(__dirname, 'licenses.json'));
  }

  return candidatePaths.find((filePath) => fs.existsSync(filePath));
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
