const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron');
const { spawn } = require('node:child_process');
const https = require('node:https');
const path = require('node:path');
const fs = require('node:fs');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

let mainWindow;
let updateInProgress = false;
const UPDATE_MANIFEST_URL = 'https://raw.githubusercontent.com/pokemon1742000-commits/Assem_CompareDataBOM/main/update.json';

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

  if (updateInProgress) {
    return { message: 'Update đang được xử lý.' };
  }

  updateInProgress = true;
  try {
    sendUpdateStatus('Đang kiểm tra phiên bản mới...');
    const manifest = await fetchUpdateManifest();
    if (!isNewerVersion(manifest.version, app.getVersion())) {
      sendUpdateStatus('Bạn đang dùng phiên bản mới nhất.');
      return { message: 'Bạn đang dùng phiên bản mới nhất.' };
    }

    sendUpdateStatus(`Có phiên bản mới ${manifest.version}. Đang tải về...`);
    const installerPath = await downloadInstaller(manifest.url, manifest.version);
    sendUpdateStatus('Đã tải xong. Đang cài đặt âm thầm và khởi động lại...');
    installUpdate(installerPath);
    return { message: `Đang cài đặt phiên bản ${manifest.version}...` };
  } catch (error) {
    const message = `Không thể kiểm tra update: ${error.message}`;
    sendUpdateStatus(message);
    updateInProgress = false;
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
    title: 'Luu bao cao',
    defaultPath: `BaoCao_${getTimestamp()}.xlsx`,
    filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
  });

  if (result.canceled || !result.filePath) {
    return null;
  }

  const khoRows = payload.khoRows || [];
  const bomRows = payload.bomRows || [];
  const compareRows = payload.compareRows || [];
  const discrepancyRows = payload.discrepancyRows || [];
  const confirmRows = payload.confirmRows || [];

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Inventory Compare App';
  workbook.created = new Date();

  addStatsSheet(workbook, {
    khoCount: khoRows.length,
    bomCount: bomRows.length,
    okCount: compareRows.filter((row) => row.status === 'Đủ').length,
    missingCount: discrepancyRows.filter((row) => row.status === 'Thiếu').length,
    extraCount: discrepancyRows.filter((row) => row.status === 'Thừa').length,
    confirmCount: confirmRows.length
  });

  addSheet(workbook, 'Dữ Liệu Kho', [
    ['STT', 'Ten ma du an', 'Ma ban ve', 'So luong/may', 'Nha san xuat', 'Ngay nhap kho', 'N/A']
  ], khoRows.map((row, index) => [
    index + 1,
    row.projectCode,
    row.drawingCode,
    row.quantity,
    row.manufacturer,
    row.importDate,
    row.note
  ]));

  addSheet(workbook, 'Dữ Liệu Thiết Kế', [
    ['STT', 'Ten mat hang', 'Ma ban ve', 'Nha san xuat', 'So luong/may']
  ], bomRows.map((row, index) => [
    index + 1,
    row.itemName,
    row.drawingCode,
    row.manufacturer,
    row.quantity
  ]));

  addCompareSheet(workbook, compareRows);
  addDiscrepancySheet(workbook, discrepancyRows);

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

function addStatsSheet(workbook, stats) {
  const headers = ['Chỉ số', 'Số lượng'];
  const rows = [
    ['Dữ Liệu Kho', stats.khoCount],
    ['Dữ Liệu Thiết Kế', stats.bomCount],
    ['Đủ hàng', stats.okCount],
    ['Thiếu hàng', stats.missingCount],
    ['Thừa hàng', stats.extraCount],
    ['Cần xác nhận', stats.confirmCount]
  ];
  return addSheet(workbook, 'Bảng Thống Kê', [headers], rows);
}

function addCompareSheet(workbook, compareRows) {
  const headers = ['STT', 'Ma BOM', 'Ma Kho', 'Ten mat hang', 'Nha san xuat', 'So luong BOM', 'So luong Kho', 'Chenh lech', 'Trang thai', 'Do tuong dong', 'Ghi chu'];
  const toValues = (row, index) => [
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
  ];

  const sufficientRows = compareRows.filter((row) => row.status === 'Đủ');
  const otherRows = compareRows.filter((row) => row.status !== 'Đủ');

  const compareSheet = addSheet(workbook, 'So Sánh', [headers], otherRows.map(toValues), {
    countLabel: 'Tổng số mã',
    countFillArgb: 'FFF4CC',
    countFontArgb: '92400E'
  });

  const sufficientSheet = addSheet(workbook, 'Đủ Hàng', [headers], sufficientRows.map(toValues), {
    countLabel: 'Tổng số mã',
    countFillArgb: 'DFF7E7',
    countFontArgb: '15803D'
  });

  return { compareSheet, sufficientSheet };
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
  const headers = ['STT', 'Nguon', 'Ma BOM', 'Ma Kho', 'Ten mat hang', 'Nha san xuat', 'So luong BOM', 'So luong Kho', 'Chenh lech', 'Trang thai', 'Ghi chu'];
  const toValues = (row, index) => [
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
  ];

  const missingRows = discrepancyRows.filter((row) => row.status === 'Thiếu');
  const extraRows = discrepancyRows.filter((row) => row.status === 'Thừa');
  const missingSheet = addSheet(workbook, 'Thiếu', [headers], missingRows.map(toValues), {
    countLabel: 'Tổng số mã',
    countFillArgb: 'FCE4E4',
    countFontArgb: '9B1C1C'
  });
  const extraSheet = addSheet(workbook, 'Thừa', [headers], extraRows.map(toValues), {
    countLabel: 'Tổng số mã',
    countFillArgb: 'EDE4FF',
    countFontArgb: '5B21B6'
  });

  missingSheet.eachRow((row, rowNumber) => {
    if (rowNumber > missingSheet.headerRowNumber) emphasizeRow(row);
  });

  return { missingSheet, extraSheet };
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

function addSheet(workbook, name, headerRows, dataRows, options = {}) {
  const sheet = workbook.addWorksheet(name);
  const columnCount = (headerRows[0] || []).length || 1;
  let headerRowNumber = 1;

  if (options.countLabel) {
    const countRow = sheet.addRow([`${options.countLabel}: ${dataRows.length}`]);
    sheet.mergeCells(countRow.number, 1, countRow.number, columnCount);
    const countCell = countRow.getCell(1);
    countCell.font = { name: 'Times New Roman', bold: true, color: { argb: options.countFontArgb || '1F2937' } };
    countCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: options.countFillArgb || 'E5E7EB' } };
    countCell.alignment = { vertical: 'middle', horizontal: 'left' };
    headerRowNumber = countRow.number + 1;
  }

  headerRows.concat(dataRows).forEach((row) => sheet.addRow(row));

  sheet.getRow(headerRowNumber).eachCell((cell) => {
    cell.font = { name: 'Times New Roman', bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'DBEAFE' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  sheet.eachRow((row) => {
    const isCountRow = options.countLabel && row.number === 1 && headerRowNumber !== 1;
    row.eachCell((cell) => {
      if (!isCountRow) {
        cell.font = { name: 'Times New Roman', bold: row.number === headerRowNumber };
      }
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

  sheet.views = [{ state: 'frozen', ySplit: headerRowNumber }];
  sheet.headerRowNumber = headerRowNumber;
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

function sendUpdateStatus(message) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('update:status', message);
  }
}

function fetchUpdateManifest() {
  return new Promise((resolve, reject) => {
    https.get(UPDATE_MANIFEST_URL, (response) => {
      if (response.statusCode !== 200) {
        response.resume();
        reject(new Error(`Manifest trả về HTTP ${response.statusCode}`));
        return;
      }

      let body = '';
      response.setEncoding('utf8');
      response.on('data', (chunk) => { body += chunk; });
      response.on('end', () => {
        try {
          const manifest = JSON.parse(body);
          if (!manifest.version || !manifest.url) {
            throw new Error('Manifest thiếu version hoặc url');
          }
          resolve(manifest);
        } catch (error) {
          reject(new Error(`Manifest không hợp lệ: ${error.message}`));
        }
      });
    }).on('error', reject);
  });
}

function parseVersion(version) {
  const match = String(version || '').match(/^v?(\d+)\.(\d+)\.(\d+)(?:-[0-9A-Za-z.-]+)?$/);
  return match ? match.slice(1, 4).map(Number) : null;
}

function isNewerVersion(candidate, current) {
  const next = parseVersion(candidate);
  const installed = parseVersion(current);
  if (!next || !installed) {
    throw new Error('Version trong manifest không hợp lệ');
  }

  for (let index = 0; index < next.length; index += 1) {
    if (next[index] !== installed[index]) return next[index] > installed[index];
  }
  return false;
}

function downloadInstaller(url, version, redirectCount = 0) {
  const installerPath = path.join(app.getPath('temp'), `inventory-compare-update-${version}.exe`);
  return new Promise((resolve, reject) => {
    const request = https.get(url, (response) => {
      if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
        response.resume();
        if (redirectCount >= 5) {
          reject(new Error('Installer chuyển hướng quá nhiều lần'));
          return;
        }
        downloadInstaller(new URL(response.headers.location, url).toString(), version, redirectCount + 1)
          .then(resolve)
          .catch(reject);
        return;
      }

      if (response.statusCode !== 200) {
        response.resume();
        reject(new Error(`Installer trả về HTTP ${response.statusCode}`));
        return;
      }

      const totalBytes = Number(response.headers['content-length']) || 0;
      let downloadedBytes = 0;
      const output = fs.createWriteStream(installerPath);
      response.on('data', (chunk) => {
        downloadedBytes += chunk.length;
        const percent = totalBytes ? Math.round((downloadedBytes / totalBytes) * 100) : 0;
        sendUpdateStatus(`Đang tải update: ${percent}%`);
      });
      response.pipe(output);
      output.on('finish', () => output.close(() => resolve(installerPath)));
      output.on('error', (error) => {
        fs.rm(installerPath, { force: true }, () => reject(error));
      });
    });
    request.on('error', reject);
  });
}

function installUpdate(installerPath) {
  const installer = spawn(installerPath, ['/S'], {
    detached: true,
    stdio: 'ignore',
    windowsHide: true
  });
  installer.unref();
  app.quit();
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
