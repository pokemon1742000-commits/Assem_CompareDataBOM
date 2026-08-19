const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron');
const path = require('node:path');
const fs = require('node:fs');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const { autoUpdater } = require('electron-updater');

let mainWindow;

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

ipcMain.handle('app:confirmPendingExport', async () => {
  const result = await dialog.showMessageBox(mainWindow, {
    type: 'warning',
    buttons: ['Tiếp tục', 'Hủy bỏ'],
    defaultId: 1,
    cancelId: 1,
    title: 'Còn mã chưa xác nhận',
    message: 'Vẫn còn mã chưa xác nhận bạn có muốn tiếp tục không?'
  });
  return result.response === 0;
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
    title: 'Luu de nghi mua hang',
    defaultPath: `DeNghiMuaHang_${getTimestamp()}.xlsx`,
    filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
  });

  if (result.canceled || !result.filePath) {
    return null;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Design Code Standardizer';
  workbook.created = new Date();
  addPurchaseRequestSheet(workbook, payload);

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

function addPurchaseRequestSheet(workbook, payload) {
  const sheet = workbook.addWorksheet('De Nghi Mua Hang');
  const metadata = payload.metadata || {};
  const rows = payload.purchaseRows || [];
  const machineQuantity = Number(metadata.machineQuantity) || 1;

  sheet.columns = [
    { width: 6 },
    { width: 34 },
    { width: 28 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 10 },
    { width: 12 },
    { width: 12 },
    { width: 12 },
    { width: 28 }
  ];

  sheet.mergeCells('A1:B2');
  sheet.getCell('A1').value = 'MEIKO AUTOMATION';
  sheet.getCell('A1').font = { name: 'Times New Roman', bold: true, size: 14, color: { argb: '1F4E8C' } };
  sheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'left' };

  sheet.mergeCells('K1:L2');
  sheet.getCell('K1').value = `Mã tài liệu: ${metadata.documentCode || 'PUR-06-12-F-04-01'}`;
  sheet.getCell('K1').font = { name: 'Times New Roman', size: 12 };
  sheet.getCell('K1').alignment = { vertical: 'middle', horizontal: 'right' };

  sheet.mergeCells('A3:H3');
  sheet.getCell('A3').value = 'Address: 1F, EMS No.2 Building, Lot CN9, Thach That-Quoc Oai IZ, Phung Xa Commune, Thach That Dist, Ha Noi City, Vietnam.';
  sheet.getCell('A3').font = { name: 'Times New Roman', italic: true };

  sheet.mergeCells('A4:B4');
  sheet.getCell('A4').value = 'Tel: 024-6662-6684';
  sheet.getCell('A4').font = { name: 'Times New Roman', italic: true };

  sheet.mergeCells('J4:L4');
  sheet.getCell('J4').value = '*Phòng mua sẽ ghi khi ký tiếp nhận';
  sheet.getCell('J4').font = { name: 'Times New Roman', bold: true, size: 12 };
  sheet.getCell('J4').alignment = { horizontal: 'right' };

  sheet.mergeCells('A5:I6');
  sheet.getCell('A5').value = 'ĐỀ NGHỊ MUA HÀNG';
  sheet.getCell('A5').font = { name: 'Times New Roman', bold: true, size: 16 };
  sheet.getCell('A5').alignment = { vertical: 'middle', horizontal: 'center' };

  sheet.mergeCells('J5:L6');
  sheet.getCell('J5').value = `Leadtime:\n\nSố DNMH\nPR-PUR-${metadata.purchaseNumber || ''}`;
  sheet.getCell('J5').font = { name: 'Times New Roman', bold: true, size: 12 };
  sheet.getCell('J5').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };

  sheet.mergeCells('A7:I7');
  sheet.getCell('A7').value = `(${metadata.department || 'Bộ phận Cơ'})`;
  sheet.getCell('A7').alignment = { horizontal: 'center' };
  sheet.getCell('A7').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2CC' } };

  sheet.getCell('A8').value = `Mã dự án: ${metadata.projectCode || ''}`;
  sheet.getCell('A8').font = { name: 'Times New Roman', bold: true, size: 12 };
  sheet.getCell('A10').value = `Tên dự án: ${metadata.projectName || ''}`;
  sheet.getCell('A10').font = { name: 'Times New Roman', bold: true, size: 12 };

  sheet.mergeCells('F8:F9');
  sheet.getCell('F8').value = 'Số lượng máy\n↓';
  sheet.getCell('F8').font = { name: 'Times New Roman', bold: true, size: 12 };
  sheet.getCell('F8').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  sheet.getCell('F10').value = machineQuantity;
  sheet.getCell('F10').font = { name: 'Times New Roman', bold: true, size: 12 };
  sheet.getCell('F10').alignment = { horizontal: 'center' };
  sheet.getCell('F10').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2CC' } };

  sheet.mergeCells('G8:I10');
  sheet.getCell('G8').value = 'PUR ký tiếp nhận';
  sheet.getCell('G8').font = { name: 'Times New Roman', bold: true, size: 12 };
  sheet.getCell('G8').alignment = { vertical: 'top', horizontal: 'center' };

  sheet.getCell('J8').value = 'Duyệt';
  sheet.getCell('K8').value = 'Kiểm tra';
  sheet.getCell('L8').value = 'Lập';
  ['J8', 'K8', 'L8'].forEach((address) => {
    sheet.getCell(address).font = { name: 'Times New Roman', size: 12 };
    sheet.getCell(address).alignment = { horizontal: 'center' };
  });

  sheet.mergeCells('A11:A12');
  sheet.mergeCells('B11:B12');
  sheet.mergeCells('C11:C12');
  sheet.mergeCells('D11:D12');
  sheet.mergeCells('E11:G11');
  sheet.mergeCells('H11:H12');
  sheet.mergeCells('I11:I12');
  sheet.mergeCells('J11:J12');
  sheet.mergeCells('K11:K12');
  sheet.mergeCells('L11:L12');

  const headers = {
    A11: 'No',
    B11: 'Name',
    C11: 'Model/ Mã bản vẽ',
    D11: 'Maker',
    E11: 'Spec',
    E12: 'Marking part',
    F12: 'Material',
    G12: 'Surface',
    H11: 'Unit',
    I11: "Q'ty/\nMachine",
    J11: 'Number\nMachine',
    K11: 'Grand\nTotal',
    L11: 'Explain'
  };
  Object.entries(headers).forEach(([address, value]) => {
    sheet.getCell(address).value = value;
  });

  rows.forEach((row, index) => {
    const excelRow = sheet.getRow(13 + index);
    const quantity = Number(row.quantity) || 0;
    const numberMachine = Number(row.machineQuantity) || machineQuantity;
    excelRow.values = [
      index + 1,
      row.itemName || '',
      row.drawingCode || '',
      row.manufacturer || '',
      row.markingPart || '',
      row.material || '',
      row.surface || '',
      row.unit || '',
      quantity || '',
      numberMachine || '',
      quantity && numberMachine ? quantity * numberMachine : '',
      row.explain || ''
    ];
  });

  formatPurchaseRequestSheet(sheet, 12 + rows.length);
  return sheet;
}

function formatPurchaseRequestSheet(sheet, lastRowNumber) {
  for (let rowNumber = 1; rowNumber <= Math.max(lastRowNumber, 13); rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    row.height = rowNumber >= 11 ? 24 : undefined;
    for (let colNumber = 1; colNumber <= 12; colNumber += 1) {
      const cell = row.getCell(colNumber);
      cell.font = cell.font || { name: 'Times New Roman', size: 11 };
      cell.border = {
        top: { style: 'thin', color: { argb: 'A6A6A6' } },
        left: { style: 'thin', color: { argb: 'A6A6A6' } },
        bottom: { style: 'thin', color: { argb: 'A6A6A6' } },
        right: { style: 'thin', color: { argb: 'A6A6A6' } }
      };
      cell.alignment = cell.alignment || { vertical: 'middle', wrapText: true };
    }
  }

  for (let rowNumber = 11; rowNumber <= 12; rowNumber += 1) {
    sheet.getRow(rowNumber).eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'BDD7EE' } };
      cell.font = { name: 'Times New Roman', bold: true, size: 12 };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    });
  }

  sheet.views = [{ state: 'frozen', ySplit: 12 }];
  sheet.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0
  };
}

function addCompareSheet(workbook, compareRows) {
  const compareSheet = addSheet(workbook, 'So Sanh', [
    ['Bang Giong Nhau'],
    ['STT', 'Ma da dat hang', 'Ma thiet ke', 'Ten ma dat hang', 'DVT thiet ke', 'DVT da dat hang', 'Ghi chu']
  ], compareRows.map((row, index) => [
    index + 1,
    row.orderDrawingCode,
    row.designDrawingCode,
    row.orderItemName,
    row.designUnit,
    row.orderUnit,
    row.note
  ]));

  compareRows.forEach((row, index) => {
    if (!row.corrected) return;
    if (row.codeCorrected) {
      compareSheet.getRow(index + 3).getCell(3).font = { name: 'Times New Roman', strike: true };
    }
    if (row.unitCorrected) {
      compareSheet.getRow(index + 3).getCell(5).font = { name: 'Times New Roman', strike: true };
    }
  });

  return compareSheet;
}

function addCompareSheetWithConfirm(workbook, compareRows, confirmRows) {
  const sheet = workbook.addWorksheet('So Sanh');

  sheet.addRow(['Bang Xac Nhan']);
  sheet.addRow(['STT', 'Ma thiet ke', 'Ma da dat hang de xuat', 'Ten mat hang', 'DVT thiet ke', 'DVT da dat hang', 'Do tuong dong ma']);
  (confirmRows || []).forEach((row, index) => {
    sheet.addRow([
      index + 1,
      row.designDrawingCode,
      row.orderDrawingCode,
      row.itemName,
      row.designUnit,
      row.orderUnit,
      row.similarity
    ]);
  });

  sheet.addRow([]);
  sheet.addRow(['Bang Giong Nhau']);
  sheet.addRow(['STT', 'Ma da dat hang', 'Ma thiet ke', 'Ten ma dat hang', 'DVT thiet ke', 'DVT da dat hang', 'Ghi chu']);
  const compareStartRow = sheet.rowCount + 1;
  (compareRows || []).forEach((row, index) => {
    sheet.addRow([
      index + 1,
      row.orderDrawingCode,
      row.designDrawingCode,
      row.orderItemName,
      row.designUnit,
      row.orderUnit,
      row.note
    ]);
    if (row.codeCorrected) {
      sheet.getRow(compareStartRow + index).getCell(3).font = { name: 'Times New Roman', strike: true };
    }
    if (row.unitCorrected) {
      sheet.getRow(compareStartRow + index).getCell(5).font = { name: 'Times New Roman', strike: true };
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
  }).map((row) => {
    row.__rowNumber = row.__rowNum__ == null ? 0 : row.__rowNum__ + 1;
    return row;
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
    ['STT', 'Ma thiet ke', 'Ma da dat hang de xuat', 'Ten mat hang', 'DVT thiet ke', 'DVT da dat hang', 'Do tuong dong ma']
  ], confirmRows.map((row, index) => [
    index + 1,
    row.designDrawingCode,
    row.orderDrawingCode,
    row.itemName,
    row.designUnit,
    row.orderUnit,
    row.similarity
  ]));
}

function addDiscrepancySheet(workbook, discrepancyRows) {
  const sheet = addSheet(workbook, 'Ma Moi', [
    ['STT', 'Ma ban ve', 'Ma da dat hang goi y', 'Ten hang goi y', 'Do giong ten hang', 'Ghi chu']
  ], discrepancyRows.map((row, index) => [
    index + 1,
    row.designDrawingCode,
    row.suggestedOrderDrawingCode,
    row.suggestedOrderItemName,
    row.nameSimilarity,
    row.note
  ]));
  sheet.getColumn(1).width = 6;
  return sheet;
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
