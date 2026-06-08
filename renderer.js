const state = {
  view: 'dashboard',
  search: '',
  appVersion: '',
  licenseStatus: null,
  khoFileName: '',
  bomFileName: '',
  khoPaths: [],
  bomPaths: [],
  khoSources: [],
  bomSources: [],
  khoRows: [],
  bomRows: [],
  compareRows: [],
  discrepancyRows: [],
  confirmRows: [],
  manualMatches: new Map(),
  rejectedMatches: new Set()
};

const columns = {
  kho: [
    ['stt', 'STT', 'number'],
    ['projectCode', 'Tên mã dự án'],
    ['drawingCode', 'Mã bản vẽ'],
    ['quantity', 'Số lượng/máy', 'number'],
    ['manufacturer', 'Nhà sản xuất'],
    ['importDate', 'Ngày nhập kho'],
    ['note', 'N/A']
  ],
  bom: [
    ['stt', 'STT', 'number'],
    ['itemName', 'Tên mặt hàng'],
    ['drawingCode', 'Mã bản vẽ'],
    ['manufacturer', 'Nhà sản xuất'],
    ['quantity', 'Số lượng/máy', 'number']
  ],
  compare: [
    ['stt', 'STT', 'number'],
    ['bomDrawingCode', 'Mã BOM'],
    ['khoDrawingCode', 'Mã Kho'],
    ['itemName', 'Tên mặt hàng'],
    ['manufacturer', 'Nhà sản xuất'],
    ['bomQuantity', 'SL BOM', 'number'],
    ['khoQuantity', 'SL Kho', 'number'],
    ['difference', 'Chênh lệch', 'number'],
    ['status', 'Trạng thái'],
    ['similarity', 'Độ tương đồng'],
    ['mergeNote', 'Ghi chú']
  ],
  discrepancy: [
    ['stt', 'STT', 'number'],
    ['source', 'Nguồn'],
    ['bomDrawingCode', 'Mã BOM'],
    ['khoDrawingCode', 'Mã Kho'],
    ['itemName', 'Tên mặt hàng'],
    ['manufacturer', 'Nhà sản xuất'],
    ['bomQuantity', 'SL BOM', 'number'],
    ['khoQuantity', 'SL Kho', 'number'],
    ['difference', 'Chênh lệch', 'number'],
    ['status', 'Trạng thái'],
    ['note', 'Ghi chú']
  ],
  confirm: [
    ['stt', 'STT', 'number'],
    ['khoDrawingCode', 'Mã Kho'],
    ['khoQuantity', 'SL Kho', 'number'],
    ['bomSelect', 'Mã BOM đề xuất'],
    ['itemName', 'Tên mặt hàng'],
    ['bomQuantity', 'SL BOM', 'number'],
    ['similarity', 'Độ tương đồng'],
    ['actions', 'Hành động']
  ]
};

const els = {};

document.addEventListener('DOMContentLoaded', () => {
  bindElements();
  bindEvents();
  renderAppVersion();
  renderAll();
  restoreRecentFiles();
});

function bindElements() {
  [
    'loadKhoBtn',
    'loadBomBtn',
    'clearBtn',
    'updateBtn',
    'infoBtn',
    'licenseBtn',
    'exportExcelBtn',
    'searchInput',
    'autoThreshold',
    'confirmThreshold',
    'appVersion',
    'trialStatus',
    'licenseOverlay',
    'expiredLicenseInput',
    'expiredLicenseBtn',
    'quitAppBtn',
    'expiredLicenseMessage',
    'khoCount',
    'bomCount',
    'sideKhoCount',
    'sideBomCount',
    'okCount',
    'missingCount',
    'extraCount',
    'confirmCount',
    'khoFileName',
    'bomFileName',
    'khoTable',
    'bomTable',
    'compareTable',
    'discrepancyTable',
    'confirmTable',
    'toast'
  ].forEach((id) => {
    els[id] = document.getElementById(id);
  });
}

function bindEvents() {
  document.querySelectorAll('.nav-item').forEach((button) => {
    button.addEventListener('click', () => setView(button.dataset.view));
  });

  els.loadKhoBtn.addEventListener('click', loadKhoFile);
  els.loadBomBtn.addEventListener('click', loadBomFile);
  els.clearBtn.addEventListener('click', clearData);
  els.updateBtn.addEventListener('click', checkForUpdates);
  els.infoBtn.addEventListener('click', openGithubInfo);
  if (els.licenseBtn) {
    els.licenseBtn.hidden = true;
    els.licenseBtn.addEventListener('click', showLicenseDialog);
  }
  els.expiredLicenseBtn.addEventListener('click', activateExpiredLicense);
  els.quitAppBtn.addEventListener('click', () => window.inventoryApi.quitApp());
  els.exportExcelBtn.addEventListener('click', exportCurrentTable);
  window.inventoryApi.onUpdateStatus((message) => showToast(message));
  document.querySelectorAll('.theme-dot').forEach((button) => {
    button.addEventListener('click', () => {
      applyTheme(button.dataset.theme);
      localStorage.setItem('inventory-theme', button.dataset.theme);
    });
  });
  els.searchInput.addEventListener('input', (event) => {
    state.search = event.target.value.trim();
    renderCurrentTable();
  });
  els.autoThreshold.addEventListener('change', handleThresholdChange);
  els.confirmThreshold.addEventListener('change', handleThresholdChange);
  applyTheme(localStorage.getItem('inventory-theme') || 'default');
}

function applyTheme(themeName) {
  document.body.classList.remove('theme-mint', 'theme-sky', 'theme-lavender');
  if (themeName !== 'default') {
    document.body.classList.add(`theme-${themeName}`);
  }
  document.querySelectorAll('.theme-dot').forEach((button) => {
    button.classList.toggle('active', button.dataset.theme === themeName);
  });
}

async function renderAppVersion() {
  try {
    state.appVersion = await window.inventoryApi.getAppVersion();
    await refreshLicenseStatus();
  } catch {
    els.appVersion.textContent = 'Trial';
  }
}

async function refreshLicenseStatus() {
  state.licenseStatus = await window.inventoryApi.getLicenseStatus();
  renderLicenseStatus();
}

function renderLicenseStatus() {
  const status = state.licenseStatus;
  if (!status) return;

  if (status.licensed) {
    els.appVersion.textContent = `v${status.appVersion || state.appVersion}`;
    els.trialStatus.textContent = '';
    els.licenseOverlay.hidden = true;
    if (els.licenseBtn) els.licenseBtn.hidden = true;
    return;
  }

  els.appVersion.textContent = 'Trial';
  els.trialStatus.textContent = `Trial còn lại: ${formatRemainingTime(status.remainingMs)}`;
  els.licenseOverlay.hidden = !status.trialExpired;
}

function startLicenseTimer() {
  clearInterval(startLicenseTimer.timer);
  startLicenseTimer.timer = setInterval(async () => {
    if (!state.licenseStatus?.licensed) {
      await refreshLicenseStatus();
    }
  }, 1000);
}

function formatRemainingTime(ms) {
  const totalSeconds = Math.max(0, Math.floor(ms / 1000));
  const days = Math.floor(totalSeconds / 86400);
  const hours = Math.floor((totalSeconds % 86400) / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  return `${days} ngày ${hours} giờ ${minutes} phút ${seconds} giây`;
}

async function loadKhoFile() {
  try {
    const files = await loadSelectedExcelFiles();
    if (!files.length) return;

    const parsedRows = files.flatMap((file) => parseKhoRows(file.rows));
    state.khoSources = appendSources(state.khoSources, files.map(toFileSource));
    state.khoPaths = appendPaths(state.khoPaths, files.map((file) => file.filePath));
    state.khoRows = renumberRows(state.khoRows.concat(parsedRows));
    state.khoFileName = appendFileName(state.khoFileName, files.map(formatFileLabel).join('; '));
    state.compareRows = [];
    state.discrepancyRows = [];
    state.confirmRows = [];
    state.manualMatches.clear();
    state.rejectedMatches.clear();

    await saveRecentFiles();
    autoCompareAfterLoad('kho');
    showToast(`Đã thêm ${parsedRows.length} dòng kho từ ${files.length} file. Tổng hiện có ${state.khoRows.length} dòng.`);
  } catch (error) {
    showToast(`Không thể load file kho: ${error.message}`);
  }
}

async function loadBomFile() {
  try {
    const files = await loadSelectedExcelFiles();
    if (!files.length) return;

    const parsedRows = files.flatMap((file) => parseBomRows(file.rows));
    state.bomSources = appendSources(state.bomSources, files.map(toFileSource));
    state.bomPaths = appendPaths(state.bomPaths, files.map((file) => file.filePath));
    state.bomRows = renumberRows(state.bomRows.concat(parsedRows));
    state.bomFileName = appendFileName(state.bomFileName, files.map(formatFileLabel).join('; '));
    state.compareRows = [];
    state.discrepancyRows = [];
    state.confirmRows = [];
    state.manualMatches.clear();
    state.rejectedMatches.clear();

    await saveRecentFiles();
    autoCompareAfterLoad('bom');
    showToast(`Đã thêm ${parsedRows.length} dòng BOM từ ${files.length} file. Tổng hiện có ${state.bomRows.length} dòng.`);
  } catch (error) {
    showToast(`Không thể load file BOM: ${error.message}`);
  }
}

async function checkForUpdates() {
  try {
    els.updateBtn.disabled = true;
    const result = await window.inventoryApi.checkForUpdates();
    if (result?.message) showToast(result.message);
  } catch (error) {
    showToast(`Không thể kiểm tra update: ${error.message}`);
  } finally {
    els.updateBtn.disabled = false;
  }
}

async function openGithubInfo() {
  const overlay = document.createElement('div');
  overlay.className = 'sheet-modal-overlay';
  overlay.innerHTML = `
    <div class="info-modal">
      <header>
        <h2>Thông tin GitHub</h2>
        <p>Trang GitHub dùng để phát hành phiên bản mới và kiểm tra update.</p>
      </header>
      <div class="info-content">
        <button class="github-link" data-action="open-github">https://github.com/pokemon1742000-commits/Assem_CompareDataBOM</button>
      </div>
      <footer>
        <button class="button" data-action="close">Đóng</button>
      </footer>
    </div>
  `;

  document.body.appendChild(overlay);
  overlay.querySelector('[data-action="close"]').addEventListener('click', () => overlay.remove());
  overlay.querySelector('[data-action="open-github"]').addEventListener('click', async () => {
    try {
      await window.inventoryApi.openGithub();
    } catch (error) {
      showToast(`Không thể mở GitHub: ${error.message}`);
    }
  });
}

function showLicenseDialog() {
  const overlay = document.createElement('div');
  overlay.className = 'sheet-modal-overlay';
  overlay.innerHTML = `
    <div class="info-modal">
      <header>
        <h2>Nhập License</h2>
        <p>Nhập mã license dạng XXX-XXX-XXX-XXXX để kích hoạt bản vĩnh viễn.</p>
      </header>
      <div class="info-content">
        <div class="license-form">
          <input class="license-input" type="text" placeholder="XXX-XXX-XXX-XXXX" maxlength="16">
          <button class="button button-primary" data-action="activate">Kích hoạt</button>
        </div>
        <div class="license-message"></div>
      </div>
      <footer>
        <button class="button" data-action="close">Đóng</button>
      </footer>
    </div>
  `;

  document.body.appendChild(overlay);
  const input = overlay.querySelector('.license-input');
  const message = overlay.querySelector('.license-message');
  input.focus();

  overlay.querySelector('[data-action="close"]').addEventListener('click', () => overlay.remove());
  overlay.querySelector('[data-action="activate"]').addEventListener('click', async () => {
    const result = await activateLicenseCode(input.value);
    message.textContent = result.message;
    message.classList.toggle('ok', result.ok);
    if (result.ok) {
      setTimeout(() => overlay.remove(), 800);
    }
  });
}

async function activateExpiredLicense() {
  const result = await activateLicenseCode(els.expiredLicenseInput.value);
  els.expiredLicenseMessage.textContent = result.message;
  els.expiredLicenseMessage.classList.toggle('ok', result.ok);
}

async function activateLicenseCode(code) {
  const result = await window.inventoryApi.activateLicense(code);
  state.licenseStatus = result.status;
  renderLicenseStatus();
  if (result.ok) {
    els.expiredLicenseInput.value = '';
  }
  return result;
}

async function loadSelectedExcelFiles() {
  const fileInfos = normalizeLoadedFiles(await window.inventoryApi.openExcel());
  if (!fileInfos.length) return [];

  const selections = await selectExcelSheets(fileInfos);
  if (!selections.length) return [];

  return window.inventoryApi.readExcelSheets(selections);
}

async function selectExcelSheets(fileInfos) {
  const multiSheetFiles = fileInfos.filter((file) => (file.sheets || []).length > 1);
  if (!multiSheetFiles.length) {
    return fileInfos.map((file) => ({
      filePath: file.filePath,
      sheetName: (file.sheets || [])[0] || ''
    }));
  }

  return showSheetPicker(fileInfos);
}

function showSheetPicker(fileInfos) {
  return new Promise((resolve) => {
    const overlay = document.createElement('div');
    overlay.className = 'sheet-modal-overlay';
    const rows = fileInfos.map((file, index) => {
      const options = (file.sheets || []).map((sheet, sheetIndex) => `
        <label class="sheet-option">
          <input type="checkbox" data-file-index="${index}" value="${escapeHtml(sheet)}"${sheetIndex === 0 ? ' checked' : ''}>
          <span>${escapeHtml(sheet)}</span>
        </label>
      `).join('');
      return `
        <div class="sheet-picker-row">
          <div class="sheet-file-name">${escapeHtml(file.fileName)}</div>
          <div class="sheet-options">${options}</div>
        </div>
      `;
    }).join('');

    overlay.innerHTML = `
      <div class="sheet-modal">
        <header>
          <h2>Chọn sheet để load</h2>
          <p>File có nhiều sheet cần chọn sheet dùng để so sánh.</p>
        </header>
        <div class="sheet-picker-list">${rows}</div>
        <footer>
          <button class="button" data-action="cancel">Hủy</button>
          <button class="button button-primary" data-action="confirm">Load sheet đã chọn</button>
        </footer>
      </div>
    `;

    document.body.appendChild(overlay);

    overlay.querySelector('[data-action="cancel"]').addEventListener('click', () => {
      overlay.remove();
      resolve([]);
    });

    overlay.querySelector('[data-action="confirm"]').addEventListener('click', () => {
      const selections = fileInfos.flatMap((file, index) => {
        const checkedSheets = Array.from(overlay.querySelectorAll(`input[data-file-index="${index}"]:checked`));
        return checkedSheets.map((input) => ({
          filePath: file.filePath,
          sheetName: input.value
        }));
      });

      if (!selections.length) {
        showToast('Hãy chọn ít nhất một sheet để load.');
        return;
      }

      overlay.remove();
      resolve(selections);
    });
  });
}

async function restoreRecentFiles() {
  try {
    const recent = await window.inventoryApi.loadRecent();
    const khoFiles = recent.khoFiles || [];
    const bomFiles = recent.bomFiles || [];
    if (!khoFiles.length && !bomFiles.length) return;

    state.khoSources = khoFiles.map(toFileSource);
    state.bomSources = bomFiles.map(toFileSource);
    state.khoPaths = state.khoSources.map((file) => file.filePath);
    state.bomPaths = state.bomSources.map((file) => file.filePath);
    state.khoRows = renumberRows(khoFiles.flatMap((file) => parseKhoRows(file.rows)));
    state.bomRows = renumberRows(bomFiles.flatMap((file) => parseBomRows(file.rows)));
    state.khoFileName = khoFiles.map(formatFileLabel).join('; ');
    state.bomFileName = bomFiles.map(formatFileLabel).join('; ');
    state.compareRows = [];
    state.discrepancyRows = [];
    state.confirmRows = [];
    state.manualMatches.clear();
    state.rejectedMatches.clear();

    runCompare();
    setView(state.khoRows.length || state.bomRows.length ? 'discrepancy' : 'dashboard');
    showToast('Đã khôi phục đường dẫn và dữ liệu file lần trước.');
  } catch (error) {
    showToast(`Không thể khôi phục file lần trước: ${error.message}`);
  }
}

function normalizeLoadedFiles(result) {
  if (!result) return [];
  return Array.isArray(result) ? result : [result];
}

function appendPaths(paths, nextPaths) {
  return Array.from(new Set(paths.concat(nextPaths)));
}

function appendSources(sources, nextSources) {
  const map = new Map();
  sources.concat(nextSources).forEach((source) => {
    map.set(`${source.filePath}::${source.sheetName}`, source);
  });
  return Array.from(map.values());
}

function toFileSource(file) {
  return {
    filePath: file.filePath,
    sheetName: file.sheetName || ''
  };
}

function formatFileLabel(file) {
  return `${file.fileName} / ${file.sheetName}`;
}

async function saveRecentFiles() {
  await window.inventoryApi.saveRecent({
    khoSources: state.khoSources,
    bomSources: state.bomSources,
    khoPaths: state.khoPaths,
    bomPaths: state.bomPaths
  });
}

function autoCompareAfterLoad(fallbackView) {
  runCompare();
  setView(state.khoRows.length && state.bomRows.length ? 'discrepancy' : fallbackView);
}

function appendFileName(current, next) {
  if (!current) return next;
  return `${current}; ${next}`;
}

function formatLoadedFiles(value) {
  return value ? value.replace(/; /g, '\n') : 'Chưa load file.';
}

function renumberRows(rows) {
  return rows.map((row, index) => ({ ...row, stt: index + 1 }));
}

function clearData() {
  state.search = '';
  state.khoFileName = '';
  state.bomFileName = '';
  state.khoPaths = [];
  state.bomPaths = [];
  state.khoSources = [];
  state.bomSources = [];
  state.khoRows = [];
  state.bomRows = [];
  state.compareRows = [];
  state.discrepancyRows = [];
  state.confirmRows = [];
  state.manualMatches.clear();
  state.rejectedMatches.clear();
  els.searchInput.value = '';
  window.inventoryApi.clearRecent();
  setView('dashboard');
  renderAll();
  showToast('Đã clear toàn bộ dữ liệu đã load.');
}

function parseKhoRows(rows) {
  return rows
    .map((row) => splitKhoCell(row[0]))
    .filter(Boolean)
    .map(normalizeKhoParts)
    .map((parts, index) => ({
      stt: index + 1,
      projectCode: parts[0],
      drawingCode: parts[1],
      drawingKey: normalizeCode(parts[1]),
      quantity: parseQuantity(parts[2]),
      manufacturer: parts[3],
      importCode: parts[4],
      importDate: parts[5],
      note: parts.slice(6).join(', ')
    }))
    .filter((row) => row.drawingKey);
}

function splitKhoCell(value) {
  if (value == null) return null;
  const text = String(value).trim();
  if (!text || !text.includes(',')) return null;
  const parts = text.split(',').map((part) => part.trim());
  if (parts.length < 4) return null;
  while (parts.length < 7) parts.push('');
  return parts;
}

function normalizeKhoParts(parts) {
  const normalized = [...parts];
  const quantityCell = normalized[2];
  const manufacturerCell = normalized[3];

  if (!isQuantityCell(quantityCell) && isQuantityCell(manufacturerCell)) {
    normalized[2] = manufacturerCell;
    normalized[3] = quantityCell;
  }

  return normalized;
}

function isQuantityCell(value) {
  const text = cleanCell(value).replace(',', '.');
  return /^-?\d+(\.\d+)?$/.test(text);
}

function parseBomRows(rows) {
  const map = detectBomColumns(rows);
  if (map.drawing < 0 || map.quantity < 0) {
    throw new Error('Không tìm thấy cột Mã bản vẽ hoặc Số lượng/Máy trong BOM.');
  }

  const startIndex = Math.max(map.headerRow + 1, 1);
  return rows
    .slice(startIndex)
    .map((row) => ({
      itemName: cleanCell(row[map.item]),
      drawingCode: cleanCell(row[map.drawing]),
      drawingKey: normalizeCode(row[map.drawing]),
      manufacturer: cleanCell(row[map.manufacturer]),
      quantity: parseQuantity(getBomQuantityCell(row, map.quantity))
    }))
    .filter((row) => row.drawingKey)
    .map((row, index) => ({ ...row, stt: index + 1 }));
}

function detectBomColumns(rows) {
  const map = {
    item: 1,
    drawing: 2,
    manufacturer: 3,
    quantity: 8,
    headerRow: 0
  };
  const found = {
    item: false,
    drawing: false,
    manufacturer: false,
    quantity: false
  };

  rows.slice(0, 20).forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      const text = normalizeText(cell);
      if (!found.item && isBomItemHeader(text)) {
        map.item = colIndex;
        found.item = true;
        map.headerRow = Math.max(map.headerRow, rowIndex);
      }
      if (!found.drawing && (text.includes('ma ban ve') || text.includes('model'))) {
        map.drawing = colIndex;
        found.drawing = true;
        map.headerRow = Math.max(map.headerRow, rowIndex);
      }
      if (!found.manufacturer && isBomManufacturerHeader(text)) {
        map.manufacturer = colIndex;
        found.manufacturer = true;
        map.headerRow = Math.max(map.headerRow, rowIndex);
      }
      if (!found.quantity && isBomQuantityHeader(text)) {
        map.quantity = colIndex;
        found.quantity = true;
        map.headerRow = Math.max(map.headerRow, rowIndex);
      }
    });
  });

  return map;
}

function isBomItemHeader(text) {
  return text.includes('ten mat hang') || matchesHeader(text, ['name', 'item name', 'part name']);
}

function isBomManufacturerHeader(text) {
  return text.includes('nha san xuat') || matchesHeader(text, ['maker', 'manufacturer']);
}

function isBomQuantityHeader(text) {
  if (text.includes('ton') || text.includes('number machine') || text.includes('grand total')) {
    return false;
  }

  return (
    text.includes('so luong') ||
    matchesHeader(text, [
      "q'ty/machine",
      "q'ty / machine",
      'qty/machine',
      'qty / machine',
      'quantity/machine',
      'quantity / machine',
      'qty per machine',
      'quantity per machine'
    ])
  );
}

function matchesHeader(text, aliases) {
  const compactText = text.replace(/\s+/g, '');
  return aliases.some((alias) => {
    const compactAlias = alias.replace(/\s+/g, '');
    return (
      text === alias ||
      text.includes(alias) ||
      compactText === compactAlias ||
      compactText.includes(compactAlias)
    );
  });
}

function getBomQuantityCell(row, detectedQuantityIndex) {
  const fixedQuantity = row[8];
  if (parseQuantity(fixedQuantity) > 0 || cleanCell(fixedQuantity) === '0') {
    return fixedQuantity;
  }
  return row[detectedQuantityIndex];
}

function runCompare() {
  const autoThreshold = clamp(Number(els.autoThreshold.value) / 100, 0.8, 1);
  const confirmThreshold = clamp(Number(els.confirmThreshold.value) / 100, 0.5, autoThreshold);
  const khoItems = aggregateRows(state.khoRows, 'kho');
  const bomItems = aggregateRows(state.bomRows, 'bom');
  if (!khoItems.length || !bomItems.length) {
    state.compareRows = [];
    state.confirmRows = [];
    state.discrepancyRows = createSingleSourceDiscrepancyRows(khoItems, bomItems);
    renderAll();
    return;
  }

  const bomByKey = new Map(bomItems.map((item) => [item.drawingKey, item]));
  const allRows = [];
  const compareRows = [];
  const matchedBomKeys = new Set();
  const confirmRows = [];
  const confirmPairKeys = new Set();

  khoItems.forEach((kho) => {
    const manualBomKey = state.manualMatches.get(kho.drawingKey);
    if (manualBomKey && bomByKey.has(manualBomKey)) {
      const bom = bomByKey.get(manualBomKey);
      matchedBomKeys.add(bom.drawingKey);
      allRows.push(createCompareRow(bom, kho, 1));
      return;
    }

    const exactBom = bomByKey.get(kho.drawingKey);
    if (exactBom) {
      matchedBomKeys.add(exactBom.drawingKey);
      allRows.push(createCompareRow(exactBom, kho, 1));
      return;
    }

    const candidates = findFuzzyMatches(kho, bomItems, confirmThreshold, 3)
      .filter((candidate) => !state.rejectedMatches.has(`${kho.drawingKey}::${candidate.item.drawingKey}`));
    const candidate = candidates[0];

    if (!candidate) {
      allRows.push(createCompareRow(null, kho, 0));
      return;
    }

    if (candidate.similarity >= autoThreshold) {
      matchedBomKeys.add(candidate.item.drawingKey);
      allRows.push(createCompareRow(candidate.item, kho, candidate.similarity));
      return;
    }

    if (candidate.similarity >= confirmThreshold) {
      if (!confirmPairKeys.has(kho.drawingKey)) {
        confirmPairKeys.add(kho.drawingKey);
        confirmRows.push(createConfirmRow(kho, candidates));
      }
      allRows.push(createCompareRow(null, kho, 0));
      return;
    }

    allRows.push(createCompareRow(null, kho, 0));
  });

  state.compareRows = allRows
    .filter((row) => row.status === 'Đủ')
    .map((row, index) => ({ ...row, stt: index + 1 }));
  state.confirmRows = confirmRows.map((row, index) => ({ ...row, stt: index + 1 }));
  state.discrepancyRows = createDiscrepancyRows(allRows, bomItems, matchedBomKeys);
  renderAll();
  showToast(`So sánh xong: ${state.compareRows.length} dòng, ${state.discrepancyRows.length} dòng thiếu/thừa, ${state.confirmRows.length} dòng cần xác nhận.`);
}

function createSingleSourceDiscrepancyRows(khoItems, bomItems) {
  const rows = [];

  if (khoItems.length && !bomItems.length) {
    khoItems.forEach((kho) => {
      rows.push({
        source: 'Chỉ có trong kho',
        bomDrawingCode: '',
        khoDrawingCode: kho.drawingCode,
        itemName: '',
        manufacturer: kho.manufacturer,
        bomQuantity: 0,
        khoQuantity: kho.quantity,
        difference: kho.quantity,
        status: 'Thừa',
        note: 'Chưa load Dữ Liệu Thiết Kế'
      });
    });
  }

  if (bomItems.length && !khoItems.length) {
    bomItems.forEach((bom) => {
      rows.push({
        source: 'Chỉ có trong BOM',
        bomDrawingCode: bom.drawingCode,
        khoDrawingCode: '',
        itemName: bom.itemName,
        manufacturer: bom.manufacturer,
        bomQuantity: bom.quantity,
        khoQuantity: 0,
        difference: -bom.quantity,
        status: 'Thiếu',
        note: 'Chưa load Dữ Liệu Kho'
      });
    });
  }

  return rows.map((row, index) => ({ ...row, stt: index + 1 }));
}

function createDiscrepancyRows(allRows, bomItems, matchedBomKeys) {
  const rows = [];

  allRows.forEach((row) => {
    if (row.status === 'Đủ') return;
    rows.push({
      source: row.bomKey ? 'Lệch số lượng' : 'Chỉ có trong kho',
      bomDrawingCode: row.bomDrawingCode,
      khoDrawingCode: row.khoDrawingCode,
      itemName: row.itemName,
      manufacturer: row.manufacturer,
      bomQuantity: row.bomQuantity,
      khoQuantity: row.khoQuantity,
      difference: row.difference,
      status: row.status,
      note: row.bomKey ? 'Số lượng kho và BOM không bằng nhau' : 'Mã có trong kho, không có trong BOM'
    });
  });

  bomItems.forEach((bom) => {
    if (matchedBomKeys.has(bom.drawingKey)) return;
    rows.push({
      source: 'Chỉ có trong BOM',
      bomDrawingCode: bom.drawingCode,
      khoDrawingCode: '',
      itemName: bom.itemName,
      manufacturer: bom.manufacturer,
      bomQuantity: bom.quantity,
      khoQuantity: 0,
      difference: -bom.quantity,
      status: 'Thiếu',
      note: 'Mã có trong BOM, không có trong kho'
    });
  });

  return rows.map((row, index) => ({ ...row, stt: index + 1 }));
}

function aggregateRows(rows, type) {
  const map = new Map();
  rows.forEach((row) => {
    if (!row.drawingKey) return;
    if (!map.has(row.drawingKey)) {
      map.set(row.drawingKey, {
        drawingKey: row.drawingKey,
        drawingCode: row.drawingCode,
        quantity: 0,
        mergeCount: 0,
        itemName: type === 'bom' ? row.itemName : '',
        manufacturer: row.manufacturer || ''
      });
    }
    const item = map.get(row.drawingKey);
    item.quantity += Number(row.quantity) || 0;
    item.mergeCount += 1;
    if (!item.itemName && row.itemName) item.itemName = row.itemName;
    if (!item.manufacturer && row.manufacturer) item.manufacturer = row.manufacturer;
  });
  return Array.from(map.values());
}

function findFuzzyMatches(source, candidates, minSimilarity, limit) {
  if (candidates.length === 0) return [];

  return candidates
    .map((item) => ({ item, similarity: stringSimilarity(source.drawingKey, item.drawingKey) }))
    .filter((match) => match.similarity >= minSimilarity)
    .sort((a, b) => b.similarity - a.similarity)
    .slice(0, limit);
}

function createCompareRow(bom, kho, similarity) {
  const bomQuantity = bom ? bom.quantity : 0;
  const khoQuantity = kho ? kho.quantity : 0;
  const difference = khoQuantity - bomQuantity;
  let status = 'Đủ';
  if (!bom && kho) status = 'Thừa';
  else if (difference < 0) status = 'Thiếu';
  else if (difference > 0) status = 'Thừa';

  return {
    bomKey: bom ? bom.drawingKey : '',
    khoKey: kho ? kho.drawingKey : '',
    bomDrawingCode: bom ? bom.drawingCode : '',
    khoDrawingCode: kho ? kho.drawingCode : '',
    itemName: bom ? bom.itemName : '',
    manufacturer: bom?.manufacturer || kho?.manufacturer || '',
    bomQuantity,
    khoQuantity,
    difference,
    status,
    similarity: similarity ? `${Math.round(similarity * 100)}%` : '',
    mergeNote: kho && kho.mergeCount > 1 ? `Có gộp ${kho.mergeCount} dòng, tổng SL ${kho.quantity}` : ''
  };
}

function createConfirmRow(kho, candidates) {
  const options = candidates.map((candidate) => ({
    bomKey: candidate.item.drawingKey,
    bomDrawingCode: candidate.item.drawingCode,
    itemName: candidate.item.itemName,
    bomQuantity: candidate.item.quantity,
    similarity: `${Math.round(candidate.similarity * 100)}%`
  }));
  const selected = options[0] || {};

  return {
    khoKey: kho.drawingKey,
    khoDrawingCode: kho.drawingCode,
    bomKey: selected.bomKey || '',
    bomDrawingCode: selected.bomDrawingCode || '',
    itemName: selected.itemName || '',
    bomQuantity: selected.bomQuantity || 0,
    khoQuantity: kho.quantity,
    similarity: selected.similarity || '',
    selectedIndex: 0,
    options
  };
}

function acceptConfirm(index) {
  const row = state.confirmRows[index];
  if (!row) return;
  const selected = row.options[row.selectedIndex] || row.options[0];
  if (!selected) return;
  state.manualMatches.set(row.khoKey, selected.bomKey);
  runCompare();
}

function rejectConfirm(index) {
  const row = state.confirmRows[index];
  if (!row) return;
  row.options.forEach((option) => {
    state.rejectedMatches.add(`${row.khoKey}::${option.bomKey}`);
  });
  runCompare();
}

async function exportCurrentTable() {
  try {
    if (state.view === 'compare') {
      if (!state.compareRows.length && !state.confirmRows.length) {
        showToast('Chưa có dữ liệu So Sánh để xuất.');
        return;
      }

      const filePath = await window.inventoryApi.exportCompare({
        compareRows: state.compareRows,
        confirmRows: state.confirmRows
      });
      if (filePath) showToast(`Đã xuất bảng So Sánh: ${filePath}`);
      return;
    }

    if (state.view === 'discrepancy') {
      if (!state.discrepancyRows.length) {
        showToast('Chưa có dữ liệu Thiếu Thừa để xuất.');
        return;
      }

      const filePath = await window.inventoryApi.exportDiscrepancy({
        discrepancyRows: state.discrepancyRows
      });
      if (filePath) showToast(`Đã xuất bảng Thiếu Thừa: ${filePath}`);
      return;
    }

    showToast('Hãy chọn bảng So Sánh hoặc Thiếu Thừa trước khi xuất Excel.');
  } catch (error) {
    showToast(`Không thể xuất Excel: ${error.message}`);
  }
}

function setView(view) {
  state.view = view;
  document.querySelectorAll('.nav-item').forEach((button) => {
    button.classList.toggle('active', button.dataset.view === view);
  });
  document.querySelectorAll('.view').forEach((section) => {
    section.classList.toggle('active', section.id === `${view}View`);
  });
  renderAll();
}

function renderAll() {
  els.khoFileName.textContent = formatLoadedFiles(state.khoFileName);
  els.bomFileName.textContent = formatLoadedFiles(state.bomFileName);
  els.loadKhoBtn.textContent = state.khoRows.length ? 'Thêm Dữ Liệu Kho' : 'Dữ Liệu Kho';
  els.loadBomBtn.textContent = state.bomRows.length ? 'Thêm Dữ Liệu Thiết Kế' : 'Dữ Liệu Thiết Kế';
  els.exportExcelBtn.disabled = !canExportCurrentTable();

  const okCount = state.compareRows.filter((row) => row.status === 'Đủ').length;
  const missingCount = state.discrepancyRows.filter((row) => row.status === 'Thiếu').length;
  const extraCount = state.discrepancyRows.filter((row) => row.status === 'Thừa').length;

  els.khoCount.textContent = state.khoRows.length;
  els.bomCount.textContent = state.bomRows.length;
  els.sideKhoCount.textContent = state.khoRows.length;
  els.sideBomCount.textContent = state.bomRows.length;
  els.okCount.textContent = okCount;
  els.missingCount.textContent = missingCount;
  els.extraCount.textContent = extraCount;
  els.confirmCount.textContent = state.confirmRows.length;

  renderCurrentTable();
}

function renderCurrentTable() {
  renderTable(els.khoTable, filterRows(state.khoRows), columns.kho);
  renderTable(els.bomTable, filterRows(state.bomRows), columns.bom);
  renderTable(els.compareTable, filterRows(state.compareRows), columns.compare, statusClass);
  renderTable(els.discrepancyTable, filterRows(state.discrepancyRows), columns.discrepancy, statusClass);
  renderConfirmTable();
}

function canExportCurrentTable() {
  if (state.view === 'compare') return state.compareRows.length > 0 || state.confirmRows.length > 0;
  if (state.view === 'discrepancy') return state.discrepancyRows.length > 0;
  return false;
}

function renderTable(container, rows, tableColumns, rowClassFn) {
  if (!rows.length) {
    container.innerHTML = '<div class="placeholder">Chưa có dữ liệu.</div>';
    return;
  }

  const thead = tableColumns.map(([, label, align]) => `<th class="${align || ''}">${escapeHtml(label)}</th>`).join('');
  const tbody = rows.map((row) => {
    const cells = tableColumns.map(([key, , align]) => `<td class="${align || ''}">${escapeHtml(row[key])}</td>`).join('');
    return `<tr class="${rowClassFn ? rowClassFn(row) : ''}">${cells}</tr>`;
  }).join('');

  container.innerHTML = `<table><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table>`;
}

function renderConfirmTable() {
  if (state.view !== 'compare') {
    els.confirmTable.innerHTML = '';
    return;
  }

  const rows = filterRows(state.confirmRows);
  if (!rows.length) {
    els.confirmTable.innerHTML = '<div class="placeholder">Không có mã cần xác nhận.</div>';
    return;
  }

  const thead = columns.confirm.map(([, label, align]) => `<th class="${align || ''}">${escapeHtml(label)}</th>`).join('');
  const tbody = rows.map((row) => {
    const actualIndex = state.confirmRows.findIndex((item) => item.khoKey === row.khoKey);
    const cells = columns.confirm.map(([key, , align]) => {
      if (key === 'actions') {
        return `<td><div class="row-actions"><button class="mini-button confirm" data-action="accept" data-index="${actualIndex}">Chọn</button><button class="mini-button reject" data-action="reject" data-index="${actualIndex}">Bỏ qua</button></div></td>`;
      }
      if (key === 'bomSelect') {
        const options = row.options.map((option, optionIndex) => {
          const selected = optionIndex === row.selectedIndex ? ' selected' : '';
          return `<option value="${optionIndex}"${selected}>${escapeHtml(option.bomDrawingCode)} - ${escapeHtml(option.similarity)}</option>`;
        }).join('');
        return `<td><select class="bom-select" data-index="${actualIndex}">${options}</select></td>`;
      }
      return `<td class="${align || ''}">${escapeHtml(row[key])}</td>`;
    }).join('');
    return `<tr>${cells}</tr>`;
  }).join('');

  els.confirmTable.innerHTML = `<table><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table>`;
  els.confirmTable.querySelectorAll('select[data-index]').forEach((select) => {
    select.addEventListener('change', () => {
      updateConfirmSelection(Number(select.dataset.index), Number(select.value));
      renderConfirmTable();
    });
  });
  els.confirmTable.querySelectorAll('button[data-action]').forEach((button) => {
    button.addEventListener('click', () => {
      const index = Number(button.dataset.index);
      if (button.dataset.action === 'accept') acceptConfirm(index);
      else rejectConfirm(index);
    });
  });
}

function updateConfirmSelection(index, selectedIndex) {
  const row = state.confirmRows[index];
  if (!row || !row.options[selectedIndex]) return;
  const selected = row.options[selectedIndex];
  row.selectedIndex = selectedIndex;
  row.bomKey = selected.bomKey;
  row.bomDrawingCode = selected.bomDrawingCode;
  row.itemName = selected.itemName;
  row.bomQuantity = selected.bomQuantity;
  row.similarity = selected.similarity;
}

function filterRows(rows) {
  if (!state.search) return rows;
  const query = normalizeText(state.search);
  return rows.filter((row) => normalizeText(Object.values(row).join(' ')).includes(query));
}

function statusClass(row) {
  if (row.status === 'Đủ') return 'status-ok';
  if (row.status === 'Thiếu') return 'status-missing';
  if (row.status === 'Thừa') return 'status-extra';
  return '';
}

function handleThresholdChange() {
  const confirmValue = Number(els.confirmThreshold.value);
  const autoValue = Number(els.autoThreshold.value);
  if (confirmValue >= autoValue) {
    els.confirmThreshold.value = Math.max(50, autoValue - 1);
  }
  if (state.khoRows.length && state.bomRows.length) {
    runCompare();
  }
}

function normalizeCode(value) {
  return cleanCell(value)
    .toUpperCase()
    .replace(/[‐‑‒–—―]/g, '-')
    .replace(/\s+/g, '')
    .trim();
}

function normalizeText(value) {
  return cleanCell(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toLowerCase();
}

function cleanCell(value) {
  if (value == null) return '';
  return String(value).replace(/\s+/g, ' ').trim();
}

function parseQuantity(value) {
  const text = cleanCell(value).replace(',', '.');
  const parsed = Number.parseFloat(text);
  return Number.isFinite(parsed) ? parsed : 0;
}

function stringSimilarity(a, b) {
  if (a === b) return 1;
  if (!a || !b) return 0;
  const distance = levenshtein(a, b);
  return 1 - distance / Math.max(a.length, b.length);
}

function levenshtein(a, b) {
  const matrix = Array.from({ length: a.length + 1 }, () => []);
  for (let i = 0; i <= a.length; i += 1) matrix[i][0] = i;
  for (let j = 0; j <= b.length; j += 1) matrix[0][j] = j;

  for (let i = 1; i <= a.length; i += 1) {
    for (let j = 1; j <= b.length; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost
      );
    }
  }
  return matrix[a.length][b.length];
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function showToast(message) {
  els.toast.textContent = message;
  els.toast.hidden = false;
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => {
    els.toast.hidden = true;
  }, 3800);
}
