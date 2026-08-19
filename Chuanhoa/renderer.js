const state = {
  view: 'dashboard',
  search: '',
  appVersion: '',
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
  rejectedMatches: new Set(),
  showBomClearPanel: false,
  khoPage: 1
};

const columns = {
  kho: [
    ['stt', 'STT', 'number'],
    ['drawingCode', 'Mã bản vẽ'],
    ['itemName', 'Tên hàng'],
    ['unit', 'Đơn vị tính']
  ],
  bom: [
    ['stt', 'STT', 'number'],
    ['itemName', 'Tên mặt hàng'],
    ['drawingCode', 'Mã bản vẽ'],
    ['manufacturer', 'Nhà sản xuất'],
    ['quantity', 'Số lượng/máy', 'number'],
    ['unit', 'Đơn vị tính']
  ],
  compare: [
    ['stt', 'STT', 'number'],
    ['orderDrawingCode', 'Mã đã đặt hàng'],
    ['designDrawingCode', 'Mã thiết kế'],
    ['orderItemName', 'Tên mã đặt hàng'],
    ['designUnit', 'ĐVT thiết kế'],
    ['orderUnit', 'ĐVT đã đặt hàng'],
    ['note', 'Ghi chú']
  ],
  discrepancy: [
    ['stt', 'STT', 'number'],
    ['designDrawingCode', 'Mã bản vẽ'],
    ['suggestedOrderDrawingCode', 'Mã đã đặt hàng gợi ý'],
    ['suggestedOrderItemName', 'Tên hàng gợi ý'],
    ['nameSimilarity', 'Độ giống tên hàng'],
    ['note', 'Ghi chú']
  ],
  confirm: [
    ['stt', 'STT', 'number'],
    ['designDrawingCode', 'Mã thiết kế'],
    ['orderSelect', 'Mã đã đặt hàng đề xuất'],
    ['itemName', 'Tên mặt hàng'],
    ['designUnit', 'ĐVT thiết kế'],
    ['orderUnit', 'ĐVT đã đặt hàng'],
    ['similarity', 'Độ tương đồng mã'],
    ['actions', 'Hành động']
  ]
};

const els = {};
const MAX_VISIBLE_ROWS = 1500;

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
    'clearKhoBtn',
    'clearBomBtn',
    'clearBomPanel',
    'clearBtn',
    'updateBtn',
    'infoBtn',
    'exportExcelBtn',
    'searchInput',
    'autoThreshold',
    'confirmThreshold',
    'appVersion',
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
  els.clearKhoBtn.addEventListener('click', clearOrderData);
  els.clearBomBtn.addEventListener('click', toggleBomClearPanel);
  els.clearBomPanel.addEventListener('click', handleBomClearPanelClick);
  els.khoTable.addEventListener('click', handleKhoTableClick);
  els.clearBtn.addEventListener('click', clearData);
  els.updateBtn.addEventListener('click', checkForUpdates);
  els.infoBtn.addEventListener('click', openGithubInfo);
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
    state.khoPage = 1;
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
  } catch {
    state.appVersion = '';
  }
}

async function loadKhoFile() {
  try {
    const files = await loadSelectedExcelFiles();
    if (!files.length) return;

    const parsedRows = files.flatMap((file) => parseOrderRows(file));
    state.khoSources = appendSources([], files.map(toFileSource));
    state.khoPaths = appendPaths([], files.map((file) => file.filePath));
    state.khoRows = renumberRows(parsedRows);
    state.khoFileName = appendFileName('', files.map(formatOrderFileLabel).join('; '));
    state.khoPage = 1;
    state.compareRows = [];
    state.discrepancyRows = [];
    state.confirmRows = [];
    state.manualMatches.clear();
    state.rejectedMatches.clear();

    await saveRecentFiles();
    autoCompareAfterLoad('kho');
    showToast(`Đã load ${parsedRows.length} mã đã đặt hàng từ ${files.length} file.`);
  } catch (error) {
    showToast(`Không thể load file Mã Đã Đặt Hàng: ${error.message}`);
  }
}

async function loadBomFile() {
  try {
    const files = await loadSelectedExcelFiles();
    if (!files.length) return;

    const parsedRows = files.flatMap((file) => parseBomRows(file));
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
    state.khoRows = renumberRows(khoFiles.flatMap((file) => parseOrderRows(file)));
    state.bomRows = renumberRows(bomFiles.flatMap((file) => parseBomRows(file)));
    state.khoFileName = khoFiles.map(formatKhoFileLabel).join('; ');
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

function formatOrderFileLabel(file) {
  return `${file.fileName} / Mã Đã Đặt Hàng`;
}

function formatKhoFileLabel(file) {
  return tryParseOrderRows(file.rows).length ? formatOrderFileLabel(file) : formatFileLabel(file);
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

function getFileNameFromPath(filePath) {
  return cleanCell(filePath).split(/[\\/]/).pop() || filePath;
}

function getBomLoadedFiles() {
  const map = new Map();
  state.bomRows.forEach((row) => {
    const filePath = row.sourceFilePath || '';
    if (!filePath || map.has(filePath)) return;
    const fileName = row.sourceFileName || getFileNameFromPath(filePath);
    map.set(filePath, {
      filePath,
      fileName,
      label: `${fileName} / Dữ Liệu Thiết Kế`
    });
  });
  return Array.from(map.values());
}

function renderBomClearPanel() {
  const files = getBomLoadedFiles();
  els.clearBomPanel.hidden = !state.showBomClearPanel || !files.length;
  if (els.clearBomPanel.hidden) {
    els.clearBomPanel.innerHTML = '';
    return;
  }

  els.clearBomPanel.innerHTML = files.map((file) => (
    `<button class="mini-button reject" data-file-path="${escapeHtml(file.filePath)}">Clear ${escapeHtml(file.fileName)}</button>`
  )).join('');
}

function renumberRows(rows) {
  return rows.map((row, index) => ({ ...row, stt: index + 1 }));
}

function clearCompareState() {
  state.compareRows = [];
  state.discrepancyRows = [];
  state.confirmRows = [];
  state.manualMatches.clear();
  state.rejectedMatches.clear();
}

async function refreshAfterPartialClear(fallbackView) {
  clearCompareState();
  await saveRecentFiles();
  runCompare();
  if (!state.khoRows.length && !state.bomRows.length) {
    setView('dashboard');
    return;
  }
  setView(state.khoRows.length && state.bomRows.length ? 'discrepancy' : fallbackView);
}

async function clearOrderData() {
  state.khoFileName = '';
  state.khoPaths = [];
  state.khoSources = [];
  state.khoRows = [];
  state.khoPage = 1;
  await refreshAfterPartialClear(state.bomRows.length ? 'bom' : 'dashboard');
  showToast('Đã clear riêng dữ liệu Mã Đã Đặt Hàng.');
}

function toggleBomClearPanel() {
  if (!state.bomRows.length) return;
  state.showBomClearPanel = !state.showBomClearPanel;
  renderBomClearPanel();
}

function handleBomClearPanelClick(event) {
  const button = event.target.closest('button[data-file-path]');
  if (!button) return;
  clearBomFile(button.dataset.filePath);
}

async function clearBomFile(filePath) {
  const file = getBomLoadedFiles().find((item) => item.filePath === filePath);
  state.bomRows = renumberRows(state.bomRows.filter((row) => row.sourceFilePath !== filePath));
  state.bomSources = state.bomSources.filter((source) => source.filePath !== filePath);
  state.bomPaths = state.bomSources.map((source) => source.filePath);
  state.bomFileName = getBomLoadedFiles().map((item) => item.label).join('; ');
  if (!state.bomRows.length) state.showBomClearPanel = false;
  await refreshAfterPartialClear(state.khoRows.length ? 'kho' : 'dashboard');
  showToast(`Đã clear file thiết kế: ${file?.fileName || getFileNameFromPath(filePath)}.`);
}

function clearData() {
  state.search = '';
  state.showBomClearPanel = false;
  state.khoPage = 1;
  state.khoFileName = '';
  state.bomFileName = '';
  state.khoPaths = [];
  state.bomPaths = [];
  state.khoSources = [];
  state.bomSources = [];
  state.khoRows = [];
  state.bomRows = [];
  clearCompareState();
  els.searchInput.value = '';
  window.inventoryApi.clearRecent();
  setView('dashboard');
  renderAll();
  showToast('Đã clear toàn bộ dữ liệu đã load.');
}

function parseKhoRows(input) {
  return parseOrderRows(input);
}

function parseOrderRows(input) {
  const parsedRows = tryParseOrderRows(input);
  if (!parsedRows.length) {
    throw new Error('Không tìm thấy cột Mã hàng tồn kho trong file Mã Đã Đặt Hàng.');
  }
  return parsedRows;
}

function tryParseOrderRows(input) {
  const { rows, filePath, fileName, sheetName } = normalizeParseInput(input);
  const map = detectOrderColumns(rows);
  if (map.drawing < 0) return [];

  return rows
    .slice(map.headerRow + 1)
    .map((row, index) => {
      const rowNumber = getExcelRowNumber(row, map.headerRow + index + 2);
      return {
        drawingCode: cleanCell(row[map.drawing]),
        drawingKey: normalizeCode(row[map.drawing]),
        itemName: cleanCell(row[map.itemName]),
        unit: cleanCell(row[map.unit]),
        sourceFilePath: filePath,
        sourceFileName: fileName,
        sourceSheetName: sheetName,
        sourceRowNumber: rowNumber,
        sourceNote: formatSourceNote(fileName, sheetName, rowNumber)
      };
    })
    .filter((row) => row.drawingKey)
    .map((row, index) => ({ ...row, stt: index + 1 }));
}

function normalizeParseInput(input) {
  if (Array.isArray(input)) {
    return { rows: input, filePath: '', fileName: '', sheetName: '' };
  }

  return {
    rows: input?.rows || [],
    filePath: input?.filePath || '',
    fileName: input?.fileName || '',
    sheetName: input?.sheetName || ''
  };
}

function getExcelRowNumber(row, fallback) {
  return Number(row?.__rowNumber) || Number(row?.__rowNum__) + 1 || fallback;
}

function formatSourceNote(fileName, sheetName, rowNumber) {
  return [
    fileName ? `File: ${fileName}` : '',
    sheetName ? `Sheet: ${sheetName}` : '',
    rowNumber ? `Dòng: ${rowNumber}` : ''
  ].filter(Boolean).join(', ');
}

function detectOrderColumns(rows) {
  const map = {
    drawing: -1,
    itemName: -1,
    unit: -1,
    headerRow: 0
  };

  rows.slice(0, 20).forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      const text = normalizeText(cell);
      if (map.drawing < 0 && isOrderCodeHeader(text)) {
        map.drawing = colIndex;
        map.headerRow = rowIndex;
      }
      if (map.itemName < 0 && isOrderItemHeader(text)) {
        map.itemName = colIndex;
      }
      if (map.unit < 0 && isOrderUnitHeader(text)) {
        map.unit = colIndex;
      }
    });
  });

  return map;
}

function isOrderCodeHeader(text) {
  return (
    text.includes('ma hang ton kho') ||
    text.includes('ma hang dat hang') ||
    text.includes('ma dat hang') ||
    matchesHeader(text, ['part no', 'part number', 'item code', 'code'])
  );
}

function isOrderItemHeader(text) {
  return text.includes('ten hang') || text.includes('ten mat hang') || matchesHeader(text, ['item name', 'name']);
}

function isOrderUnitHeader(text) {
  return (
    text.includes('don vi tinh') ||
    matchesHeader(text, [
      'dvt',
      'dv tinh',
      'don vi',
      'unit',
      'units',
      'uom',
      'measure unit',
      'unit of measure',
      'unit measure'
    ])
  );
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

function parseBomRows(input) {
  const { rows, filePath, fileName, sheetName } = normalizeParseInput(input);
  const map = detectBomColumns(rows);
  const metadata = extractDesignMetadata(rows, map.headerRow);
  if (map.drawing < 0 || map.quantity < 0) {
    throw new Error('Không tìm thấy cột Mã bản vẽ hoặc Số lượng/Máy trong BOM.');
  }

  const startIndex = Math.max(map.headerRow + 1, 1);
  return rows
    .slice(startIndex)
    .map((row, index) => {
      const rowNumber = getExcelRowNumber(row, startIndex + index + 1);
      return {
        itemName: cleanCell(row[map.item]),
        drawingCode: cleanCell(row[map.drawing]),
        drawingKey: normalizeCode(row[map.drawing]),
        manufacturer: cleanCell(row[map.manufacturer]),
        markingPart: cleanCell(row[map.markingPart]),
        material: cleanCell(row[map.material]),
        surface: cleanCell(row[map.surface]),
        quantity: parseQuantity(getBomQuantityCell(row, map.quantity)),
        hasQuantity: true,
        unit: cleanCell(row[map.unit]),
        ...metadata,
        sourceFilePath: filePath,
        sourceFileName: fileName,
        sourceSheetName: sheetName,
        sourceRowNumber: rowNumber,
        sourceNote: formatSourceNote(fileName, sheetName, rowNumber)
      };
    })
    .filter((row) => row.drawingKey)
    .map((row, index) => ({ ...row, stt: index + 1 }));
}

function detectBomColumns(rows) {
  const map = {
    item: 1,
    drawing: 2,
    manufacturer: 3,
    markingPart: -1,
    material: -1,
    surface: -1,
    quantity: 8,
    unit: -1,
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
      if (map.markingPart < 0 && matchesHeader(text, ['marking part', 'marking'])) {
        map.markingPart = colIndex;
      }
      if (map.material < 0 && matchesHeader(text, ['material', 'vat lieu'])) {
        map.material = colIndex;
      }
      if (map.surface < 0 && matchesHeader(text, ['surface', 'be mat'])) {
        map.surface = colIndex;
      }
      if (!found.quantity && isBomQuantityHeader(text)) {
        map.quantity = colIndex;
        found.quantity = true;
        map.headerRow = Math.max(map.headerRow, rowIndex);
      }
      if (map.unit < 0 && isOrderUnitHeader(text)) {
        map.unit = colIndex;
      }
    });
  });

  return map;
}

function extractDesignMetadata(rows, headerRow) {
  const topRows = rows.slice(0, Math.max(0, headerRow));
  return {
    documentCode: findMetadataValue(topRows, ['ma tai lieu', 'document code']),
    projectCode: findMetadataValue(topRows, ['ma du an', 'project code']),
    projectName: findMetadataValue(topRows, ['ten du an', 'project name']),
    machineQuantity: findMetadataValue(topRows, ['so luong may', 'number machine', 'machine quantity']),
    department: findMetadataValue(topRows, ['bo phan', 'department'])
  };
}

function findMetadataValue(rows, aliases) {
  for (const row of rows) {
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const text = normalizeText(row[colIndex]);
      if (!aliases.some((alias) => matchesHeader(text, [alias]))) continue;

      const inlineValue = extractInlineMetadataValue(row[colIndex]);
      if (inlineValue) return inlineValue;

      for (let nextIndex = colIndex + 1; nextIndex < row.length; nextIndex += 1) {
        const value = cleanCell(row[nextIndex]);
        if (value) return value;
      }
    }
  }

  return '';
}

function extractInlineMetadataValue(value) {
  const text = cleanCell(value);
  const separatorIndex = text.search(/[:：]/);
  if (separatorIndex < 0) return '';
  return text.slice(separatorIndex + 1).trim();
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
  const orderItems = aggregateRows(state.khoRows, 'order');
  const designItems = aggregateRows(state.bomRows, 'design');
  if (!orderItems.length || !designItems.length) {
    state.compareRows = [];
    state.confirmRows = [];
    state.discrepancyRows = createSingleSourceDiscrepancyRows(orderItems, designItems);
    renderAll();
    return;
  }

  const orderByKey = new Map(orderItems.map((item) => [item.drawingKey, item]));
  const orderFuzzyIndex = buildFuzzyCandidateIndex(orderItems);
  const compareRows = [];
  const newCodeRows = [];
  const confirmRows = [];

  designItems.forEach((design) => {
    const manualOrderKey = state.manualMatches.get(design.drawingKey);
    if (manualOrderKey && orderByKey.has(manualOrderKey)) {
      const order = orderByKey.get(manualOrderKey);
      compareRows.push(createCompareRow(design, order, 1, true));
      return;
    }

    const exactOrder = orderByKey.get(design.drawingKey);
    if (exactOrder) {
      if (!unitsMatch(design.unit, exactOrder.unit)) {
        confirmRows.push(createConfirmRow(design, [{ item: exactOrder, similarity: 1 }], 'unit'));
        return;
      }
      compareRows.push(createCompareRow(design, exactOrder, 1, false));
      return;
    }

    const candidates = findFuzzyMatches(design, getFuzzyCandidates(design, orderFuzzyIndex, orderItems), confirmThreshold, 3)
      .filter((candidate) => !state.rejectedMatches.has(`${design.drawingKey}::${candidate.item.drawingKey}`));
    const candidate = candidates[0];

    if (!candidate) {
      newCodeRows.push(createNewCodeRow(design, findOrderNameMatches(design, orderItems, confirmThreshold, 1)[0]));
      return;
    }

    if (candidate.similarity >= autoThreshold) {
      confirmRows.push(createConfirmRow(design, candidates));
      return;
    }

    if (candidate.similarity >= confirmThreshold) {
      confirmRows.push(createConfirmRow(design, candidates));
      return;
    }

    newCodeRows.push(createNewCodeRow(design, findOrderNameMatches(design, orderItems, confirmThreshold, 1)[0]));
  });

  state.compareRows = compareRows.map((row, index) => ({ ...row, stt: index + 1 }));
  state.confirmRows = confirmRows.map((row, index) => ({ ...row, stt: index + 1 }));
  state.discrepancyRows = newCodeRows.map((row, index) => ({ ...row, stt: index + 1 }));
  renderAll();
  showToast(`So sánh xong: ${state.compareRows.length} mã giống nhau, ${state.discrepancyRows.length} mã mới, ${state.confirmRows.length} mã cần xác nhận.`);
}

function createSingleSourceDiscrepancyRows(orderItems, designItems) {
  if (designItems.length && !orderItems.length) {
    return designItems.map((design, index) => ({ ...createNewCodeRow(design), stt: index + 1 }));
  }

  return [];
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
        hasQuantity: row.hasQuantity !== false,
        mergeCount: 0,
        itemName: row.itemName || '',
        manufacturer: row.manufacturer || '',
        markingPart: row.markingPart || '',
        material: row.material || '',
        surface: row.surface || '',
        unit: row.unit || '',
        documentCode: row.documentCode || '',
        projectCode: row.projectCode || '',
        projectName: row.projectName || '',
        machineQuantity: row.machineQuantity || '',
        department: row.department || '',
        sourceNotes: []
      });
    }
    const item = map.get(row.drawingKey);
    item.quantity = (Number(item.quantity) || 0) + (Number(row.quantity) || 0);
    item.mergeCount += 1;
    if (!item.itemName && row.itemName) item.itemName = row.itemName;
    if (!item.manufacturer && row.manufacturer) item.manufacturer = row.manufacturer;
    if (!item.markingPart && row.markingPart) item.markingPart = row.markingPart;
    if (!item.material && row.material) item.material = row.material;
    if (!item.surface && row.surface) item.surface = row.surface;
    if (!item.unit && row.unit) item.unit = row.unit;
    if (!item.documentCode && row.documentCode) item.documentCode = row.documentCode;
    if (!item.projectCode && row.projectCode) item.projectCode = row.projectCode;
    if (!item.projectName && row.projectName) item.projectName = row.projectName;
    if (!item.machineQuantity && row.machineQuantity) item.machineQuantity = row.machineQuantity;
    if (!item.department && row.department) item.department = row.department;
    if (row.sourceNote && !item.sourceNotes.includes(row.sourceNote)) {
      item.sourceNotes.push(row.sourceNote);
    }
  });
  return Array.from(map.values());
}

function findFuzzyMatches(source, candidates, minSimilarity, limit) {
  if (candidates.length === 0) return [];

  return candidates
    .map((item) => ({ item, similarity: itemSimilarity(source, item) }))
    .filter((match) => match.similarity >= minSimilarity)
    .sort((a, b) => b.similarity - a.similarity)
    .slice(0, limit);
}

function buildFuzzyCandidateIndex(items) {
  const index = new Map();
  items.forEach((item) => {
    const comparable = normalizeComparableCode(item.drawingKey);
    const prefix = comparable.slice(0, 3);
    if (!prefix) return;
    if (!index.has(prefix)) index.set(prefix, []);
    index.get(prefix).push(item);
  });
  return index;
}

function getFuzzyCandidates(source, index, fallbackItems) {
  const comparable = normalizeComparableCode(source.drawingKey);
  const prefix = comparable.slice(0, 3);
  if (!prefix) return fallbackItems.length <= 2000 ? fallbackItems : [];

  const candidates = index.get(prefix) || [];
  return candidates.filter((item) => {
    const candidateCode = normalizeComparableCode(item.drawingKey);
    const lengthGap = Math.abs(comparable.length - candidateCode.length);
    return lengthGap <= Math.max(4, Math.ceil(Math.max(comparable.length, candidateCode.length) * 0.35));
  });
}

function findOrderNameMatches(design, orderItems, minSimilarity, limit) {
  const designCode = normalizeLookupText(normalizeComparableCode(design.drawingKey));
  if (!designCode) return [];

  const prefix = designCode.slice(0, 3);
  const candidates = orderItems.filter((item) => {
    const itemName = normalizeLookupText(item.itemName);
    return itemName && (!prefix || itemName.includes(prefix));
  });

  return candidates
    .map((item) => ({ item, similarity: itemNameSimilarity(designCode, item.itemName) }))
    .filter((match) => match.similarity >= minSimilarity)
    .sort((a, b) => b.similarity - a.similarity)
    .slice(0, limit);
}

function itemNameSimilarity(designCode, itemName) {
  const nameText = normalizeLookupText(itemName);
  if (!designCode || !nameText) return 0;
  if (nameText.includes(designCode) || designCode.includes(nameText)) return 1;

  const nameParts = nameText.match(/[A-Z0-9]+/gi) || [];
  const partScore = nameParts.reduce((best, part) => Math.max(best, stringSimilarity(designCode, part)), 0);
  return Math.max(partScore, stringSimilarity(designCode, nameText));
}

function itemSimilarity(source, candidate) {
  const codeScore = Math.max(
    stringSimilarity(source.drawingKey, candidate.drawingKey),
    stringSimilarity(normalizeComparableCode(source.drawingKey), normalizeComparableCode(candidate.drawingKey))
  );
  const sourceUnit = normalizeText(source.unit);
  const candidateUnit = normalizeText(candidate.unit);
  const unitScore = sourceUnit && candidateUnit ? stringSimilarity(sourceUnit, candidateUnit) : 0;
  return sourceUnit && candidateUnit ? (codeScore * 0.8) + (unitScore * 0.2) : codeScore;
}

function unitsMatch(designUnit, orderUnit) {
  const left = normalizeUnit(designUnit);
  const right = normalizeUnit(orderUnit);
  if (!left && !right) return true;
  return left === right;
}

function normalizeUnit(value) {
  return normalizeText(value).replace(/[^a-z0-9]+/g, '').toUpperCase();
}

function createCompareRow(design, order, similarity, corrected) {
  const codeCorrected = corrected && design.drawingKey !== order.drawingKey;
  const unitCorrected = corrected && !unitsMatch(design.unit, order.unit);
  const notes = [];
  if (codeCorrected) {
    notes.push(`Mã ${design.drawingCode} không chính xác mã đúng là ${order.drawingCode}`);
  }
  if (unitCorrected) {
    notes.push(`Đơn vị tính ${design.unit || '(trống)'} sai đơn vị tính đúng là ${order.unit || '(trống)'}`);
  }

  return {
    designKey: design.drawingKey,
    orderKey: order.drawingKey,
    designDrawingCode: design.drawingCode,
    orderDrawingCode: order.drawingCode,
    orderItemName: order.itemName,
    itemName: design.itemName || order.itemName,
    manufacturer: design.manufacturer,
    markingPart: design.markingPart,
    material: design.material,
    surface: design.surface,
    quantity: design.quantity,
    designUnit: design.unit,
    orderUnit: order.unit,
    unit: order.unit || design.unit,
    corrected,
    codeCorrected,
    unitCorrected,
    similarity: similarity ? `${Math.round(similarity * 100)}%` : '',
    note: notes.join('; ')
  };
}

function createNewCodeRow(design, suggestion) {
  const suggestedItem = suggestion?.item;
  return {
    designKey: design.drawingKey,
    designDrawingCode: design.drawingCode,
    suggestedOrderDrawingCode: suggestedItem?.drawingCode || '',
    suggestedOrderItemName: suggestedItem?.itemName || '',
    itemName: design.itemName || '',
    manufacturer: design.manufacturer || '',
    markingPart: design.markingPart || '',
    material: design.material || '',
    surface: design.surface || '',
    unit: design.unit || '',
    quantity: design.quantity || '',
    nameSimilarity: suggestion ? `${Math.round(suggestion.similarity * 100)}%` : '',
    note: (design.sourceNotes || []).join('; ')
  };
}

function createConfirmRow(design, candidates, reason = 'code') {
  const options = candidates.map((candidate) => ({
    orderKey: candidate.item.drawingKey,
    orderDrawingCode: candidate.item.drawingCode,
    itemName: candidate.item.itemName,
    orderUnit: candidate.item.unit,
    designUnit: design.unit,
    similarity: `${Math.round(candidate.similarity * 100)}%`
  }));
  const selected = options[0] || {};

  return {
    designKey: design.drawingKey,
    designDrawingCode: design.drawingCode,
    orderKey: selected.orderKey || '',
    orderDrawingCode: selected.orderDrawingCode || '',
    itemName: selected.itemName || design.itemName || '',
    designUnit: selected.designUnit || design.unit || '',
    orderUnit: selected.orderUnit || '',
    unit: selected.orderUnit || design.unit || '',
    similarity: selected.similarity || '',
    reason,
    selectedIndex: 0,
    options
  };
}

function acceptConfirm(index) {
  const row = state.confirmRows[index];
  if (!row) return;
  const selected = row.options[row.selectedIndex] || row.options[0];
  if (!selected) return;
  state.manualMatches.set(row.designKey, selected.orderKey);
  runCompare();
}

function rejectConfirm(index) {
  const row = state.confirmRows[index];
  if (!row) return;
  row.options.forEach((option) => {
    state.rejectedMatches.add(`${row.designKey}::${option.orderKey}`);
  });
  runCompare();
}

async function exportCurrentTableLegacy() {
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
        showToast('Chưa có dữ liệu Mã Mới để xuất.');
        return;
      }

      const filePath = await window.inventoryApi.exportDiscrepancy({
        discrepancyRows: state.discrepancyRows
      });
      if (filePath) showToast(`Đã xuất bảng Mã Mới: ${filePath}`);
      return;
    }

    showToast('Hãy chọn bảng So Sánh hoặc Mã Mới trước khi xuất Excel.');
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
  els.loadKhoBtn.textContent = state.khoRows.length ? 'Load Lại Mã Đã Đặt Hàng' : 'Mã Đã Đặt Hàng';
  els.loadBomBtn.textContent = state.bomRows.length ? 'Thêm Dữ Liệu Thiết Kế' : 'Dữ Liệu Thiết Kế';
  els.clearBomBtn.textContent = state.showBomClearPanel ? 'Ẩn Clear Thiết Kế' : 'Clear Thiết Kế';
  els.clearKhoBtn.disabled = !state.khoRows.length;
  els.clearBomBtn.disabled = !state.bomRows.length;
  els.exportExcelBtn.disabled = !canExportCurrentTable();
  if (!state.bomRows.length) state.showBomClearPanel = false;
  renderBomClearPanel();

  const okCount = state.compareRows.length;
  const missingCount = state.discrepancyRows.length;
  const extraCount = state.confirmRows.length;

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
  if (state.view === 'kho') {
    renderKhoTable();
    return;
  }

  if (state.view === 'bom') {
    renderTable(els.bomTable, filterRows(state.bomRows), columns.bom);
    return;
  }

  if (state.view === 'compare') {
    renderConfirmTable();
    renderTable(els.compareTable, filterRows(state.compareRows), columns.compare, statusClass);
    return;
  }

  if (state.view === 'discrepancy') {
    renderTable(els.discrepancyTable, filterRows(state.discrepancyRows), columns.discrepancy, statusClass);
    return;
  }
}

function renderKhoTable() {
  const rows = filterRows(state.khoRows);
  const totalPages = Math.max(1, Math.ceil(rows.length / MAX_VISIBLE_ROWS));
  state.khoPage = clamp(Math.floor(Number(state.khoPage)) || 1, 1, totalPages);
  const startIndex = (state.khoPage - 1) * MAX_VISIBLE_ROWS;
  const pageRows = rows.slice(startIndex, startIndex + MAX_VISIBLE_ROWS);
  renderTable(els.khoTable, pageRows, columns.kho, null, {
    pagination: createPaginationHtml(rows.length, state.khoPage)
  });
}

function canExportCurrentTable() {
  return state.compareRows.length > 0 || state.discrepancyRows.length > 0 || state.confirmRows.length > 0;
}

function renderTable(container, rows, tableColumns, rowClassFn, options = {}) {
  if (!rows.length) {
    container.innerHTML = '<div class="placeholder">Chưa có dữ liệu.</div>';
    return;
  }

  const visibleRows = rows.slice(0, options.pagination ? rows.length : MAX_VISIBLE_ROWS);
  const limitNotice = !options.pagination && rows.length > MAX_VISIBLE_ROWS
    ? `<div class="table-limit">Đang hiển thị ${MAX_VISIBLE_ROWS}/${rows.length} dòng. Dữ liệu xuất Excel vẫn đầy đủ.</div>`
    : '';
  const thead = tableColumns
    .map(([key, label, align]) => `<th class="${[align || '', `col-${key}`].filter(Boolean).join(' ')}">${escapeHtml(label)}</th>`)
    .join('');
  const tbody = visibleRows.map((row) => {
    const cells = tableColumns.map(([key, , align]) => {
      const tdClass = [align || '', `col-${key}`].filter(Boolean).join(' ');
      if (key === 'designDrawingCode' && row.corrected) {
        return `<td class="${tdClass}"><s>${escapeHtml(row[key])}</s></td>`;
      }
      if (key === 'designUnit' && row.unitCorrected) {
        return `<td class="${tdClass}"><s>${escapeHtml(row[key])}</s></td>`;
      }
      if (shouldClampCell(container, key)) {
        return `<td class="${tdClass}"><div class="cell-clamp">${escapeHtml(row[key])}</div></td>`;
      }
      return `<td class="${tdClass}">${escapeHtml(row[key])}</td>`;
    }).join('');
    return `<tr class="${rowClassFn ? rowClassFn(row) : ''}">${cells}</tr>`;
  }).join('');

  const pagination = options.pagination || '';
  container.innerHTML = `${pagination}${limitNotice}<table><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table>${pagination}`;
}

function shouldClampCell(container, key) {
  return (
    (container === els.khoTable && key === 'itemName') ||
    (container === els.compareTable && key === 'orderItemName') ||
    (container === els.confirmTable && key === 'itemName')
  );
}

function createPaginationHtml(totalRows, currentPage) {
  if (totalRows <= MAX_VISIBLE_ROWS) return '';

  const totalPages = Math.ceil(totalRows / MAX_VISIBLE_ROWS);
  const pages = getVisiblePages(currentPage, totalPages);
  const buttons = pages.map((page, index) => {
    if (page === 'gap') return '<span class="page-gap">...</span>';
    const active = page === currentPage ? ' active' : '';
    const label = String(page);
    return `<button class="page-box${active}" data-page="${page}" type="button">${escapeHtml(label)}</button>`;
  }).join('');

  return `<div class="page-boxes">${buttons}</div>`;
}

function getVisiblePages(currentPage, totalPages) {
  if (totalPages <= 5) {
    return Array.from({ length: totalPages }, (_, index) => index + 1);
  }

  if (currentPage <= 3) {
    return [1, 2, 3, 4, 'gap', totalPages];
  }

  if (currentPage >= totalPages - 2) {
    return [1, 'gap', totalPages - 3, totalPages - 2, totalPages - 1, totalPages];
  }

  return [1, 'gap', currentPage - 1, currentPage, currentPage + 1, 'gap', totalPages];
}

function handleKhoTableClick(event) {
  const button = event.target.closest('button[data-page]');
  if (!button) return;
  state.khoPage = Number(button.dataset.page) || 1;
  renderKhoTable();
  els.khoTable.scrollTop = 0;
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

  const thead = columns.confirm
    .map(([key, label, align]) => `<th class="${[align || '', `col-${key}`].filter(Boolean).join(' ')}">${escapeHtml(label)}</th>`)
    .join('');
  const tbody = rows.map((row) => {
    const actualIndex = state.confirmRows.findIndex((item) => item.designKey === row.designKey);
    const cells = columns.confirm.map(([key, , align]) => {
      if (key === 'actions') {
        return `<td><div class="row-actions"><button class="mini-button confirm" data-action="accept" data-index="${actualIndex}">Chọn</button><button class="mini-button reject" data-action="reject" data-index="${actualIndex}">Bỏ qua</button></div></td>`;
      }
      if (key === 'orderSelect') {
        if (row.reason === 'unit') {
          return `<td class="${[align || '', `col-${key}`].filter(Boolean).join(' ')}">${escapeHtml(row.orderDrawingCode)}</td>`;
        }

        const options = row.options.map((option, optionIndex) => {
          const selected = optionIndex === row.selectedIndex ? ' selected' : '';
          return `<option value="${optionIndex}"${selected}>${escapeHtml(option.orderDrawingCode)} - ${escapeHtml(option.similarity)}</option>`;
        }).join('');
        return `<td class="col-${key}"><select class="bom-select" data-index="${actualIndex}">${options}</select></td>`;
      }
      if (key === 'orderUnit' && row.reason === 'unit') {
        return `<td class="${[align || '', `col-${key}`].filter(Boolean).join(' ')}">${escapeHtml(row.orderUnit)}</td>`;
      }
      if (key === 'itemName') {
        return `<td class="${[align || '', `col-${key}`].filter(Boolean).join(' ')}"><div class="cell-clamp">${escapeHtml(row[key])}</div></td>`;
      }
      return `<td class="${[align || '', `col-${key}`].filter(Boolean).join(' ')}">${escapeHtml(row[key])}</td>`;
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
  row.orderKey = selected.orderKey;
  row.orderDrawingCode = selected.orderDrawingCode;
  row.itemName = selected.itemName;
  row.designUnit = selected.designUnit || row.designUnit;
  row.orderUnit = selected.orderUnit || '';
  row.unit = selected.orderUnit || row.designUnit || '';
  row.similarity = selected.similarity;
}

async function exportCurrentTable() {
  try {
    if (!state.compareRows.length && !state.discrepancyRows.length && !state.confirmRows.length) {
      showToast('Chua co du lieu de xuat Excel.');
      return;
    }

    if (state.confirmRows.length) {
      const shouldContinue = await window.inventoryApi.confirmPendingExport();
      if (!shouldContinue) {
        showToast('Da huy xuat Excel.');
        return;
      }
    }

    const filePath = await window.inventoryApi.exportExcel(buildPurchaseExportPayload());
    if (filePath) showToast(`Da xuat De Nghi Mua Hang: ${filePath}`);
  } catch (error) {
    showToast(`Khong the xuat Excel: ${error.message}`);
  }
}

function buildPurchaseExportPayload() {
  const metadata = getExportMetadata();
  const byDesignKey = new Map();
  state.compareRows.forEach((row) => {
    byDesignKey.set(row.designKey, createPurchaseRowFromCompare(row, metadata));
  });
  state.discrepancyRows.forEach((row) => {
    if (!byDesignKey.has(row.designKey)) {
      byDesignKey.set(row.designKey, createPurchaseRowFromDiscrepancy(row, metadata));
    }
  });

  return {
    metadata,
    purchaseRows: Array.from(byDesignKey.values())
  };
}

function getExportMetadata() {
  const source = state.bomRows.find((row) => row.projectCode || row.projectName || row.machineQuantity || row.department) || {};
  return {
    documentCode: source.documentCode || '',
    projectCode: source.projectCode || '',
    projectName: source.projectName || '',
    machineQuantity: parseQuantity(source.machineQuantity) || 1,
    department: source.department || 'Bo phan Co'
  };
}

function createPurchaseRowFromCompare(row, metadata) {
  return {
    itemName: row.itemName || row.orderItemName || '',
    drawingCode: row.designDrawingCode || '',
    manufacturer: row.manufacturer || '',
    markingPart: row.markingPart || '',
    material: row.material || '',
    surface: row.surface || '',
    unit: row.designUnit || row.unit || '',
    quantity: row.quantity || '',
    machineQuantity: metadata.machineQuantity || 1,
    explain: row.note || 'Ma giong nhau'
  };
}

function createPurchaseRowFromDiscrepancy(row, metadata) {
  return {
    itemName: row.itemName || '',
    drawingCode: row.designDrawingCode || '',
    manufacturer: row.manufacturer || '',
    markingPart: row.markingPart || '',
    material: row.material || '',
    surface: row.surface || '',
    unit: row.unit || '',
    quantity: row.quantity || '',
    machineQuantity: metadata.machineQuantity || 1,
    explain: row.note || 'Ma moi'
  };
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

function normalizeComparableCode(value) {
  return normalizeCode(value).replace(/\.(PDF|DWG|DXF|STEP|STP|IGS|IGES|SLDPRT|SLDASM|XLS|XLSX|CSV)$/i, '');
}

function normalizeLookupText(value) {
  return normalizeText(value).toUpperCase().replace(/[^A-Z0-9]+/g, '');
}

function normalizeText(value) {
  return cleanCell(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
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
