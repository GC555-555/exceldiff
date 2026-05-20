let file1 = null;
let file2 = null;
let workbook1 = null;
let workbook2 = null;
let rawData1 = null;
let rawData2 = null;
let comparisonResult = null;
let currentFilter = 'diff';
let currentPreviewSheet = '';
let sheetNames1 = [];
let sheetNames2 = [];
let startCellPerSheetFile1 = {};
let startCellPerSheetFile2 = {};
let resultsData = [];
let resultsSheetNames = [];
let currentResultsSheet = 0;
let resultsStyles1 = {};
let resultsStyles2 = {};
let resultsMerges1 = {};
let resultsMerges2 = {};
let currentChapterColMap = {};
let currentStartRowFile1 = {};
let currentStartRowFile2 = {};
let previewStyles1 = {};
let previewStyles2 = {};
let previewData1 = null;
let previewData2 = null;
let previewData = [];
let previewSheetNames = [];
let currentPreviewSheetIndex = 0;
let isFullscreen = false;
let isComparing = false;
let isResultsMaximized = false;

function getSafeElement(id) { return document.getElementById(id); }

function getUniqueSheetNames(sheetNamesA, sheetNamesB) {
  const seen = new Set();
  const result = [];
  if (Array.isArray(sheetNamesA)) { sheetNamesA.forEach(n => { if (!seen.has(n)) { seen.add(n); result.push(n); } }); }
  if (Array.isArray(sheetNamesB)) { sheetNamesB.forEach(n => { if (!seen.has(n)) { seen.add(n); result.push(n); } }); }
  return result;
}

function formatFileSize(bytes) {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function setupUploadZone(zoneId, inputId, fileNum) {
  const zone = document.getElementById(zoneId);
  const input = document.getElementById(inputId);
  if (!zone || !input) return;
  zone.addEventListener('click', () => { input.click(); });
  zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragging'); });
  zone.addEventListener('dragleave', (e) => { e.preventDefault(); zone.classList.remove('dragging'); });
  zone.addEventListener('drop', (e) => { e.preventDefault(); zone.classList.remove('dragging'); if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0], fileNum); });
  input.addEventListener('change', (e) => { if (e.target.files.length > 0) handleFile(e.target.files[0], fileNum); });
}

async function handleFile(file, fileNum) {
  if (!file || !file.name) { alert(i18n.t('alert_parse_failed')); return; }
  if (!file.name.match(/\.(xlsx|xls)$/i)) { alert(i18n.t('alert_invalid_file')); return; }
  try {
    const rawData = await readFileAsArrayBuffer(file);
    const workbook = XLSX.read(rawData, {
      type: 'array',
      cellStyles: true,
      cellFormula: false,
      cellNF: true,
      cellDates: false,
      bookSST: true,
      sheetStubs: true
    });

    // 提取样式
    if (typeof extractAllStyles === 'function') {
      if (fileNum === 1) {
        window._parsedStylesFile1 = await extractAllStyles(rawData, workbook);
      } else {
        window._parsedStylesFile2 = await extractAllStyles(rawData, workbook);
      }
    }

    const sheetData = getSheetData(workbook);

    if (fileNum === 1) {
      file1 = file; workbook1 = workbook; rawData1 = rawData; sheetNames1 = workbook.SheetNames || []; previewData1 = sheetData.sheets;
      const nameEl = document.getElementById('file1Name');
      if (nameEl) nameEl.textContent = file.name + ' (' + formatFileSize(file.size) + ')';
      const infoEl = document.getElementById('file1Info');
      if (infoEl) infoEl.style.display = 'flex';
    } else {
      file2 = file; workbook2 = workbook; rawData2 = rawData; sheetNames2 = workbook.SheetNames || []; previewData2 = sheetData.sheets;
      const nameEl = document.getElementById('file2Name');
      if (nameEl) nameEl.textContent = file.name + ' (' + formatFileSize(file.size) + ')';
      const infoEl = document.getElementById('file2Info');
      if (infoEl) infoEl.style.display = 'flex';
    }
    updateCompareButton(); renderSheetTabs(); updateFilePreview();
  } catch (error) { alert(i18n.t('alert_parse_failed') + ': ' + error.message); console.error(error); }
}

async function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(new Uint8Array(e.target.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function removeFile(fileNum) {
  if (fileNum === 1) {
    file1 = null; workbook1 = null; previewData1 = null; startCellPerSheetFile1 = {};
    const infoEl = document.getElementById('file1Info');
    if (infoEl) infoEl.style.display = 'none';
    const inputEl = document.getElementById('file1Input');
    if (inputEl) inputEl.value = '';
    window._parsedStylesFile1 = null;
  } else {
    file2 = null; workbook2 = null; previewData2 = null; startCellPerSheetFile2 = {};
    const infoEl = document.getElementById('file2Info');
    if (infoEl) infoEl.style.display = 'none';
    const inputEl = document.getElementById('file2Input');
    if (inputEl) inputEl.value = '';
    window._parsedStylesFile2 = null;
  }
  updateCompareButton(); updateFilePreview();
}

function updateCompareButton() {
  const directBtn = document.getElementById('directCompareBtn');
  const selectCellBtn = document.getElementById('selectCellBtn');
  const bothFilesReady = file1 && file2;
  directBtn.disabled = !bothFilesReady; selectCellBtn.disabled = !bothFilesReady;
}

function showCellSelector() {
  document.getElementById('sheetPreviewContainer').style.display = 'block';
  document.getElementById('confirmCompareContainer').style.display = 'flex';
  renderSheetTabs(); updateFilePreview();
}

function renderSheetTabs() {
  const tabBar = document.getElementById('sheetTabBar');
  if (!tabBar) return;
  tabBar.innerHTML = '';
  if (sheetNames1.length === 0 && sheetNames2.length === 0) { tabBar.style.display = 'none'; return; }
  tabBar.style.display = 'flex';
  const allSheets = new Set([...sheetNames1, ...sheetNames2]);
  const sheetList = Array.from(allSheets);
  if (currentPreviewSheet === '' && sheetList.length > 0) currentPreviewSheet = sheetList[0];
  if (sheetList.indexOf(currentPreviewSheet) === -1 && sheetList.length > 0) currentPreviewSheet = sheetList[0];
  sheetList.forEach(sheetName => {
    const tabBtn = document.createElement('button');
    tabBtn.className = 'sheet-tab-item' + (sheetName === currentPreviewSheet ? ' active' : '');
    let badges = '';
    if (sheetNames1.indexOf(sheetName) === -1) badges += '📄 ';
    if (sheetNames2.indexOf(sheetName) === -1) badges += '📄 ';
    if (sheetNames1.indexOf(sheetName) !== -1 && sheetNames2.indexOf(sheetName) !== -1) badges += '🔄 ';
    tabBtn.textContent = badges + sheetName;
    tabBtn.addEventListener('click', () => {
      document.querySelectorAll('.sheet-tab-item').forEach(btn => btn.classList.remove('active'));
      tabBtn.classList.add('active'); currentPreviewSheet = sheetName; updateFilePreview();
    });
    tabBar.appendChild(tabBtn);
  });
}

function updateFilePreview() {
  const container = document.getElementById('sheetPreviewContainer');
  if (container.style.display === 'none') return;
  const sheet1Data = previewData1 && previewData1[currentPreviewSheet] ? previewData1[currentPreviewSheet] : [];
  const sheet2Data = previewData2 && previewData2[currentPreviewSheet] ? previewData2[currentPreviewSheet] : [];
  const hasData1 = sheet1Data.length > 0; const hasData2 = sheet2Data.length > 0;
  document.getElementById('file1PreviewBlock').style.display = hasData1 ? 'flex' : 'none';
  document.getElementById('file2PreviewBlock').style.display = hasData2 ? 'flex' : 'none';
  if (file1) document.getElementById('file1PreviewTitle').textContent = file1.name + ' - ' + (currentPreviewSheet || '');
  if (file2) document.getElementById('file2PreviewTitle').textContent = file2.name + ' - ' + (currentPreviewSheet || '');
  updateStartCellDisplay();
  if (hasData1) renderPreviewTable('file1PreviewTable', sheet1Data, 1);
  if (hasData2) renderPreviewTable('file2PreviewTable', sheet2Data, 2);
}

function updateStartCellDisplay() {
  const cell1 = startCellPerSheetFile1[currentPreviewSheet] || 'A1';
  const cell2 = startCellPerSheetFile2[currentPreviewSheet] || 'A1';
  document.getElementById('startCellValue1').textContent = cell1;
  document.getElementById('startCellValue2').textContent = cell2;
}

function setStartCell(fileNum, cellRef) {
  if (fileNum === 1) startCellPerSheetFile1[currentPreviewSheet] = cellRef;
  else startCellPerSheetFile2[currentPreviewSheet] = cellRef;
  updateStartCellDisplay(); updateFilePreview();
}

function resetStartCell(fileNum) {
  if (fileNum === 1) delete startCellPerSheetFile1[currentPreviewSheet];
  else delete startCellPerSheetFile2[currentPreviewSheet];
  updateStartCellDisplay(); updateFilePreview();
}

function renderPreviewTable(tableId, data, fileNum) {
  const table = document.getElementById(tableId);
  if (!data || data.length === 0) { table.innerHTML = '<tr><td>' + i18n.t('no_data') + '</td></tr>'; return; }
  const maxRows = 5; const displayData = data.slice(0, maxRows);
  const startCellKey = fileNum === 1 ? startCellPerSheetFile1[currentPreviewSheet] : startCellPerSheetFile2[currentPreviewSheet];
  const startCell = startCellKey ? parseCellRef(startCellKey) : null;
  const highlightedCol = startCell ? startCell.col : -1;
  let html = '<thead><tr><th class="preview-row-header"></th>';
  const maxCols = Math.min(displayData[0] ? displayData[0].length : 0, 10);
  for (let c = 0; c < maxCols; c++) html += `<th>${getColumnName(c)}</th>`;
  if (displayData[0] && displayData[0].length > 10) html += '<th>...</th>';
  html += '</tr></thead><tbody>';
  displayData.forEach((row, rowIndex) => {
    html += '<tr>'; html += `<td class="preview-row-number">${rowIndex + 1}</td>`;
    for (let col = 0; col < maxCols; col++) {
      const cellValue = row[col] !== undefined ? row[col] : '';
      const displayValue = String(cellValue).substring(0, 50);
      let cellClass = '';
      if (col === highlightedCol) { if (rowIndex + 1 === (startCell ? startCell.row + 1 : 0)) cellClass = 'start-cell-selected'; else cellClass = 'col-highlighted'; }
      html += `<td class="${cellClass}" data-file="${fileNum}" data-row="${rowIndex}" data-col="${col}" title="${cellValue}">${displayValue}</td>`;
    }
    if (row.length > 10) html += '<td>...</td>';
    html += '</tr>';
  });
  if (data.length > maxRows) html += `<tr><td class="preview-row-number"></td><td colspan="${maxCols + 1}">${i18n.t('more_rows', { count: data.length - maxRows })}</td></tr>`;
  html += '</tbody>'; table.innerHTML = html;
  table.querySelectorAll('tbody td[data-file]').forEach(td => {
    td.addEventListener('click', () => {
      const fn = parseInt(td.dataset.file); const row = parseInt(td.dataset.row); const col = parseInt(td.dataset.col);
      setStartCell(fn, getColumnName(col) + (row + 1));
    });
  });
}

async function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {
          type: 'array',
          cellStyles: true,
          cellFormula: false,
          cellNF: true,
          cellDates: false,
          bookSST: true,
          sheetStubs: true
        });
        resolve(workbook);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function getSheetData(workbook, parsedStyles) {
  const sheetNames = workbook.SheetNames;
  const sheets = {}; const sheetStyles = {}; const merges = {}; const colWidths = {}; const rowHeights = {}; const originalCells = {};
  sheetNames.forEach(name => {
    const sheet = workbook.Sheets[name];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (sheet['!merges']) merges[name] = JSON.parse(JSON.stringify(sheet['!merges'])); else merges[name] = [];
    if (sheet['!cols']) colWidths[name] = JSON.parse(JSON.stringify(sheet['!cols'])); else colWidths[name] = [];
    if (sheet['!rows']) rowHeights[name] = JSON.parse(JSON.stringify(sheet['!rows'])); else rowHeights[name] = [];
    const styles = {}; const cells = {};
    if (sheet) {
      const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
      for (let R = range.s.r; R <= range.e.r; ++R) {
        styles[R] = {}; cells[R] = {};
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cellRef = XLSX.utils.encode_cell({ c: C, r: R });
          const cell = sheet[cellRef];
          if (cell) {
            cells[R][C] = JSON.parse(JSON.stringify(cell));
            if (cell.s) styles[R][C] = JSON.parse(JSON.stringify(cell.s));
          }
        }
      }
    }
    if (parsedStyles && parsedStyles[name]) {
      Object.assign(styles, parsedStyles[name]);
    }
    sheetStyles[name] = styles; originalCells[name] = cells; sheets[name] = data;
  });
  return { sheets, styles: sheetStyles, merges, colWidths, rowHeights, originalCells };
}

function getSheetDataWithParsedStyles(workbook) {
  if (workbook === workbook1 && window._parsedStylesFile1) {
    return getSheetData(workbook, window._parsedStylesFile1);
  }
  if (workbook === workbook2 && window._parsedStylesFile2) {
    return getSheetData(workbook, window._parsedStylesFile2);
  }
  return getSheetData(workbook);
}

function compareSheets(sheets1, sheets2, chapterColumnMap, startRowFile1, startRowFile2) {
  const localSheetNames1 = Object.keys(sheets1 || {}); const localSheetNames2 = Object.keys(sheets2 || {});
  const allSheetNames = getUniqueSheetNames(localSheetNames1, localSheetNames2);
  const differences = []; let stats = { added: 0, deleted: 0, modified: 0, identical: 0 };
  allSheetNames.forEach(sheetName => {
    const raw1 = sheets1[sheetName] || []; const raw2 = sheets2[sheetName] || [];
    const startRow1 = startRowFile1 && startRowFile1[sheetName] !== undefined ? startRowFile1[sheetName] : 0;
    const startRow2 = startRowFile2 && startRowFile2[sheetName] !== undefined ? startRowFile2[sheetName] : 0;
    const data1 = raw1.slice(startRow1); const data2 = raw2.slice(startRow2);
    const sheetChapterCol = chapterColumnMap && chapterColumnMap[sheetName] !== undefined ? chapterColumnMap[sheetName] : 0;
    const groups1 = groupDataByLevel1Chapters(data1, sheetChapterCol); const groups2 = groupDataByLevel1Chapters(data2, sheetChapterCol);
    for (let g = 0; g < Math.max(groups1.length, groups2.length); g++) {
      const group1 = groups1[g] || { startIdx: Infinity, rows: [] }; const group2 = groups2[g] || { startIdx: Infinity, rows: [] };
      const chapterRows1 = new Map(); group1.rows.forEach(item => { const chapter = getChapterFromRow(item.row, sheetChapterCol); if (chapter) { if (!chapterRows1.has(chapter)) chapterRows1.set(chapter, []); chapterRows1.get(chapter).push({ row: item.row, originalIdx: item.idx }); } });
      const chapterRows2 = new Map(); group2.rows.forEach(item => { const chapter = getChapterFromRow(item.row, sheetChapterCol); if (chapter) { if (!chapterRows2.has(chapter)) chapterRows2.set(chapter, []); chapterRows2.get(chapter).push({ row: item.row, originalIdx: item.idx }); } });
      const noChapter1 = group1.rows.filter(item => !getChapterFromRow(item.row, sheetChapterCol));
      const noChapter2 = group2.rows.filter(item => !getChapterFromRow(item.row, sheetChapterCol));
      for (const item2 of group2.rows) {
        const row2 = item2.row; const idx2 = item2.idx; if (!Array.isArray(row2)) continue;
        const chapter2 = getChapterFromRow(row2, sheetChapterCol);
        if (chapter2) {
          const oldRows = chapterRows1.get(chapter2);
          if (!oldRows) { for (let col = 0; col < row2.length; col++) { const val = (row2[col] ?? '').toString().trim(); if (val !== '') { differences.push({ type: 'added', position: `${sheetName}!${getColumnName(col)}${idx2 + 1}`, oldValue: '', newValue: val, row: idx2, col: col }); stats.added++; } else stats.identical++; } }
          else { let pos = -1; for (let j = 0; j < chapterRows2.get(chapter2).length; j++) { if (chapterRows2.get(chapter2)[j].originalIdx === idx2) { pos = j; break; } } if (pos < 0) continue; const match = oldRows[pos]; if (!match) { for (let col = 0; col < row2.length; col++) { const val = (row2[col] ?? '').toString().trim(); if (val !== '') { differences.push({ type: 'added', position: `${sheetName}!${getColumnName(col)}${idx2 + 1}`, oldValue: '', newValue: val, row: idx2, col: col }); stats.added++; } else stats.identical++; } } else if (rowsEqual(match.row, row2)) { const maxLen = Math.max(match.row.length, row2.length); for (let col = 0; col < maxLen; col++) stats.identical++; } else { const maxLen = Math.max(match.row.length, row2.length); for (let col = 0; col < maxLen; col++) { const v1 = (match.row[col] ?? '').toString().trim(); const v2 = (row2[col] ?? '').toString().trim(); if (v1 !== v2) { if (v1 === '' && v2 !== '') { differences.push({ type: 'added', position: `${sheetName}!${getColumnName(col)}${idx2 + 1}`, oldValue: '', newValue: v2, row: idx2, col: col }); stats.added++; } else if (v1 !== '' && v2 === '') { differences.push({ type: 'deleted', position: `${sheetName}!${getColumnName(col)}${idx2 + 1}`, oldValue: v1, newValue: '', row: idx2, col: col }); stats.deleted++; } else { differences.push({ type: 'modified', position: `${sheetName}!${getColumnName(col)}${idx2 + 1}`, oldValue: v1, newValue: v2, row: idx2, col: col }); stats.modified++; } } else stats.identical++; } } }
        } else {
          const ncIdx = noChapter2.findIndex(e => e.idx === idx2); const match = ncIdx >= 0 && ncIdx < noChapter1.length ? noChapter1[ncIdx] : null;
          if (!match) { for (let col = 0; col < row2.length; col++) { const val = (row2[col] ?? '').toString().trim(); if (val !== '') { differences.push({ type: 'added', position: `${sheetName}!${getColumnName(col)}${idx2 + 1}`, oldValue: '', newValue: val, row: idx2, col: col }); stats.added++; } else stats.identical++; } }
          else if (rowsEqual(match.row, row2)) { const maxLen = Math.max(match.row.length, row2.length); for (let col = 0; col < maxLen; col++) stats.identical++; }
          else { const maxLen = Math.max(match.row.length, row2.length); for (let col = 0; col < maxLen; col++) { const v1 = (match.row[col] ?? '').toString().trim(); const v2 = (row2[col] ?? '').toString().trim(); if (v1 !== v2) { if (v1 === '' && v2 !== '') { differences.push({ type: 'added', position: `${sheetName}!${getColumnName(col)}${idx2 + 1}`, oldValue: '', newValue: v2, row: idx2, col: col }); stats.added++; } else if (v1 !== '' && v2 === '') { differences.push({ type: 'deleted', position: `${sheetName}!${getColumnName(col)}${idx2 + 1}`, oldValue: v1, newValue: '', row: idx2, col: col }); stats.deleted++; } else { differences.push({ type: 'modified', position: `${sheetName}!${getColumnName(col)}${idx2 + 1}`, oldValue: v1, newValue: v2, row: idx2, col: col }); stats.modified++; } } else stats.identical++; } }
        }
      }
      chapterRows1.forEach((oldRows, chapter) => { const newRows = chapterRows2.get(chapter); if (!newRows || oldRows.length > newRows.length) { const deletedCount = oldRows.length - (newRows ? newRows.length : 0); for (let k = 0; k < deletedCount; k++) { const delRow = oldRows[oldRows.length - 1 - k].row; const delIdx = oldRows[oldRows.length - 1 - k].originalIdx; for (let col = 0; col < delRow.length; col++) { const val = (delRow[col] ?? '').toString().trim(); if (val !== '') { differences.push({ type: 'deleted', position: `${sheetName}!${getColumnName(col)}${delIdx + 1}`, oldValue: val, newValue: '', row: delIdx, col: col }); stats.deleted++; } else stats.identical++; } } } });
    }
  });
  return { differences, stats };
}

function getColumnName(col) { let result = ''; while (col >= 0) { result = String.fromCharCode((col % 26) + 65) + result; col = Math.floor(col / 26) - 1; } return result; }
function parseCellRef(cellRef) { const match = cellRef.trim().toUpperCase().match(/^([A-Z]+)(\d+)$/); if (!match) return { col: 0, row: 0 }; const colStr = match[1]; const rowStr = match[2]; let col = 0; for (let i = 0; i < colStr.length; i++) col = col * 26 + (colStr.charCodeAt(i) - 64); return { col: col - 1, row: parseInt(rowStr) - 1 }; }
function extractChapterNumber(text) { if (!text) return null; const match = String(text).trim().match(/^(\d+(?:\.\d+)*)/); return match ? match[1] : null; }
function getChapterFromRow(row, columnIndex) { if (!Array.isArray(row) || row.length <= columnIndex) return null; return extractChapterNumber(row[columnIndex]); }
function parseChapterNumber(chapterStr) { if (!chapterStr) return null; return chapterStr.split('.').map(Number); }
function compareChapterNumbers(ch1, ch2) { if (!ch1 && !ch2) return 0; if (!ch1) return 1; if (!ch2) return -1; const p1 = parseChapterNumber(ch1); const p2 = parseChapterNumber(ch2); const maxLen = Math.max(p1.length, p2.length); for (let i = 0; i < maxLen; i++) { const a = i < p1.length ? p1[i] : 0; const b = i < p2.length ? p2[i] : 0; if (a !== b) return a - b; } return 0; }
function rowsEqual(row1, row2) { if (!row1 && !row2) return true; if (!row1 || !row2) return false; const len1 = Array.isArray(row1) ? row1.length : 0; const len2 = Array.isArray(row2) ? row2.length : 0; const maxLen = Math.max(len1, len2); for (let i = 0; i < maxLen; i++) { const v1 = row1[i] ?? ''; const v2 = row2[i] ?? ''; if (String(v1).trim() !== String(v2).trim()) return false; } return true; }
function isLevel1Chapter(chapterStr) { if (!chapterStr) return false; return /^\d+$/.test(chapterStr.trim()); }
function groupDataByLevel1Chapters(data, chapterColumnIndex) { const groups = []; let currentGroup = null; for (let i = 0; i < data.length; i++) { const row = data[i]; const chapter = getChapterFromRow(row, chapterColumnIndex); if (chapter && isLevel1Chapter(chapter)) { if (currentGroup && currentGroup.rows.length > 0) groups.push(currentGroup); currentGroup = { startIdx: i, rows: [{ row, idx: i }] }; } else { if (!currentGroup) currentGroup = { startIdx: i, rows: [] }; currentGroup.rows.push({ row, idx: i }); } } if (currentGroup && currentGroup.rows.length > 0) groups.push(currentGroup); return groups; }

async function compareFiles() {
  if (!file1 || !file2) return;
  if (isComparing) return;
  isComparing = true;
  const loading = getSafeElement('loading'); const results = getSafeElement('resultsSection');
  if (loading) loading.style.display = 'block';
  if (results) results.style.display = 'none';
  try {
    const result1 = getSheetDataWithParsedStyles(workbook1); const result2 = getSheetDataWithParsedStyles(workbook2);
    currentChapterColMap = {}; currentStartRowFile1 = {}; currentStartRowFile2 = {};
    const allSheetNames = getUniqueSheetNames(sheetNames1, sheetNames2);
    allSheetNames.forEach(sheetName => { const cell1 = startCellPerSheetFile1[sheetName] || 'A1'; const cell2 = startCellPerSheetFile2[sheetName] || 'A1'; const parsed1 = parseCellRef(cell1); const parsed2 = parseCellRef(cell2); currentChapterColMap[sheetName] = parsed1.col; currentStartRowFile1[sheetName] = parsed1.row; currentStartRowFile2[sheetName] = parsed2.row; });
    comparisonResult = compareSheets(result1.sheets, result2.sheets, currentChapterColMap, currentStartRowFile1, currentStartRowFile2);
    if (loading) loading.style.display = 'none'; if (results) results.style.display = 'block';
    const sheetPreviewContainer = getSafeElement('sheetPreviewContainer'); if (sheetPreviewContainer) sheetPreviewContainer.style.display = 'none';
    const confirmCompareContainer = getSafeElement('confirmCompareContainer'); if (confirmCompareContainer) confirmCompareContainer.style.display = 'none';
    renderResults();
  } catch (error) {
    if (loading) loading.style.display = 'none';
    alert(i18n.t('alert_parse_failed') + ': ' + error.message);
    console.error(error);
  } finally {
    isComparing = false;
  }
}

function prepareResultsData(chapterColumnMap, startRowFile1, startRowFile2) {
  if (!workbook1 || !workbook2) return;
  const result1 = getSheetDataWithParsedStyles(workbook1); const result2 = getSheetDataWithParsedStyles(workbook2);
  const sheets1 = result1.sheets; const sheets2 = result2.sheets;
  resultsStyles1 = result1.styles; resultsStyles2 = result2.styles;
  resultsMerges1 = result1.merges; resultsMerges2 = result2.merges;
  resultsData = []; resultsSheetNames = [];
  const localSheetNames1 = Object.keys(sheets1 || {}); const localSheetNames2 = Object.keys(sheets2 || {});
  const allSheetNames = getUniqueSheetNames(localSheetNames1, localSheetNames2);
  allSheetNames.forEach(sheetName => {
    const raw1 = sheets1[sheetName] || []; const raw2 = sheets2[sheetName] || [];
    const sheetStyles1 = resultsStyles1[sheetName] || {}; const sheetStyles2 = resultsStyles2[sheetName] || {};
    const startRow1 = startRowFile1 && startRowFile1[sheetName] !== undefined ? startRowFile1[sheetName] : 0;
    const startRow2 = startRowFile2 && startRowFile2[sheetName] !== undefined ? startRowFile2[sheetName] : 0;
    const data1 = raw1.slice(startRow1); const data2 = raw2.slice(startRow2);
    resultsSheetNames.push(sheetName);
    const sheetChapterCol = chapterColumnMap && chapterColumnMap[sheetName] !== undefined ? chapterColumnMap[sheetName] : 0;
    const groups1 = groupDataByLevel1Chapters(data1, sheetChapterCol); const groups2 = groupDataByLevel1Chapters(data2, sheetChapterCol);
    for (let g = 0; g < Math.max(groups1.length, groups2.length); g++) {
      const group1 = groups1[g] || { startIdx: Infinity, rows: [] }; const group2 = groups2[g] || { startIdx: Infinity, rows: [] };
      const groupStartIdx = resultsData.length;
      const chapters1 = new Map(); const noChapter1 = []; group1.rows.forEach(item => { const ch = getChapterFromRow(item.row, sheetChapterCol); if (ch) { if (!chapters1.has(ch)) chapters1.set(ch, []); chapters1.get(ch).push({ row: item.row, originalIdx: item.idx + startRow1 }); } else noChapter1.push({ row: item.row, originalIdx: item.idx + startRow1 }); });
      const chapters2 = new Map(); const noChapter2 = []; group2.rows.forEach(item => { const ch = getChapterFromRow(item.row, sheetChapterCol); if (ch) { if (!chapters2.has(ch)) chapters2.set(ch, []); chapters2.get(ch).push({ row: item.row, originalIdx: item.idx + startRow2 }); } else noChapter2.push({ row: item.row, originalIdx: item.idx + startRow2 }); });
      for (const item2 of group2.rows) {
        const row2 = item2.row; const absIdx2 = item2.idx + startRow2; if (!Array.isArray(row2)) continue;
        const ch = getChapterFromRow(row2, sheetChapterCol); const newRows = ch ? chapters2.get(ch) : null;
        if (newRows) { let pos = -1; for (let j = 0; j < newRows.length; j++) { if (newRows[j].originalIdx === absIdx2) { pos = j; break; } } if (pos < 0) continue; const oldRows = chapters1.get(ch); if (!oldRows) resultsData.push({ sheet: sheetName, type: 'added', row: row2, originalIdx: absIdx2 }); else { const match = oldRows[pos]; if (!match) resultsData.push({ sheet: sheetName, type: 'added', row: row2, originalIdx: absIdx2 }); else if (rowsEqual(match.row, row2)) resultsData.push({ sheet: sheetName, type: 'same', row: row2, originalIdx: absIdx2 }); else { resultsData.push({ sheet: sheetName, type: 'modified', row: row2, oldRow: match.row, originalIdx: absIdx2, oldOriginalIdx: match.originalIdx }); resultsData.push({ sheet: sheetName, type: 'old', row: match.row, originalIdx: match.originalIdx }); } } }
        else { const ncIdx = noChapter2.findIndex(e => e.originalIdx === absIdx2); const match = ncIdx >= 0 && ncIdx < noChapter1.length ? noChapter1[ncIdx] : null; if (!match) resultsData.push({ sheet: sheetName, type: 'added', row: row2, originalIdx: absIdx2 }); else if (rowsEqual(match.row, row2)) resultsData.push({ sheet: sheetName, type: 'same', row: row2, originalIdx: absIdx2 }); else { resultsData.push({ sheet: sheetName, type: 'modified', row: row2, oldRow: match.row, originalIdx: absIdx2, oldOriginalIdx: match.originalIdx }); resultsData.push({ sheet: sheetName, type: 'old', row: match.row, originalIdx: match.originalIdx }); } }
      }
      const deleted = []; chapters1.forEach((oldRows, chapter) => { const newRows = chapters2.get(chapter); if (!newRows || oldRows.length > newRows.length) { const count = newRows ? oldRows.length - newRows.length : oldRows.length; for (let k = 0; k < count; k++) deleted.push({ chapter, row: oldRows[oldRows.length - 1 - k].row, originalIdx: oldRows[oldRows.length - 1 - k].originalIdx }); } });
      deleted.sort((a, b) => compareChapterNumbers(a.chapter, b.chapter));
      for (const del of deleted) { let insertIdx = resultsData.length; for (let m = groupStartIdx; m < resultsData.length; m++) { const ch = getChapterFromRow(resultsData[m].row, sheetChapterCol); if (ch && compareChapterNumbers(del.chapter, ch) < 0) { insertIdx = m; break; } } resultsData.splice(insertIdx, 0, { sheet: sheetName, type: 'deleted', row: del.row, originalIdx: del.originalIdx }); }
    }
  });
}

function renderResultsSheet(sheetIndex) {
  const wrapper = getSafeElement('resultsPreviewWrapper'); const tabBar = getSafeElement('resultsTabBar');
  if (!wrapper || !tabBar) return;
  if (sheetIndex < 0 || sheetIndex >= resultsSheetNames.length) return;
  try {
    currentResultsSheet = sheetIndex; const sheetName = resultsSheetNames[sheetIndex];
    const sheetData = resultsData.filter(item => item.sheet === sheetName);
    if (resultsSheetNames.length > 1) { let tabHtml = ''; resultsSheetNames.forEach((name, idx) => { const active = idx === sheetIndex ? ' active' : ''; tabHtml += `<button class="results-tab-item${active}" data-sheet-idx="${idx}">${escapeHtml(name)}</button>`; }); tabBar.innerHTML = tabHtml; tabBar.style.display = 'flex'; tabBar.querySelectorAll('.results-tab-item').forEach(btn => { btn.replaceWith(btn.cloneNode(true)); }); tabBar.querySelectorAll('.results-tab-item').forEach(btn => { btn.addEventListener('click', () => { renderResultsSheet(parseInt(btn.dataset.sheetIdx)); }); }); } else tabBar.style.display = 'none';
    const filteredData = currentFilter === 'all' ? sheetData : currentFilter === 'diff' ? sheetData.filter(item => item.type !== 'same') : currentFilter === 'same' ? sheetData.filter(item => item.type === 'same') : sheetData.filter(item => item.type === currentFilter);
    let html = ''; let rowNumber = 0;
    const allMaxCols = sheetData.reduce((max, item) => Math.max(max, Array.isArray(item.row) ? item.row.length : 0), 0);
    html += `<table class="results-excel-preview"><thead><tr><th class="row-header-col"></th>`;
    for (let c = 0; c < allMaxCols; c++) html += `<th style="min-width:60px;max-width:600px;">${getColumnName(c)}</th>`;
    html += '</tr></thead><tbody>';
    filteredData.forEach(item => {
      rowNumber++; const maxCols = Array.isArray(item.row) ? item.row.length : 0; const cells = [];
      let sheetStyles = {};
      if (item.type === 'old' || item.type === 'deleted') sheetStyles = resultsStyles1[sheetName] || {}; else sheetStyles = resultsStyles2[sheetName] || {};
      const originalRowIndex = item.originalIdx;
      for (let i = 0; i < maxCols; i++) { let cellStyle = ''; if (originalRowIndex !== undefined && sheetStyles[originalRowIndex]) { const style = sheetStyles[originalRowIndex][i]; if (style) cellStyle = convertExcelStyleToCSS(style); } const styleAttr = cellStyle ? ` style="${cellStyle}"` : ''; cells.push(`<td${styleAttr}>${escapeHtml(String(item.row[i] ?? ''))}</td>`); }
      const rowClass = item.type === 'added' ? 'results-row-added' : item.type === 'deleted' ? 'results-row-deleted' : item.type === 'modified' ? 'results-row-modified' : item.type === 'old' ? 'results-row-old' : '';
      html += `<tr class="${rowClass}" data-row="${rowNumber}"><td class="row-number-cell">${rowNumber}</td>${cells.join('')}</tr>`;
    });
    html += '</tbody></table>'; wrapper.innerHTML = html;
  } catch (e) {
    console.error('渲染结果表格失败:', e);
    wrapper.innerHTML = '<tr><td>' + i18n.t('render_failed') + '</td></tr>';
  }
}

function renderResults() {
  const { stats } = comparisonResult;
  const filterCountAdded = document.getElementById('filterCountAdded');
  const filterCountDeleted = document.getElementById('filterCountDeleted');
  const filterCountModified = document.getElementById('filterCountModified');
  const filterCountSame = document.getElementById('filterCountSame');
  if (filterCountAdded) filterCountAdded.textContent = stats.added.toLocaleString();
  if (filterCountDeleted) filterCountDeleted.textContent = stats.deleted.toLocaleString();
  if (filterCountModified) filterCountModified.textContent = stats.modified.toLocaleString();
  if (filterCountSame) filterCountSame.textContent = stats.identical.toLocaleString();
  const wrapper = document.getElementById('resultsPreviewWrapper'); const emptyState = document.getElementById('emptyState');
  prepareResultsData(currentChapterColMap, currentStartRowFile1, currentStartRowFile2);
  if (resultsSheetNames.length === 0 || resultsData.length === 0) { wrapper.innerHTML = ''; emptyState.style.display = 'block'; if (resultsData.length === 0) emptyState.querySelector('h3').textContent = i18n.t('no_differences'); }
  else { emptyState.style.display = 'none'; renderResultsSheet(0); }
}

function animateNumber(elementId, target) {
  const element = document.getElementById(elementId);
  if (!element) return;
  const duration = 600; const startTime = performance.now();
  function update(currentTime) {
    const elapsed = currentTime - startTime;
    const progress = Math.min(elapsed / duration, 1);
    const easeOut = 1 - Math.pow(1 - progress, 3);
    const current = Math.floor((target - 0) * easeOut);
    element.textContent = current.toLocaleString();
    if (progress < 1) requestAnimationFrame(update);
  }
  requestAnimationFrame(update);
}
function getDiffLabel(type) { const labels = { added: i18n.t('added'), deleted: i18n.t('deleted'), modified: i18n.t('modified') }; return labels[type] || type; }
function setupFilterButtons() { const buttons = document.querySelectorAll('.filter-btn'); buttons.forEach(btn => { btn.addEventListener('click', () => { buttons.forEach(b => b.classList.remove('active')); btn.classList.add('active'); currentFilter = btn.dataset.filter; if (resultsData.length > 0) renderResultsSheet(currentResultsSheet); }); }); }

function convertExcelStyleToCSS(style) {
  if (!style) return ''; const css = [];
  if (style.font) { if (style.font.sz) css.push(`font-size: ${style.font.sz}pt`); if (style.font.name) css.push(`font-family: ${style.font.name}`); if (style.font.bold) css.push('font-weight: bold'); if (style.font.italic) css.push('font-style: italic'); if (style.font.underline) css.push('text-decoration: underline'); if (style.font.strike) css.push('text-decoration: line-through'); if (style.font.color && style.font.color.rgb) { let rgb = style.font.color.rgb; if (rgb.length === 8 && rgb.startsWith('FF')) rgb = rgb.substring(2); css.push(`color: #${rgb}`); } }
  if (style.alignment) { if (style.alignment.horizontal) css.push(`text-align: ${style.alignment.horizontal}`); if (style.alignment.vertical) { let vAlign = style.alignment.vertical; if (vAlign === 'bottom') vAlign = 'end'; if (vAlign === 'top') vAlign = 'start'; css.push(`vertical-align: ${vAlign}`); } if (style.alignment.wrapText) css.push('white-space: pre-wrap; word-wrap: break-word'); }
  if (style.border) css.push('border: 1px solid #d4d4d4');
  return css.join('; ');
}

function mergeExcelStyles(baseStyle, addFill) {
  const result = baseStyle ? JSON.parse(JSON.stringify(baseStyle)) : {};
  if (addFill) result.fill = JSON.parse(JSON.stringify(addFill));
  if (!result.font) result.font = {};
  if (!result.alignment) result.alignment = {};
  if (!result.border) result.border = {};
  return result;
}

function exportExcelReport() {
  if (!workbook1 || !workbook2) { alert(i18n.t('alert_upload_first')); return; }
  if (!file1 || !file2) { alert(i18n.t('alert_file_lost')); return; }
  if (!comparisonResult) { alert(i18n.t('alert_compare_first')); return; }
  const file1Name = file1.name; const file2Name = file2.name;
  const addedFill = { fgColor: { rgb: '90EE90' } };
  const deletedFill = { fgColor: { rgb: 'FF0000' } };
  const modifiedFill = { fgColor: { rgb: 'FFFF00' } };
  const result1 = getSheetDataWithParsedStyles(workbook1); const result2 = getSheetDataWithParsedStyles(workbook2);
  const sheets1 = result1.sheets; const sheets2 = result2.sheets;
  const styles1 = result1.styles; const styles2 = result2.styles;
  const colWidths2 = result2.colWidths; const merges2 = result2.merges;
  const rowHeights1 = result1.rowHeights; const rowHeights2 = result2.rowHeights;
  const localSheetNames1 = Object.keys(sheets1 || {}); const localSheetNames2 = Object.keys(sheets2 || {});
  const allSheetNames = getUniqueSheetNames(localSheetNames1, localSheetNames2);
  const book = XLSX.utils.book_new();
  let statsAdded = 0; let statsDeleted = 0; let statsModified = 0;

  allSheetNames.forEach(sheetName => {
    const data1 = sheets1[sheetName] || []; const data2 = sheets2[sheetName] || [];
    const sheetStyles1 = styles1[sheetName] || {}; const sheetStyles2 = styles2[sheetName] || {};
    const sheetColWidths = colWidths2[sheetName] || []; const sheetMerges = merges2[sheetName] || [];
    const sheetRowHeights1 = rowHeights1[sheetName] || []; const sheetRowHeights2 = rowHeights2[sheetName] || [];
    let maxCols1 = 0; if (Array.isArray(data1) && data1.length > 0) data1.forEach(row => { if (Array.isArray(row)) maxCols1 = Math.max(maxCols1, row.length); });
    let maxCols2 = 0; if (Array.isArray(data2) && data2.length > 0) data2.forEach(row => { if (Array.isArray(row)) maxCols2 = Math.max(maxCols2, row.length); });
    const maxCols = Math.max(maxCols1, maxCols2);
    const sheetChapterCol = currentChapterColMap[sheetName] !== undefined ? currentChapterColMap[sheetName] : 0;
    const startRow1 = currentStartRowFile1[sheetName] !== undefined ? currentStartRowFile1[sheetName] : 0;
    const startRow2 = currentStartRowFile2[sheetName] !== undefined ? currentStartRowFile2[sheetName] : 0;
    const groups1 = groupDataByLevel1Chapters(data1.slice(startRow1), sheetChapterCol);
    const groups2 = groupDataByLevel1Chapters(data2.slice(startRow2), sheetChapterCol);
    const outputRows = []; const outputStyles = []; const outputOriginalRowIndices = []; const outputSourceFile = [];

    for (let g = 0; g < Math.max(groups1.length, groups2.length); g++) {
      const group1 = groups1[g] || { startIdx: Infinity, rows: [] }; const group2 = groups2[g] || { startIdx: Infinity, rows: [] };
      const groupStartIdx = outputRows.length;
      const chapterRows1 = new Map(); group1.rows.forEach(item => { const chapter = getChapterFromRow(item.row, sheetChapterCol); if (chapter) { if (!chapterRows1.has(chapter)) chapterRows1.set(chapter, []); chapterRows1.get(chapter).push({ row: item.row, originalIdx: item.idx }); } });
      const chapterRows2 = new Map(); group2.rows.forEach(item => { const chapter = getChapterFromRow(item.row, sheetChapterCol); if (chapter) { if (!chapterRows2.has(chapter)) chapterRows2.set(chapter, []); chapterRows2.get(chapter).push({ row: item.row, originalIdx: item.idx }); } });
      const noChapter1 = group1.rows.filter(item => !getChapterFromRow(item.row, sheetChapterCol));
      const noChapter2 = group2.rows.filter(item => !getChapterFromRow(item.row, sheetChapterCol));
      for (const item2 of group2.rows) {
        const row2 = item2.row; const idx2 = item2.idx; if (!Array.isArray(row2)) continue;
        const chapter2 = getChapterFromRow(row2, sheetChapterCol);
        if (chapter2) {
          const oldRows = chapterRows1.get(chapter2);
          if (!oldRows) { addRow(row2, idx2, startRow2, maxCols, sheetStyles2, addedFill, true); }
          else { let pos = -1; for (let j = 0; j < chapterRows2.get(chapter2).length; j++) { if (chapterRows2.get(chapter2)[j].originalIdx === idx2) { pos = j; break; } } if (pos < 0) continue; const match = oldRows[pos]; if (!match) addRow(row2, idx2, startRow2, maxCols, sheetStyles2, addedFill, true); else if (rowsEqual(match.row, row2)) addRow(row2, idx2, startRow2, maxCols, sheetStyles2, null, false); else addModifiedRow(match.row, row2, idx2, match.idx, maxCols, sheetStyles1, sheetStyles2); }
        } else { const ncIdx = noChapter2.findIndex(e => e.idx === idx2); const match = ncIdx >= 0 && ncIdx < noChapter1.length ? noChapter1[ncIdx] : null; if (!match) addRow(row2, idx2, startRow2, maxCols, sheetStyles2, addedFill, true); else if (rowsEqual(match.row, row2)) addRow(row2, idx2, startRow2, maxCols, sheetStyles2, null, false); else addModifiedRow(match.row, row2, idx2, match.idx, maxCols, sheetStyles1, sheetStyles2); }
      }
      const deleted = []; chapterRows1.forEach((oldRows, chapter) => { const newRows = chapterRows2.get(chapter); if (!newRows || oldRows.length > newRows.length) { const count = newRows ? oldRows.length - newRows.length : oldRows.length; for (let k = 0; k < count; k++) deleted.push({ chapter, row: oldRows[oldRows.length - 1 - k].row, originalIdx: oldRows[oldRows.length - 1 - k].originalIdx }); } });
      deleted.sort((a, b) => compareChapterNumbers(a.chapter, b.chapter));
      for (const del of deleted) { let insertIdx = outputRows.length; for (let m = groupStartIdx; m < outputRows.length; m++) { const ch = getChapterFromRow(outputRows[m], sheetChapterCol); if (ch && compareChapterNumbers(del.chapter, ch) < 0) { insertIdx = m; break; } } addDeletedRow(del.row, del.originalIdx, startRow1, maxCols, sheetStyles1, insertIdx); }
    }

    function addRow(row, relativeIdx, startRow, maxCols, sheetStyles, fill, countStats) {
      const realIdx = relativeIdx + startRow; const outputRow = []; const outputStyle = [];
      for (let col = 0; col < Math.max(row.length, maxCols); col++) { const val = (row[col] ?? '').toString(); outputRow.push(val); let baseStyle = {}; if (sheetStyles[realIdx] && sheetStyles[realIdx][col]) baseStyle = sheetStyles[realIdx][col]; if (fill && val.trim() !== '') { outputStyle.push(mergeExcelStyles(baseStyle, fill)); if (countStats) statsAdded++; } else outputStyle.push(mergeExcelStyles(baseStyle, null)); }
      outputRows.push(outputRow); outputOriginalRowIndices.push(realIdx); outputStyles.push(outputStyle); outputSourceFile.push(2);
    }
    function addModifiedRow(row1, row2, relIdx2, relIdx1, maxCols, sStyles1, sStyles2) {
      const maxLen = Math.max(row1.length, row2.length); const realIdx2 = relIdx2 + startRow2; const realIdx1 = relIdx1 + startRow1;
      const newRow = []; const newStyle = [];
      for (let col = 0; col < maxLen; col++) { const v1 = (row1[col] ?? '').toString().trim(); const v2 = (row2[col] ?? '').toString().trim(); newRow.push(v2); let baseStyle = {}; if (sStyles2[realIdx2] && sStyles2[realIdx2][col]) baseStyle = sStyles2[realIdx2][col]; if (v1 !== v2 && v2 !== '') { newStyle.push(mergeExcelStyles(baseStyle, modifiedFill)); statsModified++; } else if (v1 !== v2 && v1 !== '' && v2 === '') { newStyle.push(mergeExcelStyles(baseStyle, deletedFill)); statsDeleted++; } else if (v1 === '' && v2 !== '') { newStyle.push(mergeExcelStyles(baseStyle, addedFill)); statsAdded++; } else newStyle.push(mergeExcelStyles(baseStyle, null)); }
      outputRows.push(newRow); outputOriginalRowIndices.push(realIdx2); outputStyles.push(newStyle); outputSourceFile.push(2);
      const oldRow = []; const oldStyle = [];
      for (let col = 0; col < maxLen; col++) { const v1 = (row1[col] ?? '').toString().trim(); if (v1 !== (row2[col] ?? '').toString().trim() && v1 !== '') { oldRow.push(v1); let baseStyle = {}; if (sStyles1[realIdx1] && sStyles1[realIdx1][col]) baseStyle = sStyles1[realIdx1][col]; oldStyle.push(mergeExcelStyles(baseStyle, modifiedFill)); } else { oldRow.push(''); let baseStyle = {}; if (sStyles1[realIdx1] && sStyles1[realIdx1][col]) baseStyle = sStyles1[realIdx1][col]; oldStyle.push(mergeExcelStyles(baseStyle, null)); } }
      outputRows.push(oldRow); outputOriginalRowIndices.push(realIdx1); outputStyles.push(oldStyle); outputSourceFile.push(1);
    }
    function addDeletedRow(row, relativeIdx, startRow, maxCols, sheetStyles, insertIdx) {
      const realIdx = relativeIdx + startRow; const outputRow = []; const outputStyle = [];
      for (let col = 0; col < Math.max(row.length, maxCols); col++) { const val = (row[col] ?? '').toString(); outputRow.push(val); let baseStyle = {}; if (sheetStyles[realIdx] && sheetStyles[realIdx][col]) baseStyle = sheetStyles[realIdx][col]; if (val.trim() !== '') { outputStyle.push(mergeExcelStyles(baseStyle, deletedFill)); statsDeleted++; } else outputStyle.push(mergeExcelStyles(baseStyle, null)); }
      outputRows.splice(insertIdx, 0, outputRow); outputOriginalRowIndices.splice(insertIdx, 0, realIdx); outputStyles.splice(insertIdx, 0, outputStyle); outputSourceFile.splice(insertIdx, 0, 1);
    }

    function isStyleValid(s) {
      if (!s || typeof s !== 'object') return false;
      if (Object.keys(s).length === 0) return false;
      if (s.font && Object.keys(s.font).length === 0 && s.alignment && Object.keys(s.alignment).length === 0 && 
          s.border && Object.keys(s.border).length === 0 && !s.fill && !s.numFmt) return false;
      return true;
    }
    const sheet = {}; const range = { s: { c: 0, r: 0 }, e: { c: maxCols - 1, r: outputRows.length - 1 } };
    for (let R = 0; R < outputRows.length; R++) {
      for (let C = 0; C < maxCols; C++) {
        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
        const value = outputRows[R][C]; const style = outputStyles[R] ? outputStyles[R][C] : null;
        if (value !== undefined && value !== null && value !== '') {
          const strValue = value.toString(); let cellType = 's';
          if (!isNaN(strValue) && strValue.trim() !== '') cellType = 'n';
          const cellObj = { t: cellType, v: cellType === 'n' ? parseFloat(strValue) : strValue };
          if (isStyleValid(style)) cellObj.s = style;
          sheet[cellRef] = cellObj;
        } else if (isStyleValid(style)) { sheet[cellRef] = { t: 's', v: '', s: style }; }
      }
    }
    sheet['!ref'] = XLSX.utils.encode_range(range);
    const colWidths = []; for (let i = 0; i < maxCols; i++) { let maxLen = 10; for (let j = 0; j < outputRows.length; j++) { const cellValue = String(outputRows[j][i] || ''); if (cellValue.length > maxLen) maxLen = Math.min(cellValue.length, 50); } colWidths.push({ wch: maxLen }); }
    sheet['!cols'] = colWidths;
    const outputRowHeights = [];
    for (let r = 0; r < outputRows.length; r++) {
      const origIdx = outputOriginalRowIndices[r];
      const src = outputSourceFile[r];
      const heights = src === 1 ? sheetRowHeights1 : sheetRowHeights2;
      if (origIdx !== undefined && heights[origIdx]) outputRowHeights[r] = heights[origIdx];
    }
    if (outputRowHeights.length > 0) sheet['!rows'] = outputRowHeights;
    if (sheetMerges && sheetMerges.length > 0) { const adjustedMerges = sheetMerges.map(merge => { const ns = merge.s.r - startRow2; const ne = merge.e.r - startRow2; if (ne >= 0 && ns < outputRows.length) return { s: { r: Math.max(0, ns), c: merge.s.c }, e: { r: Math.min(outputRows.length - 1, ne), c: merge.e.c } }; return null; }).filter(m => m !== null); if (adjustedMerges.length > 0) sheet['!merges'] = adjustedMerges; }
    const safeName = sheetName.length > 31 ? sheetName.substring(0, 28) + '...' : sheetName;
    XLSX.utils.book_append_sheet(book, sheet, safeName);
  });

  const headerStyle = { fill: { fgColor: { rgb: '6366F1' } }, font: { bold: true, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center' }, border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } };
  const normalStyle = { border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } };
  const statsSheetData = [
    [i18n.t('stats_item'), i18n.t('quantity')],
    [i18n.t('added'), statsAdded],
    [i18n.t('deleted'), statsDeleted],
    [i18n.t('modified'), statsModified],
    ['', ''],
    [i18n.t('old_file'), file1Name],
    [i18n.t('new_file'), file2Name],
    [i18n.t('generated_time'), new Date().toLocaleString()]
  ];
  const statsSheet = XLSX.utils.aoa_to_sheet(statsSheetData); statsSheet['A1'].s = headerStyle; statsSheet['B1'].s = headerStyle; for (let i = 2; i <= statsSheetData.length; i++) { statsSheet['A' + i].s = normalStyle; statsSheet['B' + i].s = normalStyle; } statsSheet['!cols'] = [{ wch: 15 }, { wch: Math.max(file1Name.length, file2Name.length, 40) }];
  const lang = i18n.getCurrentLang();
  XLSX.utils.book_append_sheet(book, statsSheet, lang === 'zh' ? '统计摘要' : (lang === 'en' ? 'Summary' : (lang === 'es' ? 'Resumen' : (lang === 'de' ? 'Zusammenfassung' : 'R\u00e9sum\u00e9'))));
  const summaryData = [
    [i18n.t('color_legend'), i18n.t('meaning')],
    [i18n.t('filter_added'), i18n.t('green_meaning')],
    [i18n.t('filter_deleted'), i18n.t('red_meaning')],
    [i18n.t('filter_modified') + ' (' + i18n.t('new_file') + ')', i18n.t('yellow_bold')],
    [i18n.t('filter_modified') + ' (' + i18n.t('old_file') + ')', i18n.t('yellow_italic')]
  ];
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData); summarySheet['A1'].s = headerStyle; summarySheet['B1'].s = headerStyle; summarySheet['A2'].s = { ...normalStyle, fill: addedFill }; summarySheet['A3'].s = { ...normalStyle, fill: deletedFill }; summarySheet['A4'].s = { ...normalStyle, fill: modifiedFill }; summarySheet['A5'].s = { ...normalStyle, fill: modifiedFill }; summarySheet['!cols'] = [{ wch: 25 }, { wch: 40 }];
  XLSX.utils.book_append_sheet(book, summarySheet, lang === 'zh' ? '说明' : (lang === 'en' ? 'Legend' : (lang === 'es' ? 'Leyenda' : (lang === 'de' ? 'Legende' : 'L\u00e9gende'))));
  XLSX.writeFile(book, `${i18n.t('file_compare_report')}_${file2Name}`, {
    bookType: 'xlsx',
    cellStyles: true,
    bookSST: true
  });
}

function exportReport() {
  if (!comparisonResult) return;
  const { differences, stats } = comparisonResult; const file1Name = file1.name; const file2Name = file2.name;
  const t = i18n.t.bind(i18n); const lang = i18n.getCurrentLang(); const langCode = lang === 'zh' ? 'zh-CN' : lang;
  const tAdded = t('added'); const tDeleted = t('deleted'); const tModified = t('modified'); const tIdentical = t('filter_same');
  const tGenerated = t('generated_time'); const tOldFile = t('old_file'); const tNewFile = t('new_file');
  const tDiffDetail = t('differences'); const tNoDiff = t('no_differences'); const tFooter = t('footer_text');
  const tReportTitle = t('file_compare_report');
  const html = `<div>${tReportTitle}</div>`;
  const statsHtml = `<div><b>${tAdded}</b>: ${stats.added} | <b>${tDeleted}</b>: ${stats.deleted} | <b>${tModified}</b>: ${stats.modified} | <b>${tIdentical}</b>: ${stats.identical}</div>`;
  const metaHtml = `<div>${tOldFile}: ${escapeHtml(file1Name)} | ${tNewFile}: ${escapeHtml(file2Name)}<br>${tGenerated}: ${new Date().toLocaleString()}</div>`;
  const diffItems = differences.map(diff => `<div class="${diff.type}"><b>${getDiffLabel(diff.type)}</b> ${diff.position} ${diff.oldValue ? '[' + escapeHtml(diff.oldValue) + ']' : ''} ${diff.newValue ? '→ [' + escapeHtml(diff.newValue) + ']' : ''}</div>`).join('\n');
  const fullHtml = `<!DOCTYPE html><html lang="${langCode}"><head><meta charset="UTF-8"><title>${tReportTitle}</title><style>body{font-family:sans-serif;padding:2rem;max-width:800px;margin:0 auto}.added{color:green}.deleted{color:red}.modified{color:orange}</style></head><body><h1>${tReportTitle}</h1>${metaHtml}<h2>${tDiffDetail}</h2>${statsHtml}<hr/>${diffItems}<footer>${tFooter}</footer></body></html>`;
  const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `${tReportTitle}_${new Date().toISOString().slice(0, 10)}.html`; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
}

function escapeHtml(text) { const div = document.createElement('div'); div.textContent = text; return div.innerHTML; }

function setupPreview() {
  if (!workbook1 || !workbook2) return;
  const result1 = getSheetDataWithParsedStyles(workbook1); const result2 = getSheetDataWithParsedStyles(workbook2);
  const sheets1 = result1.sheets; const sheets2 = result2.sheets;
  const styles2 = result2.styles; const styles1 = result1.styles;
  previewStyles1 = styles1; previewStyles2 = styles2;
  previewData = []; previewSheetNames = [];
  const localSheetNames1 = Object.keys(sheets1 || {}); const localSheetNames2 = Object.keys(sheets2 || {});
  const allSheetNames = getUniqueSheetNames(localSheetNames1, localSheetNames2);
  allSheetNames.forEach(sheetName => {
    const data1 = sheets1[sheetName] || []; const data2 = sheets2[sheetName] || [];
    previewSheetNames.push(sheetName);
    const sheetChapterCol = currentChapterColMap[sheetName] !== undefined ? currentChapterColMap[sheetName] : 0;
    const startRow1 = currentStartRowFile1[sheetName] !== undefined ? currentStartRowFile1[sheetName] : 0;
    const startRow2 = currentStartRowFile2[sheetName] !== undefined ? currentStartRowFile2[sheetName] : 0;
    const groups1 = groupDataByLevel1Chapters(data1.slice(startRow1), sheetChapterCol); const groups2 = groupDataByLevel1Chapters(data2.slice(startRow2), sheetChapterCol);
    for (let g = 0; g < Math.max(groups1.length, groups2.length); g++) {
      const group1 = groups1[g] || { startIdx: Infinity, rows: [] }; const group2 = groups2[g] || { startIdx: Infinity, rows: [] };
      const groupStartIdx = previewData.length;
      const chapters1 = new Map(); const noChapter1 = []; group1.rows.forEach(item => { const ch = getChapterFromRow(item.row, sheetChapterCol); if (ch) { if (!chapters1.has(ch)) chapters1.set(ch, []); chapters1.get(ch).push({ row: item.row, originalIdx: item.idx + startRow1 }); } else noChapter1.push({ row: item.row, originalIdx: item.idx + startRow1 }); });
      const chapters2 = new Map(); const noChapter2 = []; group2.rows.forEach(item => { const ch = getChapterFromRow(item.row, sheetChapterCol); if (ch) { if (!chapters2.has(ch)) chapters2.set(ch, []); chapters2.get(ch).push({ row: item.row, originalIdx: item.idx + startRow2 }); } else noChapter2.push({ row: item.row, originalIdx: item.idx + startRow2 }); });
      for (const item2 of group2.rows) { const row2 = item2.row; const absIdx2 = item2.idx + startRow2; if (!Array.isArray(row2)) continue; const ch = getChapterFromRow(row2, sheetChapterCol); const newRows = ch ? chapters2.get(ch) : null; if (newRows) { let pos = -1; for (let j = 0; j < newRows.length; j++) { if (newRows[j].originalIdx === absIdx2) { pos = j; break; } } if (pos < 0) continue; const oldRows = chapters1.get(ch); if (!oldRows) previewData.push({ sheet: sheetName, type: 'added', row: row2, originalIdx: absIdx2 }); else { const match = oldRows[pos]; if (!match) previewData.push({ sheet: sheetName, type: 'added', row: row2, originalIdx: absIdx2 }); else if (rowsEqual(match.row, row2)) previewData.push({ sheet: sheetName, type: 'same', row: row2, originalIdx: absIdx2 }); else { previewData.push({ sheet: sheetName, type: 'modified', row: row2, oldRow: match.row, originalIdx: absIdx2, oldOriginalIdx: match.originalIdx }); previewData.push({ sheet: sheetName, type: 'old', row: match.row, originalIdx: match.originalIdx }); } } } else { const ncIdx = noChapter2.findIndex(e => e.originalIdx === absIdx2); const match = ncIdx >= 0 && ncIdx < noChapter1.length ? noChapter1[ncIdx] : null; if (!match) previewData.push({ sheet: sheetName, type: 'added', row: row2, originalIdx: absIdx2 }); else if (rowsEqual(match.row, row2)) previewData.push({ sheet: sheetName, type: 'same', row: row2, originalIdx: absIdx2 }); else { previewData.push({ sheet: sheetName, type: 'modified', row: row2, oldRow: match.row, originalIdx: absIdx2, oldOriginalIdx: match.originalIdx }); previewData.push({ sheet: sheetName, type: 'old', row: match.row, originalIdx: match.originalIdx }); } } }
      const deleted = []; chapters1.forEach((oldRows, chapter) => { const newRows = chapters2.get(chapter); if (!newRows || oldRows.length > newRows.length) { const count = newRows ? oldRows.length - newRows.length : oldRows.length; for (let k = 0; k < count; k++) deleted.push({ chapter, row: oldRows[oldRows.length - 1 - k].row, originalIdx: oldRows[oldRows.length - 1 - k].originalIdx }); } });
      deleted.sort((a, b) => compareChapterNumbers(a.chapter, b.chapter));
      for (const del of deleted) { let insertIdx = previewData.length; for (let m = groupStartIdx; m < previewData.length; m++) { const ch = getChapterFromRow(previewData[m].row, sheetChapterCol); if (ch && compareChapterNumbers(del.chapter, ch) < 0) { insertIdx = m; break; } } previewData.splice(insertIdx, 0, { sheet: sheetName, type: 'deleted', row: del.row, originalIdx: del.originalIdx }); }
    }
  });
}

function renderPreviewSheet(sheetIndex) {
  const body = getSafeElement('previewBody'); const tabBar = getSafeElement('tabBar');
  if (!body || !tabBar) return;
  if (sheetIndex < 0 || sheetIndex >= previewSheetNames.length) return;
  try {
    currentPreviewSheetIndex = sheetIndex; const sheetName = previewSheetNames[sheetIndex];
    const sheetData = previewData.filter(item => item.sheet === sheetName);
    if (previewSheetNames.length > 1) { let tabHtml = ''; previewSheetNames.forEach((name, idx) => { const active = idx === sheetIndex ? ' active' : ''; tabHtml += `<button class="tab-item${active}" data-sheet-idx="${idx}">${escapeHtml(name)}</button>`; }); tabBar.innerHTML = tabHtml; tabBar.style.display = 'flex'; } else tabBar.style.display = 'none';
    let html = ''; let rowNumber = 0;
    const allMaxCols = sheetData.reduce((max, item) => Math.max(max, Array.isArray(item.row) ? item.row.length : 0), 0);
    html += `<div class="preview-content active"><table class="excel-preview"><thead><tr><th class="row-header-col"></th>`;
    for (let c = 0; c < allMaxCols; c++) html += `<th style="min-width:60px;max-width:600px;"><span class="col-resize-handle"></span>${getColumnName(c)}</th>`;
    html += '</tr></thead><tbody>';
    sheetData.forEach(item => {
      rowNumber++; const maxCols = Array.isArray(item.row) ? item.row.length : 0; const cells = [];
      let sheetStyles = {}; if (item.type === 'old' || item.type === 'deleted') sheetStyles = previewStyles1[sheetName] || {}; else sheetStyles = previewStyles2[sheetName] || {};
      const originalRowIndex = item.originalIdx;
      for (let i = 0; i < maxCols; i++) { let cellStyle = ''; if (originalRowIndex !== undefined && sheetStyles[originalRowIndex]) { const style = sheetStyles[originalRowIndex][i]; if (style) cellStyle = convertExcelStyleToCSS(style); } const styleAttr = cellStyle ? ` style="${cellStyle}"` : ''; cells.push(`<td${styleAttr}>${escapeHtml(String(item.row[i] ?? ''))}</td>`); }
      const rowClass = item.type === 'added' ? 'preview-row-added' : item.type === 'deleted' ? 'preview-row-deleted' : item.type === 'modified' ? 'preview-row-modified' : item.type === 'old' ? 'preview-row-old' : '';
      html += `<tr class="${rowClass}" data-row="${rowNumber}"><td class="row-number-cell">${rowNumber}</td>${cells.join('')}</tr>`;
    });
    html += '</tbody></table></div>'; body.innerHTML = html;
    setTimeout(() => { body.scrollTop = 0; body.scrollLeft = 0; }, 0);
    tabBar.querySelectorAll('.tab-item').forEach(btn => { btn.replaceWith(btn.cloneNode(true)); });
    tabBar.querySelectorAll('.tab-item').forEach(btn => { btn.addEventListener('click', () => { renderPreviewSheet(parseInt(btn.dataset.sheetIdx)); }); });
    initColumnResize();
  } catch (e) {
    console.error('渲染预览表格失败:', e);
    body.innerHTML = '<div class="preview-content active"><p>' + i18n.t('render_failed') + '</p></div>';
  }
}

function initColumnResize() {
  const table = document.querySelector('.excel-preview'); if (!table) return;
  const ths = table.querySelectorAll('th:not(.row-header-col)');
  ths.forEach((th, index) => {
    const handle = th.querySelector('.col-resize-handle'); if (!handle) return;
    let startX, startWidth;
    handle.addEventListener('mousedown', (e) => {
      e.preventDefault(); e.stopPropagation(); th.classList.add('resizing'); startX = e.clientX; startWidth = th.offsetWidth;
      const onMouseMove = (e) => { const diff = e.clientX - startX; const newWidth = Math.max(60, Math.min(600, startWidth + diff)); th.style.width = newWidth + 'px'; th.style.minWidth = newWidth + 'px'; th.style.maxWidth = 'none'; const colIndex = index + 1; table.querySelectorAll('tbody tr').forEach(row => { const td = row.children[colIndex]; if (td) { td.style.width = newWidth + 'px'; td.style.minWidth = newWidth + 'px'; } }); };
      const onMouseUp = () => { th.classList.remove('resizing'); document.removeEventListener('mousemove', onMouseMove); document.removeEventListener('mouseup', onMouseUp); };
      document.addEventListener('mousemove', onMouseMove); document.addEventListener('mouseup', onMouseUp);
    });
  });
}

function showPreview() { if (previewData.length === 0) setupPreview(); if (previewData.length === 0) return; const modal = document.getElementById('previewModal'); modal.style.display = 'flex'; renderPreviewSheet(0); requestAnimationFrame(() => { const body = document.getElementById('previewBody'); if (body) { body.scrollTop = 0; body.scrollLeft = 0; } }); }
function hidePreview() { const modal = document.getElementById('previewModal'); modal.style.display = 'none'; toggleFullscreen(false); }
function toggleFullscreen(forceState) { const container = document.querySelector('.modal-container'); const fullscreenIcon = document.querySelector('.icon-fullscreen'); const exitFullscreenIcon = document.querySelector('.icon-exit-fullscreen'); const fullscreenBtn = document.getElementById('fullscreenBtn'); isFullscreen = forceState !== undefined ? forceState : !isFullscreen; if (isFullscreen) { container.classList.add('fullscreen'); fullscreenIcon.style.display = 'none'; exitFullscreenIcon.style.display = 'block'; fullscreenBtn.setAttribute('title', i18n.t('exit_fullscreen')); } else { container.classList.remove('fullscreen'); fullscreenIcon.style.display = 'block'; exitFullscreenIcon.style.display = 'none'; fullscreenBtn.setAttribute('title', i18n.t('fullscreen')); } }
function swapFiles() {
  const tempFile = file1; const tempWorkbook = workbook1; const tempPreviewData = previewData1; const tempSheetNames = sheetNames1; const tempStartCells = startCellPerSheetFile1; const tempStyles = window._parsedStylesFile1;
  file1 = file2; workbook1 = workbook2; previewData1 = previewData2; sheetNames1 = sheetNames2; startCellPerSheetFile1 = startCellPerSheetFile2; window._parsedStylesFile1 = window._parsedStylesFile2;
  file2 = tempFile; workbook2 = tempWorkbook; previewData2 = tempPreviewData; sheetNames2 = tempSheetNames; startCellPerSheetFile2 = tempStartCells; window._parsedStylesFile2 = tempStyles;
  const file1Info = document.getElementById('file1Info'); const file2Info = document.getElementById('file2Info');
  const file1Name = document.getElementById('file1Name'); const file2Name = document.getElementById('file2Name');
  if (file1Info && file2Info && file1Name && file2Name) {
    const tempName = file1Name.textContent; const tempDisplay = file1Info.style.display;
    file1Name.textContent = file2Name.textContent; file1Info.style.display = file2Info.style.display;
    file2Name.textContent = tempName; file2Info.style.display = tempDisplay;
  }
  const resultsSection = document.getElementById('resultsSection'); if (resultsSection) resultsSection.style.display = 'none';
  updateCompareButton(); renderSheetTabs(); updateFilePreview();
}

function toggleResultsMaximize() {
  isResultsMaximized = !isResultsMaximized;
  const diffContainer = document.querySelector('.diff-container');
  const btn = getSafeElement('maximizeResultsBtn');
  if (!diffContainer) return;
  if (isResultsMaximized) {
    diffContainer.classList.add('maximized');
    const overlay = document.createElement('div');
    overlay.className = 'diff-overlay-bg';
    overlay.id = 'diffOverlayBg';
    overlay.addEventListener('click', toggleResultsMaximize);
    document.body.appendChild(overlay);
    document.addEventListener('keydown', handleResultsMaximizeKeydown);
    if (btn) {
      const maximizeIcon = btn.querySelector('.icon-maximize');
      const restoreIcon = btn.querySelector('.icon-restore');
      if (maximizeIcon) maximizeIcon.style.display = 'none';
      if (restoreIcon) restoreIcon.style.display = 'block';
      btn.setAttribute('title', i18n.t('max_preview_exit'));
    }
  } else {
    diffContainer.classList.remove('maximized');
    const overlay = getSafeElement('diffOverlayBg');
    if (overlay) overlay.remove();
    document.removeEventListener('keydown', handleResultsMaximizeKeydown);
    if (btn) {
      const maximizeIcon = btn.querySelector('.icon-maximize');
      const restoreIcon = btn.querySelector('.icon-restore');
      if (maximizeIcon) maximizeIcon.style.display = 'block';
      if (restoreIcon) restoreIcon.style.display = 'none';
      btn.setAttribute('title', i18n.t('max_preview_preview'));
    }
  }
}

function handleResultsMaximizeKeydown(e) {
  if (e.key === 'Escape' && isResultsMaximized) toggleResultsMaximize();
}

document.addEventListener('DOMContentLoaded', () => {
  setupUploadZone('file1Zone', 'file1Input', 1); setupUploadZone('file2Zone', 'file2Input', 2);
  document.getElementById('directCompareBtn').addEventListener('click', compareFiles);
  document.getElementById('selectCellBtn').addEventListener('click', showCellSelector);
  document.getElementById('confirmCompareBtn').addEventListener('click', compareFiles);
  document.getElementById('exportExcelBtn').addEventListener('click', exportExcelReport);
  document.getElementById('modalCloseBtn').addEventListener('click', hidePreview);
  document.getElementById('swapBtn').addEventListener('click', swapFiles);
  document.getElementById('previewModal').addEventListener('click', (e) => { if (e.target.id === 'previewModal') hidePreview(); });
  document.getElementById('startCellReset1').addEventListener('click', () => resetStartCell(1));
  document.getElementById('startCellReset2').addEventListener('click', () => resetStartCell(2));
  document.getElementById('removeFileBtn1').addEventListener('click', () => removeFile(1));
  document.getElementById('removeFileBtn2').addEventListener('click', () => removeFile(2));
  document.getElementById('maximizeResultsBtn').addEventListener('click', toggleResultsMaximize);
  setupFilterButtons();
});
