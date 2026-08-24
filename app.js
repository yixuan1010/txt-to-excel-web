'use strict';

const LINE_RE = /^(?<date>\d{4}-\d{2}-\d{2})\t(?<time>\d{2}:\d{2}:\d{2})\t(?<flow>[^\t]*)\t(?<message>.*)$/;
const MOVE_RE = /^Move\s+to\s+scan\s+position\s+X\s*\((?<x>[^)]+)\)\s*&\s*Y\s*\((?<y>[^)]+)\)\s*\.?\s*$/i;
const SCAN_RE = /^Scan\s+Laser\s*\((?<value>[^)]+)\)\s*\.?\s*$/i;

const state = { files: [] };
const $ = (id) => document.getElementById(id);
const elements = {
  dropZone: $('dropZone'), fileInput: $('fileInput'), folderInput: $('folderInput'),
  fileBtn: $('fileBtn'), folderBtn: $('folderBtn'), clearBtn: $('clearBtn'),
  convertBtn: $('convertBtn'), fileSummary: $('fileSummary'), fileCount: $('fileCount'),
  totalSize: $('totalSize'), fileList: $('fileList'), progressWrap: $('progressWrap'),
  progressBar: $('progressBar'), progressText: $('progressText'), resultBox: $('resultBox')
};

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 ** 2).toFixed(1)} MB`;
}

function fileKey(file) {
  return `${file.webkitRelativePath || file.name}|${file.size}|${file.lastModified}`;
}

function addFiles(fileList) {
  const existing = new Set(state.files.map(fileKey));
  for (const file of Array.from(fileList)) {
    if (!file.name.toLowerCase().endsWith('.txt')) continue;
    const key = fileKey(file);
    if (!existing.has(key)) {
      state.files.push(file);
      existing.add(key);
    }
  }
  state.files.sort((a, b) => (a.webkitRelativePath || a.name).localeCompare(
    b.webkitRelativePath || b.name, undefined, { sensitivity: 'base' }
  ));
  renderFiles();
}

function renderFiles() {
  const hasFiles = state.files.length > 0;
  elements.fileSummary.classList.toggle('hidden', !hasFiles);
  elements.clearBtn.disabled = !hasFiles;
  elements.convertBtn.disabled = !hasFiles;
  elements.fileCount.textContent = String(state.files.length);
  elements.totalSize.textContent = formatBytes(state.files.reduce((sum, file) => sum + file.size, 0));
  elements.fileList.innerHTML = '';
  for (const file of state.files) {
    const item = document.createElement('li');
    const name = document.createElement('span');
    name.className = 'file-name';
    name.textContent = file.webkitRelativePath || file.name;
    const size = document.createElement('span');
    size.className = 'file-size';
    size.textContent = formatBytes(file.size);
    item.append(name, size);
    elements.fileList.append(item);
  }
  if (!hasFiles) hideResult();
}

function hideResult() {
  elements.resultBox.className = 'result-box hidden';
  elements.resultBox.textContent = '';
  elements.progressWrap.classList.add('hidden');
  elements.progressBar.style.width = '0%';
}

function showResult(type, message) {
  elements.resultBox.className = `result-box ${type}`;
  elements.resultBox.textContent = message;
}

function decodeText(buffer) {
  const bytes = new Uint8Array(buffer);
  let encodings;
  if (bytes[0] === 0xff && bytes[1] === 0xfe) encodings = ['utf-16le'];
  else if (bytes[0] === 0xfe && bytes[1] === 0xff) encodings = ['utf-16be'];
  else if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) encodings = ['utf-8'];
  else encodings = ['utf-8', 'big5', 'utf-16le'];

  for (const encoding of encodings) {
    try {
      const text = new TextDecoder(encoding, { fatal: true }).decode(bytes);
      if (!text.includes('\u0000')) return text.replace(/^\uFEFF/, '');
    } catch (_) { /* try next encoding */ }
  }
  return new TextDecoder('utf-8').decode(bytes).replace(/^\uFEFF/, '');
}

function parseLine(line) {
  const match = line.replace(/[\r\n]+$/, '').match(LINE_RE);
  return match ? match.groups : null;
}

function truncateDecimal(rawValue, digits) {
  const text = rawValue.trim();
  const number = Number(text);
  if (!Number.isFinite(number)) return text;
  const factor = 10 ** digits;
  const truncated = Math.trunc(number * factor) / factor;
  const normalized = Object.is(truncated, -0) ? 0 : truncated;
  return normalized.toFixed(digits);
}

function extractPairs(text, fileName) {
  const lines = text.split(/\r?\n/);
  const records = [];
  let skippedMalformed = 0;

  for (let index = 0; index < lines.length - 1; index += 1) {
    const first = parseLine(lines[index]);
    const second = parseLine(lines[index + 1]);
    if (!first || !second) {
      if (!first && lines[index].trim()) skippedMalformed += 1;
      continue;
    }
    const move = first.message.trim().match(MOVE_RE);
    const scan = second.message.trim().match(SCAN_RE);
    if (!move || !scan) continue;
    if (first.flow.trim().toLowerCase() !== 'laser scan flow') continue;
    if (second.flow.trim().toLowerCase() !== 'laser scan flow') continue;

    const x = truncateDecimal(move.groups.x, 1);
    const y = truncateDecimal(move.groups.y, 1);
    const laser = truncateDecimal(scan.groups.value, 3);
    records.push({
      Date: first.date,
      Time: first.time.slice(0, 5),
      Value: `Scan Laser (${laser})`,
      cordinate: `Move to scan position X (${x}) & Y (${y})`,
      _sortTime: first.time,
      _file: fileName,
      _line: index + 1
    });
  }
  return { records, skippedMalformed };
}

function compareRecords(a, b) {
  const timestampCompare = `${a.Date} ${a._sortTime}`.localeCompare(`${b.Date} ${b._sortTime}`);
  if (timestampCompare) return timestampCompare;
  const fileCompare = a._file.localeCompare(b._file, undefined, { sensitivity: 'base' });
  return fileCompare || a._line - b._line;
}

async function buildWorkbook(records) {
  if (typeof ExcelJS === 'undefined') throw new Error('Excel 元件載入失敗，請確認網路連線後重新整理頁面。');
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Laser Scan TXT to Excel Web';
  workbook.created = new Date();
  const sheet = workbook.addWorksheet('Laser Scan', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });
  sheet.columns = [
    { header: 'Date', key: 'Date', width: 13 },
    { header: 'Time', key: 'Time', width: 10 },
    { header: 'Value', key: 'Value', width: 22 },
    { header: 'cordinate', key: 'cordinate', width: 52 }
  ];
  records.forEach((row) => sheet.addRow({
    Date: row.Date, Time: row.Time, Value: row.Value, cordinate: row.cordinate
  }));
  const header = sheet.getRow(1);
  header.font = { color: { argb: 'FFFFFFFF' }, bold: true };
  header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
  header.alignment = { horizontal: 'center' };
  sheet.autoFilter = { from: 'A1', to: `D${Math.max(1, records.length + 1)}` };
  for (let rowNumber = 2; rowNumber <= records.length + 1; rowNumber += 1) {
    sheet.getRow(rowNumber).eachCell((cell) => { cell.numFmt = '@'; });
  }
  return workbook.xlsx.writeBuffer();
}

function downloadBuffer(buffer, fileName) {
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function convert() {
  if (!state.files.length) return;
  elements.convertBtn.disabled = true;
  elements.progressWrap.classList.remove('hidden');
  elements.resultBox.classList.add('hidden');
  const allRecords = [];
  let totalSkipped = 0;

  try {
    for (let index = 0; index < state.files.length; index += 1) {
      const file = state.files[index];
      elements.progressText.textContent = `正在處理 ${index + 1}/${state.files.length}：${file.name}`;
      elements.progressBar.style.width = `${Math.round((index / state.files.length) * 85)}%`;
      const text = decodeText(await file.arrayBuffer());
      const result = extractPairs(text, file.webkitRelativePath || file.name);
      allRecords.push(...result.records);
      totalSkipped += result.skippedMalformed;
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    allRecords.sort(compareRecords);
    elements.progressText.textContent = '正在建立 Excel...';
    elements.progressBar.style.width = '92%';
    const buffer = await buildWorkbook(allRecords);
    downloadBuffer(buffer, 'result.xlsx');
    elements.progressBar.style.width = '100%';
    elements.progressText.textContent = '完成';
    const skippedText = totalSkipped ? `，另有 ${totalSkipped} 行不是預期的 Tab 格式，已略過` : '';
    showResult(allRecords.length ? 'success' : 'warning', `完成：處理 ${state.files.length} 個 TXT 檔，輸出 ${allRecords.length} 組連續動作${skippedText}。`);
  } catch (error) {
    console.error(error);
    showResult('error', `轉換失敗：${error.message || String(error)}`);
    elements.progressText.textContent = '處理失敗';
  } finally {
    elements.convertBtn.disabled = false;
  }
}

elements.fileBtn.addEventListener('click', (event) => { event.stopPropagation(); elements.fileInput.click(); });
elements.folderBtn.addEventListener('click', (event) => { event.stopPropagation(); elements.folderInput.click(); });
elements.dropZone.addEventListener('click', () => elements.fileInput.click());
elements.dropZone.addEventListener('keydown', (event) => {
  if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); elements.fileInput.click(); }
});
elements.fileInput.addEventListener('change', (event) => { addFiles(event.target.files); event.target.value = ''; });
elements.folderInput.addEventListener('change', (event) => { addFiles(event.target.files); event.target.value = ''; });
['dragenter', 'dragover'].forEach((name) => elements.dropZone.addEventListener(name, (event) => {
  event.preventDefault(); elements.dropZone.classList.add('dragover');
}));
['dragleave', 'drop'].forEach((name) => elements.dropZone.addEventListener(name, (event) => {
  event.preventDefault(); elements.dropZone.classList.remove('dragover');
}));
elements.dropZone.addEventListener('drop', (event) => addFiles(event.dataTransfer.files));
elements.clearBtn.addEventListener('click', () => { state.files = []; renderFiles(); });
elements.convertBtn.addEventListener('click', convert);
