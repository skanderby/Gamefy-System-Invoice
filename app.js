/**
 * Gamefy Invoice Builder — Plain HTML/JS (like Tuneps)
 * - Drop PDFs named day-month (e.g. 15-03.pdf)
 * - Extract income/expense table via pdf.js
 * - One Jspreadsheet Excel editor tab per file
 * - Reconciliation: Total Amounts, Total Cash, Total TPE, Difference
 * - Download Daily Excel (one sheet per day)
 * - Download Weekly Excel (all rows per week + accounting summary)
 */

let pdfjsLib;

function loadPdfJS() {
  return new Promise((resolve, reject) => {
    if (window.pdfjsLib) {
      pdfjsLib = window.pdfjsLib;
      resolve();
      return;
    }
    const s = document.createElement('script');
    s.src = 'https://unpkg.com/pdfjs-dist@4.8.69/build/pdf.min.js';
    s.onload = () => {
      pdfjsLib = window.pdfjsLib;
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://unpkg.com/pdfjs-dist@4.8.69/build/pdf.worker.min.js';
      resolve();
    };
    s.onerror = reject;
    document.head.appendChild(s);
  });
}

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let pdfFiles = [];        // { file, name, day, month, year, entries[], processed, cashAmount, tpeAmount }
let spreadsheetInstances = {}; // tabIndex -> jspreadsheet instance
let currentYear = new Date().getFullYear();
let libsReady = false;
let isProcessing = false;
let isDownloading = false;
let activeTabIndex = 0;
let isUploadCollapsed = false;

const EXCEL_COLS = [
  { key: 'rowNum', label: '#', width: 50 },
  { key: 'date', label: 'Date', width: 110 },
  { key: 'type', label: 'Type', width: 100 },
  { key: 'description', label: 'Description', width: 200 },
  { key: 'comment', label: 'Comment', width: 200 },
  { key: 'amount', label: 'Amount', width: 120 },
  { key: 'sourceFile', label: 'Source File', width: 150 },
];

const CATEGORIES = [
  { key: 'COMPUTER', label: 'COMPUTER', keywords: ['computer'] },
  { key: 'MEMBER', label: 'Member Transactions', keywords: ['member'] },
  { key: 'PLAYSTATION', label: 'Playstation Incomes', keywords: ['playstation'] },
  { key: 'ORDER', label: 'Order Incomes', keywords: ['order'] },
  { key: 'COWORKING', label: 'Co-working', keywords: ['co-working', 'coworking', 'co working'] },
  { key: 'PASS', label: 'PASS EVENEMENT', keywords: ['pass', 'evenement', 'événement'] },
];

// ---------------------------------------------------------------------------
// DOM elements
// ---------------------------------------------------------------------------
let dropZone, fileInput, browseLink, fileList, btnExtract, extractSpinner;
let stepDrop, stepTabs, tabsHeader, tabsPanels, btnMoreFiles, btnDownloadDaily, btnDownloadWeekly;
let statsFiles, statsEntries, collapseToggle;

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------
window.addEventListener('DOMContentLoaded', async () => {
  initDOM();
  setupDragDrop();
  setupFileInput();
  setupButtons();
  setupCollapseToggle();
  updateStats();
  updateDownloadButtons();

  try {
    await loadPdfJS();
    libsReady = true;
    console.log('pdf.js loaded:', typeof pdfjsLib);
  } catch (e) {
    console.error('pdf.js load failed:', e);
    showToast('error', 'Library load failed', 'pdf.js did not load. Refresh the page.');
  }
  if (btnExtract) btnExtract.disabled = false;
});

// ---------------------------------------------------------------------------
// Collapse / Expand toggle
// ---------------------------------------------------------------------------
function setupCollapseToggle() {
  if (!collapseToggle || !stepDrop) return;
  collapseToggle.addEventListener('click', () => {
    isCollapsed = !isCollapsed;
    if (isCollapsed) {
      stepDrop.style.display = 'none';
      collapseToggle.textContent = 'Expand';
      collapseToggle.classList.add('btn-primary');
      collapseToggle.classList.remove('btn-secondary');
    } else {
      stepDrop.style.display = 'block';
      collapseToggle.textContent = 'Collapse';
      collapseToggle.classList.remove('btn-primary');
      collapseToggle.classList.add('btn-secondary');
    }
  });
}

// ---------------------------------------------------------------------------
// Setup helpers
// ---------------------------------------------------------------------------
function setupDragDrop() {
  if (!dropZone) return;
  ['dragenter','dragover'].forEach(evt => {
    dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.add('dragover'); });
  });
  ['dragleave','drop'].forEach(evt => {
    dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.remove('dragover'); });
  });
  dropZone.addEventListener('drop', e => {
    const files = Array.from(e.dataTransfer.files).filter(f => f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf'));
    addFiles(files);
  });
}

function setupFileInput() {
  if (!fileInput) return;
  fileInput.addEventListener('change', e => {
    addFiles(Array.from(fileInput.files));
    fileInput.value = '';
  });
  if (browseLink) browseLink.addEventListener('click', e => { e.preventDefault(); fileInput.click(); });
}

function setupButtons() {
  if (btnExtract) btnExtract.addEventListener('click', onExtractClick);
  if (btnDownloadDaily) btnDownloadDaily.addEventListener('click', onDownloadDaily);
  if (btnDownloadWeekly) btnDownloadWeekly.addEventListener('click', onDownloadWeekly);
  if (btnMoreFiles) btnMoreFiles.addEventListener('click', () => {
    if (stepDrop) stepDrop.style.display = 'block';
    if (stepTabs) stepTabs.style.display = 'none';
  });
}

// ---------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------
function validateCashTpe() {
  const unfilled = pdfFiles.filter(f => (f.cashAmount || 0) === 0 && (f.tpeAmount || 0) === 0);
  if (unfilled.length > 0) {
    const names = unfilled.map(f => `Day ${f.day}/${f.month}`).join(', ');
    showToast('error', 'Cash/TPE not filled', `Please fill Cash or TPE for: ${names}`);
    return false;
  }
  return true;
}

function updateDownloadButtons() {
  const allFilled = pdfFiles.every(f => (f.cashAmount || 0) > 0 || (f.tpeAmount || 0) > 0);
  if (btnDownloadDaily) btnDownloadDaily.disabled = !allFilled || pdfFiles.length === 0;
  if (btnDownloadWeekly) btnDownloadWeekly.disabled = !allFilled || pdfFiles.length === 0;
}

// ---------------------------------------------------------------------------
// Toast notification
// ---------------------------------------------------------------------------
function showToast(severity, summary, detail) {
  const existing = document.getElementById('gamefy-toast');
  if (existing) existing.remove();

  const toast = document.createElement('div');
  toast.id = 'gamefy-toast';
  toast.className = `toast ${severity}`;
  toast.innerHTML = `
    <div class="toast-header">
      <strong>${summary}</strong>
      <button type="button" onclick="document.getElementById('gamefy-toast').remove()">×</button>
    </div>
    <div class="toast-body">${detail}</div>
  `;
  document.body.appendChild(toast);

  setTimeout(() => {
    if (toast.parentNode) toast.remove();
  }, 5000);
}

// ---------------------------------------------------------------------------
// Files
// ---------------------------------------------------------------------------
function addFiles(files) {
  console.log('addFiles called with', files.length, 'files');
  for (const f of files) {
    const parsed = parseDayMonth(f.name);
    if (!parsed) {
      alert(`Skipped "${f.name}": name must be like 15-03.pdf (day-month)`);
      continue;
    }
    if (pdfFiles.some(x => x.name === f.name && x.day === parsed.day && x.month === parsed.month)) continue;
    pdfFiles.push({
      file: f, name: f.name, day: parsed.day, month: parsed.month, year: currentYear,
      entries: [], processed: false, cashAmount: 0, tpeAmount: 0
    });
  }
  renderFileList();
  updateStats();
  updateDownloadButtons();
  if (stepDrop) stepDrop.style.display = 'block';
  if (stepTabs) stepTabs.style.display = 'none';
}

function removeFile(index) {
  console.log('removeFile called with index:', index);
  if (index < 0 || index >= pdfFiles.length) return;
  if (spreadsheetInstances[index]) {
    try { spreadsheetInstances[index].destroy(); } catch (e) {}
    spreadsheetInstances[index] = null;
  }
  pdfFiles.splice(index, 1);
  if (pdfFiles.length === 0) {
    if (stepDrop) stepDrop.style.display = 'block';
    if (stepTabs) stepTabs.style.display = 'none';
    activeTabIndex = 0;
  } else {
    if (activeTabIndex >= pdfFiles.length) activeTabIndex = pdfFiles.length - 1;
    renderTabs();
  }
  updateStats();
  updateDownloadButtons();
}

function parseDayMonth(name) {
  const m = name.match(/^(\d{1,2})[-_](\d{1,2})/);
  if (!m) return null;
  return { day: parseInt(m[1], 10), month: parseInt(m[2], 10) };
}

function renderFileList() {
  if (!fileList) return;
  fileList.innerHTML = '';
  for (let i = 0; i < pdfFiles.length; i++) {
    const f = pdfFiles[i];
    const card = document.createElement('div');
    card.className = 'file-card-item';
    const status = f.processed ? (f.entries.length ? '✓ ' + f.entries.length + ' entries' : '⚠ No table') : 'Pending';
    const statusClass = f.processed && f.entries.length ? 'status-ok' : (f.processed ? 'status-warn' : 'status-info');
    card.innerHTML = `
      <div class="file-info">
        <span class="file-name">${escapeHtml(f.name)}</span>
        <span class="file-badge ${statusClass}">${status}</span>
      </div>
      <div class="file-meta">
        <span class="meta-item"><i class="fa-solid fa-calendar"></i> Day ${f.day}</span>
        <span class="meta-item"><i class="fa-solid fa-calendar-days"></i> Month ${f.month}</span>
        <span class="meta-item"><i class="fa-solid fa-list"></i> ${f.entries.length}</span>
      </div>
      <button class="btn btn-secondary btn-sm" onclick="window.app.removeFile(${i})">Remove</button>
    `;
    fileList.appendChild(card);
  }
}

function updateStats() {
  if (!statsFiles || !statsEntries) return;
  statsFiles.textContent = pdfFiles.length;
  const totalEntries = pdfFiles.reduce((s, f) => s + f.entries.length, 0);
  statsEntries.textContent = totalEntries;
}

// ---------------------------------------------------------------------------
// Tabs
// ---------------------------------------------------------------------------
function renderTabs() {
  console.log('renderTabs called, activeTabIndex:', activeTabIndex, 'files:', pdfFiles.length);
  if (!tabsHeader || !tabsPanels) {
    console.error('tabsHeader or tabsPanels is null!');
    return;
  }
  tabsHeader.innerHTML = '';
  tabsPanels.innerHTML = '';

  pdfFiles.forEach((f, i) => {
    const isActive = i === activeTabIndex;

    const tabBtn = document.createElement('button');
    tabBtn.className = 'tab-btn' + (isActive ? ' active' : '');
    tabBtn.textContent = `${f.name} (${f.day}/${f.month})`;
    tabBtn.onclick = () => { activeTabIndex = i; renderTabs(); initSpreadsheet(i); };
    tabsHeader.appendChild(tabBtn);

    const panel = document.createElement('div');
    panel.className = 'tab-panel' + (isActive ? ' active' : '');
    if (!isActive) panel.style.display = 'none';
    panel.innerHTML = `
      <div class="file-tab-card">
        <div class="file-tab-actions">
          <button class="btn btn-secondary btn-sm" onclick="window.app.processSingle(${i})" ${isProcessing ? 'disabled' : ''}>
            <span class="spinner" style="display:${isProcessing ? 'inline-block' : 'none'};margin-right:6px;"></span>
            ${isProcessing ? 'Processing...' : '🔄 Re-process'}
          </button>
          <button class="btn btn-secondary btn-sm" onclick="window.app.addRow(${i})">➕ Add row</button>
          <button class="btn btn-secondary btn-sm" onclick="window.app.deleteLastRow(${i})">🗑️ Delete last</button>
          <span class="p-chip" style="background:${f.processed ? '#eef2f7' : '#fef3c7'}; color:${f.processed ? '#2563eb' : '#92400e'}; font-weight:700;">
            ${f.processed ? '✓ ' + f.entries.length + ' entries' : '⚠ Not processed'}
          </span>
        </div>
        <div id="spreadsheet-${i}" class="spreadsheet-container">
          <div class="spreadsheet-loading"><i class="fa-solid fa-spinner fa-spin"></i><p>Loading Excel editor...</p></div>
        </div>
        <div class="reconciliation-card">
          <h3 class="recon-title">Daily Reconciliation</h3>
          <div class="recon-grid">
            <div class="recon-item income-item">
              <div class="recon-label">Total Amounts</div>
              <div class="recon-value total" id="total-amounts-${i}">0.00</div>
            </div>
            <div class="recon-item cash-item">
              <div class="recon-label">Total Cash</div>
              <input type="number" id="cash-${i}" min="0" step="0.01" value="${f.cashAmount || 0}" class="recon-input" onchange="window.app.updateCashTpe(${i})" />
            </div>
            <div class="recon-item tpe-item">
              <div class="recon-label">Total TPE</div>
              <input type="number" id="tpe-${i}" min="0" step="0.01" value="${f.tpeAmount || 0}" class="recon-input" onchange="window.app.updateCashTpe(${i})" />
            </div>
            <div class="recon-item difference-item">
              <div class="recon-label">Difference</div>
              <div class="recon-value difference" id="diff-${i}">0.00</div>
            </div>
            <div class="recon-item save-item">
              <button class="btn btn-primary btn-sm" onclick="window.app.saveCashTpe(${i})" id="save-btn-${i}">
                <i class="fa-solid fa-save"></i> Save
              </button>
            </div>
          </div>
          <p class="recon-note">Difference = Total Amounts − (Total Cash + Total TPE)</p>
        </div>
      </div>
    `;
    tabsPanels.appendChild(panel);
  });

  setTimeout(() => initSpreadsheet(activeTabIndex), 100);
}

function initSpreadsheet(tabIndex) {
  const f = pdfFiles[tabIndex];
  if (!f) { console.error('No file at tabIndex', tabIndex); return; }

  const container = document.getElementById(`spreadsheet-${tabIndex}`);
  if (!container) { console.error('Container not found for tab', tabIndex); return; }
  container.innerHTML = '';

  if (spreadsheetInstances[tabIndex]) {
    try { spreadsheetInstances[tabIndex].destroy(); } catch (e) {}
    spreadsheetInstances[tabIndex] = null;
  }

  console.log('Creating jspreadsheet for tab', tabIndex, 'entries:', f.entries.length);

  const columns = [
    { type: 'text', title: '#', width: 50, readonly: true },
    { type: 'text', title: 'Date', width: 110 },
    { type: 'dropdown', title: 'Type', width: 100, source: ['Income', 'Expense'] },
    { type: 'text', title: 'Description', width: 200 },
    { type: 'text', title: 'Comment', width: 200 },
    { type: 'number', title: 'Amount', width: 120, align: 'right' },
    { type: 'text', title: 'Source File', width: 150, readonly: true },
  ];

  const entries = f.entries || [];

  let data;
  if (entries.length === 0) {
    data = [[
      '1',
      '',
      '',
      '',
      '',
      0,
      '',
    ]];
  } else {
    data = entries.map((e, i) => [
      String(i + 1),
      e.date || '',
      e.type.charAt(0).toUpperCase() + e.type.slice(1),
      e.description || '',
      e.comment || '',
      e.amount || 0,
      e.sourceFile || '',
    ]);
  }

  try {
    const inst = jspreadsheet(container, {
      data: data,
      columns: columns,
      search: true,
      filters: true,
      pagination: 25,
      tableOverflow: true,
      tableWidth: '100%',
      tableHeight: '420px',
      onchange: (worksheet, cell, x, y, value) => {
        handleSpreadsheetChange(tabIndex, y, x, value);
      },
    });

    spreadsheetInstances[tabIndex] = inst;
    console.log('jspreadsheet created successfully for tab', tabIndex);
    updateReconciliation(tabIndex);
  } catch (err) {
    console.error('jspreadsheet error:', err);
    container.innerHTML = '<div style="padding:20px;color:#ef4444;">Excel editor failed to load.</div>';
  }
}

function handleSpreadsheetChange(tabIndex, rowIndex, colIndex, value) {
  const f = pdfFiles[tabIndex];
  if (!f || rowIndex >= f.entries.length) return;
  const e = f.entries[rowIndex];
  const key = EXCEL_COLS[colIndex].key;
  switch (key) {
    case 'date': e.date = value || ''; break;
    case 'type': e.type = (String(value).toLowerCase() === 'income' ? 'income' : 'expense'); break;
    case 'description': e.description = value || ''; break;
    case 'comment': e.comment = value || ''; break;
    case 'amount': e.amount = parseFloat(value) || 0; break;
  }
  updateReconciliation(tabIndex);
  updateStats();
  updateDownloadButtons();
}

function updateReconciliation(tabIndex) {
  const f = pdfFiles[tabIndex];
  if (!f) return;

  const total = f.entries.reduce((s, e) => s + e.amount, 0);
  const cashEl = document.getElementById(`cash-${tabIndex}`);
  const tpeEl = document.getElementById(`tpe-${tabIndex}`);
  const cash = cashEl ? (parseFloat(cashEl.value) || 0) : 0;
  const tpe = tpeEl ? (parseFloat(tpeEl.value) || 0) : 0;
  const diff = total - (cash + tpe);

  const totalEl = document.getElementById(`total-amounts-${tabIndex}`);
  const diffEl = document.getElementById(`diff-${tabIndex}`);
  if (totalEl) totalEl.textContent = formatMoney(total);
  if (diffEl) {
    diffEl.textContent = (diff >= 0 ? '+' : '−') + formatMoney(Math.abs(diff));
    diffEl.className = 'recon-value difference ' + (diff >= 0 ? 'diff-positive' : 'diff-negative');
  }
}

window.updateCashTpe = function(tabIndex) {
  updateReconciliation(tabIndex);
  const f = pdfFiles[tabIndex];
  if (f) {
    f.cashAmount = parseFloat(document.getElementById(`cash-${tabIndex}`).value) || 0;
    f.tpeAmount = parseFloat(document.getElementById(`tpe-${tabIndex}`).value) || 0;
  }
  updateDownloadButtons();
};

window.saveCashTpe = function(tabIndex) {
  const f = pdfFiles[tabIndex];
  if (!f) return;
  f.cashAmount = parseFloat(document.getElementById(`cash-${tabIndex}`).value) || 0;
  f.tpeAmount = parseFloat(document.getElementById(`tpe-${tabIndex}`).value) || 0;
  updateReconciliation(tabIndex);
  updateDownloadButtons();
  showToast('success', 'Saved', `Cash/TPE saved for Day ${f.day}/${f.month}`);
};

// ---------------------------------------------------------------------------
// PDF processing
// ---------------------------------------------------------------------------
function onExtractClick() {
  console.log('Extract clicked, files:', pdfFiles.length);
  if (!pdfFiles.length) {
    showToast('error', 'No files', 'Please drop PDF files first.');
    return;
  }
  if (!libsReady) { alert('Please wait for pdf.js to load'); return; }
  isProcessing = true;
  if (btnExtract) { btnExtract.disabled = true; if (extractSpinner) extractSpinner.style.display = 'inline-block'; }
  processAllFiles().then(() => {
    if (stepDrop) stepDrop.style.display = 'none';
    if (stepTabs) stepTabs.style.display = 'block';
    isCollapsed = true;
    if (collapseToggle) {
      collapseToggle.textContent = 'Expand';
      collapseToggle.classList.add('btn-primary');
      collapseToggle.classList.remove('btn-secondary');
    }
    renderTabs();
  }).finally(() => {
    isProcessing = false;
    if (btnExtract) { btnExtract.disabled = false; if (extractSpinner) extractSpinner.style.display = 'none'; }
  });
}

function onDownloadDaily() {
  if (!validateCashTpe()) return;
  downloadDailyExcel();
}

function onDownloadWeekly() {
  if (isDownloading) return;
  if (!validateCashTpe()) return;
  isDownloading = true;
  downloadWeeklyExcel().then(() => { isDownloading = false; });
}

async function processAllFiles() {
  for (const f of pdfFiles) {
    await processPdf(f);
  }
  renderFileList();
  renderTabs();
}

window.processSingle = async function(tabIndex) {
  if (!libsReady) { alert('Please wait for pdf.js to load'); return; }
  const f = pdfFiles[tabIndex];
  isProcessing = true;
  try {
    await processPdf(f);
    renderFileList();
    renderTabs();
    initSpreadsheet(tabIndex);
  } catch (err) {
    console.error(err);
    alert('Processing failed: ' + err.message);
  } finally {
    isProcessing = false;
  }
};

async function processPdf(f) {
  const arrayBuffer = await f.file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

  const allText = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    content.items.sort((a, b) => {
      const ay = a.transform?.[5] || 0;
      const by = b.transform?.[5] || 0;
      if (Math.abs(ay - by) > 2) return by - ay;
      return (a.transform?.[4] || 0) - (b.transform?.[4] || 0);
    });
    allText.push(content.items.map(it => it.str).join(' '));
  }

  const fullText = allText.join('\n');
  console.log('PDF extracted text length:', fullText.length);
  console.log('PDF extracted text preview:', fullText.substring(0, 500));

  f.entries = parseTableText(fullText, f.day, f.month, f.year, f.name);
  console.log('PDF entries found:', f.entries.length);
  f.processed = true;
}

const dateRegexFallback = /(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})/;
const timeRegexFallback = /(\d{1,2}:\d{2}:\d{2})/;

function parseTableText(text, day, month, year, sourceFile) {
  const entries = [];

  const rowRegex = /(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}:\d{2}:\d{2})\s+(Income|Expense)\s+(Cash|Credit Card|Debit Card)\s+(.+?)\s+(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)(?:\s*TND)?\s*(\([^)]*\))?/gi;

  let match;
  while ((match = rowRegex.exec(text)) !== null) {
    const d = parseInt(match[1], 10);
    const m = parseInt(match[2], 10);
    const y = parseInt(match[3], 10);
    const type = match[5].toLowerCase() === 'income' ? 'income' : 'expense';
    let description = match[6].trim();
    let amountStr = match[8];
    let comment = match[9] ? match[8].replace(/^[\(\s]+|[\)\s]+$/g, '') : '';

    let cleaned = amountStr.replace(/[^0-9.,-]/g, '');
    const lastDot = cleaned.lastIndexOf('.');
    const lastComma = cleaned.lastIndexOf(',');
    const decSep = lastComma > lastDot ? ',' : '.';

    let intPart, decPart;
    if (decSep === ',') {
      const p = cleaned.split(',');
      decPart = p.pop();
      intPart = p.join('').replace(/\./g, '');
    } else {
      const p = cleaned.split('.');
      decPart = p.pop();
      intPart = p.join('').replace(/,/g, '');
    }

    const num = parseFloat((intPart + '.' + decPart).replace(/[^0-9.-]/g, ''));
    if (!isFinite(num)) continue;

    const amount = type === 'expense' ? -Math.abs(num) : Math.abs(num);

    description = description.replace(/\s+/g, ' ').trim();
    if (!description || description.length < 2) description = 'Entry';

    let entryDate = `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    if (d > 0 && d <= 31 && m > 0 && m <= 12) {
      entryDate = `${String(y).padStart(4, '0')}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    }

    entries.push({
      id: crypto.randomUUID?.() ?? `${day}-${month}-${Math.random()}`,
      date: entryDate,
      type, description, comment, amount,
    });
  }

  const seen = new Set();
  return entries.filter(e => {
    const k = `${e.date}|${e.type}|${e.description}|${e.amount}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

// ---------------------------------------------------------------------------
// Grid actions
// ---------------------------------------------------------------------------
window.addRow = function(tabIndex) {
  const f = pdfFiles[tabIndex];
  if (!f) return;
  const d = `${f.year}-${String(f.month).padStart(2, '0')}-${String(f.day).padStart(2, '0')}`;
  f.entries.push({
    id: crypto.randomUUID?.() ?? Math.random().toString(),
    date: d, type: 'expense', description: '', comment: '', amount: 0, sourceFile: f.name,
  });
  initSpreadsheet(tabIndex);
  updateStats();
};

window.deleteLastRow = function(tabIndex) {
  const f = pdfFiles[tabIndex];
  if (!f || f.entries.length === 0) return;
  f.entries.pop();
  initSpreadsheet(tabIndex);
  updateStats();
};

// ---------------------------------------------------------------------------
// Week grouping
// ---------------------------------------------------------------------------
function weekOfMonth(day) {
  return Math.ceil(day / 7);
}

function getAllEntries() {
  const all = [];
  for (const f of pdfFiles) {
    for (const e of f.entries) {
      all.push({ ...e, sourceFile: f.name });
    }
  }
  return all;
}

function groupByWeek() {
  const weeks = new Map();
  for (const e of getAllEntries()) {
    const d = new Date(e.date);
    if (isNaN(d.getTime())) continue;
    const wk = weekOfMonth(d.getDate());
    if (!weeks.has(wk)) weeks.set(wk, []);
    weeks.get(wk).push(e);
  }
  return weeks;
}

function categorize(description) {
  const desc = description.toLowerCase();
  for (const cat of CATEGORIES) {
    for (const kw of cat.keywords) {
      if (desc.includes(kw)) return cat.label;
    }
  }
  return 'Other';
}

// ---------------------------------------------------------------------------
// Reconciliation helpers
// ---------------------------------------------------------------------------
function totalAmounts(f) {
  return f.entries.reduce((s, e) => s + e.amount, 0);
}
function totalIncome(f) {
  return f.entries.filter(e => e.type === 'income').reduce((s, e) => s + Math.abs(e.amount), 0);
}
function totalExpense(f) {
  return f.entries.filter(e => e.type === 'expense').reduce((s, e) => s + Math.abs(e.amount), 0);
}
function difference(f) {
  const total = totalAmounts(f);
  const cash = f.cashAmount || 0;
  const tpe = f.tpeAmount || 0;
  return (cash + tpe) - total;
}

// ---------------------------------------------------------------------------
// Excel export: Daily (one sheet per day/file)
// ---------------------------------------------------------------------------
function downloadDailyExcel() {
  if (!pdfFiles.length) { alert('Nothing to export'); return; }

  const wb = XLSX.utils.book_new();

  for (const f of pdfFiles) {
    const rows = [['#', 'Date', 'Type', 'Description', 'Comment', 'Amount', 'Source File']];
    f.entries.forEach((e, i) => {
      rows.push([
        i + 1,
        e.date,
        e.type.charAt(0).toUpperCase() + e.type.slice(1),
        e.description,
        e.comment || '',
        e.type === 'income' ? +e.amount : -e.amount,
        e.sourceFile,
      ]);
    });
    rows.push(['', '', '', 'Total Amounts', '', totalAmounts(f), '']);
    rows.push(['', '', '', 'Total Cash', '', f.cashAmount || 0, '']);
    rows.push(['', '', '', 'Total TPE', '', f.tpeAmount || 0, '']);
    rows.push(['', '', '', 'Difference', '', difference(f), '']);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [
      { width: 5 }, { width: 12 }, { width: 10 }, { width: 38 }, { width: 20 }, { width: 14 }, { width: 15 },
    ];
    XLSX.utils.book_append_sheet(wb, ws, `Day ${f.day}-${f.month}`);
  }

  const fname = `Gamefy_Daily_Invoices_Month${currentYear}.xlsx`;
  XLSX.writeFile(wb, fname);
  alert('Downloaded: ' + fname);
}

// ---------------------------------------------------------------------------
// Excel export: Weekly (one sheet per week with all rows + accounting summary below)
// ---------------------------------------------------------------------------
function downloadWeeklyExcel() {
  if (!pdfFiles.length) { alert('Nothing to export'); return; }

  const weeks = groupByWeek();
  if (weeks.size === 0) { alert('No data to export'); return; }

  const wb = XLSX.utils.book_new();

  const weekNums = Array.from(weeks.keys()).sort((a, b) => a - b);
  const addedSheets = new Set();

  for (const wk of weekNums) {
    const sheetName = `Week ${wk}`;
    if (addedSheets.has(sheetName)) continue;
    addedSheets.add(sheetName);
    const entries = weeks.get(wk);
    const startDay = (wk - 1) * 7 + 1;
    const endDay = Math.min(wk * 7, 31);
    const weekTitle = `Vente ${wk}ème semaine`;

    const catTotals = {};
    for (const cat of CATEGORIES) catTotals[cat.label] = 0;
    catTotals['Other'] = 0;

    let totalIncome = 0;
    let totalExpense = 0;

    for (const e of entries) {
      if (e.type === 'income') totalIncome += Math.abs(e.amount);
      else totalExpense += Math.abs(e.amount);

      if (e.type === 'income') {
        const cat = categorize(e.description);
        if (catTotals[cat] !== undefined) {
          catTotals[cat] += Math.abs(e.amount);
        } else {
          catTotals['Other'] += Math.abs(e.amount);
        }
      }
    }

    const totalHT = totalIncome - totalExpense;

    const weekSourceFiles = new Set(entries.map(e => e.sourceFile));
    let weekCash = 0;
    let weekTpe = 0;
    for (const f of pdfFiles) {
      if (weekSourceFiles.has(f.name)) {
        weekCash += f.cashAmount || 0;
        weekTpe += f.tpeAmount || 0;
      }
    }

    const rows = [];

    rows.push([weekTitle, '', '', '', '', '', '']);
    rows.push(['Days ' + startDay + ' to ' + endDay, '', '', '', '', '', '']);
    rows.push(['', '', '', '', '', '', '']);

    rows.push(['#', 'Date', 'Type', 'Description', 'Comment', 'Amount', 'Source File']);

    let rowNum = 1;
    for (const e of entries) {
      rows.push([
        rowNum++,
        e.date,
        e.type.charAt(0).toUpperCase() + e.type.slice(1),
        e.description,
        e.comment || '',
        e.type === 'income' ? Math.abs(e.amount) : -Math.abs(e.amount),
        e.sourceFile,
      ]);
    }

    rows.push(['', '', '', '', '', '', '']);

    rows.push(['Désignation', 'Montant', '', '', '', '', '']);
    rows.push([weekTitle, '', '', '', '', '', '']);
    rows.push(['', '', '', '', '', '', '']);

    for (const cat of CATEGORIES) {
      const val = catTotals[cat.label];
      rows.push([cat.label, val === 0 ? '—' : val, '', '', '', '', '']);
    }
    rows.push(['Other', catTotals['Other'] === 0 ? '—' : catTotals['Other'], '', '', '', '', '']);

    const weekNetTotal = totalIncome - totalExpense;
    rows.push(['Total Amounts', weekNetTotal, '', '', '', '', '']);
    rows.push(['Total Cash', weekCash, '', '', '', '', '']);
    rows.push(['Total TPE', weekTpe, '', '', '', '', '']);
    rows.push(['Difference', weekCash + weekTpe - weekNetTotal, '', '', '', '', '']);

    rows.push(['TOTAL HT', totalHT, '', '', '', '', '']);
    rows.push(['TVA 19%', totalHT * 0.19, '', '', '', '', '']);
    rows.push(['Droit de timbre', 1.000, '', '', '', '', '']);
    rows.push(['Total TTC', totalHT + totalHT * 0.19 + 1.000, '', '', '', '', '']);

    const ws = XLSX.utils.aoa_to_sheet(rows);

    ws['!cols'] = [
      { width: 8 },
      { width: 12 },
      { width: 12 },
      { width: 40 },
      { width: 20 },
      { width: 14 },
      { width: 20 },
    ];

    XLSX.utils.book_append_sheet(wb, ws, `Week ${wk}`);
  }

  const fname = `Gamefy_Weekly_Invoices_Month${currentYear}.xlsx`;
  XLSX.writeFile(wb, fname);
  alert('Downloaded: ' + fname);
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function escapeHtml(t) {
  if (t == null) return '';
  return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function formatMoney(n) {
  const v = Number(n) || 0;
  return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

window.app = {
  removeFile, processSingle, addRow, deleteLastRow, updateCashTpe, saveCashTpe,
};

libsReady = true;
if (btnExtract) btnExtract.disabled = false;