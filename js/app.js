'use strict';

// ── State ──────────────────────────────────────────────────
const state = {
  workbook: null,
  sheets: {},           // { sheetName: [{...row}] }
  activeSheet: null,
  columns: [],
  allRows: [],
  filteredRows: [],
  searchQuery: '',
  filters: [],          // [{col, val}]
  sortCol: null,
  sortDir: 'asc',
  viewMode: 'table',    // 'table' | 'grid'
  page: 1,
  pageSize: 50,
  selectedRows: new Set(),  // Track selected row indices
};

// ── DOM ────────────────────────────────────────────────────
const $  = id => document.getElementById(id);
const $$ = sel => document.querySelectorAll(sel);

const uploadZone    = $('upload-zone');
const dataView      = $('data-view');
const fileInput     = $('file-input');
const searchInput   = $('search-input');
const filterBar     = $('filter-bar');
const filterSelect  = $('filter-select');
const filterValue   = $('filter-value');
const filterDropdown= $('filter-dropdown');
const activeFilters = $('active-filters');
const sheetTabs     = $('sheet-tabs');
const tableContainer = $('table-container');
const gridContainer = $('grid-container');
const paginationEl  = $('pagination');
const statsEl       = $('stats');
const loadingOverlay = $('loading-overlay');
const toast         = $('toast');
const installBanner = $('install-banner');

// ── PWA Install ────────────────────────────────────────────
let deferredPrompt = null;

window.addEventListener('beforeinstallprompt', e => {
  e.preventDefault();
  deferredPrompt = e;
  installBanner.classList.add('show');
});

$('install-btn').addEventListener('click', async () => {
  if (!deferredPrompt) return;
  deferredPrompt.prompt();
  const { outcome } = await deferredPrompt.userChoice;
  if (outcome === 'accepted') showToast('✅ App installed!');
  deferredPrompt = null;
  installBanner.classList.remove('show');
});

$('dismiss-install').addEventListener('click', () => installBanner.classList.remove('show'));

// ── Service Worker ─────────────────────────────────────────
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}

// ── File Handling ──────────────────────────────────────────
const dropTarget = document.querySelector('.drop-target');

dropTarget.addEventListener('dragover', e => {
  e.preventDefault();
  dropTarget.classList.add('drag-over');
});
dropTarget.addEventListener('dragleave', () => dropTarget.classList.remove('drag-over'));
dropTarget.addEventListener('drop', e => {
  e.preventDefault();
  dropTarget.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) loadFile(file);
});

dropTarget.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', e => {
  if (e.target.files[0]) loadFile(e.target.files[0]);
});

$('new-file-btn').addEventListener('click', () => {
  uploadZone.style.display = 'flex';
  dataView.style.display = 'none';
  $('new-file-btn').style.display = 'none';
  state.workbook = null;
  fileInput.value = '';
});

$('demo-btn').addEventListener('click', loadDemoData);

async function loadFile(file) {
  const name = file.name.toLowerCase();
  if (!name.endsWith('.xlsx') && !name.endsWith('.xls') && !name.endsWith('.csv') && !name.endsWith('.ods')) {
    showToast('⚠️ Please upload .xlsx, .xls, .csv, or .ods file');
    return;
  }
  showLoading(true);
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    processWorkbook(wb);
    showToast('✅ ' + file.name + ' loaded');
  } catch(err) {
    showToast('❌ Error reading file: ' + err.message);
    console.error(err);
  } finally {
    showLoading(false);
  }
}

function processWorkbook(wb) {
  state.workbook = wb;
  state.sheets = {};

  wb.SheetNames.forEach(name => {
    const ws = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false, dateNF: 'yyyy-mm-dd' });
    state.sheets[name] = rows;
  });

  buildSheetTabs();
  switchSheet(wb.SheetNames[0]);

  uploadZone.style.display = 'none';
  dataView.style.display = 'flex';
  $('new-file-btn').style.display = 'flex';
}

// ── Demo Data ──────────────────────────────────────────────
function loadDemoData() {
  const departments = ['Engineering', 'Marketing', 'Sales', 'Design', 'HR', 'Finance'];
  const statuses = ['Active', 'On Leave', 'Remote', 'Contractor'];
  const cities = ['New York', 'San Francisco', 'Austin', 'Chicago', 'Seattle', 'Boston'];
  const names = ['Alice Chen','Bob Kumar','Carol Smith','David Park','Eva Rossi','Frank Lima','Grace Obi','Hiro Tanaka','Iris Müller','Jay Patel'];

  const rows = Array.from({ length: 120 }, (_, i) => ({
    ID: String(i + 1).padStart(4,'0'),
    Name: names[i % names.length] + (i >= names.length ? ` ${Math.floor(i/names.length)+1}` : ''),
    Department: departments[i % departments.length],
    City: cities[i % cities.length],
    Status: statuses[i % statuses.length],
    Salary: (45000 + (i * 1337) % 80000).toLocaleString(),
    'Start Date': `202${Math.floor(i/20)}-${String((i%12)+1).padStart(2,'0')}-${String((i%28)+1).padStart(2,'0')}`,
    Projects: String(1 + i % 8),
    Score: ((60 + (i * 7) % 40) / 10).toFixed(1),
    Email: `${names[i%names.length].split(' ')[0].toLowerCase()}${i+1}@company.com`,
  }));

  const wb = { SheetNames: ['Employees', 'Summary'], Sheets: {} };
  wb.Sheets['Employees'] = XLSX.utils.json_to_sheet(rows);
  const summary = departments.map(d => ({
    Department: d,
    'Headcount': rows.filter(r => r.Department === d).length,
    'Avg Score': (rows.filter(r => r.Department === d).reduce((a, r) => a + parseFloat(r.Score), 0) / rows.filter(r => r.Department === d).length).toFixed(1),
  }));
  wb.Sheets['Summary'] = XLSX.utils.json_to_sheet(summary);
  processWorkbook(wb);
  showToast('📊 Demo data loaded — 120 employees, 2 sheets');
}

// ── Sheet Tabs ─────────────────────────────────────────────
function buildSheetTabs() {
  sheetTabs.innerHTML = '';
  Object.keys(state.sheets).forEach(name => {
    const tab = document.createElement('button');
    tab.className = 'sheet-tab';
    tab.textContent = name;
    tab.addEventListener('click', () => switchSheet(name));
    sheetTabs.appendChild(tab);
  });
}

function switchSheet(name) {
  state.activeSheet = name;
  state.allRows = state.sheets[name] || [];
  state.columns = state.allRows.length ? Object.keys(state.allRows[0]) : [];
  state.searchQuery = '';
  state.filters = [];
  state.sortCol = null;
  state.sortDir = 'asc';
  state.page = 1;
  state.selectedRows.clear();
  searchInput.value = '';

  $$('.sheet-tab').forEach(t => t.classList.toggle('active', t.textContent === name));
  buildFilterSelect();
  applyFiltersAndRender();
}

// ── Search ─────────────────────────────────────────────────
searchInput.addEventListener('input', e => {
  state.searchQuery = e.target.value.trim().toLowerCase();
  state.page = 1;
  applyFiltersAndRender();
});

// ── Filter Panel ───────────────────────────────────────────
$('filter-toggle').addEventListener('click', () => {
  filterBar.classList.toggle('open');
});

function buildFilterSelect() {
  filterSelect.innerHTML = '<option value="">— column —</option>';
  state.columns.forEach(col => {
    const opt = document.createElement('option');
    opt.value = col;
    opt.textContent = col;
    filterSelect.appendChild(opt);
  });
  updateFilterInput();
}

function updateFilterInput() {
  const column = filterSelect.value;
  const dropdownColumns = ['product category', 'speciality'];
  const isDropdown = column && dropdownColumns.includes(column.toLowerCase());
  const isMultiselect = column && column.toLowerCase() === 'speciality';

  filterValue.style.display = isDropdown ? 'none' : 'inline-block';
  filterDropdown.style.display = isDropdown ? 'inline-block' : 'none';

  if (isMultiselect) {
    filterDropdown.setAttribute('multiple', 'multiple');
    filterDropdown.setAttribute('size', '5');
  } else {
    filterDropdown.removeAttribute('multiple');
    filterDropdown.setAttribute('size', '1');
  }

  filterDropdown.innerHTML = '';

  if (isDropdown) {
    let values;
    if (column.toLowerCase() === 'speciality') {
      values = [
        'Dental', 'OBS', 'Ophthalmology', 'Dermatology', 'GP', 'Cardiology',
        'Orthopaedic', 'Physiotherapy', 'Diabetologist', 'Paediatrician',
        'Neurology', 'Radiology', 'Homeopathy', 'Not', 'Nephrology',
        'Gastroenterology', 'Dietetics', 'ENT', 'PSYCHIATRY', 'Oncology',
        'Pathology', 'Urology', 'Rheumatology', 'Pulmonology', 'General',
        'Vascular', 'Interventional', 'Critical', 'Proctology', 'Plastic'
      ];
    } else {
      values = Array.from(new Set(state.allRows.map(r => String(r[column] ?? '').trim()).filter(Boolean))).sort();
    }

    values.forEach(val => {
      const opt = document.createElement('option');
      opt.value = val;
      opt.textContent = val;
      filterDropdown.appendChild(opt);
    });
  }
}

function getFilterValue() {
  if (filterSelect.value && filterDropdown.style.display === 'inline-block') {
    if (filterDropdown.hasAttribute('multiple')) {
      const selected = Array.from(filterDropdown.selectedOptions).map(opt => opt.value);
      return selected.length > 0 ? selected : null;
    } else {
      return filterDropdown.value;
    }
  }
  return filterValue.value;
}

$('add-filter-btn').addEventListener('click', () => {
  const col = filterSelect.value;
  const val = getFilterValue();
  if (!col || (!val || (Array.isArray(val) && val.length === 0))) {
    showToast('⚠️ Select a column and enter a value');
    return;
  }

  if (Array.isArray(val)) {
    // Handle multiselect speciality
    state.filters.push({ col, val });
  } else {
    state.filters.push({ col, val: val.trim().toLowerCase() });
  }

  filterValue.value = '';
  // Clear multiselect
  Array.from(filterDropdown.options).forEach(opt => opt.selected = false);
  state.page = 1;
  applyFiltersAndRender();
  renderActiveFilters();
});

filterSelect.addEventListener('change', updateFilterInput);

filterValue.addEventListener('keydown', e => {
  if (e.key === 'Enter') $('add-filter-btn').click();
});

$('clear-filters-btn').addEventListener('click', () => {
  state.filters = [];
  state.page = 1;
  applyFiltersAndRender();
  renderActiveFilters();
});

function renderActiveFilters() {
  activeFilters.innerHTML = '';
  state.filters.forEach((f, i) => {
    const chip = document.createElement('div');
    chip.className = 'filter-chip';
    const displayVal = Array.isArray(f.val) ? f.val.join(', ') : f.val;
    chip.innerHTML = `<span>${f.col}: "${displayVal}"</span><button onclick="removeFilter(${i})">×</button>`;
    activeFilters.appendChild(chip);
  });
}

window.removeFilter = function(i) {
  state.filters.splice(i, 1);
  state.page = 1;
  applyFiltersAndRender();
  renderActiveFilters();
};

// ── Sort ───────────────────────────────────────────────────
function handleSort(col) {
  if (state.sortCol === col) {
    state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    state.sortCol = col;
    state.sortDir = 'asc';
  }
  applyFiltersAndRender();
}

// ── View Mode ──────────────────────────────────────────────
$('view-table').addEventListener('click', () => { state.viewMode = 'table'; updateViewBtns(); renderCurrentPage(); });
$('view-grid').addEventListener('click', () => { state.viewMode = 'grid'; updateViewBtns(); renderCurrentPage(); });

function updateViewBtns() {
  $('view-table').classList.toggle('active', state.viewMode === 'table');
  $('view-grid').classList.toggle('active', state.viewMode === 'grid');
}

// ── Export ─────────────────────────────────────────────────
$('export-btn').addEventListener('click', () => {
  if (!state.filteredRows.length) { showToast('No data to export'); return; }
  const ws = XLSX.utils.json_to_sheet(state.filteredRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, state.activeSheet || 'Data');
  XLSX.writeFile(wb, `export_${Date.now()}.xlsx`);
  showToast('📥 Exported ' + state.filteredRows.length + ' rows');
});

$('share-btn').addEventListener('click', shareSelected);

// ── Core: Filter + Sort + Render ──────────────────────────
function applyFiltersAndRender() {
  let rows = state.allRows;

  // Global search
  if (state.searchQuery) {
    const q = state.searchQuery;
    rows = rows.filter(row =>
      Object.values(row).some(v => String(v).toLowerCase().includes(q))
    );
  }

  // Column filters
  state.filters.forEach(({ col, val }) => {
    if (Array.isArray(val)) {
      // Handle multiselect (e.g., speciality)
      rows = rows.filter(row => {
        const rowVal = String(row[col] ?? '').toLowerCase();
        return val.some(v => rowVal.includes(v.toLowerCase()));
      });
    } else {
      rows = rows.filter(row => String(row[col] ?? '').toLowerCase().includes(val));
    }
  });

  // Sort
  if (state.sortCol) {
    const col = state.sortCol;
    const dir = state.sortDir === 'asc' ? 1 : -1;
    rows = [...rows].sort((a, b) => {
      const av = a[col] ?? '', bv = b[col] ?? '';
      const an = parseFloat(String(av).replace(/,/g, '')), bn = parseFloat(String(bv).replace(/,/g, ''));
      if (!isNaN(an) && !isNaN(bn)) return (an - bn) * dir;
      return String(av).localeCompare(String(bv)) * dir;
    });
  }

  state.filteredRows = rows;
  updateStats();
  renderCurrentPage();
  renderPagination();
}

function updateStats() {
  const total = state.allRows.length;
  const shown = state.filteredRows.length;
  statsEl.innerHTML = `<strong>${shown.toLocaleString()}</strong> / ${total.toLocaleString()} rows`;
}

// ── Pagination ─────────────────────────────────────────────
function renderPagination() {
  const total = state.filteredRows.length;
  const pages = Math.ceil(total / state.pageSize);
  paginationEl.innerHTML = '';
  if (pages <= 1) return;

  const prev = document.createElement('button');
  prev.className = 'page-btn';
  prev.textContent = '← Prev';
  prev.disabled = state.page <= 1;
  prev.onclick = () => { state.page--; renderCurrentPage(); renderPagination(); };
  paginationEl.appendChild(prev);

  const range = pageRange(state.page, pages);
  range.forEach(p => {
    if (p === '…') {
      const el = document.createElement('span');
      el.className = 'page-btn';
      el.textContent = '…';
      paginationEl.appendChild(el);
    } else {
      const btn = document.createElement('button');
      btn.className = 'page-btn' + (p === state.page ? ' active' : '');
      btn.textContent = p;
      btn.onclick = () => { state.page = p; renderCurrentPage(); renderPagination(); };
      paginationEl.appendChild(btn);
    }
  });

  const next = document.createElement('button');
  next.className = 'page-btn';
  next.textContent = 'Next →';
  next.disabled = state.page >= pages;
  next.onclick = () => { state.page++; renderCurrentPage(); renderPagination(); };
  paginationEl.appendChild(next);

  const info = document.createElement('span');
  info.className = 'page-info';
  const from = (state.page - 1) * state.pageSize + 1;
  const to = Math.min(state.page * state.pageSize, total);
  info.textContent = `${from}–${to} of ${total}`;
  paginationEl.appendChild(info);
}

function pageRange(cur, total) {
  if (total <= 7) return Array.from({ length: total }, (_, i) => i + 1);
  if (cur <= 4) return [1,2,3,4,5,'…',total];
  if (cur >= total - 3) return [1,'…',total-4,total-3,total-2,total-1,total];
  return [1,'…',cur-1,cur,cur+1,'…',total];
}

function renderCurrentPage() {
  const start = (state.page - 1) * state.pageSize;
  const rows = state.filteredRows.slice(start, start + state.pageSize);

  if (state.viewMode === 'table') {
    tableContainer.classList.add('active');
    gridContainer.classList.remove('active');
    renderTable(rows);
    updateSelectedUI();
  } else {
    tableContainer.classList.remove('active');
    gridContainer.classList.add('active');
    renderGrid(rows);
  }
}

// ── Table Render ───────────────────────────────────────────
function renderTable(rows) {
  if (!rows.length) {
    tableContainer.innerHTML = `<div class="empty-state">
      <div class="empty-icon">🔍</div>
      <h3>No matching rows</h3>
      <p>Try adjusting your search or filters</p>
    </div>`;
    return;
  }

  const q = state.searchQuery;
  const start = (state.page - 1) * state.pageSize;

  const thead = `<thead><tr><th class="checkbox-col"><input type="checkbox" id="select-all-rows" title="Select all rows"/></th>${state.columns.map(col => {
    const cls = state.sortCol === col ? (state.sortDir === 'asc' ? ' sort-asc' : ' sort-desc') : '';
    return `<th class="${cls}" data-col="${escHtml(col)}">${escHtml(col)}</th>`;
  }).join('')}</tr></thead>`;

  const tbody = '<tbody>' + rows.map((row, idx) => {
    const rowIdx = start + idx;
    const isSelected = state.selectedRows.has(rowIdx);
    const cells = state.columns.map(col => {
      let val = row[col] ?? '';
      const raw = String(val);
      const isNum = !isNaN(parseFloat(raw.replace(/,/g,''))) && raw.trim() !== '';
      const isDate = /^\d{4}-\d{2}-\d{2}/.test(raw);
      let cls = '';
      if (isNum) cls = 'cell-number';
      else if (isDate) cls = 'cell-date';
      else if (!raw.trim()) cls = 'cell-empty';

      let display = raw.trim() || '<span class="cell-empty">—</span>';
      if (q && raw.toLowerCase().includes(q)) {
        const re = new RegExp('(' + escRegex(q) + ')', 'gi');
        display = escHtml(raw).replace(re, '<mark>$1</mark>');
      } else {
        display = escHtml(raw) || '<span class="cell-empty">—</span>';
      }
      return `<td class="${cls}" title="${escHtml(raw)}">${display}</td>`;
    }).join('');
    return `<tr class="${isSelected ? 'row-selected' : ''}" data-row-idx="${rowIdx}"><td class="checkbox-col"><input type="checkbox" class="row-checkbox" ${isSelected ? 'checked' : ''} /></td>${cells}</tr>`;
  }).join('') + '</tbody>';

  tableContainer.innerHTML = `<table>${thead}${tbody}</table>`;

  // Checkbox handlers
  const selectAllCheckbox = tableContainer.querySelector('#select-all-rows');
  const rowCheckboxes = tableContainer.querySelectorAll('.row-checkbox');
  
  if (selectAllCheckbox) {
    selectAllCheckbox.addEventListener('change', () => {
      rowCheckboxes.forEach((cb, i) => {
        const rowIdx = start + i;
        if (selectAllCheckbox.checked) {
          state.selectedRows.add(rowIdx);
        } else {
          state.selectedRows.delete(rowIdx);
        }
        cb.checked = selectAllCheckbox.checked;
      });
      updateSelectedUI();
    });
  }
  
  rowCheckboxes.forEach((cb, i) => {
    const rowIdx = start + i;
    cb.addEventListener('change', () => {
      if (cb.checked) {
        state.selectedRows.add(rowIdx);
      } else {
        state.selectedRows.delete(rowIdx);
      }
      if (selectAllCheckbox) {
        selectAllCheckbox.checked = rowCheckboxes.length > 0 && Array.from(rowCheckboxes).every(c => c.checked);
      }
      updateSelectedUI();
    });
    cb.closest('tr').addEventListener('click', (e) => {
      if (e.target !== cb) {
        cb.checked = !cb.checked;
        cb.dispatchEvent(new Event('change'));
      }
    });
  });

  // Sort click handlers
  tableContainer.querySelectorAll('th[data-col]').forEach(th => {
    th.addEventListener('click', () => handleSort(th.dataset.col));
  });
}

// ── Grid Render ────────────────────────────────────────────
function renderGrid(rows) {
  gridContainer.innerHTML = '';
  if (!rows.length) {
    gridContainer.innerHTML = `<div class="empty-state" style="grid-column:1/-1">
      <div class="empty-icon">🔍</div>
      <h3>No matching rows</h3>
    </div>`;
    return;
  }
  const frag = document.createDocumentFragment();
  rows.forEach(row => {
    const card = document.createElement('div');
    card.className = 'data-card';
    card.innerHTML = state.columns.slice(0, 10).map(col => {
      const val = String(row[col] ?? '').trim() || '—';
      return `<div class="card-field">
        <span class="card-key">${escHtml(col)}</span>
        <span class="card-val" title="${escHtml(val)}">${escHtml(val)}</span>
      </div>`;
    }).join('');
    frag.appendChild(card);
  });
  gridContainer.appendChild(frag);
}

// ── Helpers ────────────────────────────────────────────────
function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function escRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
}

function showToast(msg, dur = 2800) {
  toast.textContent = msg;
  toast.classList.add('show');
  setTimeout(() => toast.classList.remove('show'), dur);
}

function showLoading(show) {
  loadingOverlay.classList.toggle('show', show);
}

function updateSelectedUI() {
  const shareBtn = $('share-btn');
  if (!shareBtn) return;
  if (state.selectedRows.size > 0) {
    shareBtn.style.display = 'flex';
    shareBtn.title = `Share ${state.selectedRows.size} selected row${state.selectedRows.size === 1 ? '' : 's'}`;
  } else {
    shareBtn.style.display = 'none';
  }
}

function shareSelected() {
  if (state.selectedRows.size === 0) {
    showToast('⚠️ Select rows to share');
    return;
  }
  
  const selectedIndexes = Array.from(state.selectedRows).sort((a, b) => a - b);
  const fields = ['Deal Name', 'Product Category', 'Speciality', 'Address'];

  const selectedData = selectedIndexes
    .map(idx => state.filteredRows[idx])
    .filter(Boolean);

  const text = selectedData.map((row, index) => {
    const values = fields.map(field => {
      let raw = row[field] ?? row[field.toLowerCase()] ?? row[field.toUpperCase()] ?? '';
      return String(raw).trim() || '—';
    }).join('\n');
    return `${index + 1}.\n${values}`;
  }).join('\n\n');
  
  if (navigator.share) {
    navigator.share({
      title: `EREMEDIUM - ${selectedData.length} items`,
      text: text,
    }).catch(() => copyToClipboard(text));
  } else {
    copyToClipboard(text);
  }
}

function copyToClipboard(text) {
  navigator.clipboard.writeText(text).then(() => {
    showToast(`✅ Copied ${state.selectedRows.size} row(s) to clipboard`);
  }).catch(() => {
    showToast('❌ Failed to copy');
  });
}
window.addEventListener('load', async () => {
  try {
    console.log("Auto loading Excel...");

    const res = await fetch('./Data.xlsx');
    const buf = await res.arrayBuffer();

    const wb = XLSX.read(buf, { type: 'array', cellDates: true });

    // ✅ THIS is the key (your app logic)
    processWorkbook(wb);

    showToast('✅ Data.xlsx auto loaded');
  } catch (err) {
    console.error("Auto load failed:", err);
    showToast('❌ Failed to auto load Excel');
  }
});
