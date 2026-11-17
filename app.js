/**
 * Executive Inventory & Asset Dashboard (PWS Analytics)
 * Features:
 * 1. Auto-loads ExportPWS.xlsx on page load into the Dashboard view directly
 * 2. Prominent "อัปเดตไฟล์ ExportPWS.xlsx" button in Header to replace data from computer
 * 3. Persists updated file in IndexedDB so custom updates remain after refresh
 * 4. Full Drag & Drop overlay across entire window
 * 5. Executive Filters, 4 KPI Cards, 22 Yellow Columns Table with high-performance pagination
 */

document.addEventListener('DOMContentLoaded', () => {
  // 22 Yellow Columns from ExportPWS.xlsx
  const YELLOW_COLUMNS = [
    'PROVINCE', 'LOCATION_ID', 'LOCATION_CODE', 'LOCATION_NAME', 'LOCATION_TYPE',
    'GROUP_ARTICLE', 'ARTICLE_ID', 'ARTICLE_NO', 'ARTICLE_TYPE', 'ARTICLE_DESC',
    'BRAND', 'MODEL', 'DOC_ID', 'UNIT_COUNT', 'SITECODE_OSPID', 'QTY',
    'RESERVED_QTY', 'RESERVED_RTF_QTY', 'RESERVED_QUOTA', 'CREATED_DATE', 'OID', 'TITLE'
  ];

  const DB_NAME = 'PWS_DASHBOARD_DB';
  const STORE_NAME = 'pws_file_store';

  // DOM Elements
  const fileInput = document.getElementById('file-input');
  const updateFileBtn = document.getElementById('update-file-btn');
  const fileStatusBadge = document.getElementById('file-status-badge');
  const activeFileName = document.getElementById('active-filename');
  const fileCustomTag = document.getElementById('file-custom-tag');

  const errorAlert = document.getElementById('error-alert');
  const errorMessage = document.getElementById('error-message');
  const errorActionBtn = document.getElementById('error-action-btn');
  const loadingSpinner = document.getElementById('loading-spinner');
  const loadingText = document.getElementById('loading-text');
  const dashboardContent = document.getElementById('dashboard-content');
  const dragDropOverlay = document.getElementById('drag-drop-overlay');
  const toastContainer = document.getElementById('toast-container');

  // Filters
  const filterLocationName = document.getElementById('filter-location-name');
  const filterLocationType = document.getElementById('filter-location-type');
  const filterArticleDesc = document.getElementById('filter-article-desc');
  const clearSearchBtn = document.getElementById('clear-search-btn');
  const resetFiltersBtn = document.getElementById('reset-filters-btn');
  const filterMatchingCount = document.getElementById('filter-matching-count');
  const locationNameCount = document.getElementById('location-name-count');
  const locationTypeCount = document.getElementById('location-type-count');

  // KPI Elements
  const kpiTotalQty = document.getElementById('kpi-total-qty');
  const kpiReservedQty = document.getElementById('kpi-reserved-qty');
  const kpiReservedPercent = document.getElementById('kpi-reserved-percent');
  const kpiTotalRecords = document.getElementById('kpi-total-records');
  const kpiRecordsSub = document.getElementById('kpi-records-sub');
  const kpiActiveLocations = document.getElementById('kpi-active-locations');
  const kpiArticlesCount = document.getElementById('kpi-articles-count');

  // Table & Pagination
  const tableHead = document.getElementById('executive-table-head');
  const tableBody = document.getElementById('executive-table-body');
  const tableRecordCountBadge = document.getElementById('table-record-count-badge');
  const yellowChipsList = document.getElementById('yellow-chips-list');
  const rowsPerPageSelect = document.getElementById('rows-per-page-select');
  const exportCsvBtn = document.getElementById('export-csv-btn');

  const paginationInfo = document.getElementById('pagination-info');
  const pageIndicator = document.getElementById('page-indicator');
  const btnFirstPage = document.getElementById('btn-first-page');
  const btnPrevPage = document.getElementById('btn-prev-page');
  const btnNextPage = document.getElementById('btn-next-page');
  const btnLastPage = document.getElementById('btn-last-page');

  // State
  let allRows = [];                  // All data rows
  let filteredRows = [];             // Filtered rows
  let activeDisplayCols = [];        // 22 yellow column definitions
  let colIndexMap = {};              // Map: colName -> columnIndex
  let currentPage = 1;
  let rowsPerPage = 25;
  let searchDebounceTimer = null;
  let dragCounter = 0;

  // Initialize Yellow Chips Bar
  renderYellowChips();

  // ==========================================
  // Auto-load on startup
  // ==========================================
  initDashboard();

  async function initDashboard() {
    showLoading(true, 'กำลังโหลดข้อมูล ExportPWS.xlsx เข้าสู่แดชบอร์ด...');

    try {
      // 1. Check if user previously saved a custom file in IndexedDB
      const savedData = await loadFromIndexedDB();
      if (savedData && savedData.buffer && savedData.name) {
        parseWorkbookBuffer(savedData.buffer, savedData.name, true);
        showToast(`โหลดข้อมูลที่บันทึกไว้ล่าสุด (${savedData.name}) สำเร็จ`);
        return;
      }

      // 2. Otherwise auto-load default ExportPWS.xlsx
      await fetchDefaultFile('ExportPWS.xlsx');

    } catch (err) {
      console.warn('Auto-load notice:', err);
      showLoading(false);
      // If running via file:// protocol where fetch is blocked by browser security
      showError(
        'กำลังเปิดผ่าน Local File (file://) : กรุณากดปุ่ม "อัปเดตไฟล์ ExportPWS.xlsx" ด้านบนขวาเพื่อเลือกไฟล์จากเครื่อง หรือเปิดผ่าน Netlify'
      );
      if (errorActionBtn) {
        errorActionBtn.classList.remove('hidden');
        errorActionBtn.onclick = () => fileInput.click();
      }
    }
  }

  async function fetchDefaultFile(filename) {
    const res = await fetch(filename);
    if (!res.ok) {
      throw new Error(`HTTP Error ${res.status}: ไม่สามารถอ่าน ${filename} ได้`);
    }
    const buf = await res.arrayBuffer();
    parseWorkbookBuffer(buf, filename, false);
  }

  // ==========================================
  // Event Listeners: Update File & Drag & Drop
  // ==========================================

  // Click Update Button
  updateFileBtn.addEventListener('click', () => {
    fileInput.click();
  });

  // File chosen via picker
  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      handleUserUploadedFile(file);
    }
  });

  // Global Window Drag & Drop
  window.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dragCounter++;
    if (e.dataTransfer && e.dataTransfer.types.includes('Files')) {
      dragDropOverlay.classList.remove('hidden');
    }
  });

  window.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dragCounter--;
    if (dragCounter <= 0) {
      dragCounter = 0;
      dragDropOverlay.classList.add('hidden');
    }
  });

  window.addEventListener('dragover', (e) => {
    e.preventDefault();
  });

  window.addEventListener('drop', (e) => {
    e.preventDefault();
    dragCounter = 0;
    dragDropOverlay.classList.add('hidden');

    const files = e.dataTransfer?.files;
    if (files && files.length > 0) {
      handleUserUploadedFile(files[0]);
    }
  });

  function handleUserUploadedFile(file) {
    hideError();
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const fileName = file.name.toLowerCase();
    const isValid = validExtensions.some((ext) => fileName.endsWith(ext));

    if (!isValid) {
      showError(`ไฟล์ "${file.name}" ไม่ใช่ไฟล์ Excel หรือ CSV ที่รองรับ (.xlsx, .xls, .csv)`);
      return;
    }

    showLoading(true, `กำลังอัปเดตข้อมูลจาก "${file.name}"...`);

    const reader = new FileReader();
    reader.onload = (e) => {
      const buffer = e.target.result;
      parseWorkbookBuffer(buffer, file.name, true);

      // Save to IndexedDB so it persists across reloads
      saveToIndexedDB(file.name, buffer);
      showToast(`อัปเดตข้อมูลจาก "${file.name}" สำเร็จเรียบร้อย!`);
    };

    reader.onerror = () => {
      showLoading(false);
      showError('ไม่สามารถอ่านไฟล์ได้ กรุณาลองใหม่อีกครั้ง');
    };

    reader.readAsArrayBuffer(file);
  }

  // ==========================================
  // Core: Parse & Process File
  // ==========================================

  function parseWorkbookBuffer(arrayBuffer, filename, isCustom = false) {
    try {
      if (typeof XLSX === 'undefined') {
        throw new Error('ไม่พบไลบรารี SheetJS (XLSX)');
      }

      const data = new Uint8Array(arrayBuffer);
      const workbook = XLSX.read(data, {
        type: 'array',
        cellDates: true,
        dense: true
      });

      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        throw new Error('ไม่พบ Sheet ข้อมูลในไฟล์นี้');
      }

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      if (!worksheet || !worksheet['!ref']) {
        throw new Error(`Sheet แรก "${firstSheetName}" ไม่มีข้อมูล`);
      }

      // Convert to 2D array
      const raw2D = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: '',
        blankrows: false
      });

      if (!raw2D || raw2D.length < 2) {
        throw new Error('ไฟล์นี้มีข้อมูลไม่เพียงพอ หรือไม่มีแถวข้อมูล');
      }

      const headers = raw2D[0].map((h) => String(h || '').trim());
      colIndexMap = {};
      headers.forEach((h, idx) => {
        colIndexMap[h.toUpperCase()] = idx;
      });

      // Match yellow columns
      activeDisplayCols = [];
      YELLOW_COLUMNS.forEach((yCol) => {
        if (colIndexMap[yCol] !== undefined) {
          activeDisplayCols.push({
            name: yCol,
            index: colIndexMap[yCol]
          });
        }
      });

      // Fallback if not matching yellow columns (e.g. another custom file)
      if (activeDisplayCols.length === 0) {
        headers.forEach((h, idx) => {
          activeDisplayCols.push({ name: h || `Column ${idx + 1}`, index: idx });
        });
      }

      // Store data rows (from row 1 onwards)
      allRows = raw2D.slice(1);

      // Populate filter dropdowns with unique values
      populateFilterOptions();

      // Reset filters and render
      filterLocationName.value = '';
      filterLocationType.value = '';
      filterArticleDesc.value = '';
      clearSearchBtn.classList.add('hidden');
      currentPage = 1;

      // Update UI Header
      activeFileName.textContent = filename;
      if (isCustom) {
        fileCustomTag.classList.remove('hidden');
      } else {
        fileCustomTag.classList.add('hidden');
      }

      dashboardContent.classList.remove('hidden');
      hideError();

      // Apply initial filter (all)
      applyFilters();

      showLoading(false);

    } catch (err) {
      showLoading(false);
      console.error(err);
      showError(`เกิดข้อผิดพลาด: ${err.message || 'ไม่สามารถแปลงไฟล์ได้'}`);
    }
  }

  // ==========================================
  // Filter & KPI Calculation
  // ==========================================

  function populateFilterOptions() {
    const locNameIdx = colIndexMap['LOCATION_NAME'];
    const locTypeIdx = colIndexMap['LOCATION_TYPE'];

    const uniqueLocNames = new Set();
    const uniqueLocTypes = new Set();

    for (let i = 0; i < allRows.length; i++) {
      const row = allRows[i];
      if (locNameIdx !== undefined) {
        const val = String(row[locNameIdx] || '').trim();
        if (val) uniqueLocNames.add(val);
      }
      if (locTypeIdx !== undefined) {
        const val = String(row[locTypeIdx] || '').trim();
        if (val) uniqueLocTypes.add(val);
      }
    }

    // Populate LOCATION_NAME dropdown
    const sortedLocNames = Array.from(uniqueLocNames).sort((a, b) => a.localeCompare(b, 'th'));
    let locOptions = '<option value="">ทั้งหมด (ทุกสาขา / สถานที่)</option>';
    sortedLocNames.forEach((name) => {
      locOptions += `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`;
    });
    filterLocationName.innerHTML = locOptions;
    locationNameCount.textContent = `${sortedLocNames.length.toLocaleString('th-TH')} สาขา`;

    // Populate LOCATION_TYPE dropdown
    const sortedLocTypes = Array.from(uniqueLocTypes).sort();
    let typeOptions = '<option value="">ทั้งหมด (ทุกประเภทสถานที่)</option>';
    sortedLocTypes.forEach((type) => {
      typeOptions += `<option value="${escapeHtml(type)}">${escapeHtml(type)}</option>`;
    });
    filterLocationType.innerHTML = typeOptions;
    locationTypeCount.textContent = `${sortedLocTypes.length} ประเภท`;
  }

  function applyFilters() {
    const targetLocName = filterLocationName.value.trim().toLowerCase();
    const targetLocType = filterLocationType.value.trim().toLowerCase();
    const targetArticleQuery = filterArticleDesc.value.trim().toLowerCase();

    const locNameIdx = colIndexMap['LOCATION_NAME'];
    const locTypeIdx = colIndexMap['LOCATION_TYPE'];
    const articleDescIdx = colIndexMap['ARTICLE_DESC'];
    const qtyIdx = colIndexMap['QTY'];
    const reservedQtyIdx = colIndexMap['RESERVED_QTY'];

    filteredRows = [];
    let sumQty = 0;
    let sumReservedQty = 0;
    const matchingLocations = new Set();
    const matchingArticles = new Set();

    for (let i = 0; i < allRows.length; i++) {
      const row = allRows[i];

      // 1. LOCATION_NAME filter
      if (targetLocName && locNameIdx !== undefined) {
        const locVal = String(row[locNameIdx] || '').trim().toLowerCase();
        if (locVal !== targetLocName) continue;
      }

      // 2. LOCATION_TYPE filter
      if (targetLocType && locTypeIdx !== undefined) {
        const typeVal = String(row[locTypeIdx] || '').trim().toLowerCase();
        if (typeVal !== targetLocType) continue;
      }

      // 3. ARTICLE_DESC partial match
      if (targetArticleQuery && articleDescIdx !== undefined) {
        const artVal = String(row[articleDescIdx] || '').toLowerCase();
        if (!artVal.includes(targetArticleQuery)) continue;
      }

      // Row matches all filters
      filteredRows.push(row);

      // Accumulate KPIs
      if (qtyIdx !== undefined) {
        const qVal = parseFloat(row[qtyIdx]);
        if (!isNaN(qVal)) sumQty += qVal;
      }

      if (reservedQtyIdx !== undefined) {
        const rVal = parseFloat(row[reservedQtyIdx]);
        if (!isNaN(rVal)) sumReservedQty += rVal;
      }

      if (locNameIdx !== undefined) {
        const loc = String(row[locNameIdx] || '').trim();
        if (loc) matchingLocations.add(loc);
      }

      if (articleDescIdx !== undefined) {
        const art = String(row[articleDescIdx] || '').trim();
        if (art) matchingArticles.add(art);
      }
    }

    // Update KPIs
    updateKPIs({
      totalQty: sumQty,
      reservedQty: sumReservedQty,
      totalRecords: filteredRows.length,
      allRecordsCount: allRows.length,
      activeLocationsCount: matchingLocations.size,
      articlesCount: matchingArticles.size
    });

    // Update Filter counter badge
    filterMatchingCount.textContent = `แสดง ${filteredRows.length.toLocaleString('th-TH')} จาก ${allRows.length.toLocaleString('th-TH')} รายการ`;
    tableRecordCountBadge.textContent = `${filteredRows.length.toLocaleString('th-TH')} รายการ`;

    // Render table
    renderTableHeader();
    renderTablePage();
  }

  function updateKPIs(data) {
    // KPI 1: Total QTY
    kpiTotalQty.textContent = Math.round(data.totalQty).toLocaleString('th-TH');

    // KPI 2: Reserved QTY & %
    kpiReservedQty.textContent = Math.round(data.reservedQty).toLocaleString('th-TH');
    const reservedPercent = data.totalQty > 0
      ? ((data.reservedQty / data.totalQty) * 100).toFixed(1)
      : '0';
    kpiReservedPercent.textContent = `${reservedPercent}% ของจำนวนทั้งหมด`;

    // KPI 3: Total Records & % of overall
    kpiTotalRecords.textContent = data.totalRecords.toLocaleString('th-TH');
    const recordsPercent = data.allRecordsCount > 0
      ? ((data.totalRecords / data.allRecordsCount) * 100).toFixed(1)
      : '100';
    kpiRecordsSub.textContent = `คิดเป็น ${recordsPercent}% จากทั้งหมด ${data.allRecordsCount.toLocaleString('th-TH')} รายการ`;

    // KPI 4: Active Locations & Unique Articles
    kpiActiveLocations.textContent = data.activeLocationsCount.toLocaleString('th-TH');
    kpiArticlesCount.textContent = `พบสินค้าไม่ซ้ำ ${data.articlesCount.toLocaleString('th-TH')} รายการ`;
  }

  // ==========================================
  // Filter Event Listeners
  // ==========================================

  filterLocationName.addEventListener('change', () => {
    currentPage = 1;
    applyFilters();
  });

  filterLocationType.addEventListener('change', () => {
    currentPage = 1;
    applyFilters();
  });

  filterArticleDesc.addEventListener('input', (e) => {
    const val = e.target.value.trim();
    if (val) {
      clearSearchBtn.classList.remove('hidden');
    } else {
      clearSearchBtn.classList.add('hidden');
    }

    clearTimeout(searchDebounceTimer);
    searchDebounceTimer = setTimeout(() => {
      currentPage = 1;
      applyFilters();
    }, 250);
  });

  clearSearchBtn.addEventListener('click', () => {
    filterArticleDesc.value = '';
    clearSearchBtn.classList.add('hidden');
    currentPage = 1;
    applyFilters();
    filterArticleDesc.focus();
  });

  resetFiltersBtn.addEventListener('click', () => {
    filterLocationName.value = '';
    filterLocationType.value = '';
    filterArticleDesc.value = '';
    clearSearchBtn.classList.add('hidden');
    currentPage = 1;
    applyFilters();
  });

  // Table pagination listeners
  rowsPerPageSelect.addEventListener('change', (e) => {
    rowsPerPage = parseInt(e.target.value, 10) || 25;
    currentPage = 1;
    renderTablePage();
  });

  btnFirstPage.addEventListener('click', () => {
    if (currentPage > 1) {
      currentPage = 1;
      renderTablePage();
    }
  });

  btnPrevPage.addEventListener('click', () => {
    if (currentPage > 1) {
      currentPage--;
      renderTablePage();
    }
  });

  btnNextPage.addEventListener('click', () => {
    const totalPages = Math.ceil(filteredRows.length / rowsPerPage) || 1;
    if (currentPage < totalPages) {
      currentPage++;
      renderTablePage();
    }
  });

  btnLastPage.addEventListener('click', () => {
    const totalPages = Math.ceil(filteredRows.length / rowsPerPage) || 1;
    if (currentPage < totalPages) {
      currentPage = totalPages;
      renderTablePage();
    }
  });

  exportCsvBtn.addEventListener('click', exportFilteredDataToCSV);

  // ==========================================
  // Render: Yellow Column Chips & Table
  // ==========================================

  function renderYellowChips() {
    yellowChipsList.innerHTML = '';
    const fragment = document.createDocumentFragment();
    YELLOW_COLUMNS.forEach((col) => {
      const chip = document.createElement('span');
      chip.className = 'chip';
      chip.textContent = col;
      fragment.appendChild(chip);
    });
    yellowChipsList.appendChild(fragment);
  }

  function renderTableHeader() {
    let thHtml = '<tr><th class="col-row-num">#</th>';
    activeDisplayCols.forEach((colDef) => {
      const isQty = colDef.name === 'QTY' || colDef.name.includes('QTY');
      const alignClass = isQty ? ' style="text-align: right;"' : '';
      thHtml += `<th title="${escapeHtml(colDef.name)}"${alignClass}>${escapeHtml(colDef.name)}</th>`;
    });
    thHtml += '</tr>';
    tableHead.innerHTML = thHtml;
  }

  function renderTablePage() {
    const totalRecords = filteredRows.length;
    const totalPages = Math.ceil(totalRecords / rowsPerPage) || 1;

    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;

    const startIdx = (currentPage - 1) * rowsPerPage;
    const endIdx = Math.min(startIdx + rowsPerPage, totalRecords);
    const pageRows = filteredRows.slice(startIdx, endIdx);

    if (totalRecords === 0) {
      paginationInfo.textContent = 'ไม่พบข้อมูลที่ตรงกับเงื่อนไขตัวกรอง';
      pageIndicator.textContent = 'หน้า 0 / 0';
      btnFirstPage.disabled = true;
      btnPrevPage.disabled = true;
      btnNextPage.disabled = true;
      btnLastPage.disabled = true;
      tableBody.innerHTML = `
        <tr>
          <td colspan="${activeDisplayCols.length + 1}" style="text-align: center; padding: 40px; color: var(--text-muted);">
            ไม่พบข้อมูลที่ตรงกับเงื่อนไขตัวกรอง กรุณาลองปรับเปลี่ยนตัวกรองใหม่
          </td>
        </tr>
      `;
      return;
    }

    paginationInfo.textContent = `แสดงแถวที่ ${(startIdx + 1).toLocaleString('th-TH')} - ${endIdx.toLocaleString('th-TH')} จากทั้งหมด ${totalRecords.toLocaleString('th-TH')} รายการ`;
    pageIndicator.textContent = `หน้า ${currentPage} / ${totalPages}`;

    btnFirstPage.disabled = currentPage === 1;
    btnPrevPage.disabled = currentPage === 1;
    btnNextPage.disabled = currentPage === totalPages;
    btnLastPage.disabled = currentPage === totalPages;

    const tbodyFragment = document.createDocumentFragment();

    pageRows.forEach((row, rIdx) => {
      const globalRowNumber = startIdx + rIdx + 1;
      const tr = document.createElement('tr');
      let rowHtml = `<td class="col-row-num">${globalRowNumber}</td>`;

      activeDisplayCols.forEach((colDef) => {
        const val = row[colDef.index];
        const isQtyCol = colDef.name === 'QTY';
        const isReservedCol = colDef.name === 'RESERVED_QTY' || colDef.name === 'RESERVED_RTF_QTY' || colDef.name === 'RESERVED_QUOTA';

        let cellClass = '';
        let formattedVal = '';

        if (val === undefined || val === null || val === '') {
          formattedVal = '<span class="cell-empty">-</span>';
        } else if (isQtyCol) {
          cellClass = 'cell-qty';
          const n = parseFloat(val);
          formattedVal = !isNaN(n) ? n.toLocaleString('th-TH') : escapeHtml(String(val));
        } else if (isReservedCol) {
          cellClass = 'cell-reserved';
          const n = parseFloat(val);
          formattedVal = !isNaN(n) ? n.toLocaleString('th-TH') : escapeHtml(String(val));
        } else if (typeof val === 'number') {
          cellClass = 'cell-numeric';
          formattedVal = val.toLocaleString('th-TH');
        } else if (val instanceof Date && !isNaN(val)) {
          formattedVal = escapeHtml(val.toLocaleDateString('th-TH'));
        } else {
          formattedVal = escapeHtml(String(val));
        }

        rowHtml += `<td class="${cellClass}">${formattedVal}</td>`;
      });

      tr.innerHTML = rowHtml;
      tbodyFragment.appendChild(tr);
    });

    tableBody.innerHTML = '';
    tableBody.appendChild(tbodyFragment);
  }

  // ==========================================
  // Export CSV
  // ==========================================

  function exportFilteredDataToCSV() {
    if (filteredRows.length === 0) {
      alert('ไม่มีข้อมูลให้ส่งออก');
      return;
    }

    const headers = activeDisplayCols.map((c) => `"${c.name.replace(/"/g, '""')}"`);
    const csvLines = [headers.join(',')];

    filteredRows.forEach((row) => {
      const line = activeDisplayCols.map((colDef) => {
        let val = row[colDef.index];
        if (val === undefined || val === null) val = '';
        if (val instanceof Date) {
          val = val.toLocaleDateString('th-TH');
        }
        val = String(val).replace(/"/g, '""');
        return `"${val}"`;
      });
      csvLines.push(line.join(','));
    });

    const csvContent = '\uFEFF' + csvLines.join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    const timestamp = new Date().toISOString().slice(0, 10);
    link.href = url;
    link.download = `ExportPWS_ExecutiveFiltered_${timestamp}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  // ==========================================
  // IndexedDB Persistence Helpers
  // ==========================================

  function getDB() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function saveToIndexedDB(filename, buffer) {
    try {
      const db = await getDB();
      const tx = db.transaction(STORE_NAME, 'readwrite');
      const store = tx.objectStore(STORE_NAME);
      store.put({ name: filename, buffer: buffer, updatedAt: Date.now() }, 'active_file');
    } catch (e) {
      console.warn('IndexedDB save failed:', e);
    }
  }

  async function loadFromIndexedDB() {
    try {
      const db = await getDB();
      return new Promise((resolve) => {
        const tx = db.transaction(STORE_NAME, 'readonly');
        const store = tx.objectStore(STORE_NAME);
        const req = store.get('active_file');
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => resolve(null);
      });
    } catch (e) {
      return null;
    }
  }

  // ==========================================
  // UI Helpers: Toast & Notifications
  // ==========================================

  function showToast(message, type = 'success') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
      <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
        <polyline points="22 4 12 14.01 9 11.01"></polyline>
      </svg>
      <span>${escapeHtml(message)}</span>
    `;
    toastContainer.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(8px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 3500);
  }

  function escapeHtml(str) {
    if (typeof str !== 'string') return String(str);
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function showLoading(isLoading, text) {
    if (isLoading) {
      if (text) loadingText.textContent = text;
      loadingSpinner.classList.remove('hidden');
    } else {
      loadingSpinner.classList.add('hidden');
    }
  }

  function showError(msg) {
    errorMessage.textContent = msg;
    errorAlert.classList.remove('hidden');
  }

  function hideError() {
    errorAlert.classList.add('hidden');
  }
});
