/**
 * Executive Inventory & Asset Dashboard (PWS Analytics)
 * Features:
 * 1. Immediate instant load on startup (from IndexedDB if previously uploaded, or embedded data.js)
 *    - Completely safe for file:/// protocol (Zero CORS errors) and Netlify
 * 2. Remembers and persists user-uploaded Excel files in IndexedDB permanently
 * 3. Prominent "อัปเดตไฟล์ ExportPWS.xlsx" button in Header
 * 4. Executive Filters, 4 KPI Cards, 22 Yellow Columns Table with high-performance pagination
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
  const activeFileDate = document.getElementById('active-filedate');
  const fileCustomTag = document.getElementById('file-custom-tag');
  const resetDefaultDataBtn = document.getElementById('reset-default-data-btn');

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
  const rowsPerPageSelect = document.getElementById('rows-per-page-select');
  const exportCsvBtn = document.getElementById('export-csv-btn');
  const exportBtnText = document.getElementById('export-btn-text');

  const paginationInfo = document.getElementById('pagination-info');
  const pageIndicator = document.getElementById('page-indicator');
  const btnFirstPage = document.getElementById('btn-first-page');
  const btnPrevPage = document.getElementById('btn-prev-page');
  const btnNextPage = document.getElementById('btn-next-page');
  const btnLastPage = document.getElementById('btn-last-page');

  // State
  let allRows = [];                  // All data rows
  let filteredRows = [];             // Filtered rows
  let activeDisplayCols = [];        // Column definitions
  let colIndexMap = {};              // Map: colName -> columnIndex
  let currentPage = 1;
  let rowsPerPage = 25;
  let searchDebounceTimer = null;
  let dragCounter = 0;

  // ==========================================
  // Startup: Auto-load Data with Priority
  // ==========================================
  initDashboard();

  async function initDashboard() {
    showLoading(true, 'กำลังโหลดข้อมูลเข้าสู่แดชบอร์ด...');

    try {
      // Priority 1: Check if user previously uploaded and saved a custom file in IndexedDB
      const saved = await loadFromIndexedDB();
      if (saved && saved.name && saved.headers && saved.rows && saved.rows.length > 0) {
        applyParsedData(saved.headers, saved.rows, saved.name, true, saved.fileDate || 'บันทึกไว้ในเครื่อง');
        showToast(`โหลดข้อมูลล่าสุดที่เคยบันทึกไว้ (${saved.name}) เรียบร้อย`);
        return;
      }

      // Priority 2: Use embedded default data (data.js) - Works 100% on file:/// and web with 0 latency
      if (window.DEFAULT_PWS_DATA && window.DEFAULT_PWS_DATA.headers && window.DEFAULT_PWS_DATA.rows) {
        applyParsedData(window.DEFAULT_PWS_DATA.headers, window.DEFAULT_PWS_DATA.rows, 'ExportPWS.xlsx', false, '17/11/2026 21:34 น.');
        return;
      }

      // Priority 3: Fetch ExportPWS.xlsx over HTTP (if hosted on web server / Netlify)
      await fetchDefaultFile('ExportPWS.xlsx');

    } catch (err) {
      console.warn('Dashboard init notice:', err);
      showLoading(false);
      showError('ไม่สามารถโหลดข้อมูลอัตโนมัติได้ กรุณากดปุ่ม "อัปเดตไฟล์ ExportPWS.xlsx" ด้านบนขวา');
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
    parseWorkbookBuffer(buf, filename, false, '17/11/2026 21:34 น.');
  }

  // ==========================================
  // Core: Apply Data to State & UI
  // ==========================================

  function applyParsedData(headers, dataRows, filename, isCustom = false, fileDate = '') {
    colIndexMap = {};
    headers.forEach((h, idx) => {
      colIndexMap[String(h).trim().toUpperCase()] = idx;
    });

    // Match 22 yellow columns
    activeDisplayCols = [];
    YELLOW_COLUMNS.forEach((yCol) => {
      if (colIndexMap[yCol] !== undefined) {
        activeDisplayCols.push({
          name: yCol,
          index: colIndexMap[yCol]
        });
      }
    });

    // Fallback if no yellow column matches
    if (activeDisplayCols.length === 0) {
      headers.forEach((h, idx) => {
        activeDisplayCols.push({ name: h || `Column ${idx + 1}`, index: idx });
      });
    }

    allRows = dataRows;

    // Populate Filters
    populateFilterOptions();

    // Reset Filter Inputs
    filterLocationName.value = '';
    filterLocationType.value = '';
    filterArticleDesc.value = '';
    clearSearchBtn.classList.add('hidden');
    currentPage = 1;

    // Header Status & Date
    activeFileName.textContent = filename;
    if (activeFileDate && fileDate) {
      activeFileDate.textContent = `ข้อมูล ณ วันที่: ${fileDate}`;
    }

    if (isCustom) {
      fileCustomTag.classList.remove('hidden');
      if (resetDefaultDataBtn) resetDefaultDataBtn.classList.remove('hidden');
    } else {
      fileCustomTag.classList.add('hidden');
      if (resetDefaultDataBtn) resetDefaultDataBtn.classList.add('hidden');
    }

    dashboardContent.classList.remove('hidden');
    hideError();

    // Calculate KPIs and render
    applyFilters();
    showLoading(false);
  }

  // ==========================================
  // File Upload & Update Handling
  // ==========================================

  updateFileBtn.addEventListener('click', () => {
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      handleUserUploadedFile(file);
    }
  });

  if (resetDefaultDataBtn) {
    resetDefaultDataBtn.addEventListener('click', async () => {
      if (confirm('ต้องการรีเซ็ตกลับไปใช้ข้อมูล ExportPWS.xlsx เริ่มต้นใช่หรือไม่?')) {
        await clearIndexedDB();
        if (window.DEFAULT_PWS_DATA && window.DEFAULT_PWS_DATA.headers) {
          applyParsedData(window.DEFAULT_PWS_DATA.headers, window.DEFAULT_PWS_DATA.rows, 'ExportPWS.xlsx', false, '17/11/2026 21:34 น.');
          showToast('รีเซ็ตกลับเป็นไฟล์ ExportPWS.xlsx เริ่มต้นแล้ว');
        } else {
          location.reload();
        }
      }
    });
  }

  // Global Drag & Drop Overlay
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

    showLoading(true, `กำลังอ่านและประมวลผล "${file.name}"...`);

    // Format file date
    const d = file.lastModified ? new Date(file.lastModified) : new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const formattedDate = `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())} น.`;

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const buffer = e.target.result;
        const parsed = parseWorkbookBuffer(buffer, file.name, true, formattedDate);

        if (parsed && parsed.headers && parsed.rows) {
          // Save parsed data & date into IndexedDB for instant reload in future visits
          await saveToIndexedDB(file.name, parsed.headers, parsed.rows, formattedDate);
          showToast(`บันทึกข้อมูลจาก "${file.name}" (${formattedDate}) เรียบร้อยแล้ว`);
        }
      } catch (err) {
        showLoading(false);
        showError(`เกิดข้อผิดพลาดในการอ่านไฟล์: ${err.message}`);
      }
    };

    reader.onerror = () => {
      showLoading(false);
      showError('ไม่สามารถอ่านไฟล์ได้ กรุณาลองใหม่อีกครั้ง');
    };

    reader.readAsArrayBuffer(file);
  }

  function parseWorkbookBuffer(arrayBuffer, filename, isCustom = false, fileDate = '') {
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

    const raw2D = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: '',
      blankrows: false
    });

    if (!raw2D || raw2D.length < 2) {
      throw new Error('ไฟล์นี้มีข้อมูลไม่เพียงพอ หรือไม่มีแถวข้อมูล');
    }

    const headers = raw2D[0].map((h) => String(h || '').trim());
    const dataRows = raw2D.slice(1);

    applyParsedData(headers, dataRows, filename, isCustom, fileDate);

    return { headers, rows: dataRows };
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

    // Populate LOCATION_NAME
    const sortedLocNames = Array.from(uniqueLocNames).sort((a, b) => a.localeCompare(b, 'th'));
    let locOptions = '<option value="">ทั้งหมด (ทุกสาขา / สถานที่)</option>';
    sortedLocNames.forEach((name) => {
      locOptions += `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`;
    });
    filterLocationName.innerHTML = locOptions;
    locationNameCount.textContent = `${sortedLocNames.length.toLocaleString('th-TH')} สาขา`;

    // Populate LOCATION_TYPE
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

      // 1. LOCATION_NAME
      if (targetLocName && locNameIdx !== undefined) {
        const locVal = String(row[locNameIdx] || '').trim().toLowerCase();
        if (locVal !== targetLocName) continue;
      }

      // 2. LOCATION_TYPE
      if (targetLocType && locTypeIdx !== undefined) {
        const typeVal = String(row[locTypeIdx] || '').trim().toLowerCase();
        if (typeVal !== targetLocType) continue;
      }

      // 3. ARTICLE_DESC partial match
      if (targetArticleQuery && articleDescIdx !== undefined) {
        const artVal = String(row[articleDescIdx] || '').toLowerCase();
        if (!artVal.includes(targetArticleQuery)) continue;
      }

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

    // Update Badges & Export Button
    filterMatchingCount.textContent = `แสดง ${filteredRows.length.toLocaleString('th-TH')} จาก ${allRows.length.toLocaleString('th-TH')} รายการ`;
    tableRecordCountBadge.textContent = `${filteredRows.length.toLocaleString('th-TH')} รายการ`;
    if (exportBtnText) {
      exportBtnText.textContent = `ส่งออกข้อมูลเป็น CSV (${filteredRows.length.toLocaleString('th-TH')} รายการ)`;
    }

    // Render Table
    renderTableHeader();
    renderTablePage();
  }

  function updateKPIs(data) {
    kpiTotalQty.textContent = Math.round(data.totalQty).toLocaleString('th-TH');

    kpiReservedQty.textContent = Math.round(data.reservedQty).toLocaleString('th-TH');
    const reservedPercent = data.totalQty > 0
      ? ((data.reservedQty / data.totalQty) * 100).toFixed(1)
      : '0';
    kpiReservedPercent.textContent = `${reservedPercent}% ของจำนวนทั้งหมด`;

    kpiTotalRecords.textContent = data.totalRecords.toLocaleString('th-TH');
    const recordsPercent = data.allRecordsCount > 0
      ? ((data.totalRecords / data.allRecordsCount) * 100).toFixed(1)
      : '100';
    kpiRecordsSub.textContent = `คิดเป็น ${recordsPercent}% จากทั้งหมด ${data.allRecordsCount.toLocaleString('th-TH')} รายการ`;

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
  // Table Rendering
  // ==========================================

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

    showToast(`ส่งออกข้อมูล CSV เรียบร้อยแล้ว (${filteredRows.length.toLocaleString('th-TH')} รายการ)`);
  }

  // ==========================================
  // IndexedDB Persistence Helpers
  // ==========================================

  function getDB() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, 2);
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

  async function saveToIndexedDB(filename, headers, rows, fileDate = '') {
    try {
      const db = await getDB();
      const tx = db.transaction(STORE_NAME, 'readwrite');
      const store = tx.objectStore(STORE_NAME);
      store.put({
        name: filename,
        headers: headers,
        rows: rows,
        fileDate: fileDate,
        updatedAt: Date.now()
      }, 'active_data');
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
        const req = store.get('active_data');
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => resolve(null);
      });
    } catch (e) {
      return null;
    }
  }

  async function clearIndexedDB() {
    try {
      const db = await getDB();
      const tx = db.transaction(STORE_NAME, 'readwrite');
      const store = tx.objectStore(STORE_NAME);
      store.delete('active_data');
    } catch (e) {
      console.warn('IndexedDB clear failed:', e);
    }
  }

  // ==========================================
  // UI Helpers
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
