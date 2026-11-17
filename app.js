/**
 * Executive Inventory & Asset Dashboard (PWS Analytics)
 * Features:
 * 1. Filter: LOCATION_NAME, LOCATION_TYPE, ARTICLE_DESC (partial text search)
 * 2. 4 KPI Cards: Total QTY, Total Reserved QTY, Total Records, Active Locations
 * 3. Detail Data Table: Filtered records showing the 22 yellow columns with high-performance pagination
 * 4. Export Filtered Records to CSV (with UTF-8 BOM for Thai Excel support)
 */

document.addEventListener('DOMContentLoaded', () => {
  // 22 Yellow Columns from ExportPWS.xlsx
  const YELLOW_COLUMNS = [
    'PROVINCE', 'LOCATION_ID', 'LOCATION_CODE', 'LOCATION_NAME', 'LOCATION_TYPE',
    'GROUP_ARTICLE', 'ARTICLE_ID', 'ARTICLE_NO', 'ARTICLE_TYPE', 'ARTICLE_DESC',
    'BRAND', 'MODEL', 'DOC_ID', 'UNIT_COUNT', 'SITECODE_OSPID', 'QTY',
    'RESERVED_QTY', 'RESERVED_RTF_QTY', 'RESERVED_QUOTA', 'CREATED_DATE', 'OID', 'TITLE'
  ];

  // DOM Elements
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input');
  const browseBtn = document.getElementById('browse-btn');
  const autoLoadBtn = document.getElementById('auto-load-btn');
  const uploadContainer = document.getElementById('upload-container');
  const errorAlert = document.getElementById('error-alert');
  const errorMessage = document.getElementById('error-message');
  const loadingSpinner = document.getElementById('loading-spinner');
  const loadingText = document.getElementById('loading-text');
  const dashboardContent = document.getElementById('dashboard-content');

  // Header status
  const fileStatusBadge = document.getElementById('file-status-badge');
  const activeFileName = document.getElementById('active-filename');
  const changeFileBtn = document.getElementById('change-file-btn');

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
  let allRows = [];                  // All data objects
  let filteredRows = [];             // Filtered data objects
  let activeDisplayCols = [];        // Column definitions to display
  let colIndexMap = {};              // Map: colName -> columnIndex in raw row
  let currentPage = 1;
  let rowsPerPage = 25;
  let searchDebounceTimer = null;

  // Initialize Yellow Chips bar
  renderYellowChips();

  // ==========================================
  // Event Listeners: File Upload & Drag & Drop
  // ==========================================

  browseBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
  });

  dropZone.addEventListener('click', () => {
    fileInput.click();
  });

  dropZone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      fileInput.click();
    }
  });

  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) processFile(file);
  });

  changeFileBtn.addEventListener('click', () => {
    uploadContainer.classList.remove('hidden');
    uploadContainer.scrollIntoView({ behavior: 'smooth' });
  });

  // Direct Auto-load Button for ExportPWS.xlsx
  autoLoadBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    tryFetchLocalFile('ExportPWS.xlsx');
  });

  // Drag & Drop
  ['dragenter', 'dragover'].forEach((eventName) => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add('dragover');
    });
  });

  ['dragleave', 'dragend'].forEach((eventName) => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (!dropZone.contains(e.relatedTarget)) {
        dropZone.classList.remove('dragover');
      }
    });
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.remove('dragover');
    const files = e.dataTransfer?.files;
    if (files && files.length > 0) {
      processFile(files[0]);
    }
  });

  // ==========================================
  // Event Listeners: Filters & Table Controls
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

  // Pagination buttons
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

  // Export CSV
  exportCsvBtn.addEventListener('click', exportFilteredDataToCSV);

  // ==========================================
  // Core: Parse & Process File
  // ==========================================

  function tryFetchLocalFile(filename) {
    showLoading(true, `กำลังอ่านไฟล์ ${filename}...`);
    fetch(filename)
      .then((res) => {
        if (!res.ok) throw new Error('ไม่สามารถอ่านไฟล์ได้ทาง HTTP (หากเปิดแบบ file:// ให้ลากไฟล์มาวาง)');
        return res.arrayBuffer();
      })
      .then((buf) => {
        parseWorkbookBuffer(buf, filename);
      })
      .catch((err) => {
        showLoading(false);
        showError(`ไม่สามารถโหลดอัตโนมัติได้: ${err.message}. กรุณาลากไฟล์ ${filename} มาวางบนกล่องด้านบนแทนครับ`);
      });
  }

  function processFile(file) {
    hideError();
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const fileName = file.name.toLowerCase();
    const isValid = validExtensions.some((ext) => fileName.endsWith(ext));

    if (!isValid) {
      showError(`ไฟล์ "${file.name}" ไม่ใช่ไฟล์ Excel หรือ CSV ที่รองรับ (.xlsx, .xls, .csv)`);
      return;
    }

    showLoading(true, `กำลังประมวลผล "${file.name}" ด้วย SheetJS...`);

    const reader = new FileReader();
    reader.onload = (e) => {
      parseWorkbookBuffer(e.target.result, file.name);
    };
    reader.onerror = () => {
      showLoading(false);
      showError('ไม่สามารถอ่านไฟล์ได้ กรุณาลองใหม่อีกครั้ง');
    };
    reader.readAsArrayBuffer(file);
  }

  function parseWorkbookBuffer(arrayBuffer, filename) {
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

      // Update UI state
      activeFileName.textContent = filename;
      fileStatusBadge.classList.remove('hidden');
      uploadContainer.classList.add('hidden');
      dashboardContent.classList.remove('hidden');

      // Apply initial filter (all)
      applyFilters();

      showLoading(false);
      window.scrollTo({ top: 0, behavior: 'smooth' });

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

    // Render table with current page
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

    // Constrain currentPage
    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;

    const startIdx = (currentPage - 1) * rowsPerPage;
    const endIdx = Math.min(startIdx + rowsPerPage, totalRecords);
    const pageRows = filteredRows.slice(startIdx, endIdx);

    // Update Pagination UI
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

    // Build Table Body rows
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
  // Export Filtered Data to CSV (UTF-8 BOM)
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

    // Add UTF-8 BOM (\uFEFF) so Excel opens Thai characters correctly
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
  // Helpers
  // ==========================================

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
    errorAlert.scrollIntoView({ behavior: 'smooth' });
  }

  function hideError() {
    errorAlert.classList.add('hidden');
  }
});
