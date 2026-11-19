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

  // 7 Core Columns to display in the Web Detail Table (matches user request & image 2)
  const DETAIL_DISPLAY_COLUMNS = [
    { key: 'CREATED_DATE', label: 'Create' },
    { key: 'LOCATION_CODE', label: 'LOCATION_CODE' },
    { key: 'ARTICLE_ID', label: 'ARTICLE_ID' },
    { key: 'ARTICLE_NO', label: 'ARTICLE_NO' },
    { key: 'QTY', label: 'QTY' },
    { key: 'TITLE', label: 'TITLE' }
  ];

  // Filters
  const btnDispatchShortcut = document.getElementById('btn-dispatch-shortcut');
  const dispatchDaysInput = document.getElementById('dispatch-days-input');
  const quickGroupButtonsContainer = document.getElementById('quick-group-buttons');
  const quickStoreButtonsContainer = document.getElementById('quick-store-buttons');
  let activeGroupPrefix = '';
  let activeStoreGroup = ''; // '', 'HUB', 'RNO', 'KONG'
  let allAvailableGroupArts = [];
  const filterLocationCode = document.getElementById('filter-location-code');
  const clearLocationBtn = document.getElementById('clear-location-btn');
  const filterLocationType = document.getElementById('filter-location-type');
  const filterGroupArticle = document.getElementById('filter-group-article');
  const groupArticleCount = document.getElementById('group-article-count');
  const filterArticleDesc = document.getElementById('filter-article-desc');
  const clearSearchBtn = document.getElementById('clear-search-btn');
  const filterTitle = document.getElementById('filter-title');
  const clearTitleBtn = document.getElementById('clear-title-btn');
  const resetFiltersBtn = document.getElementById('reset-filters-btn');
  const filterMatchingCount = document.getElementById('filter-matching-count');
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

  // Pivot Summary Elements
  const pivotTableHead = document.getElementById('pivot-table-head');
  const pivotTableBody = document.getElementById('pivot-table-body');
  const btnExpandAll = document.getElementById('btn-expand-all');
  const btnCollapseAll = document.getElementById('btn-collapse-all');
  const exportPivotCsvBtn = document.getElementById('export-pivot-csv-btn');

  // Drill-down Elements
  const drilldownBadgeContainer = document.getElementById('drilldown-badge-container');
  const drilldownLabelText = document.getElementById('drilldown-label-text');
  const clearDrilldownBtn = document.getElementById('clear-drilldown-btn');
  const detailTableCard = document.getElementById('detail-table-card');

  const paginationInfo = document.getElementById('pagination-info');
  const pageIndicator = document.getElementById('page-indicator');
  const btnFirstPage = document.getElementById('btn-first-page');
  const btnPrevPage = document.getElementById('btn-prev-page');
  const btnNextPage = document.getElementById('btn-next-page');
  const btnLastPage = document.getElementById('btn-last-page');

  // State
  let allRows = [];                  // All data rows
  let filteredRows = [];             // Filtered rows (by top filters)
  let detailRows = [];               // Displayed rows in detail table (filteredRows + drilldown)
  let activeDrilldown = null;        // { level: 1|2|3, typeName, groupName, articleDesc, locationCode }
  let isDispatchShortcutActive = false; // "กระจายของ" X วัน shortcut state
  let activeDisplayCols = [];        // Web Detail Table columns (7 columns)
  let allExportCols = [];            // Full export columns (all 22 yellow columns)
  let colIndexMap = {};              // Map: colName -> columnIndex
  let currentPage = 1;
  let rowsPerPage = 25;
  let searchDebounceTimer = null;
  let dragCounter = 0;

  // ==========================================
  // Startup: Auto-load Data with Priority
  // ==========================================
  initDashboard();

  // ==========================================
  // Supabase Cloud Storage Configuration (Real-Time 0-second sync)
  // ==========================================
  const SUPABASE_PROJECT_URL = 'https://bdtplmqsmtdxqekfheof.supabase.co';
  const SUPABASE_ANON_KEY = 'sb_publishable_mtdlPhRu-hxBjOXdYW7Oxw_meIhhf7d';
  const SUPABASE_BUCKET = 'pws-data';
  const SUPABASE_FILENAME = 'ExportPWS.xlsx';
  const SUPABASE_PUBLIC_URL = `${SUPABASE_PROJECT_URL}/storage/v1/object/public/${SUPABASE_BUCKET}/${SUPABASE_FILENAME}`;

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

      // Priority 2: Fetch latest ExportPWS.xlsx from Supabase Storage (High-Speed Singapore CDN, 0s delay)
      if (window.location.protocol.startsWith('http')) {
        try {
          const supaRes = await fetch(`${SUPABASE_PUBLIC_URL}?_nocache=${Date.now()}`);
          if (supaRes.ok) {
            const buf = await supaRes.arrayBuffer();
            const lastMod = supaRes.headers.get('last-modified');
            let formattedMod = 'ล่าสุดจาก Supabase Cloud';
            if (lastMod) {
              const d = new Date(lastMod);
              if (!isNaN(d.getTime())) {
                formattedMod = d.toLocaleString('th-TH');
              }
            }
            parseWorkbookBuffer(buf, 'ExportPWS.xlsx', false, formattedMod);
            return;
          }
        } catch (supaErr) {
          console.warn('Supabase fetch failed, falling back to repository file:', supaErr);
        }
      }

      // Priority 3: When on HTTP/HTTPS, try repository ExportPWS.xlsx with cache-busting
      if (window.location.protocol.startsWith('http')) {
        try {
          await fetchDefaultFile(`ExportPWS.xlsx?_nocache=${Date.now()}`);
          return;
        } catch (fetchErr) {
          console.warn('Network fetch of ExportPWS.xlsx failed, falling back to embedded data:', fetchErr);
        }
      }

      // Priority 4: Use embedded default data (data.js) - Works 100% on file:/// and offline
      if (window.DEFAULT_PWS_DATA && window.DEFAULT_PWS_DATA.headers && window.DEFAULT_PWS_DATA.rows) {
        applyParsedData(window.DEFAULT_PWS_DATA.headers, window.DEFAULT_PWS_DATA.rows, 'ExportPWS.xlsx', false, '17/11/2026 21:34 น.');
        return;
      }

      // Fallback
      await fetchDefaultFile('ExportPWS.xlsx');

    } catch (err) {
      console.warn('Dashboard init notice:', err);
      showLoading(false);
      showError('ไม่สามารถโหลดข้อมูลอัตโนมัติได้ กรุณากดปุ่ม "อัปเดตไฟล์" ด้านบนขวา');
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

    // 1. Build activeDisplayCols: 7 columns for the web Detail Table (per user request & image 2)
    activeDisplayCols = [];
    DETAIL_DISPLAY_COLUMNS.forEach((col) => {
      if (colIndexMap[col.key] !== undefined) {
        activeDisplayCols.push({
          name: col.label,
          realName: col.key,
          index: colIndexMap[col.key]
        });
      }
    });
    if (activeDisplayCols.length === 0) {
      headers.forEach((h, idx) => {
        activeDisplayCols.push({ name: h || `Column ${idx + 1}`, realName: h, index: idx });
      });
    }

    // 2. Build allExportCols: all 22 yellow columns for CSV export
    allExportCols = [];
    YELLOW_COLUMNS.forEach((yCol) => {
      if (colIndexMap[yCol] !== undefined) {
        allExportCols.push({
          name: yCol,
          index: colIndexMap[yCol]
        });
      }
    });
    if (allExportCols.length === 0) {
      allExportCols = [...activeDisplayCols];
    }

    allRows = dataRows;

    // Populate Filters
    populateFilterOptions();

    // Reset Filter Inputs
    filterLocationCode.value = '';

    // Set default LOCATION_TYPE to 'STORE' if exists
    let hasStoreOption = false;
    for (let i = 0; i < filterLocationType.options.length; i++) {
      if (filterLocationType.options[i].value.toUpperCase() === 'STORE') {
        hasStoreOption = true;
        break;
      }
    }
    filterLocationType.value = hasStoreOption ? 'STORE' : '';

    activeGroupPrefix = '';
    activeStoreGroup = '';
    updateGroupButtonStates();
    updateStoreButtonStates();
    updateGroupArticleDropdown('');
    if (filterGroupArticle) filterGroupArticle.value = '';
    filterArticleDesc.value = '';
    if (filterTitle) filterTitle.value = '';

    if (clearLocationBtn) clearLocationBtn.classList.add('hidden');
    clearSearchBtn.classList.add('hidden');
    if (clearTitleBtn) clearTitleBtn.classList.add('hidden');
    currentPage = 1;
    activeDrilldown = null;

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
        lastUploadedRawBuffer = buffer;
        lastUploadedFilename = file.name || 'ExportPWS.xlsx';
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
    const locTypeIdx = colIndexMap['LOCATION_TYPE'];
    const groupArtIdx = colIndexMap['GROUP_ARTICLE'];
    const uniqueLocTypes = new Set();
    const uniqueGroupArts = new Set();

    for (let i = 0; i < allRows.length; i++) {
      const row = allRows[i];
      if (locTypeIdx !== undefined) {
        const val = String(row[locTypeIdx] || '').trim();
        if (val) uniqueLocTypes.add(val);
      }
      if (groupArtIdx !== undefined) {
        const val = String(row[groupArtIdx] || '').trim();
        if (val) uniqueGroupArts.add(val);
      }
    }

    // Populate LOCATION_TYPE
    const sortedLocTypes = Array.from(uniqueLocTypes).sort();
    let typeOptions = '<option value="">All LOCATION_TYPE</option>';
    sortedLocTypes.forEach((type) => {
      typeOptions += `<option value="${escapeHtml(type)}">${escapeHtml(type)}</option>`;
    });
    filterLocationType.innerHTML = typeOptions;
    if (locationTypeCount) locationTypeCount.textContent = `${sortedLocTypes.length} Types`;

    // Store all unique groups and populate dropdown matching activeGroupPrefix
    allAvailableGroupArts = Array.from(uniqueGroupArts)
      .filter((grp) => grp && grp !== 'GROUP_ARTICLE')
      .sort();

    updateGroupArticleDropdown(activeGroupPrefix);

    // Populate Quick Group Article Buttons
    populateQuickGroupButtons(uniqueGroupArts);
  }

  function updateGroupArticleDropdown(prefix = '') {
    if (!filterGroupArticle) return;
    const currentVal = filterGroupArticle.value;

    let filtered = allAvailableGroupArts;
    if (prefix) {
      const pUpper = prefix.toUpperCase();
      filtered = allAvailableGroupArts.filter((grp) => {
        const rowPrefix = grp.split('->')[0].trim().toUpperCase();
        return rowPrefix === pUpper;
      });
    }

    const placeholderLabel = prefix
      ? `All ${formatGroupPrefixLabel(prefix)} Groups`
      : 'All GROUP_ARTICLE';

    let groupOptions = `<option value="">${escapeHtml(placeholderLabel)}</option>`;
    filtered.forEach((grp) => {
      const isSelected = grp === currentVal ? ' selected' : '';
      groupOptions += `<option value="${escapeHtml(grp)}"${isSelected}>${escapeHtml(grp)}</option>`;
    });
    filterGroupArticle.innerHTML = groupOptions;

    // Reset value if previous selection is not in the filtered options
    if (currentVal && !filtered.includes(currentVal)) {
      filterGroupArticle.value = '';
    }

    if (groupArticleCount) {
      groupArticleCount.textContent = `${filtered.length} Groups`;
    }
  }

  function formatGroupPrefixLabel(prefix) {
    const pUpper = prefix.toUpperCase();
    if (pUpper === 'TRANSMISSION') return 'Transmission';
    if (pUpper === 'FACILITY') return 'Facility';
    if (pUpper === 'MATERIAL') return 'Material';
    if (pUpper.includes('DATA NETWORK')) return 'Data Network';
    if (pUpper.includes('TOOLS')) return 'Tools & Test';
    if (pUpper.includes('IT EQUIPMENT') || pUpper.startsWith('IT')) return 'IT Equipment';
    return prefix;
  }

  function populateQuickGroupButtons(uniqueGroupArts) {
    if (!quickGroupButtonsContainer) return;

    const uniquePrefixes = new Set();
    uniqueGroupArts.forEach((grp) => {
      if (grp && grp !== 'GROUP_ARTICLE') {
        const prefix = grp.split('->')[0].trim();
        if (prefix) uniquePrefixes.add(prefix);
      }
    });

    const preferredOrder = ['TRANSMISSION', 'Material', 'FACILITY', 'Data Network Equipment', 'TOOLS&TEST_EQUIPMENT', 'IT Equipment'];
    const sortedPrefixes = Array.from(uniquePrefixes).sort((a, b) => {
      const idxA = preferredOrder.indexOf(a);
      const idxB = preferredOrder.indexOf(b);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      if (idxA !== -1) return -1;
      if (idxB !== -1) return 1;
      return a.localeCompare(b);
    });

    let btnHtml = `<button type="button" class="group-btn-chip ${!activeGroupPrefix ? 'active' : ''}" data-prefix="">ทั้งหมด</button>`;
    sortedPrefixes.forEach((prefix) => {
      const label = formatGroupPrefixLabel(prefix);
      const isActive = activeGroupPrefix.toUpperCase() === prefix.toUpperCase();
      btnHtml += `<button type="button" class="group-btn-chip ${isActive ? 'active' : ''}" data-prefix="${escapeHtml(prefix)}">${escapeHtml(label)}</button>`;
    });
    quickGroupButtonsContainer.innerHTML = btnHtml;

    quickGroupButtonsContainer.querySelectorAll('.group-btn-chip').forEach((btn) => {
      btn.addEventListener('click', () => {
        const prefix = btn.dataset.prefix || '';
        if (activeGroupPrefix === prefix && prefix !== '') {
          activeGroupPrefix = ''; // toggle off
        } else {
          activeGroupPrefix = prefix;
        }

        updateGroupButtonStates();
        updateGroupArticleDropdown(activeGroupPrefix);

        if (filterGroupArticle) {
          filterGroupArticle.value = '';
        }

        applyFilters();
      });
    });
  }

  function getStoreGroupRank(code) {
    if (!code) return 99;
    const u = String(code).toUpperCase();
    if (u.includes('HUB')) return 1; // HUB first
    if (u.includes('-M-T-')) return 3; // กองงาน third
    if (/-[MI]-[A-Z]+-/i.test(u)) return 2; // RNO second
    return 4; // Other
  }

  function getStoreGroupName(code) {
    const rank = getStoreGroupRank(code);
    if (rank === 1) return 'HUB';
    if (rank === 2) return 'RNO';
    if (rank === 3) return 'กองงาน';
    return '';
  }

  function matchStoreGroup(locCode, storeGroup) {
    if (!locCode || !storeGroup) return true;
    const u = String(locCode).toUpperCase();
    const isHub = u.includes('HUB');
    const isKong = u.includes('-M-T-');
    const isRno = /-[MI]-[A-Z]+-/i.test(u) && !isKong && !isHub;

    if (storeGroup === 'HUB') return isHub;
    if (storeGroup === 'RNO') return isRno;
    if (storeGroup === 'KONG') return isKong;
    return true;
  }

  function initQuickStoreButtons() {
    if (!quickStoreButtonsContainer) return;
    quickStoreButtonsContainer.querySelectorAll('.store-btn-chip').forEach((btn) => {
      btn.addEventListener('click', () => {
        const storeType = btn.dataset.store || '';
        if (activeStoreGroup === storeType) {
          activeStoreGroup = ''; // toggle off back to all
        } else {
          activeStoreGroup = storeType;
          // Ensure LOCATION_TYPE is STORE when filtering by a store group
          if (filterLocationType) {
            let hasStore = false;
            for (let i = 0; i < filterLocationType.options.length; i++) {
              if (filterLocationType.options[i].value.toUpperCase() === 'STORE') {
                hasStore = true;
                break;
              }
            }
            if (hasStore) filterLocationType.value = 'STORE';
          }
        }
        updateStoreButtonStates();
        currentPage = 1;
        activeDrilldown = null;
        applyFilters();
      });
    });
  }

  function updateStoreButtonStates() {
    if (!quickStoreButtonsContainer) return;
    quickStoreButtonsContainer.querySelectorAll('.store-btn-chip').forEach((b) => {
      const s = b.dataset.store || '';
      if (s === activeStoreGroup) {
        b.classList.add('active');
      } else {
        b.classList.remove('active');
      }
    });
  }

  function updateGroupButtonStates() {
    if (!quickGroupButtonsContainer) return;
    quickGroupButtonsContainer.querySelectorAll('.group-btn-chip').forEach((b) => {
      const bPref = b.dataset.prefix || '';
      if (bPref.toUpperCase() === activeGroupPrefix.toUpperCase()) {
        b.classList.add('active');
      } else {
        b.classList.remove('active');
      }
    });
  }

  function parseRowDate(rawVal) {
    if (!rawVal) return null;
    if (rawVal instanceof Date && !isNaN(rawVal.getTime())) return rawVal;
    if (typeof rawVal === 'string') {
      const trimmed = rawVal.trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) {
        const parts = trimmed.split(/[-T ]/);
        return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
      }
      const d = new Date(trimmed);
      if (!isNaN(d.getTime())) return d;
      const parts = trimmed.split(/[\/\-]/);
      if (parts.length === 3) {
        if (parts[0].length === 4) {
          return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
        } else {
          return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
        }
      }
    }
    if (typeof rawVal === 'number') {
      return new Date((rawVal - (25567 + 2)) * 86400 * 1000);
    }
    return null;
  }

  function applyFilters() {
    const targetLocCode = filterLocationCode.value.trim().toLowerCase();
    const targetLocType = filterLocationType.value.trim().toLowerCase();
    const targetGroupArticle = filterGroupArticle ? filterGroupArticle.value.trim().toLowerCase() : '';
    const targetArticleQuery = filterArticleDesc.value.trim().toLowerCase();
    const targetTitleQuery = filterTitle ? filterTitle.value.trim().toLowerCase() : '';

    const locCodeIdx = colIndexMap['LOCATION_CODE'];
    const locNameIdx = colIndexMap['LOCATION_NAME'];
    const locTypeIdx = colIndexMap['LOCATION_TYPE'];
    const groupArtIdx = colIndexMap['GROUP_ARTICLE'];
    const articleDescIdx = colIndexMap['ARTICLE_DESC'];
    const titleIdx = colIndexMap['TITLE'];
    const dateIdx = colIndexMap['CREATED_DATE'];
    const qtyIdx = colIndexMap['QTY'];
    const reservedQtyIdx = colIndexMap['RESERVED_QTY'];

    // Dispatch Shortcut cutoff date
    let cutoffDate = null;
    if (isDispatchShortcutActive) {
      const days = parseInt(dispatchDaysInput ? dispatchDaysInput.value : 7, 10) || 7;
      let maxDateMs = 0;
      if (dateIdx !== undefined) {
        for (let i = 0; i < allRows.length; i++) {
          const d = parseRowDate(allRows[i][dateIdx]);
          if (d && d.getTime() > maxDateMs) {
            maxDateMs = d.getTime();
          }
        }
      }
      const nowMs = Date.now();
      const refMs = maxDateMs > 0 ? Math.max(nowMs, maxDateMs) : nowMs;
      cutoffDate = new Date(refMs - days * 24 * 60 * 60 * 1000);
      cutoffDate.setHours(0, 0, 0, 0);
    }

    filteredRows = [];
    let sumQty = 0;
    let sumReservedQty = 0;
    const matchingLocations = new Set();
    const matchingArticles = new Set();

    for (let i = 0; i < allRows.length; i++) {
      const row = allRows[i];

      // 1. LOCATION_CODE (Partial match search)
      if (targetLocCode && locCodeIdx !== undefined) {
        const codeVal = String(row[locCodeIdx] || '').toLowerCase();
        if (!codeVal.includes(targetLocCode)) continue;
      }

      // 1.5 Quick Store Group Filter (HUB, RNO, กองงาน)
      if (activeStoreGroup) {
        if (locCodeIdx === undefined) continue;
        const code = String(row[locCodeIdx] || '').trim();
        if (!matchStoreGroup(code, activeStoreGroup)) continue;
      }

      // 2. LOCATION_TYPE
      if (targetLocType && locTypeIdx !== undefined) {
        const typeVal = String(row[locTypeIdx] || '').trim().toLowerCase();
        if (typeVal !== targetLocType) continue;
      }

      // 3. GROUP_ARTICLE / Quick Group Prefix Filter
      if (groupArtIdx !== undefined) {
        const rawGrp = String(row[groupArtIdx] || '').trim();
        if (activeGroupPrefix) {
          const rowPrefix = rawGrp.split('->')[0].trim().toUpperCase();
          if (rowPrefix !== activeGroupPrefix.toUpperCase()) continue;
        }
        if (targetGroupArticle) {
          if (rawGrp.toLowerCase() !== targetGroupArticle) continue;
        }
      }

      // 4. ARTICLE_DESC partial match
      if (targetArticleQuery && articleDescIdx !== undefined) {
        const artVal = String(row[articleDescIdx] || '').toLowerCase();
        if (!artVal.includes(targetArticleQuery)) continue;
      }

      // 5. TITLE partial match
      if (targetTitleQuery && titleIdx !== undefined) {
        const titleVal = String(row[titleIdx] || '').toLowerCase();
        if (!titleVal.includes(targetTitleQuery)) continue;
      }

      // 6. Dispatch Shortcut: CREATED_DATE within past X days
      if (cutoffDate && dateIdx !== undefined) {
        const rowDate = parseRowDate(row[dateIdx]);
        if (!rowDate || rowDate < cutoffDate) continue;
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

    // Render Pivot Summary Table
    buildAndRenderPivotSummary();

    // Render Detail Table with Drilldown
    updateDetailTableWithDrilldown(false);
  }

  function updateKPIs(data) {
    if (!kpiTotalQty) return;
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

  let locSearchDebounceTimer = null;

  filterLocationCode.addEventListener('input', (e) => {
    const val = e.target.value.trim();
    if (val) {
      if (clearLocationBtn) clearLocationBtn.classList.remove('hidden');
    } else {
      if (clearLocationBtn) clearLocationBtn.classList.add('hidden');
    }

    clearTimeout(locSearchDebounceTimer);
    locSearchDebounceTimer = setTimeout(() => {
      currentPage = 1;
      activeDrilldown = null;
      applyFilters();
    }, 250);
  });

  if (clearLocationBtn) {
    clearLocationBtn.addEventListener('click', () => {
      filterLocationCode.value = '';
      clearLocationBtn.classList.add('hidden');
      currentPage = 1;
      activeDrilldown = null;
      applyFilters();
      filterLocationCode.focus();
    });
  }

  filterLocationType.addEventListener('change', () => {
    currentPage = 1;
    activeDrilldown = null;
    applyFilters();
  });

  if (filterGroupArticle) {
    filterGroupArticle.addEventListener('change', () => {
      const selectedVal = filterGroupArticle.value.trim();
      if (selectedVal) {
        const prefix = selectedVal.split('->')[0].trim();
        activeGroupPrefix = prefix;
        updateGroupButtonStates();
      }
      currentPage = 1;
      activeDrilldown = null;
      applyFilters();
    });
  }

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
      activeDrilldown = null;
      applyFilters();
    }, 250);
  });

  clearSearchBtn.addEventListener('click', () => {
    filterArticleDesc.value = '';
    clearSearchBtn.classList.add('hidden');
    currentPage = 1;
    activeDrilldown = null;
    applyFilters();
    filterArticleDesc.focus();
  });

  let titleSearchDebounceTimer = null;

  if (filterTitle) {
    filterTitle.addEventListener('input', (e) => {
      const val = e.target.value.trim();
      if (val) {
        if (clearTitleBtn) clearTitleBtn.classList.remove('hidden');
      } else {
        if (clearTitleBtn) clearTitleBtn.classList.add('hidden');
      }

      clearTimeout(titleSearchDebounceTimer);
      titleSearchDebounceTimer = setTimeout(() => {
        currentPage = 1;
        activeDrilldown = null;
        applyFilters();
      }, 250);
    });
  }

  if (clearTitleBtn) {
    clearTitleBtn.addEventListener('click', () => {
      filterTitle.value = '';
      clearTitleBtn.classList.add('hidden');
      currentPage = 1;
      activeDrilldown = null;
      applyFilters();
      filterTitle.focus();
    });
  }

  // Dispatch Shortcut Listeners
  if (btnDispatchShortcut) {
    btnDispatchShortcut.addEventListener('click', () => {
      isDispatchShortcutActive = !isDispatchShortcutActive;
      const days = parseInt(dispatchDaysInput ? dispatchDaysInput.value : 7, 10) || 7;

      if (isDispatchShortcutActive) {
        btnDispatchShortcut.classList.add('is-active');
        filterLocationType.value = 'STORE';
        showToast(`กรองด่วน "กระจายของ": ย้อนหลัง ${days} วัน (LOCATION_TYPE = STORE)`);
      } else {
        btnDispatchShortcut.classList.remove('is-active');
        showToast('ยกเลิกตัวกรองด่วน "กระจายของ"');
      }

      currentPage = 1;
      activeDrilldown = null;
      applyFilters();
    });
  }

  if (dispatchDaysInput) {
    dispatchDaysInput.addEventListener('change', () => {
      if (isDispatchShortcutActive) {
        currentPage = 1;
        activeDrilldown = null;
        applyFilters();
        const days = parseInt(dispatchDaysInput.value, 10) || 7;
        showToast(`อัปเดตกระจายของย้อนหลัง ${days} วัน`);
      }
    });
    dispatchDaysInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        if (!isDispatchShortcutActive) {
          isDispatchShortcutActive = true;
          btnDispatchShortcut.classList.add('is-active');
          filterLocationType.value = 'STORE';
        }
        currentPage = 1;
        activeDrilldown = null;
        applyFilters();
        const days = parseInt(dispatchDaysInput.value, 10) || 7;
        showToast(`กรองด่วน "กระจายของ": ย้อนหลัง ${days} วัน`);
      }
    });
  }

  resetFiltersBtn.addEventListener('click', () => {
    isDispatchShortcutActive = false;
    if (btnDispatchShortcut) btnDispatchShortcut.classList.remove('is-active');
    if (dispatchDaysInput) dispatchDaysInput.value = '7';

    filterLocationCode.value = '';

    // Reset LOCATION_TYPE to default 'STORE'
    let hasStoreOption = false;
    for (let i = 0; i < filterLocationType.options.length; i++) {
      if (filterLocationType.options[i].value.toUpperCase() === 'STORE') {
        hasStoreOption = true;
        break;
      }
    }
    filterLocationType.value = hasStoreOption ? 'STORE' : '';

    activeGroupPrefix = '';
    activeStoreGroup = '';
    updateGroupButtonStates();
    updateStoreButtonStates();
    updateGroupArticleDropdown('');
    if (filterGroupArticle) filterGroupArticle.value = '';
    filterArticleDesc.value = '';
    if (filterTitle) filterTitle.value = '';

    if (clearLocationBtn) clearLocationBtn.classList.add('hidden');
    clearSearchBtn.classList.add('hidden');
    if (clearTitleBtn) clearTitleBtn.classList.add('hidden');

    currentPage = 1;
    activeDrilldown = null;
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
    const totalPages = Math.ceil(detailRows.length / rowsPerPage) || 1;
    if (currentPage < totalPages) {
      currentPage++;
      renderTablePage();
    }
  });

  btnLastPage.addEventListener('click', () => {
    const totalPages = Math.ceil(detailRows.length / rowsPerPage) || 1;
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
      const isQty = colDef.realName === 'QTY' || colDef.name === 'QTY';
      const alignClass = isQty ? ' style="text-align: right;"' : '';
      thHtml += `<th title="${escapeHtml(colDef.name)}"${alignClass}>${escapeHtml(colDef.name)}</th>`;
    });
    thHtml += '</tr>';
    tableHead.innerHTML = thHtml;
  }

  function renderTablePage() {
    const totalRecords = detailRows.length;
    const totalPages = Math.ceil(totalRecords / rowsPerPage) || 1;

    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;

    const startIdx = (currentPage - 1) * rowsPerPage;
    const endIdx = Math.min(startIdx + rowsPerPage, totalRecords);
    const pageRows = detailRows.slice(startIdx, endIdx);

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
        const isQtyCol = colDef.realName === 'QTY' || colDef.name === 'QTY';
        const isDateCol = colDef.realName === 'CREATED_DATE' || colDef.name === 'Create';
        const isReservedCol = colDef.realName === 'RESERVED_QTY' || colDef.name === 'RESERVED_QTY';
        const isArticleNoCol = colDef.realName === 'ARTICLE_NO' || colDef.name === 'ARTICLE_NO';

        let cellClass = '';
        let formattedVal = '';

        if (val === undefined || val === null || val === '') {
          formattedVal = '<span class="cell-empty">-</span>';
        } else if (isQtyCol) {
          cellClass = 'cell-qty';
          const n = parseFloat(val);
          formattedVal = !isNaN(n) ? n.toLocaleString('th-TH') : escapeHtml(String(val));
        } else if (isDateCol) {
          if (val instanceof Date && !isNaN(val)) {
            formattedVal = escapeHtml(val.toISOString().slice(0, 10));
          } else {
            formattedVal = escapeHtml(String(val));
          }
        } else if (isReservedCol) {
          cellClass = 'cell-reserved';
          const n = parseFloat(val);
          formattedVal = !isNaN(n) ? n.toLocaleString('th-TH') : escapeHtml(String(val));
        } else if (isArticleNoCol) {
          const artDescIdx = colIndexMap['ARTICLE_DESC'];
          const descVal = artDescIdx !== undefined ? String(row[artDescIdx] || '').trim() : '';
          if (descVal) {
            formattedVal = `<span class="article-no-hover" title="${escapeHtml(descVal)}">${escapeHtml(String(val))}</span>`;
          } else {
            formattedVal = escapeHtml(String(val));
          }
        } else if (typeof val === 'number') {
          cellClass = 'cell-numeric';
          formattedVal = val.toLocaleString('th-TH');
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
  // Pivot Summary Table with Dynamic LOCATION_CODE Breakdown
  // (Grouped by: ARTICLE_TYPE -> GROUP_ARTICLE -> ARTICLE_NO)
  // ==========================================

  let pivotTreeData = new Map();
  let activePivotLocCodes = [];

  function buildAndRenderPivotSummary() {
    if (!pivotTableBody) return;

    const articleTypeIdx = colIndexMap['ARTICLE_TYPE'];
    const groupArticleIdx = colIndexMap['GROUP_ARTICLE'];
    const articleNoIdx = colIndexMap['ARTICLE_NO'] !== undefined ? colIndexMap['ARTICLE_NO'] : colIndexMap['ARTICLE_DESC'];
    const articleDescIdx = colIndexMap['ARTICLE_DESC'];
    const locCodeIdx = colIndexMap['LOCATION_CODE'];
    const qtyIdx = colIndexMap['QTY'];

    if (articleTypeIdx === undefined || groupArticleIdx === undefined || articleNoIdx === undefined) {
      pivotTableBody.innerHTML = `
        <tr>
          <td colspan="2" style="text-align: center; padding: 24px; color: var(--text-muted);">
            ARTICLE_TYPE, GROUP_ARTICLE or ARTICLE_NO column not found
          </td>
        </tr>
      `;
      return;
    }

    if (filteredRows.length === 0) {
      if (pivotTableHead) {
        pivotTableHead.innerHTML = `
          <tr>
            <th class="col-article-name">Article NO / หมวดหมู่พัสดุ</th>
            <th class="col-total-qty">QTY</th>
          </tr>
        `;
      }
      pivotTableBody.innerHTML = `
        <tr>
          <td colspan="2" style="text-align: center; padding: 24px; color: var(--text-muted);">
            ไม่พบข้อมูลตามเงื่อนไขตัวกรอง
          </td>
        </tr>
      `;
      pivotTreeData = new Map();
      activePivotLocCodes = [];
      return;
    }

    // 1. Collect distinct location_codes present in filteredRows
    const locSet = new Set();
    if (locCodeIdx !== undefined) {
      for (let i = 0; i < filteredRows.length; i++) {
        const c = String(filteredRows[i][locCodeIdx] || '').trim();
        if (c) locSet.add(c);
      }
    }
    // Sort location columns: 1. HUB -> 2. RNO -> 3. กองงาน -> 4. Other
    const allSortedLocs = Array.from(locSet).sort((a, b) => {
      const rankA = getStoreGroupRank(a);
      const rankB = getStoreGroupRank(b);
      if (rankA !== rankB) {
        return rankA - rankB;
      }
      return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
    });
    const MAX_LOC_COLS = 60; // Safeguard to ensure smooth DOM rendering performance
    const displayLocCols = allSortedLocs.slice(0, MAX_LOC_COLS);
    activePivotLocCodes = displayLocCols;

    // 2. Render Pivot Table Header
    if (pivotTableHead) {
      let thHtml = '<tr><th class="col-article-name">Article NO / หมวดหมู่พัสดุ</th>';
      thHtml += '<th class="col-total-qty" title="ผลรวม QTY ทุกสาขา">QTY</th>';
      displayLocCols.forEach((loc) => {
        const grp = getStoreGroupName(loc);
        const titleStr = grp ? `[${grp}] สาขา: ${escapeHtml(loc)}` : `สาขา: ${escapeHtml(loc)}`;
        thHtml += `<th class="col-loc-qty" title="${titleStr}">${escapeHtml(loc)}</th>`;
      });
      if (allSortedLocs.length > MAX_LOC_COLS) {
        thHtml += `<th class="col-loc-qty" title="มีอีก ${allSortedLocs.length - MAX_LOC_COLS} สาขา">... (+${allSortedLocs.length - MAX_LOC_COLS})</th>`;
      }
      thHtml += '</tr>';
      pivotTableHead.innerHTML = thHtml;
    }

    // 3. Build Aggregation Tree with location breakdowns by ARTICLE_NO
    const tree = new Map();

    for (let i = 0; i < filteredRows.length; i++) {
      const row = filteredRows[i];
      const type = String(row[articleTypeIdx] || '(Unspecified Type)').trim();
      const group = String(row[groupArticleIdx] || '(Unspecified Group)').trim();
      const articleNo = String(row[articleNoIdx] || '(Unspecified Article)').trim();
      const articleDesc = articleDescIdx !== undefined ? String(row[articleDescIdx] || '').trim() : '';
      const loc = locCodeIdx !== undefined ? String(row[locCodeIdx] || '').trim() : '';
      const qty = parseFloat(row[qtyIdx]) || 0;

      if (!tree.has(type)) {
        tree.set(type, { totalQty: 0, locQtys: new Map(), groups: new Map() });
      }
      const typeData = tree.get(type);
      typeData.totalQty += qty;
      if (loc) typeData.locQtys.set(loc, (typeData.locQtys.get(loc) || 0) + qty);

      if (!typeData.groups.has(group)) {
        typeData.groups.set(group, { totalQty: 0, locQtys: new Map(), articles: new Map() });
      }
      const groupData = typeData.groups.get(group);
      groupData.totalQty += qty;
      if (loc) groupData.locQtys.set(loc, (groupData.locQtys.get(loc) || 0) + qty);

      if (!groupData.articles.has(articleNo)) {
        groupData.articles.set(articleNo, { totalQty: 0, locQtys: new Map(), desc: articleDesc });
      }
      const artData = groupData.articles.get(articleNo);
      artData.totalQty += qty;
      if (!artData.desc && articleDesc) artData.desc = articleDesc;
      if (loc) artData.locQtys.set(loc, (artData.locQtys.get(loc) || 0) + qty);
    }

    pivotTreeData = tree;

    // 4. Render Table Rows
    const fragment = document.createDocumentFragment();
    let typeIndex = 0;

    tree.forEach((typeObj, typeName) => {
      typeIndex++;
      const typeId = `pvt-type-${typeIndex}`;

      // Level 1 Row (ARTICLE_TYPE) - no summary totals as requested
      const trL1 = document.createElement('tr');
      trL1.className = 'pivot-row pivot-row-level-1';
      trL1.dataset.typeId = typeId;
      trL1.dataset.level = '1';
      trL1.dataset.typeName = typeName;

      let l1Html = `
        <td class="col-article-name">
          <div class="tree-cell">
            <button type="button" class="tree-toggle-btn" data-target-type="${typeId}" title="ย่อ/ขยาย">−</button>
            <span>${escapeHtml(typeName)}</span>
          </div>
        </td>
        <td class="qty-val-l1 col-total-qty cell-empty">-</td>
      `;
      displayLocCols.forEach((loc) => {
        l1Html += `<td class="qty-val-l1 col-loc-qty cell-empty">-</td>`;
      });
      if (allSortedLocs.length > MAX_LOC_COLS) {
        l1Html += `<td class="col-loc-qty cell-empty">-</td>`;
      }
      trL1.innerHTML = l1Html;
      fragment.appendChild(trL1);

      // Level 2 Rows (GROUP_ARTICLE) - no summary totals
      let groupIndex = 0;
      typeObj.groups.forEach((groupObj, groupName) => {
        groupIndex++;
        const groupId = `${typeId}-grp-${groupIndex}`;

        const trL2 = document.createElement('tr');
        trL2.className = 'pivot-row pivot-row-level-2';
        trL2.dataset.typeId = typeId;
        trL2.dataset.groupId = groupId;
        trL2.dataset.level = '2';
        trL2.dataset.typeName = typeName;
        trL2.dataset.groupName = groupName;

        let l2Html = `
          <td class="col-article-name">
            <div class="tree-cell">
              <span class="tree-indent-1"></span>
              <button type="button" class="tree-toggle-btn" data-target-group="${groupId}" title="ย่อ/ขยาย">−</button>
              <span>${escapeHtml(groupName)}</span>
            </div>
          </td>
          <td class="qty-val-l2 col-total-qty cell-empty">-</td>
        `;
        displayLocCols.forEach((loc) => {
          l2Html += `<td class="qty-val-l2 col-loc-qty cell-empty">-</td>`;
        });
        if (allSortedLocs.length > MAX_LOC_COLS) {
          l2Html += `<td class="col-loc-qty cell-empty">-</td>`;
        }
        trL2.innerHTML = l2Html;
        fragment.appendChild(trL2);

        // Level 3 Rows (ARTICLE_NO)
        groupObj.articles.forEach((artObj, articleNo) => {
          const trL3 = document.createElement('tr');
          trL3.className = 'pivot-row pivot-row-level-3';
          trL3.dataset.typeId = typeId;
          trL3.dataset.groupId = groupId;
          trL3.dataset.level = '3';
          trL3.dataset.typeName = typeName;
          trL3.dataset.groupName = groupName;
          trL3.dataset.articleNo = articleNo;

          const descTooltip = artObj.desc ? ` title="${escapeHtml(artObj.desc)}"` : '';
          let l3Html = `
            <td class="col-article-name">
              <div class="tree-cell">
                <span class="tree-indent-2"></span>
                <span${descTooltip} class="${artObj.desc ? 'article-no-hover' : ''}">${escapeHtml(articleNo)}</span>
              </div>
            </td>
            <td class="qty-val-l3 col-total-qty">
              <span class="qty-clickable" data-drill-level="3" data-type="${escapeHtml(typeName)}" data-group="${escapeHtml(groupName)}" data-article="${escapeHtml(articleNo)}" title="คลิกตัวเลขเพื่อกรอง ${escapeHtml(articleNo)} ทุกสาขา">${formatNumber(artObj.totalQty)}</span>
            </td>
          `;
          displayLocCols.forEach((loc) => {
            const lQty = artObj.locQtys.get(loc) || 0;
            l3Html += `<td class="qty-val-l3 col-loc-qty">`;
            if (lQty > 0) {
              l3Html += `<span class="qty-clickable" data-drill-level="3" data-type="${escapeHtml(typeName)}" data-group="${escapeHtml(groupName)}" data-article="${escapeHtml(articleNo)}" data-loc="${escapeHtml(loc)}" title="คลิกเพื่อกรอง ${escapeHtml(articleNo)} สาขา ${escapeHtml(loc)}">${formatNumber(lQty)}</span>`;
            } else {
              l3Html += `<span class="cell-empty">-</span>`;
            }
            l3Html += `</td>`;
          });
          if (allSortedLocs.length > MAX_LOC_COLS) {
            l3Html += `<td class="col-loc-qty cell-empty">-</td>`;
          }
          trL3.innerHTML = l3Html;
          fragment.appendChild(trL3);
        });
      });
    });

    pivotTableBody.innerHTML = '';
    pivotTableBody.appendChild(fragment);

    attachPivotToggleListeners();
  initQuickStoreButtons();
    updatePivotRowHighlights();
  }

  function attachPivotToggleListeners() {
    // Level 1 toggles
    pivotTableBody.querySelectorAll('.tree-toggle-btn[data-target-type]').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const typeId = btn.dataset.targetType;
        const isExpanded = btn.textContent === '−';

        btn.textContent = isExpanded ? '+' : '−';

        const childRows = pivotTableBody.querySelectorAll(`tr[data-type-id="${typeId}"][data-level="2"], tr[data-type-id="${typeId}"][data-level="3"]`);
        childRows.forEach((r) => {
          if (isExpanded) {
            r.classList.add('pivot-row-hidden');
          } else {
            if (r.dataset.level === '2') {
              r.classList.remove('pivot-row-hidden');
            } else if (r.dataset.level === '3') {
              const grpId = r.dataset.groupId;
              const grpToggle = pivotTableBody.querySelector(`.tree-toggle-btn[data-target-group="${grpId}"]`);
              if (!grpToggle || grpToggle.textContent === '−') {
                r.classList.remove('pivot-row-hidden');
              }
            }
          }
        });
      });
    });

    // Level 2 toggles
    pivotTableBody.querySelectorAll('.tree-toggle-btn[data-target-group]').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const groupId = btn.dataset.targetGroup;
        const isExpanded = btn.textContent === '−';

        btn.textContent = isExpanded ? '+' : '−';

        const childRows = pivotTableBody.querySelectorAll(`tr[data-group-id="${groupId}"][data-level="3"]`);
        childRows.forEach((r) => {
          if (isExpanded) {
            r.classList.add('pivot-row-hidden');
          } else {
            r.classList.remove('pivot-row-hidden');
          }
        });
      });
    });

    // Clickable QTY numbers for drill-down to Detail Table
    pivotTableBody.querySelectorAll('.qty-clickable').forEach((span) => {
      span.addEventListener('click', (e) => {
        e.stopPropagation();
        const level = parseInt(span.dataset.drillLevel, 10);
        const typeName = span.dataset.type || '';
        const groupName = span.dataset.group || '';
        const articleNo = span.dataset.article || '';
        const locationCode = span.dataset.loc || '';

        // If clicked again on the same drilldown, toggle off
        if (
          activeDrilldown &&
          activeDrilldown.level === level &&
          activeDrilldown.typeName === typeName &&
          activeDrilldown.groupName === groupName &&
          activeDrilldown.articleNo === articleNo &&
          (activeDrilldown.locationCode || '') === locationCode
        ) {
          activeDrilldown = null;
          updateDetailTableWithDrilldown(false);
          showToast('ยกเลิกการกรองเจาะจงรายละเอียดแล้ว');
          return;
        }

        activeDrilldown = { level, typeName, groupName, articleNo, locationCode: locationCode || null };
        updateDetailTableWithDrilldown(true);

        let targetDesc = '';
        if (level === 1) targetDesc = `ประเภท: ${typeName}`;
        else if (level === 2) targetDesc = `กลุ่ม: ${groupName}`;
        else if (level === 3) targetDesc = `พัสดุ: ${articleNo}`;
        if (locationCode) {
          targetDesc += ` [สาขา: ${locationCode}]`;
        }
        showToast(`เจาะจงตารางรายละเอียด: ${targetDesc}`);
      });
    });
  }

  // Expand All / Collapse All buttons
  if (btnExpandAll) {
    btnExpandAll.addEventListener('click', () => {
      pivotTableBody.querySelectorAll('tr.pivot-row').forEach((r) => {
        r.classList.remove('pivot-row-hidden');
      });
      pivotTableBody.querySelectorAll('.tree-toggle-btn').forEach((btn) => {
        btn.textContent = '−';
      });
    });
  }

  if (btnCollapseAll) {
    btnCollapseAll.addEventListener('click', () => {
      pivotTableBody.querySelectorAll('tr[data-level="2"], tr[data-level="3"]').forEach((r) => {
        r.classList.add('pivot-row-hidden');
      });
      pivotTableBody.querySelectorAll('.tree-toggle-btn').forEach((btn) => {
        btn.textContent = '+';
      });
    });
  }

  // Export Pivot Summary to CSV (including dynamic LOCATION_CODE columns)
  if (exportPivotCsvBtn) {
    exportPivotCsvBtn.addEventListener('click', exportPivotSummaryToCSV);
  }

  function exportPivotSummaryToCSV() {
    if (!pivotTreeData || pivotTreeData.size === 0) {
      alert('ไม่มีข้อมูลสรุปสำหรับส่งออก');
      return;
    }

    const locCols = activePivotLocCodes || [];
    const headers = ['Article Type', 'Group Article', 'Article NO', 'QTY', ...locCols];
    const csvLines = [headers.map((h) => `"${h.replace(/"/g, '""')}"`).join(',')];

    pivotTreeData.forEach((typeObj, typeName) => {
      typeObj.groups.forEach((groupObj, groupName) => {
        groupObj.articles.forEach((artObj, articleNo) => {
          const row = [
            `"${typeName.replace(/"/g, '""')}"`,
            `"${groupName.replace(/"/g, '""')}"`,
            `"${articleNo.replace(/"/g, '""')}"`,
            artObj.totalQty
          ];
          locCols.forEach((loc) => {
            row.push(artObj.locQtys.get(loc) || 0);
          });
          csvLines.push(row.join(','));
        });
      });
    });

    const csvContent = '\uFEFF' + csvLines.join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    const timestamp = new Date().toISOString().slice(0, 10);
    link.href = url;
    link.download = `ExportPWS_PivotSummary_${timestamp}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    showToast('ส่งออกตารางสรุป Pivot CSV เรียบร้อยแล้ว');
  }

  function formatNumber(num) {
    if (isNaN(num)) return '0';
    if (Math.floor(num) === num) {
      return Number(num).toLocaleString('th-TH');
    }
    return Number(num.toFixed(3)).toLocaleString('th-TH');
  }

  // ==========================================
  // Drill-down Management
  // ==========================================

  function updateDetailTableWithDrilldown(autoScroll = false) {
    const articleTypeIdx = colIndexMap['ARTICLE_TYPE'];
    const groupArticleIdx = colIndexMap['GROUP_ARTICLE'];
    const articleNoIdx = colIndexMap['ARTICLE_NO'] !== undefined ? colIndexMap['ARTICLE_NO'] : colIndexMap['ARTICLE_DESC'];
    const locCodeIdx = colIndexMap['LOCATION_CODE'];

    if (!activeDrilldown) {
      detailRows = [...filteredRows];
      if (drilldownBadgeContainer) drilldownBadgeContainer.classList.add('hidden');
    } else {
      const { level, typeName, groupName, articleNo, locationCode } = activeDrilldown;
      detailRows = filteredRows.filter((row) => {
        let matchHierarchy = true;
        if (level === 1) {
          matchHierarchy = articleTypeIdx !== undefined && String(row[articleTypeIdx] || '').trim() === typeName;
        } else if (level === 2) {
          const matchType = articleTypeIdx === undefined || String(row[articleTypeIdx] || '').trim() === typeName;
          const matchGroup = groupArticleIdx !== undefined && String(row[groupArticleIdx] || '').trim() === groupName;
          matchHierarchy = matchType && matchGroup;
        } else if (level === 3) {
          const matchType = articleTypeIdx === undefined || String(row[articleTypeIdx] || '').trim() === typeName;
          const matchGroup = groupArticleIdx === undefined || String(row[groupArticleIdx] || '').trim() === groupName;
          const matchArt = articleNoIdx !== undefined && String(row[articleNoIdx] || '').trim() === articleNo;
          matchHierarchy = matchType && matchGroup && matchArt;
        }

        if (!matchHierarchy) return false;

        if (locationCode && locCodeIdx !== undefined) {
          const rowLoc = String(row[locCodeIdx] || '').trim();
          if (rowLoc !== locationCode) return false;
        }

        return true;
      });

      if (drilldownBadgeContainer) {
        drilldownBadgeContainer.classList.remove('hidden');
        let badgeText = '';
        if (level === 1) {
          badgeText = `ประเภท: ${typeName}`;
        } else if (level === 2) {
          badgeText = `กลุ่ม: ${groupName}`;
        } else if (level === 3) {
          badgeText = `พัสดุ: ${articleNo}`;
        }
        if (locationCode) {
          badgeText += ` | สาขา: ${locationCode}`;
        }
        if (drilldownLabelText) {
          drilldownLabelText.textContent = badgeText;
        }
      }
    }

    // Update Detail Table badge & Export button count
    if (tableRecordCountBadge) {
      tableRecordCountBadge.textContent = `${detailRows.length.toLocaleString('th-TH')} รายการ`;
    }
    if (exportBtnText) {
      exportBtnText.textContent = `ส่งออกข้อมูลเป็น CSV (${detailRows.length.toLocaleString('th-TH')} รายการ)`;
    }

    updatePivotRowHighlights();

    currentPage = 1;
    renderTableHeader();
    renderTablePage();

    if (autoScroll && detailTableCard) {
      setTimeout(() => {
        detailTableCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }, 50);
    }
  }

  function updatePivotRowHighlights() {
    if (!pivotTableBody) return;
    const allPivotRows = pivotTableBody.querySelectorAll('tr.pivot-row');
    allPivotRows.forEach((r) => r.classList.remove('is-active-drilldown'));
    pivotTableBody.querySelectorAll('.qty-clickable').forEach((el) => el.classList.remove('is-active-cell'));

    if (!activeDrilldown) return;

    const { level, typeName, groupName, articleNo, locationCode } = activeDrilldown;
    allPivotRows.forEach((r) => {
      const rLevel = parseInt(r.dataset.level, 10);
      if (rLevel === level) {
        const matchType = r.dataset.typeName === typeName;
        const matchGroup = !groupName || r.dataset.groupName === groupName;
        const matchArt = !articleNo || r.dataset.articleNo === articleNo;

        let isMatch = false;
        if (level === 1 && matchType) isMatch = true;
        else if (level === 2 && matchType && matchGroup) isMatch = true;
        else if (level === 3 && matchType && matchGroup && matchArt) isMatch = true;

        if (isMatch) {
          r.classList.add('is-active-drilldown');
          const spans = r.querySelectorAll('.qty-clickable');
          spans.forEach((sp) => {
            const spLoc = sp.dataset.loc || '';
            if (locationCode) {
              if (spLoc === locationCode) sp.classList.add('is-active-cell');
            } else {
              if (!spLoc) sp.classList.add('is-active-cell');
            }
          });
        }
      }
    });
  }

  if (clearDrilldownBtn) {
    clearDrilldownBtn.addEventListener('click', () => {
      activeDrilldown = null;
      updateDetailTableWithDrilldown(false);
      showToast('ยกเลิกการกรองเจาะจงรายละเอียดแล้ว');
    });
  }

  // ==========================================
  // Export CSV
  // ==========================================

  function exportFilteredDataToCSV() {
    if (detailRows.length === 0) {
      alert('ไม่มีข้อมูลให้ส่งออก');
      return;
    }

    const colsToExport = (allExportCols && allExportCols.length > 0) ? allExportCols : activeDisplayCols;
    const headers = colsToExport.map((c) => `"${c.name.replace(/"/g, '""')}"`);
    const csvLines = [headers.join(',')];

    detailRows.forEach((row) => {
      const line = colsToExport.map((colDef) => {
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

    showToast(`ส่งออกข้อมูล CSV เรียบร้อยแล้ว (${detailRows.length.toLocaleString('th-TH')} รายการ)`);
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


  // ==========================================
  // Supabase Cloud Storage Sync Integration
  // ==========================================

  let lastUploadedRawBuffer = null;
  let lastUploadedFilename = 'ExportPWS.xlsx';

  const syncModal = document.getElementById('sync-modal');
  const openSyncModalBtn = document.getElementById('open-sync-modal-btn');
  const closeSyncModalBtn = document.getElementById('close-sync-modal-btn');
  const cancelSyncBtn = document.getElementById('cancel-sync-btn');
  const confirmSyncBtn = document.getElementById('confirm-sync-btn');
  const confirmSyncText = document.getElementById('confirm-sync-text');
  const syncStatus = document.getElementById('sync-status');

  if (openSyncModalBtn) {
    openSyncModalBtn.addEventListener('click', openSyncModal);
  }
  if (closeSyncModalBtn) {
    closeSyncModalBtn.addEventListener('click', closeSyncModal);
  }
  if (cancelSyncBtn) {
    cancelSyncBtn.addEventListener('click', closeSyncModal);
  }

  function openSyncModal() {
    if (!syncModal) return;
    if (syncStatus) {
      syncStatus.className = 'sync-status-box hidden';
      syncStatus.textContent = '';
    }
    syncModal.classList.remove('hidden');
  }

  function closeSyncModal() {
    if (syncModal) syncModal.classList.add('hidden');
  }

  if (confirmSyncBtn) {
    confirmSyncBtn.addEventListener('click', handleSyncToSupabase);
  }

  async function handleSyncToSupabase() {
    let bufferToSend = lastUploadedRawBuffer;
    if (!bufferToSend) {
      showSyncStatus('กำลังเตรียมไฟล์ ExportPWS.xlsx สำหรับส่งขึ้น Cloud...', 'status-info');
      try {
        const res = await fetch('ExportPWS.xlsx');
        if (!res.ok) throw new Error('ไม่พบไฟล์ ExportPWS.xlsx บนเซิร์ฟเวอร์');
        bufferToSend = await res.arrayBuffer();
      } catch (e) {
        showSyncStatus('ไม่พบไฟล์ที่จะส่ง: กรุณากดปุ่ม "อัปเดตไฟล์" เพื่อเลือกไฟล์จากเครื่องก่อน', 'status-error');
        return;
      }
    }

    try {
      showSyncStatus('กำลังอัปโหลดไฟล์ไปยัง Supabase Cloud Storage (Singapore)...', 'status-info');
      confirmSyncBtn.disabled = true;
      if (confirmSyncText) confirmSyncText.textContent = 'กำลังส่งข้อมูล...';

      const uploadUrl = `${SUPABASE_PROJECT_URL}/storage/v1/object/${SUPABASE_BUCKET}/${SUPABASE_FILENAME}`;

      const response = await fetch(uploadUrl, {
        method: 'POST',
        headers: {
          'apikey': SUPABASE_ANON_KEY,
          'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'x-upsert': 'true',
        },
        body: bufferToSend,
      });

      if (!response.ok) {
        const resData = await response.json().catch(() => ({}));
        let errMsg = resData.message || resData.error || `HTTP ${response.status}`;
        if (errMsg.includes('row-level security') || errMsg.includes('AccessDenied')) {
          errMsg = 'ยังไม่ได้เปิด Policy ใน Supabase: กรุณารัน SQL Allow Public Uploads ใน SQL Editor ของ Supabase';
        }
        throw new Error(errMsg);
      }

      const nowStr = new Date().toLocaleString('th-TH');
      showSyncStatus(`✅ ส่งขึ้น Supabase สำเร็จแล้ว! (${nowStr}) ทุกคนที่เปิดหรือรีเฟรชหน้าเว็บจะเห็นข้อมูลใหม่ทันที`, 'status-success');
      showToast('🎉 ซิงค์ไฟล์ขึ้น Supabase Cloud สำเร็จแล้ว! ทุกคนเห็นข้อมูลใหม่ทันที');

      setTimeout(() => {
        closeSyncModal();
      }, 2500);

    } catch (err) {
      showSyncStatus(`❌ ข้อผิดพลาด: ${err.message}`, 'status-error');
    } finally {
      confirmSyncBtn.disabled = false;
      if (confirmSyncText) confirmSyncText.textContent = 'ส่งขึ้น Cloud เดี๋ยวนี้ (Real-Time 0 วิ)';
    }
  }

  function showSyncStatus(msg, typeClass) {
    if (!syncStatus) return;
    syncStatus.className = `sync-status-box ${typeClass}`;
    syncStatus.textContent = msg;
    syncStatus.classList.remove('hidden');
  }
