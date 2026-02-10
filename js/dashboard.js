/* ================================================
   NLPL — Dashboard Controller
   ================================================ */
(function () {
  'use strict';

  // --- Auth guard ---
  var empId = sessionStorage.getItem('employeeId');
  var empName = sessionStorage.getItem('employeeName');
  var empBranch = sessionStorage.getItem('employeeBranch');

  if (!empId || !empName) {
    window.location.href = 'index.html';
    return;
  }

  // --- State ---
  var state = {
    parsedData: null,
    activeSheet: null,
    activeSection: 'region',
    viewMode: 'table',
    tableCoords: null,          // loaded from table_coordinates.json
    branchRankFilter: null,     // null, 'top10', 'bottom10'
    branchCategoryFilter: null, // null, 'regular', 'dpd1', 'dpd2', 'pnpa'
    rawData: null,              // drill-down data from server
    drillView: null,            // null | 'officers' | 'accounts'
    drillBranch: null,          // branch key in rawData
    drillOfficer: null          // officer name in rawData
  };

  // --- Load table_coordinates.json (best-effort, non-blocking) ---
  (function loadTableCoordinates() {
    fetch('table_coordinates.json')
      .then(function (res) {
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return res.json();
      })
      .then(function (data) {
        state.tableCoords = data;
        console.log('[dashboard] table_coordinates.json loaded successfully');
      })
      .catch(function () {
        state.tableCoords = null;
        console.log('[dashboard] table_coordinates.json not found — parser will use hardcoded offsets');
      });
  })();

  // --- DOM refs ---
  var userName = document.getElementById('userName');
  var userBranch = document.getElementById('userBranch');
  var logoutBtn = document.getElementById('logoutBtn');
  var uploadState = document.getElementById('uploadState');
  var dashboardState = document.getElementById('dashboardState');
  var uploadZone = document.getElementById('uploadZone');
  var uploadBtn = document.getElementById('uploadBtn');
  var fileInput = document.getElementById('fileInput');
  var uploadNewBtn = document.getElementById('uploadNewBtn');
  var fileNameEl = document.getElementById('fileName');
  var sheetTabs = document.getElementById('sheetTabs');
  var sectionTabs = document.getElementById('sectionTabs');
  var viewToggle = document.getElementById('viewToggle');
  var reportTitle = document.getElementById('reportTitle');
  var grandTotalCard = document.getElementById('grandTotalCard');
  var dataContainer = document.getElementById('dataContainer');
  var loadingOverlay = document.getElementById('loadingOverlay');
  var errorToast = document.getElementById('errorToast');

  // Drill-down / raw data DOM refs
  var rawFileInput = document.getElementById('rawFileInput');
  var uploadRawBtn = document.getElementById('uploadRawBtn');
  var processedStatus = document.getElementById('processedStatus');
  var rawStatusEl = document.getElementById('rawStatus');
  var drilldownBreadcrumb = document.getElementById('drilldownBreadcrumb');
  var rawUploadDashBtn = document.getElementById('rawUploadDashBtn');
  var rawDashInput = document.getElementById('rawDashInput');
  var rawDashStatus = document.getElementById('rawDashStatus');

  // --- Init header ---
  userName.textContent = empName;
  userBranch.textContent = empBranch || empId;

  // --- Helpers ---
  function formatNum(val) {
    if (val === null || val === undefined) return '--';
    if (typeof val !== 'number') return String(val);
    // Indian number format
    var s = val.toFixed(0);
    var isNeg = s[0] === '-';
    if (isNeg) s = s.slice(1);
    var lastThree = s.slice(-3);
    var rest = s.slice(0, -3);
    if (rest.length > 0) {
      lastThree = ',' + lastThree;
      rest = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',');
    }
    return (isNeg ? '-' : '') + rest + lastThree;
  }

  function formatPct(val) {
    if (val === null || val === undefined) return '--';
    if (typeof val !== 'number') return String(val);
    var pct = val <= 1 ? val * 100 : val;
    return pct.toFixed(1) + '%';
  }

  function pctClass(val) {
    if (val === null || val === undefined) return '';
    var pct = typeof val === 'number' ? (val <= 1 ? val * 100 : val) : 0;
    if (pct >= 80) return 'pct-good';
    if (pct >= 50) return 'pct-mid';
    return 'pct-bad';
  }

  function perfClass(perf) {
    if (!perf) return 'rank-na';
    if (perf.indexOf('\u25B2') !== -1 || perf.indexOf('Above') !== -1) return 'rank-good';
    if (perf.indexOf('\u25BC') !== -1 || perf.indexOf('Below') !== -1) return 'rank-bad';
    return 'rank-na';
  }

  function esc(str) {
    if (!str) return '';
    return String(str).replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  var FILTER_CATEGORIES = {
    regular: { label: 'Regular Demand vs Collection', dataKey: 'regularDemand', rankKey: 'collectionPct', type: 'pct' },
    dpd1:    { label: '1-30 DPD', dataKey: 'dpd1_30', rankKey: 'collectionPct', type: 'pct' },
    dpd2:    { label: '31-60 DPD', dataKey: 'dpd31_60', rankKey: 'collectionPct', type: 'pct' },
    pnpa:    { label: 'PNPA', dataKey: 'pnpa', rankKey: 'collectionPct', type: 'pct' },
    npa:     { label: 'NPA', dataKey: 'npa', rankKey: 'demand', type: 'demand' }
  };
  var FILTER_CATEGORY_ORDER = ['regular', 'dpd1', 'dpd2', 'pnpa', 'npa'];

  function getFilteredBranches(rows, rankFilter, categoryFilter) {
    var cat = FILTER_CATEGORIES[categoryFilter];
    if (!cat) return [];

    var valid = [];
    for (var i = 0; i < rows.length; i++) {
      var data = rows[i][cat.dataKey];
      if (!data) continue;
      var val = data[cat.rankKey];
      if (val === null || val === undefined || val === '-' || val === '') continue;
      if (typeof val !== 'number') continue;

      var sortVal;
      if (cat.type === 'pct') {
        sortVal = val <= 1 ? val * 100 : val;
        if (sortVal === 0) continue;
      } else {
        sortVal = val;
        if (sortVal === 0) continue;
      }
      valid.push({ idx: i, row: rows[i], sortVal: sortVal });
    }

    valid.sort(function(a, b) { return b.sortVal - a.sortVal; });

    if (rankFilter === 'top10') {
      return valid.slice(0, Math.min(10, valid.length));
    } else if (rankFilter === 'bottom10') {
      var bottom = valid.slice(-Math.min(10, valid.length));
      bottom.sort(function(a, b) { return a.sortVal - b.sortVal; });
      return bottom;
    }
    return [];
  }

  // --- Loading ---
  function showLoading() { loadingOverlay.classList.add('show'); }
  function hideLoading() { loadingOverlay.classList.remove('show'); }

  // --- Error ---
  var errorTimer = null;
  function showError(msg) {
    errorToast.textContent = msg;
    errorToast.classList.add('show');
    clearTimeout(errorTimer);
    errorTimer = setTimeout(function () {
      errorToast.classList.remove('show');
    }, 4000);
  }

  // --- Logout ---
  logoutBtn.addEventListener('click', function () {
    sessionStorage.clear();
    window.location.href = 'index.html';
  });

  // --- File upload (processed report) ---
  uploadBtn.addEventListener('click', function () { fileInput.click(); });
  uploadNewBtn.addEventListener('click', function () { fileInput.click(); });

  // Drag and drop
  uploadZone.addEventListener('dragover', function (e) {
    e.preventDefault();
    uploadZone.style.background = '#E0F5F2';
  });
  uploadZone.addEventListener('dragleave', function () {
    uploadZone.style.background = '';
  });
  uploadZone.addEventListener('drop', function (e) {
    e.preventDefault();
    uploadZone.style.background = '';
    if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
  });

  fileInput.addEventListener('change', function () {
    if (this.files.length > 0) handleFile(this.files[0]);
    this.value = '';
  });

  function handleFile(file) {
    if (!file) return;
    var ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (ext !== '.xlsx' && ext !== '.xlsm') {
      showError('Please upload an Excel file (.xlsx or .xlsm)');
      return;
    }

    showLoading();
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: 'array' });
        var parsed = parseCollectionReport(workbook, state.tableCoords);
        state.parsedData = parsed;
        state.activeSheet = parsed.sheetOrder[0] || 'OverAll';
        state.activeSection = 'region';
        fileNameEl.textContent = file.name;
        var displayName = file.name.replace(/\.(xlsx|xlsm)$/i, '');
        reportTitle.textContent = displayName;
        hideLoading();
        showDashboard();
      } catch (err) {
        hideLoading();
        showError('Could not read this file. Is it a valid Collection Report?');
        console.error('Parse error:', err);
      }
    };
    reader.onerror = function () {
      hideLoading();
      showError('Error reading file. Please try again.');
    };
    reader.readAsArrayBuffer(file);
  }

  // --- Raw data file upload ---
  if (uploadRawBtn) {
    uploadRawBtn.addEventListener('click', function () { rawFileInput.click(); });
  }
  if (rawFileInput) {
    rawFileInput.addEventListener('change', function () {
      if (this.files.length > 0) handleRawFile(this.files[0]);
      this.value = '';
    });
  }
  if (rawUploadDashBtn) {
    rawUploadDashBtn.addEventListener('click', function () { rawDashInput.click(); });
  }
  if (rawDashInput) {
    rawDashInput.addEventListener('change', function () {
      if (this.files.length > 0) handleRawFile(this.files[0]);
      this.value = '';
    });
  }

  function handleRawFile(file) {
    if (!file) return;
    var ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (ext !== '.xlsx') {
      showError('Raw data file must be .xlsx');
      return;
    }

    updateRawStatusUI('uploading', file.name);

    var formData = new FormData();
    formData.append('file', file);

    fetch('/api/upload-raw', {
      method: 'POST',
      body: formData
    })
    .then(function (res) {
      if (!res.ok) throw new Error('Server returned ' + res.status);
      return res.json();
    })
    .then(function (data) {
      if (data.status === 'ok') {
        state.rawData = data;
        updateRawStatusUI('ready', null);
        console.log('[dashboard] Raw data loaded:', data.meta);
        // Re-render if on branch view to make names clickable
        if (state.parsedData && state.activeSection === 'branch' && !state.drillView) {
          renderData();
        }
      } else {
        updateRawStatusUI('error', null);
        showError(data.message || 'Failed to process raw data');
      }
    })
    .catch(function (err) {
      updateRawStatusUI('error', null);
      showError('Raw data upload failed. The file may be too large or the server timed out.');
      console.error('Raw upload error:', err);
    });
  }

  function updateRawStatusUI(status, fileName) {
    var texts = {
      uploading: 'Processing' + (fileName ? ' ' + fileName.substring(0, 20) : '') + '...',
      ready: 'Drill-down ready',
      error: 'Upload failed'
    };
    var text = texts[status] || '';

    if (rawStatusEl) {
      rawStatusEl.textContent = text;
      rawStatusEl.className = 'upload-slot-status ' + status;
    }
    if (rawDashStatus) {
      rawDashStatus.textContent = text;
      rawDashStatus.className = 'raw-dash-status ' + status;
    }
    if (status === 'ready' && rawUploadDashBtn) {
      rawUploadDashBtn.classList.add('raw-ready');
    }
    if (status !== 'ready' && rawUploadDashBtn) {
      rawUploadDashBtn.classList.remove('raw-ready');
    }
  }

  // --- Switch to dashboard view ---
  function showDashboard() {
    uploadState.style.display = 'none';
    dashboardState.style.display = 'block';
    renderSheetTabs();
    renderSectionTabs();
    renderData();
  }

  function showUpload() {
    dashboardState.style.display = 'none';
    uploadState.style.display = 'block';
    state.parsedData = null;
    state.rawData = null;
    state.drillView = null;
    state.drillBranch = null;
    state.drillOfficer = null;
    updateRawStatusUI('', null);
  }

  // --- Sheet tabs ---
  function renderSheetTabs() {
    sheetTabs.innerHTML = '';
    var order = state.parsedData.sheetOrder;
    for (var i = 0; i < order.length; i++) {
      var key = order[i];
      var sheet = state.parsedData.sheets[key];
      var btn = document.createElement('button');
      btn.className = 'sheet-tab' + (key === state.activeSheet ? ' active' : '');
      btn.textContent = sheet.label;
      btn.setAttribute('data-sheet', key);
      btn.addEventListener('click', onSheetTabClick);
      sheetTabs.appendChild(btn);
    }
  }

  function onSheetTabClick() {
    var key = this.getAttribute('data-sheet');
    if (key === state.activeSheet) return;
    state.activeSheet = key;
    state.branchRankFilter = null;
    state.branchCategoryFilter = null;
    state.drillView = null;
    state.drillBranch = null;
    state.drillOfficer = null;
    renderSheetTabs();
    renderSectionTabs();
    renderData();
  }

  // --- Section tabs ---
  function renderSectionTabs() {
    sectionTabs.innerHTML = '';
    var labels = state.parsedData.sectionLabels;
    var order = state.parsedData.sectionOrder;
    for (var i = 0; i < order.length; i++) {
      var key = order[i];
      var btn = document.createElement('button');
      btn.className = 'section-tab' + (key === state.activeSection ? ' active' : '');
      btn.textContent = labels[key];
      btn.setAttribute('data-section', key);
      btn.addEventListener('click', onSectionTabClick);
      sectionTabs.appendChild(btn);
    }
  }

  function onSectionTabClick() {
    var key = this.getAttribute('data-section');
    if (key === state.activeSection) return;
    state.activeSection = key;
    state.branchRankFilter = null;
    state.branchCategoryFilter = null;
    state.drillView = null;
    state.drillBranch = null;
    state.drillOfficer = null;
    var cancelBar = document.getElementById('filterCancelBar');
    if (cancelBar) cancelBar.classList.remove('visible');
    renderSectionTabs();
    renderData();
  }

  // --- View toggle ---
  viewToggle.addEventListener('click', function (e) {
    var btn = e.target.closest('.view-btn');
    if (!btn) return;
    var mode = btn.getAttribute('data-view');
    if (mode === state.viewMode) return;
    state.viewMode = mode;
    var btns = viewToggle.querySelectorAll('.view-btn');
    for (var i = 0; i < btns.length; i++) {
      btns[i].classList.toggle('active', btns[i].getAttribute('data-view') === mode);
    }
    renderData();
  });

  // --- Drill-down helpers ---
  function findRawBranch(name) {
    if (!state.rawData || !state.rawData.branches) return null;
    var branches = state.rawData.branches;
    if (branches[name]) return name;
    var upper = name.toUpperCase().trim();
    for (var key in branches) {
      if (branches.hasOwnProperty(key) && key.toUpperCase().trim() === upper) {
        return key;
      }
    }
    return null;
  }

  function enterBranchDrilldown(branchKey) {
    state.drillView = 'officers';
    state.drillBranch = branchKey;
    state.drillOfficer = null;
    renderData();
  }

  function enterOfficerDrilldown(officerName) {
    state.drillView = 'accounts';
    state.drillOfficer = officerName;
    renderData();
  }

  function drilldownBack() {
    if (state.drillView === 'accounts') {
      state.drillView = 'officers';
      state.drillOfficer = null;
    } else {
      state.drillView = null;
      state.drillBranch = null;
      state.drillOfficer = null;
    }
    renderData();
  }

  function showDrilldownUI() {
    sheetTabs.style.display = 'none';
    sectionTabs.style.display = 'none';
    var viewBar = document.querySelector('.view-bar');
    if (viewBar) viewBar.style.display = 'none';
    reportTitle.style.display = 'none';
    var fp = document.getElementById('branchFilterPanel');
    if (fp) fp.classList.remove('visible');
    dashboardState.classList.remove('has-filter');
    var cb = document.getElementById('filterCancelBar');
    if (cb) cb.classList.remove('visible');
    drilldownBreadcrumb.style.display = 'flex';
  }

  function hideDrilldownUI() {
    sheetTabs.style.display = '';
    sectionTabs.style.display = '';
    var viewBar = document.querySelector('.view-bar');
    if (viewBar) viewBar.style.display = '';
    reportTitle.style.display = '';
    drilldownBreadcrumb.style.display = 'none';
  }

  // --- Render data ---
  function renderData() {
    if (!state.parsedData) return;

    // Drill-down mode
    if (state.drillView) {
      showDrilldownUI();
      if (state.drillView === 'officers') {
        renderOfficerTable();
      } else if (state.drillView === 'accounts') {
        renderAccountTable();
      }
      return;
    }

    // Normal mode
    hideDrilldownUI();

    var sheetData = state.parsedData.sheets[state.activeSheet];
    if (!sheetData) return;
    var section = sheetData.sections[state.activeSection];

    // Render filter panel (only visible for branch)
    renderFilterPanel();
    // Render cancel bar (only visible when filter is active)
    renderCancelBar();
    if (!section) {
      grandTotalCard.innerHTML = '';
      dataContainer.innerHTML = '<div class="empty-section"><svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke-width="1.5"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg><p>No data for this section</p></div>';
      return;
    }

    // Check if branch filter is fully active (both rank + category selected)
    var isFilterActive = state.activeSection === 'branch' &&
      state.branchRankFilter && state.branchCategoryFilter;

    if (isFilterActive) {
      grandTotalCard.innerHTML = '';
      var filtered = getFilteredBranches(section.rows, state.branchRankFilter, state.branchCategoryFilter);
      renderFilteredTable(filtered, state.branchCategoryFilter);
      return;
    }

    if (sheetData.type === 'full') {
      renderGrandTotalFull(section.grandTotal);
      if (state.viewMode === 'card') renderCardsFull(section.rows);
      else renderTableFull(section);
    } else {
      renderGrandTotalSimple(section.grandTotal);
      if (state.viewMode === 'card') renderCardsSimple(section.rows);
      else renderTableSimple(section);
    }
  }

  function ensureFilterPanel() {
    var panel = document.getElementById('branchFilterPanel');
    if (panel) return panel;
    panel = document.createElement('div');
    panel.id = 'branchFilterPanel';
    panel.className = 'filter-panel';
    document.body.appendChild(panel);

    // Single delegated listener — survives innerHTML rebuilds
    panel.addEventListener('click', function (e) {
      var rankBtn = e.target.closest('.filter-rank-btn');
      if (rankBtn) {
        var rank = rankBtn.getAttribute('data-rank');
        state.branchRankFilter = (state.branchRankFilter === rank) ? null : rank;
        renderData();
        return;
      }
      var catBtn = e.target.closest('.filter-cat-btn');
      if (catBtn) {
        var cat = catBtn.getAttribute('data-cat');
        state.branchCategoryFilter = (state.branchCategoryFilter === cat) ? null : cat;
        renderData();
        return;
      }
      var mainBtn = e.target.closest('.filter-main-btn');
      if (mainBtn) {
        state.branchRankFilter = null;
        state.branchCategoryFilter = null;
        var cancelBar = document.getElementById('filterCancelBar');
        if (cancelBar) cancelBar.classList.remove('visible');
        renderData();
      }
    });
    return panel;
  }

  function renderFilterPanel() {
    var panel = ensureFilterPanel();

    if (state.activeSection !== 'branch') {
      panel.classList.remove('visible');
      dashboardState.classList.remove('has-filter');
      return;
    }

    // Add layout class to push content left
    dashboardState.classList.add('has-filter');

    // Make visible with animation
    requestAnimationFrame(function() {
      panel.classList.add('visible');
    });

    var rankActive = state.branchRankFilter;
    var catActive = state.branchCategoryFilter;

    var noFilterActive = !rankActive && !catActive;
    var html = '<div class="filter-panel-title">Filter</div>';
    html += '<div class="filter-section">';
    html += '<button class="filter-main-btn' + (noFilterActive ? ' active' : '') + '">Main Table (No Filter)</button>';
    html += '</div>';
    html += '<div class="filter-section">';
    html += '<div class="filter-section-label">Ranking</div>';
    html += '<button class="filter-rank-btn' + (rankActive === 'top10' ? ' active' : '') + '" data-rank="top10">Top 10 Branches</button>';
    html += '<button class="filter-rank-btn' + (rankActive === 'bottom10' ? ' active' : '') + '" data-rank="bottom10">Bottom 10 Branches</button>';
    html += '</div>';
    html += '<div class="filter-section">';
    html += '<div class="filter-section-label">Report Category</div>';
    for (var i = 0; i < FILTER_CATEGORY_ORDER.length; i++) {
      var key = FILTER_CATEGORY_ORDER[i];
      var cat = FILTER_CATEGORIES[key];
      html += '<button class="filter-cat-btn' + (catActive === key ? ' active' : '') + '" data-cat="' + key + '">' + cat.label + '</button>';
    }
    html += '</div>';
    panel.innerHTML = html;
  }

  function ensureCancelBar() {
    var cancelBar = document.getElementById('filterCancelBar');
    if (cancelBar) return cancelBar;
    cancelBar = document.createElement('div');
    cancelBar.id = 'filterCancelBar';
    cancelBar.className = 'filter-cancel-bar';
    var sectionTabsEl = document.getElementById('sectionTabs');
    sectionTabsEl.parentNode.insertBefore(cancelBar, sectionTabsEl.nextSibling);

    // Single delegated listener — survives innerHTML rebuilds
    cancelBar.addEventListener('click', function (e) {
      if (!e.target.closest('.filter-cancel-btn')) return;
      state.branchRankFilter = null;
      state.branchCategoryFilter = null;
      cancelBar.classList.remove('visible');
      dataContainer.classList.add('table-fade-out');
      setTimeout(function () {
        dataContainer.classList.remove('table-fade-out');
        renderData();
      }, 200);
    });
    return cancelBar;
  }

  function renderCancelBar() {
    var cancelBar = ensureCancelBar();
    var isActive = state.activeSection === 'branch' &&
      state.branchRankFilter && state.branchCategoryFilter;

    if (!isActive) {
      cancelBar.classList.remove('visible');
      return;
    }

    var catLabel = FILTER_CATEGORIES[state.branchCategoryFilter].label;
    var rankLabel = state.branchRankFilter === 'top10' ? 'Top 10' : 'Bottom 10';
    cancelBar.innerHTML = '<span>' + rankLabel + ' \u2014 ' + catLabel + '</span><button class="filter-cancel-btn">Cancel Filter \u2715</button>';

    // Show with animation via class
    requestAnimationFrame(function() {
      cancelBar.classList.add('visible');
    });
  }

  function renderFilteredTable(rows, category) {
    var cat = FILTER_CATEGORIES[category];
    if (!cat) return;

    var isBranch = state.rawData != null;
    var headerRow, bodyHtml = '';

    if (cat.type === 'demand') {
      // NPA: 6 columns — Name, Demand, Act.A/c, Act.Amt, Cls.A/c, Cls.Amt
      headerRow = '<tr><th class="col-name">Branch Name</th><th class="col-demand">Demand</th><th class="col-demand">Act. A/c</th><th class="col-demand">Act. Amt</th><th class="col-collection">Cls. A/c</th><th class="col-collection">Cls. Amt</th></tr>';
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i].row;
        var data = r[cat.dataKey];
        var nameContent = esc(r.name);
        if (isBranch) {
          var rawKey = findRawBranch(r.name);
          if (rawKey) nameContent = '<a class="branch-link" data-branch="' + esc(rawKey) + '">' + esc(r.name) + '</a>';
        }
        bodyHtml += '<tr class="row-slide-in">' +
          '<td class="col-name">' + nameContent + '</td>' +
          '<td class="col-demand">' + formatNum(data.demand) + '</td>' +
          '<td class="col-demand">' + formatNum(data.activationAcct) + '</td>' +
          '<td class="col-demand">' + formatNum(data.activationAmt) + '</td>' +
          '<td class="col-collection">' + formatNum(data.closureAcct) + '</td>' +
          '<td class="col-collection">' + formatNum(data.closureAmt) + '</td>' +
        '</tr>';
      }
    } else {
      // Regular/DPD/PNPA: 4 columns — Name, Demand, Collection, Collection%
      headerRow = '<tr><th class="col-name">Branch Name</th><th class="col-demand">Demand</th><th class="col-collection">Collection</th><th class="col-pct">Collection %</th></tr>';
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i].row;
        var data = r[cat.dataKey];
        var nameContent = esc(r.name);
        if (isBranch) {
          var rawKey = findRawBranch(r.name);
          if (rawKey) nameContent = '<a class="branch-link" data-branch="' + esc(rawKey) + '">' + esc(r.name) + '</a>';
        }
        bodyHtml += '<tr class="row-slide-in">' +
          '<td class="col-name">' + nameContent + '</td>' +
          '<td class="col-demand">' + formatNum(data.demand) + '</td>' +
          '<td class="col-collection">' + formatNum(data.collection) + '</td>' +
          '<td class="col-pct ' + pctClass(data.collectionPct) + '">' + formatPct(data.collectionPct) + '</td>' +
        '</tr>';
      }
    }

    if (rows.length === 0) {
      dataContainer.innerHTML = '<div class="empty-section"><p>No valid data to rank</p></div>';
      return;
    }

    dataContainer.innerHTML = '';
    dataContainer.innerHTML =
      '<div class="table-wrap table-fade-in"><table class="data-table filtered-table"><thead>' +
      headerRow +
      '</thead><tbody>' + bodyHtml + '</tbody></table></div>';
  }

  // --- FULL layout renderers ---
  function metricRow(label, val, isPct) {
    var display = isPct ? formatPct(val) : formatNum(val);
    var cls = isPct ? pctClass(val) : '';
    return '<div class="metric-row"><span class="metric-label">' + label + '</span><span class="metric-value ' + cls + '">' + display + '</span></div>';
  }

  function renderGrandTotalFull(total) {
    if (!total) { grandTotalCard.innerHTML = ''; return; }
    grandTotalCard.innerHTML =
      '<div class="total-card">' +
        '<div class="total-card-label">Grand Total</div>' +
        '<div class="metric-grid">' +
          metricGroupHtml('Regular', total.regularDemand, true, 'metric-regular') +
          metricGroupHtml('1-30 DPD', total.dpd1_30, false, 'metric-dpd1') +
          metricGroupHtml('31-60 DPD', total.dpd31_60, false, 'metric-dpd2') +
          metricGroupHtml('PNPA', total.pnpa, false, 'metric-pnpa') +
          npaGroupHtml(total.npa) +
        '</div>' +
      '</div>';
  }

  function metricGroupHtml(label, data, hasFtod, className) {
    if (!data) return '';
    var cls = className ? 'metric-group ' + className : 'metric-group';
    var rows = metricRow('Demand', data.demand, false) +
      metricRow('Collection', data.collection, false);
    if (hasFtod) rows += metricRow('FTOD', data.ftod, false);
    else rows += metricRow('Balance', data.balance, false);
    rows += metricRow('Coll%', data.collectionPct, true);
    return '<div class="' + cls + '"><div class="metric-group-label">' + label + '</div>' + rows + '</div>';
  }

  function npaGroupHtml(data) {
    if (!data) return '';
    return '<div class="metric-group metric-npa"><div class="metric-group-label">NPA</div>' +
      metricRow('Demand', data.demand, false) +
      metricRow('Act. A/c', data.activationAcct, false) +
      metricRow('Act. Amt', data.activationAmt, false) +
      metricRow('Cls. A/c', data.closureAcct, false) +
      metricRow('Cls. Amt', data.closureAmt, false) +
    '</div>';
  }

  function renderCardsFull(rows) {
    dataContainer.innerHTML = '';
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var rankHtml = '';
      if (r.metrics && r.metrics.performance) {
        var pc = perfClass(r.metrics.performance);
        rankHtml = '<span class="data-card-rank ' + pc + '">' +
          (r.metrics.rank !== null ? '#' + r.metrics.rank : '') +
          ' ' + esc(r.metrics.performance) + '</span>';
      }

      var card = document.createElement('div');
      card.className = 'data-card';
      card.style.animationDelay = (i * 40) + 'ms';
      card.innerHTML =
        '<div class="data-card-header">' +
          '<span class="data-card-name">' + esc(r.name) + '</span>' +
          rankHtml +
        '</div>' +
        '<div class="metric-grid">' +
          metricGroupHtml('Regular', r.regularDemand, true, 'metric-regular') +
          metricGroupHtml('1-30 DPD', r.dpd1_30, false, 'metric-dpd1') +
          metricGroupHtml('31-60 DPD', r.dpd31_60, false, 'metric-dpd2') +
          metricGroupHtml('PNPA', r.pnpa, false, 'metric-pnpa') +
          npaGroupHtml(r.npa) +
        '</div>';
      dataContainer.appendChild(card);
    }
    if (rows.length === 0) {
      dataContainer.innerHTML = '<div class="empty-section"><p>No data in this section</p></div>';
    }
  }

  function renderTableFull(section) {
    var isBranch = state.activeSection === 'branch' && state.rawData != null;

    var groups = [
      { label: 'Regular Demand', cls: 'group-regular', cols: ['Demand', 'Coll.', 'FTOD', 'Coll%'] },
      { label: '1-30 DPD', cls: 'group-dpd1', cols: ['Demand', 'Coll.', 'Bal.', 'Coll%'] },
      { label: '31-60 DPD', cls: 'group-dpd2', cols: ['Demand', 'Coll.', 'Bal.', 'Coll%'] },
      { label: 'PNPA', cls: 'group-pnpa', cols: ['Demand', 'Coll.', 'Bal.', 'Coll%'] },
      { label: 'NPA', cls: 'group-npa', cols: ['Demand', 'Act.A/c', 'Act.Amt', 'Cls.A/c', 'Cls.Amt'] },
      { label: 'Metrics', cls: 'group-metrics', cols: ['Rank', 'Perf.'] }
    ];

    var totalCols = 1; // name column
    for (var g = 0; g < groups.length; g++) totalCols += groups[g].cols.length;

    // Group header row
    var groupRow = '<tr><th class="group-header"></th>';
    for (var g = 0; g < groups.length; g++) {
      groupRow += '<th class="group-header ' + groups[g].cls + '" colspan="' + groups[g].cols.length + '">' + groups[g].label + '</th>';
    }
    groupRow += '</tr>';

    // Sub-header row — column type classes
    var colClasses = [
      ['col-demand', 'col-collection', 'col-balance', 'col-pct'],   // Regular
      ['col-demand', 'col-collection', 'col-balance', 'col-pct'],   // 1-30 DPD
      ['col-demand', 'col-collection', 'col-balance', 'col-pct'],   // 31-60 DPD
      ['col-demand', 'col-collection', 'col-balance', 'col-pct'],   // PNPA
      ['col-demand', 'col-demand', 'col-demand', 'col-collection', 'col-collection'], // NPA
      ['col-rank', 'col-perf']                                      // Metrics
    ];
    var subRow = '<tr><th class="col-name">Name</th>';
    for (var g = 0; g < groups.length; g++) {
      for (var c = 0; c < groups[g].cols.length; c++) {
        subRow += '<th class="' + colClasses[g][c] + '">' + groups[g].cols[c] + '</th>';
      }
    }
    subRow += '</tr>';

    function rowToTds(r, isTotal) {
      var cls = isTotal ? ' class="total-row"' : '';
      var nameContent = esc(r.name);
      if (!isTotal && isBranch) {
        var rawKey = findRawBranch(r.name);
        if (rawKey) {
          nameContent = '<a class="branch-link" data-branch="' + esc(rawKey) + '">' + esc(r.name) + '</a>';
        }
      }
      var html = '<tr' + cls + '><td class="col-name">' + nameContent + '</td>';
      // Regular
      html += '<td class="col-demand">' + formatNum(r.regularDemand.demand) + '</td>';
      html += '<td class="col-collection">' + formatNum(r.regularDemand.collection) + '</td>';
      html += '<td class="col-balance">' + formatNum(r.regularDemand.ftod) + '</td>';
      html += '<td class="col-pct ' + pctClass(r.regularDemand.collectionPct) + '">' + formatPct(r.regularDemand.collectionPct) + '</td>';
      // 1-30 DPD
      html += '<td class="col-demand">' + formatNum(r.dpd1_30.demand) + '</td>';
      html += '<td class="col-collection">' + formatNum(r.dpd1_30.collection) + '</td>';
      html += '<td class="col-balance">' + formatNum(r.dpd1_30.balance) + '</td>';
      html += '<td class="col-pct ' + pctClass(r.dpd1_30.collectionPct) + '">' + formatPct(r.dpd1_30.collectionPct) + '</td>';
      // 31-60 DPD
      html += '<td class="col-demand">' + formatNum(r.dpd31_60.demand) + '</td>';
      html += '<td class="col-collection">' + formatNum(r.dpd31_60.collection) + '</td>';
      html += '<td class="col-balance">' + formatNum(r.dpd31_60.balance) + '</td>';
      html += '<td class="col-pct ' + pctClass(r.dpd31_60.collectionPct) + '">' + formatPct(r.dpd31_60.collectionPct) + '</td>';
      // PNPA
      html += '<td class="col-demand">' + formatNum(r.pnpa.demand) + '</td>';
      html += '<td class="col-collection">' + formatNum(r.pnpa.collection) + '</td>';
      html += '<td class="col-balance">' + formatNum(r.pnpa.balance) + '</td>';
      html += '<td class="col-pct ' + pctClass(r.pnpa.collectionPct) + '">' + formatPct(r.pnpa.collectionPct) + '</td>';
      // NPA
      html += '<td class="col-collection">' + formatNum(r.npa.demand) + '</td>';
      html += '<td class="col-collection">' + formatNum(r.npa.activationAcct) + '</td>';
      html += '<td class="col-collection">' + formatNum(r.npa.activationAmt) + '</td>';
      html += '<td class="col-collection">' + formatNum(r.npa.closureAcct) + '</td>';
      html += '<td class="col-collection">' + formatNum(r.npa.closureAmt) + '</td>';
      // Metrics
      html += '<td class="col-rank">' + formatNum(r.metrics.rank) + '</td>';
      html += '<td class="col-perf ' + perfClass(r.metrics.performance) + '">' + esc(r.metrics.performance) + '</td>';
      html += '</tr>';
      return html;
    }

    var bodyHtml = '';
    for (var i = 0; i < section.rows.length; i++) {
      bodyHtml += rowToTds(section.rows[i], false);
    }
    if (section.grandTotal) {
      bodyHtml += rowToTds(section.grandTotal, true);
    }

    dataContainer.innerHTML = '';
    grandTotalCard.innerHTML = '';
    dataContainer.innerHTML =
      '<div class="table-wrap"><table class="data-table"><thead>' +
      groupRow + subRow +
      '</thead><tbody>' + bodyHtml + '</tbody></table></div>';
  }

  // --- SIMPLE layout renderers (On-Date sheets) ---
  function renderGrandTotalSimple(total) {
    if (!total) { grandTotalCard.innerHTML = ''; return; }
    var pc = pctClass(total.onDate.collectionPct);
    grandTotalCard.innerHTML =
      '<div class="total-card">' +
        '<div class="total-card-label">Grand Total</div>' +
        '<div class="simple-metric-grid">' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Demand</div><div class="simple-metric-value">' + formatNum(total.onDate.demand) + '</div></div>' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Collection</div><div class="simple-metric-value">' + formatNum(total.onDate.collection) + '</div></div>' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Collection %</div><div class="simple-metric-value ' + pc + '">' + formatPct(total.onDate.collectionPct) + '</div></div>' +
        '</div>' +
      '</div>';
  }

  function renderCardsSimple(rows) {
    dataContainer.innerHTML = '';
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var pc = pctClass(r.onDate.collectionPct);
      var card = document.createElement('div');
      card.className = 'data-card';
      card.style.animationDelay = (i * 40) + 'ms';
      card.innerHTML =
        '<div class="data-card-header">' +
          '<span class="data-card-name">' + esc(r.name) + '</span>' +
        '</div>' +
        '<div class="simple-metric-grid">' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Demand</div><div class="simple-metric-value">' + formatNum(r.onDate.demand) + '</div></div>' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Collection</div><div class="simple-metric-value">' + formatNum(r.onDate.collection) + '</div></div>' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Coll%</div><div class="simple-metric-value ' + pc + '">' + formatPct(r.onDate.collectionPct) + '</div></div>' +
        '</div>';
      dataContainer.appendChild(card);
    }
    if (rows.length === 0) {
      dataContainer.innerHTML = '<div class="empty-section"><p>No data in this section</p></div>';
    }
  }

  function renderTableSimple(section) {
    var headerRow = '<tr><th>Name</th><th>Demand</th><th>Collection</th><th>Coll%</th></tr>';

    function rowToTds(r, isTotal) {
      var cls = isTotal ? ' class="total-row"' : '';
      var pc = pctClass(r.onDate.collectionPct);
      return '<tr' + cls + '><td>' + esc(r.name) + '</td>' +
        '<td>' + formatNum(r.onDate.demand) + '</td>' +
        '<td>' + formatNum(r.onDate.collection) + '</td>' +
        '<td class="' + pc + '">' + formatPct(r.onDate.collectionPct) + '</td></tr>';
    }

    var bodyHtml = '';
    for (var i = 0; i < section.rows.length; i++) {
      bodyHtml += rowToTds(section.rows[i], false);
    }
    if (section.grandTotal) {
      bodyHtml += rowToTds(section.grandTotal, true);
    }

    dataContainer.innerHTML = '';
    grandTotalCard.innerHTML = '';
    dataContainer.innerHTML =
      '<div class="table-wrap"><table class="data-table"><thead>' +
      headerRow +
      '</thead><tbody>' + bodyHtml + '</tbody></table></div>';
  }

  // --- Drill-down renderers ---
  function renderOfficerTable() {
    var branchData = state.rawData.branches[state.drillBranch];
    if (!branchData) {
      grandTotalCard.innerHTML = '';
      dataContainer.innerHTML = '<div class="empty-section"><p>No data found for this branch</p></div>';
      return;
    }

    // Breadcrumb
    drilldownBreadcrumb.innerHTML =
      '<button class="drilldown-back-btn">' +
        '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"/></svg>' +
        ' Back to Branch Table' +
      '</button>' +
      '<div class="drilldown-context">Branch: ' + esc(state.drillBranch) + '</div>';

    // Branch totals summary
    var t = branchData.totals;
    var pc = pctClass(t.collection_pct);
    grandTotalCard.innerHTML =
      '<div class="total-card">' +
        '<div class="total-card-label">' + esc(state.drillBranch) + ' \u2014 Summary</div>' +
        '<div class="simple-metric-grid" style="grid-template-columns:repeat(5,1fr)">' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Accounts</div><div class="simple-metric-value">' + formatNum(t.accounts) + '</div></div>' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Demand</div><div class="simple-metric-value">' + formatNum(t.regular_demand) + '</div></div>' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Collection</div><div class="simple-metric-value">' + formatNum(t.collection) + '</div></div>' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Coll%</div><div class="simple-metric-value ' + pc + '">' + t.collection_pct + '%</div></div>' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Officers</div><div class="simple-metric-value">' + Object.keys(branchData.officers).length + '</div></div>' +
        '</div>' +
      '</div>';

    // Officer table
    var officers = branchData.officers;
    var officerNames = Object.keys(officers);

    // Sort by collection% descending
    officerNames.sort(function (a, b) {
      return (officers[b].totals.collection_pct || 0) - (officers[a].totals.collection_pct || 0);
    });

    // Check if account detail is available
    var hasAccounts = officerNames.length > 0 && officers[officerNames[0]] && officers[officerNames[0]].accounts != null;

    var headerRow = '<tr>' +
      '<th class="col-name">Officer Name</th>' +
      '<th>Emp ID</th>' +
      '<th>Accounts</th>' +
      '<th class="col-demand">Demand</th>' +
      '<th class="col-collection">Collection</th>' +
      '<th class="col-pct">Coll%</th>' +
      '<th>0 Days</th><th>1-30</th><th>31-60</th><th>61-90</th><th>90+</th>' +
    '</tr>';

    var bodyHtml = '';
    for (var i = 0; i < officerNames.length; i++) {
      var name = officerNames[i];
      var off = officers[name];
      var ot = off.totals;
      var opc = pctClass(ot.collection_pct);
      var nameCell;
      if (hasAccounts) {
        nameCell = '<a class="officer-link" data-officer="' + esc(name) + '">' + esc(name) + '</a>';
      } else {
        nameCell = esc(name);
      }
      bodyHtml += '<tr class="row-slide-in">' +
        '<td class="col-name">' + nameCell + '</td>' +
        '<td>' + esc(off.officer_id) + '</td>' +
        '<td>' + formatNum(ot.accounts) + '</td>' +
        '<td class="col-demand">' + formatNum(ot.regular_demand) + '</td>' +
        '<td class="col-collection">' + formatNum(ot.collection) + '</td>' +
        '<td class="col-pct ' + opc + '">' + ot.collection_pct + '%</td>' +
        '<td>' + formatNum(ot.dpd_breakdown['0_days']) + '</td>' +
        '<td>' + formatNum(ot.dpd_breakdown['1_30']) + '</td>' +
        '<td>' + formatNum(ot.dpd_breakdown['31_60']) + '</td>' +
        '<td>' + formatNum(ot.dpd_breakdown['61_90']) + '</td>' +
        '<td>' + formatNum(ot.dpd_breakdown['90_plus']) + '</td>' +
      '</tr>';
    }

    if (officerNames.length === 0) {
      dataContainer.innerHTML = '<div class="empty-section"><p>No officers found for this branch</p></div>';
      return;
    }

    dataContainer.innerHTML =
      '<div class="table-wrap table-fade-in"><table class="data-table drilldown-table"><thead>' +
      headerRow +
      '</thead><tbody>' + bodyHtml + '</tbody></table></div>';
  }

  function renderAccountTable() {
    var branchData = state.rawData.branches[state.drillBranch];
    if (!branchData) return;
    var officer = branchData.officers[state.drillOfficer];
    if (!officer) {
      grandTotalCard.innerHTML = '';
      dataContainer.innerHTML = '<div class="empty-section"><p>No data found for this officer</p></div>';
      return;
    }

    // Breadcrumb
    drilldownBreadcrumb.innerHTML =
      '<button class="drilldown-back-btn">' +
        '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"/></svg>' +
        ' Back to Officers' +
      '</button>' +
      '<div class="drilldown-context">Officer: ' + esc(state.drillOfficer) + (officer.officer_id ? ' (' + esc(officer.officer_id) + ')' : '') + '</div>';

    // Officer totals
    var t = officer.totals;
    var pc = pctClass(t.collection_pct);
    grandTotalCard.innerHTML =
      '<div class="total-card">' +
        '<div class="total-card-label">' + esc(state.drillOfficer) + ' \u2014 Summary</div>' +
        '<div class="simple-metric-grid" style="grid-template-columns:repeat(4,1fr)">' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Accounts</div><div class="simple-metric-value">' + formatNum(t.accounts) + '</div></div>' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Demand</div><div class="simple-metric-value">' + formatNum(t.regular_demand) + '</div></div>' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Collection</div><div class="simple-metric-value">' + formatNum(t.collection) + '</div></div>' +
          '<div class="simple-metric-box"><div class="simple-metric-label">Coll%</div><div class="simple-metric-value ' + pc + '">' + t.collection_pct + '%</div></div>' +
        '</div>' +
      '</div>';

    // Account table
    var accounts = officer.accounts;
    if (!accounts || accounts.length === 0) {
      dataContainer.innerHTML = '<div class="empty-section"><p>Account details not available</p></div>';
      return;
    }

    var headerRow = '<tr>' +
      '<th>Account ID</th>' +
      '<th class="col-name">Client Name</th>' +
      '<th>Product</th>' +
      '<th class="col-demand">Loan Amt</th>' +
      '<th>DPD</th>' +
      '<th>Status</th>' +
      '<th class="col-demand">Demand</th>' +
      '<th class="col-collection">Collection</th>' +
      '<th>Partial</th>' +
    '</tr>';

    var bodyHtml = '';
    for (var i = 0; i < accounts.length; i++) {
      var a = accounts[i];
      bodyHtml += '<tr class="row-slide-in">' +
        '<td>' + esc(a.account_id) + '</td>' +
        '<td class="col-name">' + esc(a.client_name) + '</td>' +
        '<td>' + esc(a.product) + '</td>' +
        '<td class="col-demand">' + formatNum(a.loan_amount) + '</td>' +
        '<td>' + formatNum(a.dpd_days) + '</td>' +
        '<td>' + esc(a.status) + '</td>' +
        '<td class="col-demand">' + formatNum(a.regular_demand) + '</td>' +
        '<td class="col-collection">' + formatNum(a.collection) + '</td>' +
        '<td>' + esc(a.partial_amount) + '</td>' +
      '</tr>';
    }

    dataContainer.innerHTML =
      '<div class="table-wrap table-fade-in"><table class="data-table drilldown-table"><thead>' +
      headerRow +
      '</thead><tbody>' + bodyHtml + '</tbody></table></div>';
  }

  // --- Event delegation for drill-down clicks ---
  dataContainer.addEventListener('click', function (e) {
    var branchLink = e.target.closest('.branch-link');
    if (branchLink) {
      e.preventDefault();
      var branch = branchLink.getAttribute('data-branch');
      if (branch) enterBranchDrilldown(branch);
      return;
    }
    var officerLink = e.target.closest('.officer-link');
    if (officerLink) {
      e.preventDefault();
      var officer = officerLink.getAttribute('data-officer');
      if (officer) enterOfficerDrilldown(officer);
      return;
    }
  });

  drilldownBreadcrumb.addEventListener('click', function (e) {
    var backBtn = e.target.closest('.drilldown-back-btn');
    if (backBtn) {
      drilldownBack();
    }
  });

})();
