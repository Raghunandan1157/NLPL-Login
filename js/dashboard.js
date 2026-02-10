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
    activeSection: 'branch',
    viewMode: 'card',
    tableCoords: null          // loaded from table_coordinates.json
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
  var grandTotalCard = document.getElementById('grandTotalCard');
  var dataContainer = document.getElementById('dataContainer');
  var loadingOverlay = document.getElementById('loadingOverlay');
  var errorToast = document.getElementById('errorToast');

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

  // --- File upload ---
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
        state.activeSection = 'branch';
        fileNameEl.textContent = file.name;
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

  // --- Render data ---
  function renderData() {
    if (!state.parsedData) return;
    var sheetData = state.parsedData.sheets[state.activeSheet];
    if (!sheetData) return;
    var section = sheetData.sections[state.activeSection];
    if (!section) {
      grandTotalCard.innerHTML = '';
      dataContainer.innerHTML = '<div class="empty-section"><svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke-width="1.5"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg><p>No data for this section</p></div>';
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
          metricGroupHtml('Regular', total.regularDemand, true) +
          metricGroupHtml('1-30 DPD', total.dpd1_30, false) +
          metricGroupHtml('31-60 DPD', total.dpd31_60, false) +
          metricGroupHtml('PNPA', total.pnpa, false) +
          npaGroupHtml(total.npa) +
        '</div>' +
      '</div>';
  }

  function metricGroupHtml(label, data, hasFtod) {
    if (!data) return '';
    var rows = metricRow('Demand', data.demand, false) +
      metricRow('Collection', data.collection, false);
    if (hasFtod) rows += metricRow('FTOD', data.ftod, false);
    else rows += metricRow('Balance', data.balance, false);
    rows += metricRow('Coll%', data.collectionPct, true);
    return '<div class="metric-group"><div class="metric-group-label">' + label + '</div>' + rows + '</div>';
  }

  function npaGroupHtml(data) {
    if (!data) return '';
    return '<div class="metric-group"><div class="metric-group-label">NPA</div>' +
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
          metricGroupHtml('Regular', r.regularDemand, true) +
          metricGroupHtml('1-30 DPD', r.dpd1_30, false) +
          metricGroupHtml('31-60 DPD', r.dpd31_60, false) +
          metricGroupHtml('PNPA', r.pnpa, false) +
          npaGroupHtml(r.npa) +
        '</div>';
      dataContainer.appendChild(card);
    }
    if (rows.length === 0) {
      dataContainer.innerHTML = '<div class="empty-section"><p>No data in this section</p></div>';
    }
  }

  function renderTableFull(section) {
    var groups = [
      { label: 'Regular Demand', cols: ['Demand', 'Coll.', 'FTOD', 'Coll%'] },
      { label: '1-30 DPD', cols: ['Demand', 'Coll.', 'Bal.', 'Coll%'] },
      { label: '31-60 DPD', cols: ['Demand', 'Coll.', 'Bal.', 'Coll%'] },
      { label: 'PNPA', cols: ['Demand', 'Coll.', 'Bal.', 'Coll%'] },
      { label: 'NPA', cols: ['Demand', 'Act.A/c', 'Act.Amt', 'Cls.A/c', 'Cls.Amt'] },
      { label: 'Metrics', cols: ['Rank', 'Perf.'] }
    ];

    var totalCols = 1; // name column
    for (var g = 0; g < groups.length; g++) totalCols += groups[g].cols.length;

    // Group header row
    var groupRow = '<tr><th class="group-header"></th>';
    for (var g = 0; g < groups.length; g++) {
      groupRow += '<th class="group-header" colspan="' + groups[g].cols.length + '">' + groups[g].label + '</th>';
    }
    groupRow += '</tr>';

    // Sub-header row
    var subRow = '<tr><th>Name</th>';
    for (var g = 0; g < groups.length; g++) {
      for (var c = 0; c < groups[g].cols.length; c++) {
        subRow += '<th>' + groups[g].cols[c] + '</th>';
      }
    }
    subRow += '</tr>';

    function rowToTds(r, isTotal) {
      var cls = isTotal ? ' class="total-row"' : '';
      var html = '<tr' + cls + '><td>' + esc(r.name) + '</td>';
      // Regular
      html += '<td>' + formatNum(r.regularDemand.demand) + '</td>';
      html += '<td>' + formatNum(r.regularDemand.collection) + '</td>';
      html += '<td>' + formatNum(r.regularDemand.ftod) + '</td>';
      html += '<td class="' + pctClass(r.regularDemand.collectionPct) + '">' + formatPct(r.regularDemand.collectionPct) + '</td>';
      // 1-30 DPD
      html += '<td>' + formatNum(r.dpd1_30.demand) + '</td>';
      html += '<td>' + formatNum(r.dpd1_30.collection) + '</td>';
      html += '<td>' + formatNum(r.dpd1_30.balance) + '</td>';
      html += '<td class="' + pctClass(r.dpd1_30.collectionPct) + '">' + formatPct(r.dpd1_30.collectionPct) + '</td>';
      // 31-60 DPD
      html += '<td>' + formatNum(r.dpd31_60.demand) + '</td>';
      html += '<td>' + formatNum(r.dpd31_60.collection) + '</td>';
      html += '<td>' + formatNum(r.dpd31_60.balance) + '</td>';
      html += '<td class="' + pctClass(r.dpd31_60.collectionPct) + '">' + formatPct(r.dpd31_60.collectionPct) + '</td>';
      // PNPA
      html += '<td>' + formatNum(r.pnpa.demand) + '</td>';
      html += '<td>' + formatNum(r.pnpa.collection) + '</td>';
      html += '<td>' + formatNum(r.pnpa.balance) + '</td>';
      html += '<td class="' + pctClass(r.pnpa.collectionPct) + '">' + formatPct(r.pnpa.collectionPct) + '</td>';
      // NPA
      html += '<td>' + formatNum(r.npa.demand) + '</td>';
      html += '<td>' + formatNum(r.npa.activationAcct) + '</td>';
      html += '<td>' + formatNum(r.npa.activationAmt) + '</td>';
      html += '<td>' + formatNum(r.npa.closureAcct) + '</td>';
      html += '<td>' + formatNum(r.npa.closureAmt) + '</td>';
      // Metrics
      html += '<td>' + formatNum(r.metrics.rank) + '</td>';
      html += '<td class="' + perfClass(r.metrics.performance) + '">' + esc(r.metrics.performance) + '</td>';
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

})();
