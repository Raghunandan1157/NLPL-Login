/* ================================================
   NLPL — Excel Report Parser
   Parses Daily Collection Report .xlsx/.xlsm files
   ================================================ */
(function () {
  'use strict';

  // Section row offsets (1-indexed, matching Excel rows)
  // Data rows are between dataStart and totalRow (exclusive)
  var SECTIONS = {
    region:   { titleRow: 2,   dataStart: 6,   totalRow: 11  },
    district: { titleRow: 15,  dataStart: 19,  totalRow: 49  },
    branch:   { titleRow: 53,  dataStart: 57,  totalRow: 184 },
    igl:      { titleRow: 188, dataStart: 194, totalRow: 199 },
    fig:      { titleRow: 203, dataStart: 209, totalRow: 214 },
    il:       { titleRow: 218, dataStart: 224, totalRow: 229 }
  };

  // Sheet types: 'full' has 25 columns, 'simple' has 5 columns
  var SHEET_CONFIG = {
    'OverAll':              { type: 'full',   label: 'Overall' },
    'OverAll_On-Date':      { type: 'simple', label: 'On-Date' },
    'tom_OverAll_On-Date':  { type: 'simple', label: 'Tomorrow' },
    'FY_25-26':             { type: 'full',   label: 'FY 25-26' },
    'FY_25-26_On-Date':     { type: 'simple', label: 'FY On-Date' },
    'tom_FY_25-26_On-Date': { type: 'simple', label: 'FY Tomorrow' }
  };

  var SECTION_LABELS = {
    region: 'Region',
    district: 'District',
    branch: 'Branch',
    igl: 'IGL',
    fig: 'FIG',
    il: 'IL'
  };

  function cleanValue(val) {
    if (val === undefined || val === null || val === '') return null;
    if (typeof val === 'number') return val;
    var s = String(val).trim();
    if (s === '-' || s === '#VALUE!' || s === '#N/A' || s === '#REF!') return null;
    var num = Number(s.replace(/,/g, ''));
    return isNaN(num) ? s : num;
  }

  function formatPct(val) {
    if (val === null || val === undefined) return null;
    if (typeof val === 'number' && val <= 1) return val; // already decimal
    return val;
  }

  function parseFullRow(raw) {
    // raw is 0-indexed array from sheet_to_json header:1
    // Column B(1)=Name, C(2)=Demand... Y(24)=Performance
    return {
      name: String(raw[1] || '').trim(),
      regularDemand: {
        demand:       cleanValue(raw[2]),
        collection:   cleanValue(raw[3]),
        ftod:         cleanValue(raw[4]),
        collectionPct: formatPct(cleanValue(raw[5]))
      },
      dpd1_30: {
        demand:       cleanValue(raw[6]),
        collection:   cleanValue(raw[7]),
        balance:      cleanValue(raw[8]),
        collectionPct: formatPct(cleanValue(raw[9]))
      },
      dpd31_60: {
        demand:       cleanValue(raw[10]),
        collection:   cleanValue(raw[11]),
        balance:      cleanValue(raw[12]),
        collectionPct: formatPct(cleanValue(raw[13]))
      },
      pnpa: {
        demand:       cleanValue(raw[14]),
        collection:   cleanValue(raw[15]),
        balance:      cleanValue(raw[16]),
        collectionPct: formatPct(cleanValue(raw[17]))
      },
      npa: {
        demand:         cleanValue(raw[18]),
        activationAcct: cleanValue(raw[19]),
        activationAmt:  cleanValue(raw[20]),
        closureAcct:    cleanValue(raw[21]),
        closureAmt:     cleanValue(raw[22])
      },
      metrics: {
        rank:        cleanValue(raw[23]),
        performance: String(raw[24] || '').trim()
      }
    };
  }

  function parseSimpleRow(raw) {
    // Column B(1)=Name, C(2)=Demand, D(3)=Collection, E(4)=Collection%
    return {
      name: String(raw[1] || '').trim(),
      onDate: {
        demand:       cleanValue(raw[2]),
        collection:   cleanValue(raw[3]),
        collectionPct: formatPct(cleanValue(raw[4]))
      }
    };
  }

  function extractSection(allRows, sectionDef, parseRowFn) {
    var rows = [];
    // dataStart and totalRow are 1-indexed; allRows is 0-indexed
    for (var r = sectionDef.dataStart - 1; r < sectionDef.totalRow - 1; r++) {
      var raw = allRows[r];
      if (!raw || !raw[1] || String(raw[1]).trim() === '') continue;
      rows.push(parseRowFn(raw));
    }

    var totalRaw = allRows[sectionDef.totalRow - 1];
    var grandTotal = null;
    if (totalRaw && totalRaw[1]) {
      grandTotal = parseRowFn(totalRaw);
    }

    return { rows: rows, grandTotal: grandTotal };
  }

  function parseCollectionReport(workbook) {
    var result = {
      sheets: {},
      sheetOrder: [],
      sheetConfig: SHEET_CONFIG,
      sectionLabels: SECTION_LABELS,
      sectionOrder: ['region', 'district', 'branch', 'igl', 'fig', 'il']
    };

    var expectedSheets = Object.keys(SHEET_CONFIG);
    var foundAny = false;

    for (var i = 0; i < expectedSheets.length; i++) {
      var sheetName = expectedSheets[i];
      if (workbook.SheetNames.indexOf(sheetName) === -1) continue;

      foundAny = true;
      var config = SHEET_CONFIG[sheetName];
      var worksheet = workbook.Sheets[sheetName];
      var allRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
      var parseRowFn = config.type === 'full' ? parseFullRow : parseSimpleRow;

      var sheetData = {};
      var sectionKeys = Object.keys(SECTIONS);
      for (var j = 0; j < sectionKeys.length; j++) {
        var secKey = sectionKeys[j];
        sheetData[secKey] = extractSection(allRows, SECTIONS[secKey], parseRowFn);
      }

      result.sheets[sheetName] = {
        label: config.label,
        type: config.type,
        sections: sheetData
      };
      result.sheetOrder.push(sheetName);
    }

    if (!foundAny) {
      throw new Error('No recognized sheets found. Expected: ' + expectedSheets.join(', '));
    }

    return result;
  }

  // Expose globally
  window.parseCollectionReport = parseCollectionReport;
  window.REPORT_CONFIG = {
    SECTIONS: SECTIONS,
    SHEET_CONFIG: SHEET_CONFIG,
    SECTION_LABELS: SECTION_LABELS
  };
})();
