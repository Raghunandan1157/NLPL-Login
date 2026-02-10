/* ================================================
   NLPL — Excel Report Parser
   Parses Daily Collection Report .xlsx/.xlsm files

   Supports two modes:
     1. Dynamic coordinates from table_coordinates.json
        (passed as 2nd argument to parseCollectionReport)
     2. Hardcoded fallback when coordinates are unavailable
   ================================================ */
(function () {
  'use strict';

  // -------------------------------------------------------------------
  // Hardcoded fallback section offsets (1-indexed, matching Excel rows)
  // Used only when table_coordinates.json is NOT available
  // -------------------------------------------------------------------
  var FALLBACK_SECTIONS = {
    region:   { titleRow: 2,   dataStart: 6,   dataEnd: 10,  totalRow: 11  },
    district: { titleRow: 15,  dataStart: 19,  dataEnd: 48,  totalRow: 49  },
    branch:   { titleRow: 53,  dataStart: 57,  dataEnd: 183, totalRow: 184 },
    igl:      { titleRow: 188, dataStart: 194, dataEnd: 198, totalRow: 199 },
    fig:      { titleRow: 203, dataStart: 209, dataEnd: 213, totalRow: 214 },
    il:       { titleRow: 218, dataStart: 224, dataEnd: 228, totalRow: 229 }
  };

  // Mapping from table_coordinates.json table names to our internal keys
  var TABLE_NAME_MAP = {
    region_wise:     'region',
    district_wise:   'district',
    branch_wise:     'branch',
    igl_region_wise: 'igl',
    fig_region_wise: 'fig',
    il_region_wise:  'il'
  };

  // Sheet types: 'full' has 24 columns (B:Y), 'simple' has 4 columns (B:E)
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

  var SECTION_ORDER = ['region', 'district', 'branch', 'igl', 'fig', 'il'];

  // -------------------------------------------------------------------
  // Build per-sheet section definitions from table_coordinates.json
  // Returns an object keyed by sheet name, each value is a SECTIONS map
  // -------------------------------------------------------------------
  function buildSectionsFromCoords(tableCoords) {
    var perSheet = {};
    var sheets = tableCoords && tableCoords.sheets;
    if (!sheets) return null;

    var sheetNames = Object.keys(sheets);
    for (var s = 0; s < sheetNames.length; s++) {
      var sheetName = sheetNames[s];
      var sheetDef = sheets[sheetName];
      if (!sheetDef || !sheetDef.tables) continue;

      var sections = {};
      var tableNames = Object.keys(sheetDef.tables);
      for (var t = 0; t < tableNames.length; t++) {
        var tableName = tableNames[t];
        var sectionKey = TABLE_NAME_MAP[tableName];
        if (!sectionKey) continue; // skip unknown tables (e.g. next_day)

        var tbl = sheetDef.tables[tableName];
        if (!tbl.data_start_row || !tbl.grand_total_row) continue;

        sections[sectionKey] = {
          dataStart: tbl.data_start_row,
          dataEnd:   tbl.data_end_row,
          totalRow:  tbl.grand_total_row
        };
      }

      // Only store if we found at least one section
      if (Object.keys(sections).length > 0) {
        perSheet[sheetName] = sections;
      }
    }

    return Object.keys(perSheet).length > 0 ? perSheet : null;
  }

  // -------------------------------------------------------------------
  // Get the section definitions for a given sheet.
  // Priority: per-sheet coords from JSON > fallback hardcoded values
  // -------------------------------------------------------------------
  function getSectionsForSheet(sheetName, perSheetCoords) {
    if (perSheetCoords && perSheetCoords[sheetName]) {
      // Merge: start with fallback, override with JSON coords per section
      var merged = {};
      for (var i = 0; i < SECTION_ORDER.length; i++) {
        var key = SECTION_ORDER[i];
        if (perSheetCoords[sheetName][key]) {
          merged[key] = perSheetCoords[sheetName][key];
        } else if (FALLBACK_SECTIONS[key]) {
          merged[key] = FALLBACK_SECTIONS[key];
        }
      }
      return merged;
    }
    return FALLBACK_SECTIONS;
  }

  // -------------------------------------------------------------------
  // Value helpers
  // -------------------------------------------------------------------
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

  // -------------------------------------------------------------------
  // Row parsers
  // -------------------------------------------------------------------
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

  // -------------------------------------------------------------------
  // Extract one section from the sheet rows.
  // sectionDef has: dataStart, dataEnd (both 1-indexed inclusive),
  //                 totalRow (1-indexed)
  // allRows is 0-indexed.
  // -------------------------------------------------------------------
  function extractSection(allRows, sectionDef, parseRowFn) {
    var rows = [];
    var startIdx = sectionDef.dataStart - 1;  // convert to 0-indexed
    var endIdx   = sectionDef.dataEnd
                     ? sectionDef.dataEnd - 1   // inclusive end from JSON
                     : sectionDef.totalRow - 2; // fallback: totalRow - 1 in 0-idx, exclusive

    for (var r = startIdx; r <= endIdx; r++) {
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

  // -------------------------------------------------------------------
  // Main entry point
  //   workbook       — XLSX workbook object
  //   tableCoords    — (optional) parsed table_coordinates.json object
  // Returns the same structure as before for dashboard.js compatibility
  // -------------------------------------------------------------------
  function parseCollectionReport(workbook, tableCoords) {
    // Build per-sheet coordinate maps (null if JSON not provided)
    var perSheetCoords = tableCoords ? buildSectionsFromCoords(tableCoords) : null;

    if (perSheetCoords) {
      console.log('[excel-parser] Using table_coordinates.json for row mapping');
    } else {
      console.log('[excel-parser] Using hardcoded fallback row offsets');
    }

    var result = {
      sheets: {},
      sheetOrder: [],
      sheetConfig: SHEET_CONFIG,
      sectionLabels: SECTION_LABELS,
      sectionOrder: SECTION_ORDER
    };

    var expectedSheets = Object.keys(SHEET_CONFIG);
    var foundAny = false;

    for (var i = 0; i < expectedSheets.length; i++) {
      var sheetName = expectedSheets[i];
      if (workbook.SheetNames.indexOf(sheetName) === -1) continue;

      foundAny = true;
      var config = SHEET_CONFIG[sheetName];
      var worksheet = workbook.Sheets[sheetName];
      // Force range from A1 to include empty column A, keeping column indices aligned
      // Without this, SheetJS drops empty leading columns/rows, shifting all indices
      var sheetRange = config.type === 'full' ? 'A1:Y229' : 'A1:E229';
      // For OverAll_On-Date (17663 rows), use the standard range — aux data is beyond row 229
      var allRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', range: sheetRange });
      var parseRowFn = config.type === 'full' ? parseFullRow : parseSimpleRow;

      // Get the section definitions for this specific sheet
      var sections = getSectionsForSheet(sheetName, perSheetCoords);

      var sheetData = {};
      for (var j = 0; j < SECTION_ORDER.length; j++) {
        var secKey = SECTION_ORDER[j];
        if (sections[secKey]) {
          sheetData[secKey] = extractSection(allRows, sections[secKey], parseRowFn);
        }
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

  // Expose globally (same interface as before)
  window.parseCollectionReport = parseCollectionReport;
  window.REPORT_CONFIG = {
    SECTIONS: FALLBACK_SECTIONS,
    SHEET_CONFIG: SHEET_CONFIG,
    SECTION_LABELS: SECTION_LABELS
  };
})();
