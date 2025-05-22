/************************************************************
 * script.js (Integrated with fixes for:
 * - Removing unwanted "X" button
 * - Pivot function
 * - Dark mode toggle)
 ************************************************************/

// Flag for fuzzy matching (if needed; currently not used)
const enableFuzzyMatching = false;

/************************************************************
 * HELPER FUNCTIONS
 ************************************************************/

/**
 * Normalizes a cell value by removing non-breaking spaces and trimming.
 */
function normalizeCell(cell) {
  let normalized = String(cell).replace(/\u00A0/g, " ").trim();
  return normalized === "" ? "" : normalized;
}

/**
 * Normalizes a lookup value, removing leading zeros for numeric strings.
 */
// Excel serial date to JS Date
function excelDateToJSDate(serial) {
  const excelEpoch = new Date(1899, 11, 30); // Excel’s “day 0”
  return new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
}

/**
 * Universal date formatter: handles
 *  - Excel serial numbers (number or numeric-string)
 *  - Full mm/dd/yyyy strings (“4/2/2024”)
 *  - Two‑digit‑year strings (“4/25/50”)
 *  - JS Date objects
 */
const RX_FULL_4YR   = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
const RX_TWO_DIGIT  = /^(\d{1,2})\/(\d{1,2})\/(\d{2})$/;
const RX_NUMERIC    = /^\d+$/;

// 2) Replace your old formatDateValue entirely with this optimized version:
function formatDateValue(val) {
  // A) Excel serial (number or numeric‐string)
  if (typeof val === 'number' || (typeof val==='string' && RX_NUMERIC.test(val))) {
    let s = Number(val);
    if (!isNaN(s)) {
      if (s > 59) s--;  // Excel’s bogus Feb‑29‑1900
      const d = new Date((s - 25569) * 86400e3);
      const mm = String(d.getUTCMonth()+1).padStart(2,'0');
      const dd = String(d.getUTCDate()).padStart(2,'0');
      return `${mm}/${dd}/${d.getUTCFullYear()}`;
    }
  }
  // B) Full 4‑digit year string
  if (typeof val==='string') {
    let m = RX_FULL_4YR.exec(val);
    if (m) return m[1].padStart(2,'0') + '/' + m[2].padStart(2,'0') + '/' + m[3];
    // C) Two‑digit year string
    m = RX_TWO_DIGIT.exec(val);
    if (m) {
      const yy = parseInt(m[3],10);
      const yyyy = yy < 50 ? 2000 + yy : 1900 + yy;
      return m[1].padStart(2,'0') + '/' + m[2].padStart(2,'0') + '/' + yyyy;
    }
  }
  // D) Native Date object
  if (val instanceof Date) {
    const mm = String(val.getMonth()+1).padStart(2,'0');
    const dd = String(val.getDate()).padStart(2,'0');
    return `${mm}/${dd}/${val.getFullYear()}`;
  }
  // E) Fallback
  return val == null ? '' : String(val);
}




/**
 * Converts a Date object or date string into mm/dd/yyyy format.
 */
function formatDate(date) {
  const d = new Date(date);
  let month = '' + (d.getMonth() + 1);
  let day = '' + d.getDate();
  const year = d.getFullYear();
  return [month.padStart(2, '0'), day.padStart(2, '0'), year].join('/');
}

/**
 * Computes the blank cell counts for each column in the data.
 */
function computeBlankCellsCount(data, headers) {
  const counts = new Array(headers.length).fill(0);
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    for (let j = 0; j < headers.length; j++) {
      if (row[j] === "") {
        counts[j]++;
      }
    }
  }
  return headers.map((header, index) => [header, counts[index]]);
}

/************************************************************
 * EXPORT PIVOT TABLE TO CSV
 ************************************************************/

/**
 * Exports the pivot result table (class "pvtTable") as it is displayed in the UI.
 * This version constructs a grid that accounts for both colspans and rowspans.
 */
function exportPivotTableToCSV(containerId, filename) {
  const container = document.getElementById(containerId);
  const pivotTable = container.querySelector("table.pvtTable");
  if (!pivotTable) {
    alert("No pivot results found for export.");
    return;
  }
  
  // Build a 2D grid from the table, accounting for rowspans and colspans.
  let grid = [];
  const rows = pivotTable.rows; // HTMLCollection of <tr> elements

  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    // Ensure grid has a row for the current index.
    if (!grid[i]) grid[i] = [];
    let colIndex = 0;
    
    // Loop over all cells in the current row.
    for (let j = 0; j < row.cells.length; j++) {
      // Skip over columns already filled by previous rowspans/colspans.
      while (grid[i][colIndex] !== undefined) {
        colIndex++;
      }
      
      let cell = row.cells[j];
      let text = cell.textContent.normalize("NFKC").trim();
      // Read colspan and rowspan values; default to 1 if not specified.
      let colspan = parseInt(cell.getAttribute("colspan")) || 1;
      let rowspan = parseInt(cell.getAttribute("rowspan")) || 1;
      
      // Place the cell's text in the starting position.
      for (let r = 0; r < rowspan; r++) {
        for (let c = 0; c < colspan; c++) {
          let targetRow = i + r;
          if (!grid[targetRow]) grid[targetRow] = [];
          // Only the top-left cell gets the text; the rest are set as empty strings.
          grid[targetRow][colIndex + c] = (r === 0 && c === 0) ? text : "";
        }
      }
      colIndex += colspan;
    }
  }
  
  // Normalize every row so that all have the same number of columns.
  let maxCols = Math.max(...grid.map(r => r.length));
  grid = grid.map(row => {
    while (row.length < maxCols) {
      row.push("");
    }
    return row;
  });
  
  // Convert the grid to CSV.
  let csv = grid.map(row => 
    row.map(cell => '"' + cell.replace(/"/g, '""') + '"').join(",")
  ).join("\n");
  
  // Create a Blob and trigger a download.
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  if (link.download !== undefined) {
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", filename);
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }
}


/************************************************************
 * LOOKUP & MERGE
 ************************************************************/

/**
 * Merges two tables of data based on matching column values.
 */
function lookupMerge(fromData, fromHeaders, toData, toHeaders, fromLookupIdx, toLookupIdx) {
  let lookupMap = {};
  // Build map from "to" data
  for (let i = 0; i < toData.length; i++) {
    let key = normalizeValue(toData[i][toLookupIdx]);
    lookupMap[key] = toData[i];
  }

  // Prepare merged headers
  const mergedHeaders = fromHeaders.concat(
    toHeaders.filter((_, idx) => idx !== toLookupIdx)
  );

  let matchedRows = [];
  let unmatchedRows = [];

  // Merge row by row
  for (let i = 0; i < fromData.length; i++) {
    let row = fromData[i];
    let key = normalizeValue(row[fromLookupIdx]);
    if (lookupMap.hasOwnProperty(key)) {
      let matchingRow = lookupMap[key];
      let filteredToRow = matchingRow.filter((_, idx) => idx !== toLookupIdx);
      let mergedRow = row.concat(filteredToRow);
      matchedRows.push(mergedRow);
    } else {
      unmatchedRows.push(row);
    }
  }

  return { headers: mergedHeaders, matched: matchedRows, unmatched: unmatchedRows };
}

/**
 * XLOOKUP-style: match on one column, return a single other column
 */
function xlookupMerge(
  fromData, fromHeaders,
  toData,   toHeaders,
  fromLookupIdx, toLookupIdx, toReturnIdx
) {
  const map = {};
  toData.forEach(row => {
    map[ normalizeValue(row[toLookupIdx]) ] = row;
  });

  // merged headers = original + just the return-column header
  const mergedHeaders = fromHeaders.concat([ toHeaders[toReturnIdx] ]);

  const matched   = [];
  const unmatched = [];

  fromData.forEach(row => {
    const key = normalizeValue(row[fromLookupIdx]);
    if (map.hasOwnProperty(key)) {
      matched.push( row.concat([ map[key][toReturnIdx] ]) );
    } else {
      unmatched.push(row);
    }
  });

  return { headers: mergedHeaders, matched, unmatched };
}

/************************************************************
 * FILTER POPOVER HELPER FUNCTIONS
 ************************************************************/

/**
 * Returns the HTML for the popover container (floating).
 * -- REMOVED the "X" button for clearing values (unwanted).
 */
function getPopoverHtml(tableId) {
  return `
    <div id="excelFilterPopover-${tableId}" class="excel-filter-popover" style="display:none;">
      <div class="popover-header"><span class="popover-title">Filter</span></div>
      <div class="popover-search">
        <input type="text" id="filterSearchInput-${tableId}" placeholder="Search…" />
      </div>
      <div class="popover-checklist" id="filterChecklist-${tableId}"></div>
      <div class="popover-actions">
        <button id="selectAllBtn-${tableId}">Select All</button>
        <button id="selectNoneBtn-${tableId}">Select None</button>
        <button id="clearFilterBtn-${tableId}">Clear Filter</button>
        <button id="applyFilterBtn-${tableId}">OK</button>
        <button id="cancelFilterBtn-${tableId}">Cancel</button>
      </div>
    </div>
  `;
}

/**
 * Injects the minimal CSS for the popover (including dark mode).
 * Ensures each table instance has its own style tag only once.
 */
function injectPopoverStyles(tableId) {
  const styleTagId = `excelFilterPopoverStyles-${tableId}`;
  if (document.getElementById(styleTagId)) return;

  const style = document.createElement('style');
  style.id = styleTagId;
  style.innerHTML = `
    /* Filter row buttons */
    .filter-button {
      cursor: pointer;
      font-size: 0.8rem;
      color: #555;
      background: #f2f2f2;
      border: 1px solid #ccc;
      border-radius: 4px;
      padding: 2px 6px;
      margin-left: 2px;
    }
    .filter-button:hover { background: #e1e1e1; }
    .filter-button.filtered {
      background: #ffc107;
      color: #000;
      border: 1px solid #ccc;
    }
    body.dark-mode .filter-button.filtered {
      background: #ffc107 !important;
      color: #000 !important;
      border: 1px solid #555 !important;
    }
    /* Popover styling */
    .excel-filter-popover {
      position: absolute;
      min-width: 220px;
      max-width: 300px;
      background: #fff;
      border: 1px solid #ccc;
      padding: 8px;
      z-index: 9999;
      box-shadow: 0 2px 5px rgba(0,0,0,0.3);
      font-size: 0.9rem;
      overflow: hidden;
      white-space: normal;
    }
    .popover-search { margin-bottom: 6px; }
    .popover-search input {
      width: 100%;
      box-sizing: border-box;
      padding: 4px;
      border-radius: 4px;
      border: 1px solid #ccc;
    }
    .popover-checklist {
      max-height: 160px;
      overflow-y: auto;
      border: 1px solid #eee;
      padding: 4px;
      margin-bottom: 8px;
    }
    body.dark-mode .popover-checklist::-webkit-scrollbar { width: 8px; height: 8px; }
    body.dark-mode .popover-checklist::-webkit-scrollbar-track { background: #333; }
    body.dark-mode .popover-checklist::-webkit-scrollbar-thumb {
        background: #777;
        border-radius: 10px;
        border: 1px solid #666;
    }
    .popover-actions {
      display: flex;
      justify-content: flex-end;
      gap: 8px;
    }
    .popover-actions button {
      background: #f2f2f2;
      color: #555;
      border: 1px solid #ccc;
      border-radius: 4px;
      padding: 2px 6px;
      cursor: pointer;
    }
    .popover-actions button:hover { background: #e1e1e1; }
    body.dark-mode .filter-button {
      background: #424242;
      color: #e0e0e0;
      border: 1px solid #555;
    }
    body.dark-mode .filter-button:hover { background: #555; }
    body.dark-mode .excel-filter-popover {
      background: #1e1e1e;
      border: 1px solid #555;
      color: #e0e0e0;
    }
    body.dark-mode .popover-search input {
      background: #424242;
      color: #e0e0e0;
      border: 1px solid #555;
    }
    body.dark-mode .popover-checklist {
      background: #2e2e2e;
      border: 1px solid #555;
    }
    body.dark-mode .popover-actions button {
      background: #424242;
      color: #e0e0e0;
      border: 1px solid #555;
    }
    body.dark-mode .popover-actions button:hover { background: #555; }
  `;
  document.head.appendChild(style);
}

/**
 * Builds the filter row (one "Filter" button per column) and appends it to the table header.
 */
function buildFilterRow(api) {
  const $thead = $(api.table().header());
  $thead.find('tr.filter-row').remove();

  const $row = $('<tr class="filter-row"></tr>');
  api.columns().every(function(colIdx) {
    const visible = this.visible();
    const $th = $('<th></th>').css('display', visible ? '' : 'none');
    $th.append(`<button class="filter-button" data-colindex="${colIdx}">Filter &#x25BC;</button>`);
    $row.append($th);
  });
  $thead.prepend($row);
}

/************************************************************
 * DATATABLES INITIALIZATION
 ************************************************************/

/**
 * Initializes a DataTable for the given selector, data, and headers.
 * Also sets up the floating popover for Excel-like column filtering.
 */

function initializeTable(selector, data, headers, currentPage) {
  // 1) Destroy any existing DataTable instance
if ($.fn.DataTable.isDataTable(selector)) {
  $(selector).DataTable().clear().destroy();
}

const $table = $(selector);
$table.empty(); // Clear all content

// Rebuild thead and tbody
const newThead = $('<thead><tr></tr></thead>');
headers.forEach(h => {
  newThead.find('tr').append(`<th>${h || 'Unnamed Column'}</th>`);
});
const newTbody = $('<tbody></tbody>');

$table.append(newThead).append(newTbody);


// Rebuild the table structure
const thead = $('<thead><tr></tr></thead>');
headers.forEach(h => {
  thead.find('tr').append(`<th>${h || 'Unnamed Column'}</th>`);
});
const tbody = $('<tbody></tbody>');
$(selector).append(thead).append(tbody);


  // 2) Clear out existing thead/tbody
  $(selector + ' thead').empty();
  $(selector + ' tbody').empty();

  // 3) Determine if ordering should be enabled
  var orderingEnabled = selector.indexOf('blank-count-table') === -1;

  // 4) Build DataTable configuration
  var dtConfig = {
    data: data,
    columns: headers.map(function(h) {
      return { title: h || 'Unnamed Column' };
    }),
    pageLength: 10,
    lengthMenu: [10, 25, 50, 100],
    autoWidth: false,
    deferRender: true,
    scrollY: '400px',
    scrollCollapse: true,
    scroller: true,
    scrollX: true,
    responsive: false,
    searching: true,
    ordering: orderingEnabled,
    dom:
      '<"dt-top-bar"<"dt-search-left"f><"dt-buttons-right"B>>' +
      'rt' +
      '<"dt-bottom-bar"<"dt-length"l><"dt-info"i><"dt-pagination"p>>',
    buttons: [
      { extend: 'copy',  text: '<span>Copy Data</span>' },
      { extend: 'csv',   text: '<span>Export CSV</span>' },
      { extend: 'excel', text: '<span>Export Excel</span>' },
 
      {
        text: 'Clear All Filters',
        className: 'clear-all-filters-btn',
        action: function(e, dt) {
          dt.columns().search('').draw();
          $(dt.table().header()).find('button.filter-button').removeClass('filtered');
        }
      }
    ],

    // 5) initComplete: insert your filter-popover logic here
    initComplete: function() {
      var api     = this.api();
      var tableId = api.table().node().id;
      var columnFilters = {};

      // Inject popover HTML/CSS
      injectPopoverStyles(tableId);
      if (!document.getElementById('excelFilterPopover-' + tableId)) {
        document.body.insertAdjacentHTML('beforeend', getPopoverHtml(tableId));
      }

      // Build initial filter row
      buildFilterRow(api);

      // Keep filter row & Scroller in sync on column-visibility changes
      api.on('column-visibility.dt', function() {
        buildFilterRow(api);
        if (api.scroller) {
          var visCount = api.columns({ visible: true }).count();
          if (visCount === 0) {
            api.scroller.disable();
          } else {
            api.scroller.enable();
            api.scroller.measure();
          }
        }
      });


      // Cache jQuery elements
      const $container = $(api.table().container());
      const $popover   = $(`#excelFilterPopover-${tableId}`);
      const $checklist = $(`#filterChecklist-${tableId}`);
      const $search    = $(`#filterSearchInput-${tableId}`);
      let   currentCol = null;

      // Show popover & populate checklist on filter-button click
      $container.on('click', '.filter-button', function(e) {
        e.stopPropagation();
        e.preventDefault();
        currentCol = +$(this).data('colindex');

        // Extract distinct values from currently filtered rows
        const data = api.rows({ search: 'applied' }).data().toArray();
        const vals = [...new Set(data.map(r => r[currentCol]))].sort();
        const existingPattern = columnFilters[currentCol];

        // Build checklist
        $checklist.empty();
        vals.forEach(v => {
          const label     = v === "" ? "(Blanks)" : v;
          const safeValue = $('<div>').text(v).html();
          const isChecked = existingPattern
            ? new RegExp(existingPattern).test(v)
            : true;
          $checklist.append(`
            <div>
              <label>
                <input type="checkbox"
                       class="check-item"
                       value="${safeValue}"
                       ${isChecked ? 'checked' : ''} />
                ${label}
              </label>
            </div>
          `);
        });

        // Position & show popover
        const off = $(this).offset(), h = $(this).outerHeight();
        $popover.css({ top: off.top + h + 5, left: off.left }).show();

        // Live-search inside popover
        $search.val('').off('input').on('input', function() {
          const term     = $(this).val().toLowerCase();
          const filtered = vals.filter(x => String(x).toLowerCase().includes(term));
          $checklist.empty();
          if (filtered.length === 0) {
            $checklist.append(`<div style="padding:5px;color:#888;">No matches</div>`);
          } else {
            filtered.forEach(vv => {
              const lbl  = vv === "" ? "(Blanks)" : vv;
              const safe = $('<div>').text(vv).html();
              const chk  = existingPattern ? new RegExp(existingPattern).test(vv) : true;
              $checklist.append(`
                <div>
                  <label>
                    <input type="checkbox"
                           class="check-item"
                           value="${safe}"
                           ${chk ? 'checked' : ''} />
                    ${lbl}
                  </label>
                </div>
              `);
            });
          }
        });
      });

      // Popover action buttons
      $(`#selectAllBtn-${tableId}`).on('click', () =>
        $checklist.find('.check-item').prop('checked', true)
      );
      $(`#selectNoneBtn-${tableId}`).on('click', () =>
        $checklist.find('.check-item').prop('checked', false)
      );
      $(`#clearFilterBtn-${tableId}`).on('click', () => {
        if (currentCol === null) return $popover.hide();
        delete columnFilters[currentCol];
        applyAllFilters(); highlightButtons();
        $popover.hide();
      });
      $(`#applyFilterBtn-${tableId}`).on('click', () => {
        if (currentCol === null) return $popover.hide();
        const checked = $checklist.find('.check-item:checked')
                                  .map((_,c) => c.value).get();
        const pattern = checked.length
          ? `^(?:${checked.map(v => $.fn.dataTable.util.escapeRegex(v)).join('|')})$`
          : '(?!)';  // no-match when nothing selected
        columnFilters[currentCol] = pattern;
        applyAllFilters(); highlightButtons();
        $popover.hide();
      });
      $(`#cancelFilterBtn-${tableId}`).on('click', () => $popover.hide());
      $(document).on('click', e => {
        if (!$(e.target).closest($popover).length &&
            !$(e.target).hasClass('filter-button')) {
          $popover.hide();
        }
      });

      // Helper: apply stored filters to all columns and redraw
      function applyAllFilters() {
        api.columns().every(function(idx) { this.search(''); });
        for (let idx in columnFilters) {
          api.column(+idx).search(columnFilters[idx], true, false);
        }
        api.draw();
      }
      // Helper: highlight filter-buttons with active filters
      function highlightButtons() {
        $(api.table().header()).find('.filter-button').each((_,btn) => {
          const idx = +$(btn).data('colindex');
          $(btn).toggleClass('filtered', columnFilters[idx] !== undefined);
        });
      }
    }
  };

  // 6) Instantiate the DataTable
  var table = $(selector).DataTable(dtConfig);

  // 7) If a page was specified, go there
  if (currentPage != null) {
    table.page(currentPage).draw(false);
  }

  return table;
}


/************************************************************
 * CHART CREATION
 ************************************************************/

/**
 * Creates a Chart.js chart (bar, line, pie, etc.) with basic dark mode support.
 */
function createChart(ctx, chartType, labels, values, headerTitles) {
  const darkModeEnabled = document.body.classList.contains('dark-mode');

  function getColorPalette(count, darkMode) {
    const darkPalette = [
      '#ff6384','#36a2eb','#cc65fe','#ffce56',
      '#ffa500','#00ff7f','#ffb6c1','#40e0d0'
    ];
    const lightPalette = [
      '#f44336','#2196f3','#9c27b0','#ffeb3b',
      '#ff9800','#4caf50','#ffc0cb','#40e0d0'
    ];
    const palette = darkMode ? darkPalette : lightPalette;
    let colors = [];
    for (let i = 0; i < count; i++) {
      colors.push(palette[i % palette.length]);
    }
    return colors;
  }

  let backgroundColor, borderColor;
  if (chartType === 'pie' || chartType === 'doughnut') {
    backgroundColor = getColorPalette(values.length, darkModeEnabled);
    borderColor = '#fff';
  } else {
    backgroundColor = getColorPalette(values.length, darkModeEnabled);
    borderColor = backgroundColor;
  }

  const fontColor = darkModeEnabled ? '#e0e0e0' : '#000';
  return new Chart(ctx, {
    type: chartType,
    data: {
      labels: labels,
      datasets: [{
        label: headerTitles[1] || 'Value',
        data: values,
        backgroundColor: backgroundColor,
        borderColor: borderColor,
        borderWidth: 1
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: fontColor } },
        tooltip: { bodyColor: fontColor, titleColor: fontColor },
        datalabels: {
          anchor: 'end',
          align: 'top',
          color: fontColor,
          font: { weight: 'bold' },
          formatter: function(value) { return value; }
        }
      },
      scales: {
        x: {
          ticks: { color: fontColor },
          grid: { color: darkModeEnabled ? '#444' : '#ccc' }
        },
        y: {
          ticks: { color: fontColor },
          grid: { color: darkModeEnabled ? '#444' : '#ccc' }
        }
      }
    },
    plugins: [{
      id: 'customCanvasBackgroundColor',
      beforeDraw: (chart) => {
        const { ctx } = chart;
        ctx.save();
        ctx.fillStyle = darkModeEnabled ? '#121212' : '#ffffff';
        ctx.fillRect(0, 0, chart.width, chart.height);
        ctx.restore();
      }
    }]
  });
}

/************************************************************
 * FILE LOADING & SHEET HANDLING
 ************************************************************/

// Configuration for two "sections" (File 1, File 2).
const sections = [
  {
    dropAreaId: 'drop-area-1',
    fileInputId: 'file-input-1',
    fileNameId: 'file-name-1',
    sheetLabelId: 'sheet-label-1',
    sheetNamesId: 'sheet-names-1',
    chartButtonId: 'generate-chart-1',
    chartContainerId: 'chart-container-1',
    chartCanvasId: 'chart-canvas-1',
    tableId: 'excel-table-1',
    lookupSelectId: 'lookup-select-1',
    lookupButtonId: 'lookup-btn-1',
    state: {}
  },
  {
    dropAreaId: 'drop-area-2',
    fileInputId: 'file-input-2',
    fileNameId: 'file-name-2',
    sheetLabelId: 'sheet-label-2',
    sheetNamesId: 'sheet-names-2',
    chartButtonId: 'generate-chart-2',
    chartContainerId: 'chart-container-2',
    chartCanvasId: 'chart-canvas-2',
    tableId: 'excel-table-2',
    lookupSelectId: 'lookup-select-2',
    lookupButtonId: 'lookup-btn-2',
    state: {}
  }
];

/**
 * Loads an Excel or CSV file, reading it into the config's state object.
 */
function loadExcelFile(file, config) {
  const allowedExtensions = ['xlsx', 'csv'];
  const fileExtension = file.name.split('.').pop().toLowerCase();
  if (!allowedExtensions.includes(fileExtension)) {
    alert("Only XLSX and CSV files are accepted.");
    return;
  }
  const reader = new FileReader();
  reader.onload = function(e) {
    let workbook;
    if (fileExtension === 'csv') {
      workbook = XLSX.read(e.target.result, { type: 'string' });
    } else {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: 'array' });
    }
    config.state.workbook = workbook;

    const fileNameElem = document.getElementById(config.fileNameId);
    fileNameElem.style.display = 'inline';
    fileNameElem.innerHTML = `<br>Selected file:&nbsp;&nbsp;${file.name}`;
    document.getElementById(config.sheetLabelId).style.display = 'block';

    const sheetNamesElem = document.getElementById(config.sheetNamesId);
    sheetNamesElem.style.display = 'flex';
    sheetNamesElem.innerHTML = '';

    workbook.SheetNames.forEach(sheetName => {
      const btn = document.createElement('button');
      btn.textContent = sheetName;
      btn.addEventListener('click', () => {
        sheetNamesElem.querySelectorAll('button').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        loadSheetData(workbook, sheetName, config);
      });
      sheetNamesElem.appendChild(btn);
    });

    // Auto-load the first sheet
    if (workbook.SheetNames.length > 0) {
      const firstBtn = sheetNamesElem.querySelector('button');
      if (firstBtn) {
        firstBtn.classList.add('active');
        loadSheetData(workbook, workbook.SheetNames[0], config);
      }
    }
  };
  fileExtension === 'csv' ? reader.readAsText(file) : reader.readAsArrayBuffer(file);
}

/**
 * Loads a particular sheet from the workbook and initializes a DataTable.
 */
// Helper function to normalize sheet names more aggressively
function normalizeSheetName(name) {
  return name
    .replace(/[\u200B-\u200D\uFEFF]/g, '') // Remove zero-width spaces, etc.
    .replace(/\s+/g, ' ')                 // Replace multiple whitespace with a single space
    .trim()
    .toLowerCase();
}

// Helper function to normalize sheet names more aggressively
function normalizeSheetName(name) {
  return name
    .replace(/[\u200B-\u200D\uFEFF]/g, '') // Remove zero-width spaces, etc.
    .replace(/\s+/g, ' ')                 // Replace multiple whitespace with a single space
    .trim()
    .toLowerCase();
}

function loadSheetData(workbook, sheetName, config) {
  const state = config.state;
  
  // Log available sheet names (normalized) and keys from workbook.Sheets for debugging
  const normalizedSheetNames = workbook.SheetNames.map(s => normalizeSheetName(s));
  console.log("Desired sheet (normalized):", normalizeSheetName(sheetName));
  console.log("Available sheets (normalized):", normalizedSheetNames);
  console.log("Workbook.Sheets keys:", Object.keys(workbook.Sheets));
  
  // Hide lookup/unmatched containers
  document.getElementById('lookup-table-container').style.display = 'none';
  document.getElementById('unmatched-table-container').style.display = 'none';

  // Destroy any existing chart
  if (state.chartInstance) {
    state.chartInstance.destroy();
    state.chartInstance = null;
  }
  document.getElementById(config.chartContainerId).style.display = 'none';

  // Normalize the desired sheet name
  const desiredSheet = normalizeSheetName(sheetName);
  
  // Use a loop to search for a matching sheet key
  let sheetNameToLoad = null;
  for (let i = 0; i < workbook.SheetNames.length; i++) {
    const currentName = workbook.SheetNames[i];
    if (normalizeSheetName(currentName) === desiredSheet) {
      sheetNameToLoad = currentName;
      break;
    }
  }
  
  if (!sheetNameToLoad) {
    console.warn(`Sheet "${sheetName}" was not found. Defaulting to "${workbook.SheetNames[0]}"`);
    sheetNameToLoad = workbook.SheetNames[0];
  }
  
  // Try to get the sheet from the workbook using the key
  const sheet = workbook.Sheets[sheetNameToLoad];
  if (!sheet) {
    alert(`The sheet "${sheetNameToLoad}" was not found in the workbook.`);
    return;
  }
  
  if (!sheet['!ref']) {
    alert(`The "${sheetNameToLoad}" sheet appears to have no data or its range is not defined.`);
    return;
  }

  // Parse the sheet data
  const jsonData = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    dateNF: "mm/dd/yyyy"
  });
  
  if (!jsonData || jsonData.length === 0) {
    alert(`The "${sheetNameToLoad}" sheet is empty.`);
    return;
  }
  
  const headers = jsonData[0];
  if (headers.length === 0) {
    alert(`The "${sheetNameToLoad}" sheet has no headers.`);
    return;
  }
  
  let dataRows = jsonData.slice(1);
  
  // For files with less than 100k rows, filter out rows that are entirely blank.
  if (dataRows.length < 100000) {
    dataRows = dataRows.filter(row => row.some(cell => normalizeCell(cell) !== ""));
  }
  
  // Process each row: format date columns, normalize cells, and ensure the row has enough columns.
  for (let r = 0; r < dataRows.length; r++) {
     for (let c = 0; c < headers.length; c++) {
     if (
        headers[c] &&
       headers[c].toString().toLowerCase().includes("date")
) {
     dataRows[r][c] = formatDateValue(dataRows[r][c]);
}
      dataRows[r][c] = normalizeCell(dataRows[r][c]);
     }
    // Ensure each row has as many columns as headers
    while (dataRows[r].length < headers.length) {
      dataRows[r].push("");
    }
  }

  if (dataRows.length === 0) {
    alert(`All rows in "${sheetNameToLoad}" appear to be blank or whitespace only.`);
  }
  
  // Store the data in state for later use (for tables, pivot, lookup, etc.)
  state.currentTableData = dataRows;
  state.currentTableHeaders = headers;

  // Reset pivot table container so it doesn't show old data
  const sectionNumber = config.tableId.split('-').pop();
  const pivotContainer = document.getElementById('pivot-table-container-' + sectionNumber);
  const pivotOutput = document.getElementById('pivot-output-' + sectionNumber);
  if (pivotContainer && pivotOutput) {
    pivotContainer.style.display = 'none';
    pivotOutput.innerHTML = '';
  }
  
  // Compute blank cells for a separate table
  const blankCounts = computeBlankCellsCount(dataRows, headers);
  document.getElementById('blank-count-container-' + sectionNumber).style.display = 'block';
  initializeTable('#blank-count-table-' + sectionNumber, blankCounts, ["Column", "Blank Cells"]);

  // Main data table
  state.table = initializeTable('#' + config.tableId, dataRows, headers);
  document.getElementById(config.chartButtonId).style.display = 'none';
  document.getElementById('table-container-' + sectionNumber).style.display = 'block';
  document.getElementById('action-controls-' + sectionNumber).style.display = 'inline-flex';

  // Populate the lookup dropdown
// Populate the lookup dropdown (starts blank)
const lookupSelect = document.getElementById(config.lookupSelectId);
lookupSelect.innerHTML = '<option value="">-- Select Lookup Column --</option>';
headers.forEach((header, idx) => {
  const option = document.createElement('option');
  option.value = idx;
  option.textContent = header;
  lookupSelect.appendChild(option);
});

// Populate the return-column dropdown (starts blank)
const returnSelect = document.getElementById(
  config.lookupSelectId.replace('lookup-select', 'return-select')
);
returnSelect.innerHTML = '<option value="">-- Select Return Column --</option>';
headers.forEach((header, idx) => {
  const opt = document.createElement('option');
  opt.value = idx;
  opt.textContent = header;
  returnSelect.appendChild(opt);
});


}

/**
 * Initializes drag/drop and file selection for each section.
 */
function initSection(config) {
  const state = config.state;
  const dropArea = document.getElementById(config.dropAreaId);
  const fileInput = document.getElementById(config.fileInputId);

  dropArea.addEventListener('dragover', e => {
    e.preventDefault();
    dropArea.classList.add('dragover');
  });
  dropArea.addEventListener('dragleave', () => {
    dropArea.classList.remove('dragover');
  });
  dropArea.addEventListener('drop', e => {
    e.preventDefault();
    dropArea.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) loadExcelFile(file, config);
  });
  dropArea.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (file) loadExcelFile(file, config);
  });

  // Generate chart button
  document.getElementById(config.chartButtonId).addEventListener('click', () => {
    if (!state.table) {
      alert("No data table found. Please select a sheet first.");
      return;
    }
    const rawData = state.table.rows({ search: 'applied' }).data().toArray();
    if (rawData.length === 0) {
      alert("No data available to generate chart.");
      return;
    }
    const headerTitles = state.table.settings()[0].aoColumns.map(col => col.sTitle);

    // Aggregate numeric values by label
    const aggregation = {};
    rawData.forEach(row => {
      const label = row[0];
      const value = parseFloat(row[1]);
      if (!isNaN(value)) {
        aggregation[label] = (aggregation[label] || 0) + value;
      }
    });
    const aggLabels = Object.keys(aggregation);
    const aggValues = aggLabels.map(label => aggregation[label]);
    if (aggValues.every(val => val === 0)) {
      alert("No numeric data available for charting.");
      return;
    }

    let recommendedChartType = 'bar';
    if (headerTitles[0] && headerTitles[0].toLowerCase().includes('date')) {
      recommendedChartType = 'line';
    } else if (aggLabels.length <= 4) {
      recommendedChartType = 'pie';
    }

    state.chartData = { labels: aggLabels, values: aggValues };
    state.chartType = recommendedChartType;
    state.chartHeaderTitles = headerTitles;

    document.getElementById(config.chartContainerId).style.display = 'block';
    const ctx = document.getElementById(config.chartCanvasId).getContext('2d');
    if (state.chartInstance) {
      state.chartInstance.destroy();
      state.chartInstance = null;
    }
    state.chartInstance = createChart(ctx, recommendedChartType, aggLabels, aggValues, headerTitles);
  });

  // Lookup button
document.getElementById(config.lookupButtonId).addEventListener('click', () => {
  const src  = config;
  const dest = config.lookupButtonId === 'lookup-btn-1' ? sections[1] : sections[0];

  if (!sections[0].state.currentTableData || !sections[1].state.currentTableData) {
    alert("Both Excel files must be loaded for lookup.");
    return;
  }

  // Retrieve all dropdowns
  // Dropdown references
const fromSelect     = document.getElementById(src.lookupSelectId); // Section 1 - Lookup
const returnSelect   = document.getElementById(src.lookupSelectId.replace('lookup-select', 'return-select')); // Section 1 - Return
const typeSelect     = document.getElementById(src.lookupSelectId.replace('lookup-select', 'lookup-type'));   // Section 1 - Type
const toLookupSelect = document.getElementById(dest.lookupSelectId); // Section 2 - Lookup
const toReturnSelect = document.getElementById(dest.lookupSelectId.replace('lookup-select', 'return-select')); // Section 2 - Return

// Remove all highlights first
[fromSelect, returnSelect, toLookupSelect, toReturnSelect].forEach(el => el.classList.remove('invalid'));

// Get values
const lookupType = typeSelect.value;
const fromIdx = fromSelect.value;
const returnIdx = returnSelect.value;
const toIdx   = toLookupSelect.value;

// Initialize error state
let hasError = false;

// VLOOKUP: Require fromIdx (S1), toIdx (S2)
if (lookupType === 'vlookup') {
  if (fromIdx === "") {
    fromSelect.classList.add('invalid');
    hasError = true;
  }
  if (toIdx === "") {
    toLookupSelect.classList.add('invalid');
    hasError = true;
  }
}

// XLOOKUP: Require fromIdx (S1), returnIdx (S1), toIdx (S2)
if (lookupType === 'xlookup') {
  if (fromIdx === "") {
    fromSelect.classList.add('invalid');
    hasError = true;
  }
  if (returnIdx === "") {
    returnSelect.classList.add('invalid');
    hasError = true;
  }
  if (toIdx === "") {
    toLookupSelect.classList.add('invalid');
    hasError = true;
  }
}

// Optional: Show prompt
if (hasError) {
  alert("Please select all required fields for " + lookupType.toUpperCase() + ".");
  return;
}

  // VALIDATION ENDS

  let result;
  if (lookupType === 'vlookup') {
    // VLOOKUP: merge entire row
    result = lookupMerge(
      src.state.currentTableData,
      src.state.currentTableHeaders,
      dest.state.currentTableData,
      dest.state.currentTableHeaders,
      +fromIdx, +toIdx
    );
  } else {
    // XLOOKUP: return single column
    const returnIdx = +returnSelect.value;
    result = xlookupMerge(
      src.state.currentTableData,
      src.state.currentTableHeaders,
      dest.state.currentTableData,
      dest.state.currentTableHeaders,
      +fromIdx, +toIdx, returnIdx
    );
  }

  // Display matched & unmatched results
  document.getElementById('lookup-table-container').style.display = 'block';
  initializeTable('#lookup-table', result.matched, result.headers);

  document.getElementById('unmatched-table-container').style.display = 'block';
  initializeTable('#unmatched-table', result.unmatched, src.state.currentTableHeaders);
});


}

/************************************************************
 * PIVOT TABLE INITIALIZATION
 ************************************************************/

function initializePivot(pivotOutput, dataObjects) {
  // Remove any existing pivotUI content by replacing the container with a new element.
  const parent = pivotOutput.parentNode;
  const newPivotOutput = document.createElement("div");
  newPivotOutput.id = pivotOutput.id; // preserve the id
  parent.replaceChild(newPivotOutput, pivotOutput);
  // Initialize pivotUI on the fresh container
$(newPivotOutput).pivotUI(dataObjects, {
  rows: [],
  cols: [],
  vals: [],
  rendererName: "Table",
  aggregatorName: "Count",
  overwrite: true,
  onRefresh: function(config) {
    // Move the pivot filter popup inside the pivot container.
    var container = $(this).closest('.pivot-container');
    setTimeout(function(){
      var filterBox = $('.pvtFilterBox');
      if(filterBox.length > 0){
        filterBox.appendTo(container);
      }
    }, 50);
  }
});

  return newPivotOutput;
}

// Pivot for section 1
document.getElementById("pivot-btn-1").addEventListener("click", () => {
  const sectionNumber = 1; // Because pivot-btn-1 => file #1
  const state = sections[0].state;
  if (!state.currentTableData || !state.currentTableHeaders) {
    alert("Load a file first to generate a pivot table.");
    return;
  }

  const pivotContainer = document.getElementById(`pivot-table-container-${sectionNumber}`);
  let pivotOutput = document.getElementById(`pivot-output-${sectionNumber}`);
  pivotContainer.style.display = "block";
  pivotOutput.innerHTML = ""; // Clear existing content

  // Convert row arrays into objects keyed by headers
  const dataObjects = state.currentTableData.map(row => {
    let obj = {};
    state.currentTableHeaders.forEach((header, idx) => {
      obj[header] = row[idx];
    });
    return obj;
  });

  // Initialize the pivotUI for the first time
  $(pivotOutput).pivotUI(dataObjects, {
    rows: [],
    cols: [],
    vals: [],
    rendererName: "Table",
    aggregatorName: "Count",
    onRefresh: function(config) {
      // Handle config changes if needed
    }
  });

  // Remove any existing pivot actions to avoid duplicates
  const existingActions = pivotContainer.querySelector('.pivot-actions');
  if (existingActions) {
    existingActions.remove();
  }

// Add action buttons
let actionsDiv = document.createElement('div');
actionsDiv.className = 'pivot-actions';
actionsDiv.style.marginBottom = '10px';
actionsDiv.innerHTML = `
  <button class="dt-button" id="clear-pivot-filters-${sectionNumber}" style="margin-right:10px;">Clear All Filters</button>
  <button class="dt-button" id="export-csv-pivot-${sectionNumber}">Export CSV</button>
`;
pivotContainer.insertBefore(actionsDiv, pivotOutput);

  // Clear All Filters button: replace the pivot output container with a fresh one and reinitialize pivotUI
  document.getElementById(`clear-pivot-filters-${sectionNumber}`).addEventListener('click', () => {
    // Replace the pivot output container with a fresh clone and reinitialize pivotUI
    pivotOutput = initializePivot(pivotOutput, dataObjects);
  });

  // Export CSV button: export the current pivot table
  document.getElementById(`export-csv-pivot-${sectionNumber}`).addEventListener('click', () => {
    exportPivotTableToCSV(`pivot-table-container-${sectionNumber}`, `pivot-table-${sectionNumber}.csv`);
  });
});

// Pivot for section 2
document.getElementById("pivot-btn-2").addEventListener("click", () => {
  const sectionNumber = 2; // Because pivot-btn-2 => file #2
  const state = sections[1].state;
  if (!state.currentTableData || !state.currentTableHeaders) {
    alert("Load a file first to generate a pivot table.");
    return;
  }

  const pivotContainer = document.getElementById(`pivot-table-container-${sectionNumber}`);
  let pivotOutput = document.getElementById(`pivot-output-${sectionNumber}`);
  pivotContainer.style.display = "block";
  pivotOutput.innerHTML = "";

  // Convert row arrays into objects keyed by headers
  const dataObjects = state.currentTableData.map(row => {
    let obj = {};
    state.currentTableHeaders.forEach((header, idx) => {
      obj[header] = row[idx];
    });
    return obj;
  });

  // Initialize pivotUI
  $(pivotOutput).pivotUI(dataObjects, {
    rows: [],
    cols: [],
    vals: [],
    rendererName: "Table",
    aggregatorName: "Count",
    onRefresh: function(config) {
      // Handle config changes if needed
    }
  });

  // Remove any existing pivot actions to avoid duplicates
  const existingActions = pivotContainer.querySelector('.pivot-actions');
  if (existingActions) {
    existingActions.remove();
  }

// Add action buttons
let actionsDiv = document.createElement('div');
actionsDiv.className = 'pivot-actions';
actionsDiv.style.marginBottom = '10px';
actionsDiv.innerHTML = `
  <button class="dt-button" id="clear-pivot-filters-${sectionNumber}" style="margin-right:10px;">Clear All Filters</button>
  <button class="dt-button" id="export-csv-pivot-${sectionNumber}">Export CSV</button>
`;
pivotContainer.insertBefore(actionsDiv, pivotOutput);

  // Clear All Filters button
  document.getElementById(`clear-pivot-filters-${sectionNumber}`).addEventListener('click', () => {
    pivotOutput = initializePivot(pivotOutput, dataObjects);
  });

  // Export CSV button
  document.getElementById(`export-csv-pivot-${sectionNumber}`).addEventListener('click', () => {
    exportPivotTableToCSV(`pivot-table-container-${sectionNumber}`, `pivot-table-${sectionNumber}.csv`);
  });
});


/************************************************************
 * INITIALIZATION
 ************************************************************/

sections.forEach(initSection);

/**
 * Add a dark-mode toggle event to the "darkModeToggle" button.
 * Make sure you have a button with id="darkModeToggle" in your HTML.
 */
document.getElementById("darkModeToggle").addEventListener("click", () => {
  document.body.classList.toggle("dark-mode");
});
