(function () {
    'use strict';

    let DATA = null;
    let activeSheet = null;
    let selectedCell = null;

    // DOM refs
    const gridViewport = document.getElementById('gridViewport');
    const colHeaderTrack = document.getElementById('colHeaderTrack');
    const rowHeaderTrack = document.getElementById('rowHeaderTrack');
    const colHeaders = document.getElementById('colHeaders');
    const rowHeaders = document.getElementById('rowHeaders');
    const dataTable = document.getElementById('dataTable');
    const sheetTabs = document.getElementById('sheetTabs');
    const sheetTabsScroll = document.getElementById('sheetTabsScroll');
    const cellRef = document.getElementById('cellRef');
    const formulaDisplay = document.getElementById('formulaDisplay');
    const searchInput = document.getElementById('searchInput');
    const searchCount = document.getElementById('searchCount');
    const statusText = document.getElementById('statusText');
    const loading = document.getElementById('loading');

    // Column letter helper (A, B, ... Z, AA, AB, ...)
    function colLetter(index) {
        let s = '';
        let n = index;
        while (n >= 0) {
            s = String.fromCharCode(65 + (n % 26)) + s;
            n = Math.floor(n / 26) - 1;
        }
        return s;
    }

    // Format cell value for display
    function formatCell(val) {
        if (val === '' || val === null || val === undefined) return '';
        if (typeof val === 'number') {
            // Detect percentages (0-1 range stored as decimals)
            if (Math.abs(val) <= 3 && val !== Math.round(val) && Math.abs(val) < 1) {
                return (val * 100).toFixed(1) + '%';
            }
            // Large numbers with commas
            if (Math.abs(val) >= 1000) {
                return val.toLocaleString('en-US', { maximumFractionDigits: 1 });
            }
            if (val !== Math.round(val)) {
                return val.toFixed(2);
            }
            return val.toLocaleString('en-US');
        }
        return String(val);
    }

    // Determine cell class
    function cellClass(val, colIdx, headerRow) {
        if (val === '' || val === null || val === undefined) return 'empty-cell';
        if (typeof val === 'number') {
            if (Math.abs(val) <= 3 && val !== Math.round(val) && Math.abs(val) < 1) {
                return val < 0 ? 'pct negative' : 'pct';
            }
            return 'num';
        }
        return '';
    }

    // Detect header row index (row with the column names)
    function findHeaderRow(rows) {
        for (let i = 0; i < Math.min(rows.length, 6); i++) {
            const nonEmpty = rows[i].filter(v => v !== '').length;
            if (nonEmpty >= 3 && rows[i].every(v => typeof v === 'string' || v === '')) {
                // Check if next row has different types (data row)
                if (i + 1 < rows.length) {
                    const nextNonEmpty = rows[i + 1].filter(v => v !== '');
                    if (nextNonEmpty.some(v => typeof v === 'number')) {
                        return i;
                    }
                }
            }
        }
        return -1;
    }

    function renderSheet(name) {
        activeSheet = name;
        const sheet = DATA[name];
        const rows = sheet.rows;
        const maxCol = sheet.maxCol;

        // Find header row
        const headerRowIdx = findHeaderRow(rows);

        // Build column headers
        colHeaders.innerHTML = '';
        for (let c = 0; c < maxCol; c++) {
            const div = document.createElement('div');
            div.className = 'col-header';
            div.textContent = colLetter(c);
            div.style.width = 'var(--cell-min-w)';
            colHeaders.appendChild(div);
        }

        // Build row headers
        rowHeaders.innerHTML = '';
        for (let r = 0; r < rows.length; r++) {
            const div = document.createElement('div');
            div.className = 'row-header';
            div.textContent = r + 1;
            rowHeaders.appendChild(div);
        }

        // Build table
        dataTable.innerHTML = '';
        const fragment = document.createDocumentFragment();

        for (let r = 0; r < rows.length; r++) {
            const tr = document.createElement('tr');
            const row = rows[r];

            // Determine row type
            const isTitle = r === 0 && row.filter(v => v !== '').length <= 2;
            const isSubtitle = r === 1 && row.filter(v => v !== '').length <= 2;
            const isHeader = r === headerRowIdx;
            const isEmpty = row.every(v => v === '');

            for (let c = 0; c < maxCol; c++) {
                const td = document.createElement('td');
                const val = c < row.length ? row[c] : '';

                if (isTitle && c === 0) {
                    td.className = 'title-cell';
                    td.textContent = val;
                    td.colSpan = maxCol;
                    tr.appendChild(td);
                    break;
                }
                if (isSubtitle && c === 0) {
                    td.className = 'subtitle-cell';
                    td.textContent = val;
                    td.colSpan = maxCol;
                    tr.appendChild(td);
                    break;
                }
                if (isEmpty) {
                    td.className = 'empty-cell';
                    td.colSpan = maxCol;
                    tr.appendChild(td);
                    break;
                }

                if (isHeader) {
                    td.className = 'header-cell';
                    td.textContent = val;
                } else {
                    const cls = cellClass(val, c, isHeader);
                    if (cls) td.className = cls;
                    td.textContent = formatCell(val);
                }

                td.dataset.row = r;
                td.dataset.col = c;
                td.title = String(val);
                tr.appendChild(td);
            }
            fragment.appendChild(tr);
        }
        dataTable.appendChild(fragment);

        // Reset scroll
        gridViewport.scrollTop = 0;
        gridViewport.scrollLeft = 0;

        // Update status
        const dataRows = Math.max(0, rows.length - (headerRowIdx + 1));
        statusText.textContent = `${rows.length} rows × ${maxCol} cols`;

        // Update active tab
        document.querySelectorAll('.sheet-tab').forEach(t => {
            t.classList.toggle('active', t.dataset.sheet === name);
        });

        // Clear selection
        selectedCell = null;
        cellRef.textContent = 'A1';
        formulaDisplay.textContent = rows[0] && rows[0][0] ? String(rows[0][0]) : '';

        // Clear search
        searchInput.value = '';
        searchCount.textContent = '';

        loading.style.display = 'none';
    }

    function buildTabs() {
        sheetTabs.innerHTML = '';
        const names = Object.keys(DATA);
        names.forEach((name, i) => {
            const btn = document.createElement('button');
            btn.className = 'sheet-tab';
            btn.dataset.sheet = name;
            btn.textContent = name;
            btn.addEventListener('click', () => renderSheet(name));
            sheetTabs.appendChild(btn);
        });
    }

    // Sync scroll positions
    function syncScroll() {
        colHeaderTrack.scrollLeft = gridViewport.scrollLeft;
        rowHeaderTrack.scrollTop = gridViewport.scrollTop;
    }

    // Cell click -> select
    function handleCellClick(e) {
        const td = e.target.closest('td');
        if (!td || td.dataset.row === undefined) return;

        // Remove previous selection
        if (selectedCell) selectedCell.classList.remove('selected');
        td.classList.add('selected');
        selectedCell = td;

        const r = parseInt(td.dataset.row);
        const c = parseInt(td.dataset.col);
        cellRef.textContent = colLetter(c) + (r + 1);
        const raw = DATA[activeSheet].rows[r][c];
        formulaDisplay.textContent = raw !== '' ? String(raw) : '';
    }

    // Search functionality
    function handleSearch() {
        const query = searchInput.value.trim().toLowerCase();

        // Remove old highlights
        dataTable.querySelectorAll('.highlight').forEach(el => el.classList.remove('highlight'));

        if (!query) {
            searchCount.textContent = '';
            return;
        }

        let count = 0;
        const cells = dataTable.querySelectorAll('td');
        cells.forEach(td => {
            if (td.textContent.toLowerCase().includes(query) || (td.title && td.title.toLowerCase().includes(query))) {
                td.classList.add('highlight');
                count++;
            }
        });

        searchCount.textContent = count > 0 ? `${count} found` : 'No match';

        // Scroll to first match
        const first = dataTable.querySelector('.highlight');
        if (first) first.scrollIntoView({ block: 'center', inline: 'center' });
    }

    // Tab scroll arrows
    document.getElementById('tabScrollLeft').addEventListener('click', () => {
        sheetTabsScroll.scrollBy({ left: -120, behavior: 'smooth' });
    });
    document.getElementById('tabScrollRight').addEventListener('click', () => {
        sheetTabsScroll.scrollBy({ left: 120, behavior: 'smooth' });
    });

    // Wire events
    gridViewport.addEventListener('scroll', syncScroll);
    dataTable.addEventListener('click', handleCellClick);

    let searchTimer = null;
    searchInput.addEventListener('input', () => {
        clearTimeout(searchTimer);
        searchTimer = setTimeout(handleSearch, 200);
    });

    // Keyboard: Ctrl+F focus search
    document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
            e.preventDefault();
            searchInput.focus();
            searchInput.select();
        }
    });

    // Load data
    fetch('data.json')
        .then(r => r.json())
        .then(data => {
            DATA = data;
            buildTabs();
            // Render first sheet
            const first = Object.keys(data)[0];
            renderSheet(first);
        })
        .catch(err => {
            loading.innerHTML = '<p style="color:#f14c4c;">Failed to load data. Please refresh.</p>';
            console.error(err);
        });
})();
