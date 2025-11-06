import { 
    getVisibleColumns, 
    getCurrentPage, 
    getRowsPerPage, 
    setCurrentPage, 
    getOriginalData, 
    getSortConfig, 
    setSortConfig,
    setVisibleColumns,
    getTableActiveFilters,
    getTableFilterValues,
    setTableActiveFilters,
    setTableFilterValues,
    getCurrentHeaders
} from '../../store/index.js';
import { sortData, createElement, getElement } from '../../utils/general.js';
import { applyFilters } from '../filters/FilterManager.js';
import { getFilteredData, detectColumnTypes, parseFlexibleDate } from '../filters/FilterManager.js';
import { getCurrentCustomColumns } from '../custom/CustomColumnManager.js';
import { tableNotification } from '../../js/notifications.js';

// Track selection state
let isSelecting = false;
let selectionStart = null;
let isTextSelecting = false;

// Track hidden columns
let hiddenColumns = new Set();

// --- Utilidad debounce simple ---
function debounce(fn, delay) {
  let timeout;
  return function(...args) {
    clearTimeout(timeout);
    timeout = setTimeout(() => fn.apply(this, args), delay);
  };
}

export function displayTable(data = []) {
    console.log("📊 Starting table display with data:", { 
        dataLength: data.length,
        visibleColumns: getVisibleColumns(),
        currentPage: getCurrentPage(),
        rowsPerPage: getRowsPerPage(),
        sampleRow: data[0]
    });

    const container = getElement("#tableContainer");
    if (!container) {
        console.error("❌ Table container not found!");
        return;
    }
    
    // Los chips de filtros ahora se manejan en renderMainTabsBar()

    // Ensure container is visible
    container.classList.add('visible');
    container.innerHTML = "";

    if (!Array.isArray(data) || data.length === 0) {
        console.warn("⚠️ No data to display");
        container.innerHTML = "<p>No data found.</p>";
        return;
    }

    const table = createElement("table", "data-table");
    
    // Add selection event listeners
    table.addEventListener('mousedown', startSelection);
    table.addEventListener('mousemove', updateSelection);
    table.addEventListener('mouseup', endSelection);
    document.addEventListener('mouseup', endSelection);
    
    // Add copy event listener
    document.addEventListener('copy', handleCopy);
    document.addEventListener('keydown', handleKeyDown);
    
    table.appendChild(createTableHeader());
    table.appendChild(createTableBody(data));
    container.appendChild(table);
    
    console.log("✅ Table rendered successfully");
    
    updatePagination(data.length);
    // colorRowsByUrgencia(); // TEMPORALMENTE COMENTADO PARA DEBUG
    
    // NO llamar refreshHeaderFilterIcons aquí para evitar conflictos con filtros del modal
    // refreshHeaderFilterIcons se llamará solo cuando se modifiquen filtros de tabla directamente

    // Permitir selección fluida desde el header (versión refinada)
    const ths = table.querySelectorAll('th');
    ths.forEach(th => {
        th.addEventListener('selectstart', (e) => {
            isTextSelecting = true;
            ths.forEach(t => t.style.pointerEvents = 'none');
        });
    });
    function restorePointerEvents() {
        if (isTextSelecting) {
            ths.forEach(t => t.style.pointerEvents = '');
            isTextSelecting = false;
        }
    }
    document.addEventListener('mouseup', restorePointerEvents);
    document.addEventListener('selectionchange', () => {
        if (!window.getSelection().toString()) {
            restorePointerEvents();
        }
    });

    // --- Botón Reset Filtros debajo de la paginación ---
    let resetBtn = document.getElementById('resetAllFiltersBtn');
    if (!resetBtn && pagination) {
        resetBtn = document.createElement('button');
        resetBtn.id = 'resetAllFiltersBtn';
        resetBtn.className = 'reset-filters-btn';
        resetBtn.textContent = 'Reset Filters';
        resetBtn.setAttribute('style', 'color:#fff; background:none; border:none; font-size:16px; font-family:inherit; padding:0; margin:0;');
        resetBtn.onclick = () => window.resetAllFilters();
        if (pagination.nextSibling) {
            pagination.parentNode.insertBefore(resetBtn, pagination.nextSibling);
        } else {
            pagination.parentNode.appendChild(resetBtn);
        }
    }

    const refreshBtn = document.getElementById('resetAllFiltersBtn');
    if (refreshBtn) {
        refreshBtn.onclick = () => {
            // Solo limpiar filtros Excel de cabecera de tabla
            const filterValues = { ...getTableFilterValues() };
            const activeFilters = { ...getTableActiveFilters() };
            let changed = false;
            Object.keys(activeFilters).forEach(col => {
                if (['reference', 'date'].includes(activeFilters[col])) {
                    delete filterValues[col];
                    delete activeFilters[col];
                    changed = true;
                }
            });
            if (changed) {
                setTableFilterValues(filterValues);
                setTableActiveFilters(activeFilters);
                applyFilters();
                refreshHeaderFilterIcons();
            }
            // Limpiar visualmente los dropdowns y embudos
            document.querySelectorAll('.excel-filter-dropdown').forEach(el => el.remove());
            document.querySelectorAll('.excel-filter-icon').forEach(icon => icon.classList.remove('excel-filter-active'));
            // Show unified notification
            if (typeof window.showUnifiedNotification === 'function') {
              window.showUnifiedNotification('All table filters cleared!', 'info');
            }
        };
    }

    // Forzar popup de copiado en el botón flotante de copiar
    setTimeout(() => {
      const floatingCopyBtn = document.querySelector('.copy-btn, .copy-table-btn, .floating-copy-btn');
      if (floatingCopyBtn) {
        floatingCopyBtn.addEventListener('click', () => {
          if (typeof window.showUnifiedNotification === 'function') {
            window.showUnifiedNotification('Data copied to clipboard!', 'success');
          }
        });
      }
    }, 500);
}

function createTableHeader() {
    const thead = createElement("thead");
    const headRow = createElement("tr");
    
    const filterValues = getTableFilterValues();
    getVisibleColumns().forEach(column => {
        const th = createElement("th");
        th.draggable = true;
        th.dataset.column = column;

        // --- Icono de filtro tipo Excel (embudo minimalista azul) ---
        const filterIcon = document.createElement('span');
        filterIcon.className = 'excel-filter-icon';
        filterIcon.style.position = 'relative';
        filterIcon.innerHTML = `
          <svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align:middle;">
            <path d="M4 4h12l-5 7v4a1 1 0 0 1-2 0v-4L4 4z" stroke="#47B2E5" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        `;
        filterIcon.style.cursor = 'pointer';
        filterIcon.style.marginRight = '0.5em';
        filterIcon.title = 'Filter this column';
        filterIcon.addEventListener('click', (e) => {
            e.stopPropagation();
            showExcelFilterDropdown(th, column);
        });
        th.appendChild(filterIcon);

        // --- Título de la columna ---
        const titleSpan = document.createElement('span');
        titleSpan.textContent = column;
        th.appendChild(titleSpan);
        
        // --- Estilo especial para columnas con diferencias en duplicados ---
        // SOLO aplicar colores especiales si estamos en modo análisis de duplicados
        if (window.currentDuplicateDifferences && window.currentDuplicateColumns && window.duplicateAnalysisMode) {
            const isDuplicateColumn = window.currentDuplicateColumns.includes(column);
            
            if (isDuplicateColumn) {
                // Es una columna de duplicados - AMARILLO
                th.style.setProperty('background-color', '#fff3cd', 'important');
                th.style.setProperty('color', '#856404', 'important');
                th.title = 'Column used for duplicate detection - Click to filter';
                // Botón SVG moderno para duplicados (layers/stack)
                const filterBtn = document.createElement('button');
                filterBtn.innerHTML = `
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" style="vertical-align:middle;">
                    <path d="M12 4L2 9l10 5 10-5-10-5z" fill="#FFC107"/>
                    <path d="M2 15l10 5 10-5" stroke="#FFC107" stroke-width="1.5" fill="none"/>
                  </svg>
                `;
                filterBtn.style.cssText = `
                    position: absolute;
                    top: 2px;
                    right: 2px;
                    background: transparent;
                    border: none;
                    width: 18px;
                    height: 18px;
                    padding: 0;
                    cursor: pointer;
                    z-index: 100;
                `;
                filterBtn.title = 'Filter by duplicate columns';
                filterBtn.onclick = (e) => {
                    e.stopPropagation();
                    filterByDuplicateColumns();
                };
                th.style.position = 'relative';
                th.appendChild(filterBtn);
            }
            // Verificar si esta columna tiene diferencias
            const hasDifferences = window.currentDuplicateDifferences.some(([key, diffObj]) => diffObj[column]);
            if (hasDifferences) {
                // Columna con diferencias - ROJO
                th.style.setProperty('background-color', '#ffebee', 'important');
                th.style.setProperty('color', '#c62828', 'important');
                th.style.setProperty('border-left', '3px solid #c62828', 'important');
                th.title = 'This column has different values among duplicates - Click to filter';
                // Botón SVG moderno para diferencias (alert/warning)
                const filterBtn = document.createElement('button');
                filterBtn.innerHTML = `
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" style="vertical-align:middle;">
                    <circle cx="12" cy="12" r="10" fill="#E53935"/>
                    <rect x="11" y="7" width="2" height="6" rx="1" fill="#fff"/>
                    <rect x="11" y="15" width="2" height="2" rx="1" fill="#fff"/>
                  </svg>
                `;
                filterBtn.style.cssText = `
                    position: absolute;
                    top: 2px;
                    right: 2px;
                    background: transparent;
                    border: none;
                    width: 18px;
                    height: 18px;
                    padding: 0;
                    cursor: pointer;
                    z-index: 100;
                `;
                filterBtn.title = 'Filter by differences in this column';
                filterBtn.onclick = (e) => {
                    e.stopPropagation();
                    filterByColumnDifferences(column);
                };
                th.style.position = 'relative';
                th.appendChild(filterBtn);
            }
        }

        // --- Botón de ocultar columna (X, solo texto) ---
        const hideBtn = document.createElement('button');
        hideBtn.className = 'hide-column-x-btn';
        hideBtn.textContent = '×';
        hideBtn.type = 'button';
        hideBtn.style.marginLeft = '0.3em';
        hideBtn.style.marginRight = '1.2em';
        hideBtn.style.cursor = 'pointer';
        hideBtn.style.fontWeight = 'bold';
        hideBtn.style.fontSize = '1.2em';
        hideBtn.style.color = '#10B981';
        hideBtn.setAttribute('style', hideBtn.getAttribute('style') + ';color:#10B981 !important;');
        hideBtn.title = 'Hide column';
        hideBtn.addEventListener('click', async (e) => {
            e.stopPropagation();
            const currentVisible = getVisibleColumns();
            if (currentVisible.length <= 1) return;
            const newVisible = currentVisible.filter(col => col !== column);
            setVisibleColumns(newVisible);
            hiddenColumns.add(column);
            // Sincronizar el valor del input global search con el filtro global
            const globalSearchInput = document.getElementById('globalSearchInput');
            if (globalSearchInput) {
                const { setModuleFilterValues, getModuleFilterValues } = await import('../../store/index.js');
                const current = getModuleFilterValues();
                setModuleFilterValues({
                    ...current,
                    __globalSearch: globalSearchInput.value
                });
            }
            applyFilters();
            const checkbox = document.querySelector(`#columnList input[type='checkbox'][value='${column.replace(/'/g, "\\'") }']`);
            if (checkbox) checkbox.checked = false;
            updateHiddenColumnsButton();
        });
        th.appendChild(hideBtn);
        
        const sortConfig = getSortConfig();
        if (sortConfig && sortConfig.column === column) {
            th.classList.add(sortConfig.direction === 'asc' ? 'sorted-asc' : 'sorted-desc');
        }
        
        th.addEventListener('dragstart', handleDragStart);
        th.addEventListener('dragover', handleDragOver);
        th.addEventListener('drop', handleDrop);
        th.addEventListener('dragend', handleDragEnd);
        
        th.addEventListener('click', () => handleSort(column));
        headRow.appendChild(th);
    });

    thead.appendChild(headRow);
    return thead;
}

// Drag and drop handlers
let draggedColumn = null;

function handleDragStart(e) {
    draggedColumn = e.target;
    e.target.classList.add('dragging');
    // Set custom drag image or transparency
    e.dataTransfer.effectAllowed = 'move';
}

function handleDragOver(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    
    const th = e.target.closest('th');
    if (th && th !== draggedColumn) {
        const rect = th.getBoundingClientRect();
        const midPoint = rect.x + rect.width / 2;
        
        // Add visual indicator for drop position
        th.classList.remove('drop-right', 'drop-left');
        if (e.clientX > midPoint) {
            th.classList.add('drop-right');
        } else {
            th.classList.add('drop-left');
        }
    }
}

function handleDrop(e) {
    e.preventDefault();
    const targetTh = e.target.closest('th');
    
    if (targetTh && draggedColumn && targetTh !== draggedColumn) {
        const columns = getVisibleColumns();
        const fromIndex = columns.indexOf(draggedColumn.dataset.column);
        const toIndex = columns.indexOf(targetTh.dataset.column);
        
        // Reorder columns
        columns.splice(fromIndex, 1);
        columns.splice(toIndex, 0, draggedColumn.dataset.column);
        
        // Update state and redraw table
        setVisibleColumns(columns);
        applyFilters();
    }
    
    // Clean up visual indicators
    document.querySelectorAll('th').forEach(th => {
        th.classList.remove('drop-right', 'drop-left');
    });
}

function handleDragEnd() {
    draggedColumn = null;
    document.querySelectorAll('th').forEach(th => {
        th.classList.remove('dragging', 'drop-right', 'drop-left');
    });
}

function handleKeyDown(e) {
    // Handle Ctrl+C or Cmd+C
    if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
        e.preventDefault();
        copySelectedCells();
    }
}

function handleCopy(e) {
    const selectedCells = document.querySelectorAll('.data-table td.selected');
    if (selectedCells.length > 0) {
        e.preventDefault();
        copySelectedCells();
    }
}

function startSelection(e) {
    // Si el click es sobre un input, textarea o select, no hacer selección
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') {
        return;
    }
    const cell = e.target.closest('td');
    if (!cell) return;
    isSelecting = true;
    selectionStart = cell;
    clearSelection();
    cell.classList.add('selected');
    e.preventDefault();
}

function updateSelection(e) {
    // Si el mouse está sobre un input, textarea o select, no actualizar selección
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') {
        return;
    }
    if (!isSelecting || !selectionStart) return;
    const currentCell = e.target.closest('td');
    if (!currentCell) return;
    const table = currentCell.closest('table');
    if (!table) return;
    // Selección rectangular tipo Excel
    const allRows = Array.from(table.querySelectorAll('tbody tr'));
    const allCells = Array.from(table.querySelectorAll('td'));
    const startRow = selectionStart.parentElement.rowIndex - 1; // tbody rowIndex
    const startCol = selectionStart.cellIndex;
    const endRow = currentCell.parentElement.rowIndex - 1;
    const endCol = currentCell.cellIndex;
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);
    const minCol = Math.min(startCol, endCol);
    const maxCol = Math.max(startCol, endCol);
        clearSelection();
    for (let r = minRow; r <= maxRow; r++) {
        const row = allRows[r];
        if (!row) continue;
        for (let c = minCol; c <= maxCol; c++) {
            const cell = row.cells[c];
            if (cell) cell.classList.add('selected');
            }
    }
    e.preventDefault();
}

function endSelection() {
    isSelecting = false;
}

function getColumnIndex(cell) {
    return Array.from(cell.parentElement.children).indexOf(cell);
}

function clearSelection() {
    document.querySelectorAll('.data-table td.selected').forEach(cell => {
        cell.classList.remove('selected');
    });
}

function createTableBody(data) {
    const tbody = createElement("tbody");
    const startIndex = (getCurrentPage() - 1) * getRowsPerPage();
    const endIndex = startIndex + getRowsPerPage();
    const pageData = data.slice(startIndex, endIndex);

    // Obtener tipos de columna
    const columnTypes = detectColumnTypes(getOriginalData());
    const customColumns = getCurrentCustomColumns ? getCurrentCustomColumns() : [];
    const customHeaders = customColumns.map(c => c.header);
    const allHeaders = getVisibleColumns();

    // Detectar si la tabla es editable (opcional: puedes pasar una prop editable)
    const isEditable = document.querySelector('.main-tab-btn.active')?.textContent?.startsWith('Tab ');

    pageData.forEach((row, rowIdx) => {
        const tr = createElement("tr");
        // Índice global de la fila
        const globalIdx = startIndex + rowIdx;
        // --- NUEVO: aplicar color de fondo si existe en el objeto global ---
        if (data[globalIdx] && data[globalIdx].rowColor) {
            tr.style.background = data[globalIdx].rowColor;
        }
        // --- NUEVO: evitar que el mousedown derecho limpie la selección ---
        tr.addEventListener('mousedown', function(e) {
            const cell = e.target.closest('td');
            if (e.button === 2 && cell && cell.classList.contains('selected')) {
                e.preventDefault();
            }
        });
        // --- NUEVO: menú contextual para colorear fila ---
        if (isEditable) {
            tr.addEventListener('contextmenu', function(e) {
                const cell = e.target.closest('td');
                const isCellSelected = cell && cell.classList.contains('selected');
                if (cell && !isCellSelected) {
                    document.querySelectorAll('.data-table td.selected').forEach(c => c.classList.remove('selected'));
                    cell.classList.add('selected');
                }
                e.preventDefault();
                document.querySelectorAll('.row-color-picker-menu').forEach(el => el.remove());
                const pastelColors = [
                  '#fffbe7', '#e3fcec', '#e3f2fd', '#fce4ec', '#f3e8fd', '#f9fbe7', '#fbeee6', '#f0f4c3', '#e0f7fa', '#f8bbd0'
                ];
                const table = tr.closest('table');
                const selectedRows = new Set();
                if (table) {
                  table.querySelectorAll('td.selected').forEach(td => {
                    const rowEl = td.parentElement;
                    if (rowEl) selectedRows.add(rowEl);
                  });
                }
                if (selectedRows.size === 0) selectedRows.add(tr);
                const menu = document.createElement('div');
                menu.className = 'row-color-picker-menu';
                menu.style.position = 'fixed';
                menu.style.left = e.clientX + 'px';
                menu.style.top = e.clientY + 'px';
                menu.style.zIndex = 99999;
                menu.style.display = 'flex';
                menu.style.gap = '0.3em';
                menu.style.background = '#fff';
                menu.style.border = '1px solid #ddd';
                menu.style.borderRadius = '8px';
                menu.style.boxShadow = '0 2px 12px rgba(0,0,0,0.08)';
                menu.style.padding = '0.4em 0.5em';
                menu.style.margin = '0';
                menu.style.alignItems = 'center';
                pastelColors.forEach(color => {
                  const swatch = document.createElement('button');
                  swatch.type = 'button';
                  swatch.style.background = color;
                  swatch.style.width = '24px';
                  swatch.style.height = '24px';
                  swatch.style.border = '2px solid #eee';
                  swatch.style.borderRadius = '6px';
                  swatch.style.cursor = 'pointer';
                  swatch.style.outline = 'none';
                  swatch.style.margin = '0 2px';
                  swatch.title = color;
                  swatch.onclick = (ev) => {
                    ev.stopPropagation();
                    selectedRows.forEach(rowEl => {
                      // Índice global de la fila
                      const idx = startIndex + Array.from(rowEl.parentElement.children).indexOf(rowEl);
                      if (data[idx]) {
                        data[idx].rowColor = color;
                        rowEl.style.background = color;
                      }
                    });
                    menu.remove();
                  };
                  menu.appendChild(swatch);
                });
                const clearBtn = document.createElement('button');
                clearBtn.type = 'button';
                clearBtn.textContent = '✕';
                clearBtn.title = 'Quitar color';
                clearBtn.style.background = 'none';
                clearBtn.style.border = 'none';
                clearBtn.style.color = '#888';
                clearBtn.style.fontWeight = 'bold';
                clearBtn.style.fontSize = '1.2em';
                clearBtn.style.cursor = 'pointer';
                clearBtn.style.marginLeft = '0.5em';
                clearBtn.onclick = (ev) => {
                  ev.stopPropagation();
                  selectedRows.forEach(rowEl => {
                    const idx = startIndex + Array.from(rowEl.parentElement.children).indexOf(rowEl);
                    if (data[idx]) {
                      data[idx].rowColor = '';
                      rowEl.style.background = '';
                    }
                  });
                  menu.remove();
                };
                menu.appendChild(clearBtn);
                document.body.appendChild(menu);
                setTimeout(() => {
                  document.addEventListener('mousedown', function handler(ev) {
                    if (!menu.contains(ev.target)) {
                      menu.remove();
                      document.removeEventListener('mousedown', handler);
                    }
                  });
                }, 0);
            });
        }
        getVisibleColumns().forEach((column, colIdx) => {
            const td = createElement("td");
            const isCustom = customHeaders.includes(column);
            let value = row[column];
            if (isCustom) {
                // Buscar el valor en la columna personalizada
                const customCol = customColumns.find(c => c.header === column);
                value = customCol && customCol.values ? customCol.values[startIndex + rowIdx] : '';
                // Renderizar input editable
                const input = document.createElement('input');
                input.type = 'text';
                input.className = 'cell-input';
                input.value = value || '';
                input.removeAttribute('readonly');
                input.removeAttribute('disabled');
                // Si es fórmula, evalúa y muestra el resultado
                if (typeof value === 'string' && value.startsWith('=')) {
                    try {
                        const context = {};
                        getVisibleColumns().forEach((col, i) => {
                            context[col] = row[col] || '';
                        });
                        let expr = value.slice(1);
                        // Ordenar los headers de mayor a menor longitud para evitar reemplazos parciales
                        const sortedHeaders = Object.keys(context).sort((a, b) => b.length - a.length);
                        sortedHeaders.forEach((header) => {
                            // Reemplazar el header solo si aparece como palabra completa (usando \b y escapando espacios)
                            const safeHeader = header.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                            const regex = new RegExp(`\\b${safeHeader}\\b`, 'g');
                            expr = expr.replace(regex, `context[${JSON.stringify(header)}]`);
                        });
                        // Funciones de fechas
                        function parseDate(val) {
                            if (!val) return null;
                            if (val instanceof Date) return val;
                            if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}/.test(val)) return new Date(val);
                            const d = new Date(val);
                            return isNaN(d) ? null : d;
                        }
                        function DAYS(f2, f1) {
                            const d2 = parseDate(f2);
                            const d1 = parseDate(f1);
                            if (!d1 || !d2) return '#ERROR';
                            return Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
                        }
                        function DATEDIF(f1, f2, unit) {
                            const d1 = parseDate(f1);
                            const d2 = parseDate(f2);
                            if (!d1 || !d2) return '#ERROR';
                            if (unit === 'd') return Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
                            if (unit === 'm') return (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth());
                            if (unit === 'y') return d2.getFullYear() - d1.getFullYear();
                            return '#ERROR';
                        }
                        function YEAR(f) { const d = parseDate(f); return d ? d.getFullYear() : '#ERROR'; }
                        function MONTH(f) { const d = parseDate(f); return d ? d.getMonth() + 1 : '#ERROR'; }
                        function DAY(f) { const d = parseDate(f); return d ? d.getDate() : '#ERROR'; }
                        // Sugerencia si se intenta restar fechas directamente
                        if (/context\[.*?\]\s*-\s*context\[.*?\]/.test(expr)) {
                            input.value = '#ERROR';
                            if (!input.nextSibling || !input.nextSibling.classList || !input.nextSibling.classList.contains('formula-help-box')) {
                                const help = document.createElement('div');
                                help.className = 'formula-help-box';
                                help.style.position = 'absolute';
                                help.style.zIndex = '100000';
                                help.style.background = '#f8f9fa';
                                help.style.border = '1px solid #228be6';
                                help.style.borderRadius = '4px';
                                help.style.boxShadow = '0 2px 8px rgba(34,139,230,0.08)';
                                help.style.margin = '0';
                                help.style.padding = '0.5em 1em';
                                help.style.fontSize = '13px';
                                help.style.color = '#d32f2f';
                                help.textContent = 'Para diferencias de fechas usa: DAYS(fecha_fin, fecha_inicio) o DATEDIF(fecha_inicio, fecha_fin, "d")';
                                input.parentNode.appendChild(help);
                            }
                            return;
                        } else if (input.nextSibling && input.nextSibling.classList && input.nextSibling.classList.contains('formula-help-box')) {
                            input.nextSibling.remove();
                        }
                        // eslint-disable-next-line no-eval
                        input.value = eval(expr);
                    } catch {
                        input.value = '#ERROR';
                    }
                }
                // --- AUTOCOMPLETADO DE ENCABEZADOS Y FUNCIONES ---
                const FUNCTION_HELP = {
                    SUM: 'SUM(n1, n2, ...): Suma los valores dados. Ej: SUM(A, B)',
                    AVG: 'AVG(n1, n2, ...): Promedio de los valores. Ej: AVG(A, B)',
                    MIN: 'MIN(n1, n2, ...): Mínimo de los valores. Ej: MIN(A, B)',
                    MAX: 'MAX(n1, n2, ...): Máximo de los valores. Ej: MAX(A, B)',
                    IF: 'IF(condición, valor_si_verdadero, valor_si_falso): Ej: IF(A > 10, "Sí", "No")',
                    COUNT: 'COUNT(n1, n2, ...): Cuenta los valores no vacíos.',
                    DAYS: 'DAYS(fecha_fin, fecha_inicio): Días entre dos fechas.',
                    DATEDIF: 'DATEDIF(fecha_inicio, fecha_fin, "d"): Diferencia entre fechas en días ("d"), meses ("m") o años ("y").',
                    YEAR: 'YEAR(fecha): Año de una fecha.',
                    MONTH: 'MONTH(fecha): Mes de una fecha.',
                    DAY: 'DAY(fecha): Día de una fecha.'
                };
                const FUNCTION_LIST = Object.keys(FUNCTION_HELP);
                let autocompleteBox = null;
                let helpBox = null;
                function showAutocompleteBox(matches, isFunc = false) {
                    if (!autocompleteBox) {
                        autocompleteBox = document.createElement('ul');
                        autocompleteBox.className = 'autocomplete-header-list';
                        autocompleteBox.style.position = 'absolute';
                        autocompleteBox.style.zIndex = '99999';
                        autocompleteBox.style.background = '#fff';
                        autocompleteBox.style.border = '1px solid #228be6';
                        autocompleteBox.style.borderRadius = '4px';
                        autocompleteBox.style.boxShadow = '0 2px 8px rgba(34,139,230,0.08)';
                        autocompleteBox.style.margin = '0';
                        autocompleteBox.style.padding = '0.2em 0';
                        autocompleteBox.style.listStyle = 'none';
                        autocompleteBox.style.fontSize = '14px';
                        autocompleteBox.style.minWidth = '160px';
                        document.body.appendChild(autocompleteBox);
                    }
                    // Posicionar justo debajo del input
                    const rect = input.getBoundingClientRect();
                    autocompleteBox.style.left = `${rect.left + window.scrollX}px`;
                    autocompleteBox.style.top = `${rect.bottom + window.scrollY}px`;
                    autocompleteBox.innerHTML = '';
                    matches.forEach(header => {
                        const li = document.createElement('li');
                        li.textContent = header;
                        li.style.padding = '0.2em 0.8em';
                        li.style.cursor = 'pointer';
                        li.addEventListener('mousedown', (ev) => {
                            ev.preventDefault();
                            const val = input.value;
                            const cursorPos = input.selectionStart;
                            const before = val.slice(0, cursorPos);
                            const after = val.slice(cursorPos);
                            if (isFunc) {
                                // Insertar función con paréntesis y colocar el cursor dentro
                                const funcSyntax = header + '()';
                                input.value = before + funcSyntax + after;
                                setTimeout(() => {
                                    input.selectionStart = input.selectionEnd = before.length + header.length + 1;
                                }, 0);
                                showHelpBox(header);
                            } else {
                                // Reemplazar el término actual por el encabezado seleccionado
                                const lastTermIdx = before.lastIndexOf('=') + 1;
                                const newVal = before.slice(0, lastTermIdx) + header + after;
                                input.value = newVal;
                                hideHelpBox();
                            }
                            input.dispatchEvent(new Event('input'));
                            if (autocompleteBox) autocompleteBox.remove();
                            autocompleteBox = null;
                            input.focus();
                        });
                        li.addEventListener('mouseover', () => {
                            li.style.background = '#e3f0fc';
                            if (isFunc) showHelpBox(header);
                        });
                        li.addEventListener('mouseout', () => {
                            li.style.background = '';
                        });
                        autocompleteBox.appendChild(li);
                    });
                }
                function showHelpBox(funcName) {
                    if (!FUNCTION_HELP[funcName]) return;
                    if (!helpBox) {
                        helpBox = document.createElement('div');
                        helpBox.className = 'formula-help-box';
                        helpBox.style.position = 'absolute';
                        helpBox.style.zIndex = '100000';
                        helpBox.style.background = '#f8f9fa';
                        helpBox.style.border = '1px solid #228be6';
                        helpBox.style.borderRadius = '4px';
                        helpBox.style.boxShadow = '0 2px 8px rgba(34,139,230,0.08)';
                        helpBox.style.margin = '0';
                        helpBox.style.padding = '0.5em 1em';
                        helpBox.style.fontSize = '13px';
                        helpBox.style.color = '#1976d2';
                        document.body.appendChild(helpBox);
                    }
                    const rect = input.getBoundingClientRect();
                    helpBox.style.left = `${rect.left + window.scrollX}px`;
                    helpBox.style.top = `${rect.bottom + window.scrollY + 36}px`;
                    helpBox.textContent = FUNCTION_HELP[funcName];
                }
                function hideHelpBox() {
                    if (helpBox) { helpBox.remove(); helpBox = null; }
                }
                input.addEventListener('input', (e) => {
                    const val = input.value;
                    if (val.startsWith('=') && val.length > 1) {
                        const term = val.slice(1).split(/[^a-zA-Z0-9_ ]/).pop().trim().toLowerCase();
                        // Sugerir funciones si el término coincide
                        const funcMatches = FUNCTION_LIST.filter(f => f.toLowerCase().startsWith(term));
                        const headerMatches = allHeaders.filter(h => h.toLowerCase().includes(term) && term.length > 0);
                        if (funcMatches.length > 0) {
                            showAutocompleteBox(funcMatches, true);
                        } else if (headerMatches.length > 0) {
                            showAutocompleteBox(headerMatches, false);
                        } else if (autocompleteBox) {
                            autocompleteBox.remove();
                            autocompleteBox = null;
                            hideHelpBox();
                        }
                    } else if (autocompleteBox) {
                        autocompleteBox.remove();
                        autocompleteBox = null;
                        hideHelpBox();
                    }
                });
                input.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter') {
                        if (autocompleteBox) {
                            autocompleteBox.remove();
                            autocompleteBox = null;
                        }
                        if (helpBox) { helpBox.remove(); helpBox = null; }
                        input.blur(); // Forzar blur para guardar y evaluar la fórmula
                        e.preventDefault();
                    } else if (e.key === 'Escape') {
                        if (autocompleteBox) {
                            autocompleteBox.remove();
                            autocompleteBox = null;
                        }
                        if (helpBox) { helpBox.remove(); helpBox = null; }
                        e.preventDefault();
                    }
                });
                // --- FIN AUTOCOMPLETADO ---
                input.addEventListener('focus', () => {
                    // Al enfocar, mostrar la fórmula original si existe
                    if (typeof value === 'string' && value.startsWith('=')) {
                        input.value = value;
                    }
                });
                input.addEventListener('blur', () => {
                    // Guardar valor/fórmula en la columna personalizada
                    const customCol = customColumns.find(c => c.header === column);
                    if (customCol) {
                        customCol.values[startIndex + rowIdx] = input.value;
                        // Guardar en localStorage
                        if (typeof customColManager !== 'undefined') {
                            customColManager.saveCustomColumns();
                        }
                    }
                    // Recalcular todas las fórmulas
                    recalculateAllCustomFormulas(customColumns, pageData, startIndex);
                });
                td.innerHTML = '';
                td.appendChild(input);
                td.style.position = 'relative';
            } else {
                // Columna normal
            td.textContent = value;
                
                // --- APLICAR COLORES PARA PESTAÑAS DE ANÁLISIS DE DUPLICADOS ---
                // SOLO aplicar colores especiales si estamos en modo análisis
                if (window.currentDuplicateDifferences && window.currentDuplicateColumns && window.duplicateAnalysisMode) {
                    const isDuplicateColumn = window.currentDuplicateColumns.includes(column);
                    const rowKey = window.currentDuplicateColumns.map(c => row[c]).join('|');
                    const differences = window.currentDuplicateDifferences.find(([key, diffObj]) => key === rowKey);
                    const hasDiff = differences && differences[1][column];
                    
                    if (hasDiff) {
                        // Celda con valor diferente - ROJO
                        td.style.setProperty('background-color', '#ffebee', 'important');
                        td.style.setProperty('color', '#c62828', 'important');
                        td.style.setProperty('font-weight', 'bold', 'important');
                        td.style.setProperty('border-left', '3px solid #c62828', 'important');
                        console.log('🔴 Colored cell red:', column, row[column]);
                    } else if (isDuplicateColumn) {
                        // Columna de duplicados - AMARILLO
                        td.style.setProperty('background-color', '#fff3cd', 'important');
                        td.style.setProperty('color', '#856404', 'important');
                        td.style.setProperty('font-weight', '600', 'important');
                        console.log('🟡 Colored cell yellow:', column, row[column]);
                    }
                }
            }
            const columnType = columnTypes[column];
            if (columnType) {
                td.setAttribute('data-type', columnType);
            }
            td.classList.add('selectable');
            
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    return tbody;
}

// Función para recalcular todas las fórmulas de columnas personalizadas
function recalculateAllCustomFormulas(customColumns, pageData, startIndex) {
    customColumns.forEach((col, colIdx) => {
        if (!col.values) return;
        col.values.forEach((val, rowIdx) => {
            if (typeof val === 'string' && val.startsWith('=')) {
                try {
                    const context = {};
                    getVisibleColumns().forEach((colName) => {
                        context[colName] = pageData[rowIdx] ? pageData[rowIdx][colName] : '';
                    });
                    let expr = val.slice(1);
                    Object.entries(context).forEach(([k, v]) => {
                        expr = expr.replaceAll(new RegExp(`\\b${k}\\b`, 'g'), v || 0);
                    });
                    // eslint-disable-next-line no-eval
                    col.values[rowIdx + startIndex] = eval(expr);
                } catch {
                    col.values[rowIdx + startIndex] = '#ERROR';
                }
            }
        });
    });
    // Forzar refresco de la tabla
    displayTable(getOriginalData());
}

function handleSort(column) {
    const currentSortConfig = getSortConfig();
    const direction = currentSortConfig && currentSortConfig.column === column
        ? (currentSortConfig.direction === 'asc' ? 'desc' : 'asc')
        : 'asc';
    setSortConfig({ column, direction });
    applyFilters();
}

export { updatePagination };

function updatePagination(totalRecords) {
    const paginationContainer = getElement("#pagination");
    if (!paginationContainer) return;

    const totalPages = Math.ceil(totalRecords / getRowsPerPage());
    const currentPage = getCurrentPage();

    paginationContainer.innerHTML = "";

    // Previous button
    appendPaginationButton(paginationContainer, "«", currentPage > 1, () => handlePageChange(1), false, "First");
    appendPaginationButton(paginationContainer, "‹", currentPage > 1, () => handlePageChange(currentPage - 1), false, "Previous");

    // Calculate page range to show
    let startPage = Math.max(1, currentPage - 2);
    let endPage = Math.min(totalPages, startPage + 4);
    
    // Adjust start if we're near the end
    if (endPage === totalPages) {
        startPage = Math.max(1, endPage - 4);
    }

    // First page if not in range
    if (startPage > 1) {
        appendPaginationButton(paginationContainer, "1", true, () => handlePageChange(1));
        if (startPage > 2) {
            appendPaginationButton(paginationContainer, "...", false, null, false, "More pages");
        }
    }

    // Page numbers
    for (let i = startPage; i <= endPage; i++) {
        appendPaginationButton(paginationContainer, i.toString(), true, () => handlePageChange(i), i === currentPage);
    }

    // Last page if not in range
    if (endPage < totalPages) {
        if (endPage < totalPages - 1) {
            appendPaginationButton(paginationContainer, "...", false, null, false, "More pages");
        }
        appendPaginationButton(paginationContainer, totalPages.toString(), true, () => handlePageChange(totalPages));
    }

    // Next buttons
    appendPaginationButton(paginationContainer, "›", currentPage < totalPages, () => handlePageChange(currentPage + 1), false, "Next");
    appendPaginationButton(paginationContainer, "»", currentPage < totalPages, () => handlePageChange(totalPages), false, "Last");

    // Update record count and page info
    const recordCount = getElement("#recordCount");
    if (recordCount) {
        const start = (currentPage - 1) * getRowsPerPage() + 1;
        const end = Math.min(currentPage * getRowsPerPage(), totalRecords);
        recordCount.textContent = `${start}-${end} of ${totalRecords} records`;
    }
}

function appendPaginationButton(container, text, enabled, onClick, isActive = false, title = "") {
    const button = createElement("button", "pagination-btn");
    button.textContent = text;
    button.disabled = !enabled;
    if (isActive) button.classList.add("active");
    if (title) button.title = title;
    if (onClick) button.addEventListener("click", onClick);
    container.appendChild(button);
}

function handlePageChange(newPage) {
    setCurrentPage(newPage);
    applyFilters();
}

function copySelectedCells() {
    const selectedCells = document.querySelectorAll('.data-table td.selected');
    if (selectedCells.length === 0) return;

    // Obtener la fila y columna de cada celda seleccionada
    const cellMap = {};
    selectedCells.forEach(cell => {
        const row = cell.parentElement;
        const rowIndex = Array.from(row.parentElement.children).indexOf(row);
        const colIndex = Array.from(cell.parentElement.children).indexOf(cell);
        if (!cellMap[rowIndex]) cellMap[rowIndex] = {};
        cellMap[rowIndex][colIndex] = cell.textContent.trim();
    });
    const rowIndices = Object.keys(cellMap).map(Number).sort((a, b) => a - b);
    const colIndices = Object.keys(cellMap[rowIndices[0]]).map(Number).sort((a, b) => a - b);

    // Construir datos seleccionados (solo las celdas, sin encabezados)
    const data = rowIndices.map(r => colIndices.map(c => cellMap[r][c] || ''));

    // Formato texto plano (CSV con comas) - todas las celdas separadas por comas
    const allCells = data.flat();
    const text = allCells.map(cell => {
      const cellStr = String(cell);
      if (cellStr.includes(',') || cellStr.includes('\n') || cellStr.includes('"')) {
        return `"${cellStr.replace(/"/g, '""')}"`;
      }
      return cellStr;
    }).join(',');

    // Debug: mostrar el texto que se va a copiar
    console.log('Texto a copiar:', text);

    // Formato HTML - solo datos, sin encabezados
    const htmlTable = `
      <table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px;">
        <tbody>
          ${data.map(row =>
            `<tr>${row.map(cell => `<td style="border: 1px solid #000000; padding: 4px; text-align: left;">${cell}</td>`).join('')}</tr>`
          ).join('')}
        </tbody>
      </table>
    `;

    // Copiar al portapapeles con formato Excel - priorizar texto plano
    navigator.clipboard.write([
      new window.ClipboardItem({
        'text/plain': new Blob([text], { type: 'text/plain' }),
        'text/html': new Blob([htmlTable], { type: 'text/html' })
      })
    ]).then(() => {
        showCopyFeedback(selectedCells);
        // Usar función global si existe, sino usar local
        if (typeof window.showCopyNotification === 'function') {
            window.showCopyNotification('Selected cells copied to clipboard!');
        } else {
            showInfoModal('', 'Selected cells copied to clipboard!');
        }
    }).catch(() => {
        // Fallback para texto plano
        const textarea = document.createElement('textarea');
        textarea.value = text;
        textarea.classList.add('offscreen');
        document.body.appendChild(textarea);
        textarea.select();
        try {
            document.execCommand('copy');
            showCopyFeedback(selectedCells);
            // Usar función global si existe, sino usar local
            if (typeof window.showCopyNotification === 'function') {
                window.showCopyNotification('Selected cells copied to clipboard!');
            } else {
                showInfoModal('', 'Selected cells copied to clipboard!');
            }
        } catch (err) {
            console.error('Failed to copy:', err);
        }
        document.body.removeChild(textarea);
    });
}

function showCopyFeedback(cells) {
    cells.forEach(cell => {
        cell.classList.add('copy-feedback');
        setTimeout(() => cell.classList.remove('copy-feedback'), 300);
    });
}

// --- Dropdown básico para filtro tipo Excel ---
function showExcelFilterDropdown(th, column) {
    document.querySelectorAll('.excel-filter-dropdown').forEach(el => el.remove());

    const dropdown = document.createElement('div');
    dropdown.className = 'excel-filter-dropdown';
    dropdown.style.position = 'absolute';
    // --- Posicionamiento exacto debajo del th ---
    const rect = th.getBoundingClientRect();
    dropdown.style.top = (rect.bottom + window.scrollY) + 'px';
    dropdown.style.left = (rect.left + window.scrollX) + 'px';
    dropdown.style.width = rect.width + 'px';
    // ---
    dropdown.style.background = '#fff';
    dropdown.style.border = '1px solid var(--border-color)';
    dropdown.style.borderRadius = '0 0 6px 6px';
    dropdown.style.boxShadow = '0 2px 8px rgba(0,0,0,0.08)';
    dropdown.style.zIndex = '100';
    dropdown.style.padding = '0.5rem';
    dropdown.style.fontSize = '0.95rem';

    // Search input
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.className = 'excel-filter-search';
    searchInput.placeholder = 'Search...';
    searchInput.style.width = '100%';
    searchInput.style.padding = '0.3rem';
    searchInput.style.marginBottom = '0.5rem';
    searchInput.style.border = '1px solid var(--border-color)';
    searchInput.style.borderRadius = '4px';
    dropdown.appendChild(searchInput);

    // Options container
    const optionsDiv = document.createElement('div');
    optionsDiv.className = 'excel-filter-options';
    optionsDiv.style.maxHeight = '200px';
    optionsDiv.style.overflowY = 'auto';
    dropdown.appendChild(optionsDiv);

    // Buttons container
    const buttonsDiv = document.createElement('div');
    buttonsDiv.style.display = 'flex';
    buttonsDiv.style.justifyContent = 'space-between';
    buttonsDiv.style.marginTop = '0.5rem';
    buttonsDiv.style.paddingTop = '0.5rem';
    buttonsDiv.style.borderTop = '1px solid var(--border-color)';

    // Select all button
    const selectAllBtn = document.createElement('button');
    selectAllBtn.className = 'excel-filter-selectall-btn';
    selectAllBtn.textContent = 'Select all';
    selectAllBtn.style.marginRight = '0.5rem';
    buttonsDiv.appendChild(selectAllBtn);

    // Clear all button
    const clearAllBtn = document.createElement('button');
    clearAllBtn.className = 'excel-filter-clearall-btn';
    clearAllBtn.textContent = 'Clear all';
    clearAllBtn.style.marginRight = '0.5rem';
    buttonsDiv.appendChild(clearAllBtn);

    // Apply button
    const applyBtn = document.createElement('button');
    applyBtn.className = 'excel-filter-apply-btn';
    applyBtn.textContent = 'Apply';
    buttonsDiv.appendChild(applyBtn);

    dropdown.appendChild(buttonsDiv);
    document.body.appendChild(dropdown);

    // Get unique values
    const filteredData = getFilteredData();
    const values = filteredData.map(row => row[column] ?? '').map(String);
    const columnTypes = detectColumnTypes(getOriginalData());
    const isDateColumn = columnTypes[column] === 'date';
    // --- selectedSet debe estar disponible en todo el scope ---
    let selectedSet = new Set(getTableFilterValues()[column] || []);

    // --- Lógica especial para fechas ---
    if (isDateColumn) {
        // Agrupar por año > mes > día usando parseFlexibleDate
        const dateTree = {};
        const originalsMap = {};
        values.forEach(val => {
            const date = parseFlexibleDate(val);
            if (!date) return;
            const year = date.getFullYear();
            const month = (date.getMonth() + 1).toString().padStart(2, '0');
            const day = date.getDate().toString().padStart(2, '0');
            if (!dateTree[year]) dateTree[year] = {};
            if (!dateTree[year][month]) dateTree[year][month] = {};
            if (!dateTree[year][month][day]) dateTree[year][month][day] = [];
            dateTree[year][month][day].push(val);
            // Mapear valor original a su fecha
            if (!originalsMap[year]) originalsMap[year] = {};
            if (!originalsMap[year][month]) originalsMap[year][month] = {};
            if (!originalsMap[year][month][day]) originalsMap[year][month][day] = [];
            originalsMap[year][month][day].push(val);
        });

        // Renderizado del árbol de fechas
        function renderDateTree(filter = '') {
            optionsDiv.innerHTML = '';
            const years = Object.keys(dateTree).sort();
            years.forEach(year => {
                if (filter && !year.includes(filter)) return;
                // Año
                const yearDiv = document.createElement('div');
                yearDiv.style.fontWeight = 'bold';
                yearDiv.style.margin = '0.2em 0';
                yearDiv.style.cursor = 'pointer';
                // Checkbox año
                const yearCheckbox = document.createElement('input');
                yearCheckbox.type = 'checkbox';
                yearCheckbox.value = year;
                // Todos los originales de ese año
                const allYearOriginals = Object.values(originalsMap[year]).flatMap(monthObj => Object.values(monthObj).flatMap(dayArr => dayArr));
                yearCheckbox.checked = allYearOriginals.every(d => selectedSet.has(d));
                yearCheckbox.indeterminate = !yearCheckbox.checked && allYearOriginals.some(d => selectedSet.has(d));
                yearCheckbox.addEventListener('change', () => {
                    if (yearCheckbox.checked) {
                        allYearOriginals.forEach(d => selectedSet.add(d));
                    } else {
                        allYearOriginals.forEach(d => selectedSet.delete(d));
                    }
                    renderDateTree(filter);
                });
                yearDiv.appendChild(yearCheckbox);
                yearDiv.appendChild(document.createTextNode(' ' + year));
                // Flecha expand/collapse
                let expandedYear = false;
                const monthsDiv = document.createElement('div');
                monthsDiv.style.display = 'none';
                monthsDiv.style.marginLeft = '1em';
                yearDiv.onclick = (e) => {
                    if (e.target !== yearCheckbox) {
                        expandedYear = !expandedYear;
                        monthsDiv.style.display = expandedYear ? 'block' : 'none';
                        arrow.textContent = expandedYear ? ' ▼' : ' ▶';
                    }
                };
                const arrow = document.createElement('span');
                arrow.textContent = ' ▶';
                yearDiv.appendChild(arrow);
                optionsDiv.appendChild(yearDiv);
                optionsDiv.appendChild(monthsDiv);
                // Meses
                const months = Object.keys(dateTree[year]).sort();
                months.forEach(month => {
                    if (filter && !month.includes(filter)) return;
                    const monthDiv = document.createElement('div');
                    monthDiv.style.fontWeight = 'normal';
                    monthDiv.style.margin = '0.1em 0';
                    monthDiv.style.cursor = 'pointer';
                    // Checkbox mes
                    const monthCheckbox = document.createElement('input');
                    monthCheckbox.type = 'checkbox';
                    monthCheckbox.value = `${year}-${month}`;
                    const allMonthOriginals = Object.values(originalsMap[year][month]).flatMap(dayArr => dayArr);
                    monthCheckbox.checked = allMonthOriginals.every(d => selectedSet.has(d));
                    monthCheckbox.indeterminate = !monthCheckbox.checked && allMonthOriginals.some(d => selectedSet.has(d));
                    monthCheckbox.addEventListener('change', () => {
                        if (monthCheckbox.checked) {
                            allMonthOriginals.forEach(d => selectedSet.add(d));
                        } else {
                            allMonthOriginals.forEach(d => selectedSet.delete(d));
                        }
                        renderDateTree(filter);
                    });
                    monthDiv.appendChild(monthCheckbox);
                    monthDiv.appendChild(document.createTextNode(' ' + month));
                    // Flecha expand/collapse
                    let expandedMonth = false;
                    const daysDiv = document.createElement('div');
                    daysDiv.style.display = 'none';
                    daysDiv.style.marginLeft = '1em';
                    monthDiv.onclick = (e) => {
                        if (e.target !== monthCheckbox) {
                            expandedMonth = !expandedMonth;
                            daysDiv.style.display = expandedMonth ? 'block' : 'none';
                            monthArrow.textContent = expandedMonth ? ' ▼' : ' ▶';
                        }
                    };
                    const monthArrow = document.createElement('span');
                    monthArrow.textContent = ' ▶';
                    monthDiv.appendChild(monthArrow);
                    monthsDiv.appendChild(monthDiv);
                    monthsDiv.appendChild(daysDiv);
                    // Días
                    const days = Object.keys(dateTree[year][month]).sort();
                    days.forEach(day => {
                        if (filter && !day.includes(filter)) return;
                        const dayDiv = document.createElement('div');
                        dayDiv.style.display = 'flex';
                        dayDiv.style.alignItems = 'center';
                        dayDiv.style.gap = '0.5rem';
                        dayDiv.style.padding = '0.1rem 0.5rem';
                        // Mostrar solo valores únicos (sin duplicados)
                        const uniqueOriginals = Array.from(new Set(dateTree[year][month][day]));
                        uniqueOriginals.forEach(origVal => {
                            const checkbox = document.createElement('input');
                            checkbox.type = 'checkbox';
                            checkbox.value = origVal;
                            checkbox.checked = selectedSet.has(origVal);
                            checkbox.addEventListener('change', () => {
                                if (checkbox.checked) {
                                    selectedSet.add(origVal);
                                } else {
                                    selectedSet.delete(origVal);
                                }
                                renderDateTree(filter);
                            });
                            const label = document.createElement('span');
                            label.textContent = origVal;
                            dayDiv.appendChild(checkbox);
                            dayDiv.appendChild(label);
                        });
                        daysDiv.appendChild(dayDiv);
                    });
                });
            });
        }
        renderDateTree();
        // Buscar por año, mes o día
        searchInput.addEventListener('input', debounce(() => {
            const term = searchInput.value.trim();
            renderDateTree(term);
        }, 150));

        // Seleccionar todo
        selectAllBtn.addEventListener('click', () => {
            // Selecciona todos los valores originales
            Object.values(originalsMap).forEach(yearObj => {
                Object.values(yearObj).forEach(monthObj => {
                    Object.values(monthObj).forEach(dayArr => {
                        dayArr.forEach(val => selectedSet.add(val));
                    });
                });
            });
            renderDateTree();
        });
        // Limpiar selección
        clearAllBtn.addEventListener('click', () => {
            selectedSet.clear();
            renderDateTree();
        });
                // Aplicar filtro
        applyBtn.addEventListener('click', () => {
            const filterArray = Array.from(selectedSet);
            setTableFilterValues({ ...getTableFilterValues(), [column]: filterArray });
            setTableActiveFilters({ ...getTableActiveFilters(), [column]: 'date' });
            dropdown.remove();
            applyFilters();
            refreshHeaderFilterIcons();
            // Marcar que se aplicó un filtro desde dropdown de tabla
            window.hasTableDropdownFilters = true;
            // Resume button color update removed
            window.dispatchEvent(new CustomEvent('filtersChanged'));
        });
        // Cerrar al hacer click fuera
        setTimeout(() => {
            document.addEventListener('mousedown', function handler(e) {
                if (!dropdown.contains(e.target) && e.target !== th) {
                    dropdown.remove();
                    document.removeEventListener('mousedown', handler);
                }
            });
        }, 10);
        return;
    }

    // --- Lógica normal para columnas no fecha ---
    // Eliminar duplicados y normalizar valores
    let uniqueValues = [...new Set(values
      .map(val => String(val).trim()) // Convertir a string y eliminar espacios
      .filter(val => val !== '') // Eliminar valores vacíos
    )];
    
    // Eliminar duplicados adicionales considerando normalización
    const normalizedSet = new Set();
    uniqueValues = uniqueValues.filter(val => {
      const normalized = val.toLowerCase().replace(/\s+/g, ' ').trim();
      if (normalizedSet.has(normalized)) {
        return false;
      }
      normalizedSet.add(normalized);
      return true;
    });
    
    // Ordenar alfabéticamente
    uniqueValues.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
    let filteredValues = uniqueValues.slice();
    function renderCheckboxList() {
        // LIMPIEZA COMPLETA Y ROBUSTA
        optionsDiv.innerHTML = '';
        
        const MAX_OPTIONS = 200;
        if (filteredValues.length > MAX_OPTIONS) {
            const msg = document.createElement('div');
            msg.textContent = 'Too many values to display. Please use the search box.';
            msg.style.color = '#F44336';
            msg.style.padding = '0.5em 0';
            optionsDiv.appendChild(msg);
            return;
        }
        // Opción (Empty)
        const emptyLabel = document.createElement('label');
        emptyLabel.style.display = 'flex';
        emptyLabel.style.alignItems = 'center';
        emptyLabel.style.gap = '0.5rem';
        emptyLabel.style.padding = '0.15rem 0.5rem';
        const emptyCheckbox = document.createElement('input');
        emptyCheckbox.type = 'checkbox';
        emptyCheckbox.value = '__EMPTY__';
        emptyCheckbox.checked = selectedSet.has('__EMPTY__');
        emptyCheckbox.addEventListener('change', () => {
            if (emptyCheckbox.checked) {
                selectedSet.add('__EMPTY__');
            } else {
                selectedSet.delete('__EMPTY__');
            }
            renderCheckboxList();
        });
        emptyLabel.appendChild(emptyCheckbox);
        emptyLabel.appendChild(document.createTextNode('(Empty)'));
        optionsDiv.appendChild(emptyLabel);
        // Resto de valores - VERIFICAR DUPLICADOS EN TIEMPO REAL
        const processedValues = new Set();
        filteredValues.forEach(val => {
            if (val === '' || processedValues.has(val)) return; // Ya cubierto por (Empty) o duplicado
            processedValues.add(val);
            const label = document.createElement('label');
            label.style.display = 'flex';
            label.style.alignItems = 'center';
            label.style.gap = '0.5rem';
            label.style.padding = '0.15rem 0.5rem';
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = val;
            checkbox.checked = selectedSet.has(val);
            checkbox.addEventListener('change', () => {
                if (checkbox.checked) {
                    selectedSet.add(val);
                } else {
                    selectedSet.delete(val);
                }
                renderCheckboxList();
            });
            label.appendChild(checkbox);
            label.appendChild(document.createTextNode(val));
            optionsDiv.appendChild(label);
        });
    }
    // Buscar y mostrar dropdown (con debounce)
    const handleInput = debounce(() => {
        const term = searchInput.value.trim().toLowerCase();
        filteredValues = uniqueValues.filter(val => val.toLowerCase().includes(term));
        renderCheckboxList();
    }, 150);
    searchInput.addEventListener('input', handleInput);
    // Seleccionar todo
    selectAllBtn.addEventListener('click', () => {
        selectedSet = new Set(filteredValues);
        renderCheckboxList();
    });
    // Limpiar selección
    clearAllBtn.addEventListener('click', () => {
        selectedSet.clear();
        renderCheckboxList();
    });
    // Aplicar filtro
    applyBtn.addEventListener('click', () => {
        const filterArray = Array.from(selectedSet);
        setTableFilterValues({ ...getTableFilterValues(), [column]: filterArray });
        setTableActiveFilters({ ...getTableActiveFilters(), [column]: isDateColumn ? 'date' : 'reference' });
        dropdown.remove();
        applyFilters();
        refreshHeaderFilterIcons();
        // Marcar que se aplicó un filtro desde dropdown de tabla
        window.hasTableDropdownFilters = true;
        // Resume button color update removed
        window.dispatchEvent(new CustomEvent('filtersChanged'));
    });
    // Inicializar lista
    renderCheckboxList();
    // Cerrar al hacer click fuera
    setTimeout(() => {
        document.addEventListener('mousedown', function handler(e) {
            if (!dropdown.contains(e.target) && e.target !== th) {
                dropdown.remove();
                document.removeEventListener('mousedown', handler);
            }
        });
    }, 10);
}

// Refuerzo: tras aplicar o limpiar filtro, refrescar header y badges
function refreshHeaderFilterIcons() {
    const filterValues = getTableFilterValues();
    let totalActiveFilters = 0;
    
    document.querySelectorAll('.data-table th').forEach(th => {
        const col = th.dataset.column;
        const icon = th.querySelector('.excel-filter-icon');
        if (!icon) return;
        // Elimina badge previo
        icon.querySelectorAll('.excel-filter-badge').forEach(b => b.remove());
        if (filterValues[col] && Array.isArray(filterValues[col]) && filterValues[col].length > 0) {
            icon.classList.add('excel-filter-active');
            totalActiveFilters++;
            // Badge verde
            const badge = document.createElement('span');
            badge.className = 'excel-filter-badge';
            badge.style.position = 'absolute';
            badge.style.top = '0px';
            badge.style.right = '0px';
            badge.style.width = '8px';
            badge.style.height = '8px';
            badge.style.background = '#10B981';
            badge.style.borderRadius = '50%';
            badge.style.border = '1.5px solid #fff';
            badge.style.display = 'block';
            badge.style.zIndex = '2';
            icon.appendChild(badge);
        } else {
            icon.classList.remove('excel-filter-active');
        }
    });
    
    // Actualizar el estado del botón de reset
    updateTableResetButtonState(totalActiveFilters);
}

// Función para actualizar el estado del botón de reset basado en filtros de tabla
function updateTableResetButtonState(count) {
    const resetButtons = document.querySelectorAll('.reset-filters-btn, #resetFiltersBtn, #resetAllFiltersBtn');
    resetButtons.forEach(btn => {
        if (count > 0) {
            btn.classList.add('has-active-filters');
        } else {
            btn.classList.remove('has-active-filters');
            // Forzar limpieza de estilos inline JS
            btn.style.background = '';
            btn.style.border = '';
            btn.style.color = '';
            btn.style.opacity = '';
        }
    });
}

// Función global para resetear todos los filtros
window.resetAllFilters = function() {
    if (typeof setTableActiveFilters === 'function' && typeof setTableFilterValues === 'function') {
        setTableActiveFilters({});
        setTableFilterValues({});
        applyFilters();
        refreshHeaderFilterIcons();
        // Resetear flag de filtros de dropdown de tabla
        window.hasTableDropdownFilters = false;
        // Resume button color update removed
        // Limpiar visualmente los filtros Excel de las cabeceras
        document.querySelectorAll('.excel-filter-dropdown').forEach(el => el.remove());
        document.querySelectorAll('.excel-filter-icon').forEach(icon => icon.classList.remove('excel-filter-active'));
        document.querySelectorAll('.excel-dropdown input[type="checkbox"]').forEach(cb => { cb.checked = false; });
        document.querySelectorAll('.excel-dropdown input[type="text"]').forEach(inp => { inp.value = ''; });
        document.querySelectorAll('.excel-dropdown .excel-checkbox-list label').forEach(lbl => lbl.classList.remove('active'));
        // Actualizar el estado del botón de reset
        updateTableResetButtonState(0);
        // Refresca la tabla para forzar el render limpio de iconos y filtros
        displayTable(getOriginalData());
        window.dispatchEvent(new CustomEvent('filtersChanged'));
    }
}

// Función para actualizar el botón de mostrar columnas ocultas
function updateHiddenColumnsButton() {
    let hiddenColumnsBtn = document.getElementById('showHiddenColumnsBtn');
    if (!hiddenColumnsBtn) {
        hiddenColumnsBtn = document.createElement('button');
        hiddenColumnsBtn.id = 'showHiddenColumnsBtn';
        hiddenColumnsBtn.className = 'toolbar-button';
        hiddenColumnsBtn.textContent = 'Show Hidden Columns';
        hiddenColumnsBtn.style.display = 'none';
        document.querySelector('.toolbar-right').appendChild(hiddenColumnsBtn);
    }
    if (hiddenColumns.size > 0) {
        hiddenColumnsBtn.style.display = 'flex';
        hiddenColumnsBtn.onclick = showHiddenColumnsDropdown;
    } else {
        hiddenColumnsBtn.style.display = 'none';
    }
}

// Función para mostrar el dropdown de columnas ocultas
function showHiddenColumnsDropdown() {
    document.querySelectorAll('.hidden-columns-dropdown').forEach(el => el.remove());
    const dropdown = document.createElement('div');
    dropdown.className = 'hidden-columns-dropdown';
    const button = document.getElementById('showHiddenColumnsBtn');
    const rect = button.getBoundingClientRect();
    dropdown.style.left = rect.left + window.scrollX + 'px';
    dropdown.style.top = rect.bottom + window.scrollY + 6 + 'px';
    const title = document.createElement('div');
    title.className = 'hidden-columns-title';
    title.textContent = 'Hidden Columns';
    dropdown.appendChild(title);
    // --- Botón Show All ---
    if (hiddenColumns.size > 1) {
        const showAllBtn = document.createElement('button');
        showAllBtn.textContent = 'Show All';
        showAllBtn.className = 'show-column-btn show-all-btn';
        showAllBtn.onclick = (e) => {
            e.stopPropagation();
            const allHidden = Array.from(hiddenColumns);
            // Mostrar TODAS las columnas (visibles + ocultas)
            const allHeaders = getCurrentHeaders();
            setVisibleColumns(allHeaders); // Mostrar todas las columnas
            hiddenColumns.clear(); // Limpiar la lista de ocultas
            if (typeof applyFilters === 'function') applyFilters();
            // Marcar todos los checkboxes como checked
            allHeaders.forEach(column => {
                const checkbox = document.querySelector(`#columnList input[type='checkbox'][value='${column.replace(/'/g, "\\'") }']`);
                if (checkbox) checkbox.checked = true;
            });
            updateHiddenColumnsButton();
            dropdown.remove();
            window.dispatchEvent(new CustomEvent('filtersChanged'));
        };
        dropdown.appendChild(showAllBtn);
    }
    const list = document.createElement('div');
    list.className = 'hidden-columns-list';
    Array.from(hiddenColumns).forEach(column => {
        const item = document.createElement('div');
        item.className = 'hidden-column-item';
        const columnName = document.createElement('span');
        columnName.className = 'hidden-column-name';
        columnName.textContent = column;
        item.appendChild(columnName);
        const showBtn = document.createElement('button');
        showBtn.className = 'show-column-btn';
        showBtn.textContent = 'Show';
        showBtn.type = 'button';
        showBtn.title = 'Show column';
        showBtn.onclick = (e) => {
            e.stopPropagation();
            hiddenColumns.delete(column);
            // Restaurar el orden original
            const allHeaders = getCurrentHeaders();
            const currentVisible = getVisibleColumns();
            const newVisible = allHeaders.filter(col =>
                col === column || currentVisible.includes(col)
            );
            setVisibleColumns(newVisible);
            if (typeof applyFilters === 'function') applyFilters();
            const checkbox = document.querySelector(`#columnList input[type='checkbox'][value='${column.replace(/'/g, "\\'") }']`);
            if (checkbox) checkbox.checked = true;
            updateHiddenColumnsButton();
            dropdown.remove();
            window.dispatchEvent(new CustomEvent('filtersChanged'));
        };
        item.appendChild(showBtn);
        list.appendChild(item);
    });
    dropdown.appendChild(list);
    document.body.appendChild(dropdown);
    // Animación fadeIn
    const style = document.createElement('style');
    style.innerHTML = `@keyframes fadeInDropdown { from { opacity: 0; transform: translateY(-8px);} to { opacity: 1; transform: translateY(0);} }`;
    document.head.appendChild(style);
    // Cerrar al hacer click fuera
    setTimeout(() => {
        document.addEventListener('mousedown', function handler(e) {
            if (!dropdown.contains(e.target) && e.target !== button) {
                dropdown.remove();
                document.removeEventListener('mousedown', handler);
            }
        });
    }, 10);
}

function generateExcelReport(data) {
    // ... existing code ...
    // En generateExcelReport, pon '[Logo]' en la primera fila del array ws_data y elimina la asignación manual a ws['A1'].v y ws['A1'].s.
    // ... existing code ...
}

export function colorRowsByUrgencia() {
  const table = document.querySelector('.data-table');
  if (!table) return;
  const rows = table.querySelectorAll('tbody tr');
  const activeCards = window.activeUrgencyCards || [];
  const cardColors = {
    'critical': '#ffcdd2',
    'urgente': '#ffcdd2', // Mantener compatibilidad con datos existentes
    'media': '#fff9c4',
    'baja': '#c8e6c9'
  };
  
  // --- NUEVO: Obtener los datos actuales de la tabla ---
  // Para tabs editables, necesitamos los datos actuales, no los originales
  let currentData = [];
  const startIndex = (getCurrentPage() - 1) * getRowsPerPage();
  const endIndex = startIndex + getRowsPerPage();
  
  // Intentar obtener datos de la tab editable actual
  const currentTabName = document.querySelector('.main-tab.active')?.textContent?.trim();
  if (currentTabName && window.editableTabData && window.editableTabData[currentTabName]) {
    currentData = window.editableTabData[currentTabName].data || [];
  } else {
    // Fallback a datos originales
    currentData = getOriginalData();
  }
  
  if (activeCards.length > 0) {
    // Usa el color de la primera tarjeta activa
    const color = cardColors[activeCards[0].toLowerCase()] || '#e0e0e0';
    rows.forEach((row, index) => {
      // --- NUEVO: Solo colorear si la fila NO tiene color personalizado ---
      const globalIndex = startIndex + index;
      
      // Si la fila tiene color personalizado (rowColor), NO sobrescribirlo
      if (currentData[globalIndex] && currentData[globalIndex].rowColor) {
        // Mantener el color personalizado
        row.style.background = currentData[globalIndex].rowColor;
      } else {
        // Solo aplicar color de urgencia si no hay color personalizado
        row.style.background = color;
      }
    });
  } else {
    // Si no hay ninguna activa, solo limpiar filas sin color personalizado
    rows.forEach((row, index) => {
      const globalIndex = startIndex + index;
      
      // Solo limpiar si NO tiene color personalizado
      if (!(currentData[globalIndex] && currentData[globalIndex].rowColor)) {
        row.style.background = '';
      }
    });
  }
}

// Tooltip visual universal para chips de filtro (mejorado)
function setupFilterTagTooltips() {
  let currentTooltip = null;
  let hideTimeout = null;

  function showTooltip(span, text) {
    if (currentTooltip) currentTooltip.remove();
    const tooltip = document.createElement('div');
    tooltip.className = 'filter-tooltip';
    tooltip.textContent = text;
    tooltip.style.position = 'fixed';
    const rect = span.getBoundingClientRect();
    tooltip.style.left = (rect.left + rect.width/2) + 'px';
    tooltip.style.top = (rect.bottom + 6) + 'px';
    tooltip.style.transform = 'translateX(-50%)';
            tooltip.style.background = '#1a2332';
    tooltip.style.color = '#fff';
    tooltip.style.padding = '0.45em 1em';
    tooltip.style.borderRadius = '7px';
    tooltip.style.fontSize = '0.98em';
    tooltip.style.whiteSpace = 'pre-line';
    tooltip.style.boxShadow = '0 4px 16px rgba(25, 118, 210, 0.13)';
    tooltip.style.zIndex = '99999';
    tooltip.style.minWidth = '120px';
    tooltip.style.maxWidth = '320px';
    tooltip.style.wordBreak = 'break-word';
    document.body.appendChild(tooltip);
    currentTooltip = tooltip;
    // Ocultar solo si el ratón sale de ambos
    tooltip.addEventListener('mouseenter', () => {
      if (hideTimeout) clearTimeout(hideTimeout);
    });
    tooltip.addEventListener('mouseleave', () => {
      if (currentTooltip) currentTooltip.remove();
      currentTooltip = null;
    });
  }

  document.body.addEventListener('mouseenter', function(e) {
    const span = e.target.closest('.filter-tag span');
    if (span && span.textContent.length > 20) {
      showTooltip(span, span.textContent);
    }
  }, true);

  document.body.addEventListener('mouseleave', function(e) {
    const span = e.target.closest('.filter-tag span');
    if (span && currentTooltip) {
      hideTimeout = setTimeout(() => {
        if (currentTooltip) currentTooltip.remove();
        currentTooltip = null;
      }, 120);
    }
  }, true);
}
setupFilterTagTooltips();

// --- MODAL INFORMATIVO GLOBAL ---
export function showInfoModal(title, message) {
    // Elimina cualquier modal anterior
    const old = document.getElementById('infoModalOverlay');
    if (old) old.remove();
    // Crea overlay
    const overlay = document.createElement('div');
    overlay.id = 'infoModalOverlay';
    overlay.style.position = 'fixed';
    overlay.style.top = '0';
    overlay.style.left = '0';
    overlay.style.width = '100vw';
    overlay.style.height = '100vh';
    overlay.style.background = 'rgba(16,24,32,0.18)';
    overlay.style.backdropFilter = '';
    overlay.style.display = 'flex';
    overlay.style.alignItems = 'flex-start';
    overlay.style.justifyContent = 'center';
    overlay.style.zIndex = '99999';
    // Crea modal
    const modal = document.createElement('div');
    modal.style.background = 'rgba(30,40,60,0.85)';
    modal.style.borderRadius = '12px';
    modal.style.boxShadow = '0 2px 8px rgba(25,118,210,0.10)';
    modal.style.padding = '0.9em 1.5em 0.7em 1.5em';
    modal.style.color = '#fff';
    modal.style.fontFamily = 'Inter,Segoe UI,Arial,sans-serif';
    modal.style.maxWidth = '90vw';
    modal.style.textAlign = 'center';
    modal.style.marginTop = '1em';
    modal.style.fontSize = '1em';
    modal.innerHTML = `<div style='margin:0.2em 0 0.2em 0;font-weight:500;'>${message}</div>`;
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    setTimeout(() => overlay.remove(), 1500);
}

// Hook global para refrescar el estado del botón de reset cuando cambian los filtros
window.addEventListener('filtersChanged', () => {
  if (typeof refreshHeaderFilterIcons === 'function') refreshHeaderFilterIcons();
});

// Forzar color verde por JS si el botón tiene la clase .has-active-filters
function forceResetBtnGreen() {
  const btn = document.getElementById('resetAllFiltersBtn');
  if (!btn) return;
  if (btn.classList.contains('has-active-filters')) {
    btn.style.background = '#10B981';
    btn.style.border = '1px solid #10B981';
    btn.style.color = '#fff';
    btn.style.opacity = '1';
  } else {
    btn.style.background = '';
    btn.style.border = '';
    btn.style.color = '';
    btn.style.opacity = '';
  }
}
window.addEventListener('filtersChanged', forceResetBtnGreen); 
window.addEventListener('filtersChanged', forceResetBtnGreen); 