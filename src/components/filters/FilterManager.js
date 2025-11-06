import { 
    getOriginalData, 
    setModuleActiveFilters, 
    setModuleFilterValues, 
    getModuleActiveFilters, 
    getModuleFilterValues,
    getTableActiveFilters,
    getTableFilterValues,
    getSortConfig,
    getActiveFilters,
    setActiveFilters
} from '../../store/index.js';
import { displayTable } from '../table/Table.js';
import { sortData } from '../../utils/general.js';
import { DataDetector } from '../../utils/DataDetector.js';

// Get singleton instance
const dataDetector = DataDetector.getInstance();

// Singleton instance
let instance = null;

// Cache for unique values
const uniqueValuesCache = new Map();

// Función para actualizar valores de filtros existentes sin regenerar
function updateExistingFilterValues() {
  try {
    console.log('🔄 Updating existing filter values...');
    
    // Actualizar valores únicos en dropdowns existentes
    const filterItems = document.querySelectorAll('.filter-item');
    filterItems.forEach(item => {
      const column = item.dataset.column;
      if (!column) return;
      
      // Actualizar dropdowns de valores únicos
      const dropdown = item.querySelector('.filter-select');
      if (dropdown) {
        const currentValue = dropdown.value;
        const uniqueValues = getUniqueValues(column);
        
        // Solo actualizar si los valores han cambiado
        const currentOptions = Array.from(dropdown.options).map(opt => opt.value);
        const newOptions = uniqueValues.map(val => val.toString());
        
        if (JSON.stringify(currentOptions) !== JSON.stringify(newOptions)) {
          dropdown.innerHTML = '<option value="">All</option>';
          uniqueValues.forEach(value => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            dropdown.appendChild(option);
          });
          
          // Restaurar valor seleccionado si aún existe
          if (currentValue && newOptions.includes(currentValue)) {
            dropdown.value = currentValue;
          }
        }
      }
    });
    
    console.log('✅ Existing filter values updated');
  } catch (error) {
    console.error('❌ Error updating existing filter values:', error);
  }
}

// Función para obtener valores únicos de una columna
function getUniqueValues(column) {
  try {
    const data = getOriginalData();
    if (!data || !data.length) return [];
    
    // Usar caché si está disponible
    if (uniqueValuesCache.has(column)) {
      return uniqueValuesCache.get(column);
    }
    
    const uniqueSet = new Set();
    data.forEach(row => {
      const value = row[column];
      if (value !== null && value !== undefined && value.toString().trim() !== '') {
        uniqueSet.add(value.toString());
      }
    });
    
    const uniqueValues = Array.from(uniqueSet).sort();
    
    // Guardar en caché
    uniqueValuesCache.set(column, uniqueValues);
    
    return uniqueValues;
  } catch (error) {
    console.error('❌ Error getting unique values for column:', column, error);
    return [];
  }
}

// Initialize data
function initializeData(data) {
  if (!data || !data.length) {
    console.warn('No data provided for filter initialization');
    return false;
  }
  try {
    // Limpiar el caché de valores únicos al cargar nuevos datos
    if (uniqueValuesCache && typeof uniqueValuesCache.clear === 'function') {
      uniqueValuesCache.clear();
    }
    return true;
  } catch (error) {
    console.error('Error initializing filter data:', error);
    return false;
  }
}

// Detect column types based on data sample
function detectColumnTypes(data) {
  if (!data || !data.length) {
    console.warn('No data provided for column type detection');
    return {};
  }
  try {
    return dataDetector.detectColumnTypes(data);
  } catch (error) {
    console.error('Error detecting column types:', error);
    return {};
  }
}

// Get most frequent values for a column
function getFrequentValues(column, minCount = 5, maxItems = 10) {
  const freqMap = new Map();
  const data = getOriginalData();
  data.forEach(row => {
    const value = row[column];
    if (value !== null && value !== undefined && value.toString().trim() !== '') {
      const key = value.toString();
      freqMap.set(key, (freqMap.get(key) || 0) + 1);
    }
  });
  // Ordenar por frecuencia descendente
  const sorted = Array.from(freqMap.entries())
    .filter(([_, count]) => count >= minCount)
    .sort((a, b) => b[1] - a[1]);
  // Si no hay suficientes, muestra el top N aunque sean menos frecuentes
  const result = sorted.length > 0
    ? sorted.slice(0, maxItems).map(([val]) => val)
    : Array.from(freqMap.entries()).sort((a, b) => b[1] - a[1]).slice(0, maxItems).map(([val]) => val);
  return result;
}

// FUNCIÓN PARA FORZAR ESTILOS INLINE EN DROPDOWNS
function forceDropdownStyles(dropdown) {
  if (!dropdown) return;
  
  // Forzar estilo del dropdown principal
  dropdown.style.setProperty('background', '#fff', 'important');
  dropdown.style.setProperty('color', '#47B2E5', 'important');
  
  // Forzar estilos en todos los elementos
  const allElements = dropdown.querySelectorAll('*');
  allElements.forEach(el => {
    if (el.tagName !== 'INPUT' || el.type !== 'checkbox') {
      el.style.setProperty('color', '#47B2E5', 'important');
      el.style.setProperty('text-shadow', 'none', 'important');
    }
  });
  
  // Forzar estilos específicos en labels
  const labels = dropdown.querySelectorAll('label');
  labels.forEach(label => {
    label.style.setProperty('color', '#47B2E5', 'important');
    label.style.setProperty('text-shadow', 'none', 'important');
    label.style.setProperty('background', 'transparent', 'important');
  });
  
  // Forzar estilos en spans
  const spans = dropdown.querySelectorAll('span');
  spans.forEach(span => {
    span.style.setProperty('color', '#47B2E5', 'important');
    span.style.setProperty('text-shadow', 'none', 'important');
    span.style.setProperty('background', 'transparent', 'important');
  });
  
  // Forzar estilos en botones (excepto Apply)
  const buttons = dropdown.querySelectorAll('button:not(.apply-btn):not(.filter-apply-btn)');
  buttons.forEach(btn => {
    btn.style.setProperty('color', '#47B2E5', 'important');
    btn.style.setProperty('background', '#f8fafc', 'important');
    btn.style.setProperty('border', '1px solid #47B2E5', 'important');
  });
  
  // Botones Apply en verde
  const applyButtons = dropdown.querySelectorAll('.apply-btn, .filter-apply-btn');
  applyButtons.forEach(btn => {
    btn.style.setProperty('color', '#fff', 'important');
    btn.style.setProperty('background', '#10B981', 'important');
    btn.style.setProperty('border-color', '#10B981', 'important');
  });
}

// Debounced search function
function debounce(func, wait) {
  let timeout;
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}

// Detect if a column is categorical (has a value repeated more than threshold)
function isCategoricalColumn(column, threshold = 30) {
  // Usar los datos filtrados (visibles) en vez de todos los originales
  const dataToCheck = (typeof getFilteredData() !== 'undefined' && getFilteredData().length) ? getFilteredData() : getOriginalData();
  const freqMap = new Map();
  let maxCount = 0;
  let maxValue = null;
  for (const row of dataToCheck) {
    const value = row[column];
    if (value !== null && value !== undefined && value.toString().trim() !== '') {
      const key = value.toString();
      const count = (freqMap.get(key) || 0) + 1;
      freqMap.set(key, count);
      if (count > maxCount) {
        maxCount = count;
        maxValue = key;
      }
      if (count > threshold) {
        console.log(`[Filtro] Columna '${column}': valor más repetido '${maxValue}' (${maxCount} repeticiones, umbral ${threshold}, filas analizadas: ${dataToCheck.length})`);
        return true;
      }
    }
  }
  console.log(`[Filtro] Columna '${column}': valor más repetido '${maxValue}' (${maxCount} repeticiones, umbral ${threshold}, filas analizadas: ${dataToCheck.length})`);
  return false;
}

function createAutocomplete(input, uniqueValues) {
  let suggestionBox = input.nextSibling;
  if (!suggestionBox || !suggestionBox.classList || !suggestionBox.classList.contains('autocomplete-list')) {
    suggestionBox = document.createElement('ul');
    suggestionBox.className = 'autocomplete-list hidden';
    if (input.parentNode) {
      input.parentNode.insertBefore(suggestionBox, input.nextSibling);
    }
  }

  // Función para filtrar y mostrar sugerencias
  function showSuggestions(searchTerm) {
    const raw = input.value;
    // Obtener el último término después de la última coma
    const lastTerm = raw.split(',').pop().trim();
    const term = normalizeText(searchTerm || lastTerm);
    
    suggestionBox.innerHTML = '';
    if (!term) {
      suggestionBox.classList.add('hidden');
      return;
    }

    // Filtrar valores que ya están seleccionados
    const selectedValues = raw.split(',').map(v => v.trim()).filter(v => v);
    const matches = uniqueValues
      .filter(val => {
        const normalizedVal = normalizeText(val);
        return normalizedVal.includes(term) && !selectedValues.includes(val);
      })
      .slice(0, 10);

    if (matches.length === 0) {
      suggestionBox.classList.add('hidden');
      return;
    }

    matches.forEach(val => {
      const li = document.createElement('li');
      li.textContent = val;
      li.className = 'autocomplete-item';
      li.addEventListener('mousedown', (e) => {
        e.preventDefault();
        // Reemplazar el último término con el valor seleccionado
        const parts = raw.split(',');
        parts[parts.length - 1] = val;
        input.value = parts.join(',') + ', ';
        
        // Actualizar el estado del filtro
        const selectedColumn = input.closest('.filter-block').querySelector('.column-selector').value;
        const selectedValues = input.value.split(',').map(v => v.trim()).filter(v => v);
        setModuleFilterValues({ ...getModuleFilterValues(), [selectedColumn]: selectedValues });
        setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: 'categorical' });
        
        // Actualizar la UI
        const filterDiv = input.closest('.filter-block');
        filterDiv.classList.toggle('active', selectedValues.length > 0);
        
        // Actualizar el checkbox correspondiente en el dropdown
        const dropdown = filterDiv.querySelector('.excel-dropdown');
        if (dropdown) {
          const checkbox = dropdown.querySelector(`input[type="checkbox"][value="${val}"]`);
          if (checkbox) {
            checkbox.checked = true;
            // Disparar el evento change para asegurar que se actualice el estado
            checkbox.dispatchEvent(new Event('change', { bubbles: true }));
          }
        }
        
        updateActiveFiltersSummary();
        applyFilters();
        
        input.dispatchEvent(new Event('input'));
        suggestionBox.classList.add('hidden');
      });
      suggestionBox.appendChild(li);
    });
    suggestionBox.classList.remove('hidden');
  }

  // Mostrar sugerencias al escribir
  input.addEventListener('input', () => {
    showSuggestions();
  });

  // Mostrar sugerencias al hacer focus
  input.addEventListener('focus', () => {
    showSuggestions();
  });

  // Ocultar sugerencias al hacer blur
  input.addEventListener('blur', () => {
    setTimeout(() => suggestionBox.classList.add('hidden'), 150);
  });

  // Manejar navegación con teclado
  input.addEventListener('keydown', (e) => {
    const items = suggestionBox.querySelectorAll('.autocomplete-item');
    const activeItem = suggestionBox.querySelector('.autocomplete-item.active');
    let index = Array.from(items).indexOf(activeItem);

    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        if (index < items.length - 1) {
          index++;
        } else {
          index = 0;
        }
        break;
      case 'ArrowUp':
        e.preventDefault();
        if (index > 0) {
          index--;
        } else {
          index = items.length - 1;
        }
        break;
      case 'Enter':
        e.preventDefault();
        if (activeItem) {
          activeItem.click();
        }
        return;
      case 'Escape':
        suggestionBox.classList.add('hidden');
        return;
      default:
        return;
    }

    items.forEach(item => item.classList.remove('active'));
    items[index]?.classList.add('active');
  });
}

// Variable para evitar llamadas duplicadas
let lastGeneratedHeaders = null;
let isGenerating = false;

function generateFilterSidebar(headers) {
  console.log('🔍 generateFilterSidebar called with headers:', headers);
  const originalData = getOriginalData();
  console.log('📊 Original data available:', originalData?.length || 0, 'rows');
  
  if (!originalData || !originalData.length) {
    console.warn('No data available for generating filters');
    return;
  }
  
  if (!headers || !headers.length) {
    console.warn('No headers provided for generating filters');
    return;
  }
  
  // Evitar llamadas duplicadas con los mismos headers
  const headersString = JSON.stringify(headers?.sort());
  if (isGenerating || headersString === lastGeneratedHeaders) {
    console.log('Skipping duplicate generateFilterSidebar call');
    return;
  }
  
  isGenerating = true;
  lastGeneratedHeaders = headersString;
  
  // Timeout de seguridad para resetear la bandera
  setTimeout(() => {
    isGenerating = false;
  }, 5000);

  const genericContainer = document.getElementById("genericFilterPanel");
  const myFiltersPanel = document.getElementById("myfiltersFilterPanel");
  const referenceContainer = document.getElementById("referenceFilterPanel");
  const dateContainer = document.getElementById("dateFilterPanel");
  const filterTabs = document.querySelector('.filter-tabs');
  const filterPanels = document.querySelector('.filter-panels');

  console.log('🔍 Filter containers found:', {
    genericContainer: !!genericContainer,
    myFiltersPanel: !!myFiltersPanel,
    referenceContainer: !!referenceContainer,
    dateContainer: !!dateContainer,
    filterTabs: !!filterTabs,
    filterPanels: !!filterPanels
  });

  if (!genericContainer || !referenceContainer || !dateContainer || !filterTabs || !filterPanels || !myFiltersPanel) {
    console.warn('Filter containers or tabs not found, retrying in 100ms...');
    // Retry after a short delay in case DOM is not ready
    setTimeout(() => {
      generateFilterSidebar(headers);
    }, 100);
    return;
  }

  try {
    // Verificar si ya existen filtros generados para estos headers
    const existingFilters = genericContainer.querySelectorAll('.filter-item');
    const hasExistingFilters = existingFilters.length > 0;
    
    // Siempre limpiar y regenerar para nuevos datos
    console.log('🔧 Generating filters for new data...');
    
    // Limpiar contenido de cada panel antes de añadir el suyo
    genericContainer.innerHTML = "";
    myFiltersPanel.innerHTML = "";
    referenceContainer.innerHTML = "";
    dateContainer.innerHTML = "";
    
    // Limpiar cualquier filtro duplicado existente
    document.querySelectorAll('.filter-item').forEach(item => {
      if (item.parentNode === genericContainer || item.parentNode === dateContainer) {
        item.remove();
      }
    });

    // Remove any existing global summary and create a new one
    document.querySelectorAll('#globalActiveFiltersSummary').forEach(summary => summary.remove());
    
    // Create new global summary
    const globalSummary = document.createElement('div');
    globalSummary.id = 'globalActiveFiltersSummary';
    globalSummary.className = 'active-filters-summary';
    filterTabs.parentNode.insertBefore(globalSummary, filterPanels);

    // --- Add Active Filters Tab and Panel ---
    // Remove any existing active filters tabs and panels first
    document.querySelectorAll('.filter-tab[data-target="active"]').forEach(tab => tab.remove());
    document.querySelectorAll('#activeFilterPanel').forEach(panel => panel.remove());
    
    // Create new active filters tab
    const activeTab = document.createElement('button');
    activeTab.className = 'filter-tab';
    activeTab.dataset.target = 'active';
    activeTab.textContent = 'Active Filters';
    filterTabs.appendChild(activeTab);
    
    // Create new active filters panel
    const activePanel = document.createElement('div');
    activePanel.className = 'filter-panel';
    activePanel.id = 'activeFilterPanel';
    filterPanels.appendChild(activePanel);

    // Create summary inside the panel
    const summaryContainer = document.createElement("div");
    summaryContainer.className = "active-filters-summary";
    summaryContainer.innerHTML = `
      <div class="active-filters-list"></div>
      <button class="clear-all-filters-btn">Clear All</button>
    `;
    activePanel.appendChild(summaryContainer);

    // Render chips de filtros activos (incluyendo Today/dinámicos)
    function renderActiveFiltersList() {
      const list = summaryContainer.querySelector('.active-filters-list');
      if (!list) return;
      list.innerHTML = '';
      const filterValues = getModuleFilterValues();
      // Si no hay filtros activos, oculta completamente la sección
      if (Object.keys(filterValues).length === 0) {
        summaryContainer.style.display = 'none';
        summaryContainer.style.padding = '0';
        summaryContainer.style.margin = '0';
        summaryContainer.style.height = '0';
        summaryContainer.style.minHeight = '0';
      } else {
        summaryContainer.style.display = '';
        summaryContainer.style.padding = '';
        summaryContainer.style.margin = '';
        summaryContainer.style.height = '';
        summaryContainer.style.minHeight = '';
      }
      // Agrupa fechas
      const dateFilters = {};
      Object.entries(filterValues).forEach(([key, value]) => {
        if (key.endsWith('_start') || key.endsWith('_end') || key.endsWith('_empty')) {
          const baseKey = key.replace(/_(start|end|empty)$/, '');
          if (!dateFilters[baseKey]) {
            dateFilters[baseKey] = { start: '', end: '', empty: false };
          }
          if (key.endsWith('_start')) dateFilters[baseKey].start = value;
          if (key.endsWith('_end')) dateFilters[baseKey].end = value;
          if (key.endsWith('_empty')) dateFilters[baseKey].empty = value;
        }
      });
      // Chips de fechas
      Object.entries(dateFilters).forEach(([column, filter]) => {
        if (filter.start || filter.end || filter.empty) {
          let text = `${column}: `;
          if (filter.empty && !filter.start && !filter.end) {
            text += '(empty)';
          } else {
            if (filter.start) text += `from ${prettyDynamicDate(filter.start)}`;
            if (filter.end) text += ` to ${prettyDynamicDate(filter.end)}`;
            if (filter.empty) text += ' (including empty)';
          }
          const tag = document.createElement('div');
          tag.className = 'modal-filter-tag';
          tag.innerHTML = `<span>${text}</span><button class="modal-filter-tag-remove" data-column="${column}">×</button>`;
          list.appendChild(tag);
        }
      });
      // Chips de otros filtros
      Object.entries(filterValues).forEach(([key, value]) => {
        if (!key.endsWith('_start') && !key.endsWith('_end') && !key.endsWith('_empty') && !key.endsWith('_condition')) {
          // Mostrar tanto arrays como strings (valores únicos del input de texto)
          if (Array.isArray(value) && value.length > 0) {
            const tag = document.createElement('div');
            tag.className = 'modal-filter-tag';
            // Obtener la condición para esta columna
            const condition = filterValues[`${key}_condition`] || 'contains';
            const conditionLabel = getConditionLabel(condition);
            const conditionText = conditionLabel ? `${conditionLabel} ` : '';
            tag.innerHTML = `<span>${conditionText}${key}: ${value.join(', ')}</span><button class="modal-filter-tag-remove" data-column="${key}">×</button>`;
            list.appendChild(tag);
          } else if (typeof value === 'string' && value.trim() !== '') {
            // Mostrar valores de tipo string (del input de texto)
            const tag = document.createElement('div');
            tag.className = 'modal-filter-tag';
            // Obtener la condición para esta columna
            const condition = filterValues[`${key}_condition`] || 'contains';
            const conditionLabel = getConditionLabel(condition);
            const conditionText = conditionLabel ? `${conditionLabel} ` : '';
            tag.innerHTML = `<span>${conditionText}${key}: ${value}</span><button class="modal-filter-tag-remove" data-column="${key}">×</button>`;
            list.appendChild(tag);
          }
        }
      });
      // Listeners para eliminar chips
      // Declarar activeFilters fuera del bucle para evitar duplicidad
      let activeFilters;
      list.querySelectorAll('.modal-filter-tag-remove').forEach(btn => {
        btn.addEventListener('click', () => {
          const column = btn.dataset.column;
          const updated = { ...getModuleFilterValues() };
          // Eliminar todos los valores posibles del filtro
          delete updated[`${column}_start`];
          delete updated[`${column}_end`];
          delete updated[`${column}_empty`];
          delete updated[column];
          delete updated[`${column}_condition`];
          setModuleFilterValues(updated);
          // Eliminar de activeFilters SIEMPRE
          activeFilters = { ...getModuleActiveFilters() };
          delete activeFilters[column];
          setModuleActiveFilters(activeFilters);
          // Limpiar inputs visuales
          const filterItem = document.querySelector(`.filter-item[data-column="${column}"]`);
          if (filterItem) {
            filterItem.querySelectorAll('input[type="date"]').forEach(input => input.value = '');
            filterItem.querySelectorAll('.filter-checkbox').forEach(checkbox => checkbox.checked = false);
            filterItem.classList.remove('active');
          }
          // Refuerzo: renderizar chips y lista tras limpiar el estado
          renderActiveFiltersList();
          renderActiveFiltersSummaryChips();
          updateActiveFiltersSummary();
          applyFilters();
        });
      });
    }
    // Llama al render de chips activos al crear el panel
    renderActiveFiltersList();

    // --- Botón Save Filter SOLO en Active Filters ---
    let saveBtn = document.getElementById('saveMyFilterBtn');
    if (!saveBtn) {
      saveBtn = document.createElement('button');
      saveBtn.id = 'saveMyFilterBtn';
      saveBtn.className = 'btn btn-secondary';
      saveBtn.textContent = 'Save Filter';
      // Mensaje visual de confirmación/error
      let saveMsg = document.getElementById('saveFilterMsg');
      if (!saveMsg) {
        saveMsg = document.createElement('div');
        saveMsg.id = 'saveFilterMsg';
        saveMsg.style.display = 'none';
        saveMsg.style.marginTop = '0.5rem';
        saveMsg.style.fontWeight = 'bold';
        saveMsg.style.fontSize = '1rem';
        saveBtn.parentNode?.insertBefore?.(saveMsg, saveBtn.nextSibling);
      }
      function showSaveMsg(text, color) {
        saveMsg.textContent = text;
        saveMsg.style.color = color;
        saveMsg.style.display = 'block';
        setTimeout(() => { saveMsg.style.display = 'none'; }, 2500);
      }
      saveBtn.addEventListener('click', () => {
        // Check if user can modify data (admin only)
        if (typeof canUserModifyData === 'function' && !canUserModifyData()) {
          showSaveMsg('Only administrators can save filters', 'red');
          return;
        }
        
        const name = prompt('Enter a name for this filter:');
        if (name) {
          // Dropdown para asociar a tarjeta de urgencia
          let urgency = prompt('¿Asociar a tarjeta de urgencia? (Ninguna, Urgente, Media, Baja)', 'Ninguna');
          if (!urgency) urgency = 'Ninguna';
          try {
            saveMyFilter(name, urgency);
            // Forzar creación y renderizado de la sección My Filters
            let myFiltersPanel = document.getElementById('myfiltersFilterPanel');
            if (!myFiltersPanel) {
              myFiltersPanel = document.createElement('div');
              myFiltersPanel.id = 'myfiltersFilterPanel';
              myFiltersPanel.className = 'filter-panel';
              const filterPanels = document.querySelector('.filter-panels');
              if (filterPanels) filterPanels.appendChild(myFiltersPanel);
            }
            renderMyFiltersSection();
            showSaveMsg('Filter saved!', 'green');
          } catch (e) {
            showSaveMsg('Error saving filter', 'red');
          }
        }
      });
    }
    // --- Botón Guardar como filtro rápido ---
    let saveQuickBtn = document.getElementById('saveQuickFilterBtn');
    if (!saveQuickBtn) {
      saveQuickBtn = document.createElement('button');
      saveQuickBtn.id = 'saveQuickFilterBtn';
      saveQuickBtn.className = 'btn btn-secondary';
      saveQuickBtn.textContent = 'Save as quick filter';
      saveQuickBtn.addEventListener('click', () => {
        // Check if user can modify data (admin only)
        if (typeof canUserModifyData === 'function' && !canUserModifyData()) {
          showSaveMsg('Only administrators can save quick filters', 'red');
          return;
        }
        
        // Crear modal para guardar quick filter
        let modal = document.getElementById('saveQuickFilterModal');
        if (!modal) {
          modal = document.createElement('div');
          modal.id = 'saveQuickFilterModal';
          modal.className = 'modal-overlay hidden';
          modal.style.zIndex = '10001';
          modal.innerHTML = `
            <div class="modal-panel" style="max-width:600px;">
              <div class="modal-header">
                <div class="header-left">
                  <img src="LOGOTAB_rounded.png" alt="Logo" class="modal-logo">
                  <h3 class="panel-header-title">Save Quick Filter</h3>
                </div>
                <button id="closeSaveQuickFilterBtn" class="close-btn">×</button>
              </div>
              <div class="modal-content">
                <div style='margin-bottom:1.5em;'>
                  <label style='font-weight:600;color:white;font-family:Inter,Segoe UI,Arial,sans-serif;margin-bottom:0.5rem;display:block;'>Name:</label>
                  <input id='quickFilterNameInput' type='text' class="input" placeholder="Enter filter name" style='width:100%;margin-bottom:1rem;'>
                  <label style='font-weight:600;color:white;font-family:Inter,Segoe UI,Arial,sans-serif;margin-bottom:0.5rem;display:block;'>Associate to urgency card:</label>
                  <select id='quickFilterUrgencySelect' class="input filter-select" style='width:100%;margin-bottom:1rem;'>
                    <option value='Ninguna'>None</option>
                    <option value='Urgente'>Urgent</option>
                    <option value='Media'>Medium</option>
                    <option value='Baja'>Low</option>
                  </select>
                  <label style='font-weight:600;color:white;font-family:Inter,Segoe UI,Arial,sans-serif;margin-bottom:0.5rem;display:block;'>Save to container:</label>
                  <select id='quickFilterContainerSelect' class="input filter-select" style='width:100%;margin-bottom:1rem;'>
                    <option value=''>None</option>
                    <option value='default'>General Zone</option>
                    <option value='container1'>Order Management</option>
                    <option value='container2'>Booking Management</option>
                    <option value='container3'>Closing Operations</option>
                    <option value='container4'>Cargo Status</option>
                  </select>
                  <label style='font-weight:600;color:white;font-family:Inter,Segoe UI,Arial,sans-serif;margin-bottom:0.5rem;display:block;'>Container title (optional):</label>
                  <input id='quickFilterContainerTitleInput' type='text' class="input" placeholder='Enter custom container title' style='width:100%;margin-bottom:1rem;'>
                  <label style='font-weight:600;color:white;font-family:Inter,Segoe UI,Arial,sans-serif;margin-bottom:0.5rem;display:block;'>Save to Hub:</label>
                  <select id='quickFilterHubSelect' class="input filter-select" style='width:100%;'>
                    <option value='ops'>Operations Hub</option>
                    <option value='dq'>Data Quality Hub</option>
                  </select>
                </div>
              </div>
              <div class="modal-footer">
                <button id='quickFilterCancelBtn' class="modal-btn secondary">Cancel</button>
                <button id='quickFilterSaveBtn' class="modal-btn primary">Save</button>
              </div>
            </div>
          `;
          document.body.appendChild(modal);
        } else {
          modal.classList.remove('hidden');
        }
        
        // Limpiar campos
        modal.querySelector('#quickFilterNameInput').value = '';
        modal.querySelector('#quickFilterUrgencySelect').value = 'Ninguna';
        modal.querySelector('#quickFilterContainerSelect').value = 'default';
        modal.querySelector('#quickFilterContainerTitleInput').value = '';
        modal.querySelector('#quickFilterHubSelect').value = 'ops';
        
        // Cancelar
        modal.querySelector('#quickFilterCancelBtn').onclick = () => {
          modal.classList.add('hidden');
        };
        
        // Cerrar con el botón X
        modal.querySelector('#closeSaveQuickFilterBtn').onclick = () => {
          modal.classList.add('hidden');
        };
        
        // Cerrar al hacer click fuera del panel
        modal.onclick = (e) => {
          if (e.target === modal) {
            modal.classList.add('hidden');
          }
        };
        
        // Cambiar contenedores según el hub seleccionado
        const hubSelect = modal.querySelector('#quickFilterHubSelect');
        const containerSelect = modal.querySelector('#quickFilterContainerSelect');
        
        // Inicializar con los contenedores del Ops Hub (por defecto)
        const initializeContainers = (hubType) => {
          containerSelect.innerHTML = '<option value="">None</option>';
          
          if (hubType === 'ops') {
            // Contenedores del Ops Hub
            containerSelect.innerHTML += `
              <option value="default">General Zone</option>
              <option value="container1">Order Management</option>
              <option value="container2">Booking Management</option>
              <option value="container3">Closing Operations</option>
              <option value="container4">Cargo Status</option>
            `;
          } else if (hubType === 'dq') {
            // Contenedores del DQ Hub
            containerSelect.innerHTML += `
              <option value="dq-default">Data Quality Zone</option>
              <option value="dq-container1">Duplicate Analysis</option>
              <option value="dq-container2">Null Values</option>
              <option value="dq-container3">Format Issues</option>
              <option value="dq-container4">Completeness</option>
            `;
          }
        };
        
        // Inicializar con Ops Hub por defecto
        initializeContainers('ops');
        
        hubSelect.addEventListener('change', () => {
          const selectedHub = hubSelect.value;
          initializeContainers(selectedHub);
        });
        
        // Guardar
        modal.querySelector('#quickFilterSaveBtn').onclick = () => {
          // Check if user can modify data (admin only)
          if (typeof canUserModifyData === 'function' && !canUserModifyData()) {
            alert('Only administrators can save quick filters');
            return;
          }
          
          const name = modal.querySelector('#quickFilterNameInput').value.trim();
          const urgency = modal.querySelector('#quickFilterUrgencySelect').value;
          const container = modal.querySelector('#quickFilterContainerSelect').value;
          const containerTitle = modal.querySelector('#quickFilterContainerTitleInput').value.trim();
          const hubType = modal.querySelector('#quickFilterHubSelect').value;
          if (!name) {
            alert('Please enter a name for the quick filter.');
            return;
          }
          saveQuickFilter(name, urgency, container, containerTitle, hubType);
          modal.classList.add('hidden');
          showSaveMsg(`Quick filter saved to ${hubType.toUpperCase()} Hub!`, 'green');
        };
      });
    }
    // Contenedor flex para los dos botones
    const saveBtnsRow = document.createElement('div');
    saveBtnsRow.className = 'save-filter-btns-row';
    saveBtn.classList.add('filter-save-btn');
    saveQuickBtn.classList.add('filter-save-btn');
    saveBtnsRow.appendChild(saveBtn);
    saveBtnsRow.appendChild(saveQuickBtn);
    activePanel.appendChild(saveBtnsRow);

    // Add clear all functionality
    summaryContainer.querySelector('.clear-all-filters-btn').addEventListener('click', () => {
      try {
        // Clear all filter values
        setModuleFilterValues({});
        setModuleActiveFilters({});
        
        // Clear all input fields
        document.querySelectorAll('.filter-input').forEach(input => { 
          input.value = ''; 
          input.removeAttribute('data-dynamic');
        });
        document.querySelectorAll('.filter-input[type="date"]').forEach(input => { 
          input.value = ''; 
          input.removeAttribute('data-dynamic');
        });
        document.querySelectorAll('.filter-checkbox').forEach(checkbox => { checkbox.checked = false; });
        document.querySelectorAll('.column-selector').forEach(select => { select.value = ''; });
        
        // Remove active class from all filter items (don't remove the items themselves)
        document.querySelectorAll('.filter-item').forEach(item => { 
          item.classList.remove('active'); 
        });
        
        // Clear active filters summary
        const activeFiltersList = summaryContainer.querySelector('.active-filters-list');
        if (activeFiltersList) {
          activeFiltersList.innerHTML = '';
        }
        
        // Update UI and apply filters
        updateActiveFiltersSummary();
        renderActiveFiltersSummaryChips();
        applyFilters();
        
        if (typeof window.showUnifiedNotification === 'function') {
          window.showUnifiedNotification('All filters have been cleared!', 'info');
        }
      } catch (error) {
        console.error('Error clearing all filters:', error);
        if (typeof window.showUnifiedNotification === 'function') {
          window.showUnifiedNotification('Error clearing filters. Please try again.', 'error');
        }
      }
    });

    // --- END Active Filters Tab/Panel ---

    // Group headers by type
    const columnTypes = detectColumnTypes(originalData);
    const groupedHeaders = {
      date: [],
      categorized: []
    };

    headers.forEach(col => {
      const type = columnTypes[col] || 'generic';
      if (type === 'date') {
        groupedHeaders.date.push(col);
      } else {
        groupedHeaders.categorized.push(col);
      }
    });
    
    console.log('📋 Column types detected:', columnTypes);
    console.log('📂 Grouped headers:', groupedHeaders);

    // Create filter sections
    Object.entries(groupedHeaders).forEach(([type, cols]) => {
      const container = type === 'date' ? dateContainer : genericContainer;

      // Grid para los filtros
      const filterGrid = document.createElement('div');
      filterGrid.className = 'filter-grid';
      filterGrid.style.display = 'grid';
      filterGrid.style.gridTemplateColumns = 'repeat(4, 1fr)';
      filterGrid.style.gap = '10px';

      // Buscador arriba
      const selectorDiv = document.createElement("div");
      selectorDiv.className = "column-selector-container";
      const input = document.createElement("input");
      input.type = "text";
      input.className = "column-selector";
      input.placeholder = "Search column...";
      input.autocomplete = "off";
      // Limpiar el texto al hacer focus y mostrar todas las columnas
      input.addEventListener('focus', () => {
        input.value = '';
        // Mostrar todas las columnas al hacer focus
        filterGrid.querySelectorAll('.filter-item').forEach(div => {
          div.style.display = '';
        });
      });
      selectorDiv.appendChild(input);
      container.appendChild(selectorDiv);
      container.appendChild(filterGrid);

      // Crear todos los filtros visibles
      console.log(`🎯 Creating ${type} filters for columns:`, cols);
      cols.forEach(selectedColumn => {
        console.log(`📝 Creating filter for column: ${selectedColumn}`);
        setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: type });
        
        // Remove any existing filter for this column to prevent duplicates
        const existingFilter = filterGrid.querySelector(`[data-column="${selectedColumn}"]`);
        if (existingFilter) {
          console.log(`⚠️ Removing existing filter for ${selectedColumn} to prevent duplicates`);
          existingFilter.remove();
        }
        const filterDiv = document.createElement("div");
        filterDiv.className = "filter-item";
        filterDiv.dataset.column = selectedColumn;
        filterDiv.style.display = 'flex';
        filterDiv.style.flexDirection = 'column';
        filterDiv.style.alignItems = 'stretch';
        const headerRow = document.createElement('div');
        headerRow.style.display = 'flex';
        headerRow.style.justifyContent = 'space-between';
        headerRow.style.alignItems = 'center';
        headerRow.style.marginBottom = '0.5rem';
        
        const label = document.createElement('label');
        label.textContent = selectedColumn;
        
        const resetBtn = document.createElement('button');
        resetBtn.className = 'filter-reset-btn';
        resetBtn.textContent = '✕';
        resetBtn.title = 'Clear this filter';
        resetBtn.addEventListener('click', () => {
          // Limpiar valores del filtro
          const updated = { ...getModuleFilterValues() };
          delete updated[`${selectedColumn}_start`];
          delete updated[`${selectedColumn}_end`];
          delete updated[`${selectedColumn}_empty`];
          delete updated[selectedColumn];
          setModuleFilterValues(updated);
          
          // Eliminar de activeFilters
          const activeFilters = { ...getModuleActiveFilters() };
          delete activeFilters[selectedColumn];
          setModuleActiveFilters(activeFilters);
          
          // Limpiar inputs visuales y atributos
          filterDiv.querySelectorAll('input[type="date"]').forEach(input => {
            input.value = '';
            input.removeAttribute('data-dynamic');
          });
          filterDiv.querySelectorAll('.filter-checkbox').forEach(checkbox => checkbox.checked = false);
          filterDiv.classList.remove('active');
          
          // Refrescar chips, resumen y tabla
          renderActiveFiltersList && renderActiveFiltersList();
          renderActiveFiltersSummaryChips && renderActiveFiltersSummaryChips();
          updateActiveFiltersSummary && updateActiveFiltersSummary();
          applyFilters && applyFilters();
        });
        
        headerRow.appendChild(label);
        headerRow.appendChild(resetBtn);
        filterDiv.appendChild(headerRow);
        // Input según tipo
        if (type === 'date') {
          // Copio la lógica de creación de filtros de fecha
          const inputWrapper = document.createElement("div");
          inputWrapper.className = "filter-input-wrapper";
          // Start date
          const startInput = document.createElement("input");
          startInput.type = "date";
          startInput.className = "filter-input";
          startInput.placeholder = "Start date";
          startInput.dataset.column = selectedColumn;
          startInput.dataset.type = "start";
          const startKey = `${selectedColumn}_start`;
          const startVal = getModuleFilterValues()[startKey];
          if (typeof startVal === 'string' && startVal.startsWith('__TODAY__')) {
            startInput.value = resolveDynamicDateExpr(startVal);
            startInput.dataset.dynamic = startVal;
            filterDiv.classList.add('active');
            setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: 'date' });
          } else if (startVal) {
            startInput.value = startVal;
            filterDiv.classList.add('active');
            setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: 'date' });
          }
          // Botón Today para start
          const todayBtnStart = document.createElement('button');
          todayBtnStart.type = 'button';
          todayBtnStart.className = 'date-today-btn';
          todayBtnStart.textContent = 'Today';
          todayBtnStart.style.marginLeft = '0.5rem';
          todayBtnStart.addEventListener('click', () => {
            // Force clear any existing dynamic value
            startInput.removeAttribute('data-dynamic');
            startInput.value = '';
            
            // Set new dynamic value
            const today = new Date().toISOString().split('T')[0];
            startInput.value = today;
            startInput.dataset.dynamic = '__TODAY__';
            
            // Force update store with new value
            const key = `${selectedColumn}_start`;
            setModuleFilterValues({ ...getModuleFilterValues(), [key]: '__TODAY__' });
            setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: 'date' });
            filterDiv.classList.add("active");
            
            // Update UI and apply filters
            updateActiveFiltersSummary();
            applyFilters();
            renderActiveFiltersSummaryChips();
          });
          // Input para días personalizados (para startInput)
          const daysInputStart = document.createElement('input');
          daysInputStart.type = 'number';
          daysInputStart.value = 7;
          daysInputStart.min = 1;
          daysInputStart.max = 365;
          // daysInputStart.style.width = '48px'; // Eliminado para que el CSS lo controle
          daysInputStart.className = 'days-offset-input';
          daysInputStart.title = 'Days to add/subtract';
          // Botón + días para startInput
          const plusBtnStart = document.createElement('button');
          plusBtnStart.type = 'button';
          plusBtnStart.textContent = '+ days';
          plusBtnStart.className = 'date-offset-btn';
          plusBtnStart.addEventListener('click', () => {
            let baseExpr = startInput.dataset.dynamic || '__TODAY__';
            let offset = 0;
            const match = baseExpr.match(/__TODAY__(?:([+-])(\d+))?/);
            if (match && match[1] && match[2]) {
              offset = parseInt(match[2], 10) * (match[1] === '-' ? -1 : 1);
            }
            offset += parseInt(daysInputStart.value, 10);
            const expr = `__TODAY__${offset >= 0 ? '+' : ''}${offset}`;
            startInput.value = resolveDynamicDateExpr(expr);
            startInput.dataset.dynamic = expr;
            const key = `${selectedColumn}_start`;
            setModuleFilterValues({ ...getModuleFilterValues(), [key]: expr });
            filterDiv.classList.add("active");
            updateActiveFiltersSummary();
            applyFilters();
            renderActiveFiltersSummaryChips();
          });
          // Botón - días para startInput
          const minusBtnStart = document.createElement('button');
          minusBtnStart.type = 'button';
          minusBtnStart.textContent = '- days';
          minusBtnStart.className = 'date-offset-btn';
          minusBtnStart.addEventListener('click', () => {
            let baseExpr = startInput.dataset.dynamic || '__TODAY__';
            let offset = 0;
            const match = baseExpr.match(/__TODAY__(?:([+-])(\d+))?/);
            if (match && match[1] && match[2]) {
              offset = parseInt(match[2], 10) * (match[1] === '-' ? -1 : 1);
            }
            offset -= parseInt(daysInputStart.value, 10);
            const expr = `__TODAY__${offset >= 0 ? '+' : ''}${offset}`;
            startInput.value = resolveDynamicDateExpr(expr);
            startInput.dataset.dynamic = expr;
            const key = `${selectedColumn}_start`;
            setModuleFilterValues({ ...getModuleFilterValues(), [key]: expr });
            filterDiv.classList.add("active");
            updateActiveFiltersSummary();
            applyFilters();
            renderActiveFiltersSummaryChips();
          });
          // End date
          const endInput = document.createElement("input");
          endInput.type = "date";
          endInput.className = "filter-input";
          endInput.placeholder = "End date";
          endInput.dataset.column = selectedColumn;
          endInput.dataset.type = "end";
          const endKey = `${selectedColumn}_end`;
          const endVal = getModuleFilterValues()[endKey];
          if (typeof endVal === 'string' && endVal.startsWith('__TODAY__')) {
            endInput.value = resolveDynamicDateExpr(endVal);
            endInput.dataset.dynamic = endVal;
            filterDiv.classList.add('active');
            setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: 'date' });
          } else if (endVal) {
            endInput.value = endVal;
            filterDiv.classList.add('active');
            setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: 'date' });
          }
          // Botón Today para end
          const todayBtnEnd = document.createElement('button');
          todayBtnEnd.type = 'button';
          todayBtnEnd.className = 'date-today-btn';
          todayBtnEnd.textContent = 'Today';
          todayBtnEnd.style.marginLeft = '0.5rem';
          todayBtnEnd.addEventListener('click', () => {
            // Force clear any existing dynamic value
            endInput.removeAttribute('data-dynamic');
            endInput.value = '';
            
            // Set new dynamic value
            const today = new Date().toISOString().split('T')[0];
            endInput.value = today;
            endInput.dataset.dynamic = '__TODAY__';
            
            // Force update store with new value
            const key = `${selectedColumn}_end`;
            setModuleFilterValues({ ...getModuleFilterValues(), [key]: '__TODAY__' });
            setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: 'date' });
            filterDiv.classList.add("active");
            
            // Update UI and apply filters
            updateActiveFiltersSummary();
            applyFilters();
            renderActiveFiltersSummaryChips();
          });
          // Input para días personalizados (para endInput)
          const daysInputEnd = document.createElement('input');
          daysInputEnd.type = 'number';
          daysInputEnd.value = 7;
          daysInputEnd.min = 1;
          daysInputEnd.max = 365;
          // daysInputEnd.style.width = '48px'; // Eliminado para que el CSS lo controle
          daysInputEnd.className = 'days-offset-input';
          daysInputEnd.title = 'Days to add/subtract';
          // Botón + días para endInput
          const plusBtnEnd = document.createElement('button');
          plusBtnEnd.type = 'button';
          plusBtnEnd.textContent = '+ days';
          plusBtnEnd.className = 'date-offset-btn';
          plusBtnEnd.addEventListener('click', () => {
            let baseExpr = endInput.dataset.dynamic || '__TODAY__';
            let offset = 0;
            const match = baseExpr.match(/__TODAY__(?:([+-])(\d+))?/);
            if (match && match[1] && match[2]) {
              offset = parseInt(match[2], 10) * (match[1] === '-' ? -1 : 1);
            }
            offset += parseInt(daysInputEnd.value, 10);
            const expr = `__TODAY__${offset >= 0 ? '+' : ''}${offset}`;
            endInput.value = resolveDynamicDateExpr(expr);
            endInput.dataset.dynamic = expr;
            const key = `${selectedColumn}_end`;
            setModuleFilterValues({ ...getModuleFilterValues(), [key]: expr });
            filterDiv.classList.add("active");
            updateActiveFiltersSummary();
            applyFilters();
            renderActiveFiltersSummaryChips();
          });
          // Botón - días para endInput
          const minusBtnEnd = document.createElement('button');
          minusBtnEnd.type = 'button';
          minusBtnEnd.textContent = '- days';
          minusBtnEnd.className = 'date-offset-btn';
          minusBtnEnd.addEventListener('click', () => {
            let baseExpr = endInput.dataset.dynamic || '__TODAY__';
            let offset = 0;
            const match = baseExpr.match(/__TODAY__(?:([+-])(\d+))?/);
            if (match && match[1] && match[2]) {
              offset = parseInt(match[2], 10) * (match[1] === '-' ? -1 : 1);
            }
            offset -= parseInt(daysInputEnd.value, 10);
            const expr = `__TODAY__${offset >= 0 ? '+' : ''}${offset}`;
            endInput.value = resolveDynamicDateExpr(expr);
            endInput.dataset.dynamic = expr;
            const key = `${selectedColumn}_end`;
            setModuleFilterValues({ ...getModuleFilterValues(), [key]: expr });
            filterDiv.classList.add("active");
            updateActiveFiltersSummary();
            applyFilters();
            renderActiveFiltersSummaryChips();
          });
          // Empty button
          const emptyBtn = document.createElement('button');
          emptyBtn.type = 'button';
          emptyBtn.className = 'empty-toggle-btn';
          emptyBtn.textContent = 'Empty';
          if (getModuleFilterValues()[`${selectedColumn}_empty`]) emptyBtn.classList.add('active');
          emptyBtn.addEventListener('click', () => {
            const key = `${selectedColumn}_empty`;
            if (getModuleFilterValues()[key]) {
              const updated = { ...getModuleFilterValues() };
              delete updated[key];
              setModuleFilterValues(updated);
              emptyBtn.classList.remove('active');
              // Si no hay ningún filtro de fecha activo, eliminar de activeFilters
              if (!getModuleFilterValues()[`${selectedColumn}_start`] && 
                  !getModuleFilterValues()[`${selectedColumn}_end`]) {
                const activeFilters = { ...getModuleActiveFilters() };
                delete activeFilters[selectedColumn];
                setModuleActiveFilters(activeFilters);
                filterDiv.classList.remove('active');
              }
            } else {
              setModuleFilterValues({ ...getModuleFilterValues(), [key]: true });
              // Asegurar que el tipo 'date' está en activeFilters
              setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: 'date' });
              emptyBtn.classList.add('active');
              filterDiv.classList.add('active');
            }
            updateActiveFiltersSummary();
            applyFilters();
            renderActiveFiltersSummaryChips();
          });
          // Listeners para los inputs de fecha
          [startInput, endInput].forEach(inputEl => {
            inputEl.addEventListener("change", debounce(() => {
              const key = `${selectedColumn}_${inputEl.dataset.type}`;
              if (inputEl.value) {
                const todayStr = new Date().toISOString().slice(0, 10);
                if (inputEl.value === todayStr) {
                  setModuleFilterValues({ ...getModuleFilterValues(), [key]: '__TODAY__' });
                } else {
                  setModuleFilterValues({ ...getModuleFilterValues(), [key]: inputEl.value });
                }
                // Asegurar que el tipo 'date' está en activeFilters
                setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: 'date' });
                filterDiv.classList.add("active");
              } else {
                const updated = { ...getModuleFilterValues() };
                delete updated[key];
                setModuleFilterValues(updated);
                // Si no hay ningún filtro de fecha activo, eliminar de activeFilters
                if (!getModuleFilterValues()[`${selectedColumn}_start`] && 
                    !getModuleFilterValues()[`${selectedColumn}_end`] && 
                    !getModuleFilterValues()[`${selectedColumn}_empty`]) {
                  const activeFilters = { ...getModuleActiveFilters() };
                  delete activeFilters[selectedColumn];
                  setModuleActiveFilters(activeFilters);
                  filterDiv.classList.remove("active");
                }
              }
              updateActiveFiltersSummary();
              applyFilters();
              renderActiveFiltersSummaryChips();
            }, 300));
          });
          // Agrupar inputs y controles
          const startGroup = document.createElement('div');
          startGroup.className = 'date-input-group';
          const startLabel = document.createElement('label');
          startLabel.textContent = 'From';
          startGroup.appendChild(startLabel);
          startGroup.appendChild(startInput);
          const startControls = document.createElement('div');
          startControls.className = 'date-input-controls';
          startControls.appendChild(todayBtnStart);
          startControls.appendChild(daysInputStart);
          startControls.appendChild(plusBtnStart);
          startControls.appendChild(minusBtnStart);
          startGroup.appendChild(startControls);
          const endGroup = document.createElement('div');
          endGroup.className = 'date-input-group';
          const endLabel = document.createElement('label');
          endLabel.textContent = 'To';
          endGroup.appendChild(endLabel);
          endGroup.appendChild(endInput);
          const endControls = document.createElement('div');
          endControls.className = 'date-input-controls';
          endControls.appendChild(todayBtnEnd);
          endControls.appendChild(daysInputEnd);
          endControls.appendChild(plusBtnEnd);
          endControls.appendChild(minusBtnEnd);
          endGroup.appendChild(endControls);
          inputWrapper.innerHTML = '';
          inputWrapper.appendChild(startGroup);
          inputWrapper.appendChild(endGroup);
          inputWrapper.appendChild(emptyBtn);
          filterDiv.appendChild(inputWrapper);
        } else {
          // Filtro tipo Excel para todos los campos que no sean fecha
          let uniqueValues = [...new Set(getOriginalData()
            .map(row => row[selectedColumn] || '')
            .map(String)
            .map(val => val.trim()) // Eliminar espacios en blanco
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
          
          // Ordenar alfabéticamente para mejor UX
          uniqueValues.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
          
          if (type === 'date') {
            uniqueValues = uniqueValues.map(val => toISODateString(val));
          }
          
          // Crear selector de condición
          const conditionWrapper = document.createElement('div');
          conditionWrapper.className = 'filter-condition-wrapper';
          conditionWrapper.style.marginBottom = '0.5rem';
          
          const conditionSelect = document.createElement('select');
          conditionSelect.className = 'filter-condition-select';
          conditionSelect.style.width = '100%';
          conditionSelect.style.padding = '0.4rem';
          conditionSelect.style.borderRadius = '4px';
          conditionSelect.style.border = '1px solid var(--border-color)';
          conditionSelect.style.backgroundColor = '#1a2332';
          conditionSelect.style.color = '#ffffff';
          conditionSelect.style.fontSize = '0.9rem';
          
          // Opciones de condición
          const conditions = [
            { value: 'contains', label: 'Contains' },
            { value: 'equals', label: 'Equals' },
            { value: 'not_contains', label: 'Not Contains' },
            { value: 'not_equals', label: 'Not Equals' }
          ];
          
          conditions.forEach(cond => {
            const option = document.createElement('option');
            option.value = cond.value;
            option.textContent = cond.label;
            conditionSelect.appendChild(option);
          });
          
          // Cargar condición guardada si existe
          const savedCondition = getModuleFilterValues()[`${selectedColumn}_condition`] || 'contains';
          conditionSelect.value = savedCondition;
          
          // Guardar condición cuando cambie
          conditionSelect.addEventListener('change', () => {
            const key = `${selectedColumn}_condition`;
            const newCondition = conditionSelect.value;
            const currentValues = { ...getModuleFilterValues() };
            currentValues[key] = newCondition;
            setModuleFilterValues(currentValues);
            // Reaplicar filtros con la nueva condición
            applyFilters();
            updateActiveFiltersSummary();
            renderActiveFiltersSummaryChips();
          });
          
          conditionWrapper.appendChild(conditionSelect);
          filterDiv.appendChild(conditionWrapper);
          
          // Input de texto libre para filtrar por valor escrito
          const textInputWrapper = document.createElement('div');
          textInputWrapper.className = 'filter-text-input-wrapper';
          textInputWrapper.style.marginBottom = '0.5rem';
          
          const textInput = document.createElement('input');
          textInput.type = 'text';
          textInput.className = 'filter-text-input';
          textInput.placeholder = `Type a value to filter (comma-separated for multiple)...`;
          textInput.autocomplete = 'off';
          textInput.style.width = '100%';
          textInput.style.padding = '0.4rem';
          textInput.style.borderRadius = '4px';
          textInput.style.border = '1px solid var(--border-color)';
          textInput.style.backgroundColor = '#1a2332';
          textInput.style.color = '#ffffff';
          textInput.style.fontSize = '0.9rem';
          
          // Cargar valor guardado si existe
          const savedTextValue = getModuleFilterValues()[selectedColumn];
          if (savedTextValue && typeof savedTextValue === 'string' && !Array.isArray(savedTextValue)) {
            textInput.value = savedTextValue;
            filterDiv.classList.add('active');
          }
          
          // Aplicar filtro cuando se escribe
          const handleTextInput = () => {
            const textValue = textInput.value.trim();
            const condition = getModuleFilterValues()[`${selectedColumn}_condition`] || 'contains';
            
            if (textValue) {
              // Si hay un valor escrito, usar ese valor como filtro
              // Soporte para múltiples valores separados por comas
              const values = textValue.split(',').map(v => v.trim()).filter(v => v !== '');
              
              if (values.length === 0) {
                // Si no hay valores válidos después de filtrar, limpiar
                const updated = { ...getModuleFilterValues() };
                delete updated[selectedColumn];
                const activeFilters = { ...getModuleActiveFilters() };
                delete activeFilters[selectedColumn];
                setModuleActiveFilters(activeFilters);
                filterDiv.classList.remove('active');
                setModuleFilterValues(updated);
                updateActiveFiltersSummary();
                applyFilters();
                renderActiveFiltersSummaryChips();
                return;
              }
              
              // Limpiar checkboxes seleccionados cuando se usa el input de texto
              selectedSet.clear();
              // Actualizar checkboxes visualmente
              if (dropdown) {
                dropdown.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                  cb.checked = false;
                });
              }
              
              const currentValues = { ...getModuleFilterValues() };
              // Si hay múltiples valores, guardarlos como array, si no, como string
              if (values.length > 1) {
                currentValues[selectedColumn] = values;
              } else {
                currentValues[selectedColumn] = values[0];
              }
              currentValues[`${selectedColumn}_condition`] = condition;
              setModuleFilterValues(currentValues);
              setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: type });
              filterDiv.classList.add('active');
              
              // Aplicar filtros inmediatamente
              updateActiveFiltersSummary();
              applyFilters();
              renderActiveFiltersSummaryChips();
              // Actualizar el resumen del input del dropdown
              if (typeof updateInputSummary === 'function') {
                updateInputSummary();
              }
            } else {
              // Si está vacío, limpiar el filtro
              const updated = { ...getModuleFilterValues() };
              // Solo eliminar si no hay valores de checkbox seleccionados
              if (!Array.isArray(updated[selectedColumn]) || updated[selectedColumn].length === 0) {
                delete updated[selectedColumn];
                const activeFilters = { ...getModuleActiveFilters() };
                delete activeFilters[selectedColumn];
                setModuleActiveFilters(activeFilters);
                filterDiv.classList.remove('active');
              }
              setModuleFilterValues(updated);
              updateActiveFiltersSummary();
              applyFilters();
              renderActiveFiltersSummaryChips();
            }
          };
          
          // Debounced version para mientras escribe
          const handleTextInputDebounced = debounce(handleTextInput, 300);
          
          textInput.addEventListener('input', handleTextInputDebounced);
          textInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              handleTextInput();
            }
          });
          
          // Asegurar que el filtro se aplique cuando se presiona el botón Apply
          // Guardar referencia al input para que esté disponible globalmente
          filterDiv.textInput = textInput;
          filterDiv.handleTextInput = handleTextInput;
          
          // También aplicar cuando se pierde el foco
          textInput.addEventListener('blur', () => {
            if (textInput.value.trim()) {
              handleTextInput();
            }
          });
          
          textInputWrapper.appendChild(textInput);
          filterDiv.appendChild(textInputWrapper);
          
          const dropdownWrapper = document.createElement('div');
          dropdownWrapper.className = 'modal-filter-dropdown-wrapper';
          dropdownWrapper.style.position = 'relative';
          const input = document.createElement('input');
          input.type = 'text';
          input.className = 'modal-filter-input';
          input.placeholder = `Search or select...`;
          input.autocomplete = 'off';
          input.addEventListener('focus', () => {
            if (input.value) input.value = '';
          });
          dropdownWrapper.appendChild(input);
          const dropdown = document.createElement('div');
          dropdown.className = 'modal-filter-dropdown hidden';
          dropdown.style.position = 'absolute';
          dropdown.style.top = '100%';
          dropdown.style.left = '0';
          dropdown.style.width = '100%';
          dropdown.style.background = '#fff';
          dropdown.style.border = '1px solid var(--border-color)';
          dropdown.style.borderRadius = '0 0 6px 6px';
          dropdown.style.boxShadow = '0 2px 8px rgba(0,0,0,0.08)';
          dropdown.style.zIndex = '100';
          dropdown.style.maxHeight = '260px';
          dropdown.style.overflowY = 'auto';
          dropdown.style.fontSize = '0.95rem';
          dropdown.style.padding = '0';
          dropdown.style.setProperty('color', '#ffffff', 'important');
          dropdown.style.setProperty('background', '#1a2332', 'important');
          dropdown.style.setProperty('border', '2px solid #47B2E5', 'important');
          dropdown.style.setProperty('box-shadow', '0 12px 40px rgba(0, 0, 0, 0.8)', 'important');
          dropdown.setAttribute('style', dropdown.getAttribute('style') + '; color: #ffffff !important; background: #1a2332 !important; border: 2px solid #47B2E5 !important; box-shadow: 0 12px 40px rgba(0, 0, 0, 0.8) !important;');
          dropdownWrapper.appendChild(dropdown);
          // Añadir event listener global para cerrar el dropdown si se hace clic fuera
          document.addEventListener('mousedown', (e) => {
            if (!dropdownWrapper.contains(e.target)) {
              dropdown.classList.add('hidden');
            }
          });
          
          // MUTATION OBSERVER PARA FORZAR ESTILOS EN ELEMENTOS DINÁMICOS
          const observer = new MutationObserver((mutations) => {
            mutations.forEach((mutation) => {
              if (mutation.type === 'childList') {
                mutation.addedNodes.forEach((node) => {
                  if (node.nodeType === Node.ELEMENT_NODE) {
                    // Forzar estilos en labels
                    const labels = node.querySelectorAll ? node.querySelectorAll('label') : [];
                    labels.forEach(label => {
                      label.style.setProperty('color', '#ffffff', 'important');
                      label.style.setProperty('background', 'transparent', 'important');
                      label.style.setProperty('text-shadow', 'none', 'important');
                    });
                    
                    // Forzar estilos en spans
                    const spans = node.querySelectorAll ? node.querySelectorAll('span') : [];
                    spans.forEach(span => {
                      span.style.setProperty('color', '#ffffff', 'important');
                      span.style.setProperty('background', 'transparent', 'important');
                      span.style.setProperty('text-shadow', 'none', 'important');
                    });
                    
                    // Si el nodo mismo es un label o span
                    if (node.tagName === 'LABEL') {
                      node.style.setProperty('color', '#ffffff', 'important');
                      node.style.setProperty('background', 'transparent', 'important');
                      node.style.setProperty('text-shadow', 'none', 'important');
                    }
                    if (node.tagName === 'SPAN') {
                      node.style.setProperty('color', '#ffffff', 'important');
                      node.style.setProperty('background', 'transparent', 'important');
                      node.style.setProperty('text-shadow', 'none', 'important');
                    }
                  }
                });
              }
            });
          });
          
          observer.observe(dropdown, {
            childList: true,
            subtree: true
          });
          
          const selectAllBtn = document.createElement('button');
          selectAllBtn.textContent = 'Select all';
          selectAllBtn.type = 'button';
          selectAllBtn.className = 'modal-filter-dropdown-btn';
          selectAllBtn.style.margin = '0.3rem 0.5rem 0.3rem 0';
          selectAllBtn.style.color = '#47B2E5';
          selectAllBtn.style.background = 'rgba(71, 178, 229, 0.2)';
          selectAllBtn.style.border = '1px solid rgba(71, 178, 229, 0.3)';
          const clearAllBtn = document.createElement('button');
          clearAllBtn.textContent = 'Clear selection';
          clearAllBtn.type = 'button';
          clearAllBtn.className = 'modal-filter-dropdown-btn';
          clearAllBtn.style.margin = '0.3rem 0 0.3rem 0.5rem';
          clearAllBtn.style.color = '#47B2E5';
          clearAllBtn.style.background = 'rgba(71, 178, 229, 0.2)';
          clearAllBtn.style.border = '1px solid rgba(71, 178, 229, 0.3)';
          const btnsDiv = document.createElement('div');
          btnsDiv.style.display = 'flex';
          btnsDiv.style.justifyContent = 'space-between';
          btnsDiv.appendChild(selectAllBtn);
          btnsDiv.appendChild(clearAllBtn);
          dropdown.appendChild(btnsDiv);
          function debounce(func, wait) {
            let timeout;
            return function(...args) {
              clearTimeout(timeout);
              timeout = setTimeout(() => func.apply(this, args), wait);
            };
          }
          let selectedSet = new Set();
          let lastFilterTerm = '';
          let filteredValues = uniqueValues;
          function renderCheckboxList(filterTerm = '') {
            selectedSet = new Set(getModuleFilterValues()[selectedColumn] || []);
            
            // LIMPIEZA COMPLETA Y ROBUSTA
            dropdown.innerHTML = '';
            
            // Recrear botones de control
            const btnsDiv = document.createElement('div');
            btnsDiv.style.display = 'flex';
            btnsDiv.style.justifyContent = 'space-between';
            btnsDiv.appendChild(selectAllBtn.cloneNode(true));
            btnsDiv.appendChild(clearAllBtn.cloneNode(true));
            dropdown.appendChild(btnsDiv);
            
            // Recrear event listeners para los botones clonados
            const newSelectAllBtn = btnsDiv.children[0];
            const newClearAllBtn = btnsDiv.children[1];
            
            newSelectAllBtn.addEventListener('click', () => {
              filteredValues.forEach(val => {
                if (val !== '') selectedSet.add(val);
              });
              // Preservar la condición al actualizar los valores
              const currentValues = { ...getModuleFilterValues() };
              currentValues[selectedColumn] = Array.from(selectedSet);
              // Asegurar que la condición se preserve
              const conditionKey = `${selectedColumn}_condition`;
              if (!currentValues[conditionKey]) {
                currentValues[conditionKey] = 'contains';
              }
              setModuleFilterValues(currentValues);
              filterDiv.classList.add('active');
              renderCheckboxList();
              updateActiveFiltersSummary();
              applyFilters();
              updateInputSummary();
            });
            
            newClearAllBtn.addEventListener('click', () => {
              selectedSet.clear();
              const updated = { ...getModuleFilterValues() };
              delete updated[selectedColumn];
              setModuleFilterValues(updated);
              filterDiv.classList.remove('active');
              renderCheckboxList();
              updateActiveFiltersSummary();
              applyFilters();
              updateInputSummary();
            });
            
            const list = document.createElement('div');
            list.className = 'modal-filter-checkbox-list';
            list.style.setProperty('color', '#ffffff', 'important');
            list.style.setProperty('background', 'transparent', 'important');
            list.setAttribute('style', 'color: #ffffff !important; background: transparent !important; padding: 0.5rem !important;');
            const emptyBtn2 = document.createElement('button');
            emptyBtn2.type = 'button';
            emptyBtn2.className = 'empty-toggle-btn';
            emptyBtn2.textContent = '(Empty)';
            emptyBtn2.style.color = '#47B2E5';
            emptyBtn2.style.background = 'rgba(71, 178, 229, 0.2)';
            emptyBtn2.style.border = '1px solid rgba(71, 178, 229, 0.3)';
            if (selectedSet.has('__EMPTY__')) emptyBtn2.classList.add('active');
            emptyBtn2.addEventListener('click', () => {
              if (selectedSet.has('__EMPTY__')) {
                selectedSet.delete('__EMPTY__');
                emptyBtn2.classList.remove('active');
              } else {
                selectedSet.add('__EMPTY__');
                emptyBtn2.classList.add('active');
              }
              // Preservar la condición al actualizar los valores
              const currentValues = { ...getModuleFilterValues() };
              currentValues[selectedColumn] = Array.from(selectedSet);
              // Asegurar que la condición se preserve
              const conditionKey = `${selectedColumn}_condition`;
              if (!currentValues[conditionKey]) {
                currentValues[conditionKey] = 'contains';
              }
              setModuleFilterValues(currentValues);
              setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: type });
              filterDiv.classList.toggle('active', selectedSet.size > 0);
              updateActiveFiltersSummary();
              applyFilters();
              updateInputSummary();
            });
            list.appendChild(emptyBtn2);
            const MAX_OPTIONS = 200;
            filteredValues = uniqueValues;
            if (filterTerm) {
              const lowerTerm = filterTerm.toLowerCase();
              filteredValues = uniqueValues.filter(val => val.toLowerCase().includes(lowerTerm));
            }
            // VERIFICAR QUE NO HAYA DUPLICADOS EN TIEMPO REAL
            const processedValues = new Set();
            filteredValues.slice(0, MAX_OPTIONS).forEach(val => {
              if (val === '' || processedValues.has(val)) return;
              processedValues.add(val);
              
              // CREAR ELEMENTO USANDO createElement Y FORZAR TEXTO
              const label = document.createElement('label');
              label.setAttribute('style', 'display: flex !important; align-items: center !important; gap: 0.5rem !important; padding: 0.15rem 0.5rem !important; color: #ffffff !important; text-shadow: none !important; background: transparent !important; cursor: pointer !important; opacity: 1 !important; visibility: visible !important; border-bottom: 1px solid rgba(255, 255, 255, 0.1) !important;');
              label.setAttribute('data-text', val); // Para el pseudo-elemento
              
              const checkbox = document.createElement('input');
              checkbox.type = 'checkbox';
              checkbox.value = val;
              checkbox.checked = selectedSet.has(val);
              checkbox.setAttribute('style', 'accent-color: #6b7280 !important;');
              
              const span = document.createElement('span');
              span.textContent = val; // USAR textContent DIRECTAMENTE
              span.setAttribute('style', 'color: #ffffff !important; text-shadow: none !important; background: transparent !important; font-size: 0.9rem !important; opacity: 1 !important; visibility: visible !important; display: inline !important;');
              
              label.appendChild(checkbox);
              label.appendChild(span);
              checkbox.addEventListener('change', () => {
                if (checkbox.checked) {
                  selectedSet.add(val);
                } else {
                  selectedSet.delete(val);
                }
                // Limpiar el input de texto cuando se seleccionan checkboxes
                if (textInput) {
                  textInput.value = '';
                }
                // Preservar la condición al actualizar los valores
                const currentValues = { ...getModuleFilterValues() };
                currentValues[selectedColumn] = Array.from(selectedSet);
                // Asegurar que la condición se preserve
                const conditionKey = `${selectedColumn}_condition`;
                if (!currentValues[conditionKey]) {
                  currentValues[conditionKey] = 'contains';
                }
                setModuleFilterValues(currentValues);
                setModuleActiveFilters({ ...getModuleActiveFilters(), [selectedColumn]: type });
                filterDiv.classList.toggle('active', selectedSet.size > 0);
                updateActiveFiltersSummary();
                applyFilters();
                updateInputSummary();
              });
              

              
              list.appendChild(label);
            });
            dropdown.appendChild(list);
            
            // Los estilos se aplican automáticamente via CSS
          }
          
          // FORZAR ESTILOS SOLO CUANDO SEA NECESARIO (eliminado setInterval problemático)
          
          const handleInput = debounce(() => {
            const term = input.value.trim().toLowerCase();
            lastFilterTerm = term;
            renderCheckboxList(term);
            dropdown.classList.remove('hidden');
          }, 200);
          input.addEventListener('input', handleInput);
          input.addEventListener('focus', () => {
            renderCheckboxList(lastFilterTerm);
            dropdown.classList.remove('hidden');
            
            // Los estilos del modal se manejan via CSS
          });
          // Evita que el dropdown se cierre al marcar checkboxes
          // input.addEventListener('blur', () => {
          //   setTimeout(() => dropdown.classList.add('hidden'), 200);
          // });
          // Nueva lógica: solo cerrar si el foco realmente sale del dropdown
          input.addEventListener('blur', (e) => {
            setTimeout(() => {
              // Si el foco está dentro del dropdown, no cerrar
              if (!dropdown.contains(document.activeElement)) {
                dropdown.classList.add('hidden');
              }
            }, 200);
          });
          selectAllBtn.addEventListener('click', () => {
            filteredValues.forEach(val => {
              if (val !== '') selectedSet.add(val);
            });
            setModuleFilterValues({ ...getModuleFilterValues(), [selectedColumn]: Array.from(selectedSet) });
            filterDiv.classList.add('active');
            renderCheckboxList();
            updateActiveFiltersSummary();
            applyFilters();
            updateInputSummary();
          });
          clearAllBtn.addEventListener('click', () => {
            selectedSet.clear();
            const updated = { ...getModuleFilterValues() };
            delete updated[selectedColumn];
            setModuleFilterValues(updated);
            filterDiv.classList.remove('active');
            renderCheckboxList();
            updateActiveFiltersSummary();
            applyFilters();
            updateInputSummary();
          });
          function updateInputSummary() {
            const selected = Array.from(selectedSet).filter(v => v !== '__EMPTY__');
            if (selectedSet.has('__EMPTY__')) {
              input.value = '(Empty)';
            } else if (selected.length === 0) {
              input.value = '';
            } else if (selected.length <= 2) {
              input.value = selected.join(', ');
            } else {
              input.value = `${selected.length} selected`;
            }
          }
          if (Array.isArray(getModuleFilterValues()[selectedColumn])) {
            selectedSet = new Set(getModuleFilterValues()[selectedColumn]);
            updateInputSummary();
          }
          filterDiv.appendChild(dropdownWrapper);
        }
        filterGrid.appendChild(filterDiv);
      });

      // Filtrado visual de los bloques según el buscador
      input.addEventListener('input', () => {
        const term = input.value.trim().toLowerCase();
        filterGrid.querySelectorAll('.filter-item').forEach(div => {
          const col = div.dataset.column.toLowerCase();
          div.style.display = col.includes(term) ? '' : 'none';
        });
      });
    });

    // Setup tabs after generating content
    function setupUnifiedTabs() {
      const tabs = document.querySelectorAll('.filter-tab');
      const panels = document.querySelectorAll('.filter-panel');
      tabs.forEach(tab => {
        tab.addEventListener('click', () => {
          tabs.forEach(t => t.classList.remove('active'));
          panels.forEach(p => p.classList.remove('active'));
          tab.classList.add('active');
          const targetPanel = document.getElementById(tab.dataset.target + 'FilterPanel');
          if (targetPanel) {
            targetPanel.classList.add('active');
            // Renderiza solo el contenido del panel activo
            if (targetPanel.id === 'myfiltersFilterPanel') {
              renderMyFiltersSection();
            }
          }
        });
      });
      // Activa solo la primera pestaña y panel por defecto
      const firstTab = tabs[0];
      const firstPanel = panels[0];
      if (firstTab && firstPanel) {
        tabs.forEach(t => t.classList.remove('active'));
        panels.forEach(p => p.classList.remove('active'));
        firstTab.classList.add('active');
        firstPanel.classList.add('active');
        if (firstPanel.id === 'myfiltersFilterPanel') {
          renderMyFiltersSection();
        }
      }
    }
    setupUnifiedTabs();

    // Elimina la pestaña y panel de Reference si existen
    if (referenceContainer) referenceContainer.remove();
    const referenceTab = document.querySelector('.filter-tab[data-target="reference"]');
    if (referenceTab) referenceTab.remove();

    // --- Update summary function ---
    function updateActiveFiltersSummary() {
      const list = document.getElementById('globalActiveFiltersSummary');
      if (!list) return;
      list.innerHTML = '';
      const filterValues = getModuleFilterValues();
      if (Object.keys(filterValues).length === 0) {
        // Hide the container completely when empty
        list.style.display = 'none';
        list.style.padding = '0';
        list.style.margin = '0';
        list.style.height = '0';
        list.style.minHeight = '0';
        return;
      } else {
        // Show the container when there are filters
        list.style.display = '';
        list.style.padding = '';
        list.style.margin = '';
        list.style.height = '';
        list.style.minHeight = '';
      }
      // Group filters by type
      const dateFilters = {};
      const otherFilters = {};
      Object.entries(filterValues).forEach(([key, value]) => {
        if (key.endsWith('_start') || key.endsWith('_end') || key.endsWith('_empty')) {
          const baseKey = key.replace(/_(start|end|empty)$/, '');
          if (!dateFilters[baseKey]) {
            dateFilters[baseKey] = { start: '', end: '', empty: false };
          }
          if (key.endsWith('_start')) dateFilters[baseKey].start = value;
          if (key.endsWith('_end')) dateFilters[baseKey].end = value;
          if (key.endsWith('_empty')) dateFilters[baseKey].empty = value;
        } else if (!key.endsWith('_condition')) {
          // Excluir condiciones del objeto otherFilters
          otherFilters[key] = value;
        }
      });
      // Add date filters
      Object.entries(dateFilters).forEach(([column, filter]) => {
        if (filter.start || filter.end || filter.empty) {
          const tag = document.createElement('div');
          tag.className = 'modal-filter-tag';
          let text = `${column}: `;
          if (filter.empty && !filter.start && !filter.end) {
            text += '(empty)';
          } else {
            if (filter.start) text += `from ${prettyDynamicDate(filter.start)}`;
            if (filter.end) text += ` to ${prettyDynamicDate(filter.end)}`;
            if (filter.empty) text += ' (including empty)';
          }
          tag.innerHTML = `
            <span>${text}</span>
            <button class="modal-filter-tag-remove" data-column="${column}">×</button>
          `;
          list.appendChild(tag);
        }
      });
      // Add other filters
      Object.entries(otherFilters).forEach(([column, values]) => {
        // Mostrar tanto arrays como strings (valores únicos del input de texto)
        if (Array.isArray(values) && values.length > 0) {
          const tag = document.createElement('div');
          tag.className = 'modal-filter-tag';
          // Obtener la condición para esta columna
          const condition = filterValues[`${column}_condition`] || 'contains';
          const conditionLabel = getConditionLabel(condition);
          const conditionText = conditionLabel ? `${conditionLabel} ` : '';
          tag.innerHTML = `
            <span>${conditionText}${column}: ${values.join(', ')}</span>
            <button class="modal-filter-tag-remove" data-column="${column}">×</button>
          `;
          list.appendChild(tag);
        } else if (typeof values === 'string' && values.trim() !== '') {
          // Mostrar valores de tipo string (del input de texto)
          const tag = document.createElement('div');
          tag.className = 'modal-filter-tag';
          // Obtener la condición para esta columna
          const condition = filterValues[`${column}_condition`] || 'contains';
          const conditionLabel = getConditionLabel(condition);
          const conditionText = conditionLabel ? `${conditionLabel} ` : '';
          tag.innerHTML = `
            <span>${conditionText}${column}: ${values}</span>
            <button class="modal-filter-tag-remove" data-column="${column}">×</button>
          `;
          list.appendChild(tag);
        }
      });
      // Add remove functionality to tags
      list.querySelectorAll('.modal-filter-tag-remove').forEach(btn => {
        btn.addEventListener('click', () => {
          const column = btn.dataset.column;
            const updated = { ...getModuleFilterValues() };
          // Eliminar todos los valores posibles del filtro
          delete updated[`${column}_start`];
          delete updated[`${column}_end`];
          delete updated[`${column}_empty`];
          delete updated[column];
          delete updated[`${column}_condition`];
          setModuleFilterValues(updated);
          // Eliminar de activeFilters SIEMPRE
          const activeFilters = { ...getModuleActiveFilters() };
          delete activeFilters[column];
          setModuleActiveFilters(activeFilters);
          // Limpiar inputs visuales
            const filterItem = document.querySelector(`.filter-item[data-column="${column}"]`);
            if (filterItem) {
              filterItem.querySelectorAll('input[type="date"]').forEach(input => input.value = '');
              filterItem.querySelectorAll('.filter-checkbox').forEach(checkbox => checkbox.checked = false);
            filterItem.classList.remove('active');
          }
          // Refuerzo: renderizar chips y lista tras limpiar el estado
          renderActiveFiltersList();
          renderActiveFiltersSummaryChips();
          updateActiveFiltersSummary();
          applyFilters();
        });
      });
    }

    // Call updateActiveFiltersSummary after each filter change
    window.updateActiveFiltersSummary = updateActiveFiltersSummary;
    updateActiveFiltersSummary();

    // --- Hash simple para cabeceras (igual que en ColumnManager) ---
    function getHeaderHash(headers) {
      let hash = 0, i, chr;
      const str = headers.join('||');
      for (i = 0; i < str.length; i++) {
        chr = str.charCodeAt(i);
        hash = ((hash << 5) - hash) + chr;
        hash |= 0;
      }
      return hash.toString();
    }

    // --- Guardar y cargar filtros ---
    function saveMyFilter(name, urgencyCard) {
  const headers = Object.keys(getOriginalData()[0] || {});
  const headerHash = getHeaderHash(headers);
  const filters = loadMyFilters();
  const filterValuesToSave = { ...getModuleFilterValues() };
  // Asegurar que todas las condiciones se guarden
  const conditionKeys = Object.keys(filterValuesToSave).filter(k => k.endsWith('_condition'));
  console.log('💾 Saving filter:', name, 'with condition keys:', conditionKeys);
  conditionKeys.forEach(ck => {
    const col = ck.replace('_condition', '');
    console.log(`💾   - ${col}: condition = "${filterValuesToSave[ck]}"`);
  });
  filters[name] = { filterValues: filterValuesToSave, headerHash, headers, linkedUrgencyCard: urgencyCard };
  localStorage.setItem('myFilters', JSON.stringify(filters));
  
  // Trigger auto-save
  if (typeof window.triggerAutoSave === 'function') {
    window.triggerAutoSave();
  }
  
  // Show notification
  if (typeof window.showUnifiedNotification === 'function') {
    window.showUnifiedNotification(`Filter "${name}" saved successfully!`, 'success');
  }
}
    function loadMyFilters() {
      try {
        const saved = localStorage.getItem('myFilters');
        return saved ? JSON.parse(saved) : {};
      } catch {
        return {};
      }
    }
    function deleteMyFilter(name) {
  const filters = loadMyFilters();
  delete filters[name];
  localStorage.setItem('myFilters', JSON.stringify(filters));
  
  // Trigger auto-save
  if (typeof window.triggerAutoSave === 'function') {
    window.triggerAutoSave();
  }
  
  // Show notification
  if (typeof window.showUnifiedNotification === 'function') {
    window.showUnifiedNotification(`Filter "${name}" deleted successfully!`, 'success');
  }
}
    function applyMyFilter(name) {
      const filters = loadMyFilters();
      const headers = Object.keys(getOriginalData()[0] || {});
      const headerHash = getHeaderHash(headers);
      const filterObj = filters[name];
      if (!filterObj) {
        if (typeof window.showUnifiedNotification === 'function') {
          window.showUnifiedNotification(`Filter "${name}" not found!`, 'error');
        }
        return;
      }
      // Check if current headers match the saved filter's headers (not just hash)
      const savedHeaders = filterObj.headers || [];
      const sameColumns = headers.length === savedHeaders.length && headers.every((col, i) => col === savedHeaders[i]);
      if (!sameColumns) {
        if (typeof window.showUnifiedNotification === 'function') {
          window.showUnifiedNotification(`This filter "${name}" cannot be applied to the current CSV.`, 'error');
        }
        return;
      }
      if (filterObj.headerHash === headerHash) {
        // Asegurar que todas las condiciones se copien correctamente
        const filterValuesToApply = { ...filterObj.filterValues };
        // IMPORTANTE: Las condiciones _condition deben estar presentes antes de aplicar filtros
        const conditionKeys = Object.keys(filterValuesToApply).filter(k => k.endsWith('_condition'));
        console.log('🔍 Applying saved filter:', name);
        console.log('🔍 Condition keys found:', conditionKeys);
        conditionKeys.forEach(ck => {
          const col = ck.replace('_condition', '');
          console.log(`🔍   - ${col}: condition = "${filterValuesToApply[ck]}", value =`, filterValuesToApply[col]);
        });
        setModuleFilterValues(filterValuesToApply);
        // Reconstruct activeFilters from filterValues (excluyendo _condition)
        const newActiveFilters = {};
        for (const key of Object.keys(filterObj.filterValues)) {
          // Excluir claves _condition al reconstruir activeFilters
          if (key.endsWith('_condition')) {
            continue;
          }
          if (key.endsWith('_start') || key.endsWith('_end') || key.endsWith('_empty')) {
            const base = key.replace(/_(start|end|empty)$/, '');
            newActiveFilters[base] = 'date';
          } else if (Array.isArray(filterObj.filterValues[key])) {
            newActiveFilters[key] = 'categorical';
          } else if (filterObj.filterValues[key] !== null && filterObj.filterValues[key] !== undefined && filterObj.filterValues[key] !== '') {
            newActiveFilters[key] = 'text';
          }
        }
        setModuleActiveFilters(newActiveFilters);
        generateFilterSidebar(headers);
        // Aplicar filtros INMEDIATAMENTE con las condiciones guardadas
        // Esto es crítico para que las condiciones NOT funcionen
        applyFilters();
        renderActiveFiltersSummaryChips();
        // Restaurar condiciones en los selectores después de regenerar el sidebar (para visualización)
        setTimeout(() => {
          Object.keys(filterObj.filterValues).forEach(key => {
            if (key.endsWith('_condition')) {
              const column = key.replace('_condition', '');
              const conditionSelect = document.querySelector(`.filter-item[data-column="${column}"] .filter-condition-select`);
              if (conditionSelect) {
                conditionSelect.value = filterObj.filterValues[key];
                console.log(`✅ Restored condition selector for "${column}": "${filterObj.filterValues[key]}"`);
                // Disparar evento change para asegurar que se guarde la condición
                conditionSelect.dispatchEvent(new Event('change'));
              } else {
                console.warn(`⚠️ Condition selector not found for column "${column}"`);
              }
              // También restaurar el input de texto si existe
              const textInput = document.querySelector(`.filter-item[data-column="${column}"] .filter-text-input`);
              if (textInput && filterObj.filterValues[column] && typeof filterObj.filterValues[column] === 'string' && !Array.isArray(filterObj.filterValues[column])) {
                textInput.value = filterObj.filterValues[column];
              }
            }
          });
          // Verificar que las condiciones se aplicaron correctamente
          const currentFilterValues = getModuleFilterValues();
          const currentConditionKeys = Object.keys(currentFilterValues).filter(k => k.endsWith('_condition'));
          console.log('✅ Current filter values condition keys after restore:', currentConditionKeys);
          // Reaplicar filtros después de restaurar las condiciones visualmente (por si acaso)
          applyFilters();
        }, 200); // Aumentar el timeout para dar más tiempo a que se genere el sidebar
        // Forzar que la pestaña y panel de My Filters estén activos
        const myFiltersTab = document.querySelector('.filter-tab[data-target="myfilters"]');
        const myFiltersPanel = document.getElementById('myfiltersFilterPanel');
        if (myFiltersTab && myFiltersPanel) {
          document.querySelectorAll('.filter-panel').forEach(p => p.classList.remove('active'));
          document.querySelectorAll('.filter-tab').forEach(t => t.classList.remove('active'));
          myFiltersTab.classList.add('active');
          myFiltersPanel.classList.add('active');
        }
        
        // Show notification
        if (typeof window.showUnifiedNotification === 'function') {
          window.showUnifiedNotification(`Filter "${name}" applied successfully!`, 'success');
        }
      } else {
        if (typeof window.showUnifiedNotification === 'function') {
          window.showUnifiedNotification(`This filter "${name}" cannot be applied to the current CSV.`, 'error');
        }
      }
    }

    // --- Renderiza la sección My Filters en el panel de filtros ---
    function renderMyFiltersSection() {
      let section = document.getElementById('myFiltersSection');
      if (!section) {
        section = document.createElement('div');
        section.id = 'myFiltersSection';
        section.className = 'my-filters-section';
        // SOLO insertar en el panel correcto
        const myFiltersPanel = document.getElementById('myfiltersFilterPanel');
        if (myFiltersPanel) myFiltersPanel.appendChild(section);
      }
      section.innerHTML = '<div class="my-filters-title">My Filters</div>';
      const filters = loadMyFilters();
      const headers = Object.keys(getOriginalData()[0] || {});
      const headerHash = getHeaderHash(headers);
      const list = document.createElement('div');
      list.className = 'my-filters-list';
      let hasAny = false;
      Object.entries(filters).forEach(([name, obj]) => {
        hasAny = true;
        const item = document.createElement('div');
        item.className = 'my-filter-item';
        const label = document.createElement('span');
        label.textContent = name;
        label.className = 'my-filter-name';
        const applyBtn = document.createElement('button');
        applyBtn.textContent = 'Apply';
        applyBtn.className = 'my-filter-apply-btn';
        const deleteBtn = document.createElement('button');
        deleteBtn.textContent = 'Delete';
        deleteBtn.className = 'my-filter-delete-btn';
        deleteBtn.addEventListener('click', () => {
          // Elimina el tooltip si existe
          const tooltip = document.getElementById('myFilterTooltip');
          if (tooltip) tooltip.remove();
          deleteMyFilter(name);
          renderMyFiltersSection();
        });
        applyBtn.addEventListener('click', () => applyMyFilter(name));
        item.appendChild(label);
        item.appendChild(deleteBtn);
        item.appendChild(applyBtn);
        list.appendChild(item);
      });
      if (!hasAny) {
        const emptyMsg = document.createElement('div');
        emptyMsg.className = 'my-filters-empty';
        emptyMsg.textContent = 'No saved filters.';
        list.appendChild(emptyMsg);
      }
      section.appendChild(list);

      // --- Quick Filters Section ---
      const quickFiltersTitle = document.createElement('div');
      quickFiltersTitle.className = 'my-filters-title';
      quickFiltersTitle.style.marginTop = '2rem';
      quickFiltersTitle.textContent = 'Quick Filters';
      section.appendChild(quickFiltersTitle);
      const quickList = document.createElement('div');
      quickList.className = 'my-filters-list';
      const quickFilters = loadQuickFilters();
      let hasQuick = false;
      Object.entries(quickFilters).forEach(([name, obj]) => {
        hasQuick = true;
        const item = document.createElement('div');
        item.className = 'my-filter-item';
        const label = document.createElement('span');
        label.textContent = name;
        label.className = 'my-filter-name';
        const applyBtn = document.createElement('button');
        applyBtn.textContent = 'Apply';
        applyBtn.className = 'my-filter-apply-btn';
        const deleteBtn = document.createElement('button');
        deleteBtn.textContent = 'Delete';
        deleteBtn.className = 'my-filter-delete-btn';
        deleteBtn.addEventListener('click', () => {
          deleteQuickFilter(name);
          renderMyFiltersSection();
        });
        applyBtn.addEventListener('click', () => {
          // Aplica el quick filter directamente
          const quickFilters = loadQuickFilters();
          const filterObj = quickFilters[name];
          if (filterObj) {
            setModuleFilterValues({ ...filterObj.filterValues });
            // Reconstruir activeFilters
            const newActiveFilters = {};
            for (const key of Object.keys(filterObj.filterValues)) {
              if (key.endsWith('_start') || key.endsWith('_end') || key.endsWith('_empty')) {
                const base = key.replace(/_(start|end|empty)$/, '');
                newActiveFilters[base] = 'date';
              } else if (Array.isArray(filterObj.filterValues[key])) {
                newActiveFilters[key] = 'categorical';
              } else {
                newActiveFilters[key] = 'text';
              }
            }
            setModuleActiveFilters(newActiveFilters);
            applyFilters();
            renderActiveFiltersSummaryChips();
            // Cierra el modal de filtros si está abierto
            const filterModal = document.getElementById('filterModal');
            const filterOverlay = document.getElementById('filterModalOverlay');
            if (filterModal && filterOverlay) {
              filterModal.classList.add('hidden');
              filterOverlay.classList.add('hidden');
              filterModal.style.display = 'none';
              filterOverlay.style.display = 'none';
              filterOverlay.classList.remove('visible');
              filterOverlay.classList.remove('blur');
              filterOverlay.style.backdropFilter = 'none';
            }
          }
        });
        item.appendChild(label);
        item.appendChild(deleteBtn);
        item.appendChild(applyBtn);
        quickList.appendChild(item);
      });
      if (!hasQuick) {
        const emptyMsg = document.createElement('div');
        emptyMsg.className = 'my-filters-empty';
        emptyMsg.textContent = 'No quick filters saved.';
        quickList.appendChild(emptyMsg);
      }
      section.appendChild(quickList);
    }

    // Cambia el nombre de la pestaña 'Generic' por 'By Ref/Status'
    const genericTab = document.querySelector('.filter-tab[data-target="generic"]');
    if (genericTab) genericTab.textContent = 'By Ref/Status';

    // --- Al final de generateFilterSidebar, fuerza el render de chips ---
    renderActiveFiltersSummaryChips();

    // OCULTAR lista de filtros de abajo si ya se muestran arriba
    try {
      const summaryChips = document.getElementById('activeFiltersSummary');
      const bottomLists = document.querySelectorAll('.active-filters-list');
      if (summaryChips && bottomLists.length > 1) {
        // El primero es el de arriba, el segundo el de abajo
        if (summaryChips.style.display !== 'none') {
          bottomLists[1].style.display = 'none';
        } else {
          bottomLists[1].style.display = '';
        }
      }
    } catch (e) { /* ignore */ }
  } catch (error) {
    console.error('Error generating filter sidebar:', error);
  } finally {
    // Verificar que los filtros se generaron correctamente
    const genericContainer = document.getElementById("genericFilterPanel");
    const dateContainer = document.getElementById("dateFilterPanel");
    const filterItems = document.querySelectorAll('.filter-item');
    
    console.log('🔍 Filter generation verification:', {
      genericContainer: !!genericContainer,
      dateContainer: !!dateContainer,
      filterItemsCount: filterItems.length,
      genericItems: genericContainer ? genericContainer.querySelectorAll('.filter-item').length : 0,
      dateItems: dateContainer ? dateContainer.querySelectorAll('.filter-item').length : 0
    });
    
    // Resetear la bandera al finalizar
    isGenerating = false;
  }
}

// Función robusta para parsear fechas en múltiples formatos
function parseFlexibleDate(value) {
  if (!value || typeof value !== "string") return null;

  // Si es un número puro de 1 a 3 dígitos, nunca es fecha
  if (/^\d{1,3}$/.test(value.trim())) return null;

  // Año completo (YYYY)
  const yearOnly = value.match(/^(\d{4})$/);
  if (yearOnly) {
    return new Date(`${yearOnly[1]}-01-01T00:00:00`);
  }
  // Año y mes (YYYY-MM)
  const yearMonth = value.match(/^(\d{4})-(\d{2})$/);
  if (yearMonth) {
    return new Date(`${yearMonth[1]}-${yearMonth[2]}-01T00:00:00`);
  }
  // dd.mm.yy (formato corto con puntos)
  const shortDateWithDots = value.match(/^(\d{2})\.(\d{2})\.(\d{2})$/);
  if (shortDateWithDots) {
    const [, day, month, year] = shortDateWithDots;
    // Asumir que años 00-29 son 2000-2029, años 30-99 son 1930-1999
    const fullYear = parseInt(year) < 30 ? `20${year}` : `19${year}`;
    return new Date(`${fullYear}-${month}-${day}T00:00:00`);
  }
  // dd/mm/yyyy hh:mm:ss o dd-mm-yyyy hh:mm:ss (hora opcional)
  const match = value.match(
    /^(\d{2})[\/-](\d{2})[\/-](\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/
  );
  if (match) {
    const [, day, month, year, hours = "00", minutes = "00", seconds = "00"] = match;
    return new Date(`${year}-${month}-${day}T${hours}:${minutes}:${seconds}`);
  }
  // yyyy-mm-dd o yyyy/mm/dd
  const iso = value.match(/^(\d{4})[\/-](\d{2})[\/-](\d{2})$/);
  if (iso) {
    const [, year, month, day] = iso;
    return new Date(`${year}-${month}-${day}T00:00:00`);
  }
  // ISO 8601 fallback
  const parsed = Date.parse(value);
  return isNaN(parsed) ? null : new Date(parsed);
}

function normalizeDate(date) {
  if (!date) return null;
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  return d;
}

function normalizeText(val) {
  return (val || "").toString().toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "").trim();
}

function applyFilters() {
    const data = getOriginalData();
    const moduleActiveFilters = getModuleActiveFilters();
    const moduleFilterValues = getModuleFilterValues();
    const tableActiveFilters = getTableActiveFilters();
    const tableFilterValues = getTableFilterValues();
    const activeFilters = getActiveFilters();
    
    if (!data || !data.length) {
        console.warn('No data available for filtering');
        return [];
    }

    let filteredData = [...data];

    // Aplicar filtros del módulo
    Object.entries(moduleActiveFilters).forEach(([column, filterType]) => {
        if (filterType === 'date') {
            const arr = Array.isArray(moduleFilterValues[column]) ? moduleFilterValues[column] : null;
            const hasRange = moduleFilterValues[`${column}_start`] || moduleFilterValues[`${column}_end`] || moduleFilterValues[`${column}_empty`];
            if (arr && arr.length > 0) {
                // Filtro Excel: solo los valores seleccionados
                filteredData = filteredData.filter(row => arr.includes(row[column]));
                return;
            } else if (hasRange) {
                // Filtro por rango
                let start = resolveDynamicValue(moduleFilterValues[`${column}_start`]);
                let end = resolveDynamicValue(moduleFilterValues[`${column}_end`]);
                if (start && /^\d{4}$/.test(start)) {
                    start = `${start}-01-01`;
                    end = `${moduleFilterValues[`${column}_end`] || start.slice(0,4)}-12-31`;
                } else if (start && /^\d{4}-\d{2}$/.test(start)) {
                    const [y, m] = start.split('-');
                    start = `${y}-${m}-01`;
                    const lastDay = new Date(y, parseInt(m,10), 0).getDate();
                    end = `${y}-${m}-${lastDay}`;
                }
                if (end && /^\d{4}$/.test(end)) {
                    end = `${end}-12-31`;
                } else if (end && /^\d{4}-\d{2}$/.test(end)) {
                    const [y, m] = end.split('-');
                    const lastDay = new Date(y, parseInt(m,10), 0).getDate();
                    end = `${y}-${m}-${lastDay}`;
                }
                const empty = moduleFilterValues[`${column}_empty`];
                filteredData = filteredData.filter(row => {
                    const value = row[column] ? row[column].trim() : "";
                    // Si solo está el filtro empty, mostrar SOLO vacíos
                    if (empty && !start && !end) {
                        return value === "" || value === null || value === undefined;
                    }
                    // Si está empty + rango, mostrar vacíos o los que cumplen el rango
                    if (empty && (value === "" || value === null || value === undefined)) return true;
                    if (!value) return false;
                    const cellDate = parseFlexibleDate(value);
                    if (!cellDate) return false;
                    if (start) {
                        const startDate = parseFlexibleDate(start);
                        if (!startDate || cellDate < startDate) return false;
                    }
                    if (end) {
                        const endDate = parseFlexibleDate(end);
                        if (!endDate || cellDate > endDate) return false;
                    }
                    return true;
                });
                return;
            }
            // Si no hay ni array ni rango, no filtrar
            return;
        } else {
            const value = moduleFilterValues[column];
            if (!value || (Array.isArray(value) && value.length === 0)) return;
            const condition = moduleFilterValues[`${column}_condition`] || 'contains';
            // Debug log para verificar condiciones NOT
            if (condition === 'not_contains' || condition === 'not_equals') {
                console.log(`🔍 Filtering column "${column}" with condition "${condition}" and value:`, value, 'type:', typeof value, 'isArray:', Array.isArray(value));
                console.log(`🔍 Total rows before NOT filter:`, filteredData.length);
            }
            const beforeCount = filteredData.length;
            filteredData = filteredData.filter(row => {
                const cellValue = row[column];
                // Si el valor de la celda es null/undefined/empty, tratarlo como string vacío para comparación
                const isEmpty = cellValue === null || cellValue === undefined || cellValue === '';
                if (isEmpty) {
                    // Para condiciones NOT: si el filtro es para un valor específico (no empty), los empty NO deben pasar
                    // porque el usuario quiere ver los que NO son empty (tienen valor)
                    if (condition === 'not_contains' || condition === 'not_equals') {
                        // Si el valor del filtro es un array y contiene '__EMPTY__', entonces los empty NO deben pasar
                        if (Array.isArray(value) && value.includes('__EMPTY__')) {
                            return false; // Los empty NO deben pasar si estamos excluyendo '__EMPTY__'
                        }
                        // Si el valor del filtro es un string específico (no empty), los empty NO deben pasar
                        // porque el usuario quiere ver los que tienen ese valor específico, no los empty
                        if (typeof value === 'string' && value !== '' && value !== '__EMPTY__') {
                            return false; // Los empty NO deben pasar - el usuario quiere ver los que NO son empty
                        }
                        // Si el valor del filtro es '__EMPTY__' o está vacío, los empty NO deben pasar
                        if (value === '__EMPTY__' || value === '') {
                            return false;
                        }
                        // Por defecto, los empty NO deben pasar con NOT
                        return false;
                    }
                    // Para condiciones normales (contains, equals), los empty no pasan
                    return false;
                }
                if (Array.isArray(value)) {
                    if (value.includes('__EMPTY__') && (cellValue === '' || cellValue === null || cellValue === undefined)) {
                        // Para condiciones NOT, si está vacío y buscamos empty, no debe pasar
                        if (condition === 'not_contains' || condition === 'not_equals') {
                            return false;
                        }
                        return true;
                    }
                    const isIncluded = value.includes(cellValue?.toString());
                    // Aplicar condición NOT
                    if (condition === 'not_contains' || condition === 'not_equals') {
                        return !isIncluded;
                    }
                    return isIncluded;
                }
                // Si el valor es un string (del input de texto), puede tener múltiples valores separados por comas
                let actualValue = value;
                let valuesToCheck = [];
                if (typeof value === 'string' && value.includes(',')) {
                    // Múltiples valores separados por comas
                    valuesToCheck = value.split(',').map(v => v.trim()).filter(v => v !== '');
                } else {
                    // Un solo valor
                    valuesToCheck = [value];
                }
                
                // Si hay múltiples valores, verificar si alguno coincide
                if (valuesToCheck.length > 1) {
                    let anyMatches = false;
                    for (const val of valuesToCheck) {
                        let matches = false;
                        switch (filterType) {
                            case 'text':
                                const normalizedCell = normalizeText(cellValue.toString());
                                const normalizedVal = normalizeText(val);
                                if (condition === 'equals' || condition === 'not_equals') {
                                    matches = normalizedCell === normalizedVal;
                                } else {
                                    matches = normalizedCell.includes(normalizedVal);
                                }
                                break;
                            case 'number':
                                const numVal = parseFloat(val);
                                const cellNum = parseFloat(cellValue);
                                if (condition === 'equals' || condition === 'not_equals') {
                                    matches = !isNaN(numVal) && !isNaN(cellNum) && cellNum === numVal;
                                } else {
                                    matches = !isNaN(numVal) && !isNaN(cellNum) && cellNum === numVal;
                                }
                                break;
                            case 'categorical':
                                if (condition === 'equals' || condition === 'not_equals') {
                                    matches = cellValue.toString() === val;
                                } else {
                                    matches = cellValue.toString().includes(val) || val.includes(cellValue.toString());
                                }
                                break;
                            default:
                                // Fallback: tratar como texto
                                const normalizedCell2 = normalizeText(cellValue.toString());
                                const normalizedVal2 = normalizeText(val);
                                if (condition === 'equals' || condition === 'not_equals') {
                                    matches = normalizedCell2 === normalizedVal2;
                                } else {
                                    matches = normalizedCell2.includes(normalizedVal2);
                                }
                        }
                        // Para NOT: si alguno coincide, la fila NO debe pasar (debe ser excluida)
                        // Para condiciones normales: si alguno coincide, la fila debe pasar (debe ser incluida)
                        if (condition === 'not_contains' || condition === 'not_equals') {
                            if (matches) {
                                anyMatches = true;
                                break; // Si alguno coincide, la fila NO debe pasar
                            }
                        } else {
                            if (matches) {
                                anyMatches = true;
                                break; // Si alguno coincide, la fila debe pasar
                            }
                        }
                    }
                    // Aplicar lógica NOT
                    // Para NOT: si anyMatches es true (algún valor coincide), la fila NO debe pasar (return false)
                    // Si anyMatches es false (ningún valor coincide), la fila debe pasar (return true)
                    if (condition === 'not_contains' || condition === 'not_equals') {
                        const shouldPass = !anyMatches; // Si ningún valor coincide, pasa el filtro
                        console.log(`🔍   Multiple values NOT check: valuesToCheck=${valuesToCheck.join(',')}, anyMatches=${anyMatches}, shouldPass=${shouldPass}`);
                        return shouldPass;
                    }
                    return anyMatches; // Si algún valor coincide, pasa el filtro
                }
                
                // Un solo valor
                let matches = false;
                switch (filterType) {
                    case 'text':
                        const normalizedCell = normalizeText(cellValue.toString());
                        const normalizedValue = normalizeText(actualValue);
                        if (condition === 'equals' || condition === 'not_equals') {
                            matches = normalizedCell === normalizedValue;
                        } else {
                            matches = normalizedCell.includes(normalizedValue);
                        }
                        break;
                    case 'number':
                        const numValue = parseFloat(actualValue);
                        const cellNum = parseFloat(cellValue);
                        if (condition === 'equals' || condition === 'not_equals') {
                            matches = !isNaN(numValue) && !isNaN(cellNum) && cellNum === numValue;
                        } else {
                            // Para números, contains/not_contains no tiene mucho sentido, usar equals
                            matches = !isNaN(numValue) && !isNaN(cellNum) && cellNum === numValue;
                        }
                        break;
                    case 'categorical':
                        // Si el valor es un string con comas, tratarlo como múltiples valores
                        const selectedValues = typeof actualValue === 'string' ? actualValue.split(',').map(v => v.trim()) : [actualValue];
                        if (condition === 'equals' || condition === 'not_equals') {
                            matches = selectedValues.includes(cellValue.toString());
                        } else {
                            matches = selectedValues.some(v => cellValue.toString().includes(v) || v.includes(cellValue.toString()));
                        }
                        break;
                    default:
                        // Fallback: tratar como texto
                        const normalizedCell3 = normalizeText(cellValue.toString());
                        const normalizedValue3 = normalizeText(actualValue);
                        if (condition === 'equals' || condition === 'not_equals') {
                            matches = normalizedCell3 === normalizedValue3;
                        } else {
                            matches = normalizedCell3.includes(normalizedValue3);
                        }
                }
                // Aplicar condición NOT
                if (condition === 'not_contains' || condition === 'not_equals') {
                    const shouldPass = !matches; // Si NO coincide, la fila debe pasar
                    // Debug para valores individuales
                    if (typeof value === 'string' && !Array.isArray(value)) {
                        console.log(`🔍   Single value NOT check: cellValue="${cellValue}", value="${actualValue}", matches=${matches}, shouldPass=${shouldPass}`);
                    }
                    return shouldPass;
                }
                // Para condiciones normales (contains, equals), devolver matches tal cual
                return matches;
            });
            // Debug log después del filtro NOT
            if (condition === 'not_contains' || condition === 'not_equals') {
                console.log(`🔍 Total rows after NOT filter:`, filteredData.length, 'removed:', beforeCount - filteredData.length);
            }
        }
    });

    // Aplicar filtros de la tabla
    Object.entries(tableActiveFilters).forEach(([column, filterType]) => {
        const value = tableFilterValues[column];
        if (!value || (Array.isArray(value) && value.length === 0)) return;
        filteredData = filteredData.filter(row => {
            const cellValue = row[column];
            if (cellValue === null || cellValue === undefined) return false;
            if (Array.isArray(value)) {
                if (value.includes('__EMPTY__') && (cellValue === '' || cellValue === null || cellValue === undefined)) {
                    return true;
                }
                return value.includes(cellValue?.toString());
            }
            return true;
        });
    });

    // Aplicar filtros de duplicados (usando variable global)
    if (window.activeDuplicateFilter && window.activeDuplicateFilter.duplicateKeys) {
        console.log('🔍 Applying global duplicate filter:', window.activeDuplicateFilter.name, 'with', window.activeDuplicateFilter.duplicateKeys.size, 'duplicate keys');
        console.log('🔍 Before duplicate filter, filteredData length:', filteredData.length);
        
        // Filter the current data using the duplicate keys
        const duplicateKeys = window.activeDuplicateFilter.duplicateKeys;
        const duplicateColumns = window.activeDuplicateFilter.duplicateColumns || [];
        
        if (duplicateColumns.length > 0) {
            filteredData = filteredData.filter(row => {
                const rowKey = duplicateColumns.map(col => row[col] || '').join('|');
                return duplicateKeys.has(rowKey);
            });
        }
        
        console.log('🔍 After duplicate filter, filteredData length:', filteredData.length);
    }

    // --- FILTRO GLOBAL ---
    const globalSearch = moduleFilterValues['__globalSearch'];
    if (globalSearch && typeof globalSearch === 'string' && globalSearch.trim()) {
        // Split by comma, trim, lowercase, ignore empty terms
        const terms = globalSearch
            .split(',')
            .map(t => t.trim().toLowerCase())
            .filter(Boolean);
        
        if (terms.length > 0) {
            filteredData = filteredData.filter(row =>
                Object.values(row).some(val =>
                    terms.some(term =>
                        (val || '').toString().toLowerCase().includes(term)
                    )
                )
            );
        }
    }

    const sortConfig = getSortConfig();
    if (sortConfig && sortConfig.column) {
        filteredData = sortData(filteredData, sortConfig.column, sortConfig.direction);
    }

    displayTable(filteredData);
    return filteredData;
}

function resolveDynamicValue(value) {
  if (typeof value === 'string' && value.startsWith('__TODAY__')) {
    return resolveDynamicDateExpr(value);
  }
  if (value === "__TODAY__") {
    const today = new Date();
    return today.toISOString().split('T')[0];
  }
  return value;
}

// Setup filter tabs
function setupFilterTabs() {
  try {
    const tabs = document.querySelectorAll('.filter-tab');
    const panels = document.querySelectorAll('.filter-panel');

    tabs.forEach(tab => {
      tab.addEventListener('click', () => {
        try {
          // Remove active class from all tabs and panels
          tabs.forEach(t => t.classList.remove('active'));
          panels.forEach(p => p.classList.remove('active'));

          // Add active class to clicked tab and corresponding panel
          tab.classList.add('active');
          const targetPanel = document.getElementById(tab.dataset.target + 'FilterPanel');
          if (targetPanel) {
            targetPanel.classList.add('active');
          }
        } catch (error) {
          console.error('Error switching filter tab:', error);
        }
      });
    });

    // Ensure Generic tab is active by default if no other tab is active
    if (!Array.from(tabs).some(tab => tab.classList.contains('active'))) {
      const genericTab = Array.from(tabs).find(tab => tab.dataset.target === 'generic');
      const genericPanel = document.getElementById('genericFilterPanel');
      if (genericTab && genericPanel) {
        genericTab.classList.add('active');
        genericPanel.classList.add('active');
      }
    }
  } catch (error) {
    console.error('Error setting up filter tabs:', error);
  }
}

// Initialize the filter manager
export function initializeFilterManager(initialData) {
  if (!instance) {
    instance = {
      initializeData,
      applyFilters,
      // ... other exported functions ...
    };
  }

  if (initializeData(initialData)) {
    // Forzar reset y regeneración de filtros tras un pequeño delay
    setTimeout(() => {
      resetFilterManager();
      const headers = Object.keys(initialData[0] || {});
      generateFilterSidebar(headers);
      setupFilterTabs();
      applyFilters();
    }, 30); // 30 ms de retraso
  }

  return instance;
}

export function resetFilterManager() {
  setModuleActiveFilters({});
  setModuleFilterValues({});
  document.querySelectorAll('.filter-input').forEach(input => { input.value = ''; });
  document.querySelectorAll('.filter-input[type="date"]').forEach(input => { input.value = ''; });
  document.querySelectorAll('.filter-checkbox').forEach(checkbox => { checkbox.checked = false; });
  document.querySelectorAll('.column-selector').forEach(select => { select.value = ''; });
  // Quitar resaltado de todos los filtros
  document.querySelectorAll('.filter-item').forEach(item => { item.classList.remove('active'); });
  // Limpiar los chips de filtros activos
  const summary = document.getElementById('activeFiltersSummary');
  if (summary) {
    summary.innerHTML = '';
    summary.style.display = 'none';
  }
  // Limpiar el resumen global de filtros
  const globalSummary = document.getElementById('globalActiveFiltersSummary');
  if (globalSummary) {
    globalSummary.innerHTML = '';
  }
  updateActiveFiltersSummary();
  applyFilters();
  renderActiveFiltersSummaryChips();
}

// Helper function to get condition label text
function getConditionLabel(condition) {
  const conditionLabels = {
    'contains': '',
    'equals': '=',
    'not_contains': 'NOT',
    'not_equals': 'NOT ='
  };
  return conditionLabels[condition] || '';
}

// --- Render active filter chips above the table ---
export function renderActiveFiltersSummaryChips() {
  const summary = document.getElementById('activeFiltersSummary');
  if (!summary) return;
  summary.innerHTML = '';
  const filterValues = getModuleFilterValues();
  if (Object.keys(filterValues).length === 0) {
    summary.style.display = 'flex';
    updateFilterButtonBadge(0);
    return;
  }
  summary.style.display = 'flex';

  // Group filters for display
  const dateFilters = {};
  const otherFilters = {};
  Object.entries(filterValues).forEach(([key, value]) => {
    if (key.endsWith('_start') || key.endsWith('_end') || key.endsWith('_empty')) {
      const baseKey = key.replace(/_(start|end|empty)$/, '');
      if (!dateFilters[baseKey]) {
        dateFilters[baseKey] = { start: '', end: '', empty: false };
      }
      if (key.endsWith('_start')) dateFilters[baseKey].start = value;
      if (key.endsWith('_end')) dateFilters[baseKey].end = value;
      if (key.endsWith('_empty')) dateFilters[baseKey].empty = value;
    } else if (!key.endsWith('_condition')) {
      // Excluir condiciones del objeto otherFilters
      otherFilters[key] = value;
    }
  });

  let count = 0;
  // Date filters
  Object.entries(dateFilters).forEach(([column, filter]) => {
    if (filter.start || filter.end || filter.empty) {
      count++;
      const tag = document.createElement('div');
      tag.className = 'filter-tag';
      let text = `${column}: `;
      if (filter.empty && !filter.start && !filter.end) {
        text += '(empty)';
      } else {
        if (filter.start) text += `from ${prettyDynamicDate(filter.start)}`;
        if (filter.end) text += ` to ${prettyDynamicDate(filter.end)}`;
        if (filter.empty) text += ' (including empty)';
      }
      tag.innerHTML = `
        <span>${text}</span>
        <button class="filter-tag-remove" data-column="${column}">×</button>
      `;
      summary.appendChild(tag);
    }
  });
  // Other filters
  Object.entries(otherFilters).forEach(([column, values]) => {
    // Mostrar tanto arrays como strings (valores únicos del input de texto)
    if (Array.isArray(values) && values.length > 0) {
      count++;
      const tag = document.createElement('div');
      tag.className = 'filter-tag';
      // Obtener la condición para esta columna
      const condition = filterValues[`${column}_condition`] || 'contains';
      const conditionLabel = getConditionLabel(condition);
      const conditionText = conditionLabel ? `${conditionLabel} ` : '';
      tag.innerHTML = `
        <span>${conditionText}${column}: ${values.join(', ')}</span>
        <button class="filter-tag-remove" data-column="${column}">×</button>
      `;
      summary.appendChild(tag);
    } else if (typeof values === 'string' && values.trim() !== '') {
      // Mostrar valores de tipo string (del input de texto)
      count++;
      const tag = document.createElement('div');
      tag.className = 'filter-tag';
      // Obtener la condición para esta columna
      const condition = filterValues[`${column}_condition`] || 'contains';
      const conditionLabel = getConditionLabel(condition);
      const conditionText = conditionLabel ? `${conditionLabel} ` : '';
      tag.innerHTML = `
        <span>${conditionText}${column}: ${values}</span>
        <button class="filter-tag-remove" data-column="${column}">×</button>
      `;
      summary.appendChild(tag);
    }
  });
  updateFilterButtonBadge(count);

  // Remove filter on chip click
  summary.querySelectorAll('.filter-tag-remove').forEach(btn => {
    btn.addEventListener('click', () => {
      const column = btn.dataset.column;
        const updated = { ...getModuleFilterValues() };
      // Eliminar todos los valores posibles del filtro
          delete updated[`${column}_start`];
          delete updated[`${column}_end`];
          delete updated[`${column}_empty`];
          delete updated[column];
          delete updated[`${column}_condition`];
          setModuleFilterValues(updated);
      // Quitar resaltado si ya no hay nada filtrado
        const filterItem = document.querySelector(`.filter-item[data-column="${column}"]`);
        if (filterItem) {
        filterItem.querySelectorAll('input[type="date"]').forEach(input => {
          input.value = '';
          input.removeAttribute('data-dynamic');
        });
          filterItem.querySelectorAll('.filter-checkbox').forEach(checkbox => checkbox.checked = false);
        if (!getModuleFilterValues()[`${column}_start`] && !getModuleFilterValues()[`${column}_end`] && !getModuleFilterValues()[`${column}_empty`]) {
          filterItem.classList.remove('active');
        }
      }
      // Quitar el tipo de filtro si ya no hay nada filtrado
      const activeFilters = { ...getModuleActiveFilters() };
      if (!updated[`${column}_start`] && !updated[`${column}_end`] && !updated[`${column}_empty`]) {
        delete activeFilters[column];
        setModuleActiveFilters(activeFilters);
      }
      renderActiveFiltersSummaryChips();
      updateActiveFiltersSummary();
      applyFilters();
    });
  });
}

// --- Badge on filter button ---
function updateFilterButtonBadge(count) {
  // Use the main filter toggle button
  let btn = document.querySelector('#toggleFiltersBtn');
  if (!btn) {
    // Fallback: Try to find by text content
    btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim().toLowerCase().includes('filter'));
  }
  if (!btn) return;
  let badge = btn.querySelector('.filter-badge');
  if (!badge) {
    badge = document.createElement('span');
    badge.className = 'filter-badge';
    btn.appendChild(badge);
    btn.style.position = 'relative';
  }
  if (count > 0) {
    badge.textContent = count;
    badge.style.display = '';
  } else {
    badge.textContent = '';
    badge.style.display = 'none';
  }
  
  // Update reset button state
  updateResetButtonState(count);
}

// --- Update reset button appearance based on active filters ---
function updateResetButtonState(count) {
  // Find all reset filter buttons
  const resetButtons = document.querySelectorAll('.reset-filters-btn, #resetFiltersBtn, #resetAllFiltersBtn');
  
  resetButtons.forEach(btn => {
    if (count > 0) {
      btn.classList.add('has-active-filters');
    } else {
      btn.classList.remove('has-active-filters');
    }
  });
}

// --- Hook into filter changes ---
// Call renderActiveFiltersSummaryChips whenever filters change
const oldUpdateActiveFiltersSummary = window.updateActiveFiltersSummary;
window.updateActiveFiltersSummary = function() {
  if (oldUpdateActiveFiltersSummary) oldUpdateActiveFiltersSummary();
  renderActiveFiltersSummaryChips();
};
// Initial render
renderActiveFiltersSummaryChips();

// --- Utilidad para parsear expresiones tipo __TODAY__+N o __TODAY__-N ---
function resolveDynamicDateExpr(expr) {
  if (typeof expr !== 'string') return expr;
  const today = new Date();
  if (expr.startsWith('__TODAY__')) {
    let offset = 0;
    const match = expr.match(/__TODAY__(?:([+-])(\d+))?/);
    if (match && match[1] && match[2]) {
      offset = parseInt(match[2], 10) * (match[1] === '-' ? -1 : 1);
    }
    today.setDate(today.getDate() + offset);
    return today.toISOString().slice(0, 10);
  }
  return expr;
}

// --- Quick Filters API for Dashboard ---
function loadQuickFilters() {
  try {
    const saved = localStorage.getItem('quickFilters');
    return saved ? JSON.parse(saved) : {};
  } catch {
    return {};
  }
}
function saveQuickFilter(name, urgencyCard, container, containerTitle, hubType = 'ops') {
  const headers = Object.keys(getOriginalData()[0] || {});
  const filterValues = { ...getModuleFilterValues() };
  const activeFilters = { ...getModuleActiveFilters() };
  // Asegurar que todas las condiciones se guarden
  const conditionKeys = Object.keys(filterValues).filter(k => k.endsWith('_condition'));
  console.log('💾 Saving quick filter:', name, 'with condition keys:', conditionKeys);
  conditionKeys.forEach(ck => {
    const col = ck.replace('_condition', '');
    console.log(`💾   - ${col}: condition = "${filterValues[ck]}"`);
  });
  const quickFilters = loadQuickFilters();
  const filterObj = { filterValues, activeFilters, headers, hubType };
  if (urgencyCard && urgencyCard !== 'Ninguna') filterObj.linkedUrgencyCard = urgencyCard;
  if (container) filterObj.container = container;
  if (containerTitle) filterObj.containerTitle = containerTitle;
  quickFilters[name] = filterObj;
  localStorage.setItem('quickFilters', JSON.stringify(quickFilters));
  
  // Trigger auto-save
  if (typeof window.triggerAutoSave === 'function') {
    window.triggerAutoSave();
  }
  
  // Show notification
  if (typeof window.showUnifiedNotification === 'function') {
    window.showUnifiedNotification(`Quick filter "${name}" saved successfully!`, 'success');
  }
}
function deleteQuickFilter(name) {
  const quickFilters = loadQuickFilters();
  delete quickFilters[name];
  localStorage.setItem('quickFilters', JSON.stringify(quickFilters));
  
  // Trigger auto-save
  if (typeof window.triggerAutoSave === 'function') {
    window.triggerAutoSave();
  }
  
  // Show notification
  if (typeof window.showUnifiedNotification === 'function') {
    window.showUnifiedNotification(`Quick filter "${name}" deleted successfully!`, 'success');
  }
}
function getFilteredData() {
  const data = getOriginalData();
  const activeFilters = getModuleActiveFilters();
  const filterValues = getModuleFilterValues();
  if (!data || !data.length) return [];
  let filteredData = [...data];

  // --- BÚSQUEDA GLOBAL MULTITÉRMINO (OR, coma como separador) ---
  const globalSearch = filterValues['__globalSearch'];
  if (globalSearch && globalSearch.trim()) {
    const terms = globalSearch.split(',').map(t => t.trim().toLowerCase()).filter(Boolean);
    if (terms.length > 0) {
      filteredData = filteredData.filter(row =>
        terms.some(term =>
          Object.values(row).some(
            value => value && value.toString().toLowerCase().includes(term)
          )
        )
      );
    }
    // Si no hay términos válidos, no filtra (muestra todo)
  }

  // --- RESTO DE FILTROS POR COLUMNA ---
  Object.entries(activeFilters).forEach(([column, filterType]) => {
    if (filterType === 'date') {
      const start = resolveDynamicValue(filterValues[`${column}_start`]);
      const end = resolveDynamicValue(filterValues[`${column}_end`]);
      const empty = filterValues[`${column}_empty`];
      filteredData = filteredData.filter(row => {
        const value = row[column] ? row[column].trim() : "";
        if (empty && !start && !end) {
            return value === "" || value === null || value === undefined;
        }
        if (empty && (value === "" || value === null || value === undefined)) return true;
        if (!value) return false;
        const cellDate = parseFlexibleDate(value);
        if (!cellDate) return false;
        if (start) {
          const startDate = parseFlexibleDate(start);
          if (!startDate || cellDate < startDate) return false;
        }
        if (end) {
          const endDate = parseFlexibleDate(end);
          if (!endDate || cellDate > endDate) return false;
        }
        return true;
      });
      return;
    }
    const value = filterValues[column];
    if (!value || (Array.isArray(value) && value.length === 0)) return;
    filteredData = filteredData.filter(row => {
      const cellValue = row[column];
      if (cellValue === null || cellValue === undefined) return false;
      if (Array.isArray(value)) {
        if (value.includes('__EMPTY__') && (cellValue === '' || cellValue === null || cellValue === undefined)) {
          return true;
        }
        return value.includes(cellValue?.toString());
      }
      switch (filterType) {
        case 'text':
          return normalizeText(cellValue.toString()).includes(normalizeText(value));
        case 'number':
          const numValue = parseFloat(value);
          const cellNum = parseFloat(cellValue);
          return !isNaN(numValue) && !isNaN(cellNum) && cellNum === numValue;
        case 'categorical':
          const selectedValues = value.split(',').map(v => v.trim());
          return selectedValues.includes(cellValue.toString());
        default:
          return true;
      }
    });
  });
  const sortConfig = getSortConfig();
  if (sortConfig && sortConfig.column) {
    filteredData = sortData(filteredData, sortConfig.column, sortConfig.direction);
  }
  return filteredData;
}

// Utilidad robusta para convertir fechas a YYYY-MM-DD
function toISODateString(val) {
  if (!val) return val;
  // Si es un número puro de 1 a 3 dígitos, nunca es fecha
  if (/^\d{1,3}$/.test(val.trim())) return val;
  
  // Intenta parsear como DD.MM.YY (formato corto con puntos)
  const shortDateWithDots = val.match(/^(\d{2})\.(\d{2})\.(\d{2})$/);
  if (shortDateWithDots) {
    const [, day, month, year] = shortDateWithDots;
    // Asumir que años 00-29 son 2000-2029, años 30-99 son 1930-1999
    const fullYear = parseInt(year) < 30 ? `20${year}` : `19${year}`;
    const date = new Date(`${fullYear}-${month}-${day}`);
    if (!isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  }
  
  // Intenta parsear como DD-MM-YYYY o DD/MM/YYYY
  const match = val.match(/^([0-3]?\d)[-/]([0-1]?\d)[-/](\d{4})$/);
  if (match) {
    const [, d, m, y] = match;
    const date = new Date(`${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`);
    if (!isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  }
  // Intenta parsear como YYYY-MM-DD
  const iso = val.match(/^(\d{4})[-/](\d{2})[-/](\d{2})$/);
  if (iso) return val.replace(/\//g, '-');
  // Intenta parsear como fecha JS
  const d = new Date(val);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  return val;
}

// Utilidad para mostrar fechas dinámicas de forma legible
function prettyDynamicDate(val) {
  if (typeof val !== 'string') return val;
  if (val === '__TODAY__') return 'Today';
  const match = val.match(/^__TODAY__(\+|-)(\d+)$/);
  if (match) {
    const sign = match[1] === '+' ? '+ ' : '- ';
    return `Today ${sign}${match[2]} days`;
  }
  return val;
}

// Function to set column filter and validate column compatibility
function setColumnFilter(column, filterValues) {
  try {
    console.log(`🔍 Setting filter for column: ${column}`, filterValues);
    
    // Check if the column exists in current data
    const currentHeaders = getCurrentHeaders ? getCurrentHeaders() : [];
    const columnExists = currentHeaders.includes(column);
    
    if (!columnExists) {
      console.warn(`⚠️ Column "${column}" not found in current data. Available columns:`, currentHeaders);
      
      // Find the filter element and mark it as incompatible
      const filterElement = document.querySelector(`[data-column="${column}"]`);
      if (filterElement) {
        filterElement.classList.add('filter-incompatible');
        filterElement.style.opacity = '0.5';
        filterElement.style.pointerEvents = 'none';
        
        // Add warning tooltip
        const warningTooltip = document.createElement('div');
        warningTooltip.className = 'filter-warning-tooltip';
        warningTooltip.textContent = `Column "${column}" not available in current data`;
        warningTooltip.style.cssText = `
          position: absolute;
          background: #ff6b6b;
          color: white;
          padding: 4px 8px;
          border-radius: 4px;
          font-size: 12px;
          z-index: 1000;
          pointer-events: none;
          white-space: nowrap;
          top: -30px;
          left: 50%;
          transform: translateX(-50%);
          opacity: 0;
          transition: opacity 0.3s;
        `;
        
        filterElement.appendChild(warningTooltip);
        
        // Show tooltip on hover
        filterElement.addEventListener('mouseenter', () => {
          warningTooltip.style.opacity = '1';
        });
        
        filterElement.addEventListener('mouseleave', () => {
          warningTooltip.style.opacity = '0';
        });
      }
      
      return false; // Indicate that filter couldn't be applied
    }
    
    // Column exists, apply the filter normally
    setModuleFilterValues({ ...getModuleFilterValues(), [column]: filterValues });
    setModuleActiveFilters({ ...getModuleActiveFilters(), [column]: 'categorical' });
    
    // Update the filter element to show it's active
    const filterElement = document.querySelector(`[data-column="${column}"]`);
    if (filterElement) {
      filterElement.classList.remove('filter-incompatible');
      filterElement.style.opacity = '1';
      filterElement.style.pointerEvents = 'auto';
      
      // Remove warning tooltip if it exists
      const warningTooltip = filterElement.querySelector('.filter-warning-tooltip');
      if (warningTooltip) {
        warningTooltip.remove();
      }
      
      // Mark as active
      filterElement.classList.add('active');
    }
    
    console.log(`✅ Filter applied successfully for column: ${column}`);
    return true; // Indicate success
    
  } catch (error) {
    console.error(`❌ Error setting filter for column ${column}:`, error);
    return false;
  }
}

// Function to validate and update filter compatibility
function validateFilterCompatibility() {
  try {
    console.log('🔍 Validating filter compatibility...');
    
    const currentHeaders = getCurrentHeaders ? getCurrentHeaders() : [];
    const activeFilters = getModuleActiveFilters ? getModuleActiveFilters() : {};
    const filterValues = getModuleFilterValues ? getModuleFilterValues() : {};
    
    let incompatibleCount = 0;
    const incompatibleColumns = [];
    
    // Check each active filter
    Object.keys(activeFilters).forEach(column => {
      const columnExists = currentHeaders.includes(column);
      const filterElement = document.querySelector(`[data-column="${column}"]`);
      
      if (!columnExists && filterElement) {
        // Mark as incompatible
        filterElement.classList.add('filter-incompatible');
        filterElement.style.opacity = '0.5';
        filterElement.style.pointerEvents = 'none';
        incompatibleCount++;
        incompatibleColumns.push(column);
        
        console.log(`⚠️ Filter for column "${column}" is incompatible with current data`);
      } else if (columnExists && filterElement) {
        // Mark as compatible
        filterElement.classList.remove('filter-incompatible');
        filterElement.style.opacity = '1';
        filterElement.style.pointerEvents = 'auto';
      }
    });
    
    // Clean up incompatible filters from state
    if (incompatibleColumns.length > 0) {
      const updatedActiveFilters = { ...activeFilters };
      const updatedFilterValues = { ...filterValues };
      
      incompatibleColumns.forEach(column => {
        delete updatedActiveFilters[column];
        delete updatedFilterValues[column];
      });
      
      // Update state
      if (typeof setModuleActiveFilters === 'function') {
        setModuleActiveFilters(updatedActiveFilters);
      }
      if (typeof setModuleFilterValues === 'function') {
        setModuleFilterValues(updatedFilterValues);
      }
      
      console.log(`🧹 Cleaned up ${incompatibleColumns.length} incompatible filters from state`);
    }
    
    // Show notification if there are incompatible filters
    if (incompatibleCount > 0) {
      if (typeof window.showUnifiedNotification === 'function') {
        window.showUnifiedNotification(
          `${incompatibleCount} filter(s) were incompatible with the new CSV and have been removed.`,
          'warning'
        );
      }
    }
    
    console.log(`✅ Filter compatibility validation complete. ${incompatibleCount} incompatible filters found and cleaned up.`);
    return incompatibleCount;
    
  } catch (error) {
    console.error('❌ Error validating filter compatibility:', error);
    return 0;
  }
}

export { 
  getFilteredData, 
  loadQuickFilters, 
  saveQuickFilter, 
  deleteQuickFilter, 
  applyFilters, 
  generateFilterSidebar, 
  detectColumnTypes, 
  parseFlexibleDate,
  setColumnFilter,
  validateFilterCompatibility
}; 