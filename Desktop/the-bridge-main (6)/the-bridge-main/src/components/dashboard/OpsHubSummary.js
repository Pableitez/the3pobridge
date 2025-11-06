/**
 * Ops Hub Summary Manager
 * Handles the generation and management of professional summary reports for the Operations Hub
 */

// Import getFilteredData to include table data in the report
import { getFilteredData } from '../filters/FilterManager.js';
import { loadQuickFilters } from '../filters/FilterManager.js';
import { getOriginalData, getModuleActiveFilters, setModuleActiveFilters, getModuleFilterValues, setModuleFilterValues } from '../../store/index.js';

export class OpsHubSummary {
  constructor() {
    this.selectedCards = new Set();
    this.quickCardsData = {};
    this.currentPriority = '';
    this.cardConfigs = {}; // Store individual card configurations
    this.reportConfig = {
      includeTableData: true,
      viewType: 'both',
      removeDuplicates: false,
      duplicateColumns: [],
      tableDataLimit: 50,
      tableView: 'current',
      separateTablesPerCard: false
    };
    
    // Performance optimizations
    this.previewUpdateTimeout = null;
    this.cachedData = new Map();
    this.lastConfigHash = '';
    this.isGeneratingPreview = false;
    this.duplicateAnalysisTimeout = null;
    this.init();
  }

  init() {
    this.bindEvents();
    this.setupModalHandlers();
  }

  bindEvents() {
    // Generate summary button in dashboard
    const generateBtn = document.getElementById('generateOpsSummaryBtn');
    if (generateBtn) {
      generateBtn.addEventListener('click', () => this.openSummaryModal());
    }

    // Modal close button
    const closeBtn = document.getElementById('closeOpsSummaryBtn');
    if (closeBtn) {
      closeBtn.addEventListener('click', () => this.closeSummaryModal());
    }

    // Action buttons
    const copyBtn = document.getElementById('copySummaryBtn');
    const pdfBtn = document.getElementById('exportSummaryPdfBtn');
    const excelBtn = document.getElementById('exportSummaryExcelBtn');

    if (copyBtn) copyBtn.addEventListener('click', () => this.copyToClipboard());
    if (pdfBtn) pdfBtn.addEventListener('click', () => this.exportToPdf());
    if (excelBtn) excelBtn.addEventListener('click', () => this.exportToExcel());

    // Report options
    const includeTableDataCheckbox = document.getElementById('includeTableDataCheckbox');
    if (includeTableDataCheckbox) {
      includeTableDataCheckbox.addEventListener('change', () => this.renderSummaryPreview());
    }
    
    const includeTechnicalInfoCheckbox = document.getElementById('includeTechnicalInfoCheckbox');
    if (includeTechnicalInfoCheckbox) {
      includeTechnicalInfoCheckbox.addEventListener('change', () => this.renderSummaryPreview());
    }

    // View type selector
    const viewTypeSelect = document.getElementById('reportViewTypeSelect');
    if (viewTypeSelect) {
      viewTypeSelect.addEventListener('change', () => {
        console.log(`View type changed for card ${cardData.id} to: ${viewTypeSelect.value}`);
        this.updateCardConfig(cardData.id, 'viewType', viewTypeSelect.value);
        
        // Show immediate feedback
        const cardDiv = viewTypeSelect.closest('.card-selection-item');
        if (cardDiv) {
          cardDiv.style.borderColor = '#47B2E5';
          cardDiv.style.backgroundColor = 'rgba(71, 178, 229, 0.1)';
          setTimeout(() => {
            cardDiv.style.backgroundColor = 'rgba(71, 178, 229, 0.05)';
          }, 200);
        }
        
        // Force immediate preview update
        this.renderSummaryPreview();
      });
    }

    // Duplicate removal
    const removeDuplicatesCheckbox = document.getElementById('removeDuplicatesCheckbox');
    if (removeDuplicatesCheckbox) {
      removeDuplicatesCheckbox.addEventListener('change', () => this.handleDuplicateRemovalChange());
    }

    // Table data limit
    const tableDataLimitSelect = document.getElementById('tableDataLimitSelect');
    if (tableDataLimitSelect) {
      tableDataLimitSelect.addEventListener('change', () => this.renderSummaryPreview());
    }



    // Separate tables per card
    const separateTablesPerCardCheckbox = document.getElementById('separateTablesPerCardCheckbox');
    if (separateTablesPerCardCheckbox) {
      separateTablesPerCardCheckbox.addEventListener('change', () => this.renderSummaryPreview());
    }

    // Duplicate columns controls
    const selectAllColumnsBtn = document.getElementById('selectAllDuplicateColumnsBtn');
    const deselectAllColumnsBtn = document.getElementById('deselectAllDuplicateColumnsBtn');
    const selectCommonFieldsBtn = document.getElementById('selectCommonFieldsBtn');
    const duplicateColumnsSearch = document.getElementById('duplicateColumnsSearch');

    if (selectAllColumnsBtn) {
      selectAllColumnsBtn.addEventListener('click', () => this.selectAllDuplicateColumns());
    }
    if (deselectAllColumnsBtn) {
      deselectAllColumnsBtn.addEventListener('click', () => this.deselectAllDuplicateColumns());
    }
    if (selectCommonFieldsBtn) {
      selectCommonFieldsBtn.addEventListener('click', () => this.selectCommonFields());
    }
    if (duplicateColumnsSearch) {
      duplicateColumnsSearch.addEventListener('input', () => this.filterDuplicateColumns());
    }

    // Preferences buttons
    const savePreferencesBtn = document.getElementById('opsSavePreferencesBtn');
    const loadPreferencesBtn = document.getElementById('opsLoadPreferencesBtn');

    if (savePreferencesBtn) {
      savePreferencesBtn.addEventListener('click', () => this.saveOpsPreferences());
    }
    if (loadPreferencesBtn) {
      loadPreferencesBtn.addEventListener('click', () => this.loadOpsPreferences());
    }
  }

  setupModalHandlers() {
    const modal = document.getElementById('opsSummaryModal');
    if (modal) {
      modal.addEventListener('click', (e) => {
        if (e.target === modal) {
          this.closeSummaryModal();
        }
      });
    }
  }

  openSummaryModal() {
    // Restore current state before opening
    this.restoreCurrentState();
    
    // Force refresh of quick filters data before collecting
    if (typeof window.renderDashboardQuickFilters === 'function') {
      window.renderDashboardQuickFilters();
    }
    
    // Wait a bit for the DOM to update, then collect data
    setTimeout(() => {
      this.collectQuickCardsData();
      this.updateSummaryInfo();
      this.renderCardsSelection();

      this.renderSummaryPreview();
      
      // Debug quick filters
      this.debugQuickFilters();
      
      const modal = document.getElementById('opsSummaryModal');
      if (modal) {
        modal.classList.remove('hidden');
      }
    }, 100);
  }

  closeSummaryModal() {
    const modal = document.getElementById('opsSummaryModal');
    if (modal) {
      modal.classList.add('hidden');
    }
    
    // Save current state to localStorage before clearing
    this.saveCurrentState();
    
    // Cleanup timeouts to prevent memory leaks
    this.cleanup();
  }
  
  saveCurrentState() {
    try {
      const currentState = {
        selectedCards: Array.from(this.selectedCards),
        cardConfigs: this.cardConfigs,
        timestamp: new Date().toISOString()
      };
      localStorage.setItem('opsSummaryCurrentState', JSON.stringify(currentState));
      console.log('Ops Current state saved:', currentState);
    } catch (error) {
      console.error('Ops Error saving current state:', error);
    }
  }
  
  restoreCurrentState() {
    try {
      const savedState = localStorage.getItem('opsSummaryCurrentState');
      if (savedState) {
        const state = JSON.parse(savedState);
        
        // Restore selected cards
        if (state.selectedCards && state.selectedCards.length > 0) {
          this.selectedCards.clear();
          state.selectedCards.forEach(cardId => {
            this.selectedCards.add(cardId);
          });
        }
        
        // Restore card configurations
        if (state.cardConfigs) {
          this.cardConfigs = { ...this.cardConfigs, ...state.cardConfigs };
        }
        
        console.log('Ops Current state restored:', state);
        return true;
      }
    } catch (error) {
      console.error('Ops Error restoring current state:', error);
    }
    return false;
  }

  cleanup() {
    // Clear all timeouts
    if (this.previewUpdateTimeout) {
      clearTimeout(this.previewUpdateTimeout);
      this.previewUpdateTimeout = null;
    }
    
    if (this.duplicateAnalysisTimeout) {
      clearTimeout(this.duplicateAnalysisTimeout);
      this.duplicateAnalysisTimeout = null;
    }
    
    // Reset flags
    this.isGeneratingPreview = false;
    
    // Clear cache to free memory
    this.clearCache();
  }

  collectQuickCardsData() {
    console.log('Collecting quick cards data...');
    this.quickCardsData = {};
    this.currentPriority = '';

    // Get current priority filter
    const activeChip = document.querySelector('.ops-hub-chip.active');
    if (activeChip) {
      this.currentPriority = activeChip.getAttribute('data-urgency');
      console.log('Current priority filter:', this.currentPriority);
    } else {
      console.log('No active priority filter found');
    }

    // Collect data from quick cards - ONLY from Ops Hub containers
    // Look specifically in the Ops Hub dashboard modal
    let quickCards = document.querySelectorAll('#dashboardModal .quickfilter-cards-container .kpi-card');
    
    // If no cards found, try alternative selectors within Ops Hub
    if (quickCards.length === 0) {
      quickCards = document.querySelectorAll('#dashboardModal .quickfilters-grid .kpi-card');
    }
    
    // If still no cards, try any kpi-card specifically in the Ops Hub dashboard
    if (quickCards.length === 0) {
      quickCards = document.querySelectorAll('#dashboardModal .kpi-card');
    }
    
    // Additional filter to ensure we only get Ops Hub cards (not DQ Hub cards)
    quickCards = Array.from(quickCards).filter(card => {
      // Check if the card is within the Ops Hub context
      const isInOpsHub = card.closest('#dashboardModal') && !card.closest('#dqDashboardModal');
      return isInOpsHub;
    });

    quickCards.forEach(card => {
      const cardId = card.getAttribute('data-card-id') || this.generateCardId(card);
      
      // Get card name from kpi-title
      let cardName = card.querySelector('.kpi-title')?.textContent;
      if (!cardName) {
        cardName = card.querySelector('[class*="title"]')?.textContent;
      }
      if (!cardName) {
        cardName = card.querySelector('h4, h5, h6')?.textContent;
      }
      if (!cardName) {
        cardName = 'Quick Filter Card';
      }

      // Get count from kpi-value
      let countElement = card.querySelector('.kpi-value');
      if (!countElement) {
        countElement = card.querySelector('[class*="value"]');
      }
      if (!countElement) {
        countElement = card.querySelector('.count, [data-count]');
      }
      
      let count = 0;
      if (countElement) {
        const countText = countElement.textContent.trim();
        // Handle cases where count might be "-" or other non-numeric values
        if (countText !== '-' && countText !== '') {
          count = parseInt(countText.replace(/,/g, '')) || 0;
        }
      }
      
      const isActive = card.classList.contains('active');

      this.quickCardsData[cardId] = {
        id: cardId,
        name: cardName,
        count: count,
        active: isActive,
        element: card
      };
    });

    // If no cards found, create a default message
    if (Object.keys(this.quickCardsData).length === 0) {
      this.quickCardsData['no_cards'] = {
        id: 'no_cards',
        name: 'No Quick Cards Available',
        count: 0,
        active: false,
        element: null
      };
    }
    
    console.log('Collected quick cards data:', this.quickCardsData);
    console.log('Total cards found:', Object.keys(this.quickCardsData).length);
  }

  generateCardId(card) {
    const title = card.querySelector('.kpi-title')?.textContent || '';
    return `card_${title.toLowerCase().replace(/[^a-z0-9]/g, '_')}_${Date.now()}`;
  }

  updateSummaryInfo() {
    const dateTimeElement = document.getElementById('summaryDateTime');
    const priorityElement = document.getElementById('summaryPriority');

    if (dateTimeElement) {
      const now = new Date();
      dateTimeElement.textContent = now.toLocaleString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
    }

    if (priorityElement) {
      priorityElement.textContent = this.currentPriority || 'All Priorities';
    }
  }

  renderCardsSelection() {
    const cardsContainer = document.getElementById('quickCardsSelection');
    if (!cardsContainer) return;

    cardsContainer.innerHTML = '';

    // Add selection controls FIRST (above the cards)
    const controlsDiv = document.createElement('div');
    controlsDiv.className = 'cards-selection-controls';
    controlsDiv.style.cssText = `
      display: flex;
      gap: 1rem;
      margin-bottom: 1.5rem;
      padding: 1rem;
      background: rgba(26, 35, 50, 0.03);
      border-radius: 8px;
      flex-wrap: wrap;
    `;

    controlsDiv.innerHTML = `
      <button type="button" class="select-all-btn" style="
        background: #1a2332;
        color: white;
        border: none;
        padding: 0.6rem 1.2rem;
        border-radius: 6px;
        cursor: pointer;
        font-size: 0.9rem;
        transition: all 0.2s;
      ">Select All</button>
      
      <button type="button" class="deselect-all-btn" style="
        background: #f5f5f5;
        color: #666;
        border: 1px solid #ddd;
        padding: 0.6rem 1.2rem;
        border-radius: 6px;
        cursor: pointer;
        font-size: 0.9rem;
        transition: all 0.2s;
      ">Deselect All</button>
      

    `;

    // Insert controls at the top
    cardsContainer.appendChild(controlsDiv);

    // Now render the cards
    Object.values(this.quickCardsData).forEach(cardData => {
      const cardElement = this.createCardSelectionItem(cardData);
      cardsContainer.appendChild(cardElement);
    });

    // Bind control events
    controlsDiv.querySelector('.select-all-btn').addEventListener('click', () => this.selectAllCards());
    controlsDiv.querySelector('.deselect-all-btn').addEventListener('click', () => this.deselectAllCards());
  }

  createCardSelectionItem(cardData) {
    const cardDiv = document.createElement('div');
    cardDiv.className = 'quick-card-selection-item';
    cardDiv.style.cssText = `
      display: flex;
      align-items: center;
      gap: 1rem;
      padding: 1rem;
      border: 1px solid #e0e0e0;
      border-radius: 8px;
      margin-bottom: 0.8rem;
      background: white;
      transition: all 0.2s;
    `;

    // Get available saved views for this card
    const savedViews = this.loadSavedViews();
    const viewOptions = Object.keys(savedViews).map(name => 
      `<option value="${name}">${name}</option>`
    ).join('');

    cardDiv.innerHTML = `
      <div class="card-checkbox" style="flex-shrink: 0;">
        <label class="checkbox-label" style="
          display: flex;
          align-items: center;
          cursor: pointer;
          font-size: 0.9rem;
        ">
          <input type="checkbox" class="card-checkbox-input" data-card-id="${cardData.id}" style="
            margin-right: 0.5rem;
            transform: scale(1.1;
          ">
          <span class="checkmark" style="
            height: 18px;
            width: 18px;
            background-color: #eee;
            border-radius: 3px;
            display: inline-block;
            position: relative;
            margin-right: 0.5rem;
          "></span>
        </label>
      </div>
      
      <div class="card-info" style="flex: 1; min-width: 0;">
        <div class="card-name" style="
          font-weight: 600;
          color: #1a2332;
          margin-bottom: 0.3rem;
          font-size: 1rem;
        ">${cardData.name}</div>
        <div class="card-count" style="
          color: #47B2E5;
          font-size: 0.9rem;
          font-weight: 500;
        ">${cardData.count.toLocaleString()} records</div>
      </div>
      
      <div class="card-options" style="
        display: flex;
        gap: 0.8rem;
        align-items: center;
        flex-wrap: wrap;
      ">
                    <div class="view-type-selector" style="display: flex; align-items: center; gap: 0.5rem;">
              <label style="font-size: 0.8rem; color: #666; white-space: nowrap;">Display:</label>
              <select class="filter-select card-view-type" data-card-id="${cardData.id}" style="
                padding: 0.4rem 0.6rem;
                border: 1px solid #ddd;
                border-radius: 4px;
                font-size: 0.8rem;
                min-width: 100px;
              ">
                <option value="summary" ${this.getCardConfig(cardData.id).viewType === 'summary' ? 'selected' : ''}>Summary</option>
                <option value="table" ${this.getCardConfig(cardData.id).viewType === 'table' ? 'selected' : ''}>Table</option>
                <option value="both" ${this.getCardConfig(cardData.id).viewType === 'both' ? 'selected' : ''}>Both</option>
                <option value="charts" ${this.getCardConfig(cardData.id).viewType === 'charts' ? 'selected' : ''}>Charts</option>
              </select>
            </div>
        
        <div class="saved-view-selector" style="display: flex; align-items: center; gap: 0.5rem;">
          <label style="font-size: 0.8rem; color: #666; white-space: nowrap;">Columns:</label>
          <select class="filter-select card-saved-view" data-card-id="${cardData.id}" style="
            padding: 0.4rem 0.6rem;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 0.8rem;
            min-width: 120px;
          ">
            <option value="current" ${this.getCardConfig(cardData.id).savedView === 'current' ? 'selected' : ''}>All columns</option>
            ${viewOptions}
          </select>
        </div>
      </div>
    `;

    // Bind events
    const checkbox = cardDiv.querySelector('.card-checkbox-input');
    const viewTypeSelect = cardDiv.querySelector('.card-view-type');
    const savedViewSelect = cardDiv.querySelector('.card-saved-view');

    checkbox.addEventListener('change', (e) => {
      this.toggleCardSelection(cardData.id, cardDiv);
    });

    viewTypeSelect.addEventListener('change', () => {
      console.log(`View type changed for card ${cardData.id} to: ${viewTypeSelect.value}`);
      this.updateCardConfig(cardData.id, 'viewType', viewTypeSelect.value);
      
      // Show immediate feedback
      const cardDiv = viewTypeSelect.closest('.card-selection-item');
      if (cardDiv) {
        cardDiv.style.borderColor = '#47B2E5';
        cardDiv.style.backgroundColor = 'rgba(71, 178, 229, 0.1)';
        setTimeout(() => {
          cardDiv.style.backgroundColor = 'rgba(71, 178, 229, 0.05)';
        }, 200);
      }
      
      // Force immediate preview update
      this.renderSummaryPreview();
    });

    savedViewSelect.addEventListener('change', () => {
      this.updateCardConfig(cardData.id, 'savedView', savedViewSelect.value);
      this.renderSummaryPreview();
    });

    // Set initial state
    if (this.selectedCards.has(cardData.id)) {
      checkbox.checked = true;
      cardDiv.style.borderColor = '#47B2E5';
      cardDiv.style.backgroundColor = 'rgba(71, 178, 229, 0.05)';
    }

    return cardDiv;
  }

  toggleCardSelection(cardId, element) {
    if (this.selectedCards.has(cardId)) {
      this.selectedCards.delete(cardId);
      element.style.borderColor = '#e0e0e0';
      element.style.backgroundColor = 'white';
      const checkbox = element.querySelector('.card-checkbox-input');
      if (checkbox) checkbox.checked = false;
    } else {
      this.selectedCards.add(cardId);
      element.style.borderColor = '#47B2E5';
      element.style.backgroundColor = 'rgba(71, 178, 229, 0.05)';
      const checkbox = element.querySelector('.card-checkbox-input');
      if (checkbox) checkbox.checked = true;
    }
    
    this.renderSummaryPreview();
  }

  updateCardConfig(cardId, key, value) {
    if (!this.cardConfigs[cardId]) {
      this.cardConfigs[cardId] = {
        viewType: 'both',
        savedView: 'current'
      };
    }
    
    this.cardConfigs[cardId][key] = value;
    console.log(`Updated card config for ${cardId}:`, this.cardConfigs[cardId]);
  }

  getCardConfig(cardId) {
    return this.cardConfigs[cardId] || {
      viewType: 'both',
      savedView: 'current'
    };
  }

  selectAllCards() {
    this.selectedCards.clear();
    Object.keys(this.quickCardsData).forEach(cardId => {
      if (cardId !== 'no_cards') {
        this.selectedCards.add(cardId);
      }
    });
    this.renderCardsSelection();
    this.renderSummaryPreview();
    this.showDuplicateSummary();
  }

  deselectAllCards() {
    this.selectedCards.clear();
    this.renderCardsSelection();
    this.renderSummaryPreview();
    this.showDuplicateSummary();
  }



  renderSummaryPreview() {
    // Debounce preview updates to prevent blocking
    if (this.previewUpdateTimeout) {
      clearTimeout(this.previewUpdateTimeout);
    }
    
    console.log('renderSummaryPreview called - scheduling update');
    
    this.previewUpdateTimeout = setTimeout(() => {
      console.log('Executing preview update after debounce');
      this._renderSummaryPreviewInternal();
    }, 100); // Reduced from 300ms to 100ms for faster response
  }

  _renderSummaryPreviewInternal() {
    const preview = document.getElementById('summaryPreview');
    if (!preview || this.isGeneratingPreview) {
      console.log('Preview update skipped:', { preview: !!preview, isGenerating: this.isGeneratingPreview });
      return;
    }

    this.isGeneratingPreview = true;
    console.log('Starting preview generation...');
    
    try {
      const selectedCardsData = Object.values(this.quickCardsData)
        .filter(card => this.selectedCards.has(card.id));

      console.log('Rendering preview with', selectedCardsData.length, 'selected cards');
      
      // Log card configurations for debugging
      selectedCardsData.forEach(card => {
        const config = this.getCardConfig(card.id);
        console.log(`Card "${card.name}" config:`, config);
      });
      
      // Show loading indicator for large datasets
      if (selectedCardsData.length > 5) {
        preview.innerHTML = '<div style="text-align:center; padding:2rem; color:#666;">Generating preview...</div>';
      }
      
      // Use requestAnimationFrame to prevent blocking
      requestAnimationFrame(() => {
        const html = this.generatePreviewHTML(selectedCardsData);
        preview.innerHTML = html;
        this.isGeneratingPreview = false;
        console.log('Preview generation completed');
      });
    } catch (error) {
      console.error('Error generating preview:', error);
      preview.innerHTML = '<div style="text-align:center; padding:2rem; color:#f44336;">Error generating preview</div>';
      this.isGeneratingPreview = false;
    }
  }

  generatePreviewHTML(selectedCards) {
    const config = this.getReportConfig();
    
    // Always generate separate tables for each selected card with their individual configurations
    if (selectedCards.length > 0) {
      return this.generateSeparateTablesForCardsPreview(selectedCards, config);
    }
    
    // If no cards selected, show empty state
    return '<div style="color:#888; text-align:center; padding:2rem;">No cards selected. Please select quick cards to see the preview.</div>';
  }

  generateHtmlSummaryBlock(selectedCards) {
    const dateStr = new Date().toLocaleString();
    const selectedCount = selectedCards.length;
    const config = this.getReportConfig();
    
    // Always use the new preview logic with individual card configurations
    const content = this.generatePreviewHTML(selectedCards);
    
    // Generate summary info section only if enabled
    let summaryInfo = '';
    if (config.includeSummaryInfo) {
      summaryInfo = `
        <div style="display:flex; justify-content:space-between; color:#15364A; font-size:10px; margin-bottom:2em;">
          <div><b>Generated:</b> ${dateStr}</div>
          <div><b>Selected:</b> ${selectedCount}</div>
        </div>
      `;
    }
    
    return `
      <div style="background:#fff; border-radius:8px; box-shadow:0 2px 8px rgba(0,0,0,0.1); padding:2.5rem; margin:0 auto; max-width:600px;">
        <div style="text-align:center;">
          <h1 style="font-size:1.4rem; color:#15364A; margin-bottom:0.2em; margin-top:0;">Operations Summary Report</h1>
          <hr style="border:none; border-top:1px solid #e3f2fd; margin:1.2em 0 1.5em 0;">
        </div>
        ${summaryInfo}
        
        ${content}
        
        <div style="color:#888; font-size:9px; text-align:center; margin-top:2em;">Generated by The Bridge Operations Hub</div>
      </div>
    `;
  }

  generateSummarySection(selectedCards, config) {
    const tableRows = selectedCards.length > 0 ? selectedCards.map(card => `
      <tr>
        <td style="padding:6px 10px; border:1px solid #ddd; color:#222; font-size:9px;">${card.name}</td>
        <td style="padding:6px 10px; border:1px solid #ddd; color:#1976d2; text-align:right; font-size:9px; font-weight:600;">${card.count}</td>
      </tr>
    `).join('') : `<tr><td colspan="2" style="padding:8px 10px; color:#888; text-align:center; font-size:9px;">No cards selected.</td></tr>`;
    
    return `
      <div style="margin-bottom:2.5rem;">
        <h2 style="font-size:1.1rem; color:#15364A; margin-bottom:1.5rem; border-bottom:1px solid #e3f2fd; padding-bottom:0.5rem;">Quick Cards Summary</h2>
        <table style="width:100%; border-collapse:collapse; margin:1.5rem 0; font-size:9px; max-width:350px;">
          <thead>
            <tr style="background:#f5f5f5;">
              <th style="padding:6px 10px; color:#1976d2; text-align:left; border:1px solid #ddd; font-weight:600;">Card</th>
              <th style="padding:6px 10px; color:#1976d2; text-align:right; border:1px solid #ddd; font-weight:600;">Records</th>
            </tr>
          </thead>
          <tbody>
            ${tableRows}
          </tbody>
        </table>
      </div>
    `;
  }

  generateChartsSection() {
    return `
      <div style="margin-bottom:2rem;">
        <h2 style="font-size:1.5rem; color:#15364A; margin-bottom:1rem; border-bottom:2px solid #e3f2fd; padding-bottom:0.5rem;">Charts & Analytics</h2>
        <div style="color:#888; text-align:center; padding:2rem; background:#f8f9fa; border-radius:8px;">
          Charts functionality will be available in future updates
        </div>
      </div>
    `;
  }

  generateTableDataSection(tableData, config, customTitle = 'Table Data', isPreview = false) {
    console.log(`🔍 generateTableDataSection called for "${customTitle}":`, {
      dataLength: tableData?.length || 0,
      config: config,
      tableDataType: typeof tableData,
      isArray: Array.isArray(tableData),
      firstRow: tableData?.[0] || 'N/A'
    });
    
    // REPLICADO DEL REPOSITORIO THE BRIDGE: Manejo robusto de datos
    if (!tableData || tableData.length === 0) {
      console.warn(`⚠️ No table data available for "${customTitle}" - tableData:`, tableData);
      return `
        <div style="margin-bottom:2rem;">
          <h2 style="font-size:1.5rem; color:#15364A; margin-bottom:1rem; border-bottom:2px solid #e3f2fd; padding-bottom:0.5rem;">${customTitle}</h2>
          <div style="color:#888; text-align:center; padding:2rem; background:#f8f9fa; border-radius:8px;">
            <div style="margin-bottom:1rem;">No table data available</div>
            <div style="font-size:0.9em; color:#666;">
              Debug info: tableData length = ${tableData?.length || 'null'}, type = ${typeof tableData}
            </div>
          </div>
        </div>
      `;
    }

    // Get column headers from the first row
    let headers = Object.keys(tableData[0]);
    
    // Apply saved view columns if specified
    if (config.savedView && config.savedView !== 'current') {
      const savedViews = this.loadSavedViews();
      const selectedView = savedViews[config.savedView];
      
      if (selectedView && selectedView.columns) {
        console.log(`Applying saved view columns "${config.savedView}" to table:`, selectedView.columns);
        // Filter headers to only include columns from the saved view
        headers = headers.filter(header => selectedView.columns.includes(header));
        console.log(`Filtered headers for saved view:`, headers);
      }
    }
    
    // Limit rows only in preview, not in final report
    const totalRows = tableData.length;
    let displayData, shownRows;
    
    if (isPreview) {
      // Limit rows to prevent collapse - show max 50 rows in preview only
      const maxRows = 50;
      displayData = tableData.slice(0, maxRows);
      shownRows = displayData.length;
    } else {
      // Show all data in final report
      displayData = tableData;
      shownRows = totalRows;
    }
    
    // Generate table rows with compact email-friendly styling - NO WRAP
    const dataRows = displayData.map(row => {
      const cells = headers.map(header => {
        const value = row[header] || '';
        // Truncate very long values to prevent layout issues and make it email-friendly
        const displayValue = value.length > 25 ? value.substring(0, 25) + '...' : value;
        return `<td style="padding:4px 6px; border:1px solid #ddd; color:#333; font-size:9px; max-width:60px; overflow:hidden; text-overflow:ellipsis; text-align:left; white-space:nowrap;">${displayValue}</td>`;
      }).join('');
      return `<tr>${cells}</tr>`;
    }).join('');

    // Generate header row with compact styling - NO WRAP
    const headerRow = headers.map(header => {
      // Truncate header text if too long
      const displayHeader = header.length > 15 ? header.substring(0, 15) + '...' : header;
      return `<th style="padding:6px 8px; color:#1976d2; text-align:left; border:1px solid #ddd; background:#f5f5f5; font-weight:600; font-size:9px; white-space:nowrap; max-width:60px; overflow:hidden; text-overflow:ellipsis;">${displayHeader}</th>`;
    }).join('');

    const rowsInfo = `Showing ${shownRows} of ${totalRows} records`;

    // Add performance warning for large datasets
    let performanceWarning = '';
    if (totalRows > 1000) {
      performanceWarning = `<div style="margin-bottom:0.5rem; color:#ff9800; font-size:0.9em; font-style:italic;">⚠ Large dataset detected (${totalRows.toLocaleString()} records). Consider using filters or reducing the scope.</div>`;
    }

    // Note: Duplicate removal info is now handled in generateSeparateTablesForCards
    // The data passed to this method is already processed for duplicates

    return `
      <div style="margin-bottom:2.5rem;">
        <h2 style="font-size:1.1rem; color:#15364A; margin-bottom:1.5rem; border-bottom:1px solid #e3f2fd; padding-bottom:0.5rem;">${customTitle}</h2>
        ${performanceWarning}
        <div style="margin-bottom:1.5rem; color:#666; font-size:9px; font-style:italic;">${rowsInfo}</div>
        <div style="max-width:100%; overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; font-size:9px; background:#fff; text-align:left; table-layout:fixed; max-width:500px;">
            <thead>
              <tr>${headerRow}</tr>
            </thead>
            <tbody>
              ${dataRows}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  generateCleanTextContent() {
    const selectedCardsData = Object.values(this.quickCardsData)
      .filter(card => this.selectedCards.has(card.id));
    // Usar exactamente el mismo formato que la preview
    return this.generateHtmlSummaryBlock(selectedCardsData);
  }



  async copyToClipboard() {
    try {
      const selectedCards = Object.values(this.quickCardsData).filter(card => this.selectedCards.has(card.id));
      const htmlContent = this.generateCleanTextContent();
      if (navigator.clipboard && window.ClipboardItem) {
        const blob = new Blob([htmlContent], { type: 'text/html' });
        await navigator.clipboard.write([new window.ClipboardItem({ 'text/html': blob })]);
        this.showNotification('Summary table copied to clipboard!', 'success');
      } else {
        // Fallback: copy as plain text (still works in most email clients)
        const textArea = document.createElement('textarea');
        textArea.value = htmlContent;
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
        this.showNotification('Summary table copied as text!', 'success');
      }
    } catch (error) {
      this.showNotification('Failed to copy summary table: ' + error.message, 'error');
    }
  }

  async exportToPdf() {
    try {
      if (!window.jspdf || !window.jspdf.jsPDF) {
        this.showNotification('PDF export not available - jsPDF library not loaded', 'error');
        return;
      }
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF();
      const dateStr = new Date().toLocaleString();
      const selectedCards = Object.values(this.quickCardsData).filter(card => this.selectedCards.has(card.id));
      const selectedCount = selectedCards.length;
      doc.setFontSize(20);
      doc.setTextColor(26, 35, 50);
      doc.text('Operations Summary Report', 105, 20, { align: 'center' });
      doc.setFontSize(12);
      doc.setTextColor(71, 178, 229);
      doc.setDrawColor(227, 242, 253);
      doc.line(20, 36, 190, 36); // línea divisoria
      doc.setFontSize(10);
      doc.setTextColor(21, 54, 74);
      doc.text(`Generated: ${dateStr}`, 20, 44);
      doc.text(`Selected: ${selectedCount}`, 160, 44, { align: 'right' });
      const rows = selectedCards.length > 0 ? selectedCards.map(card => [card.name, card.count]) : [['No cards selected.', '']];
      if (doc.autoTable) {
        doc.autoTable({
          head: [['Card', 'Records']],
          body: rows,
          startY: 50,
          headStyles: { fillColor: [227, 242, 253], textColor: [25, 118, 210], fontStyle: 'bold' },
          bodyStyles: { textColor: [33, 33, 33] },
          styles: { fontSize: 12, cellPadding: 4, halign: 'right' },
          columnStyles: { 0: { halign: 'left' }, 1: { halign: 'right' } }
        });
        doc.setFontSize(10);
        doc.setTextColor(136, 136, 136);
        doc.text('Generated by The Bridge Operations Hub', 105, doc.lastAutoTable.finalY + 10, { align: 'center' });
      } else {
        // Fallback simple
        let y = 70;
        doc.setFontSize(12);
        doc.setTextColor(25, 118, 210);
        doc.text('Card', 20, y);
        doc.text('Records', 120, y);
        y += 10;
        doc.setTextColor(33, 33, 33);
        rows.forEach(row => {
          doc.text(row[0], 20, y);
          doc.text(String(row[1]), 120, y, { align: 'right' });
          y += 10;
        });
        doc.setFontSize(10);
        doc.setTextColor(136, 136, 136);
        doc.text('Generated by The Bridge Operations Hub', 105, y + 10, { align: 'center' });
      }
      const filename = `operations-summary-${new Date().toISOString().split('T')[0]}.pdf`;
      doc.save(filename);
      this.showNotification('PDF exported successfully!', 'success');
    } catch (error) {
      this.showNotification('Failed to export PDF: ' + error.message, 'error');
    }
  }

  exportToExcel() {
    try {
      // Obtener el mismo contenido que usa copyToClipboard
      const selectedCards = Object.values(this.quickCardsData).filter(card => this.selectedCards.has(card.id));
      if (selectedCards.length === 0) {
        this.showNotification('No cards selected for export.', 'warning');
        return;
      }

      // Crear el workbook
      const wb = window.XLSX.utils.book_new();
      const dateStr = new Date().toLocaleString();
      const selectedCount = selectedCards.length;

      // Hoja 1: Resumen de Quick Cards
      const summaryData = selectedCards.map(card => ({
        'Card': card.name,
        'Records': card.count
      }));
      const ws1 = window.XLSX.utils.json_to_sheet(summaryData);
      window.XLSX.utils.book_append_sheet(wb, ws1, 'Quick Cards Summary');

      // Hoja 2: Datos de tabla (si hay configuración)
      const config = this.getReportConfig();
      if (config.separateTablesPerCard) {
        // Crear una hoja por cada card seleccionada
        selectedCards.forEach(card => {
          const cardConfig = this.getCardConfig(card.id);
          const cardData = this.getDataForQuickCardWithConfig(card, cardConfig);
          if (cardData && cardData.length > 0) {
            const ws = window.XLSX.utils.json_to_sheet(cardData);
            window.XLSX.utils.book_append_sheet(wb, ws, card.name.substring(0, 31)); // Excel limita nombres de hoja a 31 chars
          }
        });
      } else {
        // Una sola hoja con todos los datos
        const tableData = this.getDataForCurrentConfiguration(config);
        if (tableData && tableData.length > 0) {
          const ws2 = window.XLSX.utils.json_to_sheet(tableData);
          window.XLSX.utils.book_append_sheet(wb, ws2, 'Table Data');
        }
      }

      // Hoja 3: Información del reporte
      const reportInfo = [
        { 'Field': 'Report Title', 'Value': 'Operations Summary Report' },
        { 'Field': 'Generated', 'Value': dateStr },
        { 'Field': 'Selected Cards', 'Value': selectedCount },
        { 'Field': 'Table View', 'Value': config.tableView || 'Current' },
        { 'Field': 'Separate Tables Per Card', 'Value': config.separateTablesPerCard ? 'Yes' : 'No' },
        { 'Field': 'Remove Duplicates', 'Value': config.removeDuplicates ? 'Yes' : 'No' },
        { 'Field': 'Generated By', 'Value': 'The Bridge Operations Hub' }
      ];
      const ws3 = window.XLSX.utils.json_to_sheet(reportInfo);
      window.XLSX.utils.book_append_sheet(wb, ws3, 'Report Info');

      // Descargar el archivo
      const filename = `operations-summary-${new Date().toISOString().split('T')[0]}.xlsx`;
      window.XLSX.writeFile(wb, filename);
      this.showNotification('Excel exported successfully!', 'success');
    } catch (error) {
      this.showNotification('Failed to export Excel: ' + error.message, 'error');
    }
  }

  showNotification(message, type = 'info') {
    if (typeof window.showUnifiedNotification === 'function') {
      window.showUnifiedNotification(message, type);
    } else {
      // Fallback simple
      console.log(`Notification [${type}]:`, message);
    }
  }

  getDataForCurrentConfiguration(config) {
    // Create a hash of the configuration for caching
    const configHash = JSON.stringify({
      selectedCards: Array.from(this.selectedCards).sort(),
      cardConfigs: this.cardConfigs
    });
    
    // Check cache first
    if (this.cachedData.has(configHash)) {
      return this.cachedData.get(configHash);
    }
    
    let result;
    
    // If separate tables per card is enabled, get data for all selected cards
    if (config.separateTablesPerCard) {
      const selectedCards = Object.values(this.quickCardsData).filter(card => this.selectedCards.has(card.id));
      let allData = [];
      
      // Process ALL selected cards (no limit)
      selectedCards.forEach(card => {
        const cardConfig = this.getCardConfig(card.id);
        const cardData = this.getDataForQuickCardWithConfig(card, cardConfig);
        if (cardData && cardData.length > 0) {
          allData = allData.concat(cardData);
        }
      });
      
      result = allData;
    }
    // If using a saved view, apply that view to the data
    else if (config.tableView && config.tableView !== 'current') {
      const savedViews = this.loadSavedViews();
      const selectedView = savedViews[config.tableView];
      
      if (selectedView && selectedView.columns) {
        console.log(`Applying saved view "${config.tableView}" to current configuration:`, selectedView);
        let data = this.getFilteredData(); // Start with filtered data (respects active filters)

        // Filter data to only include the columns specified in the saved view
        data = data.map(row => {
          const filteredRow = {};
          selectedView.columns.forEach(col => {
            if (row.hasOwnProperty(col)) {
              filteredRow[col] = row[col];
            }
          });
          return filteredRow;
        });
        console.log(`Data after applying saved view "${config.tableView}":`, data.length, 'records');
        result = data;
      } else {
        console.warn(`Saved view "${config.tableView}" not found or invalid.`);
        result = this.getFilteredData(); // Use filtered data, not original
      }
    }
    // Default: use filtered data (not original)
    else {
      result = this.getFilteredData(); // Usar los datos filtrados actuales
    }
    
    // Cache the result (limit cache size to prevent memory issues)
    if (this.cachedData.size > 10) {
      const firstKey = this.cachedData.keys().next().value;
      this.cachedData.delete(firstKey);
    }
    
    this.cachedData.set(configHash, result);
    return result;
  }

  getDataForQuickCardWithConfig(card, cardConfig) {
    // --- NUEVA LÓGICA: Forzar preview idéntica a la tabla principal ---
    // 1. Buscar el filtro rápido correspondiente
    const quickFilters = this.loadQuickFilters();
    let filterObj = quickFilters[card.name];
    let quickFilterName = card.name;
    if (!filterObj || !filterObj.filterValues) {
      // Buscar por case-insensitive o substring
      const cardNameLower = card.name.toLowerCase();
      const matchingFilter = Object.keys(quickFilters).find(filterName => filterName.toLowerCase() === cardNameLower);
      if (matchingFilter) {
        filterObj = quickFilters[matchingFilter];
        quickFilterName = matchingFilter;
      } else {
        const partialMatch = Object.keys(quickFilters).find(filterName => filterName.toLowerCase().includes(cardNameLower) || cardNameLower.includes(filterName.toLowerCase()));
        if (partialMatch) {
          filterObj = quickFilters[partialMatch];
          quickFilterName = partialMatch;
        }
      }
    }
    // 2. Si hay quick filter, actívalo temporalmente y usa getFilteredData()
    if (filterObj && filterObj.filterValues) {
      // Guardar filtros originales
      const originalActive = { ...getModuleActiveFilters() };
      const originalValues = { ...getModuleFilterValues() };
      // Activar solo el filtro rápido de la card
      const newActive = {};
      const newValues = {};
      Object.entries(filterObj.filterValues).forEach(([key, value]) => {
        if (key.endsWith('_start') || key.endsWith('_end') || key.endsWith('_empty')) {
          const base = key.replace(/_(start|end|empty)$/, '');
          newActive[base] = 'date';
        } else if (Array.isArray(value)) {
          newActive[key] = 'categorical';
        } else {
          newActive[key] = 'text';
        }
        newValues[key] = value;
      });
      setModuleActiveFilters(newActive);
      setModuleFilterValues(newValues);
      // Obtener los datos filtrados igual que la tabla principal
      const filtered = getFilteredData();
      // Restaurar filtros originales
      setModuleActiveFilters(originalActive);
      setModuleFilterValues(originalValues);
      return filtered;
    }
    // 3. Si no hay quick filter, usar los datos filtrados actuales
    return getFilteredData();
  }

  getOriginalData() {
    console.log('Getting original data...');
    
    // Use the imported getOriginalData function from store
    try {
      const data = getOriginalData();
      console.log(`Got original data from store: ${data.length} records`);
      return data;
    } catch (error) {
      console.error('❌ Error getting original data from store:', error);
    
    // Fallback: try to get data from table
    const table = document.querySelector('#dataTable tbody');
    if (table) {
      const rows = Array.from(table.querySelectorAll('tr'));
      if (rows.length > 0) {
        const headers = Array.from(table.querySelector('thead tr').querySelectorAll('th')).map(th => th.textContent.trim());
        const data = rows.map(row => {
          const cells = Array.from(row.querySelectorAll('td'));
          const rowData = {};
          headers.forEach((header, index) => {
            rowData[header] = cells[index] ? cells[index].textContent.trim() : '';
          });
          return rowData;
        });
        console.log(`Got original data from table: ${data.length} records`);
        return data;
      }
    }
    
    // Last fallback: try to get from global variable
    if (window.tableData && Array.isArray(window.tableData)) {
      console.log(`Got original data from window.tableData: ${window.tableData.length} records`);
      return window.tableData;
    }
    
    // Try to get from CSV data if available
    if (window.csvData && Array.isArray(window.csvData)) {
      console.log(`Got original data from window.csvData: ${window.csvData.length} records`);
      return window.csvData;
    }
    
    console.warn('Could not get original data, using empty array');
    return [];
    }
  }

  getFilteredData() {
    console.log('🔍 getFilteredData called - getting filtered data (respecting active filters)...');
    
    // Use the imported getFilteredData function from FilterManager
    try {
      const data = getFilteredData();
      console.log(`✅ Got filtered data from FilterManager: ${data.length} records`);
      return data;
    } catch (error) {
      console.error('❌ Error getting filtered data:', error);
    
      // Fall back to original data if there's an error
      console.log('⚠️ Falling back to original data due to error');
    const originalData = this.getOriginalData();
    console.log(`📊 Original data fallback: ${originalData.length} records`);
    return originalData;
    }
  }

  loadQuickFilters() {
    try {
      // Use the imported loadQuickFilters function from FilterManager
      const filters = loadQuickFilters();
      console.log(`📋 Loaded ${Object.keys(filters).length} quick filters from FilterManager:`, Object.keys(filters));
      return filters;
    } catch (error) {
      console.error('❌ Error loading quick filters from FilterManager:', error);
      
      // Fallback: try to load from localStorage
      try {
        const saved = localStorage.getItem('quickFilters');
        if (saved) {
          const filters = JSON.parse(saved);
          console.log(`📋 Loaded ${Object.keys(filters).length} quick filters from localStorage:`, Object.keys(filters));
          return filters;
        }
      } catch (localStorageError) {
        console.error('❌ Error loading quick filters from localStorage:', localStorageError);
      }
      
      console.log('📋 No quick filters found, returning empty object');
      return {};
    }
  }



  selectCommonFields() {
    const checkboxes = document.querySelectorAll('.duplicate-column-checkbox');
    checkboxes.forEach(checkbox => {
      const fieldName = checkbox.value;
      checkbox.checked = this.isCommonField(fieldName);
    });
    this.updateDuplicateColumns();
  }

  filterDuplicateColumns() {
    const searchInput = document.getElementById('duplicateColumnsSearch');
    const columnItems = document.querySelectorAll('#duplicateColumnsList .column-checkbox-item');
    
    console.log('filterDuplicateColumns called');
    console.log('Search input found:', !!searchInput);
    console.log('Column items found:', columnItems.length);
    
    if (!searchInput) {
      console.log('Search input not found');
      return;
    }
    
    const searchTerm = searchInput.value.toLowerCase();
    console.log('Filtering columns with term:', searchTerm);
    
    if (searchTerm === '') {
      // Show all items when search is empty
      columnItems.forEach(item => {
        item.style.display = 'block';
      });
      console.log('Search term empty, showing all columns');
      return;
    }
    
    let visibleCount = 0;
    columnItems.forEach((item, index) => {
      const fieldName = item.getAttribute('data-field');
      
      if (!fieldName) {
        console.log(`Warning: No data-field attribute found for item ${index}`);
        item.style.display = 'none';
        return;
      }
      
      const matches = fieldName.toLowerCase().includes(searchTerm);
      item.style.display = matches ? 'block' : 'none';
      if (matches) {
        visibleCount++;
        console.log(`✓ Column "${fieldName}" matches search term`);
      }
    });
    
    console.log(`Found ${visibleCount} matching columns out of ${columnItems.length} total`);
    
    // Show message if no matches found
    const noResultsMsg = document.getElementById('noDuplicateColumnsMsg');
    if (visibleCount === 0 && searchTerm !== '') {
      if (!noResultsMsg) {
        const msg = document.createElement('div');
        msg.id = 'noDuplicateColumnsMsg';
        msg.style.cssText = 'color:#888; text-align:center; padding:1rem; font-style:italic;';
        msg.textContent = `No columns found matching "${searchTerm}"`;
        document.getElementById('duplicateColumnsList').appendChild(msg);
      }
    } else if (noResultsMsg) {
      noResultsMsg.remove();
    }
  }

  getReportConfig() {
    const includeTableDataCheckbox = document.getElementById('includeTableDataCheckbox');
    const includeTechnicalInfoCheckbox = document.getElementById('includeTechnicalInfoCheckbox');
    const viewTypeSelect = document.getElementById('reportViewTypeSelect');
    const removeDuplicatesCheckbox = document.getElementById('removeDuplicatesCheckbox');
    const separateTablesPerCardCheckbox = document.getElementById('separateTablesPerCardCheckbox');

    const includeTechnicalInfo = includeTechnicalInfoCheckbox?.checked !== false;

    return {
      includeTableData: includeTableDataCheckbox?.checked !== false,
      includeCardConfig: includeTechnicalInfo,
      includeSummaryInfo: includeTechnicalInfo,
      includeDuplicateSummary: includeTechnicalInfo,
      viewType: viewTypeSelect?.value || 'both',
      removeDuplicates: removeDuplicatesCheckbox?.checked || false,
      duplicateColumns: this.reportConfig.duplicateColumns || [],
      tableDataLimit: 'all', // Always show all data in Ops Hub Summary
      separateTablesPerCard: separateTablesPerCardCheckbox?.checked || false
    };
  }

  removeDuplicatesFromData(data, config = null) {
    // Get config from parameter or current report config
    const reportConfig = config || this.getReportConfig();
    
    console.log('removeDuplicatesFromData called with:', {
      dataLength: data.length,
      removeDuplicates: reportConfig.removeDuplicates,
      duplicateColumns: reportConfig.duplicateColumns
    });
    
    if (!reportConfig.removeDuplicates || reportConfig.duplicateColumns.length === 0) {
      console.log('Duplicate removal not enabled or no columns selected, returning original data');
      return data;
    }

    const seen = new Set();
    const filteredData = data.filter(row => {
      const key = reportConfig.duplicateColumns.map(col => row[col] || '').join('|');
      if (seen.has(key)) {
        return false;
      }
      seen.add(key);
      return true;
    });
    
    console.log(`Duplicate removal applied: ${data.length} -> ${filteredData.length} records (removed ${data.length - filteredData.length})`);
    return filteredData;
  }

  loadSavedViews() {
    try {
      const saved = localStorage.getItem('tableViews');
      return saved ? JSON.parse(saved) : {};
    } catch (e) {
      console.error('Error loading saved views:', e);
      return {};
    }
  }

  generateSeparateTablesForCardsPreview(selectedCards, config) {
    if (!selectedCards || selectedCards.length === 0) {
      return '<div style="color:#888; text-align:center; padding:2rem;">No cards selected for separate tables.</div>';
    }

    let tablesHtml = '';
    let totalRecords = 0;
    let totalOriginalRecords = 0;
    let totalRemovedDuplicates = 0;
    
    // Generate summary section first
    const summarySection = this.generateSummarySection(selectedCards, config);
    tablesHtml += summarySection;
    
    // Process each selected card with its individual configuration
    selectedCards.forEach((card, index) => {
      const cardConfig = this.getCardConfig(card.id);
      console.log(`🔍 Processing card "${card.name}" (ID: ${card.id}) for PREVIEW:`, {
        cardConfig: cardConfig,
        cardCount: card.count,
        cardActive: card.active
      });
      
      const cardData = this.getDataForQuickCardWithConfig(card, cardConfig);
      console.log(`📊 Card "${card.name}" data result for PREVIEW:`, {
        dataLength: cardData?.length || 0,
        dataType: typeof cardData,
        isArray: Array.isArray(cardData),
        firstRow: cardData?.[0] || 'N/A',
        cardCount: card.count,
        expectedRecords: card.count,
        hasQuickFilter: !!(cardConfig && cardConfig.quickFilter && cardConfig.quickFilter !== 'none'),
        quickFilterName: cardConfig?.quickFilter || 'none'
      });
      
      // Debug: Show sample of filtered data
      if (cardData && cardData.length > 0) {
        console.log(`📋 Sample filtered data for "${card.name}" PREVIEW:`, cardData.slice(0, 2));
      }
      
      // Debug: Check if there's a mismatch between card count and actual data
      if (card.count > 0 && (!cardData || cardData.length === 0)) {
        console.error(`🚨 CRITICAL: Card "${card.name}" shows ${card.count} records but getDataForQuickCardWithConfig returned ${cardData?.length || 0} records`);
        console.error(`🚨 This is why "No data available for this card" is showing in the preview`);
      }
      
      if (cardData && cardData.length > 0) {
        totalOriginalRecords += cardData.length;
        
        // Apply duplicate removal if enabled
        const originalCount = cardData.length;
        console.log(`Processing card "${card.name}" for PREVIEW: ${originalCount} original records`);
        
        const processedData = this.removeDuplicatesFromData(cardData, config);
        const removedCount = originalCount - processedData.length;
        totalRemovedDuplicates += removedCount;
        
        console.log(`Card "${card.name}" after duplicate removal for PREVIEW: ${processedData.length} records (removed ${removedCount})`);
        
        // Use card's individual view type setting
        const cardViewConfig = {
          ...config,
          viewType: cardConfig.viewType || 'both'
        };
        
        // Generate appropriate content based on card's view type
        let cardContent = '';
        
        if (cardConfig.viewType === 'summary') {
          // Only show summary info for this card
          let duplicateInfo = '';
          if (config.removeDuplicates && config.duplicateColumns.length > 0 && removedCount > 0) {
            duplicateInfo = `<div style="color:#1976d2; font-size:10px; margin-top:0.2rem;">Removed ${removedCount} duplicates</div>`;
          }
          
          cardContent = `
            <div style="margin-bottom:2rem; padding:0.8rem; background:#f8f9fa; border-radius:6px; border-left:3px solid #47B2E5;">
              <h4 style="margin:0 0 0.3rem 0; color:#1a2332; font-size:12px;">${card.name}</h4>
              <div style="color:#47B2E5; font-weight:600; font-size:11px;">${processedData.length.toLocaleString()} records</div>
              <div style="color:#666; font-size:10px; margin-top:0.2rem;">View type: Summary only</div>
              ${duplicateInfo}
            </div>
          `;
        } else if (cardConfig.viewType === 'charts') {
          // Show charts placeholder for this card
          cardContent = `
            <div style="margin-bottom:2rem;">
              <h3 style="font-size:1.1rem; color:#15364A; margin-bottom:0.5rem; border-bottom:1px solid #e3f2fd; padding-bottom:0.3rem;">
                ${card.name} - Charts
              </h3>
              <div style="color:#888; text-align:center; padding:1rem; background:#f8f9fa; border-radius:6px; font-size:10px;">
                Charts functionality will be available in future updates
              </div>
            </div>
          `;
        } else {
          // Show table data (for 'table' or 'both' view types) - ONLY if includeTableData is enabled
          if (config.includeTableData) {
            const tableConfig = {
              ...cardViewConfig,
              savedView: cardConfig.savedView // Pass the saved view to the table generation
            };
            const tableHtml = this.generateTableDataSection(processedData, tableConfig, `${card.name}`, true);
            cardContent = tableHtml;
          } else {
            // Show summary only when table data is disabled
            let duplicateInfo = '';
            if (config.removeDuplicates && config.duplicateColumns.length > 0 && removedCount > 0) {
              duplicateInfo = `<div style="color:#1976d2; font-size:10px; margin-top:0.2rem;">Removed ${removedCount} duplicates</div>`;
            }
            
            cardContent = `
              <div style="margin-bottom:2rem; padding:0.8rem; background:#f8f9fa; border-radius:6px; border-left:3px solid #47B2E5;">
                <h4 style="margin:0 0 0.3rem 0; color:#1a2332; font-size:12px;">${card.name}</h4>
                <div style="color:#47B2E5; font-weight:600; font-size:11px;">${processedData.length.toLocaleString()} records</div>
                <div style="color:#666; font-size:10px; margin-top:0.2rem;">Table data disabled in configuration</div>
                ${duplicateInfo}
              </div>
            `;
          }
        }
        
        tablesHtml += cardContent;
        totalRecords += processedData.length;
      } else {
        // No data available for this card
        tablesHtml += `
          <div style="margin-bottom:2rem; padding:0.8rem; background:#fff3cd; border-radius:6px; border-left:3px solid #ffc107;">
            <h4 style="margin:0 0 0.3rem 0; color:#856404; font-size:12px;">${card.name}</h4>
            <div style="color:#856404; font-size:10px;">No data available for this card</div>
            <div style="color:#856404; font-size:9px; margin-top:0.2rem;">Card count: ${card.count}, Active: ${card.active}</div>
          </div>
        `;
      }
    });
    
    // Add duplicate removal summary if any duplicates were removed
    if (totalRemovedDuplicates > 0) {
      const duplicateSummary = `
        <div style="margin-bottom:2rem; padding:1rem; background:#e8f5e8; border-radius:6px; border-left:3px solid #28a745;">
          <h4 style="margin:0 0 0.5rem 0; color:#155724; font-size:12px;">Duplicate Removal Summary</h4>
          <div style="color:#155724; font-size:11px;">
            <div>Original records: ${totalOriginalRecords.toLocaleString()}</div>
            <div>Duplicates removed: ${totalRemovedDuplicates.toLocaleString()}</div>
            <div>Final records: ${totalRecords.toLocaleString()}</div>
          </div>
        </div>
      `;
      tablesHtml = duplicateSummary + tablesHtml;
    }
    
    return tablesHtml;
  }

  generateSeparateTablesForCards(selectedCards, config) {
    if (!selectedCards || selectedCards.length === 0) {
      return '<div style="color:#888; text-align:center; padding:2rem;">No cards selected for separate tables.</div>';
    }

    let tablesHtml = '';
    let totalRecords = 0;
    let totalOriginalRecords = 0;
    let totalRemovedDuplicates = 0;
    
    // Generate summary section first
    const summarySection = this.generateSummarySection(selectedCards, config);
    tablesHtml += summarySection;
    
    // Process each selected card with its individual configuration
    selectedCards.forEach((card, index) => {
      const cardConfig = this.getCardConfig(card.id);
      console.log(`🔍 Processing card "${card.name}" (ID: ${card.id}):`, {
        cardConfig: cardConfig,
        cardCount: card.count,
        cardActive: card.active
      });
      
      const cardData = this.getDataForQuickCardWithConfig(card, cardConfig);
      console.log(`📊 Card "${card.name}" data result:`, {
        dataLength: cardData?.length || 0,
        dataType: typeof cardData,
        isArray: Array.isArray(cardData),
        firstRow: cardData?.[0] || 'N/A',
        cardCount: card.count,
        expectedRecords: card.count,
        hasQuickFilter: !!(cardConfig && cardConfig.quickFilter && cardConfig.quickFilter !== 'none'),
        quickFilterName: cardConfig?.quickFilter || 'none'
      });
      
      // Debug: Show sample of filtered data
      if (cardData && cardData.length > 0) {
        console.log(`📋 Sample filtered data for "${card.name}":`, cardData.slice(0, 2));
      }
      
      // Debug: Check if there's a mismatch between card count and actual data
      if (card.count > 0 && (!cardData || cardData.length === 0)) {
        console.error(`🚨 CRITICAL: Card "${card.name}" shows ${card.count} records but getDataForQuickCardWithConfig returned ${cardData?.length || 0} records`);
        console.error(`🚨 This is why "No data available for this card" is showing in the preview`);
      }
      
      if (cardData && cardData.length > 0) {
        totalOriginalRecords += cardData.length;
        
        // Apply duplicate removal if enabled
        const originalCount = cardData.length;
        console.log(`Processing card "${card.name}": ${originalCount} original records`);
        
        const processedData = this.removeDuplicatesFromData(cardData, config);
        const removedCount = originalCount - processedData.length;
        totalRemovedDuplicates += removedCount;
        
        console.log(`Card "${card.name}" after duplicate removal: ${processedData.length} records (removed ${removedCount})`);
        
        // Use card's individual view type setting
        const cardViewConfig = {
          ...config,
          viewType: cardConfig.viewType || 'both'
        };
        
        // Generate appropriate content based on card's view type
        let cardContent = '';
        
        if (cardConfig.viewType === 'summary') {
          // Only show summary info for this card
          let duplicateInfo = '';
          if (config.removeDuplicates && config.duplicateColumns.length > 0 && removedCount > 0) {
            duplicateInfo = `<div style="color:#1976d2; font-size:10px; margin-top:0.2rem;">Removed ${removedCount} duplicates</div>`;
          }
          
          cardContent = `
            <div style="margin-bottom:2rem; padding:0.8rem; background:#f8f9fa; border-radius:6px; border-left:3px solid #47B2E5;">
              <h4 style="margin:0 0 0.3rem 0; color:#1a2332; font-size:12px;">${card.name}</h4>
              <div style="color:#47B2E5; font-weight:600; font-size:11px;">${processedData.length.toLocaleString()} records</div>
              <div style="color:#666; font-size:10px; margin-top:0.2rem;">View type: Summary only</div>
              ${duplicateInfo}
            </div>
          `;
        } else if (cardConfig.viewType === 'charts') {
          // Show charts placeholder for this card
          cardContent = `
            <div style="margin-bottom:2rem;">
              <h3 style="font-size:1.1rem; color:#15364A; margin-bottom:0.5rem; border-bottom:1px solid #e3f2fd; padding-bottom:0.3rem;">
                ${card.name} - Charts
              </h3>
              <div style="color:#888; text-align:center; padding:1rem; background:#f8f9fa; border-radius:6px; font-size:10px;">
                Charts functionality will be available in future updates
              </div>
            </div>
          `;
        } else {
          // Show table data (for 'table' or 'both' view types) - ONLY if includeTableData is enabled
          if (config.includeTableData) {
            const tableConfig = {
              ...cardViewConfig,
              savedView: cardConfig.savedView // Pass the saved view to the table generation
            };
            const tableHtml = this.generateTableDataSection(processedData, tableConfig, `${card.name}`, false);
            cardContent = tableHtml;
          } else {
            // Show summary only when table data is disabled
            cardContent = `
              <div style="margin-bottom:2rem; padding:0.8rem; background:#f8f9fa; border-radius:6px; border-left:3px solid #47B2E5;">
                <h4 style="margin:0 0 0.3rem 0; color:#1a2332; font-size:12px;">${card.name}</h4>
                <div style="color:#47B2E5; font-weight:600; font-size:11px;">${processedData.length.toLocaleString()} records</div>
                <div style="color:#666; font-size:10px; margin-top:0.2rem;">Table data disabled in report options</div>
              </div>
            `;
          }
        }
        
        tablesHtml += cardContent;
        totalRecords += processedData.length;
        
        // Add card configuration info with duplicate removal info - ONLY if enabled
        if (config.includeCardConfig) {
          let duplicateInfo = '';
          if (config.removeDuplicates && config.duplicateColumns.length > 0) {
            if (removedCount > 0) {
              duplicateInfo = ` | Duplicates removed: ${removedCount}`;
            } else {
              duplicateInfo = ` | No duplicates found`;
            }
          }
          
          const configInfo = `
            <div style="margin-bottom:2rem; padding:1rem; background:#e3f2fd; border-radius:6px; font-size:9px; color:#1976d2; line-height:1.5;">
              <strong>Card Configuration:</strong><br>
              View: ${cardConfig.viewType} | 
              Saved View: ${cardConfig.savedView} | 
              Active: ${cardConfig.active ? 'Yes' : 'No'} | 
              Records: ${processedData.length.toLocaleString()}${duplicateInfo}
            </div>
          `;
          tablesHtml += configInfo;
        }
              } else {
        // No data for this card - REPLICADO DEL REPOSITORIO THE BRIDGE
          tablesHtml += `
            <div style="margin-bottom:2rem;">
              <h3 style="font-size:1.1rem; color:#15364A; margin-bottom:0.5rem; border-bottom:1px solid #e3f2fd; padding-bottom:0.3rem;">
                ${card.name}
              </h3>
              <div style="color:#888; text-align:center; padding:1rem; background:#f8f9fa; border-radius:6px; font-size:10px;">
                No data available for this card
              </div>
            </div>
          `;
        }
    });

    // Add overall duplicate removal summary if enabled
    let duplicateSummary = '';
    if (config.removeDuplicates && config.duplicateColumns.length > 0 && totalRemovedDuplicates > 0 && config.includeDuplicateSummary) {
      duplicateSummary = `
        <div style="margin-bottom:0.8rem; padding:0.8rem; background:#e8f5e8; border-radius:6px; border-left:3px solid #4caf50;">
          <div style="color:#2e7d32; font-weight:600; margin-bottom:0.3rem; font-size:11px;">✓ Duplicate Removal Summary</div>
          <div style="color:#388e3c; font-size:10px;">
            Removed ${totalRemovedDuplicates.toLocaleString()} duplicate records across all cards<br>
            Based on columns: ${config.duplicateColumns.join(', ')}
          </div>
        </div>
      `;
    }

    // Generate title section only if summary info is enabled
    let titleSection = '';
    if (config.includeSummaryInfo) {
      titleSection = `
        <h2 style="font-size:1.1rem; color:#15364A; margin-bottom:1.5rem; border-bottom:1px solid #e3f2fd; padding-bottom:0.5rem;">
          Quick Cards Data (${selectedCards.length} cards, ${totalRecords.toLocaleString()} total records)
        </h2>
      `;
    }

    return `
      <div style="margin-bottom:3rem;">
        ${titleSection}
        ${duplicateSummary}
        ${tablesHtml}
      </div>
    `;
  }

  showPerformanceInfo() {
    const config = this.getReportConfig();
    const selectedCards = Object.values(this.quickCardsData).filter(card => this.selectedCards.has(card.id));
    
    console.log('Performance Info:', {
      selectedCards: selectedCards.length,
      viewType: config.viewType,
      separateTables: config.separateTablesPerCard,
      tableView: config.tableView,
      cacheSize: this.cachedData.size,
      isGeneratingPreview: this.isGeneratingPreview
    });
  }

  // Clear cache when data changes significantly
  clearCache() {
    this.cachedData.clear();
    console.log('Cache cleared');
  }

  handleDuplicateRemovalChange() {
    const removeDuplicatesCheckbox = document.getElementById('removeDuplicatesCheckbox');
    const duplicateColumnsSection = document.getElementById('duplicateColumnsSection');
    
    if (removeDuplicatesCheckbox && duplicateColumnsSection) {
      if (removeDuplicatesCheckbox.checked) {
        duplicateColumnsSection.style.display = 'block';
        this.populateDuplicateColumns();
      } else {
        duplicateColumnsSection.style.display = 'none';
      }
    }
    
    this.renderSummaryPreview();
  }

  populateDuplicateColumns() {
    const duplicateColumnsList = document.getElementById('duplicateColumnsList');
    if (!duplicateColumnsList) return;

    // Get data from selected cards or use filtered data as fallback
    const selectedCards = Object.values(this.quickCardsData).filter(card => this.selectedCards.has(card.id));
    
    let allData = [];
    
    if (selectedCards.length > 0) {
      // Combine data from all selected cards
      selectedCards.forEach(card => {
        const cardConfig = this.getCardConfig(card.id);
        const cardData = this.getDataForQuickCardWithConfig(card, cardConfig);
        if (cardData && cardData.length > 0) {
          allData = allData.concat(cardData);
        }
      });
    } else {
      // Use filtered data as fallback when no cards are selected (respects active filters)
      allData = this.getFilteredData();
    }

    if (allData.length === 0) {
      duplicateColumnsList.innerHTML = '<div style="color:#888; text-align:center; padding:1rem;">No data available. Please load data first.</div>';
      return;
    }

    const headers = Object.keys(allData[0]);
    duplicateColumnsList.innerHTML = '';

    // Calculate unique values for each column to help user choose
    const columnStats = headers.map(header => {
      const uniqueValues = new Set(allData.map(row => row[header] || '').filter(val => val !== ''));
      const totalValues = allData.length;
      const uniqueCount = uniqueValues.size;
      const duplicateRate = totalValues > 0 ? ((totalValues - uniqueCount) / totalValues * 100).toFixed(1) : 0;
      
      return {
        header,
        uniqueCount,
        totalValues,
        duplicateRate,
        isCommon: this.isCommonField(header)
      };
    });

    // Sort by duplicate rate (highest first) and then by common fields
    columnStats.sort((a, b) => {
      if (a.isCommon && !b.isCommon) return -1;
      if (!a.isCommon && b.isCommon) return 1;
      return parseFloat(b.duplicateRate) - parseFloat(a.duplicateRate);
    });

    columnStats.forEach(stat => {
      const item = document.createElement('div');
      item.className = 'column-checkbox-item';
      item.setAttribute('data-field', stat.header);
      
      const commonBadge = stat.isCommon ? '<span style="color:#47B2E5; font-size:0.8em; margin-left:0.5rem;">(common)</span>' : '';
      const duplicateInfo = stat.totalValues > 0 ? 
        `<span style="color:rgba(232,244,248,0.7); font-size:0.8em; margin-left:0.5rem;">${stat.uniqueCount}/${stat.totalValues} unique (${stat.duplicateRate}% dupes)</span>` : '';
      
      item.innerHTML = `
        <label class="checkbox-label">
          <input type="checkbox" value="${stat.header}" class="duplicate-column-checkbox">
          <span class="checkmark"></span>
          ${stat.header}${commonBadge}${duplicateInfo}
        </label>
      `;
      
      const checkbox = item.querySelector('.duplicate-column-checkbox');
      checkbox.addEventListener('change', () => this.updateDuplicateColumns());
      
      duplicateColumnsList.appendChild(item);
    });

    this.updateDuplicateColumns();
  }

  isCommonField(fieldName) {
    const commonFields = [
      'id', 'ID', 'Id', 'booking', 'Booking', 'reference', 'Reference', 'ref', 'Ref',
      'order', 'Order', 'number', 'Number', 'code', 'Code', 'tracking', 'Tracking',
      'email', 'Email', 'phone', 'Phone', 'customer', 'Customer', 'client', 'Client',
      'name', 'Name', 'title', 'Title', 'description', 'Description'
    ];
    
    return commonFields.some(common => 
      fieldName.toLowerCase().includes(common.toLowerCase())
    );
  }

  selectAllDuplicateColumns() {
    const checkboxes = document.querySelectorAll('.duplicate-column-checkbox');
    checkboxes.forEach(checkbox => {
      checkbox.checked = true;
    });
    this.updateDuplicateColumns();
  }

  deselectAllDuplicateColumns() {
    const checkboxes = document.querySelectorAll('.duplicate-column-checkbox');
    checkboxes.forEach(checkbox => {
      checkbox.checked = false;
    });
    this.updateDuplicateColumns();
  }

  updateDuplicateColumns() {
    const checkboxes = document.querySelectorAll('.duplicate-column-checkbox:checked');
    this.reportConfig.duplicateColumns = Array.from(checkboxes).map(cb => cb.value);
    
    // Update selected fields count
    const selectedFieldsCount = document.getElementById('selectedFieldsCount');
    if (selectedFieldsCount) {
      selectedFieldsCount.textContent = this.reportConfig.duplicateColumns.length;
    }
    
    // Show duplicate summary if fields are selected
    this.showDuplicateSummary();
    
    this.renderSummaryPreview();
  }

  showDuplicateSummary() {
    const duplicateSummary = document.getElementById('duplicateSummary');
    if (!duplicateSummary) return;

    if (this.reportConfig.duplicateColumns.length === 0) {
      duplicateSummary.style.display = 'none';
      return;
    }

    // Prevent excessive calls
    if (this.duplicateAnalysisTimeout) {
      clearTimeout(this.duplicateAnalysisTimeout);
    }
    
    this.duplicateAnalysisTimeout = setTimeout(() => {
      this._showDuplicateSummaryInternal(duplicateSummary);
    }, 200); // 200ms debounce
  }

  _showDuplicateSummaryInternal(duplicateSummary) {
    try {
      // Get the correct data based on current configuration
      const config = this.getReportConfig();
      let tableData = this.getDataForCurrentConfiguration(config);
      
      if (!tableData || tableData.length === 0) {
        duplicateSummary.style.display = 'none';
        return;
      }

      // Calculate potential duplicates
      const seen = new Set();
      let duplicateCount = 0;
      
      tableData.forEach(row => {
        const key = config.duplicateColumns.map(col => row[col] || '').join('|');
        if (seen.has(key)) {
          duplicateCount++;
        } else {
          seen.add(key);
        }
      });

      const totalRecords = tableData.length;
      const uniqueRecords = totalRecords - duplicateCount;
      const duplicatePercentage = totalRecords > 0 ? ((duplicateCount / totalRecords) * 100).toFixed(1) : 0;

      // Add configuration info to the summary
      let configInfo = '';
      if (config.separateTablesPerCard) {
        const selectedCards = Object.values(this.quickCardsData).filter(card => this.selectedCards.has(card.id));
        configInfo = `<br><strong>Configuration:</strong> Separate tables for ${selectedCards.length} selected cards`;
      } else if (config.tableView !== 'current') {
        configInfo = `<br><strong>Configuration:</strong> Using saved view: "${config.tableView}"`;
      }

      duplicateSummary.innerHTML = `
        <strong>Duplicate Analysis:</strong>${configInfo}<br>
        • Total records: ${totalRecords.toLocaleString()}<br>
        • Unique records: ${uniqueRecords.toLocaleString()}<br>
        • Duplicate records: ${duplicateCount.toLocaleString()} (${duplicatePercentage}%)<br>
        • Fields used: ${config.duplicateColumns.join(', ')}
      `;

      duplicateSummary.className = duplicateCount > 0 ? 'duplicate-summary' : 'duplicate-summary warning';
      duplicateSummary.style.display = 'block';
    } catch (error) {
      console.error('Error in duplicate analysis:', error);
      duplicateSummary.innerHTML = '<strong>Error:</strong> Could not analyze duplicates';
      duplicateSummary.style.display = 'block';
    }
  }

  debugQuickFilters() {
    const quickFilters = this.loadQuickFilters();
    console.log('Available quick filters:', quickFilters);
    
    Object.values(this.quickCardsData).forEach(card => {
      console.log(`\nChecking card: "${card.name}" (ID: ${card.id})`);
      
      const filterEntry = Object.entries(quickFilters).find(([name, obj]) => {
        const matchByName = obj.name === card.name;
        const matchByLinkedCard = obj.linkedUrgencyCard === card.id;
        const matchByNameContains = name.toLowerCase().includes(card.name.toLowerCase());
        const matchByObjNameContains = obj.name && obj.name.toLowerCase().includes(card.name.toLowerCase());
        
        console.log(`  Filter "${name}":`, {
          matchByName,
          matchByLinkedCard,
          matchByNameContains,
          matchByObjNameContains,
          objName: obj.name,
          linkedCard: obj.linkedUrgencyCard
        });
        
        return matchByName || matchByLinkedCard || matchByNameContains || matchByObjNameContains;
      });
      
      if (filterEntry) {
        console.log(`  ✓ Found matching filter: "${filterEntry[0]}"`);
      } else {
        console.log(`  ✗ No matching filter found`);
      }
    });
  }

  // ===== PREFERENCES SYSTEM =====

  saveOpsPreferences() {
    try {
      // Get current state for complete preview
      const selectedCards = Array.from(this.selectedCards);
      const selectedCardsNames = selectedCards.map(cardId => {
        return this.quickCardsData[cardId]?.name || cardId;
      });
      
      // Get current card configurations with saved views
      const cardConfigurations = {};
      selectedCards.forEach(cardId => {
        const cardData = this.quickCardsData[cardId];
        if (cardData) {
          cardConfigurations[cardId] = {
            id: cardId,
            name: cardData.name,
            savedView: cardData.savedView || '__all__',
            config: this.getCardConfig(cardId)
          };
        }
      });
      
      // Collect current configuration
      const preferences = {
        timestamp: new Date().toISOString(),
        selectedCards: selectedCards,
        cardConfigs: { ...this.cardConfigs },
        cardConfigurations: cardConfigurations,
        reportConfig: { ...this.reportConfig },
        // Get checkbox states
        includeTableData: document.getElementById('includeTableDataCheckbox')?.checked || false,
        includeTechnicalInfo: document.getElementById('includeTechnicalInfoCheckbox')?.checked || false,
        removeDuplicates: document.getElementById('removeDuplicatesCheckbox')?.checked || false,
        separateTablesPerCard: document.getElementById('separateTablesPerCardCheckbox')?.checked || false,
        // Get duplicate columns
        duplicateColumns: Array.from(document.querySelectorAll('#duplicateColumnsList input[type="checkbox"]:checked')).map(cb => cb.value),
        // Quick filters state
        activeQuickFilters: window.activeOpsQuickFilters || [],
        // Saved views state
        savedViews: window.loadSavedViews ? window.loadSavedViews() : {}
      };

      // Get existing preferences
      const existingPreferences = JSON.parse(localStorage.getItem('opsSummaryPreferences') || '[]');
      
      // Add new preference (limit to 10)
      existingPreferences.unshift(preferences);
      if (existingPreferences.length > 10) {
        existingPreferences.pop();
      }
      
      // Save to localStorage
      localStorage.setItem('opsSummaryPreferences', JSON.stringify(existingPreferences));
      
      this.showNotification('Preferences saved successfully!', 'success');
      this.showSavePreferencesModal();
    } catch (error) {
      console.error('Error saving preferences:', error);
      this.showNotification('Error saving preferences', 'error');
    }
  }

  loadOpsPreferences() {
    try {
      const preferences = JSON.parse(localStorage.getItem('opsSummaryPreferences') || '[]');
      
      if (preferences.length === 0) {
        this.showNotification('No saved preferences found', 'info');
        return;
      }
      
      this.showPreferencesSelectionModal(preferences);
    } catch (error) {
      console.error('Error loading preferences:', error);
      this.showNotification('Error loading preferences', 'error');
    }
  }

  showPreferencesSelectionModal(preferences) {
    // Remove existing modal if any
    const existingModal = document.getElementById('opsPreferencesModal');
    if (existingModal) {
      existingModal.remove();
    }

    const modal = document.createElement('div');
    modal.id = 'opsPreferencesModal';
    modal.className = 'modal-overlay';
    modal.style.cssText = `
      position: fixed !important; top: 0 !important; left: 0 !important; width: 100% !important; height: 100% !important;
      background: transparent !important; backdrop-filter: none !important; z-index: 10000 !important;
      display: flex !important; align-items: center !important; justify-content: center !important; opacity: 1 !important;
    `;

    const content = document.createElement('div');
    content.className = 'modal-panel';
    content.style.cssText = `
      background: #1a2332 !important; border: 2px solid rgba(255,255,255,0.3) !important; 
      border-radius: 12px !important; padding: 0 !important; max-width: 600px !important; 
      width: 90% !important; max-height: 80vh !important; overflow: hidden !important;
      box-shadow: 0 25px 80px rgba(0,0,0,0.95) !important; opacity: 1 !important;
    `;

    content.innerHTML = `
      <div class="modal-header" style="display: flex; justify-content: space-between; align-items: center; padding: 1.5rem 2rem; border-bottom: 1px solid rgba(255,255,255,0.1); background: rgba(255,255,255,0.05) !important;">
        <div class="header-left" style="display: flex; align-items: center; gap: 1rem;">
          <img src="./LOGOTAB_rounded.png" alt="Logo" style="width: 32px; height: 32px; border-radius: 6px; box-shadow: 0 2px 8px rgba(71,178,229,0.3);" onerror="console.error('Error loading logo:', this.src);">
          <h3 style="color: white; margin: 0; font-size: 1.3rem;">Load Saved Preferences</h3>
        </div>
        <button class="close-btn" style="background: none; border: none; color: white; font-size: 1.5rem; cursor: pointer; padding: 0.5rem;">×</button>
      </div>
      <div class="modal-content" style="padding: 2rem; max-height: 60vh; overflow-y: auto; background: transparent !important;">
        <div style="margin-bottom: 1.5rem;">
          <p style="color: rgba(255,255,255,0.8); margin-bottom: 1rem;">Select a saved preference to load:</p>
        </div>
        <div class="preferences-list" style="display: flex; flex-direction: column; gap: 1rem;">
          ${preferences.map((pref, index) => {
            const date = new Date(pref.timestamp).toLocaleString();
            const cardCount = pref.selectedCards.length;
            const configSummary = [
              pref.includeTableData ? 'Table Data' : null,
              pref.includeTechnicalInfo ? 'Technical Info' : null,
              pref.removeDuplicates ? 'Remove Duplicates' : null,
              pref.separateTablesPerCard ? 'Separate Tables' : null
            ].filter(Boolean).join(', ');
            
            return `
              <div class="preference-item" style="background: rgba(255,255,255,0.2) !important; border: 2px solid rgba(255,255,255,0.3) !important; border-radius: 8px; padding: 1.5rem; cursor: pointer; transition: all 0.3s ease; opacity: 1 !important; position: relative;" onclick="window.opsHubSummary.applyOpsPreferences(${index})">
                <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 0.5rem;">
                  <h4 style="color: white; margin: 0; font-size: 1.1rem;">${pref.name || `Saved Configuration ${index + 1}`}</h4>
                  <span style="color: rgba(255,255,255,0.6); font-size: 0.9rem;">${date}</span>
                </div>
                <div style="color: rgba(255,255,255,0.8); font-size: 0.9rem;">
                  <div><strong>Cards:</strong> ${cardCount} selected</div>
                  <div><strong>Options:</strong> ${configSummary || 'None'}</div>
                  ${pref.duplicateColumns.length > 0 ? `<div><strong>Duplicate Fields:</strong> ${pref.duplicateColumns.join(', ')}</div>` : ''}
                </div>
                <div style="position: absolute; bottom: 1rem; right: 1rem; opacity: 0; transition: opacity 0.3s ease;" class="action-buttons">
                  <button class="delete-btn" 
                          onclick="event.stopPropagation(); window.opsHubSummary.deleteOpsPreference(${index})"
                          style="background: rgba(244, 67, 54, 0.2); border: 1px solid rgba(244, 67, 54, 0.4); color: #f44336; padding: 0.2rem 0.4rem; border-radius: 4px; cursor: pointer; font-size: 0.7rem; transition: all 0.2s ease;"
                          onmouseover="this.style.background='rgba(244, 67, 54, 0.3)'; this.style.borderColor='rgba(244, 67, 54, 0.6)'"
                          onmouseout="this.style.background='rgba(244, 67, 54, 0.2)'; this.style.borderColor='rgba(244, 67, 54, 0.4)'"
                          title="Delete preference">
                    ×
                  </button>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      </div>
      <div class="modal-footer" style="padding: 1.5rem 2rem; border-top: 1px solid rgba(255,255,255,0.1); background: rgba(255,255,255,0.05) !important; display: flex; justify-content: flex-end; gap: 1rem;">
        <button class="modal-btn secondary" onclick="document.getElementById('opsPreferencesModal').remove()" style="background: rgba(255,255,255,0.1) !important; border: 1px solid rgba(255,255,255,0.2) !important; color: rgba(255,255,255,0.8) !important; padding: 0.75rem 1.5rem; border-radius: 6px; cursor: pointer; transition: all 0.3s ease;">Cancel</button>
      </div>
    `;

    modal.appendChild(content);
    document.body.appendChild(modal);

    // Add hover effects for action buttons
    const preferenceItems = modal.querySelectorAll('.preference-item');
    preferenceItems.forEach(item => {
      const actionButtons = item.querySelector('.action-buttons');
      item.addEventListener('mouseenter', () => {
        if (actionButtons) actionButtons.style.opacity = '1';
      });
      item.addEventListener('mouseleave', () => {
        if (actionButtons) actionButtons.style.opacity = '0';
      });
    });

    // Close button functionality
    const closeBtn = modal.querySelector('.close-btn');
    closeBtn.addEventListener('click', () => modal.remove());

    // Close on overlay click
    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        modal.remove();
      }
    });
  }

  deleteOpsPreference(prefIndex) {
    try {
      const existingPrefs = JSON.parse(localStorage.getItem('opsSummaryPreferences') || '[]');
      
      if (prefIndex >= 0 && prefIndex < existingPrefs.length) {
        const preferenceToDelete = existingPrefs[prefIndex];
        
        // Remove the preference
        existingPrefs.splice(prefIndex, 1);
        localStorage.setItem('opsSummaryPreferences', JSON.stringify(existingPrefs));
        
        // Close the modal and show success message
        const modal = document.getElementById('opsPreferencesModal');
        if (modal) modal.remove();
        
        this.showNotification(`Preference "${preferenceToDelete.name || `Configuration ${prefIndex + 1}`}" deleted successfully`, 'success');
        
        // Refresh the preferences modal if there are still preferences
        if (existingPrefs.length > 0) {
          this.loadOpsPreferences();
        }
      } else {
        this.showNotification('Invalid preference index', 'error');
      }
    } catch (error) {
      console.error('Error deleting Ops preference:', error);
      this.showNotification('Error deleting preference. Please try again.', 'error');
    }
  }

  applyOpsPreferences(prefIndex) {
    try {
      const preferences = JSON.parse(localStorage.getItem('opsSummaryPreferences') || '[]');
      const pref = preferences[prefIndex];
      
      if (!pref) {
        this.showNotification('Preference not found', 'error');
        return;
      }

      console.log('🔄 Applying complete Ops preferences:', pref);

      // Apply selected cards
      this.selectedCards.clear();
      pref.selectedCards.forEach(cardId => this.selectedCards.add(cardId));

      // Apply card configurations
      this.cardConfigs = { ...pref.cardConfigs };

      // Apply card configurations with saved views and individual settings
      if (pref.cardConfigurations) {
        Object.entries(pref.cardConfigurations).forEach(([cardId, config]) => {
          if (this.quickCardsData[cardId]) {
            // Apply individual card configurations
            if (config.viewType) {
              this.updateCardConfig(cardId, 'viewType', config.viewType);
            }
            if (config.savedView) {
              this.updateCardConfig(cardId, 'savedView', config.savedView);
            }

            // Apply individual card config if available
            if (config.config) {
              this.cardConfigs[cardId] = config.config;
            }
            
            // Also update the internal cardConfigs directly to ensure persistence
            if (!this.cardConfigs[cardId]) {
              this.cardConfigs[cardId] = {};
            }
            if (config.viewType) {
              this.cardConfigs[cardId].viewType = config.viewType;
            }
            if (config.savedView) {
              this.cardConfigs[cardId].savedView = config.savedView;
            }
          }
        });
      }

      // Apply report configuration
      this.reportConfig = { ...pref.reportConfig };

      // Apply checkbox states
      const includeTableDataCheckbox = document.getElementById('includeTableDataCheckbox');
      const includeTechnicalInfoCheckbox = document.getElementById('includeTechnicalInfoCheckbox');
      const removeDuplicatesCheckbox = document.getElementById('removeDuplicatesCheckbox');
      const separateTablesPerCardCheckbox = document.getElementById('separateTablesPerCardCheckbox');

      if (includeTableDataCheckbox) includeTableDataCheckbox.checked = pref.includeTableData;
      if (includeTechnicalInfoCheckbox) includeTechnicalInfoCheckbox.checked = pref.includeTechnicalInfo;
      if (removeDuplicatesCheckbox) removeDuplicatesCheckbox.checked = pref.removeDuplicates;
      if (separateTablesPerCardCheckbox) separateTablesPerCardCheckbox.checked = pref.separateTablesPerCard;

      // Apply duplicate columns
      if (pref.duplicateColumns && pref.duplicateColumns.length > 0) {
        const duplicateCheckboxes = document.querySelectorAll('#duplicateColumnsList input[type="checkbox"]');
        duplicateCheckboxes.forEach(cb => {
          cb.checked = pref.duplicateColumns.includes(cb.value);
        });
      }

      // Apply quick filters state
      if (pref.activeQuickFilters) {
        window.activeOpsQuickFilters = pref.activeQuickFilters;
        // Re-apply quick filters if function exists
        if (window.opsHubManager && window.opsHubManager.applyOpsQuickFilters) {
          window.opsHubManager.applyOpsQuickFilters();
        }
      }

      // Restore saved views if they exist
      if (pref.savedViews && Object.keys(pref.savedViews).length > 0) {
        // Update localStorage with saved views
        localStorage.setItem('tableViews', JSON.stringify(pref.savedViews));
        // Refresh saved views dropdown
        if (window.setupViewSelect) {
          window.setupViewSelect();
        }
        if (window.initializeColumnManager) {
          window.initializeColumnManager();
        }
      }

      // Ensure Ops Summary modal is open to show the cards
      if (!document.getElementById('opsSummaryModal')) {
        this.openSummaryModal();
      }
      
      // Update UI
      this.renderCardsSelection();
      this.renderSummaryPreview();
      this.handleDuplicateRemovalChange();
      
      // Force update the UI to reflect the selected cards with multiple attempts
      setTimeout(() => {
        this.renderCardsSelection();
        
        // Also update the checkboxes directly
        pref.selectedCards.forEach(cardId => {
          const cardElement = document.querySelector(`[data-card-id="${cardId}"]`);
          if (cardElement) {
            const checkbox = cardElement.querySelector('.card-checkbox-input');
            if (checkbox) {
              checkbox.checked = true;
            }
            // Also need to update dropdowns directly
            const viewTypeSelect = cardElement.querySelector('.card-view-type');
            const savedViewSelect = cardElement.querySelector('.card-saved-view');
            const config = pref.cardConfigurations[cardId];
            if (viewTypeSelect && config && config.viewType) {
                viewTypeSelect.value = config.viewType;
            }
            if (savedViewSelect && config && config.savedView) {
                savedViewSelect.value = config.savedView;
            }
          }
        });
      }, 200);
      
      // Additional update after a longer delay to ensure everything is rendered
      setTimeout(() => {
        this.renderCardsSelection();
      }, 500);

      // Close modal
      const modal = document.getElementById('opsPreferencesModal');
      if (modal) modal.remove();

      this.showNotification('Complete preferences applied successfully!', 'success');
      console.log('✅ Complete Ops Preferences applied:', pref);
    } catch (error) {
      console.error('Error applying preferences:', error);
      this.showNotification('Error applying preferences', 'error');
    }
  }

  showSavePreferencesModal() {
    // Remove existing modal if any
    const existingModal = document.getElementById('opsSavePreferencesModal');
    if (existingModal) {
      existingModal.remove();
    }

    const modal = document.createElement('div');
    modal.id = 'opsSavePreferencesModal';
    modal.className = 'modal-overlay';
    modal.style.cssText = `
      position: fixed !important; top: 0 !important; left: 0 !important; width: 100% !important; height: 100% !important;
      background: transparent !important; backdrop-filter: none !important; z-index: 10000 !important;
      display: flex !important; align-items: center !important; justify-content: center !important; opacity: 1 !important;
    `;

    const content = document.createElement('div');
    content.className = 'modal-panel';
    content.style.cssText = `
      background: #1a2332 !important; border: 2px solid rgba(255,255,255,0.3) !important; 
      border-radius: 12px !important; padding: 0 !important; max-width: 500px !important; 
      width: 90% !important; max-height: 80vh !important; overflow: hidden !important;
      box-shadow: 0 25px 80px rgba(0,0,0,0.95) !important; opacity: 1 !important;
    `;

    // Get current configuration summary
    const selectedCards = Array.from(this.selectedCards);
    const selectedCardsNames = selectedCards.map(cardId => {
      return this.quickCardsData[cardId]?.name || cardId;
    });
    const cardCount = selectedCards.length;
    const activeFilters = window.activeOpsQuickFilters || [];
    const savedViewsCount = window.loadSavedViews ? Object.keys(window.loadSavedViews()).length : 0;
    const configSummary = [
      document.getElementById('includeTableDataCheckbox')?.checked ? 'Table Data' : null,
      document.getElementById('includeTechnicalInfoCheckbox')?.checked ? 'Technical Info' : null,
      document.getElementById('removeDuplicatesCheckbox')?.checked ? 'Remove Duplicates' : null,
      document.getElementById('separateTablesPerCardCheckbox')?.checked ? 'Separate Tables' : null
    ].filter(Boolean).join(', ');

    content.innerHTML = `
      <div class="modal-header" style="display: flex; justify-content: space-between; align-items: center; padding: 1.5rem 2rem; border-bottom: 1px solid rgba(255,255,255,0.1); background: rgba(255,255,255,0.05) !important;">
        <div class="header-left" style="display: flex; align-items: center; gap: 1rem;">
          <img src="./LOGOTAB_rounded.png" alt="Logo" style="width: 32px; height: 32px; border-radius: 6px; box-shadow: 0 2px 8px rgba(71,178,229,0.3);" onerror="console.error('Error loading logo:', this.src);">
          <h3 style="color: white; margin: 0; font-size: 1.3rem;">Save Complete Ops Preferences</h3>
        </div>
        <button class="close-btn" style="background: none; border: none; color: white; font-size: 1.5rem; cursor: pointer; padding: 0.5rem;">×</button>
      </div>
      <div class="modal-content" style="padding: 2rem; background: transparent !important;">
        <div style="margin-bottom: 1.5rem;">
          <p style="color: rgba(255,255,255,0.8); margin-bottom: 1rem;">Complete configuration to save:</p>
          <div style="background: rgba(255,255,255,0.1) !important; border: 1px solid rgba(255,255,255,0.2) !important; border-radius: 8px; padding: 1rem; opacity: 1 !important;">
            <div style="color: rgba(255,255,255,0.9); margin-bottom: 0.5rem;"><strong>Selected Cards (${cardCount}):</strong> ${selectedCardsNames.length > 0 ? selectedCardsNames.join(', ') : 'None selected'}</div>
            <div style="color: rgba(255,255,255,0.9); margin-bottom: 0.5rem;"><strong>Active Quick Filters (${activeFilters.length}):</strong> ${activeFilters.length > 0 ? activeFilters.join(', ') : 'None'}</div>
            <div style="color: rgba(255,255,255,0.9); margin-bottom: 0.5rem;"><strong>Report Settings:</strong> ${configSummary || 'None'}</div>
            <div style="color: rgba(255,255,255,0.9); margin-bottom: 0.5rem;"><strong>Saved Views:</strong> ${savedViewsCount} views available</div>
            ${document.getElementById('removeDuplicatesCheckbox')?.checked ? `<div style="color: rgba(255,255,255,0.9);"><strong>Duplicate Fields:</strong> ${Array.from(document.querySelectorAll('#duplicateColumnsList input[type="checkbox"]:checked')).map(cb => cb.value).join(', ') || 'None'}</div>` : ''}
          </div>
          
          <div style="background: rgba(76,175,80,0.1) !important; border-radius: 8px; padding: 1rem; margin-top: 1rem; opacity: 1 !important; border: 1px solid rgba(76,175,80,0.2) !important;">
            <div style="color: #4CAF50; font-size: 0.9rem; font-weight: 600;">Complete State Save</div>
            <div style="color: rgba(255,255,255,0.8); font-size: 0.85rem; margin-top: 0.3rem;">
              This will save everything: ${cardCount} cards, ${activeFilters.length} filters, ${savedViewsCount} saved views, and all report settings for instant restoration.
            </div>
          </div>
        </div>
        <div style="margin-bottom: 1.5rem;">
          <label style="display: block; color: rgba(255,255,255,0.9); margin-bottom: 0.5rem; font-weight: 500;">Configuration Name (optional):</label>
          <input type="text" id="opsPreferenceName" placeholder="Enter a name for this configuration..." style="width: 100%; padding: 0.75rem; border: 1px solid rgba(255,255,255,0.3) !important; border-radius: 6px; background: rgba(255,255,255,0.1) !important; color: white; font-size: 0.9rem; opacity: 1 !important;">
        </div>
      </div>
      <div class="modal-footer" style="padding: 1.5rem 2rem; border-top: 1px solid rgba(255,255,255,0.1); background: rgba(255,255,255,0.05) !important; display: flex; justify-content: flex-end; gap: 1rem;">
        <button class="modal-btn secondary" onclick="document.getElementById('opsSavePreferencesModal').remove()" style="background: rgba(255,255,255,0.1) !important; border: 1px solid rgba(255,255,255,0.2) !important; color: rgba(255,255,255,0.8) !important; padding: 0.75rem 1.5rem; border-radius: 6px; cursor: pointer; transition: all 0.3s ease;">Cancel</button>
        <button class="modal-btn primary" onclick="window.opsHubSummary.confirmSaveOpsPreferences()" style="background: #47B2E5 !important; border: 1px solid #47B2E5 !important; color: white !important; padding: 0.75rem 1.5rem; border-radius: 6px; cursor: pointer; transition: all 0.3s ease; opacity: 1 !important;">Save Complete Preferences</button>
      </div>
    `;

    modal.appendChild(content);
    document.body.appendChild(modal);

    // Close button functionality
    const closeBtn = modal.querySelector('.close-btn');
    closeBtn.addEventListener('click', () => modal.remove());

    // Close on overlay click
    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        modal.remove();
      }
    });
  }

  confirmSaveOpsPreferences() {
    try {
      const nameInput = document.getElementById('opsPreferenceName');
      const name = nameInput ? nameInput.value.trim() : '';
      
      // Get current selected cards from multiple sources
      const selectedCards = [];
      const cardConfigurations = {};
      
      // First, try to get from UI checkboxes
      const cardCheckboxes = document.querySelectorAll('#opsQuickCardsSelection .card-checkbox-input:checked');
      if (cardCheckboxes.length > 0) {
        cardCheckboxes.forEach(checkbox => {
          const cardElement = checkbox.closest('.card-selection-item');
          if (cardElement) {
            const cardId = cardElement.dataset.cardId;
            if (cardId) {
              selectedCards.push(cardId);
              
              // Get card configuration from UI elements
              const viewTypeSelect = cardElement.querySelector('.card-view-type');
              const savedViewSelect = cardElement.querySelector('.card-saved-view');
              
              cardConfigurations[cardId] = {
                id: cardId,
                name: cardElement.querySelector('h4')?.textContent || cardId,
                viewType: viewTypeSelect?.value || 'both',
                savedView: savedViewSelect?.value || 'current',
                config: this.getCardConfig(cardId)
              };
            }
          }
        });
      }
      
      // If no cards found in UI, try to get from internal state
      if (selectedCards.length === 0 && this.selectedCards) {
        this.selectedCards.forEach(cardId => {
          selectedCards.push(cardId);
          
          // Get configuration from internal state
          if (this.quickCardsData[cardId]) {
            const cardData = this.quickCardsData[cardId];
            const cardConfig = this.getCardConfig(cardId);
            
            cardConfigurations[cardId] = {
              id: cardId,
              name: cardData.name,
              viewType: cardConfig.viewType || 'both',
              savedView: cardConfig.savedView || 'current',
              config: cardConfig
            };
          }
        });
      }
      
      // If still no cards, try to get all available cards as fallback
      if (selectedCards.length === 0 && this.quickCardsData) {
        Object.keys(this.quickCardsData).forEach(cardId => {
          selectedCards.push(cardId);
          
          const cardData = this.quickCardsData[cardId];
          const cardConfig = this.getCardConfig(cardId);
          
          cardConfigurations[cardId] = {
            id: cardId,
            name: cardData.name,
            viewType: cardConfig.viewType || 'both',
            savedView: cardConfig.savedView || 'current',
            config: cardConfig
          };
        });
      }
      
      const selectedCardsNames = selectedCards.map(cardId => {
        return this.quickCardsData[cardId]?.name || cardId;
      });
      
      // Get current configuration
      const preferences = {
        timestamp: new Date().toISOString(),
        name: name || `Complete Ops Configuration ${new Date().toLocaleString()}`,
        selectedCards: selectedCards,
        cardConfigs: { ...this.cardConfigs },
        cardConfigurations: cardConfigurations,
        reportConfig: { ...this.reportConfig },
        includeTableData: document.getElementById('includeTableDataCheckbox')?.checked || false,
        includeTechnicalInfo: document.getElementById('includeTechnicalInfoCheckbox')?.checked || false,
        removeDuplicates: document.getElementById('removeDuplicatesCheckbox')?.checked || false,
        separateTablesPerCard: document.getElementById('separateTablesPerCardCheckbox')?.checked || false,
        duplicateColumns: Array.from(document.querySelectorAll('#duplicateColumnsList input[type="checkbox"]:checked')).map(cb => cb.value),
        // Quick filters state
        activeQuickFilters: window.activeOpsQuickFilters || [],
        // Saved views state
        savedViews: window.loadSavedViews ? window.loadSavedViews() : {}
      };

      // Get existing preferences
      const existingPreferences = JSON.parse(localStorage.getItem('opsSummaryPreferences') || '[]');
      
      // Add new preference (limit to 10)
      existingPreferences.unshift(preferences);
      if (existingPreferences.length > 10) {
        existingPreferences.pop();
      }
      
      // Save to localStorage
      localStorage.setItem('opsSummaryPreferences', JSON.stringify(existingPreferences));
      
      // Close modal
      const modal = document.getElementById('opsSavePreferencesModal');
      if (modal) modal.remove();
      
      this.showNotification(`Configuration "${preferences.name}" saved successfully!`, 'success');
    } catch (error) {
      console.error('Error saving preferences:', error);
      this.showNotification('Error saving preferences', 'error');
    }
  }
}

// Note: This class is initialized from main.js 