/**
 * AutoVault - Car Registration & Expiry Tracker
 * Core Application Engine
 */
(function () {
  'use strict';
  // LocalStorage Key Constant
  const STORAGE_KEY = 'autovault_vehicles_v1';
  const THEME_KEY = 'autovault_theme_pref';
  // Sample Pre-populated Vehicles Data
  const INITIAL_DEMO_CARS = [
    {
      id: 'car_demo_1',
      plate: 'DXB-9874',
      model: 'Porsche 911 GT3 RS',
      expiryDate: getOffsetDateString(18), // Expiring in 18 days (Amber alert)
      type: 'Sports',
      color: '#EF4444',
      insuranceExpiry: getOffsetDateString(45),
      notes: 'Track package equipped. Annual service due at Porsche Center.',
      createdAt: Date.now() - 1000000
    },
    {
      id: 'car_demo_2',
      plate: 'CA-7890X',
      model: 'Tesla Model Y Long Range',
      expiryDate: getOffsetDateString(-5), // Expired 5 days ago (Red alert)
      type: 'Electric',
      color: '#3B82F6',
      insuranceExpiry: getOffsetDateString(60),
      notes: 'Registration renewal pending smog test verification.',
      createdAt: Date.now() - 2000000
    },
    {
      id: 'car_demo_3',
      plate: 'NY-4521K',
      model: 'Toyota Land Cruiser V8',
      expiryDate: getOffsetDateString(142), // Active (Green)
      type: 'SUV',
      color: '#10B981',
      insuranceExpiry: getOffsetDateString(180),
      notes: 'Comprehensive insurance active. Registration valid.',
      createdAt: Date.now() - 3000000
    }
  ];
  // Helper: Get YYYY-MM-DD string relative to today
  function getOffsetDateString(daysOffset) {
    const d = new Date();
    d.setDate(d.getDate() + daysOffset);
    return d.toISOString().split('T')[0];
  }
  // App State
  let vehicles = [];
  let currentFilter = 'all';
  let searchQuery = '';
  let currentSort = 'expiry-asc';
  let currentView = 'grid'; // 'grid' or 'list'
  // DOM Selectors
  const dom = {
    themeToggleBtn: document.getElementById('theme-toggle-btn'),
    exportBtn: document.getElementById('export-btn'),
    importBtn: document.getElementById('import-btn'),
    importFileInput: document.getElementById('import-file-input'),
    addCarBtn: document.getElementById('add-car-btn'),
    
    // Stats
    statTotal: document.getElementById('stat-total-count'),
    statActive: document.getElementById('stat-active-count'),
    statExpiring: document.getElementById('stat-expiring-count'),
    statExpired: document.getElementById('stat-expired-count'),
    // Alerts
    urgentAlertsContainer: document.getElementById('urgent-alerts-container'),
    // Controls
    searchInput: document.getElementById('search-input'),
    clearSearchBtn: document.getElementById('clear-search-btn'),
    filterTabs: document.querySelectorAll('.filter-tab'),
    sortSelect: document.getElementById('sort-select'),
    viewGridBtn: document.getElementById('view-grid-btn'),
    viewListBtn: document.getElementById('view-list-btn'),
    // Main List & Empty State
    vehiclesContainer: document.getElementById('vehicles-container'),
    emptyState: document.getElementById('empty-state'),
    emptyTitle: document.getElementById('empty-title'),
    emptyDescription: document.getElementById('empty-description'),
    emptyActionBtn: document.getElementById('empty-action-btn'),
    // Modal & Form
    carModal: document.getElementById('car-modal'),
    modalTitle: document.getElementById('modal-title'),
    closeModalBtn: document.getElementById('close-modal-btn'),
    cancelModalBtn: document.getElementById('cancel-modal-btn'),
    carForm: document.getElementById('car-form'),
    carIdInput: document.getElementById('car-id'),
    plateInput: document.getElementById('plate-input'),
    modelInput: document.getElementById('model-input'),
    expiryInput: document.getElementById('expiry-input'),
    typeInput: document.getElementById('type-input'),
    colorPicker: document.getElementById('color-picker'),
    colorInput: document.getElementById('color-input'),
    insuranceInput: document.getElementById('insurance-input'),
    notesInput: document.getElementById('notes-input'),
    
    // Live Plate Preview
    previewPlateText: document.getElementById('preview-plate-text'),
    
    // Toast
    toastContainer: document.getElementById('toast-container')
  };
  /* ==========================================
     Initialization & Storage
     ========================================== */
  function init() {
    loadTheme();
    loadVehicles();
    setupEventListeners();
    renderApp();
    
    // Set default date picker min/defaults
    if (dom.expiryInput) {
      const todayStr = new Date().toISOString().split('T')[0];
      dom.expiryInput.setAttribute('min', '2000-01-01');
    }
    // Auto update countdown every minute
    setInterval(renderApp, 60000);
  }
  function loadVehicles() {
    try {
      const rawData = localStorage.getItem(STORAGE_KEY);
      if (rawData) {
        vehicles = JSON.parse(rawData);
      } else {
        // Seed initial data on first visit
        vehicles = [...INITIAL_DEMO_CARS];
        saveVehicles();
      }
    } catch (e) {
      console.error('Error loading vehicles from localStorage:', e);
      vehicles = [...INITIAL_DEMO_CARS];
    }
  }
  function saveVehicles() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(vehicles));
    } catch (e) {
      console.error('Error saving vehicles to localStorage:', e);
      showToast('Failed to save to browser storage!', 'error');
    }
  }
  function loadTheme() {
    const savedTheme = localStorage.getItem(THEME_KEY) || 'dark';
    document.documentElement.setAttribute('data-theme', savedTheme);
  }
  function toggleTheme() {
    const currentTheme = document.documentElement.getAttribute('data-theme') || 'dark';
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', newTheme);
    localStorage.setItem(THEME_KEY, newTheme);
    showToast(`Switched to ${newTheme} theme`, 'info');
  }
  /* ==========================================
     Calculation & Expiry Helper Logic
     ========================================== */
  function getExpiryStatusInfo(expiryDateStr) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const expiryDate = new Date(expiryDateStr + 'T00:00:00');
    const diffTime = expiryDate.getTime() - today.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    let status = 'active'; // 'active', 'expiring', 'expired'
    let statusLabel = 'Valid / Active';
    let badgeClass = 'badge-active';
    let progressFillClass = 'fill-green';
    if (diffDays < 0) {
      status = 'expired';
      statusLabel = 'EXPIRED';
      badgeClass = 'badge-expired';
      progressFillClass = 'fill-red';
    } else if (diffDays <= 30) {
      status = 'expiring';
      statusLabel = 'EXPIRING SOON';
      badgeClass = 'badge-expiring';
      progressFillClass = 'fill-amber';
    }
    // Time text display
    let timeText = '';
    if (diffDays < 0) {
      const absDays = Math.abs(diffDays);
      timeText = `Expired ${absDays} day${absDays === 1 ? '' : 's'} ago`;
    } else if (diffDays === 0) {
      timeText = 'Expires TODAY!';
    } else if (diffDays === 1) {
      timeText = 'Expires Tomorrow!';
    } else {
      timeText = `${diffDays} Days left`;
    }
    // Progress percentage calculation (Assumes 365 day cycle max visual fill)
    let percentRemaining = 100;
    if (diffDays < 0) {
      percentRemaining = 0;
    } else {
      percentRemaining = Math.min(100, Math.max(5, Math.round((diffDays / 365) * 100)));
    }
    return {
      diffDays,
      status,
      statusLabel,
      badgeClass,
      progressFillClass,
      timeText,
      percentRemaining,
      formattedDate: formatDateDisplay(expiryDateStr)
    };
  }
  function formatDateDisplay(dateStr) {
    if (!dateStr) return '';
    const parts = dateStr.split('-');
    if (parts.length !== 3) return dateStr;
    const dateObj = new Date(dateStr + 'T00:00:00');
    return dateObj.toLocaleDateString(undefined, {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  }
  /* ==========================================
     Render Logic
     ========================================== */
  function renderApp() {
    renderStats();
    renderUrgentAlerts();
    renderVehicleList();
  }
  function renderStats() {
    let total = vehicles.length;
    let active = 0;
    let expiring = 0;
    let expired = 0;
    vehicles.forEach(v => {
      const statusInfo = getExpiryStatusInfo(v.expiryDate);
      if (statusInfo.status === 'active') active++;
      else if (statusInfo.status === 'expiring') expiring++;
      else if (statusInfo.status === 'expired') expired++;
    });
    dom.statTotal.textContent = total;
    dom.statActive.textContent = active;
    dom.statExpiring.textContent = expiring;
    dom.statExpired.textContent = expired;
  }
  function renderUrgentAlerts() {
    dom.urgentAlertsContainer.innerHTML = '';
    const urgentVehicles = vehicles
      .map(v => ({ vehicle: v, info: getExpiryStatusInfo(v.expiryDate) }))
      .filter(item => item.info.status === 'expired' || item.info.status === 'expiring')
      .sort((a, b) => a.info.diffDays - b.info.diffDays);
    if (urgentVehicles.length === 0) return;
    // Render urgent notice card for top urgent car
    urgentVehicles.slice(0, 2).forEach(item => {
      const v = item.vehicle;
      const info = item.info;
      const isExpired = info.status === 'expired';
      const banner = document.createElement('div');
      banner.className = `alert-banner ${isExpired ? 'expired' : 'expiring'}`;
      banner.innerHTML = `
        <div class="alert-content">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
            <line x1="12" y1="9" x2="12" y2="13"/>
            <line x1="12" y1="17" x2="12.01" y2="17"/>
          </svg>
          <div>
            <strong>${v.model} (${v.plate})</strong>: ${info.timeText} (${info.formattedDate}).
          </div>
        </div>
        <button class="btn btn-sm btn-secondary renew-alert-btn" data-id="${v.id}">
          +1 Year Renew
        </button>
      `;
      banner.querySelector('.renew-alert-btn').addEventListener('click', () => {
        quickRenewVehicle(v.id);
      });
      dom.urgentAlertsContainer.appendChild(banner);
    });
  }
  function getFilteredAndSortedVehicles() {
    return vehicles
      .filter(v => {
        const info = getExpiryStatusInfo(v.expiryDate);
        
        // Filter tabs
        if (currentFilter === 'active' && info.status !== 'active') return false;
        if (currentFilter === 'expiring' && info.status !== 'expiring') return false;
        if (currentFilter === 'expired' && info.status !== 'expired') return false;
        // Search query
        if (searchQuery.trim() !== '') {
          const q = searchQuery.toLowerCase().trim();
          const matchPlate = v.plate.toLowerCase().includes(q);
          const matchModel = v.model.toLowerCase().includes(q);
          const matchType = (v.type || '').toLowerCase().includes(q);
          const matchNotes = (v.notes || '').toLowerCase().includes(q);
          const matchColor = (v.color || '').toLowerCase().includes(q);
          return matchPlate || matchModel || matchType || matchNotes || matchColor;
        }
        return true;
      })
      .sort((a, b) => {
        if (currentSort === 'expiry-asc') {
          return new Date(a.expiryDate) - new Date(b.expiryDate);
        } else if (currentSort === 'expiry-desc') {
          return new Date(b.expiryDate) - new Date(a.expiryDate);
        } else if (currentSort === 'model-asc') {
          return a.model.localeCompare(b.model);
        } else if (currentSort === 'plate-asc') {
          return a.plate.localeCompare(b.plate);
        } else if (currentSort === 'added-desc') {
          return (b.createdAt || 0) - (a.createdAt || 0);
        }
        return 0;
      });
  }
  function renderVehicleList() {
    const list = getFilteredAndSortedVehicles();
    dom.vehiclesContainer.innerHTML = '';
    if (currentView === 'list') {
      dom.vehiclesContainer.classList.add('list-view');
    } else {
      dom.vehiclesContainer.classList.remove('list-view');
    }
    if (list.length === 0) {
      dom.vehiclesContainer.classList.add('hidden');
      dom.emptyState.classList.remove('hidden');
      if (vehicles.length === 0) {
        dom.emptyTitle.textContent = 'No vehicles added yet';
        dom.emptyDescription.textContent = 'Start by adding your first vehicle details and registration date.';
      } else {
        dom.emptyTitle.textContent = 'No matching vehicles found';
        dom.emptyDescription.textContent = 'Try adjusting your search query or filter tab.';
      }
      return;
    }
    dom.emptyState.classList.add('hidden');
    dom.vehiclesContainer.classList.remove('hidden');
    list.forEach(vehicle => {
      const card = createVehicleCardElement(vehicle);
      dom.vehiclesContainer.appendChild(card);
    });
  }
  function createVehicleCardElement(v) {
    const info = getExpiryStatusInfo(v.expiryDate);
    const card = document.createElement('div');
    card.className = `car-card status-${info.status}`;
    card.innerHTML = `
      <div class="car-card-header">
        <div class="license-plate">
          <div class="plate-header">REGISTRATION PLATE</div>
          <div class="plate-number">${escapeHtml(v.plate)}</div>
        </div>
        <div class="car-title-row">
          <div>
            <h3 class="car-model-name">${escapeHtml(v.model)}</h3>
            <div style="margin-top: 0.35rem; display: flex; gap: 0.5rem; align-items: center;">
              <span class="type-badge">
                <span class="color-dot" style="background-color: ${escapeHtml(v.color || '#3B82F6')}"></span>
                ${escapeHtml(v.type || 'Vehicle')}
              </span>
            </div>
          </div>
          <span class="status-badge ${info.badgeClass}">
            <span class="pulse-dot"></span>
            ${info.statusLabel}
          </span>
        </div>
      </div>
      <div class="countdown-box">
        <div class="countdown-header">
          <span class="countdown-title">Registration Expiry</span>
          <span class="countdown-date">${info.formattedDate}</span>
        </div>
        <div class="time-left-big ${info.status === 'expired' ? 'text-red' : info.status === 'expiring' ? 'text-amber' : 'text-green'}">
          ${info.timeText}
        </div>
        <div class="progress-bar-bg" title="${info.percentRemaining}% validity window remaining">
          <div class="progress-bar-fill ${info.progressFillClass}" style="width: ${info.percentRemaining}%;"></div>
        </div>
      </div>
      ${v.notes ? `<p style="font-size: 0.82rem; color: var(--text-muted); margin-bottom: 1rem; line-height: 1.4;">${escapeHtml(v.notes)}</p>` : ''}
      <div class="car-card-actions">
        <button class="btn btn-secondary btn-sm renew-btn" title="Extend expiry date by 1 Year">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/></svg>
          <span>+1 Year</span>
        </button>
        <div class="action-btn-group">
          <button class="btn btn-secondary btn-sm edit-btn" title="Edit vehicle details">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
            <span>Edit</span>
          </button>
          <button class="btn btn-secondary btn-sm delete-btn" title="Delete vehicle" style="color: var(--status-red);">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
          </button>
        </div>
      </div>
    `;
    // Action Listener Binds
    card.querySelector('.renew-btn').addEventListener('click', () => quickRenewVehicle(v.id));
    card.querySelector('.edit-btn').addEventListener('click', () => openModalForEdit(v.id));
    card.querySelector('.delete-btn').addEventListener('click', () => deleteVehicle(v.id));
    return card;
  }
  /* ==========================================
     CRUD Operations
     ========================================== */
  function openModalForAdd() {
    dom.carIdInput.value = '';
    dom.carForm.reset();
    dom.modalTitle.textContent = 'Add New Vehicle';
    dom.colorPicker.value = '#3B82F6';
    dom.colorInput.value = 'Blue';
    dom.previewPlateText.textContent = 'ABC-1234';
    clearFormErrors();
    dom.carModal.classList.remove('hidden');
    dom.plateInput.focus();
  }
  function openModalForEdit(id) {
    const v = vehicles.find(item => item.id === id);
    if (!v) return;
    clearFormErrors();
    dom.carIdInput.value = v.id;
    dom.plateInput.value = v.plate;
    dom.modelInput.value = v.model;
    dom.expiryInput.value = v.expiryDate;
    dom.typeInput.value = v.type || 'Sedan';
    dom.colorPicker.value = v.color || '#3B82F6';
    dom.colorInput.value = v.colorName || 'Blue';
    dom.insuranceInput.value = v.insuranceExpiry || '';
    dom.notesInput.value = v.notes || '';
    dom.previewPlateText.textContent = v.plate.toUpperCase();
    dom.modalTitle.textContent = 'Edit Vehicle Details';
    dom.carModal.classList.remove('hidden');
  }
  function closeModal() {
    dom.carModal.classList.add('hidden');
    clearFormErrors();
  }
  function clearFormErrors() {
    document.querySelectorAll('.error-msg').forEach(el => el.classList.remove('visible'));
    document.querySelectorAll('.form-input').forEach(el => el.classList.remove('is-invalid'));
  }
  function handleFormSubmit(e) {
    e.preventDefault();
    clearFormErrors();
    const plate = dom.plateInput.value.trim();
    const model = dom.modelInput.value.trim();
    const expiryDate = dom.expiryInput.value;
    const type = dom.typeInput.value;
    const color = dom.colorPicker.value;
    const colorName = dom.colorInput.value.trim();
    const insuranceExpiry = dom.insuranceInput.value;
    const notes = dom.notesInput.value.trim();
    const existingId = dom.carIdInput.value;
    let isValid = true;
    if (!plate) {
      showFieldError('plate', 'Please enter a valid license plate number');
      isValid = false;
    }
    if (!model) {
      showFieldError('model', 'Please enter car make & model');
      isValid = false;
    }
    if (!expiryDate) {
      showFieldError('expiry', 'Please select a valid expiry date');
      isValid = false;
    }
    if (!isValid) return;
    if (existingId) {
      // Update existing
      const index = vehicles.findIndex(v => v.id === existingId);
      if (index !== -1) {
        vehicles[index] = {
          ...vehicles[index],
          plate: plate.toUpperCase(),
          model,
          expiryDate,
          type,
          color,
          colorName,
          insuranceExpiry,
          notes,
          updatedAt: Date.now()
        };
        showToast(`Updated ${model} (${plate.toUpperCase()}) successfully`, 'success');
      }
    } else {
      // Create new vehicle
      const newVehicle = {
        id: 'car_' + Date.now(),
        plate: plate.toUpperCase(),
        model,
        expiryDate,
        type,
        color,
        colorName,
        insuranceExpiry,
        notes,
        createdAt: Date.now()
      };
      vehicles.push(newVehicle);
      showToast(`Added ${model} (${plate.toUpperCase()}) to vault`, 'success');
    }
    saveVehicles();
    closeModal();
    renderApp();
  }
  function showFieldError(fieldKey, message) {
    const inputEl = document.getElementById(`${fieldKey}-input`);
    const errorEl = document.getElementById(`${fieldKey}-error`);
    if (inputEl) inputEl.classList.add('is-invalid');
    if (errorEl) {
      errorEl.textContent = message;
      errorEl.classList.add('visible');
    }
  }
  function quickRenewVehicle(id) {
    const v = vehicles.find(item => item.id === id);
    if (!v) return;
    // Calculate current expiry + 1 year
    const currentDateObj = new Date(v.expiryDate + 'T00:00:00');
    // If already expired, start 1 year from today!
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    let baseDate = currentDateObj < today ? today : currentDateObj;
    baseDate.setFullYear(baseDate.getFullYear() + 1);
    const newExpiryStr = baseDate.toISOString().split('T')[0];
    v.expiryDate = newExpiryStr;
    v.updatedAt = Date.now();
    saveVehicles();
    renderApp();
    showToast(`Renewed registration for ${v.model} (+1 Year to ${formatDateDisplay(newExpiryStr)})!`, 'success');
  }
  function deleteVehicle(id) {
    const v = vehicles.find(item => item.id === id);
    if (!v) return;
    if (confirm(`Are you sure you want to delete ${v.model} (${v.plate}) from your tracker?`)) {
      vehicles = vehicles.filter(item => item.id !== id);
      saveVehicles();
      renderApp();
      showToast(`Deleted vehicle ${v.plate}`, 'info');
    }
  }
  /* ==========================================
     Import / Export
     ========================================== */
  function exportVehiclesJSON() {
    if (vehicles.length === 0) {
      showToast('No vehicle records to export!', 'warning');
      return;
    }
    const jsonStr = JSON.stringify(vehicles, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `autovault_vehicles_backup_${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('Exported vehicles data successfully!', 'success');
  }
  function importVehiclesJSON(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (event) {
      try {
        const importedData = JSON.parse(event.target.result);
        if (!Array.isArray(importedData)) {
          throw new Error('Invalid file format. Expected array of vehicles.');
        }
        // Simple validation
        const validVehicles = importedData.filter(item => item.plate && item.model && item.expiryDate);
        if (validVehicles.length === 0) {
          showToast('No valid vehicle records found in imported file.', 'error');
          return;
        }
        vehicles = validVehicles;
        saveVehicles();
        renderApp();
        showToast(`Successfully imported ${validVehicles.length} vehicles!`, 'success');
      } catch (err) {
        console.error(err);
        showToast('Failed to import JSON file. Please check file structure.', 'error');
      }
    };
    reader.readAsText(file);
    e.target.value = ''; // Reset input
  }
  /* ==========================================
     Toast Notifications
     ========================================== */
  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    let iconSvg = '';
    if (type === 'success') {
      iconSvg = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg>`;
    } else if (type === 'error') {
      iconSvg = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>`;
    } else {
      iconSvg = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>`;
    }
    toast.innerHTML = `${iconSvg} <span>${escapeHtml(message)}</span>`;
    dom.toastContainer.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(100%) scale(0.9)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 3500);
  }
  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }
  /* ==========================================
     Event Listeners Binds
     ========================================== */
  function setupEventListeners() {
    // Theme
    dom.themeToggleBtn.addEventListener('click', toggleTheme);
    // Export / Import
    dom.exportBtn.addEventListener('click', exportVehiclesJSON);
    dom.importBtn.addEventListener('click', () => dom.importFileInput.click());
    dom.importFileInput.addEventListener('change', importVehiclesJSON);
    // Add Car Modal
    dom.addCarBtn.addEventListener('click', openModalForAdd);
    dom.emptyActionBtn.addEventListener('click', openModalForAdd);
    dom.closeModalBtn.addEventListener('click', closeModal);
    dom.cancelModalBtn.addEventListener('click', closeModal);
    dom.carForm.addEventListener('submit', handleFormSubmit);
    // Live Plate Preview Typing Sync
    dom.plateInput.addEventListener('input', (e) => {
      const val = e.target.value.trim();
      dom.previewPlateText.textContent = val ? val.toUpperCase() : 'ABC-1234';
    });
    // Color Swatch Sync
    dom.colorPicker.addEventListener('input', (e) => {
      dom.colorInput.value = e.target.value;
    });
    // Search Input
    dom.searchInput.addEventListener('input', (e) => {
      searchQuery = e.target.value;
      if (searchQuery.trim() !== '') {
        dom.clearSearchBtn.classList.remove('hidden');
      } else {
        dom.clearSearchBtn.classList.add('hidden');
      }
      renderVehicleList();
    });
    dom.clearSearchBtn.addEventListener('click', () => {
      dom.searchInput.value = '';
      searchQuery = '';
      dom.clearSearchBtn.classList.add('hidden');
      renderVehicleList();
    });
    // Filter Tabs
    dom.filterTabs.forEach(tab => {
      tab.addEventListener('click', () => {
        dom.filterTabs.forEach(t => {
          t.classList.remove('active');
          t.setAttribute('aria-selected', 'false');
        });
        tab.classList.add('active');
        tab.setAttribute('aria-selected', 'true');
        currentFilter = tab.dataset.filter;
        renderVehicleList();
      });
    });
    // Sort Select
    dom.sortSelect.addEventListener('change', (e) => {
      currentSort = e.target.value;
      renderVehicleList();
    });
    // View Toggles
    dom.viewGridBtn.addEventListener('click', () => {
      currentView = 'grid';
      dom.viewGridBtn.classList.add('active');
      dom.viewListBtn.classList.remove('active');
      renderVehicleList();
    });
    dom.viewListBtn.addEventListener('click', () => {
      currentView = 'list';
      dom.viewListBtn.classList.add('active');
      dom.viewGridBtn.classList.remove('active');
      renderVehicleList();
    });
    // Escape Key Modal Close
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !dom.carModal.classList.contains('hidden')) {
        closeModal();
      }
    });
    // Backdrop Click Close
    dom.carModal.addEventListener('click', (e) => {
      if (e.target === dom.carModal) {
        closeModal();
      }
    });
  }
  // Run app on DOMReady
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();