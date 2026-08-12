/**
 * AutoVault - Car Registration & Expiry Tracker
 * Core Application Engine with Excel View, Excel Export & Import, and Arabic Language Support
 */
(function () {
  'use strict';

  // LocalStorage Keys
  const STORAGE_KEY = 'autovault_vehicles_v2';
  const THEME_KEY = 'autovault_theme_pref';
  const LANG_KEY = 'autovault_lang_pref';

  // Initial Vehicles Data (Empty - users add their own vehicles)
  const INITIAL_DEMO_CARS = [];

  // Popular Car Brands & Model Dictionary
  const CAR_BRANDS_DATA = {
    'Toyota': ['Camry', 'Corolla', 'RAV4', 'Land Cruiser', 'Hilux', 'Prado', 'Yaris', 'Avalon', 'Highlander', 'Fortuner', 'Supra', 'Crown', 'Sienna', 'Tacoma', 'Tundra', 'C-HR', 'Corolla Cross', '4Runner'],
    'Nissan': ['Patrol', 'Altima', 'Sunny', 'Maxima', 'Pathfinder', 'X-Trail', 'Sentra', 'Rogue', 'Armada', 'GT-R', 'Kicks', 'Murano', 'Frontier', 'Z', 'Urvan'],
    'Honda': ['Civic', 'Accord', 'CR-V', 'Pilot', 'HR-V', 'City', 'Odyssey', 'Insight', 'Passport', 'Ridgeline'],
    'Hyundai': ['Elantra', 'Sonata', 'Tucson', 'Santa Fe', 'Accent', 'Palisade', 'Kona', 'Venue', 'Azera', 'Creta', 'Ioniq 5', 'Staria'],
    'Kia': ['K5', 'Sportage', 'Sorento', 'Cerato', 'Rio', 'Telluride', 'Stinger', 'Carnival', 'Seltos', 'Picanto', 'Pegas', 'EV6', 'Optima'],
    'Ford': ['Mustang', 'Explorer', 'F-150', 'Expedition', 'Ranger', 'Edge', 'Escape', 'Bronco', 'Taurus', 'Everest', 'Territory'],
    'Chevrolet': ['Tahoe', 'Suburban', 'Silverado', 'Camaro', 'Traverse', 'Malibu', 'Equinox', 'Trailblazer', 'Captiva', 'Corvette', 'Groove'],
    'Lexus': ['ES', 'LS', 'RX', 'LX', 'GX', 'IS', 'NX', 'UX', 'LC', 'LM'],
    'Mercedes-Benz': ['C-Class', 'E-Class', 'S-Class', 'GLE', 'GLS', 'G-Class', 'CLA', 'GLA', 'GLC', 'AMG GT', 'EQE', 'EQS'],
    'BMW': ['3 Series', '5 Series', '7 Series', 'X3', 'X5', 'X6', 'X7', '4 Series', 'M3', 'M5', 'i4', 'iX'],
    'Audi': ['A4', 'A6', 'A8', 'Q5', 'Q7', 'Q8', 'Q3', 'e-tron', 'RS6', 'RS7'],
    'GMC': ['Yukon', 'Sierra', 'Acadia', 'Terrain', 'Hummer EV'],
    'Dodge': ['Charger', 'Challenger', 'Durango', 'RAM 1500'],
    'Jeep': ['Grand Cherokee', 'Wrangler', 'Cherokee', 'Compass', 'Gladiator', 'Renegade'],
    'Volkswagen': ['Golf', 'Passat', 'Tiguan', 'Touareg', 'Jetta', 'Teramont', 'ID.4', 'T-Roc'],
    'Mazda': ['Mazda3', 'Mazda6', 'CX-5', 'CX-9', 'CX-30', 'CX-50', 'MX-5 Miata', 'CX-60'],
    'Mitsubishi': ['Pajero', 'Lancer', 'Outlander', 'ASX', 'Eclipse Cross', 'Montero Sport', 'Attrage', 'Xpander'],
    'Land Rover': ['Range Rover', 'Range Rover Sport', 'Defender', 'Discovery', 'Evoque', 'Velar'],
    'Porsche': ['911', 'Cayenne', 'Macan', 'Panamera', 'Taycan', 'Boxster', 'Cayman'],
    'Tesla': ['Model 3', 'Model Y', 'Model S', 'Model X', 'Cybertruck'],
    'Subaru': ['Outback', 'Forester', 'Impreza', 'WRX', 'Crosstrek', 'BRZ'],
    'Genesis': ['G70', 'G80', 'G90', 'GV70', 'GV80'],
    'Suzuki': ['Swift', 'Jimny', 'Vitara', 'Baleno', 'Ertiga', 'Grand Vitara', 'Dzire'],
    'MG': ['MG GT', 'MG ZS', 'MG RX5', 'MG 6', 'MG HS', 'MG 5', 'MG ONE', 'MG Whale'],
    'Geely': ['Coolray', 'Monjaro', 'Tugella', 'Emgrand', 'Azkarra', 'Geometry C', 'Starray'],
    'Changan': ['CS75 Plus', 'CS35 Plus', 'CS95', 'Eado', 'UNI-K', 'UNI-V', 'UNI-T', 'Alsvin'],
    'BYD': ['Song Plus', 'Atto 3', 'Han', 'Tang', 'Seal', 'Dolphin', 'Yuan Plus', 'Qin Plus']
  };

  const FEATURED_BRANDS = ['Toyota', 'Nissan', 'Honda', 'Hyundai', 'Kia', 'Ford', 'Chevrolet', 'Mercedes-Benz', 'BMW', 'Lexus', 'GMC', 'Dodge', 'Jeep', 'Mazda'];

  // Internationalization (i18n) Dictionary
  const i18n = {
    en: {
      brandTitle: 'AutoVault',
      brandSubtitle: 'Vehicle Registration & Expiry Hub',
      langBtn: 'العربية',
      exportExcelBtn: 'Export Excel (.xlsx)',
      importExcelBtn: 'Import Excel',
      exportJsonBtn: 'Export JSON',
      importJsonBtn: 'Import JSON',
      addCarBtn: '+ Add Vehicle',
      statTotal: 'Total Vehicles',
      statActive: 'Active / Valid',
      statExpiring: 'Expiring Soon (≤30d)',
      statExpired: 'Expired',
      tabAll: 'All',
      tabActive: 'Active',
      tabExpiring: 'Expiring Soon',
      tabExpired: 'Expired',
      searchPlaceholder: 'Search by model, vehicle no, plate, driver, registration...',
      sortExpiryAsc: 'Expiry: Soonest First',
      sortExpiryDesc: 'Expiry: Furthest First',
      sortModelAsc: 'Model: A to Z',
      sortPlateAsc: 'Plate: A to Z',
      sortVehicleNoAsc: 'Vehicle No: Ascending',
      sortAddedDesc: 'Recently Added',
      colNo: 'No.',
      colVehicleNo: 'Vehicle No',
      colVehicleType: 'Vehicle Type',
      colModel: 'Model',
      colPlateNo: 'Plate No',
      colDriverName: 'Driver Name',
      colRegistrationNo: 'Registration No',
      colIssueDate: 'Issue Date',
      colExpiryDate: 'Expiry Date',
      colDaysRemaining: 'Days Remaining',
      colStatus: 'Status',
      colRemarks: 'Remarks',
      colActions: 'Actions',
      statusValid: 'Valid',
      statusExpiring: 'Expiring Soon',
      statusExpired: 'Expired',
      modalTitleAdd: 'Add New Vehicle',
      modalTitleEdit: 'Edit Vehicle Details',
      lblVehicleNo: 'Vehicle No',
      lblVehicleType: 'Vehicle Type / Make',
      lblModel: 'Model',
      lblPlate: 'Plate No',
      lblDriver: 'Driver Name',
      lblRegistration: 'Registration No',
      lblIssueDate: 'Issue Date',
      lblExpiryDate: 'Expiry Date',
      lblColor: 'Vehicle Color',
      lblRemarks: 'Remarks',
      lblNotes: 'Additional Notes / VIN (Optional)',
      lblLivePreview: 'Live Plate Preview',
      lblPopularBrands: 'Popular Brands:',
      modelPlaceholderSelectBrand: 'Select brand first...',
      btnCancel: 'Cancel',
      btnSave: 'Save Vehicle',
      emptyTitle: 'No vehicles found',
      emptyDescription: 'Get started by clicking "+ Add Vehicle" or clearing your search filter.',
      emptyAction: '+ Add Your First Car',
      footerText: '© 2026 AutoVault - Vehicle Registration & Expiry System. Local Browser Storage Protected.',
      renewSuccess: 'Renewed registration (+1 Year)!',
      deleteConfirm: 'Are you sure you want to delete vehicle',
      deleteSuccess: 'Vehicle deleted successfully',
      addSuccess: 'Vehicle added successfully',
      updateSuccess: 'Vehicle updated successfully'
    },
    ar: {
      brandTitle: 'أوتو فولت',
      brandSubtitle: 'مركز متابعة وتجديد استمارات المركبات',
      langBtn: 'English',
      exportExcelBtn: 'تصدير اكسل (.xlsx)',
      importExcelBtn: 'استيراد اكسل',
      exportJsonBtn: 'تصدير JSON',
      importJsonBtn: 'استيراد JSON',
      addCarBtn: '+ إضافة مركبة',
      statTotal: 'إجمالي المركبات',
      statActive: 'صالح / نشط',
      statExpiring: 'ينتهي قريباً (≤30 يوم)',
      statExpired: 'منتهي الصلاحية',
      tabAll: 'الكل',
      tabActive: 'المركبات الصالحة',
      tabExpiring: 'تنتهي قريباً',
      tabExpired: 'المنتهية',
      searchPlaceholder: 'ابحث بالموديل، رقم المركبة، رقم اللوحة، اسم السائق، رقم التسجيل...',
      sortExpiryAsc: 'تاريخ الانتهاء: الأقرب أولاً',
      sortExpiryDesc: 'تاريخ الانتهاء: الأبعد أولاً',
      sortModelAsc: 'الموديل: أ إلى ي',
      sortPlateAsc: 'اللوحة: أ إلى ي',
      sortVehicleNoAsc: 'رقم المركبة: تصاعدي',
      sortAddedDesc: 'المضافة حديثاً',
      colNo: 'الرقم',
      colVehicleNo: 'رقم المركبة',
      colVehicleType: 'نوع المركبة',
      colModel: 'الموديل',
      colPlateNo: 'رقم اللوحة',
      colDriverName: 'اسم السائق',
      colRegistrationNo: 'رقم التسجيل',
      colIssueDate: 'تاريخ الإصدار',
      colExpiryDate: 'تاريخ الانتهاء',
      colDaysRemaining: 'الأيام المتبقية',
      colStatus: 'الحالة',
      colRemarks: 'ملاحظات',
      colActions: 'الإجراءات',
      statusValid: 'صالح',
      statusExpiring: 'ينتهي قريباً',
      statusExpired: 'منتهي الصلاحية',
      modalTitleAdd: 'إضافة مركبة جديدة',
      modalTitleEdit: 'تعديل بيانات المركبة',
      lblVehicleNo: 'رقم المركبة',
      lblVehicleType: 'نوع المركبة / الماركة',
      lblModel: 'الموديل',
      lblPlate: 'رقم اللوحة',
      lblDriver: 'اسم السائق',
      lblRegistration: 'رقم التسجيل',
      lblIssueDate: 'تاريخ الإصدار',
      lblExpiryDate: 'تاريخ الانتهاء',
      lblColor: 'لون المركبة',
      lblRemarks: 'ملاحظات',
      lblNotes: 'ملاحظات إضافية / رقم الهيكل (اختياري)',
      lblLivePreview: 'معاينة اللوحة المباشرة',
      lblPopularBrands: 'الماركات الشهيرة:',
      modelPlaceholderSelectBrand: 'اختر الماركة أولاً...',
      btnCancel: 'إلغاء',
      btnSave: 'حفظ المركبة',
      emptyTitle: 'لم يتم العثور على مركبات',
      emptyDescription: 'ابدأ بالنقر على "+ إضافة مركبة" أو قم بمسح فلتر البحث.',
      emptyAction: '+ أضف سيارتك الأولى',
      footerText: '© 2026 أوتو فولت - نظام متابعة وتجديد استمارات السيارات. محفوط في المتصفح المحلي.',
      renewSuccess: 'تم تجديد الاستمارة لممدة سنة كاملة!',
      deleteConfirm: 'هل أنت تأكد من حذف المركبة',
      deleteSuccess: 'تم حذف المركبة بنجاح',
      addSuccess: 'تمت إضافة المركبة بنجاح',
      updateSuccess: 'تم تحديث بيانات المركبة بنجاح'
    }
  };

  // App State
  let vehicles = [];
  let currentFilter = 'all';
  let searchQuery = '';
  let currentSort = 'expiry-asc';
  let currentView = 'excel'; // 'excel', 'grid', 'list'
  let currentLang = 'en'; // 'en', 'ar'

  // DOM Selectors
  const dom = {
    // Header & Language
    themeToggleBtn: document.getElementById('theme-toggle-btn'),
    langToggleBtn: document.getElementById('lang-toggle-btn'),
    txtLangBtn: document.getElementById('txt-lang-btn'),
    
    // Excel Export & Import
    exportExcelBtn: document.getElementById('export-excel-btn'),
    txtExportExcelBtn: document.getElementById('txt-export-excel-btn'),
    importExcelBtn: document.getElementById('import-excel-btn'),
    txtImportExcelBtn: document.getElementById('txt-import-excel-btn'),
    importExcelFileInput: document.getElementById('import-excel-file-input'),

    // JSON Export & Import
    exportBtn: document.getElementById('export-btn'),
    txtExportJsonBtn: document.getElementById('txt-export-json-btn'),
    importBtn: document.getElementById('import-btn'),
    txtImportJsonBtn: document.getElementById('txt-import-json-btn'),
    importFileInput: document.getElementById('import-file-input'),

    addCarBtn: document.getElementById('add-car-btn'),
    txtAddCarBtn: document.getElementById('txt-add-car-btn'),
    txtBrandTitle: document.getElementById('txt-brand-title'),
    txtBrandSubtitle: document.getElementById('txt-brand-subtitle'),

    // Stats Labels & Values
    txtStatTotal: document.getElementById('txt-stat-total'),
    txtStatActive: document.getElementById('txt-stat-active'),
    txtStatExpiring: document.getElementById('txt-stat-expiring'),
    txtStatExpired: document.getElementById('txt-stat-expired'),
    statTotal: document.getElementById('stat-total-count'),
    statActive: document.getElementById('stat-active-count'),
    statExpiring: document.getElementById('stat-expiring-count'),
    statExpired: document.getElementById('stat-expired-count'),

    // Alerts Container
    urgentAlertsContainer: document.getElementById('urgent-alerts-container'),

    // Toolbar Controls
    searchInput: document.getElementById('search-input'),
    clearSearchBtn: document.getElementById('clear-search-btn'),
    filterTabs: document.querySelectorAll('.filter-tab'),
    tabAll: document.getElementById('tab-all'),
    tabActive: document.getElementById('tab-active'),
    tabExpiring: document.getElementById('tab-expiring'),
    tabExpired: document.getElementById('tab-expired'),
    sortSelect: document.getElementById('sort-select'),
    optSortExpiryAsc: document.getElementById('opt-sort-expiry-asc'),
    optSortExpiryDesc: document.getElementById('opt-sort-expiry-desc'),
    optSortModelAsc: document.getElementById('opt-sort-model-asc'),
    optSortPlateAsc: document.getElementById('opt-sort-plate-asc'),
    optSortVehicleNoAsc: document.getElementById('opt-sort-vehicleno-asc'),
    optSortAddedDesc: document.getElementById('opt-sort-added-desc'),

    // View Toggles
    viewExcelBtn: document.getElementById('view-excel-btn'),
    viewGridBtn: document.getElementById('view-grid-btn'),
    viewListBtn: document.getElementById('view-list-btn'),

    // Main Containers & Empty State
    vehiclesContainer: document.getElementById('vehicles-container'),
    excelViewContainer: document.getElementById('excel-view-container'),
    emptyState: document.getElementById('empty-state'),
    emptyTitle: document.getElementById('empty-title'),
    emptyDescription: document.getElementById('empty-description'),
    emptyActionBtn: document.getElementById('empty-action-btn'),

    // Modal & Form Inputs
    carModal: document.getElementById('car-modal'),
    modalTitle: document.getElementById('modal-title'),
    closeModalBtn: document.getElementById('close-modal-btn'),
    cancelModalBtn: document.getElementById('cancel-modal-btn'),
    txtCancelBtn: document.getElementById('txt-cancel-btn'),
    saveCarBtn: document.getElementById('save-car-btn'),
    txtSaveBtn: document.getElementById('txt-save-btn'),
    carForm: document.getElementById('car-form'),

    carIdInput: document.getElementById('car-id'),
    vehicleNoInput: document.getElementById('vehicleno-input'),
    typeInput: document.getElementById('type-input'),
    modelInput: document.getElementById('model-input'),
    plateInput: document.getElementById('plate-input'),
    driverInput: document.getElementById('driver-input'),
    registrationInput: document.getElementById('registration-input'),
    issueInput: document.getElementById('issue-input'),
    expiryInput: document.getElementById('expiry-input'),
    colorPicker: document.getElementById('color-picker'),
    colorInput: document.getElementById('color-input'),
    remarksInput: document.getElementById('remarks-input'),
    notesInput: document.getElementById('notes-input'),

    // Brand & Model Dynamic Lists
    brandList: document.getElementById('brand-list'),
    modelList: document.getElementById('model-list'),
    brandChips: document.getElementById('brand-chips'),
    txtPopularBrands: document.getElementById('txt-popular-brands'),

    // Form Labels
    lblVehicleNo: document.getElementById('lbl-vehicleno'),
    lblVehicleType: document.getElementById('lbl-vehicletype'),
    lblModel: document.getElementById('lbl-model'),
    lblPlate: document.getElementById('lbl-plate'),
    lblDriver: document.getElementById('lbl-driver'),
    lblRegistration: document.getElementById('lbl-registration'),
    lblIssueDate: document.getElementById('lbl-issuedate'),
    lblExpiryDate: document.getElementById('lbl-expirydate'),
    lblColor: document.getElementById('lbl-color'),
    lblRemarks: document.getElementById('lbl-remarks'),
    lblNotes: document.getElementById('lbl-notes'),
    txtLivePreviewLabel: document.getElementById('txt-live-preview-label'),
    previewPlateText: document.getElementById('preview-plate-text'),

    // Footer & Toast
    txtFooter: document.getElementById('txt-footer'),
    toastContainer: document.getElementById('toast-container')
  };

  /* ==========================================
     Brand & Model Selector Helpers
     ========================================== */
  function populateBrandDatalist() {
    if (!dom.brandList) return;
    dom.brandList.innerHTML = '';
    Object.keys(CAR_BRANDS_DATA).sort().forEach(brand => {
      const option = document.createElement('option');
      option.value = brand;
      dom.brandList.appendChild(option);
    });
  }

  function renderBrandChips(selectedBrand = '') {
    if (!dom.brandChips) return;
    dom.brandChips.innerHTML = '';
    const normSelected = (selectedBrand || '').trim().toLowerCase();

    FEATURED_BRANDS.forEach(brand => {
      const chip = document.createElement('button');
      chip.type = 'button';
      chip.className = `brand-chip ${normSelected === brand.toLowerCase() ? 'active' : ''}`;
      chip.textContent = brand;
      chip.addEventListener('click', () => {
        dom.typeInput.value = brand;
        updateBrandAndModelOptions(brand);
        dom.modelInput.focus();
      });
      dom.brandChips.appendChild(chip);
    });
  }

  function findMatchingBrandKey(inputBrand) {
    if (!inputBrand) return null;
    const cleanInput = inputBrand.trim().toLowerCase();
    return Object.keys(CAR_BRANDS_DATA).find(
      key => key.toLowerCase() === cleanInput
    ) || null;
  }

  function updateBrandAndModelOptions(brandName) {
    renderBrandChips(brandName);
    const matchedBrand = findMatchingBrandKey(brandName);
    const t = i18n[currentLang];

    if (!dom.modelList) return;
    dom.modelList.innerHTML = '';

    if (matchedBrand) {
      const models = CAR_BRANDS_DATA[matchedBrand];
      models.forEach(model => {
        const option = document.createElement('option');
        option.value = model;
        dom.modelList.appendChild(option);
      });

      const sampleModels = models.slice(0, 3).join(', ');
      const isAr = currentLang === 'ar';
      dom.modelInput.placeholder = isAr
        ? `اختر موديل ${matchedBrand} (مثل ${sampleModels}...)`
        : `Select ${matchedBrand} model (e.g. ${sampleModels}...)`;
    } else if (brandName && brandName.trim() !== '') {
      const isAr = currentLang === 'ar';
      dom.modelInput.placeholder = isAr
        ? `اكتب موديل سيارة ${brandName.trim()}...`
        : `Type ${brandName.trim()} model...`;
    } else {
      dom.modelInput.placeholder = t.modelPlaceholderSelectBrand || 'Select brand first...';
    }
  }

  /* ==========================================
     Initialization & Storage
     ========================================== */
  function init() {
    loadLanguage();
    loadTheme();
    loadVehicles();
    populateBrandDatalist();
    setupEventListeners();
    applyLanguage();
    renderApp();

    // Auto update countdown every minute
    setInterval(renderApp, 60000);
  }

  function loadLanguage() {
    const savedLang = localStorage.getItem(LANG_KEY);
    if (savedLang && (savedLang === 'en' || savedLang === 'ar')) {
      currentLang = savedLang;
    } else {
      currentLang = 'en';
    }
  }

  function toggleLanguage() {
    currentLang = currentLang === 'en' ? 'ar' : 'en';
    localStorage.setItem(LANG_KEY, currentLang);
    applyLanguage();
    renderApp();
    showToast(currentLang === 'ar' ? 'تم تغيير اللغة إلى العربية' : 'Switched language to English', 'info');
  }

  function applyLanguage() {
    const t = i18n[currentLang];
    const isAr = currentLang === 'ar';

    // HTML lang and direction
    document.documentElement.setAttribute('lang', currentLang);
    document.documentElement.setAttribute('dir', isAr ? 'rtl' : 'ltr');

    // Header & Buttons
    dom.txtBrandTitle.textContent = t.brandTitle;
    dom.txtBrandSubtitle.textContent = t.brandSubtitle;
    dom.txtLangBtn.textContent = t.langBtn;
    dom.txtExportExcelBtn.textContent = t.exportExcelBtn;
    dom.txtImportExcelBtn.textContent = t.importExcelBtn;
    dom.txtExportJsonBtn.textContent = t.exportJsonBtn;
    dom.txtImportJsonBtn.textContent = t.importJsonBtn;
    dom.txtAddCarBtn.textContent = t.addCarBtn;

    // Stats
    dom.txtStatTotal.textContent = t.statTotal;
    dom.txtStatActive.textContent = t.statActive;
    dom.txtStatExpiring.textContent = t.statExpiring;
    dom.txtStatExpired.textContent = t.statExpired;

    // Filter Tabs
    dom.tabAll.textContent = t.tabAll;
    dom.tabActive.textContent = t.tabActive;
    dom.tabExpiring.textContent = t.tabExpiring;
    dom.tabExpired.textContent = t.tabExpired;

    // Search & Sort
    dom.searchInput.placeholder = t.searchPlaceholder;
    dom.optSortExpiryAsc.textContent = t.sortExpiryAsc;
    dom.optSortExpiryDesc.textContent = t.sortExpiryDesc;
    dom.optSortModelAsc.textContent = t.sortModelAsc;
    dom.optSortPlateAsc.textContent = t.sortPlateAsc;
    dom.optSortVehicleNoAsc.textContent = t.sortVehicleNoAsc;
    dom.optSortAddedDesc.textContent = t.sortAddedDesc;

    // Modal Form Labels
    if (dom.lblVehicleNo.childNodes[0]) dom.lblVehicleNo.childNodes[0].nodeValue = t.lblVehicleNo + ' ';
    if (dom.lblVehicleType.childNodes[0]) dom.lblVehicleType.childNodes[0].nodeValue = t.lblVehicleType + ' ';
    if (dom.lblModel.childNodes[0]) dom.lblModel.childNodes[0].nodeValue = t.lblModel + ' ';
    if (dom.lblPlate.childNodes[0]) dom.lblPlate.childNodes[0].nodeValue = t.lblPlate + ' ';
    if (dom.lblDriver.childNodes[0]) dom.lblDriver.childNodes[0].nodeValue = t.lblDriver + ' ';
    if (dom.lblRegistration.childNodes[0]) dom.lblRegistration.childNodes[0].nodeValue = t.lblRegistration + ' ';
    if (dom.lblIssueDate.childNodes[0]) dom.lblIssueDate.childNodes[0].nodeValue = t.lblIssueDate + ' ';
    if (dom.lblExpiryDate.childNodes[0]) dom.lblExpiryDate.childNodes[0].nodeValue = t.lblExpiryDate + ' ';
    
    if (dom.txtPopularBrands) dom.txtPopularBrands.textContent = t.lblPopularBrands;
    dom.lblColor.textContent = t.lblColor;
    dom.lblRemarks.textContent = t.lblRemarks;
    dom.lblNotes.textContent = t.lblNotes;
    dom.txtLivePreviewLabel.textContent = t.lblLivePreview;
    dom.txtCancelBtn.textContent = t.btnCancel;
    dom.txtSaveBtn.textContent = t.btnSave;

    // Refresh model options placeholder for current language
    if (dom.typeInput) {
      updateBrandAndModelOptions(dom.typeInput.value);
    }

    // Footer
    dom.txtFooter.textContent = t.footerText;
  }

  function loadVehicles() {
    try {
      const rawData = localStorage.getItem(STORAGE_KEY);
      if (rawData) {
        vehicles = JSON.parse(rawData);
      } else {
        const legacyData = localStorage.getItem('autovault_vehicles_v1');
        if (legacyData) {
          const parsed = JSON.parse(legacyData);
          vehicles = parsed.map((v, i) => ({
            id: v.id || 'car_' + Date.now() + i,
            vehicleNo: String(i + 1).padStart(3, '0'),
            type: v.type || 'Sedan',
            model: v.model || 'Vehicle',
            plate: v.plate || '12345',
            driverName: 'Driver ' + (i + 1),
            registrationNo: String(45879 + i),
            issueDate: '2026-01-01',
            expiryDate: v.expiryDate || '2026-12-31',
            remarks: '—',
            color: v.color || '#3B82F6',
            notes: v.notes || '',
            createdAt: v.createdAt || Date.now()
          }));
        } else {
          vehicles = [];
        }
      }

      // Filter out any demo vehicles so user starts with a clean slate
      const initialLen = vehicles.length;
      vehicles = vehicles.filter(v => !v.id || !v.id.startsWith('car_demo_'));
      if (vehicles.length !== initialLen || !rawData) {
        saveVehicles();
      }
    } catch (e) {
      console.error('Error loading vehicles:', e);
      vehicles = [];
    }
  }

  function saveVehicles() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(vehicles));
    } catch (e) {
      console.error('Error saving vehicles:', e);
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

    let expiryDate = new Date(expiryDateStr + 'T00:00:00');
    if (isNaN(expiryDate.getTime())) {
      expiryDate = new Date();
    }

    const diffTime = expiryDate.getTime() - today.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    const t = i18n[currentLang];
    let status = 'active'; // 'active', 'expiring', 'expired'
    let statusLabel = t.statusValid;
    let badgeClass = 'badge-active';
    let progressFillClass = 'fill-green';

    if (diffDays < 0) {
      status = 'expired';
      statusLabel = t.statusExpired;
      badgeClass = 'badge-expired';
      progressFillClass = 'fill-red';
    } else if (diffDays <= 30) {
      status = 'expiring';
      statusLabel = t.statusExpiring;
      badgeClass = 'badge-expiring';
      progressFillClass = 'fill-amber';
    }

    let timeText = '';
    const absDays = Math.abs(diffDays);
    if (diffDays < 0) {
      timeText = currentLang === 'ar' ? `انتهى منذ ${absDays} يوم` : `Expired ${absDays} day${absDays === 1 ? '' : 's'} ago`;
    } else if (diffDays === 0) {
      timeText = currentLang === 'ar' ? 'ينتهي اليوم!' : 'Expires TODAY!';
    } else if (diffDays === 1) {
      timeText = currentLang === 'ar' ? 'ينتهي غداً!' : 'Expires Tomorrow!';
    } else {
      timeText = currentLang === 'ar' ? `متبقي ${diffDays} يوم` : `${diffDays} Days left`;
    }

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
    if (isNaN(dateObj.getTime())) return dateStr;

    return dateObj.toLocaleDateString(currentLang === 'ar' ? 'ar-SA' : 'en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  }

  function formatDateForExcel(dateStr) {
    if (!dateStr) return '';
    const parts = dateStr.split('-');
    if (parts.length !== 3) return dateStr;
    return `${parseInt(parts[1], 10)}/${parseInt(parts[2], 10)}/${parts[0]}`;
  }

  /* ==========================================
     Render Logic
     ========================================== */
  function renderApp() {
    renderStats();
    renderUrgentAlerts();
    renderMainViews();
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
            <strong>${escapeHtml(v.model)} (${escapeHtml(v.plate)})</strong>: ${info.timeText} (${info.formattedDate}).
          </div>
        </div>
        <button class="btn btn-sm btn-secondary renew-alert-btn" data-id="${v.id}">
          +1 ${currentLang === 'ar' ? 'سنة تجديد' : 'Year Renew'}
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

        // Search query: Search across Model, Vehicle No, Plate, Driver, Registration No, Type
        if (searchQuery.trim() !== '') {
          const q = searchQuery.toLowerCase().trim();
          const matchModel = (v.model || '').toLowerCase().includes(q);
          const matchVehicleNo = (v.vehicleNo || '').toLowerCase().includes(q);
          const matchPlate = (v.plate || '').toLowerCase().includes(q);
          const matchDriver = (v.driverName || '').toLowerCase().includes(q);
          const matchRegistration = (v.registrationNo || '').toLowerCase().includes(q);
          const matchType = (v.type || '').toLowerCase().includes(q);
          const matchRemarks = (v.remarks || '').toLowerCase().includes(q);

          return matchModel || matchVehicleNo || matchPlate || matchDriver || matchRegistration || matchType || matchRemarks;
        }

        return true;
      })
      .sort((a, b) => {
        if (currentSort === 'expiry-asc') {
          return new Date(a.expiryDate) - new Date(b.expiryDate);
        } else if (currentSort === 'expiry-desc') {
          return new Date(b.expiryDate) - new Date(a.expiryDate);
        } else if (currentSort === 'model-asc') {
          return (a.model || '').localeCompare(b.model || '');
        } else if (currentSort === 'plate-asc') {
          return (a.plate || '').localeCompare(b.plate || '');
        } else if (currentSort === 'vehicleNo-asc') {
          return (a.vehicleNo || '').localeCompare(b.vehicleNo || '', undefined, { numeric: true });
        } else if (currentSort === 'added-desc') {
          return (b.createdAt || 0) - (a.createdAt || 0);
        }
        return 0;
      });
  }

  function renderMainViews() {
    const list = getFilteredAndSortedVehicles();
    const t = i18n[currentLang];

    if (list.length === 0) {
      dom.vehiclesContainer.classList.add('hidden');
      dom.excelViewContainer.classList.add('hidden');
      dom.emptyState.classList.remove('hidden');

      if (vehicles.length === 0) {
        dom.emptyTitle.textContent = t.emptyTitle;
        dom.emptyDescription.textContent = t.emptyDescription;
      } else {
        dom.emptyTitle.textContent = currentLang === 'ar' ? 'لا توجد مركبات مطابقة' : 'No matching vehicles found';
        dom.emptyDescription.textContent = currentLang === 'ar' ? 'جرّب تعديل عبارة البحث أو اختيار تبويب آخر.' : 'Try adjusting your search query or filter tab.';
      }
      dom.emptyActionBtn.textContent = t.emptyAction;
      return;
    }

    dom.emptyState.classList.add('hidden');

    if (currentView === 'excel') {
      dom.vehiclesContainer.classList.add('hidden');
      dom.excelViewContainer.classList.remove('hidden');
      renderExcelTable(list);
    } else {
      dom.excelViewContainer.classList.add('hidden');
      dom.vehiclesContainer.classList.remove('hidden');

      if (currentView === 'list') {
        dom.vehiclesContainer.classList.add('list-view');
      } else {
        dom.vehiclesContainer.classList.remove('list-view');
      }

      dom.vehiclesContainer.innerHTML = '';
      list.forEach(vehicle => {
        const card = createVehicleCardElement(vehicle);
        dom.vehiclesContainer.appendChild(card);
      });
    }
  }

  /* ==========================================
     Excel Table Rendering Component
     ========================================== */
  function renderExcelTable(list) {
    const t = i18n[currentLang];

    let html = `
      <table class="excel-sheet-table">
        <thead>
          <tr>
            <th><div class="excel-header-content"><span>${t.colNo}</span> <span class="excel-filter-arrow">▼</span></div></th>
            <th><div class="excel-header-content"><span>${t.colVehicleNo}</span> <span class="excel-filter-arrow">▼</span></div></th>
            <th><div class="excel-header-content"><span>${t.colVehicleType}</span> <span class="excel-filter-arrow">▼</span></div></th>
            <th><div class="excel-header-content"><span>${t.colModel}</span> <span class="excel-filter-arrow">▼</span></div></th>
            <th><div class="excel-header-content"><span>${t.colPlateNo}</span> <span class="excel-filter-arrow">▼</span></div></th>
            <th><div class="excel-header-content"><span>${t.colDriverName}</span> <span class="excel-filter-arrow">▼</span></div></th>
            <th><div class="excel-header-content"><span>${t.colRegistrationNo}</span> <span class="excel-filter-arrow">▼</span></div></th>
            <th><div class="excel-header-content"><span>${t.colIssueDate}</span> <span class="excel-filter-arrow">▼</span></div></th>
            <th><div class="excel-header-content"><span>${t.colExpiryDate}</span> <span class="excel-filter-arrow">▼</span></div></th>
            <th><div class="excel-header-content"><span>${t.colDaysRemaining}</span> <span class="excel-filter-arrow">▼</span></div></th>
            <th><div class="excel-header-content"><span>${t.colStatus}</span> <span class="excel-filter-arrow">▼</span></div></th>
            <th><div class="excel-header-content"><span>${t.colRemarks}</span> <span class="excel-filter-arrow">▼</span></div></th>
            <th><div class="excel-header-content"><span>${t.colActions}</span></div></th>
          </tr>
        </thead>
        <tbody>
    `;

    list.forEach((v, index) => {
      const info = getExpiryStatusInfo(v.expiryDate);
      let statusText = t.statusValid;
      let statusPillClass = 'valid';

      if (info.status === 'expiring') {
        statusText = t.statusExpiring;
        statusPillClass = 'expiring';
      } else if (info.status === 'expired') {
        statusText = t.statusExpired;
        statusPillClass = 'expired';
      }

      html += `
        <tr>
          <td class="cell-no">${index + 1}</td>
          <td class="cell-center cell-num">${escapeHtml(v.vehicleNo || '')}</td>
          <td>${escapeHtml(v.type || '')}</td>
          <td><strong>${escapeHtml(v.model || '')}</strong></td>
          <td class="cell-center cell-num">${escapeHtml(v.plate || '')}</td>
          <td>${escapeHtml(v.driverName || '')}</td>
          <td class="cell-center cell-num">${escapeHtml(v.registrationNo || '')}</td>
          <td class="cell-center cell-num">${formatDateForExcel(v.issueDate)}</td>
          <td class="cell-center cell-num">${formatDateForExcel(v.expiryDate)}</td>
          <td class="cell-center cell-num"><strong>${info.diffDays}</strong></td>
          <td class="cell-center">
            <span class="status-pill ${statusPillClass}">${statusText}</span>
          </td>
          <td>${escapeHtml(v.remarks || '—')}</td>
          <td>
            <div class="excel-actions-cell">
              <button class="btn btn-secondary btn-sm renew-btn" title="Renew +1 Year" data-id="${v.id}">+1Yr</button>
              <button class="btn btn-secondary btn-sm edit-btn" title="Edit" data-id="${v.id}">✏️</button>
              <button class="btn btn-secondary btn-sm delete-btn" title="Delete" data-id="${v.id}" style="color: var(--status-red);">🗑️</button>
            </div>
          </td>
        </tr>
      `;
    });

    html += `</tbody></table>`;
    dom.excelViewContainer.innerHTML = html;

    dom.excelViewContainer.querySelectorAll('.renew-btn').forEach(btn => {
      btn.addEventListener('click', () => quickRenewVehicle(btn.dataset.id));
    });
    dom.excelViewContainer.querySelectorAll('.edit-btn').forEach(btn => {
      btn.addEventListener('click', () => openModalForEdit(btn.dataset.id));
    });
    dom.excelViewContainer.querySelectorAll('.delete-btn').forEach(btn => {
      btn.addEventListener('click', () => deleteVehicle(btn.dataset.id));
    });
  }

  function createVehicleCardElement(v) {
    const info = getExpiryStatusInfo(v.expiryDate);
    const t = i18n[currentLang];

    const card = document.createElement('div');
    card.className = `car-card status-${info.status}`;
    card.innerHTML = `
      <div class="car-card-header">
        <div class="license-plate">
          <div class="plate-header">VEHICLE NO: ${escapeHtml(v.vehicleNo || '000')}</div>
          <div class="plate-number">${escapeHtml(v.plate || '')}</div>
        </div>
        <div class="car-title-row">
          <div>
            <h3 class="car-model-name">${escapeHtml(v.model || '')}</h3>
            <div style="margin-top: 0.35rem; display: flex; gap: 0.5rem; align-items: center; flex-wrap: wrap;">
              <span class="type-badge">
                <span class="color-dot" style="background-color: ${escapeHtml(v.color || '#3B82F6')}"></span>
                ${escapeHtml(v.type || 'Vehicle')}
              </span>
              ${v.driverName ? `<span class="type-badge">👤 ${escapeHtml(v.driverName)}</span>` : ''}
              ${v.registrationNo ? `<span class="type-badge">📋 #${escapeHtml(v.registrationNo)}</span>` : ''}
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
          <span class="countdown-title">${t.colExpiryDate}</span>
          <span class="countdown-date">${info.formattedDate}</span>
        </div>
        <div class="time-left-big ${info.status === 'expired' ? 'text-red' : info.status === 'expiring' ? 'text-amber' : 'text-green'}">
          ${info.timeText}
        </div>
        <div class="progress-bar-bg" title="${info.percentRemaining}% validity window remaining">
          <div class="progress-bar-fill ${info.progressFillClass}" style="width: ${info.percentRemaining}%;"></div>
        </div>
      </div>
      ${v.remarks && v.remarks !== '—' ? `<p style="font-size: 0.82rem; color: var(--text-muted); margin-bottom: 0.75rem;"><strong>${t.colRemarks}:</strong> ${escapeHtml(v.remarks)}</p>` : ''}
      ${v.notes ? `<p style="font-size: 0.82rem; color: var(--text-dim); margin-bottom: 1rem; line-height: 1.4;">${escapeHtml(v.notes)}</p>` : ''}
      <div class="car-card-actions">
        <button class="btn btn-secondary btn-sm renew-btn" title="Extend expiry date by 1 Year">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/></svg>
          <span>+1 ${currentLang === 'ar' ? 'سنة' : 'Year'}</span>
        </button>
        <div class="action-btn-group">
          <button class="btn btn-secondary btn-sm edit-btn" title="Edit vehicle details">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
            <span>${currentLang === 'ar' ? 'تعديل' : 'Edit'}</span>
          </button>
          <button class="btn btn-secondary btn-sm delete-btn" title="Delete vehicle" style="color: var(--status-red);">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
          </button>
        </div>
      </div>
    `;

    card.querySelector('.renew-btn').addEventListener('click', () => quickRenewVehicle(v.id));
    card.querySelector('.edit-btn').addEventListener('click', () => openModalForEdit(v.id));
    card.querySelector('.delete-btn').addEventListener('click', () => deleteVehicle(v.id));

    return card;
  }

  /* ==========================================
     CRUD Operations & Form Handlers
     ========================================== */
  function openModalForAdd() {
    dom.carIdInput.value = '';
    dom.carForm.reset();

    const t = i18n[currentLang];
    dom.modalTitle.textContent = t.modalTitleAdd;

    const nextNo = String(vehicles.length + 1).padStart(3, '0');
    dom.vehicleNoInput.value = nextNo;
    dom.typeInput.value = '';
    dom.modelInput.value = '';
    updateBrandAndModelOptions('');

    dom.colorPicker.value = '#3B82F6';
    dom.colorInput.value = 'Blue';
    dom.remarksInput.value = '—';
    dom.previewPlateText.textContent = '12345';

    const today = new Date();
    const todayStr = today.toISOString().split('T')[0];
    const nextYear = new Date();
    nextYear.setFullYear(today.getFullYear() + 1);
    const nextYearStr = nextYear.toISOString().split('T')[0];

    dom.issueInput.value = todayStr;
    dom.expiryInput.value = nextYearStr;

    clearFormErrors();
    dom.carModal.classList.remove('hidden');
    dom.vehicleNoInput.focus();
  }

  function openModalForEdit(id) {
    const v = vehicles.find(item => item.id === id);
    if (!v) return;

    const t = i18n[currentLang];
    clearFormErrors();

    dom.carIdInput.value = v.id;
    dom.vehicleNoInput.value = v.vehicleNo || '';
    dom.typeInput.value = v.type || '';
    updateBrandAndModelOptions(v.type || '');
    dom.modelInput.value = v.model || '';
    dom.plateInput.value = v.plate || '';
    dom.driverInput.value = v.driverName || '';
    dom.registrationInput.value = v.registrationNo || '';
    dom.issueInput.value = v.issueDate || '';
    dom.expiryInput.value = v.expiryDate || '';
    dom.colorPicker.value = v.color || '#3B82F6';
    dom.colorInput.value = v.colorName || 'Blue';
    dom.remarksInput.value = v.remarks || '—';
    dom.notesInput.value = v.notes || '';
    dom.previewPlateText.textContent = (v.plate || '12345').toUpperCase();

    dom.modalTitle.textContent = t.modalTitleEdit;
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

    const vehicleNo = dom.vehicleNoInput.value.trim();
    const type = dom.typeInput.value.trim();
    const model = dom.modelInput.value.trim();
    const plate = dom.plateInput.value.trim();
    const driverName = dom.driverInput.value.trim();
    const registrationNo = dom.registrationInput.value.trim();
    const issueDate = dom.issueInput.value;
    const expiryDate = dom.expiryInput.value;
    const color = dom.colorPicker.value;
    const colorName = dom.colorInput.value.trim();
    const remarks = dom.remarksInput.value.trim() || '—';
    const notes = dom.notesInput.value.trim();
    const existingId = dom.carIdInput.value;

    let isValid = true;
    if (!vehicleNo) { showFieldError('vehicleno', 'Please enter vehicle number'); isValid = false; }
    if (!type) { showFieldError('type', 'Please enter vehicle type'); isValid = false; }
    if (!model) { showFieldError('model', 'Please enter model'); isValid = false; }
    if (!plate) { showFieldError('plate', 'Please enter plate number'); isValid = false; }
    if (!driverName) { showFieldError('driver', 'Please enter driver name'); isValid = false; }
    if (!registrationNo) { showFieldError('registration', 'Please enter registration number'); isValid = false; }
    if (!issueDate) { showFieldError('issue', 'Please select issue date'); isValid = false; }
    if (!expiryDate) { showFieldError('expiry', 'Please select expiry date'); isValid = false; }

    if (!isValid) return;

    const t = i18n[currentLang];

    if (existingId) {
      const index = vehicles.findIndex(v => v.id === existingId);
      if (index !== -1) {
        vehicles[index] = {
          ...vehicles[index],
          vehicleNo,
          type,
          model,
          plate: plate.toUpperCase(),
          driverName,
          registrationNo,
          issueDate,
          expiryDate,
          color,
          colorName,
          remarks,
          notes,
          updatedAt: Date.now()
        };
        showToast(`${t.updateSuccess}: ${model} (${plate.toUpperCase()})`, 'success');
      }
    } else {
      const newVehicle = {
        id: 'car_' + Date.now(),
        vehicleNo,
        type,
        model,
        plate: plate.toUpperCase(),
        driverName,
        registrationNo,
        issueDate,
        expiryDate,
        color,
        colorName,
        remarks,
        notes,
        createdAt: Date.now()
      };
      vehicles.push(newVehicle);
      showToast(`${t.addSuccess}: ${model} (${plate.toUpperCase()})`, 'success');
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

    const currentDateObj = new Date((v.expiryDate || getOffsetDateString(0)) + 'T00:00:00');
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let baseDate = currentDateObj < today ? today : currentDateObj;
    baseDate.setFullYear(baseDate.getFullYear() + 1);

    const newExpiryStr = baseDate.toISOString().split('T')[0];
    v.expiryDate = newExpiryStr;
    v.updatedAt = Date.now();

    saveVehicles();
    renderApp();

    const t = i18n[currentLang];
    showToast(`${v.model} (${v.plate}): ${t.renewSuccess}`, 'success');
  }

  function deleteVehicle(id) {
    const v = vehicles.find(item => item.id === id);
    if (!v) return;

    const t = i18n[currentLang];
    if (confirm(`${t.deleteConfirm} "${v.model} (${v.plate})"?`)) {
      vehicles = vehicles.filter(item => item.id !== id);
      saveVehicles();
      renderApp();
      showToast(t.deleteSuccess, 'info');
    }
  }

  /* ==========================================
     Excel (.xlsx) Export & Import Handlers
     ========================================== */
  function exportToExcel() {
    if (vehicles.length === 0) {
      showToast(currentLang === 'ar' ? 'لا توجد بيانات مركبات لتصديرها!' : 'No vehicle records to export!', 'warning');
      return;
    }

    try {
      const dataToExport = getFilteredAndSortedVehicles().map((v, index) => {
        const info = getExpiryStatusInfo(v.expiryDate);
        const isAr = currentLang === 'ar';

        let statusText = 'Valid';
        if (info.status === 'expiring') statusText = isAr ? 'ينتهي قريباً' : 'Expiring Soon';
        else if (info.status === 'expired') statusText = isAr ? 'منتهي الصلاحية' : 'Expired';
        else statusText = isAr ? 'صالح' : 'Valid';

        return {
          [isAr ? 'الرقم' : 'No.']: index + 1,
          [isAr ? 'رقم المركبة' : 'Vehicle No']: v.vehicleNo || '',
          [isAr ? 'نوع المركبة' : 'Vehicle Type']: v.type || '',
          [isAr ? 'الموديل' : 'Model']: v.model || '',
          [isAr ? 'رقم اللوحة' : 'Plate No']: v.plate || '',
          [isAr ? 'اسم السائق' : 'Driver Name']: v.driverName || '',
          [isAr ? 'رقم التسجيل' : 'Registration No']: v.registrationNo || '',
          [isAr ? 'تاريخ الإصدار' : 'Issue Date']: formatDateForExcel(v.issueDate),
          [isAr ? 'تاريخ الانتهاء' : 'Expiry Date']: formatDateForExcel(v.expiryDate),
          [isAr ? 'الأيام المتبقية' : 'Days Remaining']: info.diffDays,
          [isAr ? 'الحالة' : 'Status']: statusText,
          [isAr ? 'ملاحظات' : 'Remarks']: v.remarks || '—'
        };
      });

      if (window.XLSX) {
        const worksheet = XLSX.utils.json_to_sheet(dataToExport);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Vehicle Registrations');

        const fileName = `Vehicle_Registration_Tracker_${new Date().toISOString().split('T')[0]}.xlsx`;
        XLSX.writeFile(workbook, fileName);
        showToast(currentLang === 'ar' ? 'تم تصدير ملف Excel بنجاح!' : 'Exported Excel (.xlsx) file successfully!', 'success');
      } else {
        exportCSV(dataToExport);
      }
    } catch (err) {
      console.error('Error exporting to Excel:', err);
      showToast('Failed to export Excel file', 'error');
    }
  }

  function importFromExcel(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (event) {
      try {
        if (!window.XLSX) {
          showToast('Excel library not loaded.', 'error');
          return;
        }

        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        if (!Array.isArray(jsonRows) || jsonRows.length === 0) {
          showToast(currentLang === 'ar' ? 'لم يتم العثور على بيانات في ملف Excel!' : 'No rows found in Excel file.', 'warning');
          return;
        }

        const newVehicles = [];
        jsonRows.forEach((row, i) => {
          const vehicleNo = String(row['Vehicle No'] || row['رقم المركبة'] || row['No.'] || row['الرقم'] || (i + 1)).padStart(3, '0');
          const type = String(row['Vehicle Type'] || row['نوع المركبة'] || row['Type'] || 'Sedan');
          const model = String(row['Model'] || row['الموديل'] || row['Make & Model'] || 'Vehicle');
          const plate = String(row['Plate No'] || row['رقم اللوحة'] || row['Plate'] || '12345').toUpperCase();
          const driverName = String(row['Driver Name'] || row['اسم السائق'] || row['Driver'] || '');
          const registrationNo = String(row['Registration No'] || row['رقم التسجيل'] || row['Registration'] || '');
          const issueDate = parseExcelDateInput(row['Issue Date'] || row['تاريخ الإصدار']) || getOffsetDateString(-365);
          const expiryDate = parseExcelDateInput(row['Expiry Date'] || row['تاريخ الانتهاء']) || getOffsetDateString(365);
          const remarks = String(row['Remarks'] || row['ملاحظات'] || '—');

          if (model && plate) {
            newVehicles.push({
              id: 'car_' + Date.now() + '_' + i,
              vehicleNo,
              type,
              model,
              plate,
              driverName,
              registrationNo,
              issueDate,
              expiryDate,
              remarks,
              color: '#3B82F6',
              notes: 'Imported from Excel file',
              createdAt: Date.now()
            });
          }
        });

        if (newVehicles.length === 0) {
          showToast(currentLang === 'ar' ? 'لم يتم العثور على سجلات صالحة في الملف.' : 'No valid vehicle records found in file.', 'error');
          return;
        }

        vehicles = newVehicles;
        saveVehicles();
        renderApp();
        showToast(currentLang === 'ar' ? `تم استيراد ${newVehicles.length} مركبة بنجاح من ملف Excel!` : `Successfully imported ${newVehicles.length} vehicles from Excel!`, 'success');
      } catch (err) {
        console.error('Error importing Excel file:', err);
        showToast(currentLang === 'ar' ? 'فشل استيراد ملف Excel.' : 'Failed to import Excel file.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  }

  function parseExcelDateInput(val) {
    if (!val) return '';
    if (val instanceof Date && !isNaN(val.getTime())) {
      return val.toISOString().split('T')[0];
    }
    if (typeof val === 'number') {
      const dateObj = new Date(Math.round((val - 25569) * 86400 * 1000));
      if (!isNaN(dateObj.getTime())) {
        return dateObj.toISOString().split('T')[0];
      }
    }
    const str = String(val).trim();
    if (str.includes('/')) {
      const parts = str.split('/');
      if (parts.length === 3) {
        const month = parts[0].padStart(2, '0');
        const day = parts[1].padStart(2, '0');
        const year = parts[2].length === 2 ? '20' + parts[2] : parts[2];
        return `${year}-${month}-${day}`;
      }
    }
    if (str.includes('-')) {
      const parts = str.split('-');
      if (parts.length === 3) {
        if (parts[0].length === 4) return str;
        const month = parts[0].padStart(2, '0');
        const day = parts[1].padStart(2, '0');
        const year = parts[2].length === 2 ? '20' + parts[2] : parts[2];
        return `${year}-${month}-${day}`;
      }
    }
    return str;
  }

  function exportCSV(dataArray) {
    if (dataArray.length === 0) return;
    const headers = Object.keys(dataArray[0]);
    const csvRows = [headers.join(',')];

    dataArray.forEach(row => {
      const values = headers.map(h => `"${String(row[h]).replace(/"/g, '""')}"`);
      csvRows.push(values.join(','));
    });

    const csvString = '\uFEFF' + csvRows.join('\n');
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Vehicle_Registrations_${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('Exported CSV file successfully!', 'success');
  }

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
    showToast('Exported JSON backup file successfully!', 'success');
  }

  function importVehiclesJSON(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (event) {
      try {
        const importedData = JSON.parse(event.target.result);
        if (!Array.isArray(importedData)) {
          throw new Error('Invalid file format.');
        }

        const validVehicles = importedData.filter(item => item.model && item.expiryDate);
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
        showToast('Failed to import JSON file.', 'error');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  }

  /* ==========================================
     Toast Notifications & Escape Helpers
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

  function getOffsetDateString(daysOffset) {
    const d = new Date();
    d.setDate(d.getDate() + daysOffset);
    return d.toISOString().split('T')[0];
  }

  /* ==========================================
     Event Listeners Binds
     ========================================== */
  function setupEventListeners() {
    // Theme & Language Toggles
    dom.themeToggleBtn.addEventListener('click', toggleTheme);
    dom.langToggleBtn.addEventListener('click', toggleLanguage);

    // Export & Import Excel
    dom.exportExcelBtn.addEventListener('click', exportToExcel);
    dom.importExcelBtn.addEventListener('click', () => dom.importExcelFileInput.click());
    dom.importExcelFileInput.addEventListener('change', importFromExcel);

    // Export & Import JSON
    dom.exportBtn.addEventListener('click', exportVehiclesJSON);
    dom.importBtn.addEventListener('click', () => dom.importFileInput.click());
    dom.importFileInput.addEventListener('change', importVehiclesJSON);

    // Modal Triggers & Form
    dom.addCarBtn.addEventListener('click', openModalForAdd);
    dom.emptyActionBtn.addEventListener('click', openModalForAdd);
    dom.closeModalBtn.addEventListener('click', closeModal);
    dom.cancelModalBtn.addEventListener('click', closeModal);
    dom.carForm.addEventListener('submit', handleFormSubmit);

    // Live Plate Preview Typing Sync
    dom.plateInput.addEventListener('input', (e) => {
      const val = e.target.value.trim();
      dom.previewPlateText.textContent = val ? val.toUpperCase() : '12345';
    });

    // Brand Input Dynamic Model Options Sync
    dom.typeInput.addEventListener('input', (e) => {
      updateBrandAndModelOptions(e.target.value);
    });
    dom.typeInput.addEventListener('change', (e) => {
      updateBrandAndModelOptions(e.target.value);
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
      renderMainViews();
    });

    dom.clearSearchBtn.addEventListener('click', () => {
      dom.searchInput.value = '';
      searchQuery = '';
      dom.clearSearchBtn.classList.add('hidden');
      renderMainViews();
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
        renderMainViews();
      });
    });

    // Sort Select
    dom.sortSelect.addEventListener('change', (e) => {
      currentSort = e.target.value;
      renderMainViews();
    });

    // View Toggles
    dom.viewExcelBtn.addEventListener('click', () => {
      currentView = 'excel';
      dom.viewExcelBtn.classList.add('active');
      dom.viewGridBtn.classList.remove('active');
      dom.viewListBtn.classList.remove('active');
      renderMainViews();
    });

    dom.viewGridBtn.addEventListener('click', () => {
      currentView = 'grid';
      dom.viewGridBtn.classList.add('active');
      dom.viewExcelBtn.classList.remove('active');
      dom.viewListBtn.classList.remove('active');
      renderMainViews();
    });

    dom.viewListBtn.addEventListener('click', () => {
      currentView = 'list';
      dom.viewListBtn.classList.add('active');
      dom.viewExcelBtn.classList.remove('active');
      dom.viewGridBtn.classList.remove('active');
      renderMainViews();
    });

    // Keyboard & Backdrop Modal Closing
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !dom.carModal.classList.contains('hidden')) {
        closeModal();
      }
    });

    dom.carModal.addEventListener('click', (e) => {
      if (e.target === dom.carModal) {
        closeModal();
      }
    });
  }

  // DOMReady initialization
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
