(() => {
  const apiBase = 'http://localhost:8080';
  let token = sessionStorage.getItem('esa_token');
  let currentUser = sessionStorage.getItem('esa_user');
  let role = { admin: false, developer: false };
  let developerOverride = false;
  let appsCache = [];
  let formMode = 'create';
  let editingApp = null;
  let activeApp = null;
  let activeComponents = [];
  let workbookApp = null;
  let builderTarget = null;
  let builderWidgets = [];
  let builderSelectedWidget = null;
  let builderSheets = [];
  let builderSavedState = null;
  const qs = (sel) => document.querySelector(sel);
  const toastEl = qs('#toast');

  // Widget type definitions - shared between builder and runtime
  const WIDGET_TYPES = {
    // Containers
    panel: { category: 'container', isContainer: true, icon: '📦', label: 'Panel' },
    notebook: { category: 'container', isContainer: true, icon: '📑', label: 'Notebook' },
    'boxsizer-v': { category: 'container', isContainer: true, icon: '⬇️', label: 'VBox' },
    'boxsizer-h': { category: 'container', isContainer: true, icon: '➡️', label: 'HBox' },
    gridsizer: { category: 'container', isContainer: true, icon: '▦', label: 'Grid' },
    scrolled: { category: 'container', isContainer: true, icon: '📜', label: 'Scroll' },
    // Inputs
    textinput: { category: 'input', icon: '📝', label: 'Text Input' },
    number: { category: 'input', icon: '🔢', label: 'Number' },
    currency: { category: 'input', icon: '💰', label: 'Currency' },
    percentage: { category: 'input', icon: '📊', label: 'Percentage' },
    textarea: { category: 'input', icon: '📄', label: 'Text Area' },
    dropdown: { category: 'input', icon: '📋', label: 'Dropdown' },
    checkbox: { category: 'input', icon: '☑️', label: 'Checkbox' },
    radio: { category: 'input', icon: '🔘', label: 'Radio' },
    slider: { category: 'input', icon: '🎚️', label: 'Slider' },
    spinctrl: { category: 'input', icon: '🔄', label: 'Spin' },
    datepicker: { category: 'input', icon: '📅', label: 'Date' },
    colorpicker: { category: 'input', icon: '🎨', label: 'Color' },
    // Display
    label: { category: 'display', icon: '🏷️', label: 'Label' },
    output: { category: 'display', icon: '👁️', label: 'Output' },
    image: { category: 'display', icon: '🖼️', label: 'Image' },
    gauge: { category: 'display', icon: '📊', label: 'Gauge' },
    separator: { category: 'display', icon: '➖', label: 'Separator' },
    spacer: { category: 'display', icon: '⬜', label: 'Spacer' },
    // Actions
    button: { category: 'action', icon: '🔘', label: 'Button' },
    togglebtn: { category: 'action', icon: '🔀', label: 'Toggle' },
    link: { category: 'action', icon: '🔗', label: 'Link' },
    // Data
    datagrid: { category: 'data', icon: '📊', label: 'Data Grid', requiresExcel: true },
    chart: { category: 'data', icon: '📈', label: 'Chart', requiresExcel: true },
    formula: { category: 'data', icon: '∑', label: 'Formula', requiresExcel: true }
  };

  const showToast = (msg, isError = false) => {
    toastEl.textContent = msg;
    toastEl.classList.toggle('error', isError);
    toastEl.classList.remove('hidden');
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => toastEl.classList.add('hidden'), 2800);
  };

  const authHeaders = () => token ? { 'Authorization': `Bearer ${token}` } : {};
  let redirectingToLogin = false;

  const handleUnauthorized = () => {
    if (redirectingToLogin) return;
    redirectingToLogin = true;
    showToast('Session expired. Redirecting to login…', true);
    clearSession();
    window.location.replace('login.html');
  };

  const apiFetch = async (input, init = {}) => {
    const res = await fetch(input, init);
    if (res.status === 401) {
      handleUnauthorized();
      throw new Error('Unauthorized');
    }
    return res;
  };

  const sessionInfo = qs('#session-info');
  const logoutBtn = qs('#logout-btn');
  const backToLogin = qs('#back-to-login');
  const sessionUser = qs('#session-user');
  const sessionRoles = qs('#session-roles');
  const createPanel = qs('#create-panel');
  const createHint = qs('#create-hint');
  const developerPill = qs('#developer-pill');
  const developerList = qs('#developer-list');
  const devFormMode = qs('#dev-form-mode');
  const devCreateNew = qs('#dev-create-new');
  const devModal = qs('#dev-modal');
  const devModalClose = qs('#dev-modal-close');
  const devModalCancel = qs('#dev-modal-cancel');
  const bodyEl = document.body;
  const appsListEl = qs('#apps-list');
  const appSearchInput = qs('#app-search');
  const appOverlay = qs('#app-overlay');
  const appUiForm = qs('#app-ui-form');
  const appUiEmpty = qs('#app-ui-empty');
  const appWorkspaceTitle = qs('#app-workspace-title');
  const appWorkspaceDesc = qs('#app-workspace-desc');
  const appRefreshBtn = qs('#app-refresh');
  const appExitBtn = qs('#app-exit');
  const adminPanel = qs('#admin-panel');
  const adminWarning = qs('#admin-warning');
  const usersArea = qs('#users-area');
  const builderModal = qs('#ui-builder-modal');
  const builderClose = qs('#ui-builder-close');
  const builderAppLabel = qs('#builder-app-label');
  const builderSaveLayoutBtn = qs('#builder-save-layout');
  const builderImportSheet = qs('#builder-import-sheet');
  const builderImportRange = qs('#builder-import-range');
  const builderImportExcel = qs('#builder-import-excel');
  const builderWidgetTree = qs('#builder-widget-tree');
  const builderTreeRoot = qs('#builder-tree-root');
  const builderEmpty = qs('#builder-empty');
  const builderNoSelection = qs('#builder-no-selection');
  const builderPropertyForm = qs('#builder-property-form');
  const builderSash = qs('#builder-sash');
  const builderHsash = qs('#builder-hsash');
  const builderLeftPanel = qs('.builder-left-panel');
  const builderTreePanel = qs('.builder-tree-panel');
  const builderPropertiesPanel = qs('.builder-properties');
  const builderPreviewPanel = qs('.builder-preview-panel');
  const builderMain = qs('.builder-main');
  const builderPreview = qs('#builder-preview');
  const builderPreviewContent = qs('#builder-preview-content');
  const builderPreviewEmpty = qs('#builder-preview-empty');
  // Property form elements
  const propWidgetId = qs('#prop-widget-id');
  const propName = qs('#prop-name');
  const propType = qs('#prop-type'); // legacy, may be null
  const propWidgetLegend = qs('#prop-widget-legend');
  const propLabel = qs('#prop-label');
  const propPlaceholder = qs('#prop-placeholder');
  const propTooltip = qs('#prop-tooltip');
  const propWidth = qs('#prop-width');
  const propHeight = qs('#prop-height');
  const propMinWidth = qs('#prop-min-width');
  const propMinHeight = qs('#prop-min-height');
  const propProportion = qs('#prop-proportion');
  const propFlags = qs('#prop-flags');
  const propMargin = qs('#prop-margin');
  const propPadding = qs('#prop-padding');
  const propOrientation = qs('#prop-orientation');
  const propCols = qs('#prop-cols');
  const propRows = qs('#prop-rows');
  const propHgap = qs('#prop-hgap');
  const propVgap = qs('#prop-vgap');
  const propOptions = qs('#prop-options');
  const propMin = qs('#prop-min');
  const propMax = qs('#prop-max');
  const propStep = qs('#prop-step');
  const propDefault = qs('#prop-default');
  const propExcelEnabled = qs('#prop-excel-enabled');
  const propExcelFields = qs('#prop-excel-fields');
  const propSheet = qs('#prop-sheet');
  const propCell = qs('#prop-cell');
  const propMode = qs('#prop-mode');
  const propOnclick = qs('#prop-onclick');
  const propTarget = qs('#prop-target');
  const propDeleteWidget = qs('#prop-delete-widget');
  // Application form elements
  const builderAppForm = qs('#builder-app-form');
  const propAppName = qs('#prop-app-name');
  const propAppOwner = qs('#prop-app-owner');
  const propAppDescription = qs('#prop-app-description');
  const propAppPublic = qs('#prop-app-public');
  const propAppAccessGroup = qs('#prop-app-access-group');
  const propDeleteApp = qs('#prop-delete-app');
  // Property groups
  const propGroupLabel = qs('#prop-group-label');
  const propGroupSizer = qs('#prop-group-sizer');
  const propGroupInput = qs('#prop-group-input');
  const propGroupExcel = qs('#prop-group-excel');
  const propGroupAction = qs('#prop-group-action');
  const previewModal = qs('#ui-preview-modal');
  const previewClose = qs('#ui-preview-close');
  const previewCloseBtn = qs('#ui-preview-close-btn');
  const previewContent = qs('#ui-preview-content');
  const previewEmpty = qs('#ui-preview-empty');
  const previewSubtitle = qs('#ui-preview-subtitle');
  const tabButtons = {
    apps: qs('[data-tab-btn="apps"]'),
    developer: qs('[data-tab-btn="developer"]'),
    admin: qs('[data-tab-btn="admin"]')
  };
  const tabContents = {
    apps: qs('#apps-panel'),
    developer: qs('[data-tab="developer"]'),
    admin: qs('[data-tab="admin"]')
  };
  let activeTab = 'apps';

  if (!token || !currentUser) {
    window.location.replace('login.html');
    return;
  }

  sessionInfo.classList.remove('hidden');
  sessionUser.textContent = currentUser;

  Object.entries(tabButtons).forEach(([key, btn]) => {
    btn.addEventListener('click', () => setTab(key));
  });

  logoutBtn.addEventListener('click', async () => {
    try { await exitActiveApp(); } catch (_) {}
    try { await apiFetch(`${apiBase}/logout`, { method: 'POST', headers: authHeaders() }); } catch (_) {}
    clearSession();
    window.location.replace('login.html');
  });

  if (token && backToLogin) backToLogin.remove();
  if (backToLogin) {
    backToLogin.addEventListener('click', () => {
      clearSession();
      window.location.href = 'login.html';
    });
  }

  async function detectRole() {
    role = { admin: false, developer: false };
    if (!token || !currentUser) { toggleCreateVisibility(); return; }
    try {
      const res = await apiFetch(`${apiBase}/users`, { headers: { ...authHeaders() } });
      if (res.status === 200) {
        const users = await res.json();
        const me = users.find(u => u.name === currentUser);
        role.admin = true;
        role.developer = true; // admins imply developer
        if (me && me.role === 'developer') role.developer = true;
        if (me && me.role === 'user') role.developer = false;
        if (me && me.role === 'administrator') role.admin = true;
        adminWarning.classList.add('hidden');
        usersArea.classList.remove('hidden');
        renderUsers(users);
      } else {
        adminWarning.classList.remove('hidden');
        usersArea.classList.add('hidden');
      }
    } catch (_) {
      adminWarning.classList.remove('hidden');
      usersArea.classList.add('hidden');
    }
    toggleCreateVisibility();
    const adminBadge = role.admin ? ' <span class="admin-badge">Admin</span>' : '';
    sessionUser.innerHTML = `${escapeHtml(currentUser)}${adminBadge}`;
  }

  function toggleCreateVisibility() {
    const canDev = developerOverride || role.developer || role.admin;
    const signedIn = !!token;
    createPanel.classList.toggle('hidden', !canDev || !signedIn || activeTab !== 'developer');
    createHint.classList.toggle('hidden', canDev && signedIn);
    developerPill.classList.toggle('hidden', canDev && signedIn);
    tabButtons.developer.classList.toggle('hidden', !canDev || !signedIn);
    tabContents.apps.classList.toggle('hidden', !signedIn || activeTab !== 'apps');
    adminPanel.classList.toggle('hidden', !signedIn || activeTab !== 'admin');
    tabButtons.admin.classList.toggle('hidden', !role.admin || !signedIn);
    if ((!canDev || !signedIn) && devModal && !devModal.classList.contains('hidden')) {
      closeDevModal(true);
    }
    if (activeTab === 'developer' && (!canDev || !signedIn)) setTab('apps');
    if (activeTab === 'admin' && (!role.admin || !signedIn)) setTab('apps');
  }

  if (devModalClose) devModalClose.addEventListener('click', () => closeDevModal(true));
  if (devModalCancel) devModalCancel.addEventListener('click', () => closeDevModal(true));
  if (devModal) {
    devModal.addEventListener('click', (e) => {
      if (e.target === devModal) closeDevModal(true);
    });
  }
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && appOverlay && !appOverlay.classList.contains('hidden')) {
      exitActiveApp();
    }
    if (e.key === 'Escape' && devModal && !devModal.classList.contains('hidden')) {
      closeDevModal(true);
    }
    if (e.key === 'Escape' && builderModal && !builderModal.classList.contains('hidden')) {
      closeBuilderModal();
    }
    if (e.key === 'Escape' && previewModal && !previewModal.classList.contains('hidden')) {
      closePreviewModal();
    }
  });

  appsListEl?.addEventListener('click', handleAppsListClick);
  appSearchInput?.addEventListener('input', handleAppSearch);
  appRefreshBtn?.addEventListener('click', () => refreshActiveApp(true));
  appExitBtn?.addEventListener('click', () => exitActiveApp());

  if (builderClose) builderClose.addEventListener('click', () => closeBuilderModal());
  if (builderModal) {
    builderModal.addEventListener('click', (e) => {
      if (e.target === builderModal) closeBuilderModal();
    });
  }
  if (previewClose) previewClose.addEventListener('click', () => closePreviewModal());
  if (previewCloseBtn) previewCloseBtn.addEventListener('click', () => closePreviewModal());
  if (previewModal) {
    previewModal.addEventListener('click', (e) => {
      if (e.target === previewModal) closePreviewModal();
    });
  }
  
  // New builder event listeners

  builderTreeRoot?.addEventListener('click', handleWidgetTreeClick);
  builderTreeRoot?.addEventListener('change', handleAddWidgetDropdown);
  builderPropertyForm?.addEventListener('submit', handlePropertyFormSubmit);
  builderAppForm?.addEventListener('submit', handleAppFormSubmit);
  propExcelEnabled?.addEventListener('change', handleExcelToggle);
  propDeleteWidget?.addEventListener('click', handleDeleteWidget);
  propDeleteApp?.addEventListener('click', handleDeleteApp);
  builderSaveLayoutBtn?.addEventListener('click', () => saveBuilderLayout());
  builderImportExcel?.addEventListener('click', handleExcelImport);
  
  // Sash resize functionality
  if (builderSash && builderTreePanel && builderPropertiesPanel) {
    let isDragging = false;
    let startY = 0;
    let startTreeHeight = 0;
    
    builderSash.addEventListener('mousedown', (e) => {
      isDragging = true;
      startY = e.clientY;
      startTreeHeight = builderTreePanel.offsetHeight;
      builderSash.classList.add('dragging');
      document.body.style.cursor = 'ns-resize';
      document.body.style.userSelect = 'none';
      e.preventDefault();
    });
    
    document.addEventListener('mousemove', (e) => {
      if (!isDragging) return;
      const deltaY = e.clientY - startY;
      const newHeight = Math.max(80, Math.min(startTreeHeight + deltaY, window.innerHeight * 0.6));
      builderTreePanel.style.flex = 'none';
      builderTreePanel.style.height = `${newHeight}px`;
    });
    
    document.addEventListener('mouseup', () => {
      if (isDragging) {
        isDragging = false;
        builderSash.classList.remove('dragging');
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
      }
    });
  }

  // Horizontal sash resize functionality (between left panel and preview)
  if (builderHsash && builderLeftPanel && builderMain) {
    let isDraggingH = false;
    let startX = 0;
    let startLeftWidth = 0;
    
    builderHsash.addEventListener('mousedown', (e) => {
      isDraggingH = true;
      startX = e.clientX;
      startLeftWidth = builderLeftPanel.offsetWidth;
      builderHsash.classList.add('dragging');
      document.body.style.cursor = 'ew-resize';
      document.body.style.userSelect = 'none';
      e.preventDefault();
    });
    
    document.addEventListener('mousemove', (e) => {
      if (!isDraggingH) return;
      const deltaX = e.clientX - startX;
      const newWidth = Math.max(250, Math.min(startLeftWidth + deltaX, window.innerWidth * 0.5));
      builderMain.style.gridTemplateColumns = `${newWidth}px 6px 1fr`;
    });
    
    document.addEventListener('mouseup', () => {
      if (isDraggingH) {
        isDraggingH = false;
        builderHsash.classList.remove('dragging');
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
      }
    });
  }

  // Preview click to select widgets and handle notebook tabs
  builderPreviewContent?.addEventListener('click', (e) => {
    // Handle notebook tab clicks
    const tabBtn = e.target.closest('.notebook-tab');
    if (tabBtn) {
      const notebookId = tabBtn.dataset.notebookId;
      const tabIdx = parseInt(tabBtn.dataset.tabIdx, 10);
      if (notebookId && !isNaN(tabIdx)) {
        previewNotebookTabs[notebookId] = tabIdx;
        renderBuilderPreview();
        return;
      }
    }
    
    // Handle widget selection
    const previewEl = e.target.closest('[data-preview-id]');
    if (!previewEl) return;
    const widgetId = previewEl.dataset.previewId;
    if (widgetId && widgetId !== builderSelectedWidget) {
      builderSelectedWidget = widgetId;
      renderWidgetTree();
      const widget = findWidgetById(builderWidgets, widgetId);
      if (widget) showPropertyForm(widget);
    }
  });

  appUiForm?.addEventListener('change', handleAppFormChange);
  appUiForm?.addEventListener('submit', (e) => e.preventDefault());

  async function refreshApps() {
    if (!token) { showToast('Login first', true); return; }
    try {
      const res = await apiFetch(`${apiBase}/apps`, { headers: { ...authHeaders() } });
      if (!res.ok) throw new Error('Failed to load apps');
      const apps = await res.json();
      appsCache = apps;
      renderApps(apps);
      renderDeveloperApps();
    } catch (err) {
      renderApps([]);
      appsCache = [];
      renderDeveloperApps();
      showToast(err.message || 'Failed to load apps', true);
    }
  }

  function renderApps(apps) {
    const list = qs('#apps-list');
    if (!apps || apps.length === 0) {
      list.innerHTML = '<div class="hint">No applications visible.</div>';
      return;
    }
    list.innerHTML = apps.map(a => {
      const imageHtml = a.image_base64
        ? `<img src="data:image/png;base64,${a.image_base64}" alt="${escapeHtml(a.name)}" class="app-card-image" />`
        : `<div class="app-card-image" style="display: grid; place-items: center; font-size: 48px; font-weight: 700; color: var(--accent)">${escapeHtml(a.name.charAt(0).toUpperCase())}</div>`;
      const description = formatDescription(a.description);
      return `<button type="button" class="app-card app-card-clickable" data-launch-app="${escapeHtml(a.name)}" data-owner="${escapeHtml(a.owner)}" data-group="${escapeHtml(a.access_group || '')}">
        ${imageHtml}
        <div class="app-card-body">
          <h4>${escapeHtml(a.name)}</h4>
          <p>${escapeHtml(description)}</p>
        </div>
      </button>`;
    }).join('');
  }

  function handleAppsListClick(e) {
    const launchBtn = e.target.closest('[data-launch-app]');
    if (launchBtn) {
      const owner = launchBtn.dataset.owner;
      const name = launchBtn.dataset.launchApp;
      launchApp(owner, name);
    }
  }

  function handleAppSearch(e) {
    const query = (e.target.value || '').toLowerCase().trim();
    const cards = appsListEl?.querySelectorAll('.app-card-clickable');
    if (!cards) return;
    cards.forEach(card => {
      const name = (card.dataset.launchApp || '').toLowerCase();
      const group = (card.dataset.group || '').toLowerCase();
      const owner = (card.dataset.owner || '').toLowerCase();
      const text = card.textContent.toLowerCase();
      const matches = !query || name.includes(query) || group.includes(query) || owner.includes(query) || text.includes(query);
      card.style.display = matches ? '' : 'none';
    });
  }

  function canAuthorApp(owner) {
    if (!token) return false;
    if (role.admin) return true;
    if (developerOverride || role.developer) {
      return owner === currentUser;
    }
    return false;
  }

  async function launchApp(owner, name) {
    if (!token) return showToast('Login first', true);
    if (!owner || !name) return;
    try {
      showAppOverlay();
      updateWorkspaceState({ title: name, desc: 'Loading workbook and UI schema.', busy: true });
      const schema = await fetchAppSchema(owner, name);
      await ensureWorkbookLoaded(owner, name);
      activeApp = { owner, name, schema };
      renderAppUi(schema);
      await refreshActiveApp(true, false);
      updateWorkspaceState({ title: name, desc: 'Workbook synced. Edit inputs to push values back to Excel.', busy: false });
      if (appRefreshBtn) appRefreshBtn.disabled = false;
      if (appExitBtn) appExitBtn.disabled = false;
    } catch (err) {
      showToast(err.message || 'Failed to launch application', true);
      resetWorkspace();
    }
  }

  async function ensureWorkbookLoaded(owner, name) {
    if (workbookApp && workbookApp.owner === owner && workbookApp.name === name) return;
    await closeWorkbook();
    const res = await apiFetch(`${apiBase}/excel/load`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...authHeaders() },
      body: JSON.stringify({ owner, name })
    });
    if (!res.ok) throw new Error('Unable to load workbook');
    workbookApp = { owner, name };
  }

  async function closeWorkbook() {
    if (!workbookApp) return;
    try {
      await apiFetch(`${apiBase}/excel/close`, { method: 'POST', headers: { ...authHeaders() } });
    } catch (_) {}
    workbookApp = null;
  }

  function resetWorkspace() {
    activeApp = null;
    activeComponents = [];
    appUiForm?.classList.add('hidden');
    if (appUiForm) appUiForm.innerHTML = '';
    appUiEmpty?.classList.remove('hidden');
    if (appRefreshBtn) appRefreshBtn.disabled = true;
    if (appExitBtn) appExitBtn.disabled = true;
    if (appWorkspaceTitle) appWorkspaceTitle.textContent = 'Select an application';
    if (appWorkspaceDesc) appWorkspaceDesc.textContent = 'Launch an application to render its authored UI.';
    hideAppOverlay();
  }

  async function refreshActiveApp(showToastOnError = false, notifyOnSuccess = true) {
    if (!activeApp || !activeComponents.length) return;
    const failures = [];
    await Promise.all(activeComponents.map(async (component) => {
      if (!component?.sheet || !component?.cell) return;
      try {
        const value = await queryComponentValue(component);
        updateComponentDisplay(component.id, value);
      } catch (err) {
        const reason = err?.message || 'Unable to read cell';
        failures.push(`${component.label || component.cell || 'Field'}: ${reason}`);
        console.warn('Failed to refresh component', component, err);
      }
    }));
    if (showToastOnError) {
      if (failures.length) {
        const [first, ...rest] = failures;
        const extra = rest.length ? ` (+${rest.length} more)` : '';
        showToast(first + extra, true);
      } else if (notifyOnSuccess) {
        showToast('Values refreshed');
      }
    }
  }

  async function queryComponentValue(component) {
    if (!component?.sheet || !component?.cell) return '';
    const res = await apiFetch(`${apiBase}/excel/query`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...authHeaders() },
      body: JSON.stringify({ sheet: component.sheet, range: component.cell })
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
      const error = data?.error || 'Unable to read cell';
      if (error === 'range not found') {
        throw new Error(`Cell ${component.sheet}!${component.cell} not found - ensure Excel file is open and cell exists`);
      }
      throw new Error(error);
    }
    return data?.value ?? '';
  }

  function updateComponentDisplay(id, value) {
    const el = appUiForm?.querySelector(`[data-component="${id}"]`);
    if (!el) return;
    
    if (el.tagName === 'OUTPUT') {
      el.textContent = formatValue(value);
    } else if (el.type === 'checkbox') {
      el.checked = value === 'TRUE' || value === true || value === 1;
    } else if (el.type === 'range') {
      el.value = value ?? '';
      const valueSpan = appUiForm?.querySelector(`[data-value-for="${id}"]`);
      if (valueSpan) valueSpan.textContent = value ?? '—';
    } else {
      el.value = value ?? '';
    }
  }

  function formatValue(val) {
    if (Array.isArray(val)) {
      const flat = val.flat ? val.flat() : val;
      return flat.join(', ');
    }
    if (typeof val === 'object' && val !== null) {
      try { return JSON.stringify(val); } catch (_) { return String(val); }
    }
    return val ?? '';
  }

  function renderAppUi(schema) {
    const widgets = schema?.widgets || [];
    const tabs = schema?.tabs || [];
    const legacyComponents = schema?.components;
    
    // Handle new widget format
    if (widgets && widgets.length > 0) {
      renderWidgetBasedUi(widgets);
      return;
    }
    
    // Handle legacy tab format
    if (tabs && tabs.length > 0) {
      renderTabBasedAppUi(tabs);
      return;
    }
    
    // Handle very old format (no tabs)
    if (legacyComponents && legacyComponents.length > 0) {
      const components = normalizeComponents(legacyComponents);
      activeComponents = components;
      renderLegacyAppUi(components);
      return;
    }
    
    // Empty state
    if (!appUiForm || !appUiEmpty) return;
    appUiForm.classList.add('hidden');
    appUiEmpty.classList.remove('hidden');
    appUiForm.innerHTML = '';
    activeComponents = [];
  }

  function renderWidgetBasedUi(widgets) {
    if (!appUiForm || !appUiEmpty) return;
    if (!widgets.length) {
      appUiForm.classList.add('hidden');
      appUiEmpty.classList.remove('hidden');
      appUiForm.innerHTML = '';
      activeComponents = [];
      return;
    }
    
    appUiEmpty.classList.add('hidden');
    appUiForm.classList.remove('hidden');
    
    // Collect all Excel-bound widgets for refresh
    activeComponents = collectExcelWidgets(widgets);
    
    // Render widget tree
    appUiForm.innerHTML = widgets.map(w => renderWidget(w)).join('');
    
    // Add tab switching logic for notebook widgets
    setupNotebookTabHandlers();
  }

  function setupNotebookTabHandlers() {
    const tabBtns = appUiForm.querySelectorAll('.app-tab-btn');
    tabBtns.forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const target = btn.dataset.tabTarget;
        // Find the parent notebook
        const notebook = btn.closest('.widget-notebook');
        if (!notebook) return;
        
        // Toggle active state within this notebook only
        notebook.querySelectorAll('.app-tab-btn').forEach(b => b.classList.remove('active'));
        notebook.querySelectorAll('.app-tab-panel').forEach(p => p.classList.remove('active'));
        btn.classList.add('active');
        const panel = notebook.querySelector(`[data-tab-id="${target}"]`);
        if (panel) panel.classList.add('active');
      });
    });
  }

  function collectExcelWidgets(widgets) {
    let result = [];
    for (const w of widgets) {
      if (w.excel?.enabled && w.excel.sheet && w.excel.cell) {
        result.push({
          id: w.id,
          sheet: w.excel.sheet,
          cell: w.excel.cell,
          mode: w.excel.mode,
          componentType: w.type
        });
      }
      if (w.children) {
        result = result.concat(collectExcelWidgets(w.children));
      }
    }
    return result;
  }

  function renderWidget(widget) {
    const def = WIDGET_TYPES[widget.type] || {};
    const props = widget.properties || {};
    const excel = widget.excel || {};
    const label = props.label || widget.name || 'Widget';
    
    // Layout styles
    const styles = [];
    if (props.width) styles.push(`width: ${props.width}`);
    if (props.height) styles.push(`height: ${props.height}`);
    if (props.margin) styles.push(`margin: ${props.margin}px`);
    if (props.padding) styles.push(`padding: ${props.padding}px`);
    const styleAttr = styles.length ? ` style="${styles.join('; ')}"` : '';
    
    // Handle containers
    if (def.isContainer) {
      return renderContainerWidget(widget, def, props, styleAttr);
    }
    
    // Handle regular widgets
    return renderLeafWidget(widget, def, props, excel, styleAttr);
  }

  function renderContainerWidget(widget, def, props, styleAttr) {
    const children = (widget.children || []).map(c => renderWidget(c)).join('');
    const orientation = props.orientation || 'vertical';
    const gap = orientation === 'vertical' ? (props.vgap || 4) : (props.hgap || 4);
    
    switch (widget.type) {
      case 'notebook':
        // Tabbed notebook
        const tabBtns = (widget.children || []).map((child, idx) => {
          const tabLabel = child.properties?.label || child.name || `Tab ${idx + 1}`;
          return `<button type="button" class="app-tab-btn${idx === 0 ? ' active' : ''}" data-tab-target="nb-${widget.id}-${idx}">${escapeHtml(tabLabel)}</button>`;
        }).join('');
        const tabPanels = (widget.children || []).map((child, idx) => {
          const panelContent = (child.children || []).map(c => renderWidget(c)).join('');
          return `<div class="app-tab-panel${idx === 0 ? ' active' : ''}" data-tab-id="nb-${widget.id}-${idx}"${styleAttr}>${panelContent}</div>`;
        }).join('');
        return `<div class="widget-notebook" data-widget="${widget.id}"><div class="app-tabs">${tabBtns}</div><div class="app-tab-content">${tabPanels}</div></div>`;
      
      case 'boxsizer-v':
        return `<div class="widget-vbox" style="display: flex; flex-direction: column; gap: ${gap}px;"${styleAttr} data-widget="${widget.id}">${children}</div>`;
      
      case 'boxsizer-h':
        return `<div class="widget-hbox" style="display: flex; flex-direction: row; gap: ${gap}px; flex-wrap: wrap;"${styleAttr} data-widget="${widget.id}">${children}</div>`;
      
      case 'gridsizer':
        const cols = props.cols || 2;
        const hgap = props.hgap || 4;
        const vgap = props.vgap || 4;
        return `<div class="widget-grid" style="display: grid; grid-template-columns: repeat(${cols}, 1fr); gap: ${vgap}px ${hgap}px;"${styleAttr} data-widget="${widget.id}">${children}</div>`;
      
      case 'scrolled':
        return `<div class="widget-scrolled" style="overflow: auto; max-height: ${props.height || '300px'};"${styleAttr} data-widget="${widget.id}">${children}</div>`;
      
      case 'panel':
      default:
        return `<div class="widget-panel"${styleAttr} data-widget="${widget.id}">${children}</div>`;
    }
  }

  function renderLeafWidget(widget, def, props, excel, styleAttr) {
    const label = escapeHtml(props.label || widget.name || def.label);
    const id = widget.id;
    const hasExcel = excel.enabled && excel.sheet && excel.cell;
    const placeholder = escapeHtml(props.placeholder || '');
    const options = (props.options || '').split(',').map(o => o.trim()).filter(Boolean);
    // Only show label if explicitly set in properties
    const labelHtml = props.label ? `<label class="widget-field-label">${escapeHtml(props.label)}</label>` : '';
    
    switch (widget.type) {
      case 'label':
        return `<div class="app-widget"${styleAttr}><span class="widget-label">${label}</span></div>`;
      
      case 'output':
        return `<div class="app-widget"${styleAttr}>${labelHtml}<output data-component="${id}" class="widget-output">—</output></div>`;
      
      case 'textinput':
        return `<div class="app-widget"${styleAttr}>${labelHtml}<input data-component="${id}" type="text" placeholder="${placeholder || 'Enter text...'}" /></div>`;
      
      case 'number':
        return `<div class="app-widget"${styleAttr}>${labelHtml}<input data-component="${id}" type="number" placeholder="${placeholder}" min="${props.min || ''}" max="${props.max || ''}" /></div>`;
      
      case 'textarea':
        return `<div class="app-widget"${styleAttr}>${labelHtml}<textarea data-component="${id}" rows="3" placeholder="${placeholder || 'Enter text...'}"></textarea></div>`;
      
      case 'dropdown':
        const optsHtml = options.length 
          ? options.map(opt => `<option value="${escapeHtml(opt)}">${escapeHtml(opt)}</option>`).join('')
          : '';
        return `<div class="app-widget"${styleAttr}>${labelHtml}<select data-component="${id}"><option value="">Select...</option>${optsHtml}</select></div>`;
      
      case 'checkbox':
        return `<div class="app-widget checkbox"${styleAttr}><label><input type="checkbox" data-component="${id}" /> ${label}</label></div>`;
      
      case 'radio':
        const radioOpts = options.length ? options : ['Option 1', 'Option 2'];
        const radioHtml = radioOpts.map(opt => `<label class="radio-option"><input type="radio" name="radio-${id}" value="${escapeHtml(opt)}" data-component="${id}" /> ${escapeHtml(opt)}</label>`).join('');
        return `<div class="app-widget radio"${styleAttr}><div class="radio-group">${radioHtml}</div></div>`;
      
      case 'slider':
        const min = props.min ?? 0;
        const max = props.max ?? 100;
        const step = props.step ?? 1;
        const sliderDefault = props.default ?? ((min + max) / 2);
        return `<div class="app-widget"${styleAttr}>${labelHtml}<input type="range" data-component="${id}" min="${min}" max="${max}" step="${step}" value="${sliderDefault}" /></div>`;
      
      case 'spinctrl':
        const spinMin = props.min ?? 0;
        const spinMax = props.max ?? 100;
        const spinStep = props.step ?? 1;
        return `<div class="app-widget"${styleAttr}>${labelHtml}<input data-component="${id}" type="number" min="${spinMin}" max="${spinMax}" step="${spinStep}" value="${props.default || spinMin}" /></div>`;
      
      case 'datepicker':
        return `<div class="app-widget"${styleAttr}>${labelHtml}<input data-component="${id}" type="date" /></div>`;
      
      case 'colorpicker':
        return `<div class="app-widget"${styleAttr}>${labelHtml}<input data-component="${id}" type="color" value="#5ab3ff" /></div>`;
      
      case 'button':
        return `<div class="app-widget"${styleAttr}><button type="button" class="widget-button" data-component="${id}" data-action="${escapeHtml(props.onclick || '')}">${label}</button></div>`;
      
      case 'togglebtn':
        return `<div class="app-widget"${styleAttr}><button type="button" class="widget-toggle" data-component="${id}">${label}</button></div>`;
      
      case 'link':
        const href = props.target || '#';
        return `<div class="app-widget"${styleAttr}><a href="${escapeHtml(href)}" class="widget-link" data-component="${id}">${label}</a></div>`;
      
      case 'gauge':
        const gaugeMin = props.min ?? 0;
        const gaugeMax = props.max ?? 100;
        return `<div class="app-widget"${styleAttr}>${labelHtml}<progress data-component="${id}" value="${props.default || 50}" max="${gaugeMax}"></progress></div>`;
      
      case 'image':
        return `<div class="app-widget"${styleAttr}><div class="widget-image" data-component="${id}">🖼️</div></div>`;
      
      case 'separator':
        return `<hr class="widget-separator"${styleAttr} />`;
      
      case 'spacer':
        const spacerH = props.height || '16px';
        return `<div class="widget-spacer" style="height: ${spacerH};"${styleAttr}></div>`;
      
      case 'datagrid':
        return `<div class="app-widget"${styleAttr}><div class="widget-datagrid" data-component="${id}"><table><tr><th>A</th><th>B</th><th>C</th></tr><tr><td>1</td><td>2</td><td>3</td></tr></table></div></div>`;
      
      case 'chart':
        return `<div class="app-widget"${styleAttr}><div class="widget-chart" data-component="${id}">📈</div></div>`;
      
      case 'formula':
        return `<div class="app-widget"${styleAttr}>${labelHtml}<output class="widget-formula" data-component="${id}">= result</output></div>`;
      
      default:
        return `<div class="app-widget"${styleAttr}><span>${def.icon} ${label}</span></div>`;
    }
  }

  function renderTabBasedAppUi(tabs) {
    if (!appUiForm || !appUiEmpty) return;
    if (!tabs.length) {
      appUiForm.classList.add('hidden');
      appUiEmpty.classList.remove('hidden');
      appUiForm.innerHTML = '';
      activeComponents = [];
      return;
    }
    
    appUiEmpty.classList.add('hidden');
    appUiForm.classList.remove('hidden');
    
    // Flatten all components from all tabs for refresh purposes
    activeComponents = tabs.flatMap(tab => normalizeComponents(tab.components || []));
    
    // Create tab interface
    const tabsHtml = tabs.map((tab, idx) => {
      const tabId = `app-tab-${idx}`;
      return `<button type="button" class="app-tab-btn${idx === 0 ? ' active' : ''}" data-tab-target="${tabId}">${escapeHtml(tab.name)}</button>`;
    }).join('');
    
    const tabPanelsHtml = tabs.map((tab, idx) => {
      const tabId = `app-tab-${idx}`;
      const components = normalizeComponents(tab.components || []);
      const componentsHtml = components.map(comp => renderComponent(comp)).join('');
      return `<div class="app-tab-panel${idx === 0 ? ' active' : ''}" data-tab-id="${tabId}">${componentsHtml}</div>`;
    }).join('');
    
    appUiForm.innerHTML = `<div class="app-tabs">${tabsHtml}</div><div class="app-tab-content">${tabPanelsHtml}</div>`;
    
    // Add tab switching logic
    appUiForm.querySelectorAll('.app-tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const target = btn.dataset.tabTarget;
        appUiForm.querySelectorAll('.app-tab-btn').forEach(b => b.classList.remove('active'));
        appUiForm.querySelectorAll('.app-tab-panel').forEach(p => p.classList.remove('active'));
        btn.classList.add('active');
        const panel = appUiForm.querySelector(`[data-tab-id="${target}"]`);
        if (panel) panel.classList.add('active');
      });
    });
  }

  function renderLegacyAppUi(components) {
    activeComponents = components;
    if (!appUiForm || !appUiEmpty) return;
    if (!components.length) {
      appUiForm.classList.add('hidden');
      appUiEmpty.classList.remove('hidden');
      appUiForm.innerHTML = '';
      return;
    }
    appUiEmpty.classList.add('hidden');
    appUiForm.classList.remove('hidden');
    appUiForm.innerHTML = components.map(comp => renderComponent(comp)).join('');
  }

  function renderComponent(comp) {
    const type = comp.componentType || 'text';
    const label = escapeHtml(comp.label || comp.cell || 'Field');
    const cellRef = `${escapeHtml(comp.sheet || '')}!${escapeHtml(comp.cell || '')}`;
    
    if (type === 'display') {
      return `<div class="app-ui-field"><label>${label}</label><output data-component="${comp.id}">—</output></div>`;
    }
    
    if (type === 'textarea') {
      return `<div class="app-ui-field"><label>${label}</label><textarea data-component="${comp.id}" rows="4"></textarea></div>`;
    }
    
    if (type === 'dropdown') {
      const options = (comp.options || '').split(',').map(o => o.trim()).filter(Boolean);
      const optionsHtml = options.map(opt => `<option value="${escapeHtml(opt)}">${escapeHtml(opt)}</option>`).join('');
      return `<div class="app-ui-field"><label>${label}</label><select data-component="${comp.id}"><option value="">Select...</option>${optionsHtml}</select></div>`;
    }
    
    if (type === 'slider') {
      const min = comp.min ?? 0;
      const max = comp.max ?? 100;
      const step = comp.step ?? 1;
      return `<div class="app-ui-field"><label>${label} <span class="slider-value" data-value-for="${comp.id}">—</span></label><input type="range" data-component="${comp.id}" min="${min}" max="${max}" step="${step}" /></div>`;
    }
    
    if (type === 'checkbox') {
      return `<div class="app-ui-field checkbox"><label><input type="checkbox" data-component="${comp.id}" /> ${label}</label></div>`;
    }
    
    if (type === 'number') {
      return `<div class="app-ui-field"><label>${label}</label><input data-component="${comp.id}" type="number" /></div>`;
    }
    
    return `<div class="app-ui-field"><label>${label}</label><input data-component="${comp.id}" type="text" /></div>`;
  }

  function handleAppFormChange(e) {
    const el = e.target.closest('[data-component]');
    if (!el) return;
    const id = el.dataset.component;
    const component = activeComponents.find(c => c.id === id);
    if (!component || component.componentType === 'display') return;
    
    let value;
    if (el.type === 'checkbox') {
      value = el.checked ? 'TRUE' : 'FALSE';
    } else if (el.type === 'range') {
      value = el.value;
      const valueSpan = appUiForm?.querySelector(`[data-value-for="${id}"]`);
      if (valueSpan) valueSpan.textContent = value;
    } else {
      value = el.value;
    }
    
    commitComponentValue(component, value);
  }

  async function commitComponentValue(component, rawValue) {
    if (!component?.sheet || !component?.cell) return;
    const payload = { sheet: component.sheet, range: component.cell };
    const type = component.componentType || 'text';
    
    if (type === 'number' || type === 'slider') {
      const trimmed = (rawValue ?? '').trim();
      if (trimmed === '') return showToast('Enter a number', true);
      const num = Number(trimmed);
      if (!Number.isFinite(num)) return showToast('Enter a valid number', true);
      payload.value_number = num;
    } else if (type === 'checkbox') {
      payload.value = rawValue;
    } else {
      payload.value = rawValue ?? '';
    }
    
    try {
      const res = await apiFetch(`${apiBase}/excel/set`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', ...authHeaders() },
        body: JSON.stringify(payload)
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data?.error || 'Failed to push value');
      await refreshActiveApp();
    } catch (err) {
      showToast(err.message || 'Failed to push value', true);
    }
  }

  function updateWorkspaceState({ title, desc, busy }) {
    if (title && appWorkspaceTitle) appWorkspaceTitle.textContent = title;
    if (desc && appWorkspaceDesc) appWorkspaceDesc.textContent = desc;
    if (busy) {
      if (appRefreshBtn) appRefreshBtn.disabled = true;
      if (appExitBtn) appExitBtn.disabled = true;
    }
  }

  async function exitActiveApp() {
    await closeWorkbook();
    resetWorkspace();
    showToast('Workbook session closed');
  }

  function showAppOverlay() {
    if (appOverlay) {
      appOverlay.classList.remove('hidden');
      appOverlay.setAttribute('aria-hidden', 'false');
      bodyEl.classList.add('modal-open');
    }
  }

  function hideAppOverlay() {
    if (appOverlay) {
      appOverlay.classList.add('hidden');
      appOverlay.setAttribute('aria-hidden', 'true');
      bodyEl.classList.remove('modal-open');
    }
  }

  async function fetchAppSchema(owner, name) {
    const res = await apiFetch(`${apiBase}/apps/ui/get`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...authHeaders() },
      body: JSON.stringify({ owner, name })
    });
    if (!res.ok) throw new Error('Unable to load UI schema');
    const data = await res.json();
    return data.schema || { components: [] };
  }

  async function fetchWorkbookSheets(owner, name) {
    const res = await apiFetch(`${apiBase}/excel/sheets`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...authHeaders() },
      body: JSON.stringify({ owner, name })
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data?.error || 'Unable to load sheets');
    return Array.isArray(data?.sheets) ? data.sheets : [];
  }

  async function openBuilder(owner, name) {
    if (!token) return showToast('Login first', true);
    if (!canAuthorApp(owner)) return showToast('Developer access required', true);
    closeDevModal();
    try {
      showSheetLoadingState();
      const [schema, sheets] = await Promise.all([
        fetchAppSchema(owner, name),
        fetchWorkbookSheets(owner, name)
      ]);
      builderTarget = { owner, name };
      
      // Load or create widget tree from schema
      if (schema?.widgets) {
        // New widget format
        builderWidgets = normalizeWidgets(schema.widgets);
      } else if (schema?.tabs) {
        // Migrate from old tab format
        builderWidgets = migrateTabsToWidgets(schema.tabs);
      } else if (schema?.components?.length) {
        // Very old component format
        builderWidgets = migrateComponentsToWidgets(schema.components);
      } else {
        builderWidgets = [];
      }
      
      builderSelectedWidget = null;
      builderSavedState = JSON.stringify(builderWidgets);
      setSheetOptions(sheets);
      if (builderAppLabel) builderAppLabel.textContent = `${owner}/${name}`;
      renderWidgetTree();
      hidePropertyForm();
      openBuilderModal();
    } catch (err) {
      showToast(err.message || 'Unable to load builder', true);
    }
  }

  function normalizeWidgets(widgets) {
    if (!Array.isArray(widgets)) return [];
    return widgets.map(w => ({
      ...w,
      id: w.id || createWidgetId(),
      children: normalizeWidgets(w.children || [])
    }));
  }

  function migrateTabsToWidgets(tabs) {
    // Convert old tab format to notebook widget with panels
    if (!tabs || !tabs.length) return [];
    const notebook = {
      id: createWidgetId(),
      type: 'notebook',
      name: 'MainNotebook',
      children: tabs.map(tab => ({
        id: tab.id || createWidgetId(),
        type: 'panel',
        name: tab.name,
        properties: { label: tab.name },
        children: (tab.components || []).map(comp => migrateComponentToWidget(comp))
      }))
    };
    return [notebook];
  }

  function migrateComponentsToWidgets(components) {
    // Wrap old components in a VBox
    return [{
      id: createWidgetId(),
      type: 'boxsizer-v',
      name: 'MainLayout',
      children: components.map(comp => migrateComponentToWidget(comp))
    }];
  }

  function migrateComponentToWidget(comp) {
    const typeMap = {
      text: 'textinput',
      number: 'number',
      textarea: 'textarea',
      dropdown: 'dropdown',
      slider: 'slider',
      checkbox: 'checkbox',
      display: 'output'
    };
    return {
      id: comp.id || createWidgetId(),
      type: typeMap[comp.componentType] || 'textinput',
      name: comp.label || comp.cell || 'Widget',
      properties: {
        label: comp.label || '',
        options: comp.options || '',
        min: comp.min,
        max: comp.max,
        step: comp.step
      },
      excel: {
        enabled: true,
        sheet: comp.sheet || '',
        cell: comp.cell || '',
        mode: 'bidirectional'
      },
      children: []
    };
  }

  function createWidgetId() {
    return 'w-' + Math.random().toString(36).slice(2, 9);
  }

  // Get valid child widget types for a given parent type
  function getValidChildWidgets(parentType) {
    // All widget types can be added to the app root or containers
    const allWidgets = Object.entries(WIDGET_TYPES).map(([key, def]) => ({
      type: key,
      icon: def.icon,
      label: def.label,
      category: def.category
    }));
    
    // Non-container widgets cannot have children
    if (parentType && !WIDGET_TYPES[parentType]?.isContainer) {
      return [];
    }
    
    return allWidgets;
  }

  function renderAddWidgetDropdown(parentId, parentType) {
    const validWidgets = getValidChildWidgets(parentType);
    if (!validWidgets.length) return '';
    
    const optionsHtml = validWidgets.map(w => 
      `<option value="${w.type}">${w.icon} ${w.label}</option>`
    ).join('');
    
    return `<select class="tree-add-widget" data-parent-id="${parentId}" title="Add child widget">
      <option value="">+</option>
      ${optionsHtml}
    </select>`;
  }

  function renderWidgetTree() {
    if (!builderTreeRoot) return;
    
    const appName = builderTarget?.name || 'Application';
    const appCollapsed = builderAppCollapsed || false;
    const appSelected = builderSelectedWidget === 'app-root';
    
    // Always show app node, even with no widgets
    builderEmpty?.classList.add('hidden');
    
    const widgetsHtml = builderWidgets.map(w => renderWidgetNode(w, 1)).join('');
    const emptyHint = !builderWidgets.length 
      ? '<div class="tree-empty-hint" style="padding-left: 24px; opacity: 0.6; font-size: 12px; margin: 8px 0;">Use the + dropdown to add widgets.</div>' 
      : '';
    
    // App root can have any widgets added to it
    const appAddDropdown = renderAddWidgetDropdown('app-root', null);
    
    builderTreeRoot.innerHTML = `
      <div class="tree-widget-group tree-app-root">
        <div class="tree-widget-node tree-app-node ${appSelected ? 'selected' : ''}" data-widget-id="app-root">
          <button type="button" class="tree-expand-btn" data-toggle-app>
            <span class="tree-expand-icon">${appCollapsed ? '▶' : '▼'}</span>
          </button>
          <span class="widget-icon">�️</span>
          <span class="widget-name">${escapeHtml(appName)}</span>
          <span class="widget-type">Application</span>
          <div class="tree-node-actions">
            ${appAddDropdown}
          </div>
        </div>
        <div class="tree-children ${appCollapsed ? 'hidden' : ''}">
          ${widgetsHtml}
          ${emptyHint}
        </div>
      </div>
    `;
    
    // Update live preview
    renderBuilderPreview();
  }

  let builderAppCollapsed = false;

  function renderWidgetNode(widget, depth) {
    const def = WIDGET_TYPES[widget.type] || { icon: '❓', label: widget.type };
    const isContainer = def.isContainer;
    const hasChildren = widget.children && widget.children.length > 0;
    const collapsed = widget.collapsed || false;
    const selected = builderSelectedWidget === widget.id;
    
    const indent = depth * 10;
    const childrenHtml = isContainer && hasChildren && !collapsed
      ? `<div class="tree-children">${widget.children.map(c => renderWidgetNode(c, depth + 1)).join('')}</div>`
      : '';
    
    const expandBtn = isContainer
      ? `<button type="button" class="tree-expand-btn" data-toggle-widget="${widget.id}">
           <span class="tree-expand-icon">${collapsed ? '▶' : '▼'}</span>
         </button>`
      : '<span class="tree-expand-spacer"></span>';
    
    // Add dropdown for containers, delete button for all widgets
    const addDropdown = isContainer ? renderAddWidgetDropdown(widget.id, widget.type) : '';
    const moveUpBtn = `<button type="button" class="tree-move-btn" data-move-widget="${widget.id}" data-direction="up" title="Move up">↑</button>`;
    const moveDownBtn = `<button type="button" class="tree-move-btn" data-move-widget="${widget.id}" data-direction="down" title="Move down">↓</button>`;
    const deleteBtn = `<button type="button" class="tree-delete-btn" data-delete-widget="${widget.id}" title="Delete widget">×</button>`;
    
    return `<div class="tree-widget-group">
      <div class="tree-widget-node ${selected ? 'selected' : ''}" 
           data-widget-id="${widget.id}" 
           style="padding-left: ${indent}px">
        ${expandBtn}
        <span class="widget-icon">${def.icon}</span>
        <span class="widget-name">${escapeHtml(widget.name || def.label)}</span>
        <span class="widget-type">${escapeHtml(def.label)}</span>
        <div class="tree-node-actions">
          ${moveUpBtn}
          ${moveDownBtn}
          ${addDropdown}
          ${deleteBtn}
        </div>
      </div>
      ${childrenHtml}
    </div>`;
  }

  // Live preview rendering
  function renderBuilderPreview() {
    if (!builderPreviewContent || !builderPreviewEmpty) return;
    
    if (!builderWidgets || !builderWidgets.length) {
      builderPreviewContent.innerHTML = '';
      builderPreviewEmpty.classList.remove('hidden');
      return;
    }
    
    builderPreviewEmpty.classList.add('hidden');
    builderPreviewContent.innerHTML = builderWidgets.map(w => renderPreviewWidget(w)).join('');
  }

  function getWidgetStyles(props) {
    const styles = [];
    if (props.width && props.width !== 'auto') styles.push(`width: ${props.width}px`);
    if (props.height && props.height !== 'auto') styles.push(`height: ${props.height}px`);
    if (props.minWidth) styles.push(`min-width: ${props.minWidth}px`);
    if (props.minHeight) styles.push(`min-height: ${props.minHeight}px`);
    if (props.margin) styles.push(`margin: ${props.margin}px`);
    if (props.padding) styles.push(`padding: ${props.padding}px`);
    if (props.proportion && props.proportion > 0) styles.push(`flex: ${props.proportion}`);
    return styles.join('; ');
  }

  // Track active tab for each notebook in preview
  let previewNotebookTabs = {};

  function renderPreviewNotebook(widget, selected) {
    const selectedClass = selected ? ' selected' : '';
    const children = widget.children || [];
    
    // Each child of a notebook is a tab/page (typically panels)
    // Get the active tab index for this notebook
    const activeTabIdx = previewNotebookTabs[widget.id] || 0;
    
    if (!children.length) {
      return `<div class="preview-notebook${selectedClass}" data-preview-id="${widget.id}">
        <div class="notebook-tabs"><div class="notebook-tab-empty">No tabs</div></div>
        <div class="notebook-content"><div class="preview-empty-hint">(empty notebook)</div></div>
      </div>`;
    }
    
    // Build tabs from children - use child's label or name for tab title
    const tabsHtml = children.map((child, idx) => {
      const childProps = child.properties || {};
      const tabLabel = escapeHtml(childProps.label || child.name || `Tab ${idx + 1}`);
      const activeClass = idx === activeTabIdx ? ' active' : '';
      return `<button type="button" class="notebook-tab${activeClass}" data-notebook-id="${widget.id}" data-tab-idx="${idx}">${tabLabel}</button>`;
    }).join('');
    
    // Render only the active tab's content
    const activeChild = children[activeTabIdx];
    const contentHtml = activeChild ? renderPreviewWidget(activeChild) : '';
    
    return `<div class="preview-notebook${selectedClass}" data-preview-id="${widget.id}">
      <div class="notebook-tabs">${tabsHtml}</div>
      <div class="notebook-content">${contentHtml}</div>
    </div>`;
  }

  function renderPreviewWidget(widget) {
    const def = WIDGET_TYPES[widget.type] || { icon: '❓', label: widget.type };
    const props = widget.properties || {};
    const label = escapeHtml(props.label || widget.name || def.label);
    const selected = builderSelectedWidget === widget.id;
    const selectedClass = selected ? ' selected' : '';
    const widgetStyles = getWidgetStyles(props);
    const styleAttr = widgetStyles ? ` style="${widgetStyles}"` : '';
    
    // Special handling for notebook (tabbed container)
    if (widget.type === 'notebook') {
      return renderPreviewNotebook(widget, selected);
    }
    
    // Container/Sizer widgets - render as invisible layout containers
    if (def.isContainer) {
      const childrenHtml = widget.children?.length 
        ? widget.children.map(c => renderPreviewWidget(c)).join('') 
        : '';
      
      // Determine layout type and styles
      let containerClass = 'preview-sizer';
      let containerStyles = [];
      
      if (widget.type === 'boxsizer-v') {
        containerClass += ' sizer-vertical';
        if (props.vgap) containerStyles.push(`gap: ${props.vgap}px`);
      } else if (widget.type === 'boxsizer-h') {
        containerClass += ' sizer-horizontal';
        if (props.hgap) containerStyles.push(`gap: ${props.hgap}px`);
      } else if (widget.type === 'gridsizer') {
        containerClass += ' sizer-grid';
        const cols = props.cols || 2;
        containerStyles.push(`grid-template-columns: repeat(${cols}, 1fr)`);
        if (props.hgap) containerStyles.push(`column-gap: ${props.hgap}px`);
        if (props.vgap) containerStyles.push(`row-gap: ${props.vgap}px`);
      } else if (widget.type === 'panel') {
        containerClass = 'preview-panel';
        containerStyles.push('flex-direction: column');
      } else if (widget.type === 'scrolled') {
        containerClass = 'preview-scrolled';
      }
      
      // Add widget-level styles
      if (props.margin) containerStyles.push(`margin: ${props.margin}px`);
      if (props.padding) containerStyles.push(`padding: ${props.padding}px`);
      if (props.proportion && props.proportion > 0) containerStyles.push(`flex: ${props.proportion}`);
      
      const styleStr = containerStyles.length ? ` style="${containerStyles.join('; ')}"` : '';
      
      // Empty state for visible containers
      const emptyHint = !widget.children?.length && (widget.type === 'panel' || widget.type === 'notebook') 
        ? '<div class="preview-empty-hint">(empty)</div>' : '';
      
      return `<div class="${containerClass}${selectedClass}" data-preview-id="${widget.id}"${styleStr}>
        ${childrenHtml}${emptyHint}
      </div>`;
    }
    
    // Non-container widgets - render as actual controls
    const controlHtml = renderPreviewControl(widget, def, props, label);
    
    return `<div class="preview-widget${selectedClass}" data-preview-id="${widget.id}"${styleAttr}>
      ${controlHtml}
    </div>`;
  }

  function renderPreviewControl(widget, def, props, label) {
    const placeholder = escapeHtml(props.placeholder || '');
    const options = (props.options || '').split(',').map(o => o.trim()).filter(Boolean);
    const labelHtml = props.label ? `<label class="preview-label">${escapeHtml(props.label)}</label>` : '';
    
    switch (widget.type) {
      case 'textinput':
        return `${labelHtml}<input type="text" placeholder="${placeholder || 'Enter text...'}">`;
      case 'number':
      case 'spinctrl':
        return `${labelHtml}<input type="number" value="${props.default || props.min || 0}" min="${props.min || 0}" max="${props.max || 100}">`;
      case 'textarea':
        return `${labelHtml}<textarea rows="3" placeholder="${placeholder || 'Enter text...'}"></textarea>`;
      case 'dropdown':
        const opts = options.length ? options : ['Select...'];
        return `${labelHtml}<select>${opts.map(o => `<option>${escapeHtml(o)}</option>`).join('')}</select>`;
      case 'checkbox':
        return `<label class="preview-checkbox"><input type="checkbox"> ${escapeHtml(props.label || label)}</label>`;
      case 'radio':
        const radioOpts = options.length ? options : ['Option 1', 'Option 2'];
        return `<div class="preview-radio-group">${radioOpts.map(o => `<label class="preview-radio"><input type="radio" name="${widget.id}"> ${escapeHtml(o)}</label>`).join('')}</div>`;
      case 'slider':
        return `${labelHtml}<input type="range" min="${props.min || 0}" max="${props.max || 100}" value="${props.default || ((props.min || 0) + (props.max || 100)) / 2}">`;
      case 'datepicker':
        return `${labelHtml}<input type="date">`;
      case 'colorpicker':
        return `${labelHtml}<input type="color" value="#5ab3ff">`;
      case 'button':
        return `<button type="button" class="preview-button">${escapeHtml(props.label || label)}</button>`;
      case 'togglebtn':
        return `<button type="button" class="preview-button toggle">${escapeHtml(props.label || label)}</button>`;
      case 'link':
        return `<a href="#" class="preview-link">${escapeHtml(props.label || label)}</a>`;
      case 'label':
        return `<span class="preview-text">${escapeHtml(props.label || label)}</span>`;
      case 'output':
        return `${labelHtml}<span class="preview-output">—</span>`;
      case 'image':
        return `<div class="preview-image">🖼️</div>`;
      case 'gauge':
        return `${labelHtml}<progress value="${props.default || 50}" max="${props.max || 100}"></progress>`;
      case 'separator':
        return `<hr class="preview-separator">`;
      case 'spacer':
        return `<div class="preview-spacer"></div>`;
      case 'datagrid':
        return `<div class="preview-datagrid"><table><tr><th>A</th><th>B</th><th>C</th></tr><tr><td>1</td><td>2</td><td>3</td></tr><tr><td>4</td><td>5</td><td>6</td></tr></table></div>`;
      case 'chart':
        return `<div class="preview-chart">📈</div>`;
      case 'formula':
        return `${labelHtml}<span class="preview-formula">= result</span>`;
      default:
        return `<span>${def.icon} ${escapeHtml(props.label || label)}</span>`;
    }
  }

  function getDefaultProperties(type) {
    const defaults = {
      slider: { min: 0, max: 100, step: 1 },
      spinctrl: { min: 0, max: 100, step: 1 },
      gauge: { min: 0, max: 100 },
      gridsizer: { cols: 2, rows: 0, hgap: 4, vgap: 4 },
      'boxsizer-v': { orientation: 'vertical', vgap: 4 },
      'boxsizer-h': { orientation: 'horizontal', hgap: 4 }
    };
    return defaults[type] || {};
  }

  function findWidgetById(widgets, id) {
    for (const w of widgets) {
      if (w.id === id) return w;
      if (w.children) {
        const found = findWidgetById(w.children, id);
        if (found) return found;
      }
    }
    return null;
  }

  function findWidgetParent(widgets, id, parent = null) {
    for (const w of widgets) {
      if (w.id === id) return parent;
      if (w.children) {
        const found = findWidgetParent(w.children, id, w);
        if (found !== undefined) return found;
      }
    }
    return undefined;
  }

  function removeWidgetById(widgets, id) {
    const idx = widgets.findIndex(w => w.id === id);
    if (idx >= 0) {
      widgets.splice(idx, 1);
      return true;
    }
    for (const w of widgets) {
      if (w.children && removeWidgetById(w.children, id)) return true;
    }
    return false;
  }

  function moveWidgetInSiblings(widgetId, direction) {
    // Find the widget and its parent array
    function findWidgetAndSiblings(widgets, id) {
      for (let i = 0; i < widgets.length; i++) {
        if (widgets[i].id === id) {
          return { siblings: widgets, index: i };
        }
        if (widgets[i].children) {
          const result = findWidgetAndSiblings(widgets[i].children, id);
          if (result) return result;
        }
      }
      return null;
    }
    
    const result = findWidgetAndSiblings(builderWidgets, widgetId);
    if (!result) return;
    
    const { siblings, index } = result;
    const newIndex = direction === 'up' ? index - 1 : index + 1;
    
    // Check bounds
    if (newIndex < 0 || newIndex >= siblings.length) return;
    
    // Swap positions
    [siblings[index], siblings[newIndex]] = [siblings[newIndex], siblings[index]];
    renderWidgetTree();
  }

  function handleWidgetTreeClick(e) {
    // Toggle app collapse/expand
    const toggleAppBtn = e.target.closest('[data-toggle-app]');
    if (toggleAppBtn) {
      builderAppCollapsed = !builderAppCollapsed;
      renderWidgetTree();
      return;
    }
    
    // Toggle widget expand/collapse
    const toggleBtn = e.target.closest('[data-toggle-widget]');
    if (toggleBtn) {
      const widgetId = toggleBtn.dataset.toggleWidget;
      const widget = findWidgetById(builderWidgets, widgetId);
      if (widget) {
        widget.collapsed = !widget.collapsed;
        renderWidgetTree();
      }
      return;
    }
    
    // Handle move button click
    const moveBtn = e.target.closest('[data-move-widget]');
    if (moveBtn) {
      e.stopPropagation();
      const widgetId = moveBtn.dataset.moveWidget;
      const direction = moveBtn.dataset.direction;
      moveWidgetInSiblings(widgetId, direction);
      return;
    }
    
    // Handle delete button click
    const deleteBtn = e.target.closest('[data-delete-widget]');
    if (deleteBtn) {
      e.stopPropagation();
      const widgetId = deleteBtn.dataset.deleteWidget;
      if (confirm('Delete this widget and all its children?')) {
        removeWidgetById(builderWidgets, widgetId);
        if (builderSelectedWidget === widgetId) {
          builderSelectedWidget = null;
          hidePropertyForm();
        }
        renderWidgetTree();
        showToast('Widget deleted');
      }
      return;
    }
    
    // Select widget or app node (but not if clicking on actions)
    if (e.target.closest('.tree-node-actions')) return;
    
    const node = e.target.closest('.tree-widget-node');
    if (node) {
      const widgetId = node.dataset.widgetId;
      builderSelectedWidget = widgetId;
      renderWidgetTree();
      
      if (widgetId === 'app-root') {
        showAppPropertyForm();
      } else {
        const widget = findWidgetById(builderWidgets, widgetId);
        if (widget) showPropertyForm(widget);
      }
    }
  }

  // Handle add widget dropdown change
  function handleAddWidgetDropdown(e) {
    const select = e.target.closest('.tree-add-widget');
    if (!select) return;
    
    const widgetType = select.value;
    if (!widgetType || !WIDGET_TYPES[widgetType]) {
      select.value = '';
      return;
    }
    
    const parentId = select.dataset.parentId;
    const def = WIDGET_TYPES[widgetType];
    
    const newWidget = {
      id: createWidgetId(),
      type: widgetType,
      name: def.label,
      properties: getDefaultProperties(widgetType),
      excel: { enabled: false, sheet: '', cell: '', mode: 'bidirectional' },
      children: []
    };
    
    if (parentId === 'app-root') {
      builderWidgets.push(newWidget);
    } else {
      const parent = findWidgetById(builderWidgets, parentId);
      if (parent && WIDGET_TYPES[parent.type]?.isContainer) {
        parent.children = parent.children || [];
        parent.children.push(newWidget);
      }
    }
    
    // Select the new widget
    builderSelectedWidget = newWidget.id;
    renderWidgetTree();
    showPropertyForm(newWidget);
    
    // Reset dropdown
    select.value = '';
  }

  async function showAppPropertyForm() {
    if (!builderAppForm || !builderNoSelection) return;
    
    // Hide other forms, show app form
    builderNoSelection.classList.add('hidden');
    builderPropertyForm?.classList.add('hidden');
    builderAppForm.classList.remove('hidden');
    
    // Find app in cache
    const app = appsCache.find(a => a.name === builderTarget?.name && a.owner === builderTarget?.owner);
    
    if (propAppName) propAppName.value = builderTarget?.name || '';
    if (propAppOwner) propAppOwner.value = builderTarget?.owner || '';
    if (propAppDescription) propAppDescription.value = app?.description || '';
    if (propAppPublic) propAppPublic.checked = app?.public || false;
    if (propAppAccessGroup) propAppAccessGroup.value = app?.access_group || '';
  }

  async function handleAppFormSubmit(e) {
    e.preventDefault();
    if (!builderTarget) return showToast('No app selected', true);
    
    const payload = {
      description: propAppDescription?.value || '',
      public: propAppPublic?.checked || false
    };
    
    const accessGroup = propAppAccessGroup?.value?.trim();
    if (accessGroup) payload.access_group = accessGroup;
    
    try {
      const res = await apiFetch(`${apiBase}/apps/${encodeURIComponent(builderTarget.name)}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json', ...authHeaders() },
        body: JSON.stringify(payload)
      });
      if (!res.ok) throw new Error('Failed to update app');
      showToast('Application updated');
      await refreshApps();
    } catch (err) {
      showToast(err.message || 'Failed to update app', true);
    }
  }

  async function handleDeleteApp() {
    if (!builderTarget) return showToast('No app selected', true);
    
    const confirmMsg = `Are you sure you want to delete "${builderTarget.name}"? This cannot be undone.`;
    if (!confirm(confirmMsg)) return;
    
    try {
      const res = await apiFetch(`${apiBase}/apps/${encodeURIComponent(builderTarget.name)}`, {
        method: 'DELETE',
        headers: authHeaders()
      });
      if (!res.ok) throw new Error('Failed to delete app');
      showToast('Application deleted');
      closeBuilderModal(true);
      await refreshApps();
    } catch (err) {
      showToast(err.message || 'Failed to delete app', true);
    }
  }

  function showPropertyForm(widget) {
    if (!builderPropertyForm || !builderNoSelection) return;
    
    // Hide other forms, show widget property form
    builderNoSelection.classList.add('hidden');
    builderAppForm?.classList.add('hidden');
    builderPropertyForm.classList.remove('hidden');
    
    const def = WIDGET_TYPES[widget.type] || {};
    const props = widget.properties || {};
    const excel = widget.excel || { enabled: false };
    
    // Fill common fields
    if (propWidgetId) propWidgetId.value = widget.id;
    if (propName) propName.value = widget.name || '';
    const typeLabel = def.label || widget.type;
    if (propWidgetLegend) propWidgetLegend.textContent = `Widget - ${typeLabel}`;
    if (propLabel) propLabel.value = props.label || '';
    if (propPlaceholder) propPlaceholder.value = props.placeholder || '';
    if (propTooltip) propTooltip.value = props.tooltip || '';
    
    // Layout
    if (propWidth) propWidth.value = props.width || '';
    if (propHeight) propHeight.value = props.height || '';
    if (propMinWidth) propMinWidth.value = props.minWidth || '';
    if (propMinHeight) propMinHeight.value = props.minHeight || '';
    if (propProportion) propProportion.value = props.proportion ?? 0;
    if (propMargin) propMargin.value = props.margin ?? 4;
    if (propPadding) propPadding.value = props.padding ?? 0;
    
    // Flags
    if (propFlags) {
      const flags = props.flags || [];
      Array.from(propFlags.options).forEach(opt => {
        opt.selected = flags.includes(opt.value);
      });
    }
    
    // Sizer options
    if (propOrientation) propOrientation.value = props.orientation || 'vertical';
    if (propCols) propCols.value = props.cols ?? 2;
    if (propRows) propRows.value = props.rows ?? 0;
    if (propHgap) propHgap.value = props.hgap ?? 4;
    if (propVgap) propVgap.value = props.vgap ?? 4;
    
    // Input options
    if (propOptions) propOptions.value = props.options || '';
    if (propMin) propMin.value = props.min ?? 0;
    if (propMax) propMax.value = props.max ?? 100;
    if (propStep) propStep.value = props.step ?? 1;
    if (propDefault) propDefault.value = props.defaultValue || '';
    
    // Excel binding
    if (propExcelEnabled) propExcelEnabled.checked = excel.enabled;
    if (propExcelFields) propExcelFields.classList.toggle('hidden', !excel.enabled);
    if (propSheet) propSheet.value = excel.sheet || '';
    if (propCell) propCell.value = excel.cell || '';
    if (propMode) propMode.value = excel.mode || 'bidirectional';
    
    // Action
    if (propOnclick) propOnclick.value = props.onclick || '';
    if (propTarget) propTarget.value = props.target || '';
    
    // Show/hide sections based on widget type
    updatePropertyVisibility(widget.type);
  }

  function updatePropertyVisibility(type) {
    const def = WIDGET_TYPES[type] || {};
    const category = def.category;
    const isContainer = def.isContainer;
    
    // Label/content section - show for most widgets
    if (propGroupLabel) {
      propGroupLabel.classList.toggle('hidden', type === 'separator' || type === 'spacer');
    }
    
    // Sizer options - only for container types
    if (propGroupSizer) {
      propGroupSizer.classList.toggle('hidden', !isContainer);
    }
    
    // Grid-specific options
    const propGridOptions = qs('#prop-grid-options');
    if (propGridOptions) {
      propGridOptions.classList.toggle('hidden', type !== 'gridsizer');
    }
    
    // Input options - only for input types with options/range
    if (propGroupInput) {
      const hasOptions = ['dropdown', 'radio'].includes(type);
      const hasRange = ['slider', 'spinctrl', 'gauge', 'number'].includes(type);
      propGroupInput.classList.toggle('hidden', !hasOptions && !hasRange);
      
      const optionsRow = qs('#prop-options-row');
      const rangeOptions = qs('#prop-range-options');
      const stepRow = qs('#prop-step-row');
      if (optionsRow) optionsRow.classList.toggle('hidden', !hasOptions);
      if (rangeOptions) rangeOptions.classList.toggle('hidden', !hasRange);
      if (stepRow) stepRow.classList.toggle('hidden', !['slider', 'spinctrl'].includes(type));
    }
    
    // Excel binding - available for most widgets except pure containers/decorators
    if (propGroupExcel) {
      const noExcel = ['separator', 'spacer', 'boxsizer-v', 'boxsizer-h', 'gridsizer'].includes(type);
      propGroupExcel.classList.toggle('hidden', noExcel);
    }
    
    // Action section - only for buttons/links
    if (propGroupAction) {
      propGroupAction.classList.toggle('hidden', category !== 'action');
    }
    
    // Placeholder row - only for text inputs
    const placeholderRow = qs('#prop-placeholder-row');
    if (placeholderRow) {
      placeholderRow.classList.toggle('hidden', !['textinput', 'textarea', 'number'].includes(type));
    }
  }

  function handleExcelToggle(e) {
    if (propExcelFields) {
      propExcelFields.classList.toggle('hidden', !e.target.checked);
    }
  }

  function handlePropertyFormSubmit(e) {
    e.preventDefault();
    
    const widgetId = propWidgetId?.value;
    if (!widgetId) return;
    
    const widget = findWidgetById(builderWidgets, widgetId);
    if (!widget) return;
    
    // Update widget from form
    widget.name = propName?.value || widget.name;
    widget.properties = widget.properties || {};
    widget.properties.label = propLabel?.value || '';
    widget.properties.placeholder = propPlaceholder?.value || '';
    widget.properties.tooltip = propTooltip?.value || '';
    widget.properties.width = propWidth?.value || '';
    widget.properties.height = propHeight?.value || '';
    widget.properties.minWidth = propMinWidth?.value || '';
    widget.properties.minHeight = propMinHeight?.value || '';
    widget.properties.proportion = Number(propProportion?.value ?? 0);
    widget.properties.margin = Number(propMargin?.value ?? 4);
    widget.properties.padding = Number(propPadding?.value ?? 0);
    
    // Flags
    if (propFlags) {
      widget.properties.flags = Array.from(propFlags.selectedOptions).map(o => o.value);
    }
    
    // Sizer
    widget.properties.orientation = propOrientation?.value || 'vertical';
    widget.properties.cols = Number(propCols?.value ?? 2);
    widget.properties.rows = Number(propRows?.value ?? 0);
    widget.properties.hgap = Number(propHgap?.value ?? 4);
    widget.properties.vgap = Number(propVgap?.value ?? 4);
    
    // Input
    widget.properties.options = propOptions?.value || '';
    widget.properties.min = Number(propMin?.value ?? 0);
    widget.properties.max = Number(propMax?.value ?? 100);
    widget.properties.step = Number(propStep?.value ?? 1);
    widget.properties.defaultValue = propDefault?.value || '';
    
    // Excel
    widget.excel = widget.excel || {};
    widget.excel.enabled = propExcelEnabled?.checked || false;
    widget.excel.sheet = propSheet?.value || '';
    widget.excel.cell = propCell?.value || '';
    widget.excel.mode = propMode?.value || 'bidirectional';
    
    // Action
    widget.properties.onclick = propOnclick?.value || '';
    widget.properties.target = propTarget?.value || '';
    
    renderWidgetTree();
    showToast('Properties updated');
  }

  function handleDeleteWidget() {
    const widgetId = propWidgetId?.value;
    if (!widgetId) return;
    
    if (!confirm('Delete this widget and all its children?')) return;
    
    removeWidgetById(builderWidgets, widgetId);
    builderSelectedWidget = null;
    hidePropertyForm();
    renderWidgetTree();
    showToast('Widget deleted');
  }

  function hidePropertyForm() {
    if (builderPropertyForm) builderPropertyForm.classList.add('hidden');
    if (builderAppForm) builderAppForm.classList.add('hidden');
    if (builderNoSelection) builderNoSelection.classList.remove('hidden');
  }

  function showSheetLoadingState() {
    if (!propSheet) return;
    propSheet.disabled = true;
    propSheet.innerHTML = '<option value="">Loading sheets…</option>';
    propSheet.value = '';
  }

  function setSheetOptions(sheets, preferred) {
    builderSheets = Array.isArray(sheets) ? sheets.filter(Boolean) : [];
    renderSheetOptions(preferred);
    renderImportSheetOptions();
  }

  function renderImportSheetOptions() {
    if (!builderImportSheet) return;
    if (!builderSheets.length) {
      builderImportSheet.innerHTML = '<option value="">No sheets</option>';
      builderImportSheet.disabled = true;
      return;
    }
    builderImportSheet.disabled = false;
    builderImportSheet.innerHTML = '<option value="">Sheet...</option>' +
      builderSheets.map(sheet => `<option value="${escapeHtml(sheet)}">${escapeHtml(sheet)}</option>`).join('');
  }

  function ensureSheetOption(sheet) {
    if (!sheet) return;
    if (!builderSheets.includes(sheet)) {
      builderSheets.push(sheet);
    }
    renderSheetOptions(sheet);
  }

  function renderSheetOptions(preferred) {
    if (!propSheet) return;
    const targetValue = preferred ?? propSheet.value;
    if (!builderSheets.length) {
      propSheet.disabled = true;
      propSheet.innerHTML = '<option value="">No sheets found</option>';
      propSheet.value = '';
      return;
    }
    propSheet.disabled = false;
    propSheet.innerHTML = builderSheets
      .map(sheet => `<option value="${escapeHtml(sheet)}">${escapeHtml(sheet)}</option>`)
      .join('');
    if (targetValue && builderSheets.includes(targetValue)) {
      propSheet.value = targetValue;
    } else {
      propSheet.value = builderSheets[0];
    }
  }

  async function saveBuilderLayout() {
    if (!builderTarget) return showToast('Select an app to design', true);
    const schema = { widgets: builderWidgets };
    try {
      const res = await apiFetch(`${apiBase}/apps/ui/save`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', ...authHeaders() },
        body: JSON.stringify({ owner: builderTarget.owner, name: builderTarget.name, schema_json: JSON.stringify(schema) })
      });
      if (!res.ok) throw new Error('Failed to save layout');
      builderSavedState = JSON.stringify(builderWidgets);
      showToast('Layout saved');
    } catch (err) {
      showToast(err.message || 'Failed to save layout', true);
    }
  }

  async function handleExcelImport() {
    if (!builderTarget) return showToast('No app selected', true);
    const sheet = builderImportSheet?.value;
    const range = builderImportRange?.value?.trim();
    if (!sheet) return showToast('Select a sheet first', true);
    if (!range) return showToast('Enter a cell range (e.g., A1:C5)', true);
    
    try {
      const res = await apiFetch(`${apiBase}/excel/analyze`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', ...authHeaders() },
        body: JSON.stringify({
          owner: builderTarget.owner,
          name: builderTarget.name,
          sheet,
          range
        })
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data?.error || 'Failed to analyze Excel range');
      
      if (!data.cells?.length) {
        return showToast('No cells found in range', true);
      }
      
      // Create a grid sizer with widgets matching the Excel layout
      const gridWidget = createGridFromCells(data.cells, data.rowCount, data.colCount, sheet);
      
      // Add the grid to the builder
      addImportedWidgets([gridWidget]);
      const widgetCount = data.cells.filter(c => c.type !== 'empty').length;
      showToast(`Imported ${data.rowCount}×${data.colCount} grid with ${widgetCount} widget${widgetCount !== 1 ? 's' : ''}`);
      
    } catch (err) {
      showToast(err.message || 'Failed to import from Excel', true);
    }
  }

  function createGridFromCells(cells, rowCount, colCount, sheet) {
    // Create a grid sizer to match Excel layout
    const grid = {
      id: createWidgetId(),
      type: 'gridsizer',
      name: 'ExcelGrid',
      properties: {
        rows: rowCount,
        cols: colCount,
        hgap: '8',
        vgap: '8'
      },
      children: []
    };
    
    // Create a 2D map of cells by position (0-indexed relative to range)
    const cellMap = new Map();
    let minRow = Infinity, minCol = Infinity;
    for (const cell of cells) {
      minRow = Math.min(minRow, cell.row);
      minCol = Math.min(minCol, cell.col);
    }
    for (const cell of cells) {
      const relRow = cell.row - minRow;
      const relCol = cell.col - minCol;
      cellMap.set(`${relRow},${relCol}`, cell);
    }
    
    // Fill grid in row-major order
    for (let r = 0; r < rowCount; r++) {
      for (let c = 0; c < colCount; c++) {
        const cell = cellMap.get(`${r},${c}`);
        const widget = cell ? cellToWidget(cell, sheet) : createSpacer();
        grid.children.push(widget);
      }
    }
    
    return grid;
  }

  function createSpacer() {
    return {
      id: createWidgetId(),
      type: 'spacer',
      name: 'spacer',
      properties: {}
    };
  }

  function cellToWidget(cell, sheet) {
    const { address, type, value, options, format } = cell;
    const baseName = `cell_${address.replace(/[^A-Za-z0-9]/g, '')}`;
    
    // Empty cells become spacers
    if (type === 'empty') {
      return createSpacer();
    }
    
    // Normalize value to string for defaults
    const defaultValue = value != null ? String(value) : '';
    
    const common = {
      id: createWidgetId(),
      name: baseName,
      properties: { defaultValue },
      excel: { enabled: true, sheet, cell: address, mode: 'output' }
    };
    
    switch (type) {
      case 'checkbox':
        // Checkbox - bidirectional
        return { 
          ...common, 
          type: 'checkbox',
          properties: { label: '', defaultValue: value === true || value === 'TRUE' },
          excel: { ...common.excel, mode: 'bidirectional' } 
        };
      
      case 'dropdown':
        // Dropdown - bidirectional
        const optList = options?.split(',').map(o => o.trim()).filter(Boolean) || [];
        return {
          ...common,
          type: 'dropdown',
          properties: { label: '', options: optList.join('\n'), defaultValue },
          excel: { ...common.excel, mode: 'bidirectional' }
        };
      
      case 'currency':
        // Currency - bidirectional input with money formatting
        return { 
          ...common, 
          type: 'currency',
          properties: { label: '', defaultValue, format: format || '$#,##0.00' },
          excel: { ...common.excel, mode: 'bidirectional' } 
        };
      
      case 'percentage':
        // Percentage - bidirectional with percent formatting
        return { 
          ...common, 
          type: 'percentage',
          properties: { label: '', defaultValue, format: format || '0%' },
          excel: { ...common.excel, mode: 'bidirectional' } 
        };
      
      case 'number':
        // Number without formula - bidirectional input (feeds equations)
        return { 
          ...common, 
          type: 'number',
          properties: { label: '', defaultValue },
          excel: { ...common.excel, mode: 'bidirectional' } 
        };
      
      case 'date':
        // Date - bidirectional
        return { 
          ...common, 
          type: 'datepicker',
          properties: { label: '', defaultValue },
          excel: { ...common.excel, mode: 'bidirectional' } 
        };
      
      case 'formula':
        // Formula result - read-only output from Excel
        return { 
          ...common, 
          type: 'output',
          properties: { label: '', defaultValue },
          excel: { ...common.excel, mode: 'output' }
        };
      
      case 'text':
      default:
        // Plain text - label bound to Excel for live updates
        return { 
          ...common, 
          type: 'label',
          name: baseName,
          properties: { label: defaultValue, defaultValue },
          excel: { ...common.excel, mode: 'output' }
        };
    }
  }

  function addImportedWidgets(widgets) {
    // If a container widget is selected, add inside it
    let parent = null;
    if (builderSelectedWidget && builderSelectedWidget !== 'app-root') {
      const selectedWidgetObj = findWidgetById(builderWidgets, builderSelectedWidget);
      if (selectedWidgetObj) {
        const def = WIDGET_TYPES[selectedWidgetObj.type];
        if (def?.isContainer) {
          parent = selectedWidgetObj;
        }
      }
    }
    
    if (parent) {
      // Add inside selected container
      if (!parent.children) parent.children = [];
      widgets.forEach(w => parent.children.push(w));
    } else {
      // Add to root
      widgets.forEach(w => builderWidgets.push(w));
    }
    
    renderWidgetTree();
    renderBuilderPreview();
  }

  function hasUnsavedBuilderChanges() {
    if (!builderSavedState) return builderWidgets.length > 0;
    return JSON.stringify(builderWidgets) !== builderSavedState;
  }

  async function previewAppUi(owner, name) {
    if (!token) return showToast('Login first', true);
    if (!owner || !name) return;
    try {
      const schema = await fetchAppSchema(owner, name);
      renderPreviewComponents(schema?.widgets || schema?.components || []);
      openPreviewModal(owner, name);
    } catch (err) {
      showToast(err.message || 'Unable to preview UI', true);
    }
  }

  function renderPreviewComponents(widgets) {
    if (!previewContent || !previewEmpty) return;
    if (!widgets || !widgets.length) {
      previewContent.innerHTML = '';
      previewEmpty.classList.remove('hidden');
      return;
    }
    previewEmpty.classList.add('hidden');
    previewContent.innerHTML = widgets.map(w => previewWidgetMarkup(w)).join('');
  }

  function previewWidgetMarkup(widget) {
    const def = WIDGET_TYPES[widget.type] || {};
    const props = widget.properties || {};
    const excel = widget.excel || {};
    const label = escapeHtml(props.label || widget.name || 'Widget');
    const location = excel.enabled ? `${escapeHtml(excel.sheet || '')}!${escapeHtml(excel.cell || '')}` : 'No binding';
    
    if (def.isContainer && widget.children?.length) {
      const childrenHtml = widget.children.map(c => previewWidgetMarkup(c)).join('');
      return `<div class="app-ui-container preview"><div class="container-label">${def.icon} ${label}</div><div class="container-children">${childrenHtml}</div></div>`;
    }
    
    return `<div class="app-ui-field preview"><label>${def.icon} ${label}</label><div class="app-meta">${def.label} • ${location}</div></div>`;
  }

  function openPreviewModal(owner, name) {
    if (!previewModal) return;
    if (previewSubtitle) previewSubtitle.textContent = `${owner}/${name}`;
    previewModal.classList.remove('hidden');
    previewModal.setAttribute('aria-hidden', 'false');
    bodyEl.classList.add('modal-open');
  }

  function closePreviewModal() {
    if (!previewModal) return;
    previewModal.classList.add('hidden');
    previewModal.setAttribute('aria-hidden', 'true');
    const otherModalOpen = (devModal && !devModal.classList.contains('hidden')) || (builderModal && !builderModal.classList.contains('hidden'));
    if (!otherModalOpen) bodyEl.classList.remove('modal-open');
    if (previewContent) previewContent.innerHTML = '';
    previewEmpty?.classList.remove('hidden');
    if (previewSubtitle) previewSubtitle.textContent = 'Select an app to preview its interface.';
  }

  const createForm = qs('#create-form');
  const nameInput = createForm.querySelector('input[name="name"]');
  const accessInput = createForm.querySelector('input[name="access_group"]');
  const descInput = createForm.querySelector('input[name="description"]');
  const publicInput = createForm.querySelector('input[name="public"]');
  const fileInput = createForm.querySelector('input[name="file"]');
  const imageInput = createForm.querySelector('input[name="image"]');

  if (devCreateNew) {
    devCreateNew.addEventListener('click', () => {
      setCreateMode('create');
      openDevModal();
    });
  }
  if (developerList) {
    developerList.addEventListener('click', (e) => {
      const editBtn = e.target.closest('[data-edit-app]');
      if (editBtn) {
        const target = editBtn.dataset.editApp;
        const app = appsCache.find(a => a.name === target && (role.admin || a.owner === currentUser));
        if (!app) return showToast('App not editable', true);
        setCreateMode('update', app);
        openDevModal();
        return;
      }
      const builderBtn = e.target.closest('[data-builder-app]');
      if (builderBtn) {
        const target = builderBtn.dataset.builderApp;
        const owner = builderBtn.dataset.owner || currentUser;
        openBuilder(owner, target);
        return;
      }
      const previewBtn = e.target.closest('[data-preview-app]');
      if (previewBtn) {
        const target = previewBtn.dataset.previewApp;
        const owner = previewBtn.dataset.owner || currentUser;
        previewAppUi(owner, target);
      }
    });
  }

  createForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!token) return showToast('Login first', true);
    if (!(role.developer || role.admin || developerOverride)) return showToast('Developer access required', true);
    const data = new FormData(createForm);
    const name = data.get('name').trim();
    const description = data.get('description') || '';
    const access_group = data.get('access_group').trim();
    const isPublic = data.get('public') === 'on';
    const file = data.get('file');
    const hasFile = file && file.size > 0;
    const isUpdate = formMode === 'update';
    if (!name || (!hasFile && !isUpdate)) return showToast('Name and file required', true);
    if (!isUpdate && !/^[a-zA-Z0-9_\-. ]+$/.test(name)) {
      return showToast('App name can only contain letters, numbers, spaces, underscores, hyphens, and periods', true);
    }
    if (isUpdate && !editingApp) return showToast('Select an app to edit', true);
    try {
      const imageFile = imageInput && imageInput.files ? imageInput.files[0] : null;
      const image_base64 = imageFile ? await resizeImageFile(imageFile) : null;
      if (!isUpdate) {
        const file_base64 = await fileToBase64(file);
        if (!file_base64) return showToast('Failed to read file', true);
        const payload = { name, description, file_base64, public: isPublic };
        if (access_group) payload.access_group = access_group;
        if (image_base64) payload.image_base64 = image_base64;
        console.log('Creating app with payload:', { name, description, hasFile: !!file_base64, fileLen: file_base64?.length, public: isPublic, access_group });
        const res = await apiFetch(`${apiBase}/apps`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', ...authHeaders() },
          body: JSON.stringify(payload)
        });
        const responseText = await res.clone().text();
        console.log('Server response:', res.status, responseText);
        if (!res.ok) throw new Error(responseText || 'Create failed');
        showToast('Application created');
      } else if (hasFile) {
        const payload = { name, description, public: isPublic };
        if (access_group) payload.access_group = access_group;
        payload.file_base64 = await fileToBase64(file);
        if (image_base64) payload.image_base64 = image_base64;
        const res = await apiFetch(`${apiBase}/apps/version`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', ...authHeaders() },
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error('Version publish failed');
        showToast('New version published');
      } else {
        const payload = { description, public: isPublic };
        if (access_group) payload.access_group = access_group;
        if (image_base64) payload.image_base64 = image_base64;
        const res = await apiFetch(`${apiBase}/apps/${encodeURIComponent(editingApp.name)}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json', ...authHeaders() },
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error('Update failed');
        showToast('Application updated');
      }
      setCreateMode('create');
      closeDevModal();
      await refreshApps();
    } catch (err) {
      showToast(err.message || 'Create failed', true);
    }
  });

  const userForm = qs('#user-form');
  userForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!role.admin) return showToast('Admin required', true);
    const data = new FormData(userForm);
    const payload = {
      username: data.get('username').trim(),
      password: data.get('password'),
      groups: data.get('groups').trim(),
      role: data.get('role')
    };
    if (!payload.username || !payload.password) return showToast('Username and password required', true);
    try {
      const res = await apiFetch(`${apiBase}/users`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', ...authHeaders() },
        body: JSON.stringify(payload)
      });
      if (!res.ok) throw new Error('User save failed');
      showToast('User saved');
      await loadUsers();
    } catch (err) {
      showToast(err.message || 'User save failed', true);
    }
  });

  async function loadUsers() {
    if (!role.admin) return;
    try {
      const res = await apiFetch(`${apiBase}/users`, { headers: { ...authHeaders() } });
      if (!res.ok) throw new Error('Cannot load users');
      const users = await res.json();
      renderUsers(users);
      adminWarning.classList.add('hidden');
      usersArea.classList.remove('hidden');
    } catch (err) {
      adminWarning.classList.remove('hidden');
      usersArea.classList.add('hidden');
      showToast(err.message || 'Cannot load users', true);
    }
  }

  function renderUsers(users) {
    const tbody = qs('#users-table tbody');
    tbody.innerHTML = users.map(u => `<tr><td>${escapeHtml(u.name)}</td><td>${escapeHtml((u.groups || []).join(', '))}</td><td>${escapeHtml(u.role)}</td></tr>`).join('');
  }

  const escapeHtml = (s) => (s || '').replace(/[&<>"]+/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
  const formatDescription = (desc) => {
    const normalized = (desc || '').trim();
    if (!normalized) return 'No description yet.';
    if (normalized.length <= 120) return normalized;
    return `${normalized.slice(0, 117).trimEnd()}...`;
  };

  const fileToBase64 = (file) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const res = reader.result;
      const base64 = res.substring(res.indexOf(',') + 1);
      resolve(base64);
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });

  const resizeImageFile = (file, maxSize = 512) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const maxDim = Math.max(img.width, img.height);
        const scale = maxDim > maxSize ? maxSize / maxDim : 1;
        const canvas = document.createElement('canvas');
        canvas.width = Math.max(1, Math.round(img.width * scale));
        canvas.height = Math.max(1, Math.round(img.height * scale));
        const ctx = canvas.getContext('2d');
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        const base64 = canvas.toDataURL('image/png', 0.85).split(',')[1];
        resolve(base64);
      };
      img.onerror = () => reject(new Error('Image decode failed'));
      img.src = reader.result;
    };
    reader.onerror = () => reject(reader.error || new Error('Image read failed'));
    reader.readAsDataURL(file);
  });

  function clearSession() {
    sessionStorage.removeItem('esa_token');
    sessionStorage.removeItem('esa_user');
    token = null;
    currentUser = null;
    role = { admin: false, developer: false };
    developerOverride = false;
  }

  // Initialize
  toggleCreateVisibility();
  (async () => {
    await detectRole();
    await refreshApps();
    if (role.admin) await loadUsers();
  })();

  function setTab(tab) {
    activeTab = tab;
    Object.entries(tabButtons).forEach(([key, btn]) => {
      btn.classList.toggle('active', key === tab);
    });
    Object.entries(tabContents).forEach(([key, el]) => {
      if (!el) return;
      const show = key === tab;
      if (key === 'developer') {
        const canDev = developerOverride || role.developer || role.admin;
        el.classList.toggle('hidden', !show || !canDev || !token);
        if (show && canDev) renderDeveloperApps();
      } else if (key === 'admin') {
        el.classList.toggle('hidden', !show || !role.admin || !token);
      } else {
        el.classList.toggle('hidden', !show || !token);
      }
    });
  }

  function renderDeveloperApps() {
    if (!developerList) return;
    const canEdit = developerOverride || role.developer || role.admin;
    if (!token || !canEdit || activeTab !== 'developer') return;
    const mine = appsCache.filter(a => role.admin || a.owner === currentUser);
    if (mine.length === 0) {
      developerList.innerHTML = '<div class="hint">No editable applications yet.</div>';
      return;
    }
    developerList.innerHTML = mine.map(a => {
      const imageHtml = a.image_base64
        ? `<img src="data:image/png;base64,${a.image_base64}" alt="${escapeHtml(a.name)}" class="app-card-image" />`
        : `<div class="app-card-image" style="display: grid; place-items: center; font-size: 48px; font-weight: 700; color: var(--accent)">${escapeHtml(a.name.charAt(0).toUpperCase())}</div>`;
      const description = formatDescription(a.description);
      return `<div class="app-card dev-app-card" data-builder-app="${escapeHtml(a.name)}" data-owner="${escapeHtml(a.owner)}">
        ${imageHtml}
        <div class="app-card-body">
          <h4>${escapeHtml(a.name)}</h4>
          <p>${escapeHtml(description)}</p>
        </div>
      </div>`;
    }).join('');
  }

  function setCreateMode(mode, app = null) {
    formMode = mode;
    editingApp = app;
    if (mode === 'create') {
      devFormMode.textContent = 'Mode: Create new';
      nameInput.readOnly = false;
      createForm.reset();
      publicInput.checked = false;
      fileInput.required = true;
      if (imageInput) imageInput.value = '';
    } else if (app) {
      devFormMode.textContent = `Mode: Update ${app.name} (v${app.latest_version})`;
      nameInput.value = app.name;
      nameInput.readOnly = true;
      accessInput.value = app.access_group || '';
      descInput.value = app.description || '';
      publicInput.checked = !!app.public;
      fileInput.value = '';
      fileInput.required = false;
      if (imageInput) imageInput.value = '';
    }
  }

  function openDevModal() {
    if (!devModal) return;
    devModal.classList.remove('hidden');
    devModal.setAttribute('aria-hidden', 'false');
    bodyEl.classList.add('modal-open');
  }

  function closeDevModal(resetForm = false) {
    if (!devModal) return;
    devModal.classList.add('hidden');
    devModal.setAttribute('aria-hidden', 'true');
    bodyEl.classList.remove('modal-open');
    if (resetForm) setCreateMode('create');
  }

  function openBuilderModal() {
    if (!builderModal) return;
    builderModal.classList.remove('hidden');
    builderModal.setAttribute('aria-hidden', 'false');
    bodyEl.classList.add('modal-open');
  }

  function closeBuilderModal(force = false) {
    if (!builderModal) return;
    if (!force && hasUnsavedBuilderChanges()) {
      if (!confirm('You have unsaved changes. Are you sure you want to close the App Builder?')) {
        return;
      }
    }
    builderModal.classList.add('hidden');
    builderModal.setAttribute('aria-hidden', 'true');
    bodyEl.classList.remove('modal-open');
    builderTarget = null;
    builderWidgets = [];
    builderSelectedWidget = null;
    builderAppCollapsed = false;
    builderSavedState = null;
    hidePropertyForm();
    if (builderTreeRoot) builderTreeRoot.innerHTML = '';
    if (builderAppLabel) builderAppLabel.textContent = 'No app';
  }

  function normalizeComponents(list) {
    if (!Array.isArray(list)) return [];
    return list.map(item => ({ ...item, id: item.id || createWidgetId() }));
  }

  function renderAppIcon(app) {
    if (app?.image_base64) {
      return `<img class="app-icon" src="data:image/png;base64,${app.image_base64}" alt="${escapeHtml(app.name)} icon">`;
    }
    const letter = escapeHtml((app?.name || '?').charAt(0).toUpperCase());
    return `<div class="app-icon placeholder">${letter}</div>`;
  }
})();
