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
  let builderComponents = [];
  let builderSheets = [];
  const qs = (sel) => document.querySelector(sel);
  const toastEl = qs('#toast');

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
  const builderCancel = qs('#ui-builder-cancel');
  const builderComponentsEl = qs('#builder-components');
  const builderEmpty = qs('#builder-empty');
  const builderForm = qs('#builder-form');
  const builderLabelInput = qs('#builder-label');
  const builderSheetSelect = qs('#builder-sheet');
  const builderCellInput = qs('#builder-cell');
  const builderModeSelect = qs('#builder-mode');
  const builderInputTypeSelect = qs('#builder-input-type');
  const builderComponentIdInput = qs('#builder-component-id');
  const builderResetBtn = qs('#builder-reset');
  const builderAppLabel = qs('#builder-app-label');
  const builderSaveLayoutBtn = qs('#builder-save-layout');
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

  qs('#developer-override').addEventListener('click', () => {
    if (!token) { showToast('Login first', true); return; }
    developerOverride = true;
    toggleCreateVisibility();
    showToast('Developer UI forced for this session');
  });

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
  if (builderCancel) builderCancel.addEventListener('click', () => closeBuilderModal());
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
  builderResetBtn?.addEventListener('click', () => setBuilderEditing(null));
  builderForm?.addEventListener('submit', handleBuilderFormSubmit);
  builderComponentsEl?.addEventListener('click', handleBuilderListClick);
  builderSaveLayoutBtn?.addEventListener('click', () => saveBuilderLayout());
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
    if (!res.ok) throw new Error(data?.error || 'Unable to read cell');
    return data?.value ?? '';
  }

  function updateComponentDisplay(id, value) {
    const el = appUiForm?.querySelector(`[data-component="${id}"]`);
    if (!el) return;
    if (el.tagName === 'INPUT') {
      el.value = value ?? '';
    } else {
      el.textContent = formatValue(value);
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
    const components = normalizeComponents(schema?.components);
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
    appUiForm.innerHTML = components.map(comp => {
      const mode = comp.mode || 'display';
      const inputType = comp.inputType || 'text';
      if (mode === 'input') {
        return `<div class="app-ui-field"><label>${escapeHtml(comp.label || comp.cell || 'Input')}</label><input data-component="${comp.id}" type="${escapeHtml(inputType)}" placeholder="${escapeHtml(comp.sheet || '')}!${escapeHtml(comp.cell || '')}" /></div>`;
      }
      return `<div class="app-ui-field"><label>${escapeHtml(comp.label || comp.cell || 'Value')}</label><output data-component="${comp.id}">—</output><div class="app-meta">${escapeHtml(comp.sheet || '')}!${escapeHtml(comp.cell || '')}</div></div>`;
    }).join('');
  }

  function handleAppFormChange(e) {
    const input = e.target.closest('input[data-component]');
    if (!input) return;
    const id = input.dataset.component;
    const component = activeComponents.find(c => c.id === id);
    if (!component || (component.mode || 'display') !== 'input') return;
    commitComponentValue(component, input.value);
  }

  async function commitComponentValue(component, rawValue) {
    if (!component?.sheet || !component?.cell) return;
    const payload = { sheet: component.sheet, range: component.cell };
    if ((component.inputType || 'text') === 'number') {
      const trimmed = (rawValue ?? '').trim();
      if (trimmed === '') return showToast('Enter a number', true);
      const num = Number(trimmed);
      if (!Number.isFinite(num)) return showToast('Enter a valid number', true);
      payload.value_number = num;
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
      builderComponents = normalizeComponents(schema?.components);
      setSheetOptions(sheets);
      if (builderAppLabel) builderAppLabel.textContent = `${owner}/${name}`;
      setBuilderEditing(null);
      renderBuilderComponents();
      openBuilderModal();
    } catch (err) {
      showToast(err.message || 'Unable to load builder', true);
    }
  }

  function renderBuilderComponents() {
    if (!builderComponentsEl) return;
    if (!builderComponents.length) {
      builderComponentsEl.innerHTML = '';
      builderEmpty?.classList.remove('hidden');
      return;
    }
    builderEmpty?.classList.add('hidden');
    builderComponentsEl.innerHTML = builderComponents.map(comp => {
      const meta = `${escapeHtml(comp.sheet || '')}!${escapeHtml(comp.cell || '')} • ${(comp.mode || 'display') === 'input' ? 'Input' : 'Display'}`;
      return `<div class="builder-row"><div><h5>${escapeHtml(comp.label || comp.cell || 'Component')}</h5><div class="app-meta">${meta}</div></div><div class="actions"><button type="button" class="ghost" data-edit-component="${comp.id}">Edit</button><button type="button" class="ghost" data-remove-component="${comp.id}">Remove</button></div></div>`;
    }).join('');
  }

  function showSheetLoadingState() {
    if (!builderSheetSelect) return;
    builderSheetSelect.disabled = true;
    builderSheetSelect.innerHTML = '<option value="">Loading sheets…</option>';
    builderSheetSelect.value = '';
  }

  function setSheetOptions(sheets, preferred) {
    builderSheets = Array.isArray(sheets) ? sheets.filter(Boolean) : [];
    renderSheetOptions(preferred);
  }

  function ensureSheetOption(sheet) {
    if (!sheet) return;
    if (!builderSheets.includes(sheet)) {
      builderSheets.push(sheet);
    }
    renderSheetOptions(sheet);
  }

  function renderSheetOptions(preferred) {
    if (!builderSheetSelect) return;
    const targetValue = preferred ?? builderSheetSelect.value;
    if (!builderSheets.length) {
      builderSheetSelect.disabled = true;
      builderSheetSelect.innerHTML = '<option value="">No sheets found</option>';
      builderSheetSelect.value = '';
      return;
    }
    builderSheetSelect.disabled = false;
    builderSheetSelect.innerHTML = builderSheets
      .map(sheet => `<option value="${escapeHtml(sheet)}">${escapeHtml(sheet)}</option>`)
      .join('');
    if (targetValue && builderSheets.includes(targetValue)) {
      builderSheetSelect.value = targetValue;
    } else {
      builderSheetSelect.value = builderSheets[0];
    }
  }

  function setBuilderEditing(id) {
    if (!builderForm) return;
    if (!id) {
      builderComponentIdInput.value = '';
      builderForm.reset();
      builderModeSelect.value = 'display';
      builderInputTypeSelect.value = 'text';
      if (builderSheetSelect) builderSheetSelect.value = builderSheets[0] || '';
      return;
    }
    const comp = builderComponents.find(c => c.id === id);
    if (!comp) return setBuilderEditing(null);
    builderComponentIdInput.value = comp.id;
    builderLabelInput.value = comp.label || '';
    ensureSheetOption(comp.sheet || '');
    if (builderSheetSelect) builderSheetSelect.value = comp.sheet || builderSheets[0] || '';
    builderCellInput.value = comp.cell || '';
    builderModeSelect.value = comp.mode || 'display';
    builderInputTypeSelect.value = comp.inputType || 'text';
  }

  function handleBuilderFormSubmit(e) {
    e.preventDefault();
    if (!builderTarget) return showToast('Select an app to design', true);
    const label = builderLabelInput.value.trim() || 'Untitled field';
    const sheet = (builderSheetSelect?.value || '').trim();
    if (!sheet) return showToast('Select a sheet', true);
    const cell = builderCellInput.value.trim();
    if (!cell) return showToast('Cell / range is required', true);
    const mode = builderModeSelect.value || 'display';
    const inputType = builderInputTypeSelect.value || 'text';
    const id = builderComponentIdInput.value || createComponentId();
    const next = { id, label, sheet, cell, mode, inputType };
    const idx = builderComponents.findIndex(c => c.id === id);
    if (idx >= 0) builderComponents[idx] = next; else builderComponents.push(next);
    setBuilderEditing(null);
    renderBuilderComponents();
  }

  function handleBuilderListClick(e) {
    const editBtn = e.target.closest('[data-edit-component]');
    if (editBtn) {
      setBuilderEditing(editBtn.dataset.editComponent);
      return;
    }
    const removeBtn = e.target.closest('[data-remove-component]');
    if (removeBtn) {
      builderComponents = builderComponents.filter(c => c.id !== removeBtn.dataset.removeComponent);
      setBuilderEditing(null);
      renderBuilderComponents();
    }
  }

  async function saveBuilderLayout() {
    if (!builderTarget) return showToast('Select an app to design', true);
    const schema = { components: builderComponents };
    try {
      const res = await apiFetch(`${apiBase}/apps/ui/save`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', ...authHeaders() },
        body: JSON.stringify({ owner: builderTarget.owner, name: builderTarget.name, schema_json: JSON.stringify(schema) })
      });
      if (!res.ok) throw new Error('Failed to save layout');
      showToast('Layout saved');
      closeBuilderModal();
      await refreshApps();
    } catch (err) {
      showToast(err.message || 'Failed to save layout', true);
    }
  }

  async function previewAppUi(owner, name) {
    if (!token) return showToast('Login first', true);
    if (!owner || !name) return;
    try {
      const schema = await fetchAppSchema(owner, name);
      renderPreviewComponents(schema?.components || []);
      openPreviewModal(owner, name);
    } catch (err) {
      showToast(err.message || 'Unable to preview UI', true);
    }
  }

  function renderPreviewComponents(components) {
    if (!previewContent || !previewEmpty) return;
    const normalized = normalizeComponents(components || []);
    if (!normalized.length) {
      previewContent.innerHTML = '';
      previewEmpty.classList.remove('hidden');
      return;
    }
    previewEmpty.classList.add('hidden');
    previewContent.innerHTML = normalized.map(previewComponentMarkup).join('');
  }

  function previewComponentMarkup(comp) {
    const mode = comp.mode || 'display';
    const inputType = comp.inputType || 'text';
    const label = escapeHtml(comp.label || comp.cell || 'Field');
    const location = `${escapeHtml(comp.sheet || '')}!${escapeHtml(comp.cell || '')}`;
    if (mode === 'input') {
      return `<div class="app-ui-field preview"><label>${label}</label><input type="${escapeHtml(inputType)}" disabled placeholder="${location}" /><div class="app-meta">Two-way • ${location}</div></div>`;
    }
    return `<div class="app-ui-field preview"><label>${label}</label><output>Preview value</output><div class="app-meta">Display • ${location}</div></div>`;
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
    const isUpdate = formMode === 'update';
    if (!name || (!file && !isUpdate)) return showToast('Name and file required', true);
    if (isUpdate && !editingApp) return showToast('Select an app to edit', true);
    try {
      const imageFile = imageInput && imageInput.files ? imageInput.files[0] : null;
      const image_base64 = imageFile ? await resizeImageFile(imageFile) : null;
      if (!isUpdate) {
        const file_base64 = await fileToBase64(file);
        const payload = { name, description, file_base64, public: isPublic };
        if (access_group) payload.access_group = access_group;
        if (image_base64) payload.image_base64 = image_base64;
        const res = await apiFetch(`${apiBase}/apps`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', ...authHeaders() },
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error('Create failed');
        showToast('Application created');
      } else if (file) {
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
      return `<div class="app-card dev-app-card">
        ${imageHtml}
        <div class="app-card-body">
          <h4>${escapeHtml(a.name)}</h4>
          <p>${escapeHtml(description)}</p>
          <div class="app-buttons">
            <button type="button" data-edit-app="${escapeHtml(a.name)}">Edit</button>
            <button type="button" class="ghost" data-builder-app="${escapeHtml(a.name)}" data-owner="${escapeHtml(a.owner)}">UI Builder</button>
          </div>
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

  function closeBuilderModal() {
    if (!builderModal) return;
    builderModal.classList.add('hidden');
    builderModal.setAttribute('aria-hidden', 'true');
    bodyEl.classList.remove('modal-open');
    builderTarget = null;
    builderComponents = [];
    setBuilderEditing(null);
    renderBuilderComponents();
    if (builderAppLabel) builderAppLabel.textContent = 'No app';
  }

  function normalizeComponents(list) {
    if (!Array.isArray(list)) return [];
    return list.map(item => ({ ...item, id: item.id || createComponentId() }));
  }

  function createComponentId() {
    return 'cmp-' + Math.random().toString(36).slice(2, 9);
  }

  function renderAppIcon(app) {
    if (app?.image_base64) {
      return `<img class="app-icon" src="data:image/png;base64,${app.image_base64}" alt="${escapeHtml(app.name)} icon">`;
    }
    const letter = escapeHtml((app?.name || '?').charAt(0).toUpperCase());
    return `<div class="app-icon placeholder">${letter}</div>`;
  }
})();
