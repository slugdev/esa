(() => {
  const apiBase = 'http://localhost:8080'; // hardcoded ESA API base
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

  const qs = (sel) => document.querySelector(sel);
  const qsa = (sel) => Array.from(document.querySelectorAll(sel));
  const toastEl = qs('#toast');

  const showToast = (msg, isError = false) => {
    toastEl.textContent = msg;
    toastEl.classList.toggle('error', isError);
    toastEl.classList.remove('hidden');
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => toastEl.classList.add('hidden'), 2800);
  };

  const authHeaders = () => token ? { 'Authorization': `Bearer ${token}` } : {};

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
  const builderSheetInput = qs('#builder-sheet');
  const builderCellInput = qs('#builder-cell');
  const builderModeSelect = qs('#builder-mode');
  const builderInputTypeSelect = qs('#builder-input-type');
  const builderComponentIdInput = qs('#builder-component-id');
  const builderResetBtn = qs('#builder-reset');
  const builderAppLabel = qs('#builder-app-label');
  const builderSaveLayoutBtn = qs('#builder-save-layout');
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
  sessionUser.textContent = `Signed in as ${currentUser}`;

  Object.entries(tabButtons).forEach(([key, btn]) => {
    btn.addEventListener('click', () => setTab(key));
  });

  logoutBtn.addEventListener('click', async () => {
    try { await fetch(`${apiBase}/logout`, { method: 'POST', headers: authHeaders() }); } catch (_) {}
    clearSession();
    window.location.replace('login.html');
  });

  if (token && backToLogin) backToLogin.remove();
  backToLogin?.addEventListener('click', () => {
    clearSession();
    window.location.href = 'login.html';
  });

  async function detectRole() {
    role = { admin: false, developer: false };
    if (!token || !currentUser) { toggleCreateVisibility(); return; }
    try {
      const res = await fetch(`${apiBase}/users`, { headers: { ...authHeaders() } });
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
    sessionRoles.textContent = `Roles: ${role.admin ? 'Admin' : 'User'}${developerOverride || role.developer ? ' • Developer' : ''}`;
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
    if (e.key === 'Escape' && devModal && !devModal.classList.contains('hidden')) {
      closeDevModal(true);
    }
    if (e.key === 'Escape' && builderModal && !builderModal.classList.contains('hidden')) {
      closeBuilderModal();
    }
  });

  appsListEl?.addEventListener('click', handleAppsListClick);
  appRefreshBtn?.addEventListener('click', () => refreshActiveApp(true));
  appExitBtn?.addEventListener('click', () => exitActiveApp());

  if (builderClose) builderClose.addEventListener('click', () => closeBuilderModal());
  if (builderCancel) builderCancel.addEventListener('click', () => closeBuilderModal());
  if (builderModal) {
    builderModal.addEventListener('click', (e) => {
      if (e.target === builderModal) closeBuilderModal();
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
      const res = await fetch(`${apiBase}/apps`, { headers: { ...authHeaders() } });
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
      const icon = renderAppIcon(a);
      const canAuthor = (developerOverride || role.developer || role.admin) && (role.admin || a.owner === currentUser);
      const launchLabel = a.has_ui ? 'Launch' : 'Launch (No UI)';
      const launchBtn = `<button type="button" data-launch-app="${escapeHtml(a.name)}" data-owner="${escapeHtml(a.owner)}">${launchLabel}</button>`;
      const designBtn = canAuthor ? `<button type="button" class="ghost" data-design-app="${escapeHtml(a.name)}" data-design-owner="${escapeHtml(a.owner)}">Design UI</button>` : '';
      const statusMeta = a.has_ui ? 'UI ready' : 'UI not authored yet';
      return `<div class="app-card"><div class="app-card-top">${icon}<div><h4>${escapeHtml(a.name)}</h4><div class="app-meta">Owner: ${escapeHtml(a.owner)} • Version ${a.latest_version}</div></div></div><div class="app-meta">${a.public ? 'Public' : 'Restricted'} • ${statusMeta}</div><p>${escapeHtml(a.description || '')}</p><div class="app-buttons">${launchBtn}${designBtn}</div></div>`;
    }).join('');
  }

  function handleAppsListClick(e) {
    const launchBtn = e.target.closest('[data-launch-app]');
    if (launchBtn) {
      const owner = launchBtn.dataset.owner;
      const name = launchBtn.dataset.launchApp;
      launchApp(owner, name);
      return;
    }
    const designBtn = e.target.closest('[data-design-app]');
    if (designBtn) {
      const owner = designBtn.dataset.designOwner || currentUser;
      const name = designBtn.dataset.designApp;
      openBuilder(owner, name);
    }
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
      updateWorkspaceState({ title: `Launching ${name}…`, desc: 'Loading workbook and UI schema.', busy: true });
      const schema = await fetchAppSchema(owner, name);
      await ensureWorkbookLoaded(owner, name);
      activeApp = { owner, name, schema };
      renderAppUi(schema);
      await refreshActiveApp();
      updateWorkspaceState({ title: `${owner}/${name}`, desc: 'Workbook synced. Edit inputs to push values back to Excel.', busy: false });
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
    const res = await fetch(`${apiBase}/excel/load`, {
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
      await fetch(`${apiBase}/excel/close`, { method: 'POST', headers: { ...authHeaders() } });
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
  }

  async function refreshActiveApp(showToastOnError = false) {
    if (!activeApp || !activeComponents.length) return;
    try {
      for (const component of activeComponents) {
        const value = await queryComponentValue(component);
        updateComponentDisplay(component.id, value);
      }
      if (showToastOnError) showToast('Values refreshed');
    } catch (err) {
      if (showToastOnError) showToast(err.message || 'Refresh failed', true);
    }
  }

  async function queryComponentValue(component) {
    if (!component?.sheet || !component?.cell) return '';
    const res = await fetch(`${apiBase}/excel/query`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...authHeaders() },
      body: JSON.stringify({ sheet: component.sheet, range: component.cell })
    });
    if (!res.ok) throw new Error('Unable to read cell');
    const data = await res.json();
    return data.value;
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
      const res = await fetch(`${apiBase}/excel/set`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', ...authHeaders() },
        body: JSON.stringify(payload)
      });
      if (!res.ok) throw new Error('Failed to push value');
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

  async function fetchAppSchema(owner, name) {
    const res = await fetch(`${apiBase}/apps/ui/get`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...authHeaders() },
      body: JSON.stringify({ owner, name })
    });
    if (!res.ok) throw new Error('Unable to load UI schema');
    const data = await res.json();
    return data.schema || { components: [] };
  }

  async function openBuilder(owner, name) {
    if (!token) return showToast('Login first', true);
    if (!canAuthorApp(owner)) return showToast('Developer access required', true);
    closeDevModal();
    try {
      const schema = await fetchAppSchema(owner, name);
      builderTarget = { owner, name };
      builderComponents = normalizeComponents(schema?.components);
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

  function setBuilderEditing(id) {
    if (!builderForm) return;
    if (!id) {
      builderComponentIdInput.value = '';
      builderForm.reset();
      builderModeSelect.value = 'display';
      builderInputTypeSelect.value = 'text';
      if (builderSheetInput) builderSheetInput.value = 'Sheet1';
      return;
    }
    const comp = builderComponents.find(c => c.id === id);
    if (!comp) return setBuilderEditing(null);
    builderComponentIdInput.value = comp.id;
    builderLabelInput.value = comp.label || '';
    builderSheetInput.value = comp.sheet || '';
    builderCellInput.value = comp.cell || '';
    builderModeSelect.value = comp.mode || 'display';
    builderInputTypeSelect.value = comp.inputType || 'text';
  }

  function handleBuilderFormSubmit(e) {
    e.preventDefault();
    if (!builderTarget) return showToast('Select an app to design', true);
    const label = builderLabelInput.value.trim() || 'Untitled field';
    const sheet = builderSheetInput.value.trim() || 'Sheet1';
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
      const res = await fetch(`${apiBase}/apps/ui/save`, {
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
        const res = await fetch(`${apiBase}/apps`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', ...authHeaders() },
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error('Create failed');
        showToast('Application created');
      } else {
        const payload = { description, public: isPublic };
        if (access_group) payload.access_group = access_group;
        if (file) {
          payload.file_base64 = await fileToBase64(file);
          payload.new_version = true;
        } else {
          payload.new_version = false;
        }
        if (image_base64) payload.image_base64 = image_base64;
        const res = await fetch(`${apiBase}/apps/${encodeURIComponent(editingApp.name)}`, {
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
      const res = await fetch(`${apiBase}/users`, {
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
      const res = await fetch(`${apiBase}/users`, { headers: { ...authHeaders() } });
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
      const meta = `${escapeHtml(a.owner)} • Version ${a.latest_version} • ${a.public ? 'Public' : 'Restricted'}`;
      const icon = renderAppIcon(a);
      return `<div class="app-card"><div class="app-card-top">${icon}<div><h4>${escapeHtml(a.name)}</h4><div class="app-meta">${meta}</div></div></div><div class="app-card-header"><div class="app-buttons"><button type="button" data-edit-app="${escapeHtml(a.name)}">Edit</button><button type="button" class="ghost" data-builder-app="${escapeHtml(a.name)}" data-owner="${escapeHtml(a.owner)}">Design UI</button></div></div><p>${escapeHtml(a.description || '')}</p></div>`;
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
    bodyEl.classList.add('modal-open');
  }

  function closeDevModal(resetForm = false) {
    if (!devModal) return;
    devModal.classList.add('hidden');
    bodyEl.classList.remove('modal-open');
    if (resetForm) setCreateMode('create');
  }

  function openBuilderModal() {
    if (!builderModal) return;
    builderModal.classList.remove('hidden');
    bodyEl.classList.add('modal-open');
  }

  function closeBuilderModal() {
    if (!builderModal) return;
    builderModal.classList.add('hidden');
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
