(() => {
  const apiBase = 'http://localhost:8080'; // hardcoded ESA API base
  let token = sessionStorage.getItem('esa_token');
  let currentUser = sessionStorage.getItem('esa_user');
  let role = { admin: false, developer: false };
  let developerOverride = false;
  let appsCache = [];
  let formMode = 'create';
  let editingApp = null;

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
  const adminPanel = qs('#admin-panel');
  const adminWarning = qs('#admin-warning');
  const usersArea = qs('#users-area');
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
  });

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
      return `<div class="app-card"><h4>${escapeHtml(a.name)}</h4><div class="app-meta">Owner: ${escapeHtml(a.owner)} • Version ${a.latest_version}</div><div class="app-meta">${a.public ? 'Public' : 'Restricted'}</div><p>${escapeHtml(a.description || '')}</p></div>`;
    }).join('');
  }

  const createForm = qs('#create-form');
  const nameInput = createForm.querySelector('input[name="name"]');
  const accessInput = createForm.querySelector('input[name="access_group"]');
  const descInput = createForm.querySelector('input[name="description"]');
  const publicInput = createForm.querySelector('input[name="public"]');
  const fileInput = createForm.querySelector('input[name="file"]');

  if (devCreateNew) {
    devCreateNew.addEventListener('click', () => {
      setCreateMode('create');
      openDevModal();
    });
  }
  if (developerList) {
    developerList.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-edit-app]');
      if (!btn) return;
      const target = btn.dataset.editApp;
      const app = appsCache.find(a => a.name === target && (role.admin || a.owner === currentUser));
      if (!app) return showToast('App not editable', true);
      setCreateMode('update', app);
      openDevModal();
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
      if (!isUpdate) {
        const file_base64 = await fileToBase64(file);
        const payload = { name, description, file_base64, public: isPublic };
        if (access_group) payload.access_group = access_group;
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
      return `<div class="app-card"><div class="app-card-header"><div><h4>${escapeHtml(a.name)}</h4><div class="app-meta">${meta}</div></div><button type="button" data-edit-app="${escapeHtml(a.name)}">Edit</button></div><p>${escapeHtml(a.description || '')}</p></div>`;
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
    } else if (app) {
      devFormMode.textContent = `Mode: Update ${app.name} (v${app.latest_version})`;
      nameInput.value = app.name;
      nameInput.readOnly = true;
      accessInput.value = app.access_group || '';
      descInput.value = app.description || '';
      publicInput.checked = !!app.public;
      fileInput.value = '';
      fileInput.required = false;
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
})();
