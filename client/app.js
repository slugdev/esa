(() => {
  const urlApi = new URLSearchParams(location.search).get('api');
  const apiBase = urlApi || (location.origin === 'null' ? 'http://localhost:8080' : window.location.origin); // adjust if server runs elsewhere
  let token = sessionStorage.getItem('esa_token');
  let currentUser = sessionStorage.getItem('esa_user');
  let role = { admin: false, developer: false };
  let developerOverride = false;

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
  const adminPanel = qs('#admin-panel');
  const adminWarning = qs('#admin-warning');
  const usersArea = qs('#users-area');

  if (!token || !currentUser) {
    window.location.replace('login.html');
    return;
  }

  sessionInfo.classList.remove('hidden');
  sessionUser.textContent = `Signed in as ${currentUser}`;

  logoutBtn.addEventListener('click', async () => {
    try { await fetch(`${apiBase}/logout`, { method: 'POST', headers: authHeaders() }); } catch (_) {}
    clearSession();
    window.location.replace('login.html');
  });

  backToLogin.addEventListener('click', () => {
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
    createPanel.classList.toggle('hidden', !canDev || !signedIn);
    createHint.classList.toggle('hidden', canDev && signedIn);
    developerPill.classList.toggle('hidden', canDev && signedIn);
    qs('#apps-panel').classList.toggle('hidden', !signedIn);
    adminPanel.classList.toggle('hidden', !signedIn);
  }

  qs('#developer-override').addEventListener('click', () => {
    if (!token) { showToast('Login first', true); return; }
    developerOverride = true;
    toggleCreateVisibility();
    showToast('Developer UI forced for this session');
  });

  qs('#refresh-apps').addEventListener('click', refreshApps);

  async function refreshApps() {
    if (!token) { showToast('Login first', true); return; }
    try {
      const res = await fetch(`${apiBase}/apps`, { headers: { ...authHeaders() } });
      if (!res.ok) throw new Error('Failed to load apps');
      const apps = await res.json();
      renderApps(apps);
    } catch (err) {
      renderApps([]);
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
    if (!name || !file) return showToast('Name and file required', true);
    try {
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
      createForm.reset();
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
})();
