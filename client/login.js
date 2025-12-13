(() => {
  const urlApi = new URLSearchParams(location.search).get('api');
  const apiBase = urlApi || (location.origin === 'null' ? 'http://localhost:8080' : window.location.origin);
  const toastEl = document.querySelector('#toast');
  const showToast = (msg, isError = false) => {
    toastEl.textContent = msg;
    toastEl.classList.toggle('error', isError);
    toastEl.classList.remove('hidden');
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => toastEl.classList.add('hidden'), 2800);
  };

  const loginForm = document.querySelector('#login-form');

  loginForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const data = new FormData(loginForm);
    const username = data.get('username').trim();
    const password = data.get('password');
    if (!username || !password) return;
    try {
      const res = await fetch(`${apiBase}/login`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password })
      });
      if (!res.ok) throw new Error('Login failed');
      const json = await res.json();
      sessionStorage.setItem('esa_token', json.token);
      sessionStorage.setItem('esa_user', username);
      showToast('Logged in');
      setTimeout(() => window.location.href = 'index.html', 400);
    } catch (err) {
      showToast(err.message || 'Login failed', true);
    }
  });

})();
