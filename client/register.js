(() => {
  const apiBase = window.location.origin;
  const toastEl = document.querySelector('#toast');
  const showToast = (msg, isError = false) => {
    toastEl.textContent = msg;
    toastEl.classList.toggle('error', isError);
    toastEl.classList.remove('hidden');
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => toastEl.classList.add('hidden'), 2800);
  };

  const signupForm = document.querySelector('#signup-form');

  signupForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const data = new FormData(signupForm);
    const payload = {
      username: data.get('username').trim(),
      password: data.get('password'),
      groups: data.get('groups').trim(),
      role: 'user'
    };
    if (!payload.username || !payload.password) return showToast('Username and password required', true);
    try {
      const res = await fetch(`${apiBase}/users`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (!res.ok) {
        const text = await res.text();
        throw new Error(text || 'Account creation failed');
      }
      showToast('Account created. Return to login.');
      signupForm.reset();
    } catch (err) {
      showToast(err.message || 'Account creation failed', true);
    }
  });
})();
