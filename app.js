(() => {
  const body = document.body;
  const menuBtn = document.getElementById('menuBtn');
  if (!menuBtn) return;

  menuBtn.addEventListener('click', () => {
    const open = body.classList.toggle('menu-open');
    menuBtn.setAttribute('aria-expanded', String(open));
  });

  document.querySelectorAll('#mainNav a').forEach(link => {
    link.addEventListener('click', () => {
      body.classList.remove('menu-open');
      menuBtn.setAttribute('aria-expanded', 'false');
    });
  });
})();
