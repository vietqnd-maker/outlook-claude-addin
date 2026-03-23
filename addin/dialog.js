Office.onReady(() => {
  // Lire le texte passé via le hash URL (base64 encodé)
  const hash = window.location.hash.slice(1);
  if (hash) {
    try {
      const text = decodeURIComponent(escape(atob(hash)));
      document.getElementById('content').value = text;
    } catch (e) {
      console.error('Erreur décodage:', e);
    }
  }

  document.getElementById('btnApply').addEventListener('click', () => {
    const text = document.getElementById('content').value;
    Office.context.ui.messageParent(JSON.stringify({ action: 'apply', text }));
  });

  document.getElementById('btnClose').addEventListener('click', () => {
    Office.context.ui.messageParent(JSON.stringify({ action: 'close' }));
  });
});
