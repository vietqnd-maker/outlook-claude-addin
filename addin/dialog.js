Office.onReady(() => {
  // Lire le HTML passé via le hash URL (base64 encodé)
  const hash = window.location.hash.slice(1);
  if (hash) {
    try {
      const html = decodeURIComponent(escape(atob(hash)));
      document.getElementById('content').innerHTML = html;
    } catch (e) {
      console.error('Erreur décodage:', e);
    }
  }

  document.getElementById('btnApply').addEventListener('click', () => {
    const html = document.getElementById('content').innerHTML;
    Office.context.ui.messageParent(JSON.stringify({ action: 'apply', html }));
  });

  document.getElementById('btnClose').addEventListener('click', () => {
    Office.context.ui.messageParent(JSON.stringify({ action: 'close' }));
  });
});
