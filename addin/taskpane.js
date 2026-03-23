/* ─── Configuration ─────────────────────────────────────────────────────────── */
const SERVER_URL = 'https://localhost:3000';

/* ─── Initialisation Office.js ───────────────────────────────────────────────── */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadEmailInfo();
    document.getElementById('btnApply').addEventListener('click', appliquerRevision);
    document.getElementById('toggleDiagnostic').addEventListener('click', () => {
      const body = document.getElementById('diagnostic');
      const arrow = document.querySelector('#toggleDiagnostic .toggle-arrow');
      const isOpen = body.style.display !== 'none';
      body.style.display = isOpen ? 'none' : 'block';
      arrow.classList.toggle('open', !isOpen);
    });
    // Lancer la révision automatiquement dès l'ouverture du panneau
    reviserCourriel();
  }
});

/* ─── Lecture du courriel ouvert ─────────────────────────────────────────────── */
function loadEmailInfo() {
  const item = Office.context.mailbox.item;
  const subjectEl = document.getElementById('emailSubject');
  if (!item) {
    subjectEl.textContent = 'Aucun courriel sélectionné';
    return;
  }
  subjectEl.textContent = `📧 ${item.subject || '(sans objet)'}`;
}

/* ─── Révision principale ────────────────────────────────────────────────────── */
async function reviserCourriel() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  setLoading(true);
  hideError();
  hideResult();

  // Réinitialiser l'état de la session précédente
  window._emailSignature = null;
  window._originalRevision = null;

  try {
    // Lire le corps du courriel en texte brut
    const rawBody = await getEmailBody(item);

    // Extraire et mettre de côté la signature
    const { body: emailBody, signature } = splitSignature(rawBody);
    window._emailSignature = signature; // conservée pour réinsertion

    // Métadonnées — APIs différentes en compose vs lecture
    const subject = await getSubject(item);
    const from = item.from?.emailAddress || Office.context.mailbox.userProfile.emailAddress || '';

    // Appel au serveur proxy local (corps sans signature)
    const response = await fetch(`${SERVER_URL}/api/reviser`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ body: emailBody, subject, from }),
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({ error: `Erreur HTTP ${response.status}` }));
      throw new Error(err.error || `Erreur ${response.status}`);
    }

    const { revision } = await response.json();
    afficherRevision(revision);

  } catch (err) {
    showError(err.message || 'Erreur inattendue. Vérifie que le serveur est démarré.');
  } finally {
    setLoading(false);
  }
}

/* ─── Lecture du corps du courriel (Promise) ────────────────────────────────── */
function getEmailBody(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, { asyncContext: 'body' }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || '');
      } else {
        reject(new Error('Impossible de lire le corps du courriel.'));
      }
    });
  });
}

/* ─── Lecture du sujet (compose = objet async, lecture = string) ─────────────── */
function getSubject(item) {
  if (typeof item.subject === 'string') return Promise.resolve(item.subject);
  return new Promise((resolve) => {
    item.subject.getAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || '' : '');
    });
  });
}

/* ─── Détection et extraction de la signature ───────────────────────────────── */
function splitSignature(body) {
  // Séparateur standard --
  const dashSep = body.search(/\n--[ \t]*\n/);
  if (dashSep !== -1) {
    return { body: body.substring(0, dashSep).trim(), signature: body.substring(dashSep) };
  }

  // Heuristique : détecter la signature après la formule de politesse + nom
  // (ex: "Bonne journée,\n\nViet Nguyen-Duong\nCourtier...")
  const closingMatch = body.match(
    /(Cordialement|Bonne journée|Bonne fin|Merci|À bientôt|Sincèrement)[^\n]*\n[\s\S]{0,10}\n([\s\S]+)$/i
  );
  if (closingMatch && closingMatch[2] && closingMatch[2].trim().split('\n').length >= 2) {
    const sigStart = body.lastIndexOf(closingMatch[2]);
    return {
      body: body.substring(0, sigStart).trim(),
      signature: '\n\n' + closingMatch[2].trim()
    };
  }

  return { body: body.trim(), signature: '' };
}

/* ─── Affichage des résultats ────────────────────────────────────────────────── */
function afficherRevision(markdown) {
  // Parser les 3 sections du format de réponse Claude
  const sections = parseRevision(markdown);

  if (sections.diagnostic) {
    document.getElementById('diagnostic').innerHTML = formatMarkdown(sections.diagnostic);
  }
  if (sections.revision) {
    // textarea — on assigne .value directement (éditable par Viet avant d'appliquer)
    document.getElementById('emailRevised').value = sections.revision;
    // Mémoriser la version originale de Claude pour détecter les modifications de Viet
    window._originalRevision = sections.revision;
  }
  if (sections.changements) {
    document.getElementById('changements').innerHTML = formatMarkdown(sections.changements);
  }

  showResult();
}

/* ─── Parser les sections Markdown de Claude ────────────────────────────────── */
function parseRevision(text) {
  const sections = { diagnostic: '', revision: '', changements: '' };

  const diagMatch = text.match(/###\s*Diagnostic\s*([\s\S]*?)(?=###\s*Courriel révisé|$)/i);
  const revMatch  = text.match(/###\s*Courriel révisé\s*([\s\S]*?)(?=###\s*Changements clés|$)/i);
  const chgMatch  = text.match(/###\s*Changements clés\s*([\s\S]*?)$/i);

  if (diagMatch) sections.diagnostic  = diagMatch[1].trim();
  if (revMatch)  sections.revision    = revMatch[1].trim();
  if (chgMatch)  sections.changements = chgMatch[1].trim();

  // Fallback : si le parsing échoue, afficher le tout dans revision
  if (!sections.revision && text.trim()) {
    sections.revision = text.trim();
  }

  return sections;
}

/* ─── Convertir Markdown basique en HTML ─────────────────────────────────────── */
function formatMarkdown(text) {
  return text
    // Gras **texte**
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    // Bullets - texte
    .replace(/^[-•]\s+(.+)/gm, '<li>$1</li>')
    // Alertes ⚠️
    .replace(/(⚠️[^\n]+)/g, '<span class="warning">$1</span>')
    // Sauts de ligne
    .replace(/\n{2,}/g, '<br><br>')
    .replace(/\n/g, '<br>')
    // Envelopper les <li> dans <ul>
    .replace(/(<li>.*?<\/li>(\s*<br>)*)+/gs, (match) => {
      const items = match.replace(/<br>/g, '').trim();
      return `<ul>${items}</ul>`;
    });
}

/* ─── Appliquer le courriel révisé dans Outlook ─────────────────────────────── */
function appliquerRevision() {
  const text = document.getElementById('emailRevised').value;
  const btn = document.getElementById('btnApply');
  if (!text) return;

  // Réinsérer la signature qui avait été mise de côté
  const fullText = window._emailSignature ? text + '\n\n' + window._emailSignature : text;

  // Si Viet a modifié la révision → envoyer le feedback pour apprendre
  if (window._originalRevision && text.trim() !== window._originalRevision.trim()) {
    envoyerFeedback(window._originalRevision, text);
  }

  Office.context.mailbox.item.body.setAsync(
    fullText,
    { coercionType: Office.CoercionType.Text },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        btn.textContent = '✓ Appliqué!';
        btn.classList.add('applied');
        setTimeout(() => {
          btn.textContent = '✓ Appliquer';
          btn.classList.remove('applied');
        }, 2500);
      } else {
        showError('Impossible d\'appliquer : ' + (result.error?.message || 'erreur inconnue'));
      }
    }
  );
}

/* ─── Feedback d'apprentissage — envoyer les modifications de Viet ───────────── */
async function envoyerFeedback(original, revised) {
  try {
    const resp = await fetch(`${SERVER_URL}/api/feedback`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ original, revised }),
    });
    const data = await resp.json().catch(() => ({}));
    if (data.incoherences && data.incoherences.length > 0) {
      showError('⚠️ Incohérence dans tes préférences apprises : ' + data.incoherences.join(' — '));
    }
  } catch (err) {
    // Silencieux — le feedback n'est pas critique
    console.warn('Feedback non envoyé:', err.message);
  }
}

/* ─── Helpers UI ─────────────────────────────────────────────────────────────── */
function setLoading(show) {
  document.getElementById('loading').style.display = show ? 'flex' : 'none';
}

function showResult() {
  document.getElementById('resultContainer').style.display = 'flex';
}

function hideResult() {
  document.getElementById('resultContainer').style.display = 'none';
}

function showError(msg) {
  const box = document.getElementById('errorBox');
  document.getElementById('errorMessage').textContent = msg;
  box.style.display = 'block';
}

function hideError() {
  document.getElementById('errorBox').style.display = 'none';
}
