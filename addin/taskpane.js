/* ─── Configuration ─────────────────────────────────────────────────────────── */
const SERVER_URL = 'https://localhost:3000';

/* ─── Initialisation Office.js ───────────────────────────────────────────────── */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadEmailInfo();
    document.getElementById('btnApply').addEventListener('click', appliquerRevision);
    document.getElementById('btnFullscreen').addEventListener('click', ouvrirPleinEcran);

    // Toggle Changements clés
    document.getElementById('toggleChangements').addEventListener('click', () => {
      const body = document.getElementById('changements');
      const arrow = document.querySelector('#toggleChangements .toggle-arrow');
      const isOpen = body.style.display !== 'none';
      body.style.display = isOpen ? 'none' : 'block';
      arrow.classList.toggle('open', !isOpen);
    });

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
    // Lire le corps en HTML pour préserver la mise en forme et détecter la signature
    const rawHtml = await getEmailBodyHtml(item);

    // Extraire la signature HTML et le corps texte à envoyer à Claude
    const { bodyHtml, signatureHtml } = splitSignatureHtml(rawHtml);
    window._emailSignature = signatureHtml;

    // Convertir le HTML en texte brut pour Claude
    const emailBodyText = htmlToPlainText(bodyHtml);

    // Métadonnées — APIs différentes en compose vs lecture
    const subject = await getSubject(item);
    const from = item.from?.emailAddress || Office.context.mailbox.userProfile.emailAddress || '';

    // Appel au serveur proxy local
    const response = await fetch(`${SERVER_URL}/api/reviser`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ body: emailBodyText, subject, from }),
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

/* ─── Lecture du corps en HTML ───────────────────────────────────────────────── */
function getEmailBodyHtml(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, (result) => {
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

/* ─── Détection et extraction de la signature Outlook (HTML) ─────────────────── */
function splitSignatureHtml(html) {
  // Outlook place la signature dans plusieurs patterns possibles selon la version
  const sigPatterns = [
    /(<div[^>]+id=["']Signature["'][^>]*>[\s\S]*)/i,          // Outlook desktop classique
    /(<div[^>]+id=["']appendonsend["'][^>]*>[\s\S]*)/i,       // Outlook 365 web
    /(<div[^>]+class=["'][^"']*signature[^"']*["'][^>]*>[\s\S]*)/i, // classe contenant "signature"
    /(<div[^>]+class=["'][^"']*Signature[^"']*["'][^>]*>[\s\S]*)/i, // classe avec majuscule
  ];

  for (const pattern of sigPatterns) {
    const match = html.match(pattern);
    if (match) {
      const sigStart = html.indexOf(match[1]);
      return {
        bodyHtml: html.substring(0, sigStart),
        signatureHtml: match[1],
      };
    }
  }

  // Fallback : pas de signature détectée — retourner tout le corps
  return { bodyHtml: html, signatureHtml: '' };
}

/* ─── Convertir HTML en texte brut pour Claude ───────────────────────────────── */
function htmlToPlainText(html) {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<li[^>]*>/gi, '- ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/* ─── Ajouter styles inline Outlook sur HTML natif de Claude ────────────────── */
function addOutlookStyles(html) {
  return html
    .replace(/<p(?![^>]*style)[^>]*>/gi, '<p style="margin:0 0 12px 0;">')
    .replace(/<ul(?![^>]*style)[^>]*>/gi, '<ul style="margin:0 0 12px 20px; padding-left:0;">')
    .replace(/<ol(?![^>]*style)[^>]*>/gi, '<ol style="margin:0 0 12px 20px; padding-left:0;">')
    .replace(/<li(?![^>]*style)[^>]*>/gi, '<li style="margin-bottom:4px;">');
}

/* ─── Convertir texte brut de Claude en HTML pour Outlook ───────────────────── */
function plainTextToHtml(text) {
  const lines = text.split('\n');
  const blocks = [];
  let i = 0;

  while (i < lines.length) {
    const trimmed = lines[i].trim();

    // Ligne vide ou séparateur --- → on saute
    if (!trimmed || /^-{3,}$/.test(trimmed)) { i++; continue; }

    // Blockquote > → paragraphe simple
    if (/^>\s*/.test(trimmed)) {
      blocks.push(`<p><em>${inlineFormat(trimmed.replace(/^>\s*/, ''))}</em></p>`);
      i++;
      continue;
    }

    // Liste numérotée (1. item)
    if (/^\d+\.\s+/.test(trimmed)) {
      const items = [];
      while (i < lines.length && /^\d+\.\s+/.test(lines[i].trim())) {
        items.push(`<li style="margin-bottom:4px;">${inlineFormat(lines[i].trim().replace(/^\d+\.\s+/, ''))}</li>`);
        i++;
      }
      blocks.push(`<ol style="margin:0 0 12px 0; padding-left:20px;">${items.join('')}</ol>`);
      continue;
    }

    // Liste à puces (- ou * ou • en début de ligne)
    if (/^[-*•]\s+/.test(trimmed)) {
      const items = [];
      while (i < lines.length && /^[-*•]\s+/.test(lines[i].trim())) {
        items.push(`<li style="margin-bottom:4px;">${inlineFormat(lines[i].trim().replace(/^[-*•]\s+/, ''))}</li>`);
        i++;
      }
      blocks.push(`<ul style="margin:0 0 12px 0; padding-left:20px;">${items.join('')}</ul>`);
      continue;
    }

    // Paragraphe — chaque ligne non-vide consécutive = même paragraphe
    const paraLines = [];
    while (i < lines.length) {
      const t = lines[i].trim();
      if (!t || /^-{3,}$/.test(t) || /^[-*•]\s+/.test(t) || /^\d+\.\s+/.test(t) || /^>\s*/.test(t)) break;
      paraLines.push(inlineFormat(t));
      i++;
    }
    if (paraLines.length > 0) {
      blocks.push(`<p style="margin:0 0 12px 0;">${paraLines.join('<br>')}</p>`);
    }
  }

  return blocks.join('');
}

/* ─── Formater les styles inline (gras, italique markdown) ──────────────────── */
function inlineFormat(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.*?)\*/g, '<em>$1</em>');
}

/* ─── Affichage des résultats ────────────────────────────────────────────────── */
function afficherRevision(markdown) {
  const sections = parseRevision(markdown);
  const editor   = document.getElementById('emailRevised');
  const btnApply = document.getElementById('btnApply');
  const btnFS    = document.getElementById('btnFullscreen');

  if (sections.noChanges) {
    // Courriel déjà bon — afficher message, masquer les boutons d'action
    editor.innerHTML = `<p style="color:#7EC8A0; margin:0;">✓ ${sections.noChanges}</p>`;
    editor.contentEditable = 'false';
    btnApply.style.display = 'none';
    btnFS.style.display    = 'none';
    window._originalHtml   = null;
    document.getElementById('sectionChangements').style.display = 'none';
  } else if (sections.revision) {
    // Courriel révisé — injecter HTML complet
    editor.innerHTML = addOutlookStyles(sections.revision);
    editor.contentEditable = 'true';
    btnApply.style.display = '';
    btnFS.style.display    = '';
    window._originalHtml   = editor.innerHTML;
    if (sections.changements) {
      document.getElementById('changements').innerHTML = formatMarkdown(sections.changements);
      document.getElementById('sectionChangements').style.display = '';
    }
  }

  showResult();
}

/* ─── Parser les sections Markdown de Claude ────────────────────────────────── */
function parseRevision(text) {
  const sections = { revision: '', changements: '', noChanges: '' };

  const noChgMatch = text.match(/###\s*Aucune correction recommandée\s*([\s\S]*?)$/i);
  const revMatch   = text.match(/###\s*Courriel révisé\s*([\s\S]*?)(?=###\s*Changements clés|$)/i);
  const chgMatch   = text.match(/###\s*Changements clés\s*([\s\S]*?)$/i);

  if (noChgMatch) sections.noChanges  = noChgMatch[1].trim() || 'Courriel déjà clair et bien tourné.';
  if (revMatch)   sections.revision   = revMatch[1].trim().replace(/(\n\s*---+\s*)+$/, '').trim();
  if (chgMatch)   sections.changements = chgMatch[1].trim();

  return sections;
}

/* ─── Convertir Markdown basique en HTML ─────────────────────────────────────── */
function formatMarkdown(text) {
  return text
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/^[-•]\s+(.+)/gm, '<li>$1</li>')
    .replace(/(⚠️[^\n]+)/g, '<span class="warning">$1</span>')
    .replace(/\n{2,}/g, '<br><br>')
    .replace(/\n/g, '<br>')
    .replace(/(<li>.*?<\/li>(\s*<br>)*)+/gs, (match) => {
      const items = match.replace(/<br>/g, '').trim();
      return `<ul>${items}</ul>`;
    });
}

/* ─── Appliquer le courriel révisé dans Outlook ─────────────────────────────── */
async function appliquerRevision() {
  const editor = document.getElementById('emailRevised');
  const btn = document.getElementById('btnApply');
  if (!editor.innerHTML.trim()) return;

  btn.disabled = true;
  btn.textContent = '…';

  try {
    // Relire le corps HTML au moment d'appliquer — signature toujours fraîche
    const currentHtml = await getEmailBodyHtml(Office.context.mailbox.item);
    const { signatureHtml } = splitSignatureHtml(currentHtml);

    // HTML du div contenteditable → injection directe avec styles Outlook
    const bodyHtml = addOutlookStyles(editor.innerHTML);
    const revisedHtml = bodyHtml + (signatureHtml || '');

    // Si modifié → envoyer le feedback pour apprendre
    if (window._originalHtml && editor.innerHTML !== window._originalHtml) {
      envoyerFeedback(window._originalHtml, editor.innerHTML);
    }

    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.setAsync(
        revisedHtml,
        { coercionType: Office.CoercionType.Html },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message || 'erreur inconnue'));
        }
      );
    });

    btn.textContent = '✓ Appliqué!';
    btn.classList.add('applied');
    setTimeout(() => {
      btn.textContent = '✓ Appliquer';
      btn.classList.remove('applied');
      btn.disabled = false;
    }, 2500);

  } catch (err) {
    showError('Impossible d\'appliquer : ' + err.message);
    btn.textContent = '✓ Appliquer';
    btn.disabled = false;
  }
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
    console.warn('Feedback non envoyé:', err.message);
  }
}


/* ─── Bouton plein écran ─────────────────────────────────────────────────────── */
function ouvrirPleinEcran() {
  const html = document.getElementById('emailRevised').innerHTML;
  if (!html.trim()) return;

  // Encoder le HTML en base64 pour le passer via l'URL hash
  const encoded = btoa(unescape(encodeURIComponent(html)));
  const dialogUrl = `https://vietqnd-maker.github.io/outlook-claude-addin/addin/dialog.html#${encoded}`;

  Office.context.ui.displayDialogAsync(dialogUrl, { height: 80, width: 65 }, (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      showError('Impossible d\'ouvrir la fenêtre plein écran.');
      return;
    }
    const dialog = result.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
      const data = JSON.parse(msg.message);
      if (data.action === 'apply') {
        document.getElementById('emailRevised').innerHTML = data.html;
        dialog.close();
        appliquerRevision();
      } else {
        dialog.close();
      }
    });
  });
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
