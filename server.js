require('dotenv').config();

const https = require('https');
const fs = require('fs');
const path = require('path');
const express = require('express');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');

// ─── Configuration ────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
const MODEL = process.env.CLAUDE_MODEL || 'claude-sonnet-4-6';

if (!process.env.ANTHROPIC_API_KEY) {
  console.error('❌  ANTHROPIC_API_KEY manquante dans .env — arrêt du serveur.');
  process.exit(1);
}

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ─── Certificats HTTPS (générés par office-addin-dev-certs) ──────────────────
const certDir = path.join(process.env.USERPROFILE || process.env.HOME, '.office-addin-dev-certs');
let httpsOptions;
try {
  httpsOptions = {
    key:  fs.readFileSync(path.join(certDir, 'localhost.key')),
    cert: fs.readFileSync(path.join(certDir, 'localhost.crt')),
  };
} catch {
  console.error('❌  Certificats HTTPS introuvables dans', certDir);
  console.error('   Lance d\'abord : npx office-addin-dev-certs install');
  process.exit(1);
}

// ─── Prompt système de base ───────────────────────────────────────────────────
const SYSTEM_PROMPT_BASE = `Tu es un assistant de révision de courriels pour A&V Courtiers Hypothécaires (Viet Nguyen-Duong, courtier hypothécaire, Blainville QC).

Ton rôle : réviser, clarifier et améliorer les brouillons de courriels de Viet, en préservant son style et ses intentions.

## Identification du contexte

Détermine automatiquement :
- **Mode expéditeur** : viet@avhypotheques.com (personnel) ou info@avhypotheques.com (collectif)
  - Défaut si non précisé : Mode A (viet@)
- **Registre** : tutoiement ou vouvoiement
  - "Madame / Monsieur [Nom]" dans la salutation → vouvoiement partout ("vous", "votre")
  - "Salut [Prénom]" ou "Bonjour [Prénom]" → tutoiement partout ("tu", "ta", "ton")
  - Si courriel reçu fourni en contexte → calquer son niveau de familiarité
  - Si impossible à déduire → tutoiement par défaut
  - **Règle absolue** : cohérence complète entre salutation et corps — jamais de mélange

## Règles de voix

**Mode A — viet@ (personnel) :**
- "on" quand Viet parle en tant qu'A&V ("on te recommande", "on a regardé")
- Aller droit au but — jamais de remplissage ni de formules corporate
- 1 seule demande d'action à la fois si action requise
- Fermeture contextuelle ("Bonne fin de semaine!", "N'hésite pas si t'as des questions.", "Merci!")

**Mode B — info@ (collectif) :**
- Ton neutre, professionnel mais chaleureux
- Voix collective "on", sans tutoiement ni vouvoiement explicite

## Ce qui ne change jamais
- Préserver les faits, chiffres, et intentions de Viet — ne rien inventer
- Vocabulaire hypothécaire québécois naturel (taux, terme, amortissement, mise de fonds, prêteur, SCHL, équité, ratio ABD/ATD)
- Ne jamais promettre un taux ou une approbation sans que le dossier soit confirmé
- **Ne jamais inclure de signature dans le courriel révisé** — la signature est gérée séparément
- **Ne jamais modifier ni proposer de changement à l'objet du courriel** — l'objet est fourni pour contexte seulement

## Format de réponse obligatoire

Réponds TOUJOURS avec exactement cette structure Markdown :

### Diagnostic
[bullets sur les axes problématiques seulement — clarté, ton, structure, appel à l'action, longueur]
[Si le brouillon est solide, une ligne courte le confirmant]
[⚠️ Signale les risques de malentendu si pertinent]

### Courriel révisé
[courriel complet, prêt à copier-coller, SANS signature]

### Changements clés
[2 à 4 bullets max : **Ce qui a changé** — pourquoi ça améliore le courriel]
[Ne liste pas les corrections triviales]`;

// ─── Préférences apprises — cache mémoire + sauvegarde ───────────────────────
const PREFERENCES_FILE = path.join(__dirname, 'preferences.json');
let _prefsCache = null; // null = pas encore chargé

function loadPreferences() {
  if (_prefsCache !== null) return _prefsCache;
  try {
    if (fs.existsSync(PREFERENCES_FILE)) {
      const data = JSON.parse(fs.readFileSync(PREFERENCES_FILE, 'utf8'));
      _prefsCache = Array.isArray(data.preferences) ? data.preferences : [];
    } else {
      _prefsCache = [];
    }
  } catch { _prefsCache = []; }
  return _prefsCache;
}

function savePreferences(preferences) {
  _prefsCache = preferences;
  fs.writeFileSync(PREFERENCES_FILE, JSON.stringify({ preferences }, null, 2), 'utf8');
}

// Construire le prompt avec les préférences apprises
function buildSystemPrompt() {
  const prefs = loadPreferences();
  if (prefs.length === 0) return SYSTEM_PROMPT_BASE;
  const section = '\n\n## Préférences apprises de Viet\n\nCes règles ont été identifiées à partir des corrections que Viet a apportées aux révisions précédentes. Applique-les systématiquement :\n'
    + prefs.map(p => `- ${p}`).join('\n');
  return SYSTEM_PROMPT_BASE + section;
}

// ─── Express app ──────────────────────────────────────────────────────────────
const app = express();

// Autoriser GitHub Pages + localhost, et Private Network Access (Chrome/Edge 94+)
app.use(cors({
  origin: [
    'https://vietqnd-maker.github.io',
    'https://localhost:3000',
  ],
  methods: ['GET', 'POST', 'OPTIONS'],
}));
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Private-Network', 'true');
  next();
});
app.use(express.json({ limit: '50kb' }));
app.use(express.static(path.join(__dirname, 'addin')));

// ─── Route principale : révision de courriel ──────────────────────────────────
app.post('/api/reviser', async (req, res) => {
  const { body: emailBody, subject, from } = req.body;

  if (!emailBody || emailBody.trim().length === 0) {
    return res.status(400).json({ error: 'Corps du courriel manquant.' });
  }

  // Construire le contexte utilisateur
  let userMessage = 'Révise ce courriel :\n\n';
  if (subject) userMessage += `**Objet :** ${subject}\n`;
  if (from)    userMessage += `**De :** ${from}\n`;
  if (subject || from) userMessage += '\n';
  userMessage += emailBody;

  try {
    const response = await client.messages.create({
      model: MODEL,
      max_tokens: 2048,
      system: buildSystemPrompt(),
      messages: [{ role: 'user', content: userMessage }],
    });

    const revision = response.content.find(b => b.type === 'text')?.text ?? '';
    res.json({ revision });

  } catch (err) {
    console.error('Erreur Claude API:', err.message);

    if (err.status === 401) {
      return res.status(401).json({ error: 'Clé API invalide. Vérifie ton .env.' });
    }
    if (err.status === 429) {
      return res.status(429).json({ error: 'Limite de débit atteinte. Réessaie dans quelques secondes.' });
    }
    res.status(500).json({ error: 'Erreur serveur. Vérifie les logs.' });
  }
});

// ─── Route feedback : apprendre des modifications de Viet ────────────────────
app.post('/api/feedback', async (req, res) => {
  const { original, revised } = req.body;
  if (!original || !revised) return res.status(400).json({ error: 'Données manquantes.' });

  const analysisPrompt = `Viet Nguyen-Duong (courtier hypothécaire, A&V) a modifié la révision de courriel proposée avant de l'appliquer.

**Révision proposée :**
${original}

**Version finale de Viet :**
${revised}

Analyse les différences et identifie 1 à 3 préférences stylistiques ou règles de rédaction à retenir pour améliorer les prochaines révisions. Formule chaque préférence en une phrase courte, directe et actionnable (ex: "Préférer 'Tel que discuté' à 'Suite à notre échange'").

Réponds UNIQUEMENT avec un tableau JSON valide, sans texte autour : [{"preference": "..."}]`;

  try {
    // Paralléliser : appel Claude + lecture des préférences existantes
    const [response, existing] = await Promise.all([
      client.messages.create({
        model: MODEL,
        max_tokens: 512,
        messages: [{ role: 'user', content: analysisPrompt }],
      }),
      Promise.resolve(loadPreferences()),
    ]);

    const text = response.content.find(b => b.type === 'text')?.text ?? '';
    const jsonMatch = text.match(/\[[\s\S]*?\]/);
    if (!jsonMatch) return res.json({ saved: false, reason: 'Parsing échoué' });

    const newPrefs = JSON.parse(jsonMatch[0])
      .map(p => p.preference)
      .filter(p => typeof p === 'string' && p.trim());

    if (newPrefs.length === 0) return res.json({ saved: false, reason: 'Aucune préférence extraite' });

    // Combiner avec l'existant et demander à Claude de dédoublonner + détecter les incohérences
    const allPrefs = [...existing, ...newPrefs];

    const cleanupPrompt = `Voici une liste de préférences de rédaction apprises de Viet Nguyen-Duong (courtier hypothécaire, A&V Hypothèques).

${allPrefs.map((p, i) => `${i + 1}. ${p}`).join('\n')}

Tâche :
1. Combine les préférences redondantes ou similaires en une seule formulation plus précise
2. Identifie toute incohérence ou contradiction entre des préférences (ex : deux règles opposées)
3. Garde au maximum 20 préférences, les plus spécifiques et utiles en priorité

Réponds UNIQUEMENT avec un JSON valide sans aucun texte autour :
{"preferences":["règle 1","règle 2"],"incoherences":["description si applicable"]}`;

    const cleanupResponse = await client.messages.create({
      model: MODEL,
      max_tokens: 1024,
      messages: [{ role: 'user', content: cleanupPrompt }],
    });

    const cleanupText = cleanupResponse.content.find(b => b.type === 'text')?.text ?? '';
    const cleanupJson = cleanupText.match(/\{[\s\S]*\}/);
    if (!cleanupJson) {
      // Fallback : déduplication simple et troncature
      const updated = [...new Set(allPrefs)].slice(-20);
      savePreferences(updated);
      return res.json({ saved: true, learned: newPrefs, incoherences: [] });
    }

    const { preferences: cleaned, incoherences = [] } = JSON.parse(cleanupJson[0]);
    savePreferences(cleaned);

    if (incoherences.length > 0) {
      console.warn(`⚠️  Incohérences détectées dans les préférences :`, incoherences);
    }
    console.log(`✅  Feedback appris — ${cleaned.length} règle(s) après dédup :`, newPrefs);
    res.json({ saved: true, learned: newPrefs, incoherences });

  } catch (err) {
    console.error('Erreur analyse feedback:', err.message);
    if (err.status === 401) return res.status(401).json({ error: 'Clé API invalide.' });
    if (err.status === 429) return res.status(429).json({ error: 'Limite de débit atteinte.' });
    res.status(500).json({ error: 'Erreur analyse feedback.' });
  }
});

// ─── Health check ─────────────────────────────────────────────────────────────
app.get('/health', (_, res) => res.json({ status: 'ok', model: MODEL }));

// ─── Démarrage HTTPS ──────────────────────────────────────────────────────────
https.createServer(httpsOptions, app).listen(PORT, () => {
  console.log(`\n✅  Serveur A&V démarré → https://localhost:${PORT}`);
  console.log(`   Modèle Claude : ${MODEL}`);
  console.log(`   Prêt à réviser des courriels dans Outlook\n`);
});
