#!/usr/bin/env python3
"""
validate_addin.py — Validates the Outlook add-in project for common issues.

Checks:
1. DOM consistency: every getElementById/querySelector in JS must have a matching ID in HTML
2. Manifest validity: runs npx office-addin-manifest validate
3. Server CORS config: GitHub Pages origin whitelisted + Private Network Access header present
"""

import io
import os
import re
import subprocess
import sys
from pathlib import Path

# Force UTF-8 sur Windows (évite les erreurs cp1252 avec les emojis)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PROJECT_ROOT = Path("C:/Users/vietq/Documents/outlook-claude-addin")
ADDIN_DIR = PROJECT_ROOT / "addin"
MANIFEST = PROJECT_ROOT / "manifest.xml"

issues = []

# ─── 1. DOM consistency check ────────────────────────────────────────────────

# Collect all IDs defined in HTML files
html_ids = set()
for html_file in ADDIN_DIR.glob("**/*.html"):
    content = html_file.read_text(encoding="utf-8", errors="ignore")
    for match in re.finditer(r'id=["\']([^"\']+)["\']', content):
        html_ids.add(match.group(1))

# Collect all getElementById / querySelector('#id') calls in JS files
js_refs = []  # list of (file, line_number, id_referenced)
for js_file in ADDIN_DIR.glob("**/*.js"):
    lines = js_file.read_text(encoding="utf-8", errors="ignore").splitlines()
    for lineno, line in enumerate(lines, start=1):
        # getElementById('id') or getElementById("id")
        for match in re.finditer(r'getElementById\(["\']([^"\']+)["\']\)', line):
            js_refs.append((js_file.name, lineno, match.group(1)))
        # querySelector('#id') — sélecteurs ID simples seulement (pas de .class ou espace)
        for match in re.finditer(r'querySelector\(["\']#([\w-]+)["\']\)', line):
            js_refs.append((js_file.name, lineno, match.group(1)))

dom_issues = []
for filename, lineno, ref_id in js_refs:
    if ref_id not in html_ids:
        dom_issues.append(f"  - {filename}:{lineno} — getElementById('{ref_id}') introuvable dans le HTML")

if dom_issues:
    issues.append(("DOM consistency", dom_issues))

# ─── 2. Manifest validation ──────────────────────────────────────────────────

manifest_issues = []
try:
    npx_path = r"C:\Program Files\nodejs\npx.cmd"
    result = subprocess.run(
        [npx_path, "office-addin-manifest", "validate", str(MANIFEST)],
        capture_output=True,
        text=True,
        timeout=60,
    )
    output = result.stdout + result.stderr
    if "The manifest is not valid" in output or result.returncode != 0:
        # Extract error lines
        error_lines = [l.strip() for l in output.splitlines() if "Error" in l or "Details" in l]
        manifest_issues = error_lines if error_lines else ["Manifest invalide (voir sortie complète)"]
        issues.append(("Manifest XML", manifest_issues))
except Exception as e:
    issues.append(("Manifest XML", [f"Impossible de lancer le validateur : {e}"]))

# ─── 3. Server CORS config ───────────────────────────────────────────────────

SERVER_JS = PROJECT_ROOT / "server.js"
TASKPANE_JS = ADDIN_DIR / "taskpane.js"
cors_issues = []

try:
    server_src = SERVER_JS.read_text(encoding="utf-8", errors="ignore")
    taskpane_src = TASKPANE_JS.read_text(encoding="utf-8", errors="ignore")

    # Extraire l'URL du serveur dans taskpane.js
    server_url_match = re.search(r"SERVER_URL\s*=\s*['\"]([^'\"]+)['\"]", taskpane_src)
    server_url = server_url_match.group(1) if server_url_match else None

    # Extraire les origines CORS autorisées dans server.js
    cors_origins = re.findall(r"['\"]https?://[^'\"]+['\"]", server_src[server_src.find("cors("):server_src.find("cors(")+500] if "cors(" in server_src else "")
    cors_origins = [o.strip("'\"") for o in cors_origins]

    # Vérifier que GitHub Pages est dans les origines CORS
    gh_pages_origins = [o for o in cors_origins if "github.io" in o]
    if not gh_pages_origins:
        cors_issues.append("  - server.js : GitHub Pages (github.io) absent des origines CORS autorisées")

    # Vérifier le header Private Network Access
    if "Access-Control-Allow-Private-Network" not in server_src:
        cors_issues.append("  - server.js : header 'Access-Control-Allow-Private-Network: true' manquant (bloque les requêtes depuis GitHub Pages vers localhost)")

    # Vérifier que SERVER_URL dans taskpane.js est cohérent avec le port du serveur
    port_match = re.search(r"PORT\s*=\s*process\.env\.PORT\s*\|\|\s*(\d+)", server_src)
    server_port = port_match.group(1) if port_match else "3000"
    if server_url and f":{server_port}" not in server_url:
        cors_issues.append(f"  - taskpane.js : SERVER_URL='{server_url}' ne correspond pas au port {server_port} du serveur")

except Exception as e:
    cors_issues.append(f"  - Impossible de lire server.js : {e}")

if cors_issues:
    issues.append(("Configuration serveur (CORS)", cors_issues))

# ─── Report ──────────────────────────────────────────────────────────────────

print("\n=== Validation add-in ===\n")

if not issues:
    print("✅ PASS — Aucun problème trouvé. Prêt à pusher.\n")
    sys.exit(0)

for section, section_issues in issues:
    print(f"❌ {section} — {len(section_issues)} problème(s) :")
    for issue in section_issues:
        print(issue)
    print()

total = sum(len(s) for _, s in issues)
print(f"❌ FAIL — {total} problème(s) à corriger avant de pusher.\n")
sys.exit(1)
