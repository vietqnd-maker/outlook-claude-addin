#!/usr/bin/env python3
"""
validate_addin.py — Validates the Outlook add-in project for common issues.

Checks:
1. DOM consistency: every getElementById/querySelector in JS must have a matching ID in HTML
2. Manifest validity: runs npx office-addin-manifest validate
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
