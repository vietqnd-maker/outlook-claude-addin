@echo off
cd /d "%~dp0"
echo.
echo  ╔══════════════════════════════════════╗
echo  ║   A^&V — Révision de courriels        ║
echo  ╚══════════════════════════════════════╝
echo.

:: Vérifier que .env existe
if not exist ".env" (
  echo  ❌  Fichier .env introuvable.
  echo     Copie .env.example en .env et ajoute ta clé API Anthropic.
  echo.
  pause
  exit /b 1
)

:: Vérifier que node_modules existe
if not exist "node_modules" (
  echo  📦  Installation des dépendances...
  npm install
  echo.
)

echo  ✅  Démarrage du serveur A^&V...
echo  ℹ️   Garde cette fenêtre ouverte pendant que tu travailles dans Outlook.
echo  ℹ️   Ferme-la pour arrêter le serveur.
echo.
node server.js
pause
