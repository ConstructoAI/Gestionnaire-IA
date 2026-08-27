@echo off
setlocal
title Constructo AI - courriels, calendrier, projets, comptabilite

rem ---------------------------------------------------------------------------
rem  OU DOIT VIVRE CE FICHIER
rem  A cote du dossier .claude, pas dedans. Mais on le livre DANS .claude pour
rem  qu'un seul dossier suffise a tout transporter : ce bloc detecte les deux cas
rem  et se place toujours au bon endroit. Le dossier de travail doit etre le
rem  PARENT de .claude, sinon Claude Code chercherait sa configuration dans
rem  .claude\.claude\ et ne trouverait rien.
rem ---------------------------------------------------------------------------
cd /d "%~dp0"
for %%I in ("%CD%") do set "NOMDOSSIER=%%~nxI"
if /I "%NOMDOSSIER%"==".claude" cd ..

set "CFG=%CD%\.claude"
if not exist "%CFG%\CLAUDE.md" (
  echo.
  echo   INTROUVABLE : %CFG%\CLAUDE.md
  echo.
  echo   Ce fichier doit se trouver dans le dossier .claude, ou juste a cote.
  echo   Copiez le dossier .claude dans votre dossier de travail, puis relancez.
  echo.
  pause
  exit /b 1
)

rem  Chemins complets : Git for Windows livre un `find` Unix qui masque celui
rem  de Windows des que son dossier bin se trouve dans le PATH.
set "TASKLIST_EXE=%SystemRoot%\System32\tasklist.exe"
set "FIND_EXE=%SystemRoot%\System32\find.exe"

echo.
echo   ============================================================
echo                      C O N S T R U C T O   A I
echo.
echo      Courriels Outlook, calendrier, projets, comptabilite.
echo      Claude lit, mesure, classe et redige.
echo.
echo      Une gracieusete de Sylvain Leduc, president
echo      Ecosysteme intelligent pour la construction au Quebec
echo      www.constructoai.ca
echo   ============================================================
echo.

rem ---------------------------------------------------------------------------
rem  [1/5] Outlook d'abord : il se charge pendant tout le reste.
rem ---------------------------------------------------------------------------
echo   [1/5] Ouverture d'Outlook...
echo.

if not exist "%CFG%\scripts\outlook_mail.py" (
  echo   MOTEUR INTROUVABLE : %CFG%\scripts\
  echo.
  echo   Les scripts qui pilotent Outlook manquent. Les courriels et le
  echo   calendrier seront inaccessibles ; les dossiers restent lisibles.
  echo.
  pause
  set "MODE_DEGRADE=1"
  goto dependances
)

"%TASKLIST_EXE%" /FI "IMAGENAME eq OUTLOOK.EXE" | "%FIND_EXE%" /I "OUTLOOK.EXE" >nul
if not errorlevel 1 (
  echo   Outlook est deja lance.
  goto dependances
)

set "PF=%ProgramFiles%"
set "PFX86=%ProgramFiles(x86)%"
set "OUTLOOK_EXE="
for %%P in (
  "%PF%\Microsoft Office\root\Office16\OUTLOOK.EXE"
  "%PF%\Microsoft Office\Office16\OUTLOOK.EXE"
  "%PF%\Microsoft Office\Office15\OUTLOOK.EXE"
  "%PFX86%\Microsoft Office\root\Office16\OUTLOOK.EXE"
  "%PFX86%\Microsoft Office\Office16\OUTLOOK.EXE"
  "%PFX86%\Microsoft Office\Office15\OUTLOOK.EXE"
) do if not defined OUTLOOK_EXE if exist %%P set "OUTLOOK_EXE=%%~P"

if not defined OUTLOOK_EXE (
  echo   Outlook classique est introuvable aux emplacements habituels.
  echo   Le nouveau Outlook, celui du Microsoft Store, n'expose pas
  echo   MAPI/COM : le moteur ne sait pas le piloter.
  echo.
  echo   Ouvrez Outlook classique a la main si vous voulez les courriels.
  echo   On continue : les dossiers de projets fonctionnent sans lui.
  echo.
  set "MODE_DEGRADE=1"
  goto dependances
)

start "" "%OUTLOOK_EXE%"
echo   Lance : %OUTLOOK_EXE%

rem ---------------------------------------------------------------------------
rem  [2/5] Dependances. UN DOUBLE-CLIC DOIT SUFFIRE : si pywin32 manque, on
rem        l'installe ici plutot que d'exiger une commande pip de l'utilisateur.
rem        C'est la SEULE dependance du poste.
rem ---------------------------------------------------------------------------
:dependances
echo.
echo   [2/5] Verification des dependances...
echo.

set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY where py >nul 2>&1 && set "PY=py"

if not defined PY (
  echo   Python est introuvable dans le PATH.
  echo.
  echo   Les courriels, le calendrier et la comptabilite en ont besoin.
  echo   Installez Python depuis www.python.org - cochez bien
  echo   "Add python.exe to PATH" pendant l'installation - puis relancez.
  echo.
  echo   On continue : les dossiers de projets fonctionnent sans Python.
  echo.
  set "MODE_DEGRADE=1"
  goto maj_claude
)

%PY% -c "import win32com.client" >nul 2>&1
if not errorlevel 1 (
  echo   pywin32 est deja installe.
  goto maj_claude
)

echo   pywin32 est absent - installation automatique en cours...
echo   Une seule fois, environ trente secondes.
echo.
%PY% -m pip install --quiet --disable-pip-version-check pywin32
%PY% -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
  echo.
  echo   L'installation automatique a echoue.
  echo   Ouvrez une invite de commandes et lancez :
  echo.
  echo       %PY% -m pip install pywin32
  echo.
  echo   Cause frequente : pas de connexion, ou un pare-feu d'entreprise.
  echo.
  pause
  set "MODE_DEGRADE=1"
) else (
  echo   pywin32 installe.
)

rem ---------------------------------------------------------------------------
rem  [3/5] Mise a jour de Claude Code pendant qu'Outlook demarre.
rem ---------------------------------------------------------------------------
:maj_claude
echo.
echo   [3/5] Mise a jour de Claude Code...
echo.

where claude.cmd >nul 2>&1
if errorlevel 1 (
  echo   claude.cmd est introuvable dans le PATH.
  echo   Installez Claude Code, puis relancez ce fichier :
  echo   https://claude.com/claude-code
  echo.
  pause
  exit /b 1
)

call claude.cmd update
if errorlevel 1 (
  echo.
  echo   Mise a jour impossible - on continue quand meme.
  echo   Cause frequente : une autre session Claude est ouverte,
  echo   Windows verrouille alors le fichier en cours d'utilisation.
)

rem ---------------------------------------------------------------------------
rem  [4/5] Le processus ne suffit pas : c'est MAPI qui doit repondre.
rem        check_setup.py sort 0 quand la boite est reellement accessible.
rem ---------------------------------------------------------------------------
echo.
echo   [4/5] Attente d'Outlook...
echo.

if defined MODE_DEGRADE (
  echo   Ignoree : mode degrade, sans courriels ni calendrier.
  goto lancer_claude
)

set /a TENTATIVE=0

:attente
set /a TENTATIVE+=1
if %TENTATIVE% GTR 20 goto attente_expiree
<nul set /p "=."
rem  ping plutot que timeout : timeout refuse de dormir quand l'entree est
rem  redirigee (raccourci, tache planifiee). ping dort ~2s partout.
ping -n 3 127.0.0.1 >nul
"%TASKLIST_EXE%" /FI "IMAGENAME eq OUTLOOK.EXE" | "%FIND_EXE%" /I "OUTLOOK.EXE" >nul
if errorlevel 1 goto attente
%PY% "%CFG%\scripts\check_setup.py" >nul 2>&1
if errorlevel 1 goto attente

:outlook_pret
echo.
echo   Outlook repond. Boite accessible.
goto lancer_claude

:attente_expiree
echo.
echo   Outlook met plus de temps que prevu a repondre. Diagnostic :
echo.
%PY% "%CFG%\scripts\check_setup.py"
echo.
echo   On demarre Claude quand meme - demandez-lui de regarder.

rem ---------------------------------------------------------------------------
rem  [5/5] La session. Une seule, quatre metiers.
rem ---------------------------------------------------------------------------
:lancer_claude
echo.
echo   [5/5] Demarrage : Opus 5 contexte 1M, effort MAX, autonomie complete.
echo.

call claude.cmd --dangerously-skip-permissions --model "claude-opus-5[1m]" --effort max

endlocal
