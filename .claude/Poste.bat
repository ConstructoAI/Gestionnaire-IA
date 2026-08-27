@echo off
setlocal
title Poste - courriels, calendrier, projets, comptabilite

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
echo      POSTE DE TRAVAIL
echo.
echo      Courriels Outlook, calendrier, projets, comptabilite.
echo      Claude lit, mesure, classe et redige.
echo.
echo      Une gracieusete de Sylvain Leduc - Constructo AI inc.
echo      Ecosysteme intelligent pour la construction au Quebec
echo      www.constructoai.ca
echo   ============================================================
echo.

rem ---------------------------------------------------------------------------
rem  [1/4] Outlook d'abord : il se charge pendant la mise a jour de Claude.
rem ---------------------------------------------------------------------------
echo   [1/4] Ouverture d'Outlook...
echo.

if not exist "%CFG%\scripts\outlook_mail.py" (
  echo   MOTEUR INTROUVABLE : %CFG%\scripts\
  echo.
  echo   Les scripts qui pilotent Outlook manquent. Les courriels et le
  echo   calendrier seront inaccessibles ; les dossiers restent lisibles.
  echo.
  pause
  set "MODE_DEGRADE=1"
  goto maj_claude
)

"%TASKLIST_EXE%" /FI "IMAGENAME eq OUTLOOK.EXE" | "%FIND_EXE%" /I "OUTLOOK.EXE" >nul
if not errorlevel 1 (
  echo   Outlook est deja lance.
  goto maj_claude
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
  echo   Le nouveau Outlook - celui du Microsoft Store - n'expose pas
  echo   MAPI/COM : le moteur ne sait pas le piloter.
  echo.
  echo   Ouvrez Outlook classique a la main si vous voulez les courriels.
  echo   On continue : les dossiers de projets fonctionnent sans lui.
  echo.
  set "MODE_DEGRADE=1"
  goto maj_claude
)

start "" "%OUTLOOK_EXE%"
echo   Lance : %OUTLOOK_EXE%

rem ---------------------------------------------------------------------------
rem  [2/4] Mise a jour de Claude Code pendant qu'Outlook demarre.
rem ---------------------------------------------------------------------------
:maj_claude
echo.
echo   [2/4] Mise a jour de Claude Code...
echo.

where claude.cmd >nul 2>&1
if errorlevel 1 (
  echo   claude.cmd est introuvable dans le PATH.
  echo   Installez Claude Code, puis relancez ce fichier.
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
rem  [3/4] Le processus ne suffit pas : c'est MAPI qui doit repondre.
rem        check_setup.py sort 0 quand la boite est reellement accessible.
rem ---------------------------------------------------------------------------
echo.
echo   [3/4] Attente d'Outlook...
echo.

if defined MODE_DEGRADE (
  echo   Ignoree : mode degrade, sans courriels ni calendrier.
  goto lancer_claude
)

set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY where py >nul 2>&1 && set "PY=py"

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
if not defined PY goto mapi_non_verifie
%PY% "%CFG%\scripts\check_setup.py" >nul 2>&1
if errorlevel 1 goto attente

:outlook_pret
echo.
echo   Outlook repond. Boite accessible.
goto lancer_claude

rem  Le processus tourne mais MAPI n'a PAS pu etre interroge : sans Python, pas
rem  de check_setup.py. Ne PAS annoncer "Boite accessible" - ce serait affirmer
rem  une verification qui n'a pas eu lieu.
:mapi_non_verifie
echo.
echo   Outlook est lance, mais Python est introuvable dans le PATH :
echo   MAPI n'a PAS ete verifie. La boite peut ne pas repondre.
echo   Demandez a Claude de lancer check_setup.py s'il y a un doute.
goto lancer_claude

:attente_expiree
echo.
echo   Outlook met plus de temps que prevu a repondre. Diagnostic :
echo.
if defined PY %PY% "%CFG%\scripts\check_setup.py"
echo.
echo   On demarre Claude quand meme - demandez-lui de regarder.

rem ---------------------------------------------------------------------------
rem  [4/4] La session. Une seule, quatre metiers.
rem ---------------------------------------------------------------------------
:lancer_claude
echo.
echo   [4/4] Demarrage : Opus 5 contexte 1M, effort MAX, autonomie complete.
echo.

call claude.cmd --dangerously-skip-permissions --model "claude-opus-5[1m]" --effort max

endlocal
