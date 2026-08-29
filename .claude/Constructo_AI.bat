@echo off
setlocal enabledelayedexpansion
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

rem  Dossier courant NORMALISE, avec un antislash final garanti. Sert a ecarter
rem  un executable pose dans le dossier de travail (voir les sous-routines).
rem  Corrige le 2026-08-29 : a la racine d'un lecteur, %CD% vaut deja "X:\" et
rem  la comparaison "%CD%\" testait "X:\" contre "X:\\" - jamais vrai. Mesure :
rem  avec `subst X:` puis `cd X:\`, un faux claude.exe du dossier courant etait
rem  RETENU. Cas reel : lecteur reseau mappe, SharePoint monte en racine.
set "CDS=%CD%"
if not "%CDS:~-1%"=="\" set "CDS=%CDS%\"

rem ---------------------------------------------------------------------------
rem  RIEN N'EST CODE EN DUR : CHAQUE POSTE EST DIFFERENT
rem  Tous les chemins ci-dessous sont ABSOLUS mais CALCULES a l'execution.
rem  %SystemRoot%, %USERPROFILE% et %LOCALAPPDATA% sont resolus par Windows sur
rem  chaque machine : C:\Users\Marie chez Marie, D:\Profils\jd ailleurs. Le
rem  fichier reste identique d'un poste a l'autre.
rem
rem  Chemins complets pour tasklist et find : Git for Windows livre un `find`
rem  Unix qui masque celui de Windows des que son dossier bin est dans le PATH.
rem  powershell.exe en chemin absolu : sur un poste avec WSL, `bash` du PATH
rem  resout vers WSL et `pwsh` est souvent absent - voir CLAUDE.md section 8.
rem ---------------------------------------------------------------------------
set "TASKLIST_EXE=%SystemRoot%\System32\tasklist.exe"
set "FIND_EXE=%SystemRoot%\System32\find.exe"
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

rem  L'installateur natif de Claude Code depose claude.exe ICI, sans droits
rem  administrateur. Une session cmd deja ouverte ne voit pas le PATH mis a jour
rem  par l'installateur : on l'ajoute nous-memes, sinon on croirait a tort que
rem  Claude Code est absent juste apres l'avoir installe.
set "CLAUDE_BIN=%USERPROFILE%\.local\bin"
if exist "%CLAUDE_BIN%" set "PATH=%CLAUDE_BIN%;%PATH%"

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

rem ===========================================================================
rem  [1/6] INVENTAIRE - on regarde AVANT de proposer quoi que ce soit.
rem        Rien n'est installe a cette etape. Aucun effet de bord.
rem ===========================================================================
echo   [1/6] Inventaire de l'outillage...
echo.

set "MANQUE_CLAUDE="
set "MANQUE_PYTHON="
set "MANQUE_PYWIN32="
set "MANQUE_GIT="
set "RIEN_A_FAIRE=1"

rem  Claude Code : on resout un CHEMIN ABSOLU une fois pour toutes, et on exige
rem  que l'executable REPONDE. On ne teste JAMAIS `claude.cmd` : mesure du
rem  2026-08-29, l'installateur natif produit claude.exe et le nom claude.cmd
rem  n'existe qu'avec une installation npm. L'ancienne version de ce fichier
rem  testait claude.cmd et refusait de demarrer sur un poste ou Claude Code
rem  etait pourtant installe et fonctionnel.
call :resoudre_claude
if not defined CLAUDE_EXE (
  echo     Claude Code . . . . . ABSENT
  set "MANQUE_CLAUDE=1"
  set "RIEN_A_FAIRE="
) else (
  echo     Claude Code . . . . . present   !VER_CLAUDE!
)

call :resoudre_python
if not defined PY (
  echo     Python  . . . . . . . ABSENT
  where python >nul 2>&1 && echo                             ^(raccourci Microsoft Store ignore : ce n'est pas un interpreteur^)
  echo     pywin32 . . . . . . . ABSENT   -- Python requis d'abord
  set "MANQUE_PYTHON=1"
  set "MANQUE_PYWIN32=1"
  set "RIEN_A_FAIRE="
) else (
  echo     Python  . . . . . . . present   !VER_PY!
  rem  `call` obligatoire : si `python` resout vers un .bat ou un .cmd - shim
  rem  pyenv-win, pixi, conda - une invocation nue transfererait le controle et
  rem  ne reviendrait JAMAIS. Mesure du 2026-08-29 : le script s'arretait net
  rem  apres cette ligne, code de sortie 0, sans un mot.
  call "!PY!" -c "import win32com.client" >nul 2>&1
  if errorlevel 1 (
    echo     pywin32 . . . . . . . ABSENT
    set "MANQUE_PYWIN32=1"
    set "RIEN_A_FAIRE="
  ) else (
    echo     pywin32 . . . . . . . present
  )
)

rem  Git pour Windows : OPTIONNEL, mais il porte l'outil Bash de Claude Code.
rem  Sans lui, Claude Code bascule sur PowerShell - or les commandes des
rem  sections 2 et 3 de CLAUDE.md sont ecrites en shell POSIX.
call :resoudre_git
if not defined GIT_EXE (
  echo     Git pour Windows  . . ABSENT   -- optionnel, mais recommande ici
  set "MANQUE_GIT=1"
  set "RIEN_A_FAIRE="
) else (
  echo     Git pour Windows  . . present
)

rem  Le moteur : les SIX scripts, pas seulement le premier. Un moteur ampute de
rem  check_setup.py passerait le controle puis ferait echouer l'etape [5/6] pour
rem  une raison invisible.
set "MOTEUR_MANQUANT="
for %%S in (outlook_mail outlook_calendar veille_poste check_setup factures ost_reader) do (
  if not exist "%CFG%\scripts\%%S.py" set "MOTEUR_MANQUANT=!MOTEUR_MANQUANT! %%S.py"
)
if defined MOTEUR_MANQUANT (
  echo     Moteur  . . . . . . . INCOMPLET :!MOTEUR_MANQUANT!
) else (
  echo     Moteur  . . . . . . . present   6 scripts
)

rem  Outlook classique : on ne peut pas l'installer, seulement le signaler.
set "OUTLOOK_EXE="
set "PF=%ProgramFiles%"
set "PFX86=%ProgramFiles(x86)%"
for %%P in (
  "%PF%\Microsoft Office\root\Office16\OUTLOOK.EXE"
  "%PF%\Microsoft Office\Office16\OUTLOOK.EXE"
  "%PF%\Microsoft Office\Office15\OUTLOOK.EXE"
  "%PFX86%\Microsoft Office\root\Office16\OUTLOOK.EXE"
  "%PFX86%\Microsoft Office\Office16\OUTLOOK.EXE"
  "%PFX86%\Microsoft Office\Office15\OUTLOOK.EXE"
) do if not defined OUTLOOK_EXE if exist %%P set "OUTLOOK_EXE=%%~P"
if not defined OUTLOOK_EXE (
  echo     Outlook classique . . INTROUVABLE  -- ne s'installe pas d'ici
) else (
  echo     Outlook classique . . present
)

rem  De quoi installer. curl est livre avec Windows 10 1803 et plus.
set "A_CURL="
where curl >nul 2>&1 && set "A_CURL=1"
set "A_WINGET="
where winget >nul 2>&1 && set "A_WINGET=1"

echo.

rem ===========================================================================
rem  [2/6] INSTALLATION - une seule question, pour tout le lot.
rem        Niveau utilisateur : aucun droit administrateur demande.
rem ===========================================================================
echo   [2/6] Installation de ce qui manque...
echo.

if defined RIEN_A_FAIRE (
  echo   Tout est deja en place. Rien a installer.
  echo.
  goto ouvrir_outlook
)

echo   Il manque :
if defined MANQUE_CLAUDE  echo     - Claude Code       installateur officiel, sans droits admin
if defined MANQUE_PYTHON  echo     - Python 3          requis pour les courriels et le calendrier
if defined MANQUE_PYWIN32 echo     - pywin32           le pont vers Outlook
if defined MANQUE_GIT     echo     - Git pour Windows  optionnel : donne l'outil Bash a Claude
echo.
echo   Tout s'installe pour VOTRE compte seulement. Aucune elevation UAC,
echo   rien n'est modifie pour les autres utilisateurs du poste.
echo.

rem  Pre-remplie a N : si l'entree est redirigee - raccourci, tache planifiee -
rem  set /p laisse la variable inchangee. On n'installe alors RIEN plutot que
rem  d'installer des logiciels sans accord.
rem  Le prompt n'est pas indente : cmd retire les espaces de tete.
set "REPONSE=N"
set /p "REPONSE=Installer maintenant ? [O/N, defaut N] "
echo.

rem  Neutraliser les guillemets AVANT tout test : une reponse contenant " faisait
rem  sortir le script en 255 avec " La syntaxe de la commande n'est pas correcte ",
rem  fenetre fermee sur une erreur systeme. Mesure du 2026-08-29.
set "REPONSE=%REPONSE:"=%"
if not defined REPONSE set "REPONSE=N"

rem  Accepter ce qu'un francophone tape naturellement : O, o, oui, OUI, Y, yes.
rem  Mesure du 2026-08-29 : `oui` etait traite comme un REFUS, sans rien signaler.
set "ACCORD="
if /I "%REPONSE%"=="O"   set "ACCORD=1"
if /I "%REPONSE%"=="OUI" set "ACCORD=1"
if /I "%REPONSE%"=="Y"   set "ACCORD=1"
if /I "%REPONSE%"=="YES" set "ACCORD=1"

if not defined ACCORD (
  echo   Rien n'a ete installe.
  set "MODE_DEGRADE=1"
  if defined MANQUE_CLAUDE (
    echo.
    echo   Or Claude Code est justement ce qui manque : sans lui il n'y a
    echo   aucune session a demarrer. Le poste ne peut pas continuer.
    echo.
    echo   Relancez ce fichier et repondez O, ou installez-le vous-meme :
    echo       https://code.claude.com/docs/en/setup
    echo.
    pause
    exit /b 1
  )
  echo   Le poste demarre en mode degrade : ce qui manque restera inaccessible.
  echo.
  goto ouvrir_outlook
)

rem --- Claude Code -----------------------------------------------------------
if not defined MANQUE_CLAUDE goto install_python
echo   Installation de Claude Code...
echo.
set "VOIES="

rem  Voie 1 : installateur natif par curl. Methode officielle en CMD.
rem  On travaille en chemins ABSOLUS : si %TEMP% est injoignable, on ne veut
rem  surtout pas telecharger puis EXECUTER un installateur dans le dossier de
rem  travail synchronise. Et `call` sur un chemin absolu, jamais en relatif nu :
rem  avec NoDefaultCurrentDirectoryInExePath, `call cc_install.cmd` echoue alors
rem  que `if exist` etait vrai. Les deux mesures datent du 2026-08-29.
if defined A_CURL (
  set "VOIES=!VOIES! curl"
  echo   Tentative par curl...
  if exist "%TEMP%\" (
    curl -fsSL https://claude.ai/install.cmd -o "%TEMP%\cc_install.cmd"
    if exist "%TEMP%\cc_install.cmd" call "%TEMP%\cc_install.cmd"
    del "%TEMP%\cc_install.cmd" >nul 2>&1
  ) else (
    echo   Dossier temporaire injoignable - voie ignoree.
  )
)
call :apres_install_claude
if defined CLAUDE_EXE goto claude_ok

rem  Voie 2 : installateur natif par PowerShell, chemin absolu.
rem  PAS de caret devant le pipe : entre guillemets, cmd ne le consomme pas et
rem  PowerShell recoit `^|`, ce qui leve " Jeton inattendu ^ ". Mesure du
rem  2026-08-29 : cette voie n'avait JAMAIS pu reussir, sur aucun poste.
if exist "%PS_EXE%" (
  set "VOIES=!VOIES! PowerShell"
  echo   Tentative par PowerShell...
  "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "irm https://claude.ai/install.ps1 | iex"
)
call :apres_install_claude
if defined CLAUDE_EXE goto claude_ok

rem  Voie 3 : winget. Attention, une installation winget ne se met PAS a jour
rem  toute seule, contrairement a l'installation native.
if defined A_WINGET (
  set "VOIES=!VOIES! winget"
  echo   Tentative par winget...
  winget install --id Anthropic.ClaudeCode --silent --accept-source-agreements --accept-package-agreements
)
call :apres_install_claude
if defined CLAUDE_EXE goto claude_ok

rem  Voie 4 : Node.js puis npm. Dernier recours.
rem  Le paquet npm installe le MEME binaire natif ; Node ne sert qu'a
rem  l'installer, claude ne l'appelle pas a l'execution.
where npm >nul 2>&1
if errorlevel 1 if defined A_WINGET (
  echo   Installation de Node.js...
  winget install --id OpenJS.NodeJS.LTS --silent --accept-source-agreements --accept-package-agreements
  if exist "%ProgramFiles%\nodejs" set "PATH=%ProgramFiles%\nodejs;%PATH%"
)
where npm >nul 2>&1
if not errorlevel 1 (
  set "VOIES=!VOIES! npm"
  echo   Tentative par npm...
  call npm install -g @anthropic-ai/claude-code
)
call :apres_install_claude
if defined CLAUDE_EXE goto claude_ok

echo.
echo   ECHEC : Claude Code n'a pas pu etre installe automatiquement.
rem  Ne PAS pretendre que quatre voies ont ete tentees : winget et npm sont
rem  conditionnels et souvent sautes. Annoncer ce qui a REELLEMENT ete tente
rem  envoie l'utilisateur chercher au bon endroit. Mesure du 2026-08-29.
if defined VOIES (
  echo   Voies reellement tentees :!VOIES!
) else (
  echo   AUCUNE voie n'a pu etre tentee : ni curl, ni PowerShell, ni winget,
  echo   ni npm ne sont disponibles sur ce poste.
)
if not defined A_WINGET echo   winget est absent - c'est souvent la cause.
echo.
echo   Installez Claude Code a la main, puis relancez ce fichier :
echo       https://code.claude.com/docs/en/setup
echo.
pause
exit /b 1

:claude_ok
echo   Claude Code installe : %VER_CLAUDE%
echo   Emplacement : %CLAUDE_EXE%
call :ajouter_au_path "%CLAUDE_BIN%"
echo.

rem --- Python ----------------------------------------------------------------
:install_python
if not defined MANQUE_PYTHON goto install_pywin32
echo   Installation de Python...
echo.
if defined A_WINGET (
  winget install --id Python.Python.3.13 --scope user --silent --accept-source-agreements --accept-package-agreements
) else (
  echo   winget est absent : Python ne peut pas etre installe automatiquement.
)
call :apres_install_python
if not defined PY (
  echo.
  echo   Python n'a pas pu etre installe automatiquement.
  echo   Installez-le depuis www.python.org - cochez bien
  echo   "Add python.exe to PATH" pendant l'installation - puis relancez.
  echo.
  echo   On continue : les dossiers de projets fonctionnent sans Python.
  echo.
  set "MODE_DEGRADE=1"
  goto install_git
)
echo   Python installe : %PY%
echo.

rem --- pywin32 ---------------------------------------------------------------
:install_pywin32
if not defined MANQUE_PYWIN32 goto install_git
if not defined PY (
  echo   pywin32 : saute, Python est absent.
  echo.
  goto install_git
)
echo   Installation de pywin32...
echo   Une seule fois, environ trente secondes.
echo.
call "%PY%" -m pip install --quiet --disable-pip-version-check pywin32
call "%PY%" -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
  echo.
  echo   L'installation de pywin32 a echoue.
  echo   Ouvrez une invite de commandes et lancez :
  echo.
  echo       "%PY%" -m pip install pywin32
  echo.
  echo   Cause frequente : pas de connexion, ou un pare-feu d'entreprise.
  echo.
  pause
  set "MODE_DEGRADE=1"
) else (
  echo   pywin32 installe.
)
echo.

rem --- Git pour Windows ------------------------------------------------------
:install_git
if not defined MANQUE_GIT goto ouvrir_outlook
echo   Installation de Git pour Windows...
echo.
if defined A_WINGET (
  winget install --id Git.Git --silent --accept-source-agreements --accept-package-agreements
) else (
  echo   winget est absent : Git ne peut pas etre installe automatiquement.
)
call :apres_install_git
if not defined GIT_EXE (
  echo   Git n'a pas pu etre installe. Ce n'est pas bloquant.
  echo   Claude Code utilisera PowerShell au lieu de Bash.
  echo   Pour l'ajouter plus tard : https://git-scm.com/downloads/win
) else (
  echo   Git pour Windows installe : %GIT_EXE%
)
echo.

rem ===========================================================================
rem  [3/6] Outlook. Il se charge pendant tout le reste.
rem ===========================================================================
:ouvrir_outlook
echo   [3/6] Ouverture d'Outlook...
echo.

if defined MOTEUR_MANQUANT (
  echo   MOTEUR INCOMPLET :%MOTEUR_MANQUANT%
  echo.
  echo   Des scripts du dossier %CFG%\scripts\ manquent. Les courriels et le
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

if not defined OUTLOOK_EXE (
  echo   Outlook classique est introuvable aux emplacements habituels.
  echo   Le nouveau Outlook, celui du Microsoft Store, n'expose pas
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

rem ===========================================================================
rem  [4/6] Mise a jour de Claude Code pendant qu'Outlook demarre.
rem        L'installation native se met deja a jour toute seule en arriere-plan ;
rem        cet appel ne fait que la declencher tout de suite.
rem ===========================================================================
:maj_claude
echo.
echo   [4/6] Mise a jour de Claude Code...
echo.

if not defined CLAUDE_EXE (
  echo   Claude Code reste introuvable : il n'y a pas de session a demarrer.
  echo   Relancez ce fichier et acceptez l'installation, ou installez-le
  echo   vous-meme : https://code.claude.com/docs/en/setup
  echo.
  pause
  exit /b 1
)

call "%CLAUDE_EXE%" update
if errorlevel 1 (
  echo.
  echo   Mise a jour impossible - on continue quand meme.
  echo   Cause frequente : une autre session Claude est ouverte,
  echo   Windows verrouille alors le fichier en cours d'utilisation.
)

rem ===========================================================================
rem  [5/6] Le processus ne suffit pas : c'est MAPI qui doit repondre.
rem        check_setup.py sort 0 quand la boite est reellement accessible.
rem ===========================================================================
echo.
echo   [5/6] Attente d'Outlook...
echo.

if defined MODE_DEGRADE (
  echo   Ignoree : mode degrade, sans courriels ni calendrier.
  goto lancer_claude
)
if not defined PY (
  echo   Ignoree : Python absent, la boite ne peut pas etre testee.
  goto lancer_claude
)

set /a TENTATIVE=0

:attente
set /a TENTATIVE+=1
if %TENTATIVE% GTR 20 goto attente_expiree
<nul set /p "=."
rem  ping plutot que timeout : timeout refuse de dormir quand l'entree est
rem  redirigee - raccourci, tache planifiee. ping dort ~2s partout.
ping -n 3 127.0.0.1 >nul
"%TASKLIST_EXE%" /FI "IMAGENAME eq OUTLOOK.EXE" | "%FIND_EXE%" /I "OUTLOOK.EXE" >nul
if errorlevel 1 goto attente
call "%PY%" "%CFG%\scripts\check_setup.py" >nul 2>&1
if errorlevel 1 goto attente

:outlook_pret
echo.
echo   Outlook repond. Boite accessible.
goto lancer_claude

:attente_expiree
echo.
echo   Outlook met plus de temps que prevu a repondre. Diagnostic :
echo.
call "%PY%" "%CFG%\scripts\check_setup.py"
echo.
echo   On demarre Claude quand meme - demandez-lui de regarder.

rem ===========================================================================
rem  [6/6] La session. Une seule, quatre metiers.
rem ===========================================================================
:lancer_claude
echo.
echo   [6/6] Demarrage : Opus 5 contexte 1M, effort MAX, autonomie complete.
echo.

rem  --fallback-model : un abonnement Pro n'a pas forcement acces a Opus 1M.
rem  Sans repli, la session est refusee net. Ajoute le 2026-08-29.
call "%CLAUDE_EXE%" --dangerously-skip-permissions --model "claude-opus-5[1m]" --fallback-model "claude-sonnet-5" --effort max
set "CODE_SESSION=%errorlevel%"

rem  Sans ce pause, un echec immediat de `claude` - plan sans acces au modele,
rem  session non authentifiee, abonnement absent - ferme la fenetre sur son
rem  propre message d'erreur : rien a lire, rien a rapporter.
if not "%CODE_SESSION%"=="0" (
  echo.
  echo   La session s'est terminee avec le code %CODE_SESSION%.
  echo   Si un message d'erreur s'affiche ci-dessus, c'est lui qui compte.
  echo.
  pause
)

endlocal & exit /b %CODE_SESSION%

rem ===========================================================================
rem  PATH UTILISATEUR
rem
rem  L'installateur natif le signale lui-meme : "Native installation exists but
rem  ...\.local\bin is not in your PATH". Ce .bat s'en sort en le prefixant
rem  pour SA session, mais l'utilisateur qui ouvre ensuite un terminal et tape
rem  `claude` ne trouve rien.
rem
rem  On l'ajoute PAR LE REGISTRE, jamais par `setx` : setx tronque a 1024
rem  caracteres et detruit silencieusement un PATH long. Ajoute le 2026-08-29.
rem ===========================================================================

:ajouter_au_path
if not exist "%~f1" goto :eof
rem  Pas de `echo %PATH% ^| findstr` ici : le PATH contient des parentheses
rem  (Program Files (x86)) qui cassent la ligne. PowerShell verifie lui-meme si
rem  le chemin y est deja, et ne fait rien si c'est le cas.
"%PS_EXE%" -NoProfile -Command "$p = '%~f1'; $u = [Environment]::GetEnvironmentVariable('PATH','User'); if ($null -eq $u) { $u = '' }; if ($u -notlike ('*' + $p + '*')) { [Environment]::SetEnvironmentVariable('PATH', ($u.TrimEnd(';') + ';' + $p).TrimStart(';'), 'User') }" >nul 2>&1
if errorlevel 1 (
  echo   Ajout au PATH impossible - sans effet sur cette session, qui fonctionne.
  echo   Pour taper `claude` depuis n'importe quel terminal, ajoutez a la main :
  echo       %~f1
) else (
  echo   Ajoute au PATH de votre compte : %~f1
  echo   ^(actif dans les NOUVEAUX terminaux ; celui-ci fonctionne deja^)
)
goto :eof

rem ===========================================================================
rem  SOUS-ROUTINES DE RESOLUTION
rem
rem  POURQUOI un chemin absolu plutot que la commande nue.
rem  `where` regarde le DOSSIER COURANT avant le PATH. Or le dossier courant de
rem  ce script est le dossier de travail : un espace OneDrive/SharePoint
rem  synchronise ou n'importe quel fichier peut atterrir, y compris depose par
rem  un tiers. Une commande nue y prendrait donc un executable etranger. On
rem  ecarte explicitement toute correspondance dans le dossier courant, et on ne
rem  se sert ensuite que du chemin absolu retenu.
rem
rem  POURQUOI on exige une REPONSE.
rem  Un executable present n'est pas un executable qui marche. Mesure du
rem  2026-08-29 : %LOCALAPPDATA%\Microsoft\WindowsApps\python.exe fait 0 octet -
rem  c'est l'alias du Microsoft Store - et `where python` le renvoie. Il etait
rem  retenu, annonce " present ", et le poste partait avec un outillage
rem  inexistant sans meme poser la question. On exige donc que `--version`
rem  rende une chaine NON VIDE.
rem
rem  Ces chemins ne sont JAMAIS codes en dur : ils sont recalcules a chaque
rem  execution, sur chaque poste, a partir des variables d'environnement.
rem ===========================================================================

:resoudre_claude
set "CLAUDE_EXE="
set "VER_CLAUDE="
if exist "%CLAUDE_BIN%\claude.exe" call :essai_claude "%CLAUDE_BIN%\claude.exe"
if defined CLAUDE_EXE goto :eof
for /f "delims=" %%E in ('where claude 2^>nul') do (
  if not defined CLAUDE_EXE call :essai_claude "%%~fE"
)
goto :eof

:essai_claude
if /I "%~dp1"=="%CDS%" goto :eof
rem  La sortie doit RESSEMBLER a une version (chiffre.chiffre), pas seulement
rem  etre non vide : un stub bavard repondrait n'importe quoi.
rem  La validation se fait HORS du for /f. Mesure du 2026-08-29 : un tube vers
rem  findstr a l'interieur du for /f produit deux arguments entre guillemets,
rem  cmd retire alors le premier et le dernier, et la detection echoue - Claude
rem  Code etait annonce ABSENT alors qu'il etait installe.
set "VTEMP="
for /f "delims=" %%V in ('"%~f1" --version 2^>nul') do if not defined VTEMP set "VTEMP=%%V"
if not defined VTEMP goto :eof
echo %VTEMP%| "%SystemRoot%\System32\findstr.exe" /R "[0-9]\.[0-9]" >nul
if errorlevel 1 goto :eof
set "VER_CLAUDE=%VTEMP%"
set "CLAUDE_EXE=%~f1"
goto :eof

:apres_install_claude
rem  Apres une installation, le PATH de CETTE fenetre est encore l'ancien.
rem  On rajoute l'emplacement natif avant de re-resoudre, sinon on conclurait
rem  a tort que l'installation a echoue.
if exist "%CLAUDE_BIN%" set "PATH=%CLAUDE_BIN%;%PATH%"
call :resoudre_claude
goto :eof

:resoudre_python
set "PY="
set "VER_PY="
for /f "delims=" %%E in ('where python 2^>nul') do (
  if not defined PY call :essai_python "%%~fE"
)
if defined PY goto :eof
for /f "delims=" %%E in ('where py 2^>nul') do (
  if not defined PY call :essai_python "%%~fE"
)
goto :eof

:essai_python
if /I "%~dp1"=="%CDS%" goto :eof
rem  ON FAIT EXECUTER DU PYTHON, on ne se contente pas d'une reponse.
rem  Corrige le 2026-08-29 apres essai sur un poste Windows 11 neuf : le stub du
rem  Microsoft Store (%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe, 0 octet)
rem  est BAVARD. Il rend une vraie version quand un Python du Store est installe
rem  derriere, et "Python est introuvable..." quand il n'y en a pas. Les deux
rem  sorties etant non vides, exiger une reponse laissait passer un faux
rem  interpreteur : le poste n'installait jamais le vrai Python, et
rem  `pip install pywin32` echouait ensuite sans que la cause soit visible.
rem  On passe par un FICHIER SONDE, pas par `for /f ('commande')`.
rem  Mesure du 2026-08-29 : `for /f ... in ('"chemin" -c "code"')` casse, parce
rem  que cmd retire le premier et le dernier guillemet des qu'il y a DEUX
rem  arguments entre guillemets - "La syntaxe du nom de fichier est
rem  incorrecte". Le vrai Python etait alors rejete lui aussi.
set "PYOK="
set "PYSONDE=%TEMP%\cc_pysonde.txt"
del "%PYSONDE%" >nul 2>&1
rem  `call` obligatoire : si le candidat est un .bat ou un .cmd - shim
rem  pyenv-win, pixi, conda - une invocation nue transfere le controle et
rem  ne revient JAMAIS. Le script mourait en silence pendant l'inventaire.
call "%~f1" -c "print(84)" > "%PYSONDE%" 2>nul
if exist "%PYSONDE%" for /f "usebackq delims=" %%V in ("%PYSONDE%") do if "%%V"=="84" set "PYOK=1"
del "%PYSONDE%" >nul 2>&1
if not defined PYOK goto :eof
for /f "delims=" %%V in ('"%~f1" --version 2^>^&1') do if not defined VER_PY set "VER_PY=%%V"
if not defined VER_PY goto :eof
set "PY=%~f1"
goto :eof

:apres_install_python
rem  winget --scope user ecrit dans le PATH UTILISATEUR : une fenetre cmd deja
rem  ouverte ne le verra jamais. Sans ce rattrapage, le script annoncait
rem  " Python n'a pas pu etre installe " APRES une installation reussie.
rem  Mesure du 2026-08-29.
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
  if exist "%%~fD\python.exe" set "PATH=%%~fD;%%~fD\Scripts;!PATH!"
)
if exist "%LOCALAPPDATA%\Programs\Python\Launcher" set "PATH=%LOCALAPPDATA%\Programs\Python\Launcher;!PATH!"
call :resoudre_python
goto :eof

:resoudre_git
set "GIT_EXE="
for /f "delims=" %%E in ('where git 2^>nul') do (
  if not defined GIT_EXE if /I not "%%~dpE"=="%CDS%" set "GIT_EXE=%%~fE"
)
goto :eof

:apres_install_git
rem  Meme rattrapage que pour Python : winget installe Git sans que la fenetre
rem  courante voie le nouveau PATH.
if exist "%ProgramFiles%\Git\cmd\git.exe" set "PATH=%ProgramFiles%\Git\cmd;%PATH%"
if exist "%LOCALAPPDATA%\Programs\Git\cmd\git.exe" set "PATH=%LOCALAPPDATA%\Programs\Git\cmd;%PATH%"
call :resoudre_git
goto :eof
