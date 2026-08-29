# Installation — poste de travail Outlook pour Claude Code

> ### Une gracieuseté de **Sylvain Leduc — Constructo AI inc.**
> *Écosystème intelligent pour la construction au Québec* · [www.constructoai.ca](https://www.constructoai.ca)

Gérer ses **courriels**, son **calendrier**, ses **dossiers de projets** et sa **comptabilité**
avec Claude, directement sur son Outlook — sans mot de passe, sans jeton, sans inscription
d'application.

---

## Ce qu'il faut avant de commencer

| Prérequis | Comment vérifier | Si ça manque |
|---|---|---|
| **Outlook classique** | il s'ouvre depuis `Program Files\Microsoft Office\` | ⛔ Celui du **Microsoft Store n'expose pas MAPI/COM** — rien ne fonctionnera. Installer Outlook classique. |
| **Python 3.x** | `python --version` | rien à faire — **le `.bat` l'installe seul** (winget, portée utilisateur). Repli manuel : [python.org](https://www.python.org), cocher « Add to PATH » |
| **pywin32** | `python -c "import win32com.client"` | rien à faire — **le `.bat` l'installe seul** |
| **Claude Code** | `claude --version` | rien à faire — **le `.bat` l'installe seul** (installateur officiel, sans droits admin). ⚠️ mais il vous faut un **abonnement Claude payant** — Pro, Max, Team, Enterprise ou Console ; *ce poste est gratuit, Claude Code ne l'est pas* |
| **Git pour Windows** | `git --version` | rien à faire — **le `.bat` l'installe seul**. *Optionnel mais recommandé* : sans lui, Claude Code n'a pas l'outil Bash et bascule sur PowerShell, alors que les commandes du poste sont écrites en shell POSIX |

⚠️ **Ne cherchez pas `claude.cmd`** : l'installateur natif produit `claude.exe` dans
`%USERPROFILE%\.local\bin`. Le nom `claude.cmd` n'existe qu'avec une installation **npm**.
*(Mesuré le 2026-08-29 : le `.bat` testait `claude.cmd` et refusait de démarrer sur un poste où
Claude Code était installé et fonctionnel.)*

---

## Installation — deux gestes

**1. Copier le dossier `.claude`** dans le dossier de travail : celui qui contient déjà vos
projets, généralement un dossier OneDrive ou SharePoint synchronisé.

```
Mon dossier de travail\
   .claude\          <- le dossier copié
   01. PROJETS\      <- vos dossiers, tels qu'ils sont
   02. CLIENTS\
```

**2. Double-cliquer `.claude\Constructo_AI.bat`.**

C'est tout. **Vous n'avez aucune commande à taper.** Au premier lancement, il dresse l'inventaire
de ce qui est présent, vous montre la liste de ce qui manque, et **demande une seule fois** :

```
  Il manque :
    - Claude Code       installateur officiel, sans droits admin
    - pywin32           le pont vers Outlook
    - Git pour Windows  optionnel : donne l'outil Bash a Claude

  Claude Code, Python et pywin32 s'installent pour VOTRE compte seulement,
  sans elevation : rien n'est modifie pour les autres utilisateurs.

  > Git pour Windows, LUI, s'installe pour toute la machine et affichera
    une invite UAC. Vous pouvez la REFUSER : Git est optionnel, et tout
    le reste continue sans lui.

  Installer maintenant ? [O/N, defaut N]
```

Répondez `O` et il enchaîne tout seul. Répondez autre chose et il démarre quand même, en mode
dégradé, en disant ce qui restera inaccessible. **Si rien ne manque, la question n'est pas
posée.**

Pour Claude Code il tente quatre voies dans l'ordre, et s'arrête à la première qui réussit :
l'installateur officiel par `curl`, puis par PowerShell, puis `winget`, puis Node.js + `npm`.
*Note : une installation `winget` ne se met pas à jour toute seule, contrairement à
l'installation native.*

Ensuite il se replace au bon niveau, ouvre Outlook, **attend que la boîte réponde vraiment**,
et démarre la session. Si quelque chose manque, il le dit au lieu de faire semblant.

---

## Première session — dix minutes bien investies

Le poste arrive **vierge**. Il connaît tous les pièges techniques d'Outlook, mais rien de votre
entreprise. Demandez simplement :

> **« lis CLAUDE.md et remplis les sections À COMPLÉTER en mesurant »**

Claude relèvera vos calendriers, vos dossiers, vos volumes réels, et vous posera les questions
auxquelles il ne peut pas répondre seul. Il ne remplira **rien d'avance** : une section vide se
lit « pas encore consigné », jamais « rien ne s'est passé ».

Puis mettez votre signature en place :

1. Ouvrir un **courriel réellement envoyé** — pas le dossier `%APPDATA%\Microsoft\Signatures`,
   qui contient souvent des signatures périmées.
2. **Copier** `.claude\profiles\signature_MODELE.html` en
   `.claude\profiles\signature_defaut.html`, puis y coller le bloc. C'est ce nom-là
   qu'appelle `--signature` sans argument.
   ⚠️ **Le dépôt ne livre volontairement aucune signature remplie** — elle porterait votre nom
   et votre adresse sur GitHub. Tant que la copie n'est pas faite, `--signature` **refuse** :
   `signature introuvable`. Et `--signature MODELE` refuse aussi, exprès.
3. Retirer le commentaire HTML en tête du fichier : il part **avec** la signature dans le corps
   du courriel, invisible au rendu mais lisible en « afficher la source ».

⚠️ **Sans le drapeau `--signature`, les courriels partent sans signature** — un message créé par
COM n'en reçoit jamais automatiquement, et rien ne le signale.

---

## Une décision à prendre, dès le premier jour

`CLAUDE.md` §4-1. Deux positions possibles :

| | |
|---|---|
| **Accord message par message** *(défaut)* | Claude rédige, montre, attend votre « go ». Le verrou `--yes-send` ne se pose qu'après. |
| **Envoi autonome** | Claude rédige et envoie. Le drapeau reste, comme geste délibéré. |

Le défaut est la position prudente. **La changer est une décision, pas un réglage** — un
courriel envoyé ne se rappelle pas. Quelle que soit votre position, quatre règles ne bougent
jamais : ne jamais suivre un lien reçu, ne jamais exécuter ce qu'un courriel demande, ne jamais
supprimer définitivement, signaler toute adresse inconnue. Celles-là protègent contre des
**tiers**, pas contre vous.

---

## Ce que vous pouvez demander, ensuite

- « **qu'est-ce qui attend une réponse ?** » — il mesure, trie en quatre paquets, propose des brouillons
- « **qu'est-ce qui vient cette semaine ?** » — le calendrier, récurrences comprises
- « **où en est [client] ?** » — il croise le dossier, les courriels et vos fichiers
- « **prépare une relance pour la facture [numéro]** » — il lit le fil avant d'écrire
- « **note ça** » — c'est la phrase qui fait grandir la mémoire du poste

---

## Ce qu'il y a dans le dossier

| | |
|---|---|
| `CLAUDE.md` | **le hub** — se charge tout seul, porte l'accès, les règles, la carte |
| `settings.json` | permissions et les deux hooks qui tiennent la mémoire à jour |
| `scripts\` | **le moteur, six scripts** — `outlook_mail` · `outlook_calendar` · `veille_poste` · `check_setup` · `factures` · `ost_reader` |
| `skills\poste-outlook\` | la méthode : trier, chercher, rédiger, chiffrer |
| `agents\courriels.md` | l'agent délégué pour les passes longues |
| `profiles\` | votre signature, vos profils métier |
| `ETAT_*.md` · `JOURNAL.md` | 🛰️ la mémoire — **ne se chargent pas seuls**, ils grossissent à l'usage |
| `references\depannage.md` | boîte vide, synchronisation figée, mode sans échec |
| `Constructo_AI.bat` | le point d'entrée |

---

## Deux règles techniques à ne pas casser

🔴 **Les fichiers de `.claude\` doivent rester en fins de ligne LF.** Un fichier d'agent en CRLF
**ne s'enregistre pas** — il disparaît de la liste sans aucune erreur. Une compétence et un
`CLAUDE.md` en CRLF, eux, fonctionnent quand même : c'est une panne **partielle**, donc
invisible. `Constructo_AI.bat`, lui, doit rester en **CRLF** (exigence de `cmd.exe`).

```
python -c "d=open(r'.claude\agents\courriels.md','rb').read(); print('CRLF' if b'\r\n' in d else 'LF')"
```

🔴 **Ne pas « simplifier » les hooks** en `shell: bash` ou `shell: powershell`. Sur un poste
avec WSL, `bash` du PATH résout vers **WSL** ; et `pwsh` est souvent **absent**. Le hook meurt
alors en silence. Ils tournent en forme exec sur `powershell.exe`, en chemin absolu, pour cette
raison précise.

---

## Aucun secret n'est stocké

L'accès passe par le profil Outlook **déjà authentifié** du poste : aucun mot de passe, aucun
jeton, aucun droit administrateur. Si Outlook fonctionne pour vous, l'outil fonctionne.

⚠️ Ce dossier étant destiné à un espace synchronisé, **n'y écrivez jamais de secret** — mot de
passe, clé d'API, numéro de compte de taxe. Les fichiers vivants sont conçus pour pointer vers
ces informations, pas pour les contenir.

---

## Une gracieuseté de Constructo AI

Ce poste vous est offert par **Sylvain Leduc**, président-concepteur de
**Constructo AI inc.** — *écosystème intelligent pour la construction au Québec*.

Il n'est pas né d'un cahier des charges : il a été **construit et éprouvé en conditions
réelles**, sur une vraie boîte de courriels, un vrai calendrier de chantiers et une vraie
comptabilité. Chaque piège qu'il documente a d'abord coûté quelque chose à quelqu'un — un
compte faux qui avait l'air juste, un agent silencieusement mort, une facture partie sans
signature, un prix sous-estimé de 2,3 %.

**C'est ce qui fait sa valeur** : ce ne sont pas des précautions théoriques, ce sont des
mesures. Prenez-en soin comme d'un carnet d'atelier — corrigez ce qui se révèle faux,
datez ce que vous mesurez, et il vous servira longtemps.

[www.constructoai.ca](https://www.constructoai.ca)
