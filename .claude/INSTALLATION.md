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
| **Python 3.x** | `python --version` | [python.org](https://www.python.org) — cocher « Add to PATH » |
| **pywin32** | `python -c "import win32com.client"` | `pip install pywin32` |
| **Claude Code** | `claude --version` | [claude.com/claude-code](https://claude.com/claude-code) |

---

## Installation — trois minutes

**1. Copier le dossier `.claude`** dans le dossier de travail : celui qui contient déjà vos
projets, généralement un dossier OneDrive ou SharePoint synchronisé.

```
Mon dossier de travail\
   .claude\          <- le dossier copié
   01. PROJETS\      <- vos dossiers, tels qu'ils sont
   02. CLIENTS\
```

**2. Installer les dépendances Python :**

```
pip install -r ".claude\requirements.txt"
```

**3. Double-cliquer `.claude\Poste.bat`.**

Il se replace tout seul au bon niveau, ouvre Outlook, **attend que la boîte réponde vraiment**,
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
2. Coller le bloc dans `.claude\profiles\signature_defaut.html` — il est déjà en place et
   c'est celui qu'appelle `--signature` sans argument. (`signature_MODELE.html` reste comme
   exemple commenté.)

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
| `scripts\` | **le moteur** — `outlook_mail` · `outlook_calendar` · `veille_poste` · `check_setup` |
| `skills\poste-outlook\` | la méthode : trier, chercher, rédiger, chiffrer |
| `agents\courriels.md` | l'agent délégué pour les passes longues |
| `profiles\` | votre signature, vos profils métier |
| `ETAT_*.md` · `JOURNAL.md` | 🛰️ la mémoire — **ne se chargent pas seuls**, ils grossissent à l'usage |
| `references\depannage.md` | boîte vide, synchronisation figée, mode sans échec |
| `Poste.bat` | le point d'entrée |

---

## Deux règles techniques à ne pas casser

🔴 **Les fichiers de `.claude\` doivent rester en fins de ligne LF.** Un fichier d'agent en CRLF
**ne s'enregistre pas** — il disparaît de la liste sans aucune erreur. Une compétence et un
`CLAUDE.md` en CRLF, eux, fonctionnent quand même : c'est une panne **partielle**, donc
invisible. `Poste.bat`, lui, doit rester en **CRLF** (exigence de `cmd.exe`).

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
