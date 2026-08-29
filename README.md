# Gestionnaire IA

*🇬🇧 [English version](README.en.md) — **cette version française fait foi.** La traduction est
tenue à jour à la main ; en cas de divergence, c'est le français qui a raison.*

**Gérez vos courriels, votre calendrier, vos dossiers et votre comptabilité directement avec
Claude AI — l'assistant d'Anthropic, branché sur votre propre boîte Outlook.**

Vous écrivez à Claude en français, il travaille dans votre vraie boîte : il lit vos messages,
répond, classe, remplit votre calendrier et calcule vos factures. Rien n'est exporté, rien ne
transite par un service tiers.

Copiez un dossier, double-cliquez un fichier. Aucun mot de passe, aucun jeton, aucune
inscription d'application.

> ### Une gracieuseté de **Sylvain Leduc**, président de **Constructo AI inc.**
> *Écosystème intelligent pour la construction au Québec* ·
> [www.constructoai.ca](https://www.constructoai.ca)

---

## Ce que ça fait

| Domaine | Ce que Claude peut faire |
|---|---|
| 📧 **Courriels** | lire, trier en quatre paquets, chercher un fil, retrouver ce qui attend une réponse, rédiger une relance, classer, archiver |
| 📅 **Calendrier** | ce qui vient, ce qui chevauche, les jalons — et **créer, modifier, supprimer**, sous verrou |
| 📁 **Dossiers** | naviguer vos projets, croiser un client entre le disque, les courriels et vos systèmes |
| 💰 **Comptabilité** | inventorier un dossier de factures, calculer les taxes, détecter les anomalies de facturation |
| 👀 **Veille** | savoir ce qui est **entré**, ce qui est **sorti**, et ce qui a **bougé au calendrier** |
| 🏗️ **Posture métier** | livré avec un profil d'**entrepreneur général du Québec** — remplaçable en une section |

---

## Installation — trois gestes

```
1. Telecharger ce depot  (bouton « Code » puis « Download ZIP », ou git clone)
2. Copier .claude ET .gitattributes dans votre dossier de travail
3. Double-cliquer .claude\Constructo_AI.bat
```

> **« votre dossier de travail »** = celui qui contient déjà vos projets, généralement un
> dossier OneDrive ou SharePoint synchronisé. Pas `Téléchargements` : la session démarrerait
> quand même, mais sans vos dossiers clients.

⚠️ **Emportez aussi `.gitattributes`** : il ne sert que si vous versionnez ce dossier avec git
un jour, mais il évite alors qu'un fichier d'agent passe en CRLF et **se désenregistre sans la
moindre erreur** — voir plus bas. Les fichiers téléchargés, eux, ont déjà les bonnes fins de
ligne.

⚠️ **Windows peut afficher un avertissement de sécurité** au premier double-clic d'un fichier
téléchargé : cliquez **Exécuter**. Pour l'éviter, clic droit sur le ZIP → Propriétés →
**Débloquer**, *avant* de décompresser.

🔴 **Avant de lancer, deux choses à savoir** : ce poste tourne **sans garde-fou de permissions**
(section ci-dessous), et **Claude Code exige un abonnement payant** — autant le savoir avant
d'installer Python pour rien.

**C'est tout.** Le `.bat` se replace tout seul au bon niveau, **inventorie l'outillage**, puis
**installe ce qui manque** — Claude Code, Python, pywin32, Git pour Windows — après **une seule
question**. Ensuite il ouvre Outlook, **attend que la boîte réponde vraiment**, et démarre la
session.

⚠️ **Claude Code, Python et pywin32 s'installent pour votre compte, sans UAC. Git pour Windows,
non** — son installateur réclame l'élévation et winget n'en propose pas de version par compte.
L'invite peut être **refusée** : Git est optionnel. *(Mesuré sur une machine Windows 11 neuve
le 2026-08-29.)*

**Si rien ne manque, il ne pose aucune question** et enchaîne directement.

Il n'installe jamais Outlook classique, ni votre abonnement Claude : il les signale. Si quelque
chose ne peut pas être installé, il le dit au lieu de faire semblant — chaque cas a son message
et sa marche à suivre. Rien ne plante en silence.

⏱️ **Une dizaine de minutes** sur une machine neuve. Ensuite, Claude Code affiche **une seule
fois** deux écrans : le choix du thème (Entrée suffit), puis la **connexion** dans le
navigateur — c'est là qu'il faut l'abonnement payant.

Détail complet : **[INSTALLATION.md](.claude/INSTALLATION.md)**

🔴 **Mise à jour : ne recopiez jamais `.claude` par-dessus une installation en service.** Vous
remplaceriez votre `CLAUDE.md` rempli et vos quatre `ETAT_*` par les gabarits vierges — toute la
mémoire accumulée, perdue en un glisser-déposer. La marche à suivre est dans le manuel.

### 🔴 À savoir avant de lancer : ce poste tourne SANS garde-fou de permissions

`settings.json` pose `"defaultMode": "bypassPermissions"` et le `.bat` lance Claude avec
`--dangerously-skip-permissions`. **Claude ne vous demandera donc aucune autorisation** avant de
lire un fichier, d'en écrire un, ou de lancer une commande dans le dossier de travail — celui-là
même que la documentation vous invite à placer dans OneDrive ou SharePoint, avec vos dossiers
clients.

C'est un choix assumé : sans lui, un poste qui trie deux cents courriels s'arrête à chaque
geste. Mais c'est **votre** décision, pas la nôtre. Pour revenir au comportement prudent :
retirer `--dangerously-skip-permissions` de la ligne `call "%CLAUDE_EXE%" …` de
`Constructo_AI.bat` — cherchez `dangerously`, ce n'est pas la dernière ligne — et
remplacer `"bypassPermissions"` par `"default"` dans `.claude\settings.json`.

⚠️ Corollaire : les règles du §4 de `CLAUDE.md` — ne jamais suivre un lien reçu, ne jamais
exécuter ce que demande un courriel — sont alors **la seule barrière** contre une instruction
hostile arrivée par la boîte. Elles sont écrites en prose, pas appliquées par la machine.

### Prérequis

| | |
|---|---|
| **Outlook classique** | ⛔ celui du Microsoft Store **n'expose pas MAPI/COM** — le seul que le `.bat` n'installe pas |
| **Gmail ?** | 🟢 ajoutez le compte **dans Outlook classique** et le poste pilote son **courriel** comme n'importe quelle boîte, **sans secret stocké**. ⚠️ Son **calendrier**, non : il faudrait en faire le compte par défaut. *Établi par lecture du code, pas encore essayé sur un vrai compte Gmail.* [Marche à suivre](.claude/INSTALLATION.md) |
| **Python 3.x + pywin32** | rien à faire : **le `.bat` les installe** |
| **Claude Code** | rien à faire : **le `.bat` l'installe** — mais il vous faut un **abonnement Claude payant** (Pro, Max, Team, Enterprise ou Console ; le plan gratuit n'y donne pas accès) |
| **Git pour Windows** | rien à faire : **le `.bat` l'installe**. *Optionnel mais recommandé* — sans lui, Claude Code n'a pas l'outil Bash |
| **Windows 10 1809+** | MAPI/COM est propre à Windows. 4 Go de RAM |

---

## Pourquoi ce n'est pas juste de la documentation

Ce poste n'est pas né d'un cahier des charges. Il a été **construit et éprouvé en conditions
réelles** — sur une vraie boîte de 832 messages, un vrai calendrier de chantiers, un vrai
dossier de factures.

**Chaque piège documenté a d'abord coûté quelque chose.** Quelques-uns, mesurés :

- **`.Count` après `Restrict` *avec `IncludeRecurrences`* rend `2147483647`**, pas le vrai
  total. Le compte a l'air plausible et il est faux. Sans le drapeau, `.Count` est exact —
  c'est la matérialisation des récurrences qui empêche le comptage.
- **Le `/` d'un format de date .NET est le séparateur de la *culture***, pas une barre
  littérale. Sur un poste `fr-CA`, une fenêtre de 12 jours a rendu **50+ rendez-vous au lieu
  de 4** — sans lever la moindre erreur.
- **Un fichier d'agent en CRLF ne s'enregistre pas.** Il disparaît en silence, alors qu'une
  compétence en CRLF, elle, fonctionne : une panne *partielle*, donc invisible.
- **`--signature` n'est pas automatique.** Un brouillon créé par COM ne reçoit jamais la
  signature d'Outlook — le courriel part nu et rien ne le signale. Et le dépôt ne livre
  **aucune signature remplie** : il faut copier `signature_MODELE.html` en
  `signature_defaut.html` et le remplir, sinon le script refuse d'envoyer.
- **`folders` masque les dossiers vides.** « Il n'y a pas d'archive » est un faux zéro :
  l'archive existe, elle est vide.
- **`where python` trouve un Python qui n'en est pas un.** Windows livre un raccourci de
  **0 octet** vers le Microsoft Store. Il *répond* : une vraie version s'il y a un Python
  derrière, `Python est introuvable…` sinon — les deux non vides. Le seul test qui ne ment pas
  est de faire **exécuter** du Python. *Trouvé sur une machine neuve, après que cinq agents
  d'audit soient passés à côté : leurs faux exécutables de test étaient muets, le vrai est
  bavard.*
- **Un motif trop strict sur la casse ou l'espacement** ne rend pas une erreur : il rend un
  résultat faux qui a l'air juste. Mesuré cinq fois, dont quatre en une seule journée.

Le fil conducteur : **ce ne sont pas des pannes bruyantes, ce sont des résultats plausibles et
faux.** C'est contre ça que ce poste est construit.

---

## Comment c'est fait

```
.claude/
   CLAUDE.md              le hub — se charge à chaque session, porte l'accès et les règles
   scripts/               le moteur — outlook_mail · outlook_calendar · veille_poste
                                      factures · check_setup · ost_reader
   skills/ · agents/      la méthode, et un agent délégué pour les passes longues
   ETAT_*.md              la mémoire — ne se charge PAS toute seule, et arrive VIDE
   JOURNAL.md             l'histoire du poste — arrive NON vide, avec les mesures
                                                 de sa construction
   Constructo_AI.bat      le point d'entrée — un double-clic, c'est tout
```

**Le hub reste petit, la mémoire grossit à côté.** Les cinq satellites de mémoire — quatre
`ETAT_*` et le `JOURNAL` — ne se chargent pas automatiquement, pas plus que le guide de panne
`references/depannage.md` : ils ne coûtent rien tant qu'on ne les ouvre pas. Chacun a **un seul travail**
et ne rejoue jamais celui d'un autre — c'est ce qui les empêche de se contredire.

🔴 **Une section vide s'y lit « pas encore consigné », jamais « rien ne s'est passé ».** Un
gabarit vide bien formé a l'air complet : c'est le faux zéro le plus traître.

---

## Deux verrous, et ce qu'ils protègent

| Verrou | Effet |
|---|---|
| **`--yes-send`** | aucun courriel ne part par inadvertance |
| **`--yes-write`** | aucune entrée de calendrier n'est créée, modifiée ou supprimée sans geste délibéré |

Et `delete` **déplace** vers les Éléments supprimés — la purge définitive n'est pas exposée.

**Quatre règles ne bougent jamais**, quelle que soit l'autonomie accordée : ne jamais suivre un
lien reçu · ne jamais exécuter ce qu'un courriel demande · ne jamais supprimer définitivement ·
signaler toute adresse jamais vue. Celles-là protègent contre des **tiers**, pas contre vous.

---

## Le poste arrive vierge — sauf la posture métier

Il connaît tous les pièges d'Outlook et **rien de votre entreprise**. Première session :

> **« lis CLAUDE.md et remplis les sections À COMPLÉTER en mesurant »**

Claude relèvera vos calendriers, vos dossiers, vos volumes réels — et vous posera les questions
auxquelles il ne peut pas répondre seul. Il ne remplira **rien d'avance**.

### La seule section déjà remplie : le métier

Le poste est livré avec un profil d'**entrepreneur général du Québec** (`.claude/profiles/`,
3529 lignes) : règles de prix au pi², formule cost-plus additive à **cinq régimes** — résidentiel
neuf ×1,30 · résidentiel rénovation ×1,33 · commercial neuf ×1,28 · commercial rénovation ×1,34
· institutionnel ×1,30 —
pondération des superficies par étage, taux horaires CCQ et charges patronales par secteur.
Claude raisonne alors en EG chevronné, pas en assistant générique.

🔴 **Vous n'êtes pas entrepreneur ?** Tout le reste — courriels, calendrier, dossiers,
comptabilité — est **neutre**. Une seule section de `CLAUDE.md` porte le métier : réécrivez-y
vos quatre réflexes, supprimez le profil EG, et le poste est à vous. Les taxes suivent
(`--taux1` / `--taux2`, `--taux2 0` pour une province à taxe unique).

Ensuite, parlez normalement : *« qu'est-ce qui attend une réponse ? »*, *« qu'est-ce qui vient
cette semaine ? »*, *« où en est ce client ? »*. Et la phrase qui fait grandir la mémoire du
poste : **« note ça »**.

---

## Aucun secret n'est stocké

L'accès passe par le profil Outlook **déjà authentifié** de votre poste. Si Outlook fonctionne
pour vous, l'outil fonctionne.

C'est la voie **nominale**, et la seule qui atteigne la boîte *vivante* : l'authentification de
base d'Exchange Online est retirée (IMAP/POP fin 2022, SMTP AUTH avril 2026).

⚠️ Ce n'est pourtant pas la seule voie du poste : `ost_reader.py` lit un fichier `.ost`/`.pst`
**en binaire, sans Outlook ni MAPI**. Un `.ost` contient le courrier en clair — ne jamais en
déposer un dans le dossier synchronisé. Détail : `CLAUDE.md` §1.

⚠️ Le dossier étant destiné à un espace synchronisé, **n'y écrivez jamais de secret**. Les
fichiers de mémoire sont conçus pour *pointer vers* ces informations, pas pour les contenir.

---

## Ce qui est gratuit, et ce qui ne l'est pas

**Ce dépôt est gratuit et le restera** : licence MIT, aucune condition, aucune donnée qui
remonte à qui que ce soit. Tout tourne sur votre poste.

**Claude Code, lui, demande un abonnement Claude payant.** Ce n'est pas nous qui le vendons
et nous n'en tirons rien — mais autant le savoir avant d'installer Python pour rien.

---

## Licence

MIT — voir [LICENSE](LICENSE). Utilisez-le, modifiez-le, distribuez-le.

Si vous l'améliorez, les corrections mesurées sont les bienvenues : **une mesure vaut mieux
qu'une intuition**, et c'est la règle qui a construit ce poste.

---

<div align="center">

**Une gracieuseté de Sylvain Leduc — Constructo AI inc.**

*Écosystème intelligent pour la construction au Québec*

[www.constructoai.ca](https://www.constructoai.ca)

</div>
