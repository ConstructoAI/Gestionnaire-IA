# JOURNAL — l'histoire du poste

> **Satellite — ne se charge PAS automatiquement**, et c'est voulu : il ne coute rien
> tant qu'on ne l'ouvre pas. Le hub `CLAUDE.md` porte les regles ; ce fichier porte
> **ce qui s'est passe et pourquoi** : chaque decision avec la mesure qui l'a produite.
>
> 🔴 **Une section vide se lit « pas encore consigné », JAMAIS « rien ne s'est passé ».**
> Un gabarit vide bien formé a l'air complet : c'est un faux zéro. Ne pas le remplir
> artificiellement — l'alimenter au fil des passes.
>
> **Ne duplique rien.** Le hub garde les regles et la carte ; ici on garde l'histoire.
>
> **Append-only.** Une entree datee ne se reecrit pas : c'est un journal, pas un etat. Les
> corrections d'etat vont dans le hub ou les satellites ; ici **on ajoute**.

> Cree le : **2026-08-29**. Derniere mise a jour : **2026-08-29**.

---

## Quand l'ouvrir

Avant de **defaire une decision**, pour savoir ce qu'elle a coute · quand un **piege
reapparait**, pour verifier s'il a deja ete paye · quand une regle du hub **parait arbitraire**,
pour lire la mesure qui l'a produite.

## Format

```
- **AAAA-MM-JJ** — <ce qui a ete fait ou trouve>.
  <La mesure : la commande, sa sortie reelle, le chiffre.>
  ⚠️ <Ce qui a failli mal tourner, ou l'ecart entre ce qu'on croyait et ce qui est.>
```

Une entree sans mesure ne vaut rien. Une entree qui dit seulement « corrige » non plus :
c'est **ce qui etait faux et comment on l'a su** qui servira dans six mois.

---

- **2026-08-29** — Premiere passe d'audit complete du poste, avant toute mise en service.
  26 fichiers, 8 dossiers, ~8000 lignes (dont 3529 pour le seul profil EG). Cinq agents sur
  cinq perimetres disjoints (moteur courriel · calendrier+veille · comptabilite+profil ·
  amorcage+config · coherence documentaire), lecture seule ; corrections faites par le chef.
  Aucun secret, aucune adresse reelle, aucun numero de taxe, aucun chemin nominatif —
  zeros PROUVES, chaque motif ayant d'abord montre qu'il savait trouver.
  Fins de ligne conformes au §8 : LF sur les 25 fichiers de `.claude\`, CRLF pour le seul `.bat`.

- **2026-08-29** — 🔴 La signature livree EST le gabarit vierge. Le piege le plus couteux du lot.
  `md5sum .claude/profiles/signature_*.html` rend `15480acb8000e7152aeefd9d58cdb9f9` pour les
  DEUX ; `cmp` silencieux. `signature_defaut.html` == `signature_MODELE.html` au bit pres.
  ⚠️ Le hub disait « renommer signature_MODELE.html en signature_defaut.html » : c'etait
  DEJA fait dans le poste livre, donc rien ne signalait qu'il restait a remplir. Et le hub
  ordonne par ailleurs `--signature` sur TOUT draft. Les deux consignes justes, combinees,
  faisaient partir « PRENOM NOM / 1 (000) 000-0000 / adresse@exemple.ca » chez un vrai client.
  Le garde-fou du script ne mord pas : `charger_signature` (`outlook_mail.py:431-435`) ne leve
  que si le fichier MANQUE — il existe. Et la verification que le hub recommandait (« chercher
  le bloc de coordonnees ») PASSE AU VERT : le bloc est la, il est faux.
  Hub corrige : section « PIRE ENCORE » ajoutee au §2, avec le temoin `grep -c "PRENOM NOM"`.

- **2026-08-29** — Le hub affirmait que MAPI/COM etait « la seule voie qui tienne ». Faux.
  `.claude\scripts\ost_reader.py` (346 l.) parse le format binaire OST/PST **sans Outlook, sans
  MAPI, sans profil authentifie**, et contourne par balayage le verrou qu'Outlook pose sur les
  1024 premiers octets (`ost_reader.py:1-23,34,49-91` ; aucun `win32com` dans le fichier).
  Un `.ost` contient le courrier en clair et le script s'ouvre sur n'importe quel chemin.
  ⚠️ Ce script n'etait cite NULLE PART : ni dans la carte du hub, ni dans SKILL.md, ni dans
  l'agent. Idem `factures.py`. La carte annoncait 4 scripts ; `ls .claude/scripts/` en rend 6.
  Hub corrige : carte (6 scripts + `references\` + `INSTALLATION.md`), avertissement au §1,
  et section `factures.py` ajoutee au §6 — le hub citait ses drapeaux `--taux1`/`--taux2`
  sans jamais nommer le script qui les accepte.

- **2026-08-29** — Deux faux zeros dans `outlook_calendar.py`, tous deux silencieux.
  `--jours 0` : `:133-135` porte trois affectations dont les deux premieres sont integralement
  ecrasees par la troisieme. Fenetre `[maintenant, maintenant]`, donc 0 rendez-vous, code 0,
  aucun avertissement. `--limite` : plafond a 100 (`:141,248`) jamais annonce en sortie ; avec
  `IncludeRecurrences` une serie quotidienne le sature seule.
  ⚠️ Et `delete` NE PREVIENT PAS pour une serie, contrairement a ce que le hub affirmait :
  `update` avertit avant (`:201-203`), `delete` enregistre `serie=True` puis supprime puis
  l'annonce (`:227-229`) — l'info arrive apres que toute la serie est partie.
  Hub corrige : pieges 7 et 8 ajoutes au §3, et la phrase « le script avertit » remplacee.

- **2026-08-29** — Le hub se trompait sur son propre piege `--json`.
  Il annoncait « code de sortie 0 : la commande semble avoir reussi ». Mesure :
  `outlook_mail.py:645` fait `p.parse_args()`, pas `parse_known_args` — argparse sort en **2**,
  sur stderr. La regle de placement, elle, est juste (`--json` declare l.588, avant
  `add_subparsers` l.589).
  ⚠️ Une justification fausse sous une regle juste : elle decredibilise la regle quand on la teste.

- **2026-08-29** — Le hub n'exposait que 2 regimes de prix sur 5, et figeait l'admin a 3 %.
  `grep -inE "contingences\s*[0-9]+\s*%" profil.txt` rend : resid neuf ×1,30 (l.57) · resid reno
  ×1,33 (l.1872) · **commercial neuf ×1,28** (l.2277) · **commercial reno ×1,34, admin 4 %**
  (l.2690) · **institutionnel ×1,30, admin 5 %** (l.3151).
  ⚠️ Un chiffrage commercial ou institutionnel fait « selon le hub » etait donc faux.
  Arithmetique revalidee : majorations ADDITIVES (3+12+15=30 %, donc 1,30), jamais composees
  (1,03×1,12×1,15 = 1,3266). Taxes justes au cent : 4160 donne 208,00 + 414,96 = 4782,96, sans
  cascade. L'ecart neuf/reno de 3 449,25 $ tombe pile — mais c'est du TTC ; en HT c'est 3 000 $,
  ce que le hub ne precisait pas. Table des 5 regimes substituee a la table des 2.

- **2026-08-29** — Plages du profil EG mal etiquetees, et taux CCQ mal dates.
  `346–600` annoncait « verification arithmetique d'un agrandissement » : le bloc s'arrete a
  ~385 ; 386-600 couvre le toit plat, la structure seule, la coherence inter-pages.
  Les taux CCQ : le hub disait « janvier 2026 », la source (`profil:3334`) dit **« EN VIGUEUR
  AU 28 DECEMBRE 2025 »**, et le fichier est genere le 17 mars 2026. Trois dates, le hub
  retenait la plus flatteuse — sur une grille de main-d'oeuvre destinee a une soumission.
  ⚠️ Le socle tarifaire au pi² est millesime **Quebec 2025** et rien ne le signalait.
  ⚠️ Le compte de 3529 lignes du hub est JUSTE : c'est `wc -l` qui rend 3528, le fichier
  n'ayant pas de saut de ligne final. Ma propre mesure d'ouverture etait la mauvaise.

- **2026-08-29** — Le hub violait sa propre regle §7 : la table « les calendriers du poste »
  existait a DEUX endroits — hub §3 et `ETAT_calendrier.md:19-25` — aucun ne pointant vers
  l'autre, alors que le satellite proclame « Ne duplique rien » en tete.
  Le satellite gagne (la donnee grossit et vieillit) ; le hub renvoie desormais, comme le
  §6 le fait deja pour la comptabilite. Modele a suivre pour les autres doublons.
  ⚠️ Reste non tranche : le « registre par correspondant » (§2) est du detail par personne,
  qui croit sans borne, dans le fichier AUTO-CHARGE — et aucun des cinq satellites n'a ce
  travail. Ce n'est pas un doublon, c'est un trou dans le decoupage. A loger.

- **2026-08-29** — ⚠️ Deux faux zeros produits par MES PROPRES mesures pendant cette passe.
  (1) Un `grep -oE` pour balayer les chemins en dur du poste : **0 sur 10**, sans erreur —
  motif casse par l'echappement du shell. Refait en Python avec temoin positif : 10 chemins,
  tous generiques. Consigne ajoutee au §4.
  (2) Le script qui a corrige ce hub a ecrit `"\references"` en chaine Python NON BRUTE :
  `\r` est un retour chariot. La ligne est partie sur disque en `.claude` + CR + `eferences`.
  Python n'a averti que pour les AUTRES echappements, jamais pour celui-la, qui est valide.
  Consigne ajoutee au §8 : textes de correction en chaine brute ou via un fichier de donnees,
  et compter les octets CR apres ecriture. Les corrections suivantes ont transite par un
  fichier de donnees ; verification finale : 0 octet CR.

- **2026-08-29** — Ce qui TIENT, verifie et non modifie (a ne pas re-auditer sans raison) :
  les deux hooks sont en forme exec sur le chemin absolu de `powershell.exe`
  (`settings.json:11,28`) — le piege §8 est integralement evite ; `settings.json` est portable
  (aucun chemin nominatif) ; le `.bat` distingue reellement « processus present » de « MAPI
  repond » (`:198-201`), borne a ~40 s, et se replace tout seul qu'il soit dans `.claude\` ou a
  la racine ; les verrous `--yes-send` et `--yes-write` sont dans le CODE et levent avant tout
  acces MAPI ; aucune purge definitive n'est exposee ; `factures.py` et `ost_reader.py` sont
  strictement en lecture seule et sans secret ; les cinq satellites sont des gabarits vierges
  corrects — **0 section remplie sur 15**, aucune donnee client, aucune fuite.
  ⚠️ Restent ouverts, non corriges ce jour : `.gitignore:12` reamet `signature_defaut.html`
  (le seul fichier qu'on demande de remplir avec ses vraies coordonnees) et n'ignore aucun
  `ETAT_*` ; `check_setup.py` n'a pas de `reconfigure` et fait de `tasklist` un verrou sur le
  test MAPI ; le garde `delete` d'`outlook_mail.py:555` liste des noms de dossier traduits
  (fr/en/de/es) et laisse passer it/pt/nl ; `SKILL.md` recopie 8 blocs du hub dont 3 deja
  appauvris. Ce sont des decisions du proprietaire du poste, pas des corrections de doc.

- **2026-08-29** — 🔴 Le lanceur refusait de demarrer sur le poste de son propre auteur.
  `Constructo_AI.bat:158` testait `where claude.cmd` puis `exit /b 1`. Mesure en cmd, avec
  temoins positif et negatif : `where claude.cmd` INTROUVABLE · `where claude` TROUVE
  `C:\Users\<user>\.local\bin\claude.exe` v2.1.251 · temoin `where cmd.exe` TROUVE · temoin
  `where zzz_inexistant` INTROUVABLE.
  ⚠️ L'installateur NATIF produit `claude.exe` ; le nom `claude.cmd` n'existe qu'avec une
  installation **npm** (`%APPDATA%\npm\claude.cmd`, absent ici). Claude Code etait installe,
  fonctionnel, et le poste declarait qu'il manquait.
  ⚠️ Premiere erreur de MA part ce jour-la : j'ai d'abord lu un `errorlevel=0` rassurant, qui
  etait un artefact — `cmd` developpe `%errorlevel%` a l'analyse de la ligne, donc AVANT que
  `where` ne s'execute. Refait avec `&&` / `||`, qui eux ne mentent pas.

- **2026-08-29** — Le `.bat` devient un INSTALLATEUR, pas seulement un lanceur.
  Decision du proprietaire du poste, trois choix actes : installateur natif d'abord et Node en
  secours · UNE seule question pour tout le lot · niveau utilisateur, aucune elevation UAC.
  Six etapes desormais : inventaire sans effet de bord · installation de ce qui manque ·
  Outlook · mise a jour · attente MAPI · session. **Si rien ne manque, aucune question n'est
  posee.** Quatre voies tentees dans l'ordre pour Claude Code : curl, PowerShell, winget, npm.
  Commandes prises dans la doc officielle (code.claude.com/docs/en/setup, consultee ce jour),
  jamais de memoire.
  ⚠️ Non verifie : le comportement reel des installations. Tester le `.bat` de bout en bout
  supposerait d'installer des logiciels ; seules les phases SANS effet de bord ont ete
  executees. Ce qui a ete mesure : l'inventaire, les deux branches de la question, la
  resolution des chemins, l'absence de `goto`/`call` orphelin, l'encodage.

- **2026-08-29** — Les chemins sont ABSOLUS mais jamais CODES EN DUR. La question posee etait
  bonne : le depot part sur GitHub, donc chaque poste sera different.
  Mesure 1 — `grep -nE "[A-Za-z]:\\\\"` sur les 438 lignes : **une seule occurrence, dans un
  commentaire d'illustration**. Aucun nom propre de ce poste (`ThinkPad`, `Python313`,
  `AppData`, `nodejs`) : zero occurrence. Tout passe par `%USERPROFILE%`, `%SystemRoot%`,
  `%ProgramFiles%`, `%CD%`, resolus par Windows a l'execution.
  Mesure 2 — le MEME fichier lance avec `USERPROFILE=C:\Users\Marie` et un `PATH` minimal
  rend : Claude ABSENT, Python ABSENT, Git ABSENT, `CLAUDE_EXE` vide, verdict « il manque
  quelque chose ». Sur ce poste-ci, sans rien changer au fichier : tout present, verdict
  « rien a installer ». **Le fichier s'adapte, il ne suppose rien.**
  ⚠️ Durcissement au passage : `where` regarde le DOSSIER COURANT avant le PATH, et ce dossier
  est l'espace OneDrive/SharePoint synchronise. Une commande nue y prendrait un executable
  etranger. Le `.bat` resout donc `claude` et `python` en chemin absolu UNE fois, en ecartant
  explicitement toute correspondance dans le dossier courant, et n'utilise que ce chemin.
  *(Ce point precis — `where` regarde-t-il vraiment le dossier courant — n'a PAS ete mesure
  proprement : deux tentatives cassees par le quoting. Rendu sans objet par le durcissement
  plutot qu'affirme.)*

- **2026-08-29** — 🔴 `.gitignore` reintroduisait la fuite qu'il pretendait empecher.
  `.gitignore:10` ignorait `signature_*.html`, puis `:12` portait `!signature_defaut.html` —
  c'est-a-dire exactement le fichier que le poste demande de remplir avec le vrai nom, le vrai
  telephone, la vraie adresse et le vrai courriel. Rempli puis `git add .`, il partait en clair.
  Devenu urgent : le depot est destine a GitHub.
  ➜ La negation est supprimee. Seul `signature_MODELE.html` se publie. Effet voulu : un depot
  fraichement clone n'a PAS de `signature_defaut.html`, donc `--signature` **echoue franchement**
  en listant les signatures disponibles, au lieu d'envoyer en silence « PRENOM NOM /
  adresse@exemple.ca » a un vrai client. Le danger silencieux devient une panne bruyante.
  ⚠️ Les cinq satellites, eux, restent publies : ce sont des gabarits, ils doivent arriver chez
  celui qui telecharge. Le `.gitignore` porte desormais la commande de verification a lancer
  avant chaque push depuis un poste EN SERVICE, et les quatre lignes a decommenter si le meme
  depot sert a la fois de poste et de gabarit.

- **2026-08-29** — Sauvegarde de l'ancien lanceur :
  `.claude\Constructo_AI.bat.2026-08-29.bak`, ignoree par git (`*.bak`). Supprimable sans
  risque une fois le nouveau lanceur eprouve en conditions reelles.

---

## Seconde passe du 2026-08-29 — campagne de tests, cinq agents

- **2026-08-29** — 🔴 CORRECTION D'UNE ERREUR DE LA PREMIERE PASSE : les dates CCQ.
  J'avais ecrit ce matin que la source ne dit « nulle part » janvier 2026. **Elle l'ecrit en
  toutes lettres.** Le profil porte DEUX blocs distincts : `profil:3333` « EN VIGUEUR AU
  28 DECEMBRE 2025 » (charges patronales, grilles APCHQ consolidees, l.3331-3381) et
  `profil:3383` « TAUX HORAIRE CCQ JANVIER 2026 » (les taux par metier, l.3382-3529).
  ⚠️ Ma correction avait colle la date du premier bloc sur le second. La formulation
  d'origine du hub etait juste et sourcee ; je l'avais degradee. Sur une grille de
  main-d'oeuvre qui sert a soumissionner.
  ➜ Le hub porte desormais les deux blocs avec leurs deux dates, et l'avertissement de ne pas
  les confondre. Lecon : une correction se verifie a la source aussi rigoureusement que
  l'affirmation qu'elle remplace.

- **2026-08-29** — 🔴 `factures.py` retenait le PLUS GROS quadruplet qui boucle.
  `resoudre()` parcourait `sorted(S, reverse=True)` et rendait le premier succes.
  Reproduction : un document portant la vraie facture du jour (HT 1000 / TPS 50 / TVQ 99,75 /
  1149,75) ET une facture anterieure rappelee (2000 / 100 / 199,50 / 2299,50) rendait
  **2299,50 marque `fiable=true`** — 100 % de trop, en silence.
  ⚠️ N'importe quel document citant un autre jeu de montants declenchait le cas : facture
  precedente, extra, soumission rappelee, note de credit.
  ➜ Si PLUSIEURS HT bouclent, on n'en retient AUCUN : l'ambiguite s'isole en « a verifier ».
  Verifie sur corpus fabrique : le cas Epsilon sort desormais en « a verifier ».

- **2026-08-29** — 🔴 Trois autres pertes silencieuses dans `factures.py`, toutes mesurees
  sur un corpus fabrique de 7 factures.
  (1) Deux fichiers de meme horodatage ET meme client : le second ECRASAIT le premier, et la
  facture perdue n'apparaissait dans AUCUN compteur. Mesure : deux factures de 2299,50 ne
  comptaient que pour une. ➜ isolee en « DOUBLON horodatage+client, NON ADDITIONNE ».
  (2) Fine insecable U+2009 : « 1 234,56 » lu **234,56**. Virgule anglo : « 4,160.00 » lu
  **160.00**. Le millier etait ampute ; seule la regle « le quadruplet doit boucler » empechait
  la valeur fausse de sortir. ➜ toutes les espaces de milliers tolerees, le DERNIER separateur
  est le decimal. Verifie : les deux formats sont maintenant lus juste.
  (3) `if x` jetait un montant de 0,00 $. ➜ `if x is not None`.

- **2026-08-29** — 🔴 FUITE DE DONNEES fermee avant publication GitHub.
  `outlook_mail.py:414-415` nommait en clair **deux domaines d'entreprises tierces** dans une
  docstring — des employeurs anterieurs. Le depot part sur GitHub.
  ➜ Anonymise : « deux d'entreprises anterieures ». La mesure garde toute sa valeur sans les
  noms. Balayage final : zero occurrence, temoin positif au vert.

- **2026-08-29** — 🔴 `.gitignore` : quatre trous mesures, dont un grave.
  `git status` classait comme PUBLIABLES : `*.ost` et `*.pst` (**un fichier de donnees Outlook
  contient TOUT le courrier en clair**, et le hub interdit deja d'en deposer un ici),
  `.claude/settings.local.json` (regles de permission du poste), `.env`, et les sauvegardes
  hors motif `*.backup` `*.sav` `*.tmp`.
  ➜ Tous ignores. Verifie dans un depot jetable : `git ls-files` ne les rend plus.

- **2026-08-29** — 🔴 `.gitattributes` normalisait les BINAIRES de `.claude/`.
  `git check-attr` rendait `text: set, eol: lf` sur `.claude/profiles/logo.png` : la regle
  `.claude/** text eol=lf` ecrase le `* -text` du debut. Un PNG de signature ou un `.xlsx`
  commite en serait ressorti **corrompu**. Et un `.cmd` depose dans `.claude/` subissait la
  panne meme que la ligne `*.bat text eol=crlf` evite au `.bat`.
  ➜ `*.cmd` en CRLF, et onze extensions binaires marquees `binary`. Verifie : `logo.png` rend
  desormais `text: unset`, `lanceur.cmd` rend `eol: crlf`.

- **2026-08-29** — 🔴 Le hook `SessionStart` fabriquait une fausse mesure, en silence.
  Lance ailleurs qu'a la racine du poste, `Join-Path (Get-Location)` ne trouvait pas
  CLAUDE.md : le hook annoncait « modifie il y a **-1 jour(s)** », **perdait le bloc GABARIT**,
  et sortait quand meme en **0**. Un message d'apparence normale, entierement faux.
  ➜ Il utilise desormais `$env:CLAUDE_PROJECT_DIR` avec repli sur le dossier courant, et
  **echoue bruyamment** si le fichier reste introuvable. Mesure apres correction : depuis
  `C:\Windows` sans la variable, exit **1** + « CLAUDE.md introuvable sous C:\Windows » ; avec
  la variable posee, exit 0 et message juste. Au passage, `[int]` arrondissait (13 h annoncees
  « 1 jour ») : remplace par `Floor`.

- **2026-08-29** — 🔴 `check_setup.py` — le PORTIER du poste — mourait sur un accent.
  Seul des six scripts a n'avoir pas de `reconfigure`, et il imprime des noms de comptes
  Outlook, donnee externe. Sur console cp1252, un nom hors latin-1 levait `UnicodeEncodeError`
  EN PLEIN diagnostic : le script mourait apres avoir affiche des `[ OK ]`, sortait != 0, et le
  `.bat` basculait en mode degrade pour une raison purement cosmetique.
  ➜ Ajoute, sous `try/except` — indispensable, ce script pretend diagnostiquer Python < 3.8 ou
  `.reconfigure` peut manquer. Et `stderr` reconfigure dans les six, en `backslashreplace`
  (qui conserve le caractere perdu) plutot que `replace` (qui le remplace par `?` sans trace).

- **2026-08-29** — 🔴 `--jours 0` corrige DANS LE CODE, pas seulement documente.
  La premiere passe avait documente le faux zero sans le reparer. `outlook_calendar.py`
  portait trois affectations de `fin` dont les deux premieres etaient mortes. Reproduction
  isolee : `jours=0` donnait une fenetre de largeur **0:00:00** ; `jours=-5` une fenetre
  INVERSEE, meme symptome, sans avertissement.
  ➜ `--jours 0` = aujourd'hui jusqu'a 23:59:59, comme l'intention le disait ; `--jours` negatif
  refuse net. L'import local `from datetime import timedelta` remonte en tete.

- **2026-08-29** — 🔴 Le filtre des feries de `veille_poste.py` ne matchait JAMAIS en francais.
  `if nom in ("United States holidays", "Anniversaires")` : comparaison exacte, sensible a la
  casse, sur des noms de dossiers Outlook. Sur un profil fr-CA ils s'appellent « Jours feries
  au Canada », « Birthdays ». L'exclusion ne se declenchait jamais et la veille se noyait dans
  le bruit. Regle du motif trop strict, section 4.
  ➜ Sous-chaine insensible a la casse sur cinq racines, francais et anglais.

- **2026-08-29** — 🔴 `--signature MODELE` reussissait et envoyait le gabarit a un vrai client.
  Consequence non anticipee du `.gitignore` : un depot clone ne porte plus qu'UNE signature,
  `MODELE`, et `charger_signature` la listait comme « disponible ».
  ➜ `MODELE` explicitement refuse, avec la marche a suivre. Mesure : le refus tombe AVANT tout
  contact MAPI. Et le commentaire HTML en tete du fichier partait **dans le corps du courriel**
  (invisible au rendu, lisible en « afficher la source ») : signale dans le fichier lui-meme.

- **2026-08-29** — 🔴 L'agent `courriels.md` se declarait responsable d'ECRIRE un satellite.
  `:99-100` disait « c'est ta passe qui le nourrit, et personne d'autre ne le peut » ; ses
  outils sont `Bash, PowerShell, Read, Grep, Glob` — **ni `Write` ni `Edit`** — et onze lignes
  plus bas il disait l'inverse. Comme le poste tourne sans controle de permission, il aurait
  ecrit via Bash malgre tout, avec le risque CRLF de la section 8.
  ➜ Tranche en lecture seule : il MESURE, le chef ECRIT.

- **2026-08-29** — 🔴 SEPT defauts du nouveau lanceur, tous mesures, tous corriges.
  (1) `-Command "irm ... ^| iex"` : entre guillemets cmd ne consomme pas le caret, PowerShell
  recoit `^|` et leve « Jeton inattendu ^ ». **La voie 2 n'avait JAMAIS pu reussir, sur aucun
  poste.** Temoin : sans caret, `BBB` et code 0.
  (2) Faux executable a la RACINE d'un lecteur : `%CD%` vaut deja `X:\`, donc la comparaison
  `"%CD%\"` testait `X:\` contre `X:\\` — jamais vrai. Un faux `claude` du dossier courant
  etait RETENU. Cas reel : lecteur reseau mappe, SharePoint monte en racine, `subst`.
  ➜ `%CD%` normalise une fois. Verifie avec temoin positif : le faux REPOND (`9.9.9`) et il est
  quand meme REJETE.
  (3) Executable MUET declare « present » : le seul test etait l'existence. Or
  `%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe` fait **0 octet** — l'alias du Microsoft
  Store — et `where python` le renvoie. Le poste partait avec un outillage inexistant **sans
  meme poser la question**. ➜ On exige que `--version` rende une chaine NON VIDE. Verifie.
  (4) `%PY%` invoque sans `call` : si `python` resout vers un `.bat` (shim pyenv-win, pixi,
  conda), le controle est transfere et ne revient JAMAIS — script arrete net, code 0, silence.
  ➜ `call` partout.
  (5) Une reponse contenant `"` sortait en **255** avec une erreur systeme en francais. ➜ les
  guillemets sont neutralises avant tout test ; le cas sort maintenant en 1, proprement.
  (6) « Les quatre voies ont ete tentees » etait FAUX : winget et npm sont conditionnels et
  souvent sautes. Le diagnostic envoyait chercher un probleme de reseau alors que la cause
  etait « winget absent ». ➜ On annonce les voies REELLEMENT tentees.
  (7) Refus d'installation + Claude absent : le script mourait en `exit /b 1` avec un message
  qui inventait une cause (« une installation toute fraiche n'est parfois visible qu'au
  prochain demarrage ») alors que RIEN n'avait ete installe. ➜ Message veridique.

- **2026-08-29** — ⚠️ Le PATH n'etait rafraichi que pour Claude Code.
  `winget install --scope user` ecrit dans le PATH UTILISATEUR : une fenetre cmd deja ouverte
  ne le verra jamais. Claude avait son rattrapage, Python et Git non — le `.bat` annoncait donc
  « Python n'a pas pu etre installe » APRES une installation reussie. ➜ `:apres_install_python`
  et `:apres_install_git` ajoutes.
  Autres corrections du meme lot : `oui`/`OUI`/`Y` acceptes (ils valaient REFUS, en silence, sur
  un poste entierement francophone) · `call "%TEMP%\cc_install.cmd"` en chemin absolu, et
  `%TEMP%` teste avant, pour ne jamais telecharger puis EXECUTER un installateur dans le
  dossier synchronise · les SIX scripts du moteur verifies, pas seulement le premier ·
  `--fallback-model` ajoute, un abonnement Pro n'ayant pas forcement acces a Opus 1M · `pause`
  si la session sort != 0, sinon la fenetre se ferme sur le message d'erreur · le code de sortie
  de `claude` propage au lieu d'un `exit /b 0` systematique.

- **2026-08-29** — Documentation realignee sur le code, apres la campagne.
  `README.md` etait le seul fichier non touche de la journee et il etait devenu faux sur cinq
  points, dont « c'est la **seule voie** qui tienne » — l'assertion que le hub avait dementie le
  matin meme a cause d'`ost_reader.py`. C'est le premier fichier que lit un visiteur GitHub.
  ⚠️ Et il taisait, comme `INSTALLATION.md`, que le poste tourne en `bypassPermissions` +
  `--dangerously-skip-permissions` : Claude n'y demande AUCUNE autorisation avant de lire,
  d'ecrire ou d'executer dans un dossier OneDrive plein de donnees clients. Le README vendait
  « copiez un dossier, double-cliquez » en listant quatre regles de securite, ce qui le faisait
  paraitre plus prudent qu'il n'est. ➜ Section explicite ajoutee, avec la marche a suivre pour
  revenir au mode prudent.
  Aussi : `INSTALLATION.md` listait 4 scripts sur 6 et decrivait un `signature_defaut.html`
  qui n'existe plus dans un clone · `SKILL.md` disait « six pieges » quand le hub en numerote
  huit, mettait `update` et `delete` dans le meme sac, et ignorait `factures.py` ·
  `depannage.md` portait le seul chemin relatif faux des dix invocations documentees.

- **2026-08-29** — CE QUI N'A PAS PU ETRE TESTE, et ne doit pas etre pris pour verifie :
  les INSTALLATIONS reelles (curl vers claude.ai, winget, npm, pip) — les verifier supposerait
  d'installer des logiciels sur le poste · le lancement reel de la session et `claude update` ·
  toute ecriture Outlook, et les 19 sous-commandes MAPI · `ost_reader.py` au-dela de `--help`,
  faute de `.ost` d'essai · le Mark-of-the-Web sur un `.bat` extrait d'un ZIP telecharge depuis
  GitHub, qui est pourtant le tout PREMIER ecran que verrait un cloneur.
  ⚠️ Le vrai test reste une machine Windows nue.

- **2026-08-29** — Quatre derniers faux zeros fermes dans `outlook_mail.py` et `veille_poste.py`.
  (1) `folders` descendait a 5 niveaux, `search` et `thread` a 3 : un message dans un dossier
  profond etait VISIBLE dans `folders` et INVISIBLE a la recherche, sans un mot. ➜ aligne a 5.
  (2) Le plafond de 500 messages par dossier etait MUET dans `search` et `thread` : une
  correspondance plus ancienne etait perdue en silence. ➜ avertissement sur stderr des que le
  plafond est atteint. Un zero ne vaut que prouve.
  (3) Une `reply` SANS `--signature` affectait `.Body` : sur un original HTML, TOUT le fil cite
  etait converti en texte brut. Le commentaire juste au-dessus identifiait ce risque et ne le
  neutralisait que dans l'autre branche. ➜ passe par `HTMLBody`, comme la branche signature.
  (4) `veille_poste --intervalle 0` levait `ZeroDivisionError`, et `--remise-a-zero` effacait
  le fichier d'etat AVANT de planter. ➜ validation posee avant tout effet de bord ; verifie :
  l'arret tombe avant `_win32()` et avant `os.remove`.

- **2026-08-29** — ⚠️ TROISIEME occurrence du piege d'echappement de la section 8, dans MES
  outils cette fois. Un heredoc bash a mange un antislash dans un motif de recherche de
  chemins nominatifs : `[A-Za-z]:\\Users\\` est arrive comme `[A-Za-z]:\Users\`, et Python a
  leve `incomplete escape \U`. **Echec bruyant, donc sans dommage** — contrairement a la
  premiere occurrence du matin, ou le meme accident avait produit un ZERO silencieux.
  ➜ Regle de travail : tout motif contenant des antislash passe par un FICHIER ecrit, jamais
  par un heredoc. Et le temoin positif reste obligatoire.

- **2026-08-29** — ETAT A LA FIN DE LA CAMPAGNE. 27 fichiers, 0 non-conforme (LF partout,
  CRLF pour le seul `.bat` et son `.bak`, aucun BOM). Les 6 scripts compilent. `settings.json`
  valide. Simulation de publication : **25 fichiers sur 27 partiraient**, les deux ecartes
  etant le `.bak` et `signature_defaut.html` — tous deux volontairement.
  Balayage de fuite, chaque motif valide par un temoin positif : aucun courriel reel, aucun
  telephone, aucun numero TPS/TVQ, aucune cle, aucun chemin nominatif. Les deux seules
  occurrences de `C:\Users\<nom>` sont l'exemple fictif « Marie ».

- **2026-08-29** — PUBLICATION sur GitHub : `ConstructoAI/Gestionnaire-IA`, depot **PUBLIC**.
  Le dossier local n'etait PAS un depot git ; le distant portait deja **5 commits**. Un
  `git init` suivi d'un push aurait ecrase cet historique. ➜ Clone du distant, application de
  l'etat corrige par-dessus, commit sur une branche `audit-2026-08-29`, PR #1. Historique
  preserve, correction revocable par `git revert`.
  ⚠️ **`signature_defaut.html` etait DEJA suivi par git** : ajouter une regle au `.gitignore`
  ne detrack pas un fichier deja versionne. Il a fallu `git rm --cached`. Sans ce geste, le
  jour ou quelqu'un remplit sa signature avec son vrai nom et son vrai telephone, un
  `git add .` l'aurait publiee — la protection du matin n'aurait servi a rien.
  ➜ Regle a retenir : une regle `.gitignore` ne vaut que pour les fichiers PAS ENCORE suivis.
  Verification finale sur un clone frais de la branche, avec `core.autocrlf=true` — le defaut
  Windows, celui qui casse tout : **25 fichiers, 0 fin de ligne non conforme**, les 6 scripts
  compilent, `signature_defaut.html` absent et `signature_MODELE.html` present, comme voulu.

---

## Le PREMIER essai sur une machine reelle — 2026-08-29

- **2026-08-29** — 🔴 UNE MACHINE WINDOWS 11 NEUVE A TROUVE CE QUE CINQ AGENTS ONT RATE.
  Premier lancement du `.bat` sur un poste tiers, jamais utilise pour ce projet. Resultat :
  **Claude Code s'est installe** par la voie 1 (curl), v2.1.251 dans
  `%USERPROFILE%\.local\bin`, sans droits admin — la partie neuve du lanceur fonctionne.
  **Puis le poste a cru Python present alors qu'il ne l'etait pas**, n'a donc jamais installe
  le vrai Python, et `pip install pywin32` a echoue en affichant le chemin du coupable :
  `%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe`.
  ⚠️ **Windows livre un raccourci d'execution de 0 octet a cet endroit**, et `where python` le
  renvoie. Mon garde-fou de la veille exigeait que `--version` rende une chaine NON VIDE. Ca ne
  suffit pas : ce raccourci est BAVARD. Mesure des deux comportements :
  · avec un Python du Store installe derriere, il relaie et rend une VRAIE version (`Python
  3.13.14` sur mon poste) ; · sans rien derriere, il imprime « Python est introuvable; executez
  sans arguments a installer a partir du Microsoft Store » (poste d'essai).
  **Les deux sorties sont non vides.** Aucune simulation ne l'avait produit : mes faux
  executables de test etaient MUETS, jamais bavards.
  ➜ Le seul test qui ne ment pas : **faire EXECUTER du Python**. `python -c "print(84)"` doit
  rendre exactement 84. Un stub repond ; un interpreteur calcule. Verifie sur les deux cas :
  le faux est rejete avec la mention explicite du raccourci Store, le vrai passe.

- **2026-08-29** — ⚠️ DEUX REGRESSIONS QUE J'AI INTRODUITES EN CORRIGEANT, et rattrapees par
  le test avant publication. Meme cause batch, deux fois.
  (1) `for /f "delims=" %%V in ('"%~f1" -c "print(42*2)"')` : **DEUX arguments entre
  guillemets** dans un `for /f` font retirer par `cmd` le premier et le dernier guillemet.
  Resultat : « La syntaxe du nom de fichier, de repertoire ou de volume est incorrecte », et le
  VRAI Python etait rejete lui aussi. ➜ La sonde passe desormais par un FICHIER
  (`> "%TEMP%\cc_pysonde.txt"` puis `for /f "usebackq" ... ("fichier")`), qui n'a pas ce piege.
  (2) Le meme piege dans `:essai_claude`, ou j'avais mis un tube vers `findstr` a l'interieur du
  `for /f` : **Claude Code etait annonce ABSENT alors qu'il etait installe**. ➜ La validation de
  la version se fait maintenant HORS du `for /f`.
  ⚠️ Et un troisieme, trouve par mon propre test : la sonde invoquait le candidat **sans
  `call`**. Sur un shim `.bat` — pyenv-win, pixi, conda — le controle est transfere et ne revient
  jamais : le script mourait en silence pendant l'inventaire. C'est exactement le defaut que
  l'audit avait releve la veille, et que j'ai reintroduit dans une ligne neuve.
  ➜ Lecon : en batch, **toute** invocation d'un chemin resolu prend `call`, sans exception.

- **2026-08-29** — ⚠️ L'installateur natif n'ajoute PAS `.local\bin` au PATH, et le dit :
  « Native installation exists but ...\.local\bin is not in your PATH ». Le `.bat` s'en sortait
  pour SA session en le prefixant, mais l'utilisateur qui ouvrait ensuite un terminal et tapait
  `claude` ne trouvait rien.
  ➜ Ajout au PATH **utilisateur**, par le registre via PowerShell — **jamais par `setx`**, qui
  tronque a 1024 caracteres et detruit silencieusement un PATH long.

- **2026-08-29** — ⚠️ QUATRIEME occurrence du piege d'echappement de la section 8 dans mes
  outils, plus une cinquieme d'un genre voisin. Un heredoc bash a de nouveau mange les
  antislash d'un `replace("\r\n","\n")`, produisant un fichier Python syntaxiquement invalide.
  Et un script de correction a echoue parce qu'il lisait le `.bat` en CRLF tout en cherchant
  des motifs en LF — **zero occurrence trouvee**.
  ⚠️ Les deux ont echoue BRUYAMMENT, donc sans dommage : rien n'a ete ecrit. C'est la
  difference avec le faux zero du matin, ou le meme accident avait produit un resultat
  silencieux et faux.
  ➜ Regle ajoutee a mes propres passes : lire un fichier CRLF en normalisant d'abord, et faire
  transiter tout texte a antislash par un FICHIER, jamais par un heredoc.

- **2026-08-29** — ✅ LA CHAINE COMPLETE D'INSTALLATION EST DESORMAIS PROUVEE.
  Les entrees precedentes de ce journal disaient, a juste titre au moment ou elles ont ete
  ecrites, que les installations reelles n'avaient jamais tourne. **Ce n'est plus le cas.**
  Deroulement observe sur une machine Windows 11 neuve, de bout en bout :
  · Claude Code par curl (voie 1), v2.1.251, sans UAC · Python 3.13.15 par winget
  `--scope user`, telecharge depuis python.org, hachage verifie, **sans UAC** · pywin32 par pip
  · Git 2.55.0.3 par winget · `[3/6]` Outlook · `[4/6]` `Claude Code is up to date` ·
  `[5/6]` attente MAPI.
  ⚠️ Points qui restaient incertains et qui sont maintenant **mesures** :
  · `--scope user` sur le paquet Python de winget **fonctionne** — je pensais pouvoir me tromper ;
  · les deux rattrapages de PATH trouvent bien leur cible :
  `%LOCALAPPDATA%\Programs\Python\Python313\python.exe` et `%ProgramFiles%\Git\cmd\git.exe` ;
  · le message d'accord `msstore` passe tout seul grace a `--accept-source-agreements`.

- **2026-08-29** — 🔴 UNE PROMESSE QUE LE POSTE NE TENAIT PAS : "Aucune elevation UAC".
  Le `.bat` l'affichait pour TOUT le lot. Mesure sur la meme machine : Git for Windows imprime
  **"Le programme d'installation demande a s'executer en tant qu'administrateur"**.
  `winget search Git.Git` ne rend qu'un seul paquet, sans variante par compte : l'installation
  est forcement machine, donc UAC. La promesse etait donc fausse, et elle l'etait sur l'ecran
  ou l'utilisateur donne son accord.
  ⚠️ C'est exactement ce que le reflexe n°4 de la posture metier interdit : **ne rien promettre
  qu'on ne tienne**. Le poste se l'appliquait aux tarifs et aux delais, pas a lui-meme.
  ➜ Corrige dans le `.bat` ET dans les trois documents qui repetaient la promesse
  (`CLAUDE.md` §0, `README.md`, `INSTALLATION.md`) : Claude Code, Python et pywin32 s'installent
  sans elevation ; Git, non, et l'invite **peut etre refusee** puisqu'il est optionnel.
  Git installe quand meme sur la machine d'essai, l'invite ayant ete acceptee.

- **2026-08-29** — ✅ LES SIX ETAPES SONT ALLEES JUSQU'AU BOUT. `[5/6]` puis
  `[6/6] Demarrage : Opus 5 contexte 1M, effort MAX, autonomie complete`, suivi de
  `Welcome to Claude Code v2.1.251` et du choix de theme au premier lancement.
  **Sur une machine qui, dix minutes plus tot, n'avait ni Python, ni Git, ni Claude Code.**
  Le double-clic unique promis par le README est donc tenu, mesure a l'appui.
  ⚠️ Nuance a ne pas gommer : la capture ne permet pas de distinguer si l'attente MAPI a
  REUSSI (`Outlook repond. Boite accessible.`) ou si elle a EXPIRE au bout des ~40 s avant de
  demarrer quand meme. Les deux chemins menent a `[6/6]`. Le demarrage de la session est
  prouve ; la reponse effective de MAPI sur cette boite-la ne l'est pas encore.

- **2026-08-29** — Le MANUEL avait cinq trous, mesures et combles.
  Controle de `INSTALLATION.md` contre ce qui s'est REELLEMENT passe sur la machine d'essai :
  · `bypassPermissions` **jamais explique** — la seule occurrence de "permission" etait dans la
  table des fichiers, alors que le README en a une section entiere. C'est pourtant le manuel
  qu'on suit AVANT de lancer ; · le **raccourci Microsoft Store pour Python** absent — la seule
  mention du Store visait Outlook, or c'est ce raccourci qui a bloque l'installation ;
  · **rien apres `[6/6]`** : zero occurrence de theme, connexion ou navigateur, alors que Claude
  Code affiche DEUX ecrans avant la session, et que c'est la que bloque un plan gratuit ;
  · **aucun renvoi vers `depannage.md`** depuis la section installation, et `depannage.md` ne
  couvre que la boite, pas l'installation ; · le **telechargement GitHub** et `.gitattributes`
  presents dans le README, absents du manuel.
  ➜ Les cinq combles, plus une estimation de duree (~10 min : Python 28 Mo, Git 62 Mo) et un
  tableau des causes d'echec d'installation. La verification de Python passe de
  `python --version` a `python -c "print(84)"`, qui est le seul test qui ne ment pas.
  ⚠️ Le README avait les deux memes trous cote pieges : le raccourci Store et l'apres-demarrage.
  Combles aussi. Et son titre annoncait "deux gestes" pour un bloc qui en listait trois.

- **2026-08-29** — Versions anglaises : `README.en.md` et `.claude\INSTALLATION.en.md`.
  🔴 **Une traduction est un DOUBLON, et le §7 dit ce qu'il advient des doublons.** On ne peut
  pas l'eviter ici — le depot est public et l'anglais elargit sa portee — donc on la rend
  gerable : chacun des quatre fichiers porte en tete un renvoi vers l'autre langue et la
  mention **"la version francaise fait foi"**. En cas de divergence, on sait laquelle corriger
  en premier, ce qui est exactement ce que le §7 reproche aux doublons de rendre impossible.
  ⚠️ Les deux fichiers anglais signalent ce que la traduction ne peut pas resoudre : le `.bat`
  ne parle que francais (`O` = oui), `CLAUDE.md` et les gabarits sont en francais, et
  `depannage.md` n'est pas traduit. Mieux vaut le dire que laisser un anglophone le decouvrir
  devant l'invite `Installer maintenant ? [O/N]`.
  Carte du hub mise a jour : elle porte desormais les quatre fichiers et la regle de preseance.

- **2026-08-29** — 🔴 UN CHEMIN AVEC UNE PARENTHESE TUAIT LE LANCEUR. Trouve par l'utilisateur.
  Symptome rapporte : "le .bat s'est ouvert et ferme automatiquement", sans rien afficher.
  Le dossier etait `Gestionnaire-IA-main (1)` — le nom que Windows donne au SECOND
  telechargement d'un meme ZIP. Cas donc tres courant.
  Reproduit a l'identique : lance depuis ce dossier, le `.bat` rend
  `\.claude\CLAUDE.md etait inattendu.` et sort en **255**.
  ⚠️ Cause : `cmd` developpe `%VAR%` a l'ANALYSE, avant d'executer. La parenthese fermante de
  `(1)` terminait le bloc `if not exist (` des la ligne 20 — sur un `echo` qui n'aurait JAMAIS
  du s'executer, la condition etant fausse. Tout le fichier devenait syntaxiquement faux.
  ➜ Dix expansions converties en `!VAR!`, developpee a l'execution. Deux exceptions gardees :
  `endlocal & exit /b %CODE%` (que `!...!` viderait) et tout ce qui est entre guillemets.
  `%~f1` dans un bloc avait le meme defaut : range dans une variable d'abord.
  Verifie APRES correction, depuis le meme dossier a parentheses : inventaire complet, **sortie
  0**. Et les deux hooks PowerShell testes depuis ce chemin : sortie 0, JSON correct — seul
  `cmd` avait ce defaut.
  ⚠️ Ce que ca dit de la campagne d'hier : cinq agents, deux passes, et personne n'a essaye un
  chemin contenant une parenthese. Agent 1 avait pourtant teste l'espace, l'apostrophe, l'accent
  et meme la racine d'un lecteur. La parenthese manquait a la liste — et c'est le seul de ces
  cas que Windows fabrique TOUT SEUL, sans que l'utilisateur ait rien demande.

- **2026-08-29** — ✅ Correctif confirme sur un DEUXIEME poste reel, celui du proprietaire.
  Lancement du `.bat` corrige : "cela a fonctionne du premier coup", et **il a ouvert Outlook**.
  Le poste avait deja tout l'outillage, donc le chemin parcouru est l'autre : inventaire tout au
  vert, aucune question posee, puis les etapes 3 a 6 jusqu'a la session.
  ⚠️ Detail qui compte : c'est la PREMIERE fois que la branche `start "" "%OUTLOOK_EXE%"` est
  exercee. Sur le poste d'essai precedent, Outlook tournait deja et le `.bat` affichait
  "Outlook est deja lance" — la detection des six chemins d'installation d'Office, elle, n'avait
  jamais servi a lancer quoi que ce soit. Elle sert, et elle trouve.
  ⚠️ **A NE PAS SUR-INTERPRETER.** Ce qui est etabli : le lanceur va de bout en bout sur un
  poste deja equipe. Ce qui ne l'est PAS, faute d'avoir ete rapporte : laquelle des deux sorties
  de l'etape `[5/6]` s'est affichee — `Outlook repond. Boite accessible.` (MAPI a repondu) ou
  `Outlook met plus de temps que prevu` (l'attente a expire et le poste a demarre quand meme).
  **Les deux menent a `[6/6]`**, et un succes visible ne distingue pas les deux. La reponse
  effective de MAPI reste donc **non mesuree**, comme la veille.
