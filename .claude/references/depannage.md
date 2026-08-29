# Dépannage — Outlook ne synchronise plus

Procédure établie sur un cas réel : une boîte Microsoft 365 restée **quatre mois
sans télécharger un seul message**, alors qu'Outlook s'ouvrait normalement et
que le courrier arrivait bien sur le serveur.

Le fil conducteur : **mesurer avant de réparer**. La plupart du temps perdu sur
ce genre de panne vient de gestes correctifs appliqués sans diagnostic.

---

## 1. Le diagnostic en une commande

```bash
fsutil file queryvaliddata "%LOCALAPPDATA%\Microsoft\Outlook\<compte>.ost"
```

La **VDL** (valid data length) est le nombre d'octets réellement écrits dans le
fichier, indépendamment de sa taille allouée. C'est l'indicateur le plus honnête
qui existe : un OST peut afficher 16 Mo tout en n'ayant jamais rien reçu.

| VDL | Interprétation |
|---|---|
| ~280 Ko | coquille vide — la structure des dossiers existe, **aucun message** |
| quelques Mo | synchronisation amorcée puis interrompue |
| centaines de Mo à plusieurs Go | synchronisation normale |

Relancer la commande à une minute d'intervalle : si la VDL grimpe, la
synchronisation est en cours et il n'y a **rien à réparer**, juste à attendre.

## 2. Séparer « pas de courrier » de « pas de synchronisation »

Avant de toucher au client, vérifier le serveur : ouvrir la boîte dans un
navigateur (`outlook.office.com` pour Microsoft 365). Si les messages y sont
mais pas dans Outlook, le problème est **local** — le serveur est hors de cause.

Cette vérification coûte trente secondes et évite d'aller chercher un problème
d'hébergement qui n'existe pas.

## 3. Les causes, de la plus fréquente à la plus rare

### a) La boîte a été déplacée côté serveur — cause la plus sournoise

**Symptôme** : au démarrage, Outlook affiche

> « Votre boîte aux lettres a été temporairement déplacée sur Microsoft
> Exchange. Une boîte aux lettres temporaire est à votre disposition… »

avec le choix entre **« Utiliser la boîte aux lettres temporaire »** et
**« Utiliser les anciennes données »**.

**Le piège** : « Utiliser les anciennes données » place Outlook **hors ligne de
façon permanente** — le texte le dit d'ailleurs, *« vous ne pourrez ni envoyer,
ni recevoir de courriers »*, mais la phrase se lit vite. Le choix est mémorisé.
La boîte ne se remplira plus jamais, sans aucun message d'erreur ensuite.

**La réponse est « Utiliser la boîte aux lettres temporaire »** — le mot
« temporaire » désigne la boîte à son nouvel emplacement, pas un cache jetable.

Si la fenêtre **revient à chaque démarrage**, le profil contient un pointeur
périmé : passer à l'étape 4.

### b) Outlook démarre en mode sans échec

**Symptôme** : « (Mode sans échec) » dans la barre de titre, et le bouton
**Réparer** grisé dans les paramètres du compte.

Le mode sans échec bride le courtier d'authentification : la connexion moderne
n'aboutit pas et la synchronisation ne démarre jamais. **Toujours répondre
« Non » à la proposition de mode sans échec** — c'est lui qui empêche la
réparation, pas l'inverse.

Vérifier si Outlook y est entré seul plutôt que par un `/safe` explicite :

```powershell
Get-CimInstance Win32_Process -Filter "Name='OUTLOOK.EXE'" |
  Select-Object ProcessId, CommandLine
```

Une ligne de commande sans `/safe` signifie qu'Outlook a planté et que Windows
a proposé le mode dégradé.

### c) Plantage du fournisseur Exchange

```powershell
Get-WinEvent -FilterHashtable @{LogName='Application'; Id=1000} |
  Where-Object { $_.Message -match 'OUTLOOK' } |
  Select-Object -First 3 TimeCreated, Message
```

Lire le **module défaillant** :

- **`EMSMDB32.DLL`** → fournisseur MAPI Exchange. Ce n'est **pas** un complément
  tiers : inutile de désactiver les add-ins. Presque toujours un profil périmé
  → étape 4.
- **`OUTLLIB.DLL`**, `MSPST32.DLL` → cœur d'Outlook ou magasin de données.
- Une DLL portant un nom d'éditeur tiers → là, oui, un complément est en cause :
  démarrer avec `outlook.exe /safe`, puis désactiver les compléments un à un.

### d) Aucun jeton d'authentification

```powershell
cmdkey /list | Select-String "OAUTH2|MicrosoftOffice"
```

Si aucun jeton ne correspond au domaine du compte configuré, la connexion n'a
jamais abouti. Souvent accompagné de jetons orphelins d'anciens comptes, qui
brouillent l'invite de connexion.

---

## 4. La réparation : recréer le profil

**C'est le remède aux cas (a), (b) et (c-EMSMDB32).** Le profil Outlook est un
petit objet du registre qui mémorise l'emplacement de la boîte ; quand cette
information est périmée, aucun redémarrage ne la corrigera — elle est relue à
chaque lancement.

> **Réinstaller Office ne sert à rien ici.** Le profil vit dans
> `HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles` et **survit intacte** à
> une désinstallation. On y perd une heure pour retomber sur la même boucle.

> **Supprimer l'OST ne sert à rien non plus** si la VDL montre qu'il est vide :
> il n'y a rien à corrompre dans un fichier qui n'a jamais rien reçu.

### Sauvegarder d'abord

```powershell
reg export "HKCU\Software\Microsoft\Office\16.0\Outlook" "$env:USERPROFILE\Desktop\profil_outlook_backup.reg" /y
```

Un double-clic sur ce fichier restaure l'état d'origine. **Ne pas continuer si
cette commande échoue.**

### Voie graphique (recommandée)

1. Fermer Outlook.
2. Panneau de configuration → **Courrier (Microsoft Outlook)**
   *(affichage en « Petites icônes » s'il n'apparaît pas)*
3. **Afficher les profils… → Ajouter…**, nommer le nouveau profil.
4. Saisir **une seule** adresse — celle qui pose problème. Les autres comptes se
   rajoutent une fois le fonctionnement confirmé.
5. Cocher **« Toujours utiliser ce profil »**, démarrer Outlook.

### Voie scriptée

Outlook fermé. Purger le cache d'autodécouverte, qui mémorise l'emplacement de
la boîte et la fait revenir obstinément au mauvais endroit :

```powershell
# Sauvegarde
$bk = "$env:USERPROFILE\Desktop\outlook_backup"
New-Item -ItemType Directory -Force $bk | Out-Null
Copy-Item "$env:LOCALAPPDATA\Microsoft\Outlook\*Autodiscover.xml" $bk -EA SilentlyContinue
reg export "HKCU\Software\Microsoft\Office\16.0\Outlook" "$bk\profil.reg" /y

# Purge
Remove-Item "$env:LOCALAPPDATA\Microsoft\Outlook\*Autodiscover.xml" -Force -EA SilentlyContinue
Remove-Item "HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover" -Recurse -Force -EA SilentlyContinue
Remove-Item "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook" -Recurse -Force -EA SilentlyContinue
Remove-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name DefaultProfile -Force -EA SilentlyContinue
```

Sans profil, Outlook lance l'assistant de configuration au démarrage suivant :
il interroge l'autodécouverte à zéro et trouve la boîte à son emplacement réel.

> Adapter `16.0` à la version installée (Microsoft 365 et Office 2016 à 2024
> utilisent tous `16.0`).

### Vérifier

```bash
fsutil file queryvaliddata "%LOCALAPPDATA%\Microsoft\Outlook\<nouveau>.ost"
```

La VDL doit grimper de façon continue. La barre d'état d'Outlook affiche
« Connecté à Microsoft Exchange » et le volume restant à télécharger.

---

## 5. Ce qu'il ne faut pas faire

| Geste réflexe | Pourquoi c'est inutile ici |
|---|---|
| Supprimer l'OST | Sans mesure préalable, on détruit un cache sain pour rien |
| Désactiver les compléments | Sans effet si le module fautif est `EMSMDB32.DLL` |
| Réinstaller Office | Le profil défectueux survit à la réinstallation |
| Accepter le mode sans échec | C'est précisément ce qui bloque la reconnexion |
| Changer le mot de passe | Le compte s'authentifie déjà : ce n'est pas la cause |

---

## 6. Analyse hors ligne du fichier OST

`scripts/ost_reader.py` lit un `.ost` **sans Outlook** : dossiers, messages,
expéditeurs, dates. Utile pour prouver qu'un cache est vide, inventorier une
boîte sur un poste où Outlook ne démarre plus, ou auditer un fichier récupéré.

```bash
python .claude/scripts/ost_reader.py "C:\chemin\vers\fichier.ost"
```

Points à connaître sur le format, vérifiés en pratique :

- OST « large » 4 Ko (Outlook 2013+) : `wVer = 36`, pages de 4096 octets,
  `PAGETRAILER` à **+0xFE8** et non en fin de page, `cEnt`/`cEntMax` sur
  **16 bits**. Le format 512 octets décrit dans MS-PST ne s'applique pas.
- **Aucun chiffrement** dans le cas courant (`bCryptMethod = 0x00`) : le contenu
  se lit en clair. Même chiffré, les modes disponibles sont de la substitution
  d'octets à table publique — de l'obscurcissement, pas de la cryptographie.
  **Un fichier OST n'est protégé que par les permissions du système de fichiers
  et le chiffrement de disque.**
- Certains blocs volumineux sont un flux **zlib brut** (`78 9C`) : sans
  décompression, leur signature paraît invalide.
- Outlook pose un verrou sur les **1024 premiers octets** tant qu'il tourne :
  `Copy-Item` et `esentutl /y` échouent tous deux. Copier bloc par bloc en
  sautant la plage verrouillée, ou fermer Outlook.
