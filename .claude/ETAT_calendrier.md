# ETAT — calendrier

> **Satellite — ne se charge PAS automatiquement**, et c'est voulu : il ne coute rien
> tant qu'on ne l'ouvre pas. Le hub `CLAUDE.md` porte les regles ; ce fichier porte
> ce qui a **glisse**, ce qui **revient**, et ce qui s'est revele **faux**.
>
> 🔴 **Une section vide se lit « pas encore consigné », JAMAIS « rien ne s'est passé ».**
> Un gabarit vide bien formé a l'air complet : c'est un faux zéro. Ne pas le remplir
> artificiellement — l'alimenter au fil des passes.
>
> **Ne duplique rien.** Le calendrier Outlook est sa propre source de verite et change tous les jours :
> **ne jamais recopier de rendez-vous ici**. La recette de lecture et les six pieges sont
> dans `CLAUDE.md` §3.

> Cree le : *(a dater)*. Derniere mise a jour : *(a dater)*.

---

## 0. Les calendriers du poste

Lancer `python .claude/scripts/outlook_calendar.py calendriers` et remplir.

| Calendrier | Elements | Ce qu'on y trouve |
|---|---:|---|
| *(a relever)* | | |

⚠️ Le compte **stocke** et le compte **avec recurrences** different. Toujours dire lequel on cite.

---

## 1. Jalons qui ont glisse

**Le coeur de ce fichier.** Un jalon deplace raconte ce que le calendrier efface en le
deplacant : **la date d'origine disparait.** C'est ici qu'elle survit.

*(vide — a alimenter)*

| Chantier / dossier | Jalon | Prevu a l'origine | Reel | Glissement | Cause |
|---|---|---|---|---|---|

➜ Une **cause** vaut mieux qu'un nombre de jours : « fournisseur en rupture » se reutilise,
« +9 j » non.

---

## 2. Recurrences connues

Ce qui revient — et parce qu'une serie est precisement ce qu'une lecture sans
`IncludeRecurrences` **rate en silence**.

| Quoi | Rythme | Calendrier | Mesure le |
|---|---|---|---|
| *(a alimenter)* | | | |

⚠️ Une observation unique ne prouve pas une recurrence : lire `IsRecurring` et
`GetRecurrencePattern()`.

---

## 3. Ce qui s'est revele faux

Les corrections payees, pour ne pas les repayer. *(vide)*

---

## 4. Ecriture

L'ecriture est possible via `outlook_calendar.py`, verrouillee par **`--yes-write`**.
⚠️ Tout jalon deplace par Claude se consigne au **§1** — sinon la date d'origine est perdue
pour de bon.
