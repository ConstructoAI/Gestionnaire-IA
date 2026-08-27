# ETAT — les ENGAGEMENTS

> **Satellite — ne se charge PAS automatiquement**, et c'est voulu : il ne coute rien
> tant qu'on ne l'ouvre pas. Le hub `CLAUDE.md` porte les regles ; ce fichier porte
> **le journal des engagements** : ce qui a ete promis, a qui, pour quand.
>
> 🔴 **Une section vide se lit « pas encore consigné », JAMAIS « rien ne s'est passé ».**
> Un gabarit vide bien formé a l'air complet : c'est un faux zéro. Ne pas le remplir
> artificiellement — l'alimenter au fil des passes.
>
> **Ne duplique rien.** 🔴 **Aucun etat de boite ici** — ni liste de messages, ni fils en souffrance, ni
> compteurs. Outlook est deja le magasin et il est a jour ; le recopier creerait deux
> versions d'une meme verite.

> Cree le : *(a dater)*. Derniere mise a jour : *(a dater)*.

---

## Pourquoi ce fichier existe

Une detection mecanique de « fil sans reponse » **ne peut pas** produire un engagement :
repondre « je vous envoie le prix vendredi » sort le fil de la liste, alors que la promesse
court toujours. Seule une lecture des **Elements envoyes** la voit.

---

## 1. Engagements ouverts

**Un engagement ouvert ne s'elague jamais.** *(vide — a alimenter)*

| Pris le | Envers | Engagement | Echeance | Source (EntryID) | Statut |
|---|---|---|---|---|---|

Statuts : `ouvert` · `tenu` (avec la date) · `caduc` (avec la raison).

**Ce qui compte comme engagement** — une phrase qui cree une attente chez l'autre : un prix ou
un chiffre annonce (ils engagent, meme verbalement), une date, un document promis, une decision
annoncee comme a venir (« je vous reviens la-dessus »).
**Ce qui n'en est pas** : un accuse de reception, une politesse, une hypothese au conditionnel.

---

## 2. Comment on l'alimente

A chaque passe de courriels, deux gestes en plus du tri : **relire les Elements envoyes**
depuis la derniere passe pour en extraire les promesses — c'est la qu'elles vivent, jamais
dans la reception — et **fermer ce qui est tenu**.

⚠️ **Un brouillon n'est pas un engagement.** Tant qu'il n'est pas envoye, rien n'a ete promis.
⚠️ **Ne rien inventer** : lire le montant et l'echeance dans le message envoye, avec son
`EntryID` en source.

---

## 3. Engagements clos

`tenu` reste 90 jours en clair puis passe en une ligne. `caduc` s'efface apres 90 jours, sauf
si sa raison eclaire une decision — elle migre alors dans `ETAT_projets.md` §3. *(vide)*
