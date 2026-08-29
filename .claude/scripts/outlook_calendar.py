# -*- coding: utf-8 -*-
"""Calendrier Outlook : lire ET ecrire. L'ecriture est verrouillee par --yes-write.

    python scripts/outlook_calendar.py calendriers
    python scripts/outlook_calendar.py list   [--cal "<NOM DU CALENDRIER>"] [--jours 30]
    python scripts/outlook_calendar.py show   --id <EntryID>
    python scripts/outlook_calendar.py create --sujet "..." --debut "2026-09-03 08:00" --duree 480 \
                                              [--cal "..."] [--lieu "..."] [--corps "..."] [--jour-entier] --yes-write
    python scripts/outlook_calendar.py update --id <EntryID> [--sujet ...] [--debut ...] [--duree ...] \
                                              [--lieu ...] [--corps ...] --yes-write
    python scripts/outlook_calendar.py delete --id <EntryID> --yes-write

LE VERROU, ET POURQUOI IL EXISTE
    `--yes-write` est calque sur le `--yes-send` de outlook_mail.py. Sans lui, toute
    commande qui MODIFIE refuse de s'executer et sort en code 1. Il ne se pose que
    lorsque le proprietaire du poste a approuve CETTE entree-la : un accord donne pour un rendez-vous
    ne vaut pas pour le suivant.

    Ce n'est pas une precaution decorative : un calendrier partage porte souvent
    l'echeancier REEL du travail, et une entree fautive s'y voit et se propage.

CE QUE `delete` FAIT VRAIMENT
    Il DEPLACE vers les Elements supprimes — jamais de purge definitive. Un
    rendez-vous efface par erreur se retrouve. La suppression definitive n'est pas
    exposee, et elle ne doit pas l'etre.

QUATRE PIEGES DEJA PAYES, respectes ici
    1. DATES EN ISO, toujours. Le `/` d'un format .NET est le separateur de la
       CULTURE, pas une barre litterale ; ce poste est en fr-CA. `MM/dd/yyyy` y rend
       `09-01-2026`. Mesure du 2026-08-27 : une fenetre lue au format francais a rendu
       50+ rendez-vous au lieu de 4.
    2. `IncludeRecurrences = True` PUIS `Sort("[Start]")`, dans cet ordre — sinon une
       serie ne rend qu'UNE occurrence. Le calendrier principal passe de 11 a 22.
    3. JAMAIS `.Count` apres `Restrict` : il rend 2147483647. On itere, avec une borne.
    4. Filtrer sur `Class == 26` (olAppointment) : un dossier de calendrier peut porter
       autre chose, et les proprietes de rendez-vous levent alors.
    5. Console en cp1252 : sans `reconfigure`, un accent tue le script.

ECRIRE DANS UNE SERIE RECURRENTE
    `update` et `delete` agissent sur l'element STOCKE, donc sur toute la serie si
    l'EntryID en designe une. Modifier une seule occurrence n'est PAS expose : c'est
    une operation a faire dans Outlook, ou l'on voit ce qu'on touche.
"""
import argparse
import json
import sys
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")

OL_CALENDRIER, OL_SUPPRIMES = 9, 3
CLASS_RDV = 26
CREATE_RDV = 1


def _win32():
    try:
        import win32com.client as w
        return w
    except ImportError:
        sys.exit("ARRET : pywin32 manque. `pip install pywin32`, puis relancer.")


def _ns():
    w = _win32()
    try:
        return w.Dispatch("Outlook.Application"), w.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        sys.exit("ARRET : MAPI ne repond pas (%s). Outlook est-il lance ?" % e)


def _iso(d):
    try:
        return d.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return "?"


def _lire_date(s):
    """N'accepte QUE l'ISO. Refuser le reste est volontaire : voir piege 1."""
    for f in ("%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    sys.exit("ARRET : date '%s' non reconnue. Format attendu : AAAA-MM-JJ [HH:MM] (ISO, "
             "seule forme insensible a la culture)." % s)


def _dossiers(ns):
    cal = ns.GetDefaultFolder(OL_CALENDRIER)
    out = [(cal.Name, cal)]
    for s in cal.Folders:
        if getattr(s, "DefaultItemType", None) == 1:
            out.append((s.Name, s))
    return out


def _cible(ns, nom):
    d = _dossiers(ns)
    if not nom:
        return d[0][1]
    for n, f in d:
        if n.lower() == nom.lower():
            return f
    for n, f in d:
        if nom.lower() in n.lower():
            return f
    sys.exit("ARRET : calendrier '%s' introuvable. Disponibles : %s"
             % (nom, " | ".join(n for n, _ in d)))


def _verrou(a, quoi):
    if not getattr(a, "yes_write", False):
        print(json.dumps(dict(status="refus", raison="l'ecriture exige --yes-write",
                              operation=quoi), ensure_ascii=False))
        sys.exit(1)


def cmd_calendriers(a):
    _, ns = _ns()
    for n, f in _dossiers(ns):
        print("  %-42s %5d element(s)" % (n, f.Items.Count))


def cmd_list(a):
    _, ns = _ns()
    f = _cible(ns, a.cal)
    items = f.Items
    items.IncludeRecurrences = True          # AVANT le tri — piege 2
    items.Sort("[Start]")
    # Corrige le 2026-08-29. Les deux premieres affectations de `fin` etaient
    # mortes — ecrasees par la troisieme — et `--jours 0` produisait une fenetre
    # [maintenant, maintenant] : 0 rendez-vous, code 0, aucun avertissement.
    if a.jours < 0:
        sys.exit("ARRET : --jours doit valoir 0 ou plus (0 = aujourd'hui jusqu'a 23:59).")
    debut = datetime.now()
    fin = (debut.replace(hour=23, minute=59, second=59) if a.jours == 0
           else debut + timedelta(days=a.jours))
    res = items.Restrict("[Start] >= '%s' AND [Start] <= '%s'"
                         % (debut.strftime("%Y-%m-%d %H:%M"), fin.strftime("%Y-%m-%d %H:%M")))
    out, n = [], 0
    for x in res:                            # on ITERE — jamais .Count, piege 3
        if n >= a.limite:
            break
        try:
            if x.Class != CLASS_RDV:         # piege 4
                continue
            out.append(dict(id=x.EntryID, sujet=x.Subject or "", debut=_iso(x.Start),
                            fin=_iso(x.End), lieu=x.Location or "",
                            recurrent=bool(x.IsRecurring), jour_entier=bool(x.AllDayEvent),
                            calendrier=f.Name))
            n += 1
        except Exception:
            continue
    if a.json:
        print(json.dumps(out, ensure_ascii=False, indent=1))
    else:
        print("  %d rendez-vous sur %d jours — %s" % (len(out), a.jours, f.Name))
        for x in out:
            r = " [serie]" if x["recurrent"] else ""
            l = ("  @ " + x["lieu"]) if x["lieu"] else ""
            print("  %s  %s%s%s" % (x["debut"], x["sujet"], l, r))


def cmd_show(a):
    _, ns = _ns()
    x = ns.GetItemFromID(a.id)
    print(json.dumps(dict(id=x.EntryID, sujet=x.Subject or "", debut=_iso(x.Start),
                          fin=_iso(x.End), duree_min=x.Duration, lieu=x.Location or "",
                          organisateur=x.Organizer or "", recurrent=bool(x.IsRecurring),
                          jour_entier=bool(x.AllDayEvent),
                          modifie=_iso(x.LastModificationTime),
                          calendrier=x.Parent.Name, corps=(x.Body or "")[:400]),
                     ensure_ascii=False, indent=1))


def cmd_create(a):
    _verrou(a, "create")
    ol, ns = _ns()
    f = _cible(ns, a.cal)
    x = f.Items.Add(CREATE_RDV)
    x.Subject = a.sujet
    x.Start = _lire_date(a.debut)
    if a.jour_entier:
        x.AllDayEvent = True
    else:
        x.Duration = a.duree
    if a.lieu:
        x.Location = a.lieu
    if a.corps:
        x.Body = a.corps
    x.Save()
    print(json.dumps(dict(status="rendez-vous cree", id=x.EntryID, sujet=x.Subject,
                          debut=_iso(x.Start), fin=_iso(x.End), calendrier=f.Name),
                     ensure_ascii=False))


def cmd_update(a):
    _verrou(a, "update")
    _, ns = _ns()
    x = ns.GetItemFromID(a.id)
    avant = dict(sujet=x.Subject, debut=_iso(x.Start), duree=x.Duration, lieu=x.Location or "")
    if x.IsRecurring:
        print("[avertissement] cet element est une SERIE : la modification porte sur "
              "toute la serie, pas sur une occurrence.", file=sys.stderr)
    if a.sujet:
        x.Subject = a.sujet
    if a.debut:
        x.Start = _lire_date(a.debut)
    if a.duree:
        x.Duration = a.duree
    if a.lieu:
        x.Location = a.lieu
    if a.corps:
        x.Body = a.corps
    x.Save()
    print(json.dumps(dict(status="rendez-vous modifie", id=x.EntryID, avant=avant,
                          apres=dict(sujet=x.Subject, debut=_iso(x.Start),
                                     duree=x.Duration, lieu=x.Location or "")),
                     ensure_ascii=False, indent=1))


def cmd_delete(a):
    _verrou(a, "delete")
    _, ns = _ns()
    x = ns.GetItemFromID(a.id)
    if x.Parent.Name == ns.GetDefaultFolder(OL_SUPPRIMES).Name:
        sys.exit("ARRET : deja dans les Elements supprimes. La purge definitive n'est pas exposee.")
    info = dict(sujet=x.Subject or "", debut=_iso(x.Start), calendrier=x.Parent.Name,
                serie=bool(x.IsRecurring))
    x.Delete()                               # -> Elements supprimes, jamais definitif
    print(json.dumps(dict(status="deplace vers les Elements supprimes", **info),
                     ensure_ascii=False))


def main():
    ap = argparse.ArgumentParser(description="Calendrier Outlook. L'ecriture exige --yes-write.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def add(nom, fn, aide):
        s = sub.add_parser(nom, help=aide)
        s.set_defaults(fn=fn)
        return s

    add("calendriers", cmd_calendriers, "lister les calendriers et leur volume")

    s = add("list", cmd_list, "les rendez-vous a venir")
    s.add_argument("--cal", default=None, help="nom du calendrier (defaut : le principal)")
    s.add_argument("--jours", type=int, default=30)
    s.add_argument("--limite", type=int, default=100)
    s.add_argument("--json", action="store_true")

    s = add("show", cmd_show, "le detail d'un rendez-vous")
    s.add_argument("--id", required=True)

    s = add("create", cmd_create, "creer un rendez-vous")
    s.add_argument("--sujet", required=True)
    s.add_argument("--debut", required=True, metavar="AAAA-MM-JJ[ HH:MM]")
    s.add_argument("--duree", type=int, default=60, metavar="MINUTES")
    s.add_argument("--cal", default=None)
    s.add_argument("--lieu", default=None)
    s.add_argument("--corps", default=None)
    s.add_argument("--jour-entier", action="store_true", dest="jour_entier")
    s.add_argument("--yes-write", action="store_true", dest="yes_write")

    s = add("update", cmd_update, "modifier un rendez-vous")
    s.add_argument("--id", required=True)
    s.add_argument("--sujet", default=None)
    s.add_argument("--debut", default=None, metavar="AAAA-MM-JJ[ HH:MM]")
    s.add_argument("--duree", type=int, default=None)
    s.add_argument("--lieu", default=None)
    s.add_argument("--corps", default=None)
    s.add_argument("--yes-write", action="store_true", dest="yes_write")

    s = add("delete", cmd_delete, "deplacer un rendez-vous vers les Elements supprimes")
    s.add_argument("--id", required=True)
    s.add_argument("--yes-write", action="store_true", dest="yes_write")

    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
