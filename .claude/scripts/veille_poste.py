# -*- coding: utf-8 -*-
"""Veille sur TROIS surfaces : courriels ENTRANTS, courriels SORTANTS, CALENDRIER.

    python scripts/veille_poste.py                    # une passe : ce qui a change depuis la derniere fois
    python scripts/veille_poste.py --boucle           # veille continue (defaut 480 min, tour de 60 s)
    python scripts/veille_poste.py --boucle --minutes 120 --intervalle 30
    python scripts/veille_poste.py --remise-a-zero    # oublie l'etat et reprend au present

POURQUOI CE SCRIPT PLUTOT QUE veille_entrants.py
    Celui-la ne voit que les ENTRANTS. Ici on suit aussi ce qui PART (une autre
    session, une regle Outlook, ou l'utilisateur lui-meme peuvent envoyer) et ce qui
    bouge au CALENDRIER, y compris l'echeancier de chantiers.

LE .ost DECLENCHE LA MESURE, IL NE LA REMPLACE PAS
    Sa date de modification bouge des dizaines de fois par heure pour des raisons
    qui ne sont pas du courrier : un message marque lu, un cache reindexe. On s'en
    sert comme d'un signal bon marche (un `stat` contre une connexion COM), puis on
    VERIFIE par MAPI. Regle heritee de veille_entrants.py, et elle tient.

CE QUI EST COMPARE, SURFACE PAR SURFACE
    Entrants  : EntryID nouveaux dans la Boite de reception
    Sortants  : EntryID nouveaux dans les Elements envoyes
    Calendrier: (Subject, Start, LastModificationTime) — un rendez-vous MODIFIE se
                voit, la ou un simple compte d'elements ne verrait rien.

TROIS PIEGES DEJA PAYES, respectes ici
    1. Console Windows en cp1252 : sans reconfigure, un accent combinant tue le
       script. Mesure du 2026-08-27 sur nos propres titres de fichiers.
    2. Dates en ISO. Le `/` d'un format .NET est le separateur de la CULTURE ; ce
       poste est en fr-CA. On n'ecrit jamais de borne a la main ici.
    3. Pas de .Count apres Restrict : on itere. Restrict + IncludeRecurrences rend
       2147483647.
    4. line_buffering : hors terminal, Python retient sa sortie. Une veille de huit
       heures resterait MUETTE et on la croirait morte.

LECTURE SEULE ABSOLUE. Ce script n'ecrit rien dans Outlook : ni marquage, ni
deplacement, ni envoi, ni rendez-vous. Il observe et il rapporte.
"""
import argparse
import glob
import io
import json
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)

ICI = os.path.dirname(os.path.abspath(__file__))
RACINE = os.path.dirname(ICI)
ETAT = os.path.join(RACINE, ".veille_poste_state.json")
OST = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Microsoft", "Outlook", "*.ost")

OL_ENVOYES, OL_RECEPTION, OL_CALENDRIER = 5, 6, 9


def _win32():
    try:
        import win32com.client as w
        return w
    except ImportError:
        sys.exit("ARRET : pywin32 manque. Installer avec `pip install pywin32`, puis relancer.")


def empreinte_ost():
    e = {}
    for p in glob.glob(OST):
        try:
            e[p] = os.path.getmtime(p)
        except OSError:
            pass
    return e


def _iso(d):
    try:
        return d.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return "?"


def releve(ns, limite=60):
    """Photographie des trois surfaces. Rend None si MAPI ne repond pas."""
    try:
        snap = {"entrants": {}, "sortants": {}, "calendrier": {}}

        for cle, dossier in (("entrants", OL_RECEPTION), ("sortants", OL_ENVOYES)):
            f = ns.GetDefaultFolder(dossier)
            items = f.Items
            items.Sort("[ReceivedTime]" if cle == "entrants" else "[SentOn]", True)
            n = 0
            for m in items:
                if n >= limite:
                    break
                try:
                    if m.Class != 43:          # 43 = olMail ; un dossier peut porter autre chose
                        continue
                    quand = m.ReceivedTime if cle == "entrants" else m.SentOn
                    qui = (m.SenderName if cle == "entrants" else m.To) or ""
                    snap[cle][m.EntryID] = {"quand": _iso(quand), "objet": (m.Subject or "")[:70],
                                            "qui": str(qui)[:50]}
                    n += 1
                except Exception:
                    continue

        cal = ns.GetDefaultFolder(OL_CALENDRIER)
        for dossier, nom in [(cal, cal.Name)] + [(s, s.Name) for s in cal.Folders]:
            if getattr(dossier, "DefaultItemType", None) != 1:
                continue
            if nom in ("United States holidays", "Anniversaires"):
                continue                      # feries et anniversaires : du bruit ici
            for x in dossier.Items:           # PAS de Restrict : on veut les elements STOCKES
                try:
                    if x.Class != 26:          # 26 = olAppointment
                        continue
                    snap["calendrier"][x.EntryID] = {
                        "cal": nom, "objet": (x.Subject or "")[:70],
                        "debut": _iso(x.Start), "modifie": _iso(x.LastModificationTime)}
                except Exception:
                    continue
        return snap
    except Exception as e:
        print("[mesure ratee] %s" % e)
        return None


def charger():
    try:
        return json.load(io.open(ETAT, encoding="utf-8"))
    except Exception:
        return None


def sauver(snap):
    try:
        with io.open(ETAT, "w", encoding="utf-8") as f:
            json.dump(snap, f, ensure_ascii=False)
    except OSError as e:
        print("[avertissement] etat non sauvegarde : %s" % e)


def comparer(avant, apres):
    """Rend (entrants, sortants, cal_nouveaux, cal_modifies)."""
    ent = [v for k, v in apres["entrants"].items() if k not in avant["entrants"]]
    sor = [v for k, v in apres["sortants"].items() if k not in avant["sortants"]]
    neufs, modif = [], []
    for k, v in apres["calendrier"].items():
        a = avant["calendrier"].get(k)
        if a is None:
            neufs.append(v)
        elif a.get("modifie") != v.get("modifie"):
            modif.append((a, v))
    return (sorted(ent, key=lambda m: m["quand"]),
            sorted(sor, key=lambda m: m["quand"]),
            sorted(neufs, key=lambda m: m["debut"]), modif)


def rapporter(ent, sor, neufs, modif):
    rien = True
    if ent:
        rien = False
        print("\n=== %d COURRIEL(S) ENTRANT(S) ===" % len(ent))
        for m in ent:
            print("  %s  de %-32s  %s" % (m["quand"], m["qui"], m["objet"]))
    if sor:
        rien = False
        print("\n=== %d COURRIEL(S) SORTANT(S) ===" % len(sor))
        for m in sor:
            print("  %s  a  %-32s  %s" % (m["quand"], m["qui"], m["objet"]))
    if neufs:
        rien = False
        print("\n=== %d RENDEZ-VOUS AJOUTE(S) ===" % len(neufs))
        for x in neufs:
            print("  %s  [%s]  %s" % (x["debut"], x["cal"], x["objet"]))
    if modif:
        rien = False
        print("\n=== %d RENDEZ-VOUS MODIFIE(S) ===" % len(modif))
        for a, b in modif:
            deplace = " — DEPLACE : %s -> %s" % (a["debut"], b["debut"]) if a["debut"] != b["debut"] else ""
            print("  [%s]  %s%s" % (b["cal"], b["objet"], deplace))
    return rien


def main():
    ap = argparse.ArgumentParser(description="Veille entrants + sortants + calendrier. Lecture seule.")
    ap.add_argument("--boucle", action="store_true", help="veille continue au lieu d'une seule passe")
    ap.add_argument("--minutes", type=int, default=480, help="duree de la boucle (defaut 480)")
    ap.add_argument("--intervalle", type=int, default=60, help="secondes entre deux tours (defaut 60)")
    ap.add_argument("--remise-a-zero", action="store_true", help="oublie l'etat et reprend au present")
    a = ap.parse_args()

    w = _win32()
    try:
        ns = w.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        sys.exit("ARRET : MAPI ne repond pas (%s). Outlook est-il lance ?" % e)

    if a.remise_a_zero and os.path.exists(ETAT):
        os.remove(ETAT)
        print("etat oublie.")

    snap = releve(ns)
    if snap is None:
        sys.exit("ARRET : impossible de photographier la boite. On ne conclut RIEN.")

    avant = charger()
    if avant is None:
        sauver(snap)
        print("Premiere passe — repere pose. Rien a comparer.")
        print("  entrants suivis : %d | sortants : %d | rendez-vous : %d"
              % (len(snap["entrants"]), len(snap["sortants"]), len(snap["calendrier"])))
        if not a.boucle:
            return
        avant = snap

    if not a.boucle:
        rien = rapporter(*comparer(avant, snap))
        sauver(snap)
        if rien:
            print("Rien de neuf depuis la derniere passe.")
        return

    ost = empreinte_ost()
    print("veille demarree — %d magasin(s) surveille(s), tour de %d s, %d min"
          % (len(ost), a.intervalle, a.minutes))
    if not ost:
        print("AVERTISSEMENT : aucun .ost lisible — MAPI sera interroge a chaque tour.")

    tours = max(1, (a.minutes * 60) // a.intervalle)
    for i in range(1, tours + 1):
        time.sleep(a.intervalle)
        neuf_ost = empreinte_ost()
        bouge = (neuf_ost != ost)
        ost = neuf_ost
        # Le calendrier peut bouger sans que le .ost change : on mesure quand meme
        # tous les 5 tours, sinon un rendez-vous deplace passerait inapercu.
        if ost and not bouge and (i % 5):
            continue
        s = releve(ns)
        if s is None:
            print("[tour %d] MESURE RATEE — on ne conclut rien" % i)
            continue
        if not rapporter(*comparer(avant, s)):
            sauver(s)
        avant = s
    print("\nveille terminee (%d tours)." % tours)


if __name__ == "__main__":
    main()
