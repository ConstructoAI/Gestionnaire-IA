# -*- coding: utf-8 -*-
"""Inventaire d'un dossier de FACTURES sur disque. Lecture seule.

    python .claude/scripts/factures.py --dossier "10. ADMIN/COMPTABILITE/02. FACTURES"
    python .claude/scripts/factures.py --dossier "..." --annee 2026 --json
    python .claude/scripts/factures.py --dossier "..." --conventions

CE QUE FAIT CE SCRIPT, ET CE QU'IL NE FAIT PAS
    Il LIT un dossier de factures et rend : combien de factures distinctes, pour
    quel total, par client, et ce qu'il n'a PAS pu lire. Il n'ecrit rien, ne
    renomme rien, ne touche a aucun systeme de gestion.

POURQUOI IL EXISTE
    Un dossier de factures reel n'est jamais propre. Quatre defauts s'y trouvent
    presque toujours, et chacun produit un compte FAUX qui a l'air juste :

    1. LE MEME DOCUMENT EN .html ET .pdf, sous des noms differents.
       Compter les fichiers surestime le nombre de factures.
    2. L'HORODATAGE N'EST PAS UN IDENTIFIANT. Deux factures emises la meme minute
       a des clients differents portent le meme nombre. Deduplique par horodatage
       seul, on FUSIONNE des factures et on perd de l'argent.
    3. LE NOM DE FICHIER N'IDENTIFIE PAS LE CLIENT. Le suffixe est tantot le
       client, tantot l'emetteur. La seule source fiable est le champ
       << Facture a >> DANS le document.
    4. LA CASSE N'EST NORMALISEE NULLE PART : FACTURE / Facture / facture
       coexistent. Un motif sensible a la casse rate un tiers du dossier.

    D'ou la cle retenue : (horodatage + client lu DANS le document).

LA VALIDATION ARITHMETIQUE EST LE COEUR DU SCRIPT
    On ne retient un montant que si le quadruplet BOUCLE :
        TPS = HT x taux1, TVQ = HT x taux2, TOTAL = HT x (1 + taux1 + taux2)
    au cent pres. Sans cette regle, une extraction naive sort des aberrations du
    genre << TPS 442,65 $ sur un sous-total de 38,40 $ >> et les livre sans broncher.
    Ce qui ne boucle pas est ISOLE, jamais devine.

TAUX PAR DEFAUT : Quebec (TPS 5 %, TVQ 9,975 %, toutes deux sur le HT).
    Ailleurs : --taux1 / --taux2. Une seule taxe : --taux2 0
"""
import argparse
import io
import json
import os
import re
import sys
from collections import defaultdict

sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")

RE_HORO = re.compile(r"(\d{8,12})")
# Tolerer TOUTES les espaces de separation de milliers, pas seulement l'espace
# ordinaire et l'insecable. Corrige le 2026-08-29 : la fine insecable U+2009 et
# la virgule anglo DECAPITAIENT le montant en silence - "1 234,56" devenait
# 234,56 et "4,160.00" devenait 160.00. Regle du motif trop strict, section 4.
ESPACES = "\u00a0\u202f\u2009\u2007\u2008\u2002\u2003 "
RE_MONTANT = re.compile(r"(\d[\d" + ESPACES + r",.]*\d[.,]\d{2})\s*\$")
RE_DEST = re.compile(r"(?i)^\s*factur[ée]?\s*[àa]\s*:?\s*(.*)$")


def montant(s):
    """Lit un montant quel que soit le separateur de milliers.

    Le DERNIER separateur rencontre est le decimal : "4,160.00" -> 4160.00 et
    "4 160,00" -> 4160.00. Tout ce qui precede n'est que du bruit de milliers.
    """
    t = s
    for c in ESPACES:
        t = t.replace(c, "")
    i = max(t.rfind(","), t.rfind("."))
    if i < 0:
        return None
    try:
        return round(float(re.sub(r"[.,]", "", t[:i]) + "." + t[i + 1:]), 2)
    except ValueError:
        return None


def texte(chemin):
    """Texte lisible d'un HTML, bloc de commentaire (gabarit) retire."""
    brut = io.open(chemin, encoding="utf-8", errors="replace").read()
    corps = re.sub(r"<!--[\s\S]*?-->", " ", brut)
    corps = re.sub(r"<script[\s\S]*?</script>|<style[\s\S]*?</style>", " ", corps, flags=re.I)
    plat = re.sub(r"<[^>]+>", "\n", corps)
    plat = plat.replace("&nbsp;", " ").replace("&amp;", "&").replace("&#39;", "'")
    return [l.strip() for l in plat.split("\n") if l.strip()]


def destinataire(lignes):
    for i, l in enumerate(lignes):
        m = RE_DEST.match(l)
        if m:
            v = m.group(1).strip()
            if v:
                return v
            if i + 1 < len(lignes):
                return lignes[i + 1].strip()
    return None


def resoudre(montants, t1, t2):
    """Cherche le HT dont les trois derives sont TOUS presents.

    Si PLUSIEURS HT bouclent, on n'en retient AUCUN : l'ambiguite s'isole.

    Corrige le 2026-08-29. La version precedente parcourait `sorted(S, reverse=True)`
    et rendait le PREMIER succes : elle retenait donc le PLUS GROS quadruplet.
    Mesure : un document portant la vraie facture du jour (1149,75) ET une facture
    anterieure rappelee (2299,50) rendait 2299,50 marque `fiable=true` - 100 % de
    trop, en silence. N'importe quel document citant un autre jeu de montants -
    facture precedente, extra, soumission rappelee, note de credit - declenchait
    le cas.
    """
    S = set(montants)
    sols = []
    for ht in sorted(S, reverse=True):
        if ht <= 0:
            continue
        a, b, tot = round(ht * t1, 2), round(ht * t2, 2), round(ht * (1 + t1 + t2), 2)
        attendus = [(a, .02), (tot, .03)] + ([(b, .02)] if t2 else [])
        if all(any(abs(x - v) <= tol for x in S) for v, tol in attendus):
            sols.append((ht, a, b, tot))
    if len(sols) == 1:
        return sols[0]
    return None, None, None, None


def canoniser(nom):
    if not nom:
        return None
    n = re.sub(r"\s+", " ", nom).strip(" .,;:-")
    n = re.sub(r"(?i)\b(inc|ltd|ltee|ltée|enr|senc|srl|sa|sas)\b\.?", "", n).strip(" .,-")
    return n[:44] or None


def main():
    ap = argparse.ArgumentParser(description="Inventaire d'un dossier de factures. Lecture seule.")
    ap.add_argument("--dossier", required=True, help="dossier a inventorier (recursif)")
    ap.add_argument("--annee", default=None, help="ne garder que cette annee (ex. 2026)")
    ap.add_argument("--taux1", type=float, default=0.05, help="1re taxe (defaut TPS 0.05)")
    ap.add_argument("--taux2", type=float, default=0.09975, help="2e taxe (defaut TVQ 0.09975 ; 0 si aucune)")
    ap.add_argument("--conventions", action="store_true", help="analyser les formes de nommage")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()

    if not os.path.isdir(a.dossier):
        sys.exit("ARRET : dossier introuvable — %s" % a.dossier)

    fac, hors, formes = {}, [], defaultdict(int)
    pdf_par_horo = defaultdict(list)

    for racine, _, noms in os.walk(a.dossier):
        for f in sorted(noms):
            ext = os.path.splitext(f)[1].lower()
            if ext not in (".html", ".htm", ".pdf"):
                continue
            m = RE_HORO.search(f)
            horo = m.group(1) if m else None
            # forme de nommage, horodatage neutralise
            formes[re.sub(r"\d{8,12}", "<H>", os.path.splitext(f)[0])[:38]] += 1
            if ext == ".pdf":
                if horo:
                    pdf_par_horo[horo].append(f)
                else:
                    hors.append((f, os.path.relpath(racine, a.dossier), "PDF sans horodatage"))
                continue
            chemin = os.path.join(racine, f)
            lignes = texte(chemin)
            blob = "\n".join(lignes)
            cli = canoniser(destinataire(lignes))
            if cli and "{{" in cli:
                cli = None
            if not horo:
                hors.append((f, os.path.relpath(racine, a.dossier), "aucun horodatage dans le nom"))
                continue
            if a.annee and not horo.startswith(a.annee):
                continue
            mts = [x for x in (montant(v) for v in RE_MONTANT.findall(blob)) if x is not None]
            ht, t1, t2, tot = resoudre(mts, a.taux1, a.taux2)
            cle = (horo, cli or "(client non lu dans le document)")
            # Deux fichiers de MEME horodatage et MEME client : la version
            # precedente ecrasait la premiere sans un mot, et la facture perdue
            # n'apparaissait dans AUCUN compteur. Mesure du 2026-08-29 : deux
            # factures de 2299,50 ne comptaient que pour une. On isole plutot que
            # d'additionner - c'est a l'humain de trancher.
            if cle in fac and fac[cle]["fichier"] != f:
                hors.append((f, os.path.relpath(racine, a.dossier),
                             "DOUBLON horodatage+client de %s - NON ADDITIONNE, a trancher"
                             % fac[cle]["fichier"]))
                continue
            fac[cle] = dict(horodatage=horo, client=cle[1], destinataire_brut=destinataire(lignes),
                            ht=ht, taxe1=t1, taxe2=t2, total=tot,
                            fiable=tot is not None, fichier=f,
                            dossier=os.path.relpath(racine, a.dossier), pdf=False)

    for horo, noms in pdf_par_horo.items():
        cles = [k for k in fac if k[0] == horo]
        if len(cles) == 1:
            fac[cles[0]]["pdf"] = True
        elif not cles:
            for n in noms:
                hors.append((n, "", "PDF sans HTML correspondant — montant non lisible"))

    liste = sorted(fac.values(), key=lambda x: (x["horodatage"], x["client"]))
    fiables = [f for f in liste if f["fiable"]]
    collisions = {}
    for k in fac:
        collisions.setdefault(k[0], []).append(k[1])
    collisions = {h: c for h, c in collisions.items() if len(c) > 1}
    par_client = defaultdict(float)
    for f in fiables:
        par_client[f["client"]] += f["total"]

    res = dict(dossier=a.dossier, annee=a.annee,
               factures_distinctes=len(liste), exploitees=len(fiables),
               a_verifier=len(liste) - len(fiables), hors_perimetre=len(hors),
               total_ttc=round(sum(f["total"] for f in fiables), 2),
               total_ht=round(sum(f["ht"] for f in fiables), 2),
               total_taxe1=round(sum(f["taxe1"] for f in fiables), 2),
               total_taxe2=round(sum(f["taxe2"] or 0 for f in fiables), 2),
               collisions_horodatage=collisions,
               par_client={k: round(v, 2) for k, v in sorted(par_client.items(), key=lambda x: -x[1])},
               formes_de_nommage=dict(sorted(formes.items(), key=lambda x: -x[1])),
               factures=liste,
               hors_perimetre_detail=[dict(fichier=f, dossier=d, raison=r) for f, d, r in hors])

    if a.json:
        print(json.dumps(res, ensure_ascii=False, indent=1))
        return

    print("  Dossier : %s%s" % (a.dossier, ("  (annee %s)" % a.annee) if a.annee else ""))
    print("  %d facture(s) distincte(s) — %d exploitee(s), %d a verifier, %d hors perimetre"
          % (res["factures_distinctes"], res["exploitees"], res["a_verifier"], res["hors_perimetre"]))
    print()
    print("  TOTAL TTC   %12.2f $" % res["total_ttc"])
    print("    dont HT   %12.2f $" % res["total_ht"])
    print("    taxe 1    %12.2f $" % res["total_taxe1"])
    if a.taux2:
        print("    taxe 2    %12.2f $" % res["total_taxe2"])
    print()
    print("  PAR CLIENT")
    for k, v in res["par_client"].items():
        print("    %-42s %12.2f $" % (k[:42], v))

    if collisions:
        print()
        print("  ATTENTION — horodatages partages par plusieurs clients :")
        for h, c in sorted(collisions.items()):
            print("    %s -> %s" % (h, " | ".join(sorted(c))))
        print("    Dedupliquer sur l'horodatage seul FUSIONNERAIT ces factures.")

    aver = [f for f in liste if not f["fiable"]]
    if aver:
        print()
        print("  A VERIFIER — montants non concordants, non devines :")
        for f in aver:
            print("    %-46s %s" % (f["fichier"][:46], f["client"]))

    if hors:
        print()
        print("  HORS PERIMETRE :")
        for f, d, r in hors:
            print("    %-46s %s" % (f[:46], r))

    if a.conventions:
        print()
        print("  FORMES DE NOMMAGE (horodatage neutralise en <H>) :")
        for k, v in res["formes_de_nommage"].items():
            print("    %3d x  %s" % (v, k))
        print("    %d formes distinctes — chercher en INSENSIBLE A LA CASSE." % len(res["formes_de_nommage"]))


if __name__ == "__main__":
    main()
