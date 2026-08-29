"""
outlook_mail.py — Contrôle complet d'une boîte Outlook depuis Python (MAPI/COM).

Lire, chercher, rédiger, répondre, envoyer, classer, supprimer.
Fonctionne avec le profil Outlook déjà authentifié sur le poste : aucun mot de
passe, aucun jeton, aucune inscription d'application, aucun accès administrateur.

Prérequis : Windows, Outlook (classique) installé et LANCÉ, `pip install pywin32`.

    python outlook_mail.py accounts
    python outlook_mail.py folders
    python outlook_mail.py list   --folder inbox --limit 25 [--unread] [--query TEXTE]
    python outlook_mail.py search --query TEXTE [--folder all] [--limit 50]
    python outlook_mail.py read   --id <EntryID> [--html]
    python outlook_mail.py thread --id <EntryID>
    python outlook_mail.py draft  --to a@b.com --subject "..." --body "..." [--cc] [--attach F...]
    python outlook_mail.py reply  --id <EntryID> --body "..." [--all]
    python outlook_mail.py send   --id <EntryID> --yes-send
    python outlook_mail.py move   --id <EntryID> --folder archive
    python outlook_mail.py mark   --id <EntryID> [--unread]
    python outlook_mail.py flag   --id <EntryID> [--off]
    python outlook_mail.py delete --id <EntryID>

Ajoutez --json à n'importe quelle commande pour une sortie exploitable par script.
Ajoutez --account "adresse@domaine" pour cibler une boîte précise (sinon : compte
par défaut du profil).

Deux garde-fous sont codés en dur, volontairement :
  * `send` refuse de partir sans --yes-send.
  * `delete` déplace vers les Éléments supprimés et refuse d'agir sur un message
    qui s'y trouve déjà. Aucune purge définitive n'est exposée.

Licence MIT.
"""
import argparse, json, sys, datetime, os, io

try:
    import win32com.client as win32
    import pythoncom
except ImportError:
    sys.exit("pywin32 manquant.  ->  pip install pywin32")

# Identifiants de dossiers par défaut d'Outlook (OlDefaultFolders)
OL = dict(deleted=3, outbox=4, sent=5, inbox=6, calendar=9, contacts=10,
          journal=11, notes=12, tasks=13, drafts=16, junk=23)
OL_MAIL_ITEM = 43          # olMail
CREATE_MAIL = 0            # olMailItem


# --------------------------------------------------------------- connexion
def ns():
    return win32.Dispatch("Outlook.Application").GetNamespace("MAPI")


def list_accounts():
    n, out = ns(), []
    for i in range(1, n.Folders.Count + 1):
        try:
            f = n.Folders.Item(i)
            out.append(dict(name=f.Name, items=f.Folders.Count))
        except Exception:
            pass
    return out


def store_root(account=None):
    """Racine du magasin visé. Sans --account : celui de la boîte par défaut."""
    n = ns()
    if account:
        for i in range(1, n.Folders.Count + 1):
            f = n.Folders.Item(i)
            if account.lower() in f.Name.lower():
                return f
        raise SystemExit(f"compte introuvable : {account}\n"
                         f"comptes disponibles : "
                         f"{', '.join(a['name'] for a in list_accounts())}")
    try:
        return n.GetDefaultFolder(OL["inbox"]).Parent
    except Exception:
        return n.Folders.Item(1)


def folder_path(f):
    parts, cur, guard = [], f, 0
    while cur is not None and guard < 20:
        guard += 1
        try:
            parts.append(cur.Name)
            cur = cur.Parent
            if cur is None or not hasattr(cur, "Name"):
                break
        except Exception:
            break
    return "/".join(reversed(parts))


def resolve_folder(name, account=None):
    """Accepte un alias (inbox, sent, drafts…), un nom simple, ou un chemin A/B."""
    name = (name or "inbox").strip()
    key = name.lower()
    if key in OL:
        f = ns().GetDefaultFolder(OL[key])
        if account:                       # forcer le bon magasin
            root = store_root(account)
            if not folder_path(f).startswith(root.Name):
                found = _find(root, [f.Name.lower()])
                if found:
                    return found
        return f
    root = store_root(account)
    if key in ("root", "racine"):
        return root
    parts = [p for p in key.replace("\\", "/").split("/") if p]
    found = _find(root, parts)
    if found is None:
        raise SystemExit(f"dossier introuvable : {name}")
    return found


def _find(root, parts):
    best = None
    for f in _walk(root):
        try:
            fp = folder_path(f).lower()
        except Exception:
            continue
        if fp.endswith("/".join(parts)) or f.Name.lower() == parts[-1]:
            if best is None or len(fp) < len(folder_path(best)):
                best = f
    return best


def _walk(f, depth=0, maxdepth=5):
    yield f
    if depth >= maxdepth:
        return
    try:
        for i in range(1, f.Folders.Count + 1):
            yield from _walk(f.Folders.Item(i), depth + 1, maxdepth)
    except Exception:
        return


# --------------------------------------------------------------- lecture
def _dt(v):
    try:
        return datetime.datetime(v.year, v.month, v.day,
                                 v.hour, v.minute, v.second).isoformat(" ")
    except Exception:
        return None


def brief(m, body_chars=0):
    """Dictionnaire normalisé décrivant un message. body_chars=-1 => corps entier."""
    d = {}
    for attr, key in (("EntryID", "id"), ("Subject", "subject"),
                      ("SenderName", "from"), ("SenderEmailAddress", "from_addr"),
                      ("To", "to"), ("CC", "cc"), ("Categories", "categories"),
                      ("ConversationTopic", "topic")):
        try:
            d[key] = getattr(m, attr)
        except Exception:
            d[key] = None
    for attr, key in (("ReceivedTime", "received"), ("SentOn", "sent")):
        try:
            d[key] = _dt(getattr(m, attr))
        except Exception:
            d[key] = None
    for attr, key, dflt in (("UnRead", "unread", False), ("Size", "size", 0),
                            ("FlagStatus", "flag", 0), ("Importance", "importance", 1)):
        try:
            d[key] = getattr(m, attr)
        except Exception:
            d[key] = dflt
    try:
        d["attachments"] = [a.FileName for a in m.Attachments]
    except Exception:
        d["attachments"] = []
    try:
        d["folder"] = folder_path(m.Parent)
    except Exception:
        d["folder"] = None
    if body_chars:
        try:
            b = " ".join((m.Body or "").split())
            d["body"] = b if body_chars < 0 else b[:body_chars]
        except Exception:
            d["body"] = None
    return d


def iter_items(folder, limit=None, unread=False, newest=True):
    items = folder.Items
    try:
        items.Sort("[ReceivedTime]", newest)
    except Exception:
        pass
    if unread:
        try:
            items = items.Restrict("[UnRead] = True")
            items.Sort("[ReceivedTime]", newest)
        except Exception:
            pass
    out = []
    for it in items:
        try:
            if it.Class != OL_MAIL_ITEM:
                continue
        except Exception:
            continue
        out.append(it)
        if limit and len(out) >= limit:
            break
    return out


def matches(m, q):
    q = q.lower()
    for attr in ("Subject", "SenderName", "SenderEmailAddress", "To", "CC"):
        try:
            v = getattr(m, attr)
            if v and q in str(v).lower():
                return True
        except Exception:
            pass
    try:
        if m.Body and q in m.Body.lower():
            return True
    except Exception:
        pass
    return False


# --------------------------------------------------------------- commandes
def cmd_accounts(a):
    rows = list_accounts()
    print(json.dumps(rows, ensure_ascii=False, indent=2) if a.json
          else "\n".join("  " + r["name"] for r in rows))


def cmd_folders(a):
    rows = []
    for f in _walk(store_root(a.account)):
        try:
            rows.append(dict(path=folder_path(f), items=f.Items.Count,
                             unread=f.UnReadItemCount,
                             depth=folder_path(f).count("/")))
        except Exception:
            pass
    if a.json:
        print(json.dumps(rows, ensure_ascii=False, indent=2))
    else:
        base = min((r["depth"] for r in rows), default=0)
        for r in rows:
            if r["items"] or r["unread"]:
                pad = "  " * (r["depth"] - base)
                print(f"{pad}{r['path'].split('/')[-1]:<38} "
                      f"{r['items']:>6} élts  {r['unread']:>4} non lus")


def cmd_list(a):
    f = resolve_folder(a.folder, a.account)
    items = iter_items(f, (a.limit * 6 if a.query else a.limit), a.unread)
    if a.query:
        items = [m for m in items if matches(m, a.query)][:a.limit]
    emit([brief(m, 160) for m in items], a)


def cmd_search(a):
    if a.folder and a.folder.lower() != "all":
        folders = [resolve_folder(a.folder, a.account)]
    else:
        # maxdepth aligne sur `folders` (5) le 2026-08-29 : a 3, un message
        # dans un dossier profond etait VISIBLE dans `folders` et INVISIBLE
        # a `search`, sans un mot. Faux zero de la section 4.
        folders = list(_walk(store_root(a.account), maxdepth=5))
    rows = []
    for f in folders:
        try:
            vus = iter_items(f, 500)
            # Plafond rendu VISIBLE le 2026-08-29 : au-dela, une correspondance
            # plus ancienne etait perdue sans un mot. Un zero ne vaut que prouve.
            if len(vus) >= 500:
                print("[avertissement] %s : plafond de 500 messages atteint - les plus "
                      "anciens n'ont PAS ete examines." % folder_path(f), file=sys.stderr)
            for m in vus:
                if matches(m, a.query):
                    rows.append(brief(m, 160))
                    if len(rows) >= a.limit:
                        break
        except Exception:
            continue
        if len(rows) >= a.limit:
            break
    emit(rows, a)


def cmd_read(a):
    m = ns().GetItemFromID(a.id)
    d = brief(m, -1)
    if a.html:
        try:
            d["html"] = m.HTMLBody
        except Exception:
            d["html"] = None
    if a.json:
        print(json.dumps(d, ensure_ascii=False, indent=2))
        return
    print(f"De      : {d['from']} <{d['from_addr']}>")
    print(f"À       : {d['to']}")
    if d["cc"]:
        print(f"Cc      : {d['cc']}")
    print(f"Date    : {d['received'] or d['sent']}")
    print(f"Dossier : {d['folder']}")
    print(f"Objet   : {d['subject']}")
    if d["attachments"]:
        print(f"PJ      : {', '.join(d['attachments'])}")
    print("-" * 72)
    print(d.get("body") or "")


def cmd_thread(a):
    src = ns().GetItemFromID(a.id)
    topic = (src.ConversationTopic or src.Subject or "").strip().lower()
    rows = []
    for f in _walk(store_root(a.account), maxdepth=5):   # aligne sur folders, 2026-08-29
        try:
            vus = iter_items(f, 500)
            if len(vus) >= 500:
                print("[avertissement] %s : plafond de 500 messages atteint - le fil peut "
                      "etre incomplet." % folder_path(f), file=sys.stderr)
            for m in vus:
                if (m.ConversationTopic or "").strip().lower() == topic:
                    rows.append(brief(m, 400))
        except Exception:
            continue
    rows.sort(key=lambda r: r["received"] or r["sent"] or "")
    emit(rows, a)


def compte_envoi(adresse):
    """Objet Account Outlook correspondant à `adresse`, ou None.

    Sans lui, un brouillon part TOUJOURS du compte par défaut du profil, quel
    que soit `--account` : le drapeau ne servait qu'à choisir le magasin où
    LIRE. Avec deux comptes dans le profil (l'adresse principale et une adresse
    Gmail), écrire depuis le mauvais expéditeur est silencieux et se découvre
    chez le destinataire.
    """
    if not adresse:
        return None
    comptes = win32.Dispatch("Outlook.Application").Session.Accounts
    besoin = adresse.lower().strip()
    for i in range(1, comptes.Count + 1):
        try:
            c = comptes.Item(i)
        except Exception:
            continue
        for attribut in ("SmtpAddress", "DisplayName", "UserName"):
            try:
                v = (getattr(c, attribut) or "").lower()
            except Exception:
                v = ""
            if v and (besoin == v or besoin in v):
                return c
    raise SystemExit(
        f"compte expéditeur introuvable : {adresse}\n"
        f"comptes du profil : " + ", ".join(
            (comptes.Item(i).SmtpAddress or comptes.Item(i).DisplayName or "?")
            for i in range(1, comptes.Count + 1)))


def _poser_expediteur(m, a):
    """Applique --account comme compte expéditeur, si demandé.

    ⚠️ L'affectation peut être SILENCIEUSEMENT IGNORÉE par Outlook : mesuré le
    2026-08-17 sur un compte Gmail (IMAP) ajouté au profil — `SendUsingAccount`
    relisait `None` après affectation, avant ET après `Save()`, en liaison
    tardive comme précoce, et créer l'élément dans les Brouillons du compte
    levait « Vous n'avez pas l'autorisation de créer une entrée dans ce
    dossier ». Le compte était pourtant bien dans `Session.Accounts` et sa
    boîte lisible. Cause probable : compte non pleinement configuré pour
    l'envoi côté Outlook, pas un défaut d'API.

    On VÉRIFIE donc que l'affectation a pris, et on refuse plutôt que de
    laisser partir un brouillon depuis le mauvais expéditeur — une erreur qui
    ne se découvrirait que chez le destinataire.
    """
    demande = getattr(a, "account", None)
    compte = compte_envoi(demande)
    if compte is None:
        return None

    # `SendUsingAccount` n'est PAS relisible via pywin32 : elle rend `None`
    # après affectation quel que soit le compte, y compris le compte par défaut
    # pour lequel l'envoi fonctionne. Vérifier par relecture était donc un faux
    # test — il refusait aussi le cas qui marchait. On raisonne sur le RANG.
    try:
        defaut = ns().Accounts.Item(1)
        est_defaut = (compte.SmtpAddress or "").lower() == (defaut.SmtpAddress or "").lower()
    except Exception:
        est_defaut = False

    if est_defaut:
        return compte  # rien à poser : c'est déjà l'expéditeur naturel

    m.SendUsingAccount = compte
    print(f"[avertissement] expéditeur « {compte.SmtpAddress} » demandé sur un compte "
          f"NON PAR DÉFAUT.\n"
          f"    Outlook ne permet pas de relire ce réglage : je ne peux pas garantir "
          f"qu'il a pris.\n"
          f"    OUVRE LE BROUILLON DANS OUTLOOK et vérifie le champ « De » avant "
          f"d'autoriser l'envoi.", file=sys.stderr)
    return compte


RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def charger_signature(nom):
    """Lit `profiles/signature_<nom>.html`.

    POURQUOI CE FICHIER PLUTOT QU'OUTLOOK. Les signatures de l'utilisateur sont
    INFONUAGIQUES : elles vivent dans la boîte Microsoft 365 et se gèrent depuis
    Outlook Web. Elles ne redescendent PAS dans
    `%APPDATA%\\Microsoft\\Signatures` — mesure du 2026-08-17 : ce dossier ne
    contenait que cinq signatures périmées, dont deux d'entreprises antérieures,
    et AUCUNE des cinq réellement utilisées. Et un brouillon créé par COM ne
    reçoit jamais de signature :
    Outlook ne les applique que dans sa propre fenêtre de rédaction.

    La signature a donc été extraite d'un COURRIEL RÉELLEMENT ENVOYÉ — la seule
    source qui dise ce qui part vraiment. ⚠️ En la ré-extrayant un jour,
    resserrer les bornes : une première extraction avait emporté la réponse
    citée, donc l'adresse du destinataire de ce message-là.
    """
    # Le gabarit n'est pas une signature. Ajouté le 2026-08-29 : un dépôt cloné
    # ne porte QUE `MODELE`, et `--signature MODELE` réussissait — le courriel
    # partait signé « PRENOM NOM / adresse@exemple.ca » chez un vrai client.
    if nom.strip().upper() == "MODELE":
        raise SystemExit(
            "refus : MODELE est le GABARIT, pas une signature.\n"
            "Il porte encore PRENOM NOM et adresse@exemple.ca.\n"
            "Copiez profiles/signature_MODELE.html en profiles/signature_defaut.html,\n"
            "remplissez-le depuis un courriel RÉELLEMENT ENVOYÉ, puis réessayez.")
    chemin = os.path.join(RACINE, "profiles", f"signature_{nom}.html")
    if not os.path.exists(chemin):
        dispo = []
        dossier = os.path.join(RACINE, "profiles")
        if os.path.isdir(dossier):
            dispo = [f[10:-5] for f in os.listdir(dossier)
                     if f.startswith("signature_") and f.endswith(".html")
                     and f != "signature_MODELE.html"]
        raise SystemExit(
            f"signature introuvable : {chemin}\n"
            f"disponibles : {', '.join(dispo) if dispo else '(aucune)'}\n"
            f"Le dépôt ne livre volontairement AUCUNE signature remplie.\n"
            f"Copiez profiles/signature_MODELE.html en profiles/signature_defaut.html,\n"
            f"puis remplissez-le depuis un courriel réellement envoyé.")
    with io.open(chemin, encoding="utf-8") as f:
        contenu = f.read()

    # Refuser un gabarit COPIE mais PAS ENCORE REMPLI.
    # Ajouté le 2026-08-29 après mesure : refuser `--signature MODELE` ne suffit
    # pas. Le manuel demande de copier signature_MODELE.html en
    # signature_defaut.html PUIS d'y coller son bloc ; entre les deux, le fichier
    # existe, `charger_signature` le charge sans broncher, et un brouillon part
    # signé « PRENOM NOM / adresse@exemple.ca ». Mesuré : un `draft --signature`
    # a réellement produit ce brouillon dans une vraie boîte.
    MARQUEURS = ("PRENOM NOM", "NOM DE L'ENTREPRISE", "adresse@exemple.ca",
                 "(000) 000-0000", "A0A 0A0", "www.exemple.ca")
    restants = [m for m in MARQUEURS if m in contenu]
    if restants:
        raise SystemExit(
            f"refus : la signature '{nom}' n'est pas remplie.\n"
            f"Elle porte encore : {', '.join(restants)}\n"
            f"Fichier : {chemin}\n"
            f"Remplacez ces marqueurs par vos vraies coordonnées, extraites d'un\n"
            f"courriel RÉELLEMENT ENVOYÉ, puis réessayez.")
    return contenu


def _corps_avec_signature(a):
    """Rend (html, utilise_html) en tenant compte de --signature."""
    corps = a.body or ""
    if not getattr(a, "signature", None):
        return corps, bool(getattr(a, "html_body", False))
    sig = charger_signature(a.signature)
    if getattr(a, "html_body", False):
        return corps + "<br>" + sig, True
    # Corps en texte + signature HTML : on passe en HTML en préservant les
    # sauts de ligne, sinon le texte s'affiche en un seul bloc.
    echappe = (corps.replace("&", "&amp;").replace("<", "&lt;")
                    .replace(">", "&gt;").replace("\n", "<br>"))
    return f'<div style="font-family:Aptos,Calibri,sans-serif;font-size:12pt">{echappe}</div><br>{sig}', True


def cmd_draft(a):
    m = win32.Dispatch("Outlook.Application").CreateItem(CREATE_MAIL)
    _poser_expediteur(m, a)
    m.To = a.to
    if a.cc:
        m.CC = a.cc
    if a.bcc:
        m.BCC = a.bcc
    m.Subject = a.subject or ""
    corps, en_html = _corps_avec_signature(a)
    if en_html:
        m.HTMLBody = corps
    else:
        m.Body = corps
    for p in (a.attach or []):
        if os.path.exists(p):
            m.Attachments.Add(os.path.abspath(p))
        else:
            print(f"[avertissement] pièce jointe introuvable : {p}", file=sys.stderr)
    m.Save()
    print(json.dumps(dict(status="brouillon créé", id=m.EntryID,
                          subject=m.Subject, to=m.To,
                          expediteur=_adresse_expediteur(m)),
                     ensure_ascii=False, indent=2))


def _adresse_expediteur(m):
    """Adresse réellement utilisée à l'envoi, pour la montrer avant d'envoyer."""
    try:
        c = m.SendUsingAccount
        if c is not None:
            return c.SmtpAddress or c.DisplayName
    except Exception:
        pass
    try:
        return ns().Accounts.Item(1).SmtpAddress + " (compte par défaut)"
    except Exception:
        return "(compte par défaut)"


def cmd_reply(a):
    src = ns().GetItemFromID(a.id)
    r = src.ReplyAll() if a.all else src.Reply()
    # Par défaut Outlook répond depuis le compte qui a REÇU le message, ce qui
    # est le bon comportement ; --account permet de le forcer autrement.
    _poser_expediteur(r, a)
    if getattr(a, "signature", None):
        # On écrit AVANT la citation, sans y toucher : `r.HTMLBody` contient
        # déjà l'original mis en forme par Outlook, et le réécrire en texte
        # brut le détruirait.
        tete, _ = _corps_avec_signature(a)
        r.HTMLBody = tete + (r.HTMLBody or "")
    else:
        # Corrige le 2026-08-29 : affecter `.Body` sur une reponse dont
        # l'original est HTML convertit TOUT le fil cite en texte brut. On passe
        # par HTMLBody, comme la branche --signature juste au-dessus.
        echappe = ((a.body or "").replace("&", "&amp;").replace("<", "&lt;")
                   .replace(">", "&gt;").replace("\n", "<br>"))
        r.HTMLBody = ('<div style="font-family:Aptos,Calibri,sans-serif;font-size:12pt">'
                      + echappe + '</div><br>' + (r.HTMLBody or ""))
    r.Save()
    print(json.dumps(dict(status="réponse enregistrée en brouillon", id=r.EntryID,
                          subject=r.Subject, to=r.To,
                          expediteur=_adresse_expediteur(r)),
                     ensure_ascii=False, indent=2))


def cmd_send(a):
    if not a.yes_send:
        raise SystemExit(
            "refus : l'envoi exige --yes-send.\n"
            "Ce drapeau matérialise un accord explicite de l'utilisateur pour CET envoi.")
    m = ns().GetItemFromID(a.id)
    to, subj = m.To, m.Subject
    m.Send()
    print(json.dumps(dict(status="envoyé", to=to, subject=subj),
                     ensure_ascii=False, indent=2))


def cmd_move(a):
    m = ns().GetItemFromID(a.id)
    dest = resolve_folder(a.folder, a.account)
    src = folder_path(m.Parent)
    m.Move(dest)
    print(json.dumps(dict(status="déplacé", de=src, vers=folder_path(dest)),
                     ensure_ascii=False, indent=2))


def cmd_mark(a):
    m = ns().GetItemFromID(a.id)
    m.UnRead = bool(a.unread)
    m.Save()
    print(json.dumps(dict(status="marqué", unread=m.UnRead), ensure_ascii=False, indent=2))


def cmd_flag(a):
    m = ns().GetItemFromID(a.id)
    m.FlagStatus = 0 if a.off else 2
    m.Save()
    print(json.dumps(dict(status="indicateur", flag=m.FlagStatus),
                     ensure_ascii=False, indent=2))


def cmd_delete(a):
    """Déplace vers les Éléments supprimés. RÉVERSIBLE. Jamais de purge définitive."""
    m = ns().GetItemFromID(a.id)
    cur = folder_path(m.Parent).lower()
    if any(k in cur for k in ("supprim", "deleted", "gelöscht", "eliminado")):
        raise SystemExit(
            "refus : ce message est déjà dans les Éléments supprimés.\n"
            "La suppression définitive n'est pas exposée par cet outil.\n"
            "Si vous la voulez, faites-la vous-même depuis Outlook.")
    subj = m.Subject
    m.Delete()
    print(json.dumps(dict(status="déplacé vers les Éléments supprimés (réversible)",
                          subject=subj), ensure_ascii=False, indent=2))


def emit(rows, a):
    if a.json:
        print(json.dumps(rows, ensure_ascii=False, indent=2))
        return
    for r in rows:
        mark = ("*" if r.get("unread") else " ") + ("@" if r.get("attachments") else " ")
        print(f"\n{mark} {r.get('received') or r.get('sent')}  [{r.get('folder')}]")
        print(f"    De    : {r.get('from')} <{r.get('from_addr')}>")
        print(f"    Objet : {r.get('subject')}")
        if r.get("body"):
            print(f"    Corps : {r['body'][:200]}")
        print(f"    ID    : {r.get('id')}")
    print(f"\n({len(rows)} message(s))")


# --------------------------------------------------------------- CLI
def main():
    p = argparse.ArgumentParser(
        description="Contrôle complet d'une boîte Outlook via MAPI/COM.")
    p.add_argument("--account", default=None,
                   help="adresse ou nom du magasin (défaut : compte principal)")
    p.add_argument("--json", action="store_true", help="sortie JSON")
    sub = p.add_subparsers(dest="cmd", required=True)

    def add(name, fn, help=""):
        s = sub.add_parser(name, help=help)
        s.set_defaults(func=fn)
        return s

    add("accounts", cmd_accounts, "lister les comptes du profil")
    add("folders", cmd_folders, "arborescence des dossiers avec compteurs")
    s = add("list", cmd_list, "lister les messages d'un dossier")
    s.add_argument("--folder", default="inbox")
    s.add_argument("--limit", type=int, default=25)
    s.add_argument("--unread", action="store_true")
    s.add_argument("--query", default=None)
    s = add("search", cmd_search, "rechercher dans plusieurs dossiers")
    s.add_argument("--query", required=True)
    s.add_argument("--folder", default="all")
    s.add_argument("--limit", type=int, default=50)
    s = add("read", cmd_read, "lire un message complet")
    s.add_argument("--id", required=True)
    s.add_argument("--html", action="store_true")
    s = add("thread", cmd_thread, "reconstituer un fil de discussion")
    s.add_argument("--id", required=True)
    s = add("draft", cmd_draft, "créer un brouillon")
    s.add_argument("--to", required=True)
    s.add_argument("--subject", default="")
    s.add_argument("--body", default="")
    s.add_argument("--cc", default="")
    s.add_argument("--bcc", default="")
    s.add_argument("--attach", nargs="*")
    s.add_argument("--html-body", action="store_true")
    s.add_argument("--signature", nargs="?", const="defaut", default=None,
                   metavar="NOM",
                   help="ajouter la signature profiles/signature_<NOM>.html "
                        "(défaut : defaut). Le corps passe alors en HTML.")
    s = add("reply", cmd_reply, "préparer une réponse en brouillon")
    s.add_argument("--id", required=True)
    s.add_argument("--body", default="")
    s.add_argument("--all", action="store_true")
    s.add_argument("--signature", nargs="?", const="defaut", default=None,
                   metavar="NOM", help="idem pour une réponse")
    s = add("send", cmd_send, "envoyer un brouillon existant")
    s.add_argument("--id", required=True)
    s.add_argument("--yes-send", action="store_true")
    s = add("move", cmd_move, "déplacer vers un dossier")
    s.add_argument("--id", required=True)
    s.add_argument("--folder", required=True)
    s = add("mark", cmd_mark, "marquer lu / non lu")
    s.add_argument("--id", required=True)
    s.add_argument("--unread", action="store_true")
    s = add("flag", cmd_flag, "poser / retirer l'indicateur de suivi")
    s.add_argument("--id", required=True)
    s.add_argument("--off", action="store_true")
    s = add("delete", cmd_delete, "déplacer vers les Éléments supprimés")
    s.add_argument("--id", required=True)

    a = p.parse_args()
    pythoncom.CoInitialize()
    try:
        a.func(a)
    except Exception as e:
        sys.exit(f"erreur : {e}\n"
                 f"(Outlook est-il lancé ? Le compte existe-t-il ? "
                 f"Essayez : python outlook_mail.py accounts)")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
        sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
    except Exception:
        pass
    main()
