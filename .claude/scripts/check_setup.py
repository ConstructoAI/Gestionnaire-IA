"""
check_setup.py — Vérifie que l'environnement permet de piloter Outlook.

    python scripts/check_setup.py

Ne modifie rien. Affiche un diagnostic et, en cas de problème, la marche à suivre.
"""
import sys, os, platform, subprocess

OK, KO, WARN = "[ OK ]", "[ECHEC]", "[ ! ]"
problems = []


def line(state, label, detail=""):
    print(f"{state} {label}" + (f" — {detail}" if detail else ""))


# 1. Système -----------------------------------------------------------------
if platform.system() != "Windows":
    line(KO, "Système", f"{platform.system()} — ce kit exige Windows (MAPI/COM)")
    problems.append("Windows requis : MAPI/COM n'existe pas ailleurs.")
else:
    line(OK, "Système", f"Windows {platform.release()}")

# 2. Python ------------------------------------------------------------------
v = sys.version_info
if v < (3, 8):
    line(KO, "Python", f"{v.major}.{v.minor} — 3.8 minimum requis")
    problems.append("Installer Python 3.8 ou plus récent.")
else:
    line(OK, "Python", f"{v.major}.{v.minor}.{v.micro}")

# 3. pywin32 -----------------------------------------------------------------
try:
    import win32com.client as win32
    import pythoncom
    line(OK, "pywin32", "installé")
except ImportError:
    line(KO, "pywin32", "absent")
    problems.append("pip install pywin32")
    win32 = None

# 4. Outlook installé --------------------------------------------------------
outlook_exe = None
for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
             os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
    for sub in (r"Microsoft Office\root\Office16\OUTLOOK.EXE",
                r"Microsoft Office\Office16\OUTLOOK.EXE",
                r"Microsoft Office\Office15\OUTLOOK.EXE"):
        p = os.path.join(base, sub)
        if os.path.exists(p):
            outlook_exe = p
            break
    if outlook_exe:
        break
if outlook_exe:
    line(OK, "Outlook installé", outlook_exe)
else:
    line(WARN, "Outlook installé", "introuvable aux emplacements habituels")

# 5. Outlook lancé -----------------------------------------------------------
running = False
try:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE"],
                         capture_output=True, text=True, timeout=20).stdout
    running = "OUTLOOK.EXE" in out
except Exception:
    pass
if running:
    line(OK, "Outlook lancé", "processus actif")
else:
    line(KO, "Outlook lancé", "non démarré")
    problems.append("Démarrer Outlook : l'automatisation pilote une instance vivante.")

# 6. Accès MAPI --------------------------------------------------------------
if win32 and running:
    try:
        pythoncom.CoInitialize()
        ns = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        names = [ns.Folders.Item(i).Name for i in range(1, ns.Folders.Count + 1)]
        line(OK, "Accès MAPI", f"{len(names)} compte(s)")
        for n in names:
            print(f"        • {n}")
        try:
            inbox = ns.GetDefaultFolder(6)
            line(OK, "Boîte de réception",
                 f"{inbox.Items.Count} éléments, {inbox.UnReadItemCount} non lus")
            if inbox.Items.Count == 0:
                line(WARN, "Boîte vide",
                     "voir .claude/references/depannage.md")
        except Exception as e:
            line(WARN, "Boîte de réception", str(e))
    except Exception as e:
        line(KO, "Accès MAPI", str(e))
        problems.append("MAPI inaccessible. Outlook est-il configuré avec un compte ?")
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
elif win32:
    line(WARN, "Accès MAPI", "non testé (Outlook n'est pas lancé)")

# 7. Synthèse ----------------------------------------------------------------
print()
if problems:
    print("À corriger :")
    for p in problems:
        print(f"  → {p}")
    sys.exit(1)
print("Environnement prêt.  Essayez :")
print("    python scripts/outlook_mail.py folders")
print("    python scripts/outlook_mail.py list --limit 10")
