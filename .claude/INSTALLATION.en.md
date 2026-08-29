# Installation — Outlook workstation for Claude Code

*🇫🇷 [Version française](INSTALLATION.md) — **the French version is the source of truth.** This
translation is kept in step by hand; if the two ever disagree, the French one is right.*

> ### Courtesy of **Sylvain Leduc — Constructo AI inc.**
> *Intelligent ecosystem for construction in Quebec* · [www.constructoai.ca](https://www.constructoai.ca)

Manage your **email**, **calendar**, **project folders** and **bookkeeping** with Claude,
directly on your own Outlook — no password, no token, no app registration.

---

## What you need before you start

| Requirement | How to check | If it is missing |
|---|---|---|
| **Classic Outlook** | it opens from `Program Files\Microsoft Office\` | ⛔ The **Microsoft Store one does not expose MAPI/COM** — nothing will work. Install classic Outlook. |
| **Python 3.x** | `python -c "print(84)"` must return **84** | nothing to do — **the `.bat` installs it** (winget, user scope). Manual fallback: [python.org](https://www.python.org), tick "Add to PATH" |
| **pywin32** | `python -c "import win32com.client"` | nothing to do — **the `.bat` installs it** |
| **Claude Code** | `claude --version` | nothing to do — **the `.bat` installs it** (official installer, no admin rights). ⚠️ but you need a **paid Claude subscription** — Pro, Max, Team, Enterprise or Console; *this workstation is free, Claude Code is not* |
| **Git for Windows** | `git --version` | nothing to do — **the `.bat` installs it**. *Optional but recommended*: without it, Claude Code has no Bash tool and falls back to PowerShell, while this workstation's commands are written in POSIX shell |

⚠️ **Do not look for `claude.cmd`**: the native installer produces `claude.exe` in
`%USERPROFILE%\.local\bin`. The name `claude.cmd` only exists with an **npm** install.
*(Measured 2026-08-29: the `.bat` tested `claude.cmd` and refused to start on a machine where
Claude Code was installed and working.)*

### 🔴 Why `python --version` is not enough to check Python

Windows ships a **0-byte execution shortcut** at
`%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe`, and `where python` returns it. It **answers**,
which makes it undetectable by a simple version check:

| Situation | What `python --version` returns |
|---|---|
| A Store Python is installed behind it | a **real** version, `Python 3.13.14` |
| Nothing behind it | `Python was not found; run without arguments…` |

Both outputs are non-empty. ➜ **The only test that does not lie is to make Python execute
something**: `python -c "print(84)"` must return exactly `84`. A shortcut answers; an
interpreter computes. The `.bat` runs that test and says so when it discards the shortcut.
*(Measured 2026-08-29 on a fresh Windows 11 machine: the workstation believed Python was
installed, therefore never installed it, and `pip install pywin32` failed afterwards with no
visible cause.)*

---

## 🔴 Before you launch: this workstation runs WITHOUT permission guardrails

`settings.json` sets `"defaultMode": "bypassPermissions"` and the `.bat` launches Claude with
`--dangerously-skip-permissions`. **Claude will ask for no authorization** before reading a
file, writing one, or running a command in the working folder — the very folder this manual
invites you to place in OneDrive or SharePoint, alongside your client files.

It is a deliberate choice: without it, a workstation triaging two hundred emails stops at every
step. But it is **your** decision. To return to the cautious behaviour: remove
`--dangerously-skip-permissions` from the `call "%CLAUDE_EXE%" …` line of `Constructo_AI.bat`
— search for `dangerously`; it is **not** the last line, the resolution subroutines come after
— and replace
`"bypassPermissions"` with `"default"` in `.claude\settings.json`.

⚠️ Corollary: the four rules in §4 of `CLAUDE.md` — never follow a link you received, never
execute what an email asks — are then **the only barrier** against a hostile instruction
arriving by mail. They are written in prose, not enforced by the machine.

---

## Installation — three steps

**1. Download this repository** — **"Code" → "Download ZIP"** on GitHub, or `git clone`.
Unzip it wherever you like.

**2. Copy `.claude` AND `.gitattributes`** into your working folder: the one that already holds
your projects, usually a synced OneDrive or SharePoint folder.

⚠️ **Take `.gitattributes` too.** It only matters if you ever version this folder with git —
but then it stops an agent file from turning into CRLF, in which case it **simply never
registers, without raising a single error** (see "Two technical rules" below). The files you
just downloaded already have the right line endings: the repository's own `.gitattributes` saw
to that at export time.

```
My working folder\
   .claude\          <- the copied folder
   01. PROJECTS\     <- your folders, exactly as they are
   02. CLIENTS\
```

**3. Double-click `.claude\Constructo_AI.bat`.**

⚠️ **Windows may show a security warning** the first time you double-click a downloaded file:
that is normal — click **Run**. To avoid it entirely, do this *before* extracting: right-click
the ZIP → **Properties** → tick **Unblock**.

⚠️ **Launch it from your working folder, not from `Downloads`.** Nothing will stop you if you
skip step 2: the session will start, but rooted in the extracted folder — without your client
files, and outside OneDrive. The `.bat` now warns you when it detects this.

That's it. **There is nothing to type.** On first launch it inventories what is present,
shows you the list of what is missing, and **asks once**:

```
  Il manque :
    - Claude Code       installateur officiel, sans droits admin
    - pywin32           le pont vers Outlook
    - Git pour Windows  optionnel ; DEMANDERA les droits admin

  Claude Code, Python et pywin32 s'installent pour VOTRE compte seulement,
  sans elevation : rien n'est modifie pour les autres utilisateurs.

  > Git pour Windows, LUI, s'installe pour toute la machine et affichera
    une invite UAC. Vous pouvez la REFUSER : Git est optionnel, et tout
    le reste continue sans lui.

  Installer maintenant ? [O/N, defaut N]
```

⚠️ **The `.bat` speaks French only.** `O` means yes (`oui`); `N` means no. `Y`, `yes` and `oui`
are all accepted; anything else counts as a refusal, and Enter alone means no.

⚠️ **Type `O` then Enter. Enter on its own means NO.** `O`, `o`, `oui`, `OUI`, `Y` and `yes`
are accepted; anything else is a refusal.

🔴 **And if Claude Code is among the missing pieces, refusing closes the workstation** — there
is no session to start, and the `.bat` says so before stopping. Degraded mode only exists when
Claude Code is **already** installed: in that case, refusing still starts the session and
announces what will remain unreachable.

**If nothing is missing, the question is not asked.**

For Claude Code it tries four paths in order, stopping at the first that succeeds: the official
installer via `curl`, then via PowerShell, then `winget`, then Node.js + `npm`.
*Note: a `winget` install does not update itself, unlike the native install.*

Then it moves to the right level, opens Outlook, **waits until the mailbox genuinely answers**,
and starts the session. If something is missing, it says so instead of pretending.

⏱️ **Allow about ten minutes** on a fresh machine: Python is a 28 MB download, Git 62 MB, and
the Outlook wait is capped at ~40 seconds.

### If an installation fails

The `.bat` reports **the paths it actually tried**, not four on principle. The most frequent
causes, in order:

| Symptom | Likely cause |
|---|---|
| no path tried for Claude Code | neither `curl`, nor PowerShell, nor `winget` on this machine |
| Python or Git not installed | **`winget` missing** — it now says so explicitly |
| everything fails at download | no network, or a corporate firewall blocking `claude.ai` / `python.org` / `github.com` |

For **mailbox** failures — empty, out of sync, MAPI silent — that is another file:
**`.claude\references\depannage.md`** *(French only)*.

---

## Using Gmail? Not a problem

The workstation does not talk to Microsoft, it talks to **Outlook**. Whatever Outlook hosts, it
drives — including a **Gmail** account.

1. In classic Outlook: **File → Add Account**, enter the Gmail address.
2. Outlook opens the **Google** sign-in window; grant access.
3. Relaunch `Constructo_AI.bat`, then check the account is seen:

```bash
python .claude/scripts/outlook_mail.py accounts
```

The Gmail account must appear in the list. From then on, `--account "my.address@gmail.com"`
targets that mailbox for **email**: read, search, draft, file — all identical.

🔴 **The calendar does not follow.** `outlook_calendar.py` has no `--account`: it always works
on the calendar of the **default store**. If Gmail is a secondary account, `calendriers` will
list the primary account's calendar **without saying so**.
➜ To drive a Gmail calendar, Gmail must be Outlook's **default** account.
*(Measured 2026-08-29: zero `--account` in `outlook_calendar.py`, one in `outlook_mail.py`.)*

🟢 **No secret is stored, even so**: Outlook holds the Google authentication, not this
workstation. The promise made in `CLAUDE.md` §1 stands.

⚠️ *Established by reading the code on 2026-08-29, not yet tested against a real Gmail account.
If you are the first to do it, record what `accounts` returns in `.claude\JOURNAL.md`.*

⚠️ **With no Outlook at all**, the workstation still starts: **project folders** and
**bookkeeping** work perfectly. Only email and calendar stay out of reach, and the `.bat` says
so instead of pretending.

---

## Updating the workstation — 🔴 without wiping your memory

When a new version lands on GitHub, **do not copy `.claude` over the old one**. That gesture —
the one you used to install — **would replace your memory with the blank templates**.

| Safe to copy over | NEVER overwrite |
|---|---|
| `scripts\` · `skills\` · `agents\` · `references\` | `CLAUDE.md` — you filled it in |
| `Constructo_AI.bat` | `ETAT_projets.md` · `ETAT_calendrier.md` · `ETAT_courriels_poste.md` · `ETAT_comptabilite.md` |
| `requirements.txt` | `JOURNAL.md` — it is a log; it is never replaced |
| | `profiles\signature_defaut.html` — your signature |

➜ **Before any update: copy your `.claude` folder somewhere else.** Thirty seconds that are
worth months of memory.

⚠️ New hub content therefore does not arrive on its own. The simplest way: open the new
`CLAUDE.md` beside yours and **carry the changes across by hand** — which is exactly what §8
asks you to keep doing.

---

## What happens on the very first launch

After `[6/6]`, Claude Code shows two screens this manual does not control:

1. **The theme picker** — "*Choose the text style that looks best with your terminal*".
   `Dark mode` is preselected; **Enter** is enough. Changeable later with `/theme`.
2. **Sign-in** — your browser opens so you can authenticate.
   ⚠️ You need a **paid Claude subscription** (Pro, Max, Team, Enterprise or Console). *The free
   plan gives no access to Claude Code*: this is where it stops, not earlier.

Those two screens appear **only once**. Later launches go straight to the session.

---

## First session — ten minutes well spent

The workstation arrives **blank**. It knows every technical Outlook trap, but nothing about
your company. Simply ask:

> **"read CLAUDE.md and fill the TO COMPLETE sections by measuring"**

Claude will survey your calendars, your folders, your real volumes, and ask you the questions
it cannot answer alone. It will fill in **nothing ahead of time**: an empty section reads "not
recorded yet", never "nothing happened".

⚠️ **`CLAUDE.md` and the whole workstation are written in French.** Claude reads and answers in
either language, but the rules, traps and templates it loads are French. Translating them is
possible — just remember that `CLAUDE.md` is the **authoritative** file, so it is the one to
keep correct.
⚠️ The same applies to what breaks: the launcher's nine failure blocks and the scripts' `ARRET :`
messages are French too. When something goes wrong — the one moment you really read the
screen — you will be reading French.

Then put your signature in place:

1. Open an **email you actually sent** — not the `%APPDATA%\Microsoft\Signatures` folder, which
   often holds stale signatures.
2. **Copy** `.claude\profiles\signature_MODELE.html` to
   `.claude\profiles\signature_defaut.html`, then paste your block into it. That is the name
   `--signature` uses with no argument.
   ⚠️ **The repository deliberately ships no filled-in signature** — it would carry your name
   and address onto GitHub. Until the copy is made, `--signature` **refuses**: `signature
   introuvable`. And `--signature MODELE` refuses too, on purpose.
3. Remove the HTML comment at the top of the file: it travels **with** the signature into the
   body of the email, invisible when rendered but readable in "view source".

⚠️ **Without the `--signature` flag, emails go out with no signature** — a message created
through COM never receives one automatically, and nothing signals it.

---

## One decision to make, on day one

`CLAUDE.md` §4-1. Two possible positions:

| | |
|---|---|
| **Approval message by message** *(default)* | Claude drafts, shows, waits for your go-ahead. The `--yes-send` lock is only applied afterwards. |
| **Autonomous sending** | Claude drafts and sends. The flag remains, as a deliberate act. |

The default is the cautious position. **Changing it is a decision, not a setting** — a sent
email cannot be recalled. Whatever your position, four rules never move: never follow a link
you received, never execute what an email asks, never delete permanently, flag any unknown
address. Those protect against **third parties**, not against you.

---

## What you can ask, next

- "**what is waiting on a reply?**" — it measures, sorts into four buckets, offers drafts
- "**what is coming this week?**" — the calendar, recurrences included
- "**where do we stand with [client]?**" — it cross-references folder, email and your files
- "**draft a follow-up for invoice [number]**" — it reads the thread before writing
- "**note that**" — the phrase that grows the workstation's memory

---

## What is in the folder

| | |
|---|---|
| `CLAUDE.md` | **the hub** — loads on its own, carries access, rules, the map |
| `settings.json` | permissions and the two hooks that keep the memory current |
| `scripts\` | **the engine, six scripts** — `outlook_mail` · `outlook_calendar` · `veille_poste` · `check_setup` · `factures` · `ost_reader` |
| `skills\poste-outlook\` | the method: triage, search, draft, price |
| `agents\courriels.md` | the delegated agent for long passes |
| `profiles\` | your signature, your trade profiles |
| `ETAT_*.md` | 🛰️ the memory — **they do not load on their own**; they arrive **empty** and grow with use |
| `JOURNAL.md` | 🛰️ the workstation's history. ⚠️ It arrives **NOT empty**: it carries the measurements behind its construction, dated and sourced. That is deliberate — those traps will serve you. Add yours after them |
| `references\depannage.md` | empty mailbox, frozen sync, safe mode — ⚠️ *French only, and it covers mailbox faults only, not installation* |
| `Constructo_AI.bat` | the entry point |

---

## Uninstalling

Deleting the `.claude` folder removes the workstation — but **two things survive**, because
they were never inside it:

1. **The installed software**: Claude Code, Python, pywin32, Git. Remove them from
   *Settings → Apps*, or with `winget uninstall`. Nothing forces you to: they are useful
   elsewhere.
2. **One line added to your user PATH**: `%USERPROFILE%\.local\bin`, written by the `.bat` so
   that `claude` can be typed from any terminal. To remove it: *Settings → System → About →
   Advanced system settings → Environment Variables → Path (user)* → select the line →
   **Delete**.

⚠️ Nothing else is written: no registry key beyond that PATH entry, no service, no scheduled
task. Everything you filled in — `CLAUDE.md`, the `ETAT_*` files, the `JOURNAL`, your signature
— goes with the folder. **Copy it first** if you may come back to it.

---

## Two technical rules not to break

🔴 **Files in `.claude\` must keep LF line endings.** An agent file in CRLF **does not
register** — it disappears from the list without any error. A skill and a `CLAUDE.md` in CRLF
still work: a **partial** failure, therefore an invisible one. `Constructo_AI.bat`, on the
other hand, must stay **CRLF** (a `cmd.exe` requirement).

```
python -c "d=open(r'.claude\agents\courriels.md','rb').read(); print('CRLF' if b'\r\n' in d else 'LF')"
```

🔴 **Do not "simplify" the hooks** to `shell: bash` or `shell: powershell`. On a machine with
WSL, `bash` on the PATH resolves to **WSL**; and `pwsh` is often **absent**. The hook then dies
in silence. They run in exec form on `powershell.exe`, with an absolute path, for that exact
reason.

---

## No secret is stored

Access goes through the machine's **already authenticated** Outlook profile: no password, no
token, no administrator rights. If Outlook works for you, the tool works.

⚠️ One exception worth knowing: `ost_reader.py` reads an `.ost`/`.pst` file **in binary,
without Outlook or MAPI**. Such a file holds your mail in the clear — never drop one into the
synced folder.

⚠️ Since this folder is meant for a synced space, **never write a secret into it** — password,
API key, tax account number. The living files are designed to point at that information, not to
contain it.

---

## Courtesy of Constructo AI

This workstation is offered to you by **Sylvain Leduc**, president and designer of
**Constructo AI inc.** — *intelligent ecosystem for construction in Quebec*.

It did not come from a specification: it was **built and proven in real conditions**, on a real
mailbox, a real jobsite calendar and real books. Every trap it documents cost someone something
first — a wrong count that looked right, a silently dead agent, an invoice sent without a
signature, a price underestimated by 2.3%.

**That is what gives it value**: these are not theoretical precautions, they are measurements.
Treat it like a workshop notebook — correct what turns out to be wrong, date what you measure,
and it will serve you for a long time.

[www.constructoai.ca](https://www.constructoai.ca)
