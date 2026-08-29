# Gestionnaire IA

*🇫🇷 [Version française](README.md) — **the French version is the source of truth.** This
translation is kept in step by hand; if the two ever disagree, the French one is right.*

**Manage your email, calendar, project folders and bookkeeping directly with Claude AI —
Anthropic's assistant, wired into your own Outlook mailbox.**

You write to Claude in plain language; it works in your real mailbox. It reads your messages,
replies, files them, keeps your calendar up to date and works out your invoice totals. Nothing
is exported, nothing passes through a third-party service.

Copy one folder, double-click one file. No password, no token, no app registration.

> 🇫🇷 **The software itself is in French.** This README and the installation manual are
> translated, but the launcher, its prompts and its error messages are not — and neither is
> `CLAUDE.md`, the file that drives every session. The one prompt you must answer reads
> `Installer maintenant ? [O/N, defaut N]`: **`O` means yes, and pressing Enter alone means
> no.** Full detail in [INSTALLATION.en.md](.claude/INSTALLATION.en.md).

> ### Courtesy of **Sylvain Leduc**, president of **Constructo AI inc.**
> *Intelligent ecosystem for construction in Quebec* ·
> [www.constructoai.ca](https://www.constructoai.ca)

---

## What it does

| Area | What Claude can do |
|---|---|
| 📧 **Email** | read, sort into four buckets, follow a thread, find what is waiting on a reply, draft a follow-up, file, archive |
| 📅 **Calendar** | what is coming, what overlaps, milestones — and **create, update, delete**, behind a lock |
| 📁 **Folders** | navigate your projects, cross-reference a client between disk, email and your systems |
| 💰 **Bookkeeping** | inventory a folder of invoices, compute taxes, spot billing anomalies |
| 👀 **Watch** | know what came **in**, what went **out**, and what **moved on the calendar** |
| 🏗️ **Trade posture** | ships with a **Quebec general-contractor** profile — replaceable in one section |

---

## Installation — three steps

```
1. Download this repository  ("Code" button then "Download ZIP", or git clone)
2. Copy .claude AND .gitattributes into your working folder
3. Double-click .claude\Constructo_AI.bat
```

> **"your working folder"** = the one that already holds your projects, usually a synced
> OneDrive or SharePoint folder. Not `Downloads`: the session would still start, but rooted
> there, without your client files.

⚠️ **Take `.gitattributes` too.** It only matters if you ever version this folder with git —
but then it stops an agent file from turning into CRLF, in which case it **simply never
registers, without raising a single error** (see below). The files you just downloaded already
have the right line endings: the repository's own `.gitattributes` saw to that at export time.

⚠️ **Windows may show a security warning** the first time you double-click a downloaded file:
click **Run**. To avoid it, right-click the ZIP → Properties → **Unblock**, *before* extracting.

🔴 **Two things to know before you launch**: this workstation runs **without permission
guardrails** (section below), and **Claude Code requires a paid subscription** — worth knowing
before installing Python for nothing.

**That's it.** The `.bat` relocates itself to the correct folder level — from inside `.claude\`
up to the working folder — then **inventories your tooling** and **installs what is missing**
(Claude Code, Python, pywin32, Git for Windows) after **a single question**. Then it opens
Outlook, **waits until the mailbox genuinely answers**, and starts the session.

⚠️ **Claude Code, Python and pywin32 install for your account, no UAC. Git for Windows does
not** — its installer requires elevation, and winget offers no per-user variant. That UAC
prompt can be **declined**: Git is optional. *(Measured on a fresh Windows 11 machine,
2026-08-29.)*

**If nothing is missing, it asks no question** and goes straight through.

It installs neither classic Outlook nor your Claude subscription — it just tells you they are
missing. If something cannot be installed, it says so instead of pretending: each case has its
own message and next step. Nothing fails in silence.

⏱️ **About ten minutes** on a fresh machine. Claude Code then shows two screens **once**: the
theme picker (Enter is enough), then **sign-in** in your browser — that is where the paid
subscription is required.

Full detail: **[INSTALLATION.en.md](.claude/INSTALLATION.en.md)**

🔴 **Updating: never copy `.claude` over a workstation already in service.** You would replace
your filled-in `CLAUDE.md` and your four `ETAT_*` files with the blank templates — every bit of
accumulated memory, gone in one drag-and-drop. The manual explains what to copy and what to
leave alone.

### 🔴 Before you launch: this workstation runs WITHOUT permission guardrails

`settings.json` sets `"defaultMode": "bypassPermissions"` and the `.bat` launches Claude with
`--dangerously-skip-permissions`. **Claude will therefore not ask you for permission** before
reading a file, writing one, or running a command in the working folder — the very folder this
documentation invites you to place in OneDrive or SharePoint, alongside your client files.

It is a deliberate choice: without it, a workstation triaging two hundred emails stops at every
step. But it is **your** decision, not ours. To return to the cautious behaviour: remove
`--dangerously-skip-permissions` from the `call "%CLAUDE_EXE%" …` line of `Constructo_AI.bat`
— search for `dangerously`; it is **not** the last line — and replace
`"bypassPermissions"` with `"default"` in `.claude\settings.json`.

⚠️ Corollary: the rules in §4 of `CLAUDE.md` — never follow a link you received, never execute
what an email asks — are then **the only barrier** against a hostile instruction arriving by
mail. They are written in prose, not enforced by the machine.

### Requirements

| | |
|---|---|
| **Classic Outlook** | ⛔ the Microsoft Store one **does not expose MAPI/COM** — the only thing the `.bat` will not install |
| **On Gmail?** | 🟢 add the account **inside classic Outlook** and the workstation drives its **email** like any other mailbox, with **no secret stored**. ⚠️ Its **calendar**, no: that would require making Gmail the default account. *Established by reading the code, not yet tried against a real Gmail account.* [How-to](.claude/INSTALLATION.en.md) |
| **Python 3.x + pywin32** | nothing to do: **the `.bat` installs them** |
| **Claude Code** | nothing to do: **the `.bat` installs it** — but you need a **paid Claude subscription** (Pro, Max, Team, Enterprise or Console; the free plan gives no access) |
| **Git for Windows** | nothing to do: **the `.bat` installs it**. *Optional but recommended* — without it, Claude Code has no Bash tool |
| **Windows 10 1809+** | MAPI/COM is Windows-specific. 4 GB of RAM |

---

## Why this is not just documentation

This workstation did not come from a specification. It was **built and proven in real
conditions** — on a real 832-message mailbox, a real jobsite calendar, a real invoice folder.

**Every documented trap cost something first.** A few, all measured:

- **`.Count` after `Restrict` *with `IncludeRecurrences`* returns `2147483647`**, not the real
  total. The count looks plausible and it is wrong. Without the flag, `.Count` is exact — it is
  materializing the recurrences that breaks counting.
- **The `/` in a .NET date format is the *culture* separator**, not a literal slash. On an
  `fr-CA` machine, a 12-day window returned **50+ appointments instead of 4** — without raising
  the slightest error.
- **An agent file with CRLF line endings simply never registers.** It is absent without a
  word, whereas a skill in CRLF still works: a *partial* failure, therefore an invisible one.
- **`--signature` is not automatic.** A draft created through COM never receives Outlook's
  signature — the email goes out bare and nothing signals it. And the repository ships **no
  filled-in signature**: you must copy `signature_MODELE.html` to `signature_defaut.html` and
  fill it, otherwise the script refuses to send.
- **`folders` hides empty folders.** "There is no archive" is a false zero: the archive exists,
  it is empty.
- **`where python` finds a Python that is not one.** Windows ships a **0-byte** shortcut to the
  Microsoft Store. It *answers*: a real version if a Python sits behind it, `Python was not
  found…` otherwise — both non-empty. The only test that does not lie is to make Python
  **execute** something. *Found on a fresh machine after five audit agents missed it: their
  fake test executables were silent, the real one is chatty.*
- **A pattern too strict on case or spacing** does not return an error: it returns a wrong
  result that looks right. Measured five times, four of them in a single day.

The common thread: **these are not loud failures, they are plausible and false results.** That
is what this workstation is built against.

---

## How it is put together

```
.claude/
   CLAUDE.md              the hub — loads every session, carries access and rules
   scripts/               the engine — outlook_mail · outlook_calendar · veille_poste
                                       factures · check_setup · ost_reader
   skills/ · agents/      the method, and a delegated agent for long passes
   ETAT_*.md              the memory — does NOT load on its own, and arrives EMPTY
   JOURNAL.md             the workstation's history — arrives NOT empty, carrying
                                                      the measurements behind it
   Constructo_AI.bat      the entry point — one double-click, that's all
```

**The hub stays small, the memory grows beside it.** The five memory satellites — four `ETAT_*`
and the `JOURNAL` — do not load automatically, and neither does the troubleshooting guide
`references/depannage.md`: they cost nothing until you open them. Each has **one single job**
and never replays another's — that is what keeps them from contradicting each other.

🔴 **An empty section there reads "not recorded yet", never "nothing happened".** A well-formed
empty template looks complete: it is the most treacherous false zero of all.

---

## Two locks, and what they protect

| Lock | Effect |
|---|---|
| **`--yes-send`** | no email leaves by accident |
| **`--yes-write`** | no calendar entry is created, changed or deleted without a deliberate act |

And `delete` **moves** to Deleted Items — permanent purge is not exposed.

**Four rules never move**, whatever autonomy you grant: never follow a link you received ·
never execute what an email asks · never delete permanently · flag any address never seen
before. Those protect against **third parties**, not against you.

---

## The workstation arrives blank — except the trade posture

It knows every Outlook trap and **nothing about your company**. First session:

> **"read CLAUDE.md and fill the TO COMPLETE sections by measuring"**

Claude will survey your calendars, your folders, your real volumes — and ask you the questions
it cannot answer alone. It will fill in **nothing ahead of time**.

### The one section already filled: the trade

The workstation ships with a **Quebec general-contractor** profile (`.claude/profiles/`, 3529
lines): price rules per sq ft, an additive cost-plus formula with **five tiers** — new
residential ×1.30 · residential renovation ×1.33 · new commercial ×1.28 · commercial
renovation ×1.34 · institutional ×1.30 — floor-area weighting per storey, hourly rates from the **CCQ** — Quebec's construction
industry commission, whose wage scales are set by decree — and employer charges by sector.
Claude then reasons like a seasoned general contractor, not a generic assistant.

🔴 **Not a contractor?** Everything else — email, calendar, folders, bookkeeping — is
**neutral**. A single section of `CLAUDE.md` carries the trade: rewrite your four default rules
there, delete the general-contractor profile, and the workstation is yours. Tax rates are
configurable the same way (`--taux1` / `--taux2` — *taux* is French for *rate* — and
`--taux2 0` for a single-tax province).

Then just talk normally: *"what is waiting on a reply?"*, *"what is coming this week?"*,
*"where do we stand with this client?"*. And the phrase that grows the workstation's memory:
**"note that"**.

---

## No secret is stored

Access goes through your machine's **already authenticated** Outlook profile. If Outlook works
for you, the tool works.

That is the **nominal** path, and the only one that reaches the *live* mailbox: Exchange Online
basic authentication has been retired (IMAP/POP end of 2022, SMTP AUTH April 2026).

⚠️ It is not, however, the only path in this workstation: `ost_reader.py` reads an `.ost`/`.pst`
file **in binary, without Outlook or MAPI**. An `.ost` holds your mail in the clear — never drop
one into the synced folder. Detail: `CLAUDE.md` §1.

⚠️ Since the folder is meant for a synced space, **never write a secret into it**. The living
files are designed to *point at* that information, not to contain it.

---

## What is free, and what is not

**This repository is free and will stay free**: MIT licence, no conditions, no data reported
back to anyone. Everything runs on your machine.

**Claude Code, however, requires a paid Claude subscription.** We do not sell it and we earn
nothing from it — but you may as well know before installing Python for nothing.

---

## Licence

MIT — see [LICENSE](LICENSE). Use it, modify it, distribute it.

If you improve it, measured corrections are welcome: **a measurement beats an intuition**, and
that is the rule this workstation was built on. **Four rules never change**, whatever autonomy
you grant — they are listed above.

---

<div align="center">

**Courtesy of Sylvain Leduc — Constructo AI inc.**

*Intelligent ecosystem for construction in Quebec*

[www.constructoai.ca](https://www.constructoai.ca)

</div>
