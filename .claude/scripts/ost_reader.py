"""
ost_reader.py — Lit un fichier Outlook .ost (ou .pst) SANS Outlook.

Extrait l'arborescence des dossiers et l'inventaire des messages (expéditeur,
objet, date, dossier) directement depuis le format binaire. Aucune dépendance.

    python ost_reader.py "C:\\...\\compte.ost" [--limit 40] [--json] [--folders]

À quoi ça sert :
  * prouver qu'un cache est vide (ou plein) sans ouvrir Outlook ;
  * inventorier une boîte sur un poste où Outlook ne démarre plus ;
  * auditer un fichier récupéré d'une sauvegarde.

Portée : format OST/PST « large » 4 Ko (Outlook 2013 et suivants), non chiffré —
le cas courant. Les corps de messages ne sont pas extraits : Exchange les stocke
en RTF compressé ou en HTML, et la synchronisation télécharge souvent les
en-têtes avant les corps. Pour le contenu complet, utiliser `outlook_mail.py`.

Note : si Outlook tourne, il verrouille les 1024 premiers octets du fichier. Le
script le détecte et retrouve les racines des B-trees par balayage.

Licence MIT.
"""
import struct, zlib, datetime, collections, sys, io, os, json, argparse

PAGE, TRAILER_OFF, META_OFF = 4096, 0xFE8, 0xFD8
NID_FOLDER, NID_SEARCH_FOLDER, NID_MESSAGE = 0x02, 0x03, 0x04


class OstReader:
    def __init__(self, path, verbose=True):
        self.path = path
        self.v = verbose
        self.f = open(path, "rb")
        self.f.seek(0, io.SEEK_END)
        self.size = self.f.tell()
        self.nbt, self.bbt = {}, {}
        self._load()

    def log(self, *a):
        if self.v:
            print(*a, file=sys.stderr)

    def rd(self, off, n):
        self.f.seek(off)
        return self.f.read(n)

    # ------------------------------------------------------------- racines
    def _roots(self):
        try:
            h = self.rd(0, 0x100)
            if h[:4] == b"!BDN":
                _, nbt_ib, _, bbt_ib = struct.unpack_from("<QQQQ", h, 0xD8)
                self.log(f"  en-tete lu : NBT=0x{nbt_ib:X} BBT=0x{bbt_ib:X}")
                return nbt_ib, bbt_ib
        except OSError:
            pass
        self.log("  en-tete verrouille (Outlook actif) -> balayage des pages")
        cand, refs, pages = {0x80: [], 0x81: []}, set(), {}
        CH, off = 1 << 22, 0
        while off < self.size:
            try:
                blk = self.rd(off, CH)
            except OSError:
                off += CH
                continue
            if not blk:
                break
            for i in range(0, len(blk) - PAGE + 1, PAGE):
                t = i + TRAILER_OFF
                pt, pr = blk[t], blk[t + 1]
                if pt != pr or pt not in (0x80, 0x81):
                    continue
                bid, = struct.unpack_from("<Q", blk, t + 8)
                cE, = struct.unpack_from("<H", blk, i + META_OFF)
                cb, lvl = blk[i + META_OFF + 4], blk[i + META_OFF + 5]
                if cb < 24 or cE == 0 or cE * cb > META_OFF:
                    continue                       # page perimee
                pages[off + i] = (pt, bid)
                if lvl > 0:
                    for k in range(cE):
                        _, _, ib = struct.unpack_from("<QQQ", blk, i + k * cb)
                        refs.add(ib)
            off += CH
        for o, (pt, bid) in pages.items():
            if o not in refs:
                cand[pt].append((bid, o))
        if not cand[0x81] or not cand[0x80]:
            raise SystemExit("racines B-tree introuvables — fichier illisible "
                             "(format non supporté, chiffré, ou corrompu)")
        return max(cand[0x81])[1], max(cand[0x80])[1]

    def _btwalk(self, off, want):
        cE, = struct.unpack("<H", self.rd(off + META_OFF, 2))
        cb, lvl = self.rd(off + META_OFF + 4, 1)[0], self.rd(off + META_OFF + 5, 1)[0]
        if cb < 24 or cE == 0 or cE * cb > META_OFF:
            return
        buf = self.rd(off, cE * cb)
        if lvl == 0:
            for k in range(cE):
                yield buf[k * cb:(k + 1) * cb]
        else:
            for k in range(cE):
                _, _, ib = struct.unpack_from("<QQQ", buf, k * cb)
                try:
                    if self.rd(ib + TRAILER_OFF, 1)[0] == want:
                        yield from self._btwalk(ib, want)
                except (OSError, IndexError):
                    continue

    def _load(self):
        nbt_root, bbt_root = self._roots()
        for e in self._btwalk(nbt_root, 0x81):
            nid, bd, bs, par = struct.unpack("<QQQI", e[:28])
            self.nbt[nid & 0xFFFFFFFF] = (bd, bs, par)
        for e in self._btwalk(bbt_root, 0x80):
            bid, ib, cb, _ = struct.unpack("<QQHH", e[:20])
            self.bbt[bid & ~1] = (ib, cb)
        self.log(f"  noeuds : {len(self.nbt):,}   blocs : {len(self.bbt):,}")

    # ------------------------------------------------------------- blocs
    @staticmethod
    def _inflate(b):
        if len(b) > 2 and b[0] == 0x78 and b[1] in (0x01, 0x5E, 0x9C, 0xDA):
            try:
                return zlib.decompress(b)
            except zlib.error:
                pass
        return b

    def blocks(self, bid):
        if bid == 0:
            return []
        e = self.bbt.get(bid & ~1)
        if e is None:
            return []
        buf = self.rd(e[0], e[1])
        if bid & 0x02 and len(buf) >= 8 and buf[0] == 0x01:      # XBLOCK
            cEnt, = struct.unpack_from("<H", buf, 2)
            out = []
            for k in struct.unpack_from(f"<{cEnt}Q", buf, 8):
                out.extend(self.blocks(k))
            return out
        return [self._inflate(buf)]

    def subnodes(self, bs):
        out = {}
        if bs == 0:
            return out
        e = self.bbt.get(bs & ~1)
        if e is None:
            return out
        buf = self.rd(e[0], e[1])
        if len(buf) < 8 or buf[0] != 0x02:
            return out
        cEnt, = struct.unpack_from("<H", buf, 2)
        if buf[1] == 0:                                          # SLBLOCK
            for i in range(cEnt):
                if 8 + i * 24 + 24 > len(buf):
                    break
                nid, bd, sb = struct.unpack_from("<QQQ", buf, 8 + i * 24)
                out[nid & 0xFFFFFFFF] = (bd, sb)
        else:                                                    # SIBLOCK
            for i in range(cEnt):
                if 8 + i * 16 + 16 > len(buf):
                    break
                _, b2 = struct.unpack_from("<QQ", buf, 8 + i * 16)
                out.update(self.subnodes(b2))
        return out

    # ------------------------------------------------------------- proprietes
    FIXED = {0x0002: 2, 0x0003: 4, 0x0004: 4, 0x000A: 4, 0x000B: 2}

    def read_pc(self, nid):
        bd, bs, _ = self.nbt[nid]
        blks = self.blocks(bd)
        if not blks:
            return {}
        hdr = blks[0]
        ibHnpm, sig, _, hidUserRoot = struct.unpack_from("<HBBI", hdr, 0)
        if sig != 0xEC:
            raise ValueError(f"signature HN 0x{sig:02X}")
        subs = self.subnodes(bs)

        def hget(hid):
            if hid == 0 or (hid & 0x1F):
                return b""
            idx, bi = (hid >> 5) & 0x7FF, hid >> 16
            if bi >= len(blks) or idx < 1:
                return b""
            blk = blks[bi]
            ib = ibHnpm if bi == 0 else struct.unpack_from("<H", blk, 0)[0]
            cAlloc, = struct.unpack_from("<H", blk, ib)
            rg = struct.unpack_from(f"<{cAlloc + 1}H", blk, ib + 4)
            return blk[rg[idx - 1]:rg[idx]] if idx < len(rg) else b""

        def value(pt, hnid):
            if pt in self.FIXED:
                raw = struct.pack("<I", hnid & 0xFFFFFFFF)
                if pt == 0x0003:
                    return struct.unpack("<i", raw)[0]
                if pt == 0x0002:
                    return struct.unpack("<h", raw[:2])[0]
                if pt == 0x000B:
                    return bool(hnid & 0xFFFF)
                return hnid
            if hnid == 0:
                return None
            if (hnid & 0x1F) == 0:
                b = hget(hnid)
            else:
                s = subs.get(hnid & 0xFFFFFFFF)
                b = b"".join(self.blocks(s[0])) if s else None
            if b is None:
                return None
            if pt == 0x001F:
                return b.decode("utf-16-le", "replace")
            if pt == 0x001E:
                return b.decode("cp1252", "replace")
            if pt == 0x0040 and len(b) >= 8:
                v, = struct.unpack("<Q", b[:8])
                try:
                    return datetime.datetime(1601, 1, 1) + datetime.timedelta(microseconds=v // 10)
                except Exception:
                    return None
            if pt in (0x0014, 0x0006) and len(b) >= 8:
                return struct.unpack("<q", b[:8])[0]
            return b

        root = hget(hidUserRoot)
        _, cbKey, cbEnt, lvls, hidRoot = struct.unpack_from("<BBBBI", root, 0)
        props = {}

        def rec(hid, lv):
            buf = hget(hid)
            if not buf:
                return
            if lv > 0:
                w = cbKey + 4
                for i in range(len(buf) // w):
                    nxt, = struct.unpack_from("<I", buf, i * w + cbKey)
                    rec(nxt, lv - 1)
            else:
                w = cbKey + cbEnt
                for i in range(len(buf) // w):
                    pid, = struct.unpack_from("<H", buf, i * w)
                    pt, hnid = struct.unpack_from("<HI", buf, i * w + cbKey)
                    try:
                        props[pid] = value(pt, hnid)
                    except Exception:
                        props[pid] = None

        rec(hidRoot, lvls)
        return props

    # ------------------------------------------------------------- API
    def folders(self):
        out = {}
        for nid in self.nbt:
            if (nid & 0x1F) in (NID_FOLDER, NID_SEARCH_FOLDER):
                try:
                    out[nid] = self.read_pc(nid).get(0x3001) or ""
                except Exception:
                    out[nid] = "?"
        return out

    def folder_path(self, nid, folders):
        parts, cur, seen = [], nid, set()
        while cur in folders and cur not in seen:
            seen.add(cur)
            if folders[cur]:
                parts.append(folders[cur])
            cur = self.nbt[cur][2]
        return "/".join(reversed(parts)) or "(racine)"

    def messages(self):
        folders = self.folders()
        rows, errors = [], 0
        for nid in self.nbt:
            if (nid & 0x1F) != NID_MESSAGE:
                continue
            try:
                p = self.read_pc(nid)
            except Exception:
                errors += 1
                continue
            subj = p.get(0x0037) or ""
            if isinstance(subj, str) and subj[:1] == "\x01":
                subj = subj[2:] if len(subj) > 2 else ""
            d = p.get(0x0E06) or p.get(0x0039)
            rows.append(dict(
                folder=self.folder_path(self.nbt[nid][2], folders),
                subject=subj,
                sender=p.get(0x0C1A) or p.get(0x0042) or "",
                sender_addr=p.get(0x0C1F) or "",
                date=d.isoformat(" ") if isinstance(d, datetime.datetime) else None,
                unread=not ((p.get(0x0E07) or 0) & 0x01),
                size=p.get(0x0E08) or 0))
        rows.sort(key=lambda r: r["date"] or "", reverse=True)
        return rows, errors, folders


def main():
    ap = argparse.ArgumentParser(description="Lecteur de fichier Outlook .ost/.pst")
    ap.add_argument("path", help="chemin du fichier .ost ou .pst")
    ap.add_argument("--limit", type=int, default=40, help="messages à afficher")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--folders", action="store_true", help="dossiers seulement")
    a = ap.parse_args()

    if not os.path.exists(a.path):
        sys.exit(f"fichier introuvable : {a.path}")
    size = os.path.getsize(a.path)
    print(f"fichier : {size:,} octets", file=sys.stderr)

    r = OstReader(a.path)
    rows, errors, folders = r.messages()

    if a.json:
        print(json.dumps(dict(file=a.path, size=size, folders=len(folders),
                              messages=len(rows), errors=errors,
                              items=rows[:a.limit]), ensure_ascii=False, indent=2))
        return

    print(f"\ndossiers : {len(folders)}   messages : {len(rows):,}   erreurs : {errors}")
    print("\n" + "=" * 78 + "\nPAR DOSSIER\n" + "=" * 78)
    for k, v in collections.Counter(x["folder"] for x in rows).most_common():
        print(f"  {v:6,}   {k}")
    if a.folders:
        return
    print("\n" + "=" * 78 + f"\n{min(a.limit, len(rows))} MESSAGES LES PLUS RECENTS\n" + "=" * 78)
    for x in rows[:a.limit]:
        print(f"\n{'*' if x['unread'] else ' '} {x['date'] or '?'}  [{x['folder']}]")
        print(f"    De    : {x['sender']}  <{x['sender_addr']}>")
        print(f"    Objet : {x['subject']}")
    if not rows:
        print("\nAucun message : ce cache n'a jamais été synchronisé.")
        print("Voir .claude/references/depannage.md")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    main()
