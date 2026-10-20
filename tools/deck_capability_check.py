#!/usr/bin/env python3
"""deck_capability_check.py - PitWall capability scan for a .pptx deck (the deck analogue of Turbo's nav_capability_check.py).

A symbol scan of the deck's text: no rendering, no judgment of quality. Reads the capability lines of PITWALL/REGRESSION.md
(plus PITWALL/capabilities/*.md and PITWALL/board/*.md when present), checks each against the slides, prints PASS or FAIL
per line and one RESULT line; the exit code is 1 when anything fails. Stdlib plus python-pptx; ASCII only.

Line grammar (one per line, anywhere in the file; a leading "-" or "*" is allowed):
  capability: <words> => <regex>[, <regex>]      PASS when any regex matches the text of at least one slide (case-insensitive)
  forbid: <words> => <regex>[, <regex>]          PASS when no slide matches any regex
  check: <words> => <key>[=<value>]              structural check: max_slides=N | max_presented=N | notes | validates | no_raster_charts
  (max_presented counts the slides before the first divider slide whose text starts with "Backup" or "Appendix")
A "capability:" line whose right-hand side starts with "check:" is treated as a check line.

usage: python3 -I tools/deck_capability_check.py <deck.pptx> [--pitwall DIR] [--quiet] [--self-test]
  --pitwall  the PITWALL folder (default: <repo>/PITWALL)
  --quiet    print the FAIL lines and the RESULT only
"""
import argparse
import os
import re
import subprocess
import sys
import tempfile

_LINE = re.compile(r"^\s*(?:[-*]\s*)?(capability|cap|forbid|check)\s*:\s*(.+?)\s*(?:=>|->)\s*(.+?)\s*$", re.I)
_CHECK_KEYS = ("max_slides", "max_presented", "notes", "validates", "no_raster_charts")
VALIDATOR = os.environ.get("PPTX_VALIDATOR")  # optional: path to the pptx skill's office/validate.py


# ---------------------------------------------------------------- the deck
def read_deck(path):
    """-> list of dicts per slide: text (all shapes and tables), notes, n_pictures."""
    from pptx import Presentation
    from pptx.util import Emu  # noqa: F401  (import check only)

    prs = Presentation(path)
    slides = []
    for i, s in enumerate(prs.slides, 1):
        parts = []
        pics = 0
        for sh in s.shapes:
            if sh.shape_type == 13:  # PICTURE
                pics += 1
            if sh.has_text_frame:
                parts.append(sh.text_frame.text)
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    parts.append(" | ".join(c.text for c in row.cells))
        notes = ""
        if s.has_notes_slide:
            notes = s.notes_slide.notes_text_frame.text.strip()
        slides.append({"n": i, "text": "\n".join(parts), "notes": notes, "pics": pics})
    return slides


# ---------------------------------------------------------------- the lists
def read_lines(pitwall):
    files = []
    if pitwall and os.path.isdir(pitwall):
        reg = os.path.join(pitwall, "REGRESSION.md")
        if os.path.isfile(reg):
            files.append(reg)
        for sub in ("capabilities", "board"):
            d = os.path.join(pitwall, sub)
            if os.path.isdir(d):
                for n in sorted(os.listdir(d)):
                    if n.endswith((".md", ".txt")):
                        files.append(os.path.join(d, n))
    items = []
    for f in files:
        with open(f, encoding="utf-8-sig", errors="replace") as fh:
            for line in fh:
                m = _LINE.match(line)
                if not m:
                    continue
                kind, name, rhs = m.group(1).lower(), m.group(2), m.group(3)
                if kind == "cap":
                    kind = "capability"
                if rhs.lower().startswith("check:"):
                    kind, rhs = "check", rhs[6:].strip()
                items.append((os.path.basename(f), kind, name, rhs))
    return items, files


def split_patterns(rhs):
    # split on ", " only when the comma is followed by a space and the next char is not inside a regex group
    out, depth, cur = [], 0, ""
    for ch in rhs:
        if ch in "([{":
            depth += 1
        elif ch in ")]}":
            depth = max(0, depth - 1)
        if ch == "," and depth == 0:
            out.append(cur.strip())
            cur = ""
        else:
            cur += ch
    if cur.strip():
        out.append(cur.strip())
    return [p for p in out if p]


def _compile(p):
    try:
        return re.compile(p, re.I | re.S)
    except re.error as e:
        return e


# ---------------------------------------------------------------- the check
def check(slides, items, deck_path=None):
    """-> list of (status, name, detail)."""
    results = []
    core = [s for s in slides if len(s["text"].strip()) > 40]  # divider and title cards are not "core"
    for src, kind, name, rhs in items:
        label = "[%s] %s" % (src, name)
        if kind in ("capability", "forbid"):
            pats = split_patterns(rhs)
            hits = []
            bad = []
            for p in pats:
                rx = _compile(p)
                if isinstance(rx, re.error):
                    bad.append("%s (%s)" % (p, rx))
                    continue
                for s in slides:
                    m = rx.search(s["text"])
                    if m:
                        snippet = re.sub(r"\s+", " ", m.group(0))[:60]
                        hits.append((s["n"], p, snippet))
            if bad:
                results.append(("FAIL", label, "bad regex: " + "; ".join(bad)))
            elif kind == "capability":
                if hits:
                    n, p, snip = hits[0]
                    results.append(("PASS", label, "slide %d: '%s'" % (n, snip)))
                else:
                    results.append(("FAIL", label, "no slide matches: " + " | ".join(pats)))
            else:  # forbid
                if hits:
                    where = ", ".join("slide %d '%s'" % (n, snip) for n, _, snip in hits[:4])
                    results.append(("FAIL", label, "found on " + where))
                else:
                    results.append(("PASS", label, "absent"))
        elif kind == "check":
            key, _, val = rhs.partition("=")
            key = key.strip().lower()
            if key == "max_slides":
                try:
                    limit = int(val)
                except ValueError:
                    results.append(("FAIL", label, "max_slides needs a number"))
                    continue
                st = "PASS" if len(slides) <= limit else "FAIL"
                results.append((st, label, "%d slides, limit %d" % (len(slides), limit)))
            elif key == "max_presented":
                try:
                    limit = int(val)
                except ValueError:
                    results.append(("FAIL", label, "max_presented needs a number"))
                    continue
                presented = len(slides)
                for s in slides:
                    if re.match(r"\s*(backup|appendix)\b", s["text"], re.I) and len(s["text"].strip()) < 400:
                        presented = s["n"] - 1
                        break
                st = "PASS" if presented <= limit else "FAIL"
                results.append((st, label, "%d presented slides before the backup divider, limit %d" % (presented, limit)))
            elif key == "notes":
                missing = [s["n"] for s in core if not s["notes"]]
                if missing:
                    results.append(("FAIL", label, "%d of %d core slides have no notes: %s" % (
                        len(missing), len(core), ", ".join(map(str, missing)))))
                else:
                    results.append(("PASS", label, "all %d core slides have notes" % len(core)))
            elif key == "no_raster_charts":
                withpics = [s["n"] for s in slides if s["pics"]]
                if withpics:
                    results.append(("FAIL", label, "raster picture(s) on slide(s) " + ", ".join(map(str, withpics))))
                else:
                    results.append(("PASS", label, "no raster pictures"))
            elif key == "validates":
                st, why = validate(deck_path)
                results.append((st, label, why))
            else:
                results.append(("FAIL", label, "unknown check key '%s' (known: %s)" % (key, ", ".join(_CHECK_KEYS))))
    return results


def validate(deck_path):
    """OOXML schema validation through the pptx skill's validator when available, else a minimal zip/XML sanity check."""
    if not deck_path:
        return "FAIL", "no deck path"
    v = VALIDATOR
    if not v:
        for root in (os.path.expanduser("~/.claude/skills"),):
            for dp, dn, fn in os.walk(root):
                if "validate.py" in fn and dp.endswith(os.path.join("scripts", "office")):
                    v = os.path.join(dp, "validate.py")
                    break
                if v:
                    break
    if v and os.path.isfile(v):
        try:
            r = subprocess.run([sys.executable, v, deck_path], capture_output=True, text=True, timeout=240)
        except Exception as e:  # noqa: BLE001
            return "FAIL", "validator did not run: %s" % e
        out = (r.stdout + r.stderr).strip().splitlines()
        last = out[-1] if out else ""
        if r.returncode == 0 and "PASSED" in (r.stdout + r.stderr):
            return "PASS", "schema validation passed (%s)" % os.path.basename(os.path.dirname(os.path.dirname(v)))
        return "FAIL", "schema validation: " + (last[:160] or "exit %d" % r.returncode)
    # fallback: python-pptx could open it (read_deck already did) and the embedded font list is ordered
    import zipfile
    try:
        with zipfile.ZipFile(deck_path) as z:
            x = z.read("ppt/presentation.xml").decode("utf-8", "replace")
    except Exception as e:  # noqa: BLE001
        return "FAIL", "cannot read ppt/presentation.xml: %s" % e
    m = re.search(r"<[\w:]*embeddedFontLst>(.*?)</[\w:]*embeddedFontLst>", x, re.S)
    if m:
        for ent in re.findall(r"<[\w:]*embeddedFont>(.*?)</[\w:]*embeddedFont>", m.group(1), re.S):
            order = re.findall(r"<[\w:]*(regular|bold|italic|boldItalic)\b", ent)
            want = [k for k in ("regular", "bold", "italic", "boldItalic") if k in order]
            if order != want:
                return "FAIL", "embedded font entries out of schema order: %s (no full validator found)" % order
    return "PASS", "opened by python-pptx; embedded font order ok (no full validator found: set PPTX_VALIDATOR for the schema check)"


def print_results(results, files, quiet):
    npass = sum(1 for r in results if r[0] == "PASS")
    nfail = sum(1 for r in results if r[0] == "FAIL")
    for st, name, why in results:
        if quiet and st == "PASS":
            continue
        print("%s  %s  (%s)" % (st, name, why))
    print("capability lines read: %d from %d file(s)%s" % (
        len(results), len(files), "" if results else " (none: add lines 'capability: <words> => <regex>' to PITWALL/REGRESSION.md)"))
    print("RESULT %d pass, %d fail" % (npass, nfail))
    return nfail


# ---------------------------------------------------------------- self test
def self_test():
    from pptx import Presentation

    tmp = tempfile.mkdtemp()
    deck = os.path.join(tmp, "t.pptx")
    prs = Presentation()
    for txt in ("The ask: authorize $25M across two gates. This is the first content slide with enough text to count.",
                "Risks: a reversal trigger is stated here. Seamless is a banned word that must be caught by the guard."):
        s = prs.slides.add_slide(prs.slide_layouts[5])
        s.shapes.title.text = txt
    prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = "Backup"
    prs.slides[0].notes_slide.notes_text_frame.text = "notes"
    prs.save(deck)
    pit = os.path.join(tmp, "PITWALL")
    os.makedirs(pit)
    with open(os.path.join(pit, "REGRESSION.md"), "w") as fh:
        fh.write("- capability: ask => authorize, approve\n"
                 "- capability: missing thing => unobtainium\n"
                 "- forbid: banned word => seamless\n"
                 "- forbid: absent word => nonexistentword\n"
                 "- check: size => max_slides=5\n"
                 "- capability: notes => check: notes\n"
                 "- check: native => no_raster_charts\n"
                 "- check: presented => max_presented=2\n"
                 "- check: presented tight => max_presented=1\n")
    slides = read_deck(deck)
    items, files = read_lines(pit)
    res = check(slides, items, deck)
    got = [r[0] for r in res]
    want = ["PASS", "FAIL", "FAIL", "PASS", "PASS", "FAIL", "PASS", "PASS", "FAIL"]
    ok = got == want and len(files) == 1
    print("self-test %s: %s (want %s)" % ("PASS" if ok else "FAIL", got, want))
    return 0 if ok else 1


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("deck", nargs="?")
    ap.add_argument("--pitwall", default=None)
    ap.add_argument("--quiet", action="store_true")
    ap.add_argument("--self-test", action="store_true")
    a = ap.parse_args()
    if a.self_test:
        return self_test()
    if not a.deck or not os.path.isfile(a.deck):
        print("FAIL  deck not found: %s" % a.deck)
        print("RESULT 0 pass, 1 fail")
        return 1
    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    pitwall = a.pitwall or os.environ.get("DECK_PITWALL") or os.path.join(repo, "PITWALL")
    slides = read_deck(a.deck)
    items, files = read_lines(pitwall)
    print("deck: %s (%d slides); lists: %s" % (a.deck, len(slides), ", ".join(os.path.relpath(f, repo) for f in files) or "none"))
    results = check(slides, items, a.deck)
    return 1 if print_results(results, files, a.quiet) else 0


if __name__ == "__main__":
    sys.exit(main())
