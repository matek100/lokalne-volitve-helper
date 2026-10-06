"""Pomožne funkcije za izpolnjevanje praznih mest (podčrtajev) v Word obrazcih.

Besedilo odstavka je v Wordu razdeljeno na več "runov" s svojim oblikovanjem. Funkcije tukaj
delajo nad zlepljenim besedilom odstavka in spremembe preslikajo nazaj v rune, tako da se
oblikovanje ohrani.
"""
import copy
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BLANK = re.compile(r"_{2,}")

# vrstni red elementov v w:rPr po shemi OOXML (za pravilno vstavljanje w:bdr)
_RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
              "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
              "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
              "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
              "specVanish", "oMath"]


def text(p):
    return "".join(r.text for r in p.runs)


def _offsets(p):
    out, pos = [], 0
    for r in p.runs:
        t = r.text
        out.append((r, pos, pos + len(t)))
        pos += len(t)
    return out


def replace_span(p, start, end, value):
    """Zamenja znake [start, end) v besedilu odstavka z `value` (oblika prvega runa)."""
    first = True
    for r, s, e in _offsets(p):
        if e <= start or s >= end:
            continue
        t = r.text
        a, b = max(start, s) - s, min(end, e) - s
        if first:
            r.text = t[:a] + value + t[b:]
            first = False
        else:
            r.text = t[:a] + t[b:]


def isolate(p, start, end):
    """Razdeli rune tako, da [start, end) pokrivajo točno celi runi; vrne te rune."""
    result = []
    for r, s, e in _offsets(p):
        if e <= start or s >= end:
            continue
        t = r.text
        a, b = max(start, s) - s, min(end, e) - s
        el = r._r
        if b < len(t):  # odreži rep
            tail = copy.deepcopy(el)
            el.addnext(tail)
            r.text = t[:b]
            type(r)(tail, r._parent).text = t[b:]
            t = t[:b]
        if a > 0:  # odreži glavo
            head = copy.deepcopy(el)
            el.addprevious(head)
            type(r)(head, r._parent).text = t[:a]
            r.text = t[a:]
        result.append(r)
    return result


def _set_border(run):
    rPr = run._r.get_or_add_rPr()
    for old in rPr.findall(qn("w:bdr")):
        rPr.remove(old)
    bdr = OxmlElement("w:bdr")
    bdr.set(qn("w:val"), "single")
    bdr.set(qn("w:sz"), "8")
    bdr.set(qn("w:space"), "0")
    bdr.set(qn("w:color"), "auto")
    idx = _RPR_ORDER.index("bdr")
    for child in rPr:
        name = child.tag.split("}")[1]
        if name in _RPR_ORDER and _RPR_ORDER.index(name) > idx:
            child.addprevious(bdr)
            break
    else:
        rPr.append(bdr)


def mark(p, start=None, end=None):
    """'Obkroži' del odstavka: krepko + okvir okoli besedila."""
    t = text(p)
    if start is None:
        start, end = len(t) - len(t.lstrip()), len(t.rstrip())
    for r in isolate(p, start, end):
        r.bold = True
        _set_border(r)


def mark_text(p, needle):
    i = text(p).find(needle)
    if i < 0:
        return False
    mark(p, i, i + len(needle))
    return True


def fill_after(p, label, value):
    """Prvi niz podčrtajev za oznako `label` (regex) zamenja z vrednostjo. Prazna vrednost pusti
    podčrtaje. Vrne True, če je oznaka najdena."""
    t = text(p)
    m = re.search(label, t)
    if not m:
        return False
    if value in (None, ""):
        return True
    u = BLANK.search(t, m.end())
    if not u:
        return False
    val = str(value)
    sprednja_opomba = any(s == e == u.start() and r._r.find(qn("w:footnoteReference")) is not None
                          for r, s, e in _offsets(p))
    if sprednja_opomba or (u.start() > 0 and not t[u.start() - 1].isspace()):
        val = " " + val
    if u.end() < len(t) and not t[u.end()].isspace() and t[u.end()] not in ",.;":
        val = val + " "
    replace_span(p, u.start(), u.end(), val)
    return True


def fill_blank_par(p, value):
    """Zamenja niz podčrtajev v odstavku, ki vsebuje le prazno mesto (npr. druga vrstica naslova)."""
    u = BLANK.search(text(p))
    if not u:
        return False
    if value not in (None, ""):
        replace_span(p, u.start(), u.end(), str(value))
    return True


def clear_blank_par(p):
    t = text(p)
    u = BLANK.search(t)
    if u:
        replace_span(p, u.start(), u.end(), "")


def find(pars, pattern, lo=0, hi=None):
    hi = len(pars) if hi is None else hi
    rx = re.compile(pattern)
    for i in range(lo, hi):
        if rx.search(text(pars[i])):
            return i
    return -1


def fill_in(pars, lo, hi, label, value):
    i = find(pars, label, lo, hi)
    if i >= 0:
        fill_after(pars[i], label, value)
    return i


def delete_par(p):
    el = p._p
    el.getparent().remove(el)
