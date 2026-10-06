"""Čiščenje in razčlenjevanje podatkov iz obrazca za evidentiranje kandidatov."""
import datetime as dt
import difflib
import re
import unicodedata

# Ravni izobrazbe točno tako, kot so zapisane na obrazcih (uporablja se za označevanje).
RAVNI_IZOBRAZBE = [
    "1. raven: nedokončana osnovnošolska izobrazba",
    "2. raven: osnovnošolska izobrazba",
    "3. raven: nižja poklicna izobrazba",
    "4. raven: srednja poklicna izobrazba",
    "5. raven: srednja strokovna izobrazba",
    "5. raven: srednja izobrazba",
    "6. raven: višja strokovna izobrazba",
    "6. raven: višješolska izobrazba",
    "7. raven: visokošolska univerzitetna/ visokošolska strokovna izobrazba ali visoka strokovna izobrazba,",
    "8. raven: diploma o univerzitetnem izobraževanju",
    "8. raven: diploma druge stopnje (magisterij, pridobljen po magistrskem študijskem programu ali enovitem magistrskem študijskem programu)",
    "8. raven: diploma o specializaciji po visokošolski strokovni izobrazbi, diploma o visokošolskem izobraževanju,",
    "9. raven: magisterij znanosti/umetnosti, diploma o specializaciji po visoki univerzitetni izobrazbi, diploma o specializaciji po visoki strokovni izobrazbi,",
    "10. raven oz. stopnja izobrazbe: diploma tretje stopnje oz. doktorat  znanosti.",
]
R = RAVNI_IZOBRAZBE

TIPI_PREDLAGATELJA = ["Politična stranka", "Skupna lista", "Skupina volivcev"]
SISTEMI = ["proporcionalni", "večinski"]
TIPI_SKUPNOSTI = ["četrtna", "krajevna", "vaška"]

# pogosti vzdevki občin
OBCINE_VZDEVKI = {
    "mol": "Ljubljana", "mom": "Maribor", "mok": "Koper", "mok koper": "Koper",
    "mong": "Nova Gorica", "monm": "Novo mesto", "moc": "Celje", "mokr": "Kranj",
    "mov": "Velenje", "mops": "Murska Sobota", "mosg": "Slovenj Gradec", "moptuj": "Ptuj",
}


def cell(v):
    """Vrednost celice kot očiščen niz ('' za prazno)."""
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:  # NaN
            return ""
        if v.is_integer():
            v = int(v)
    if isinstance(v, (dt.datetime, dt.date)):
        return datum_str(v)
    return re.sub(r"\s+", " ", str(v)).strip()


def da(v):
    return cell(v).lower() in ("da", "yes", "true", "1", "x")


def brez_sumnikov(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


# --- EMŠO --------------------------------------------------------------------------------------
DAN_VOLITEV = dt.date(2026, 11, 15)
EMSO_UTEZI = [7, 6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2]


def emso_norm(v):
    s = re.sub(r"\D", "", cell(v))
    if len(s) == 12:
        s = "0" + s  # Excel izgubi vodilno ničlo (dan rojstva 01-09)
    return s


def emso_kontrolna(e):
    """Kontrolna števka za prvih 12 števk ali None, če je ostanek 10 (takšen EMŠO ne obstaja)."""
    m = 11 - sum(int(a) * b for a, b in zip(e[:12], EMSO_UTEZI)) % 11
    if m == 10:
        return None
    return 0 if m == 11 else m


def preveri_emso(vnos, datum_rojstva=None, spol=None, dan_volitev=DAN_VOLITEV):
    """Preveri EMŠO in vrne seznam napak (prazen seznam = veljaven).

    Preveri dolžino, datum rojstva v EMŠO, kontrolno števko, polnoletnost na dan volitev ter
    (neobvezno) ujemanje z vpisanim datumom rojstva in spolom."""
    surovo = cell(vnos)
    if not surovo:
        return ["EMŠO manjka"]
    if re.search(r"[^\d\s.\-/]", surovo):
        return [f"EMŠO '{surovo}' vsebuje nedovoljene znake"]
    e = emso_norm(surovo)
    if len(e) != 13:
        return [f"EMŠO ima {len(e)} števk namesto 13"]
    napake = []
    rojen = emso_datum(e)
    if not rojen:
        napake.append(f"EMŠO vsebuje neveljaven datum rojstva ({e[0:2]}.{e[2:4]}.{e[4:7]})")
    elif rojen > dt.date.today():
        napake.append("datum rojstva v EMŠO je v prihodnosti")
    k = emso_kontrolna(e)
    if k is None:
        napake.append("EMŠO s temi prvimi 12 števkami ne more obstajati (kontrolni ostanek 10)")
    elif k != int(e[12]):
        napake.append(f"napačna kontrolna števka EMŠO ({e[12]}, pričakovana {k}) - verjetno tipkarska napaka")
    if rojen and dan_volitev and rojen <= dt.date.today():
        starost = dan_volitev.year - rojen.year - ((dan_volitev.month, dan_volitev.day) < (rojen.month, rojen.day))
        if starost < 18:
            napake.append(f"na dan volitev ({datum_str(dan_volitev)}) kandidat ne bo polnoleten ({starost} let)")
    if rojen and datum_rojstva:
        d = datum_rojstva if isinstance(datum_rojstva, dt.date) else datum_iz_besedila(cell(datum_rojstva))
        if d and d != rojen:
            napake.append(f"datum rojstva ({datum_str(d)}) se ne ujema z EMŠO ({datum_str(rojen)})")
    if spol and cell(spol) in ("M", "Ž") and cell(spol) != emso_spol(e):
        napake.append(f"spol ({cell(spol)}) se ne ujema z EMŠO ({emso_spol(e)})")
    return napake


def sestavi_emso(dan, mesec, leto, spol="M", regija=50, zaporedna=0):
    """Sestavi veljaven (izmišljen) EMŠO - za primere in teste."""
    zap = zaporedna % 500 + (500 if spol == "Ž" else 0)
    while True:
        e = f"{dan:02d}{mesec:02d}{leto % 1000:03d}{regija:02d}{zap:03d}"
        k = emso_kontrolna(e)
        if k is not None:
            return e + str(k)
        zap += 1


def emso_veljaven(e):
    return not preveri_emso(e, dan_volitev=None)


def emso_datum(e):
    if not re.fullmatch(r"\d{13}", e):
        return None
    d, m, y = int(e[0:2]), int(e[2:4]), int(e[4:7])
    y += 1000 if y >= 800 else 2000
    try:
        return dt.date(y, m, d)
    except ValueError:
        return None


def emso_spol(e):
    if not re.fullmatch(r"\d{13}", e):
        return ""
    return "M" if int(e[9:12]) < 500 else "Ž"


def datum_str(d):
    return f"{d.day}. {d.month}. {d.year}" if d else ""


def datum_iz_besedila(s):
    m = re.search(r"(\d{1,2})\s*\.\s*(\d{1,2})\s*\.\s*(\d{4})", s or "")
    if not m:
        return None
    try:
        return dt.date(int(m[3]), int(m[2]), int(m[1]))
    except ValueError:
        return None


# --- ime in priimek ----------------------------------------------------------------------------
def razdeli_ime(polno):
    parts = cell(polno).split(" ")
    if len(parts) < 2:
        return cell(polno), ""
    return parts[0], " ".join(parts[1:])


# --- naslov ------------------------------------------------------------------------------------
def razcleni_naslov(s):
    """'Slovenska cesta 10a, 1000 Ljubljana' -> dict(ulica, hs, dodatek, posta, kraj, vrstica1, vrstica2).

    Dopušča tudi zapise brez vejic ('Ob železnici 14 Ljubljana'), s krajem spredaj
    ('Domžale, Dragomelj 153c') in s pošto brez kraja ('..., 1411, Izlake')."""
    s = cell(s)
    out = dict(ulica="", hs=None, dodatek=None, posta=None, kraj="", vrstica1=s, vrstica2="")
    deli = [d.strip() for d in s.split(",") if d.strip()]
    ulica_del, ostalo = None, []
    for d in deli:
        m = re.fullmatch(r"(\d{4})\s*(\D*)", d)
        if m and not out["posta"]:
            out["posta"] = m[1]
            if m[2].strip():
                ostalo.append(m[2].strip())
            continue
        if ulica_del is None:
            h = re.fullmatch(r"(.*?\D)\s*(\d+)\s*([a-zA-Z])?(?:\s+(\d{4}))?(?:\s+([^\d]+))?", d)
            if h and h[1].strip(" ."):
                ulica_del = d
                out["ulica"], out["hs"] = h[1].strip(" ,"), int(h[2])
                out["dodatek"] = h[3].upper() if h[3] else None
                if h[4] and not out["posta"]:
                    out["posta"] = h[4]
                if h[5]:
                    ostalo.append(h[5].strip())
                    ulica_del = d[:h.start(4) if h[4] else h.start(5)].strip()
                continue
        ostalo.append(d)
    out["kraj"] = ", ".join(ostalo)
    out["vrstica1"] = ulica_del or s
    out["vrstica2"] = " ".join(x for x in (out["posta"], out["kraj"]) if x) if ulica_del else ""
    return out


# --- občina ------------------------------------------------------------------------------------
def _kljuc_obcine(s):
    s = brez_sumnikov(cell(s).lower())
    s = re.sub(r"^(mestna obcina|obcina|mo)\s+", "", s)
    return re.sub(r"[^a-z ]", "", s).strip()


def normaliziraj_obcino(vnos, uradne):
    """Vrne (uradni naziv ali '', opomba)."""
    v = cell(vnos)
    if not v:
        return "", ""
    low = v.lower().strip()
    if low in OBCINE_VZDEVKI:
        return OBCINE_VZDEVKI[low], ""
    kljuci = {_kljuc_obcine(o): o for o in uradne}
    k = _kljuc_obcine(v)
    if k in kljuci:
        return kljuci[k], ""
    if re.search(r"\d", v):
        return "", f"vnos občine '{v}' je videti kot naslov"
    m = difflib.get_close_matches(k, list(kljuci), n=1, cutoff=0.8)
    if m:
        return kljuci[m[0]], f"občina '{v}' ugotovljena približno"
    return "", f"občine '{v}' ni bilo mogoče prepoznati"


# --- izobrazba ---------------------------------------------------------------------------------
def oceni_raven(besedilo):
    """Hevristično preslika prosto besedilo v raven izobrazbe. Vrne '' če ni zanesljivo."""
    t = " " + brez_sumnikov(cell(besedilo).lower()) + " "
    if not t.strip():
        return ""
    if re.search(r"\bdr\.|doktor znanosti|doktorat(?! ?ski)|phd", t) and "studij" not in t:
        return R[13]
    if re.search(r"mag\. ?znan|magister znanosti|magisterij znanosti|viii/1", t):
        return R[12]
    if re.search(r"\bmag\.|magist|master|viii", t):
        return R[10]
    if re.search(r"\(un\)|\(vs\)|bolonj|\b1\. stopnja|prva stopnja|vi/2|\bvs\b", t):
        return R[8]
    if re.search(r"univ\.|univerzitetn|\bvii\b|\b7\. ?stopnj|\b7\.? ?raven", t):
        return R[9]
    if re.search(r"dipl\.|diplomiran|visoka strokovna|visokosolsk", t):
        return R[8]
    if re.search(r"vis(ja|jo|je)|\bvi\b|vi/1|\b6\. ?stopnj|\b6\.? ?raven|\bvi\.", t):
        return R[6]
    if re.search(r"gimnazij|matur", t):
        return R[5]
    if re.search(r"tehnik|srednja strokovna|\bv\.|\bv\b|\b5\. ?stopnj|\b5\.? ?raven|srednj", t):
        return R[4]
    if re.search(r"nizja poklicna|\b3\. ?stopnj", t):
        return R[2]
    if re.search(r"poklicn|\biv\b|\b4\. ?stopnj", t):
        return R[3]
    if re.search(r"osnovn", t):
        return R[1]
    return ""


def strokovni_naslov(besedilo):
    """Odstrani oznako stopnje ('V.', '7. stopnja,' ...) in vrne preostanek kot strokovni naslov."""
    s = cell(besedilo)
    s = re.sub(r"^\s*(?:[IVX]+(?:/\d)?\b\.?|\d+\.?)\s*(?:stopnja|raven)?\s*[,\-–:]?\s*", "", s, flags=re.I)
    return s.strip(" ,-–")
