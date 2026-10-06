"""Korak 2: iz priprava.xlsx ustvari izpolnjene obrazce (.docx) v mapi output/<Občina>/."""
import re
import shutil
from collections import defaultdict
from pathlib import Path

import pandas as pd

from . import obrazci, zapisniki
from . import podatki as P
from .preverjanje import napake_emso
from .priprava import K, OBCINE_PREPIS, vec_volilnih_enot


def _ime_datoteke(s):
    return re.sub(r'[<>:"/\\|?*]+', "-", s).strip(" .")


def _red(v):
    try:
        return float(P.cell(v))
    except ValueError:
        return float("inf")


def _preveri_red(kk, stolpec, kje):
    """Opozorila za podvojen ali neštevilski vrstni red v eni listi."""
    out, videni = [], {}
    for r in kk:
        v = P.cell(r.get(stolpec))
        if _red(v) == float("inf"):
            out.append(f"{kje}: {_oseba(r)} nima številskega vrstnega reda ('{v}') - uvrščen na konec")
        elif v in videni:
            out.append(f"{kje}: vrstni red {v} imata {videni[v]} in {_oseba(r)}")
        videni.setdefault(v, _oseba(r))
    return out


def _deli_naslova(naslov1, naslov2):
    """Ulica, hišna številka in naselje iz dveh vrstic naslova (za zapisnike)."""
    n = P.razcleni_naslov(", ".join(x for x in (naslov1, naslov2) if x))
    hs = f"{n['hs']}{(n['dodatek'] or '').lower()}" if n["hs"] is not None else ""
    return dict(ulica=n["ulica"], hs=hs, naselje=re.sub(r"\s+-\s+.*$", "", n["kraj"]))


def _ve_st(ve, vec_ve=False):
    """Številka volilne enote za zapisnike ('Volilna enota 7' -> '7'; ena VE -> '1')."""
    m = re.search(r"\d+", ve or "")
    return m[0] if m else ("" if vec_ve else "1")


def _nastavitve(nast, obcina_row):
    """Združi privzete nastavitve in morebitne prepise za občino v obliko za obrazce."""
    v = dict(nast)
    for k in OBCINE_PREPIS:
        if P.cell((obcina_row or {}).get(k)):
            v[k] = P.cell(obcina_row[k])
    return dict(
        predlagatelj=v.get("Predlagatelj", ""),
        tip_predlagatelja=v.get("Tip predlagatelja", ""),
        organ=v.get("Organ", ""),
        ime_liste=v.get("Ime liste", ""),
        podpisi=v.get("Število podpisov", ""),
        kraj=v.get("Kraj podpisa", ""),
        datum=v.get("Datum podpisa", ""),
        predstavnik=dict(
            emso=P.emso_norm(v.get("Predstavnik EMŠO", "")),
            ime=v.get("Predstavnik ime", ""), priimek=v.get("Predstavnik priimek", ""),
            datum_rojstva=v.get("Predstavnik datum rojstva", ""),
            naslov1=v.get("Predstavnik naslov - ulica", ""), naslov2=v.get("Predstavnik naslov - pošta", ""),
            eposta=v.get("Predstavnik e-pošta", ""), telefon=v.get("Predstavnik telefon", ""),
            obcina=v.get("Predstavnik občina", ""),
            **_deli_naslova(v.get("Predstavnik naslov - ulica", ""), v.get("Predstavnik naslov - pošta", "")),
        ),
        seja=dict(
            datum=v.get("Seja - datum", ""), ura=v.get("Seja - ura", ""), kraj=v.get("Seja - kraj", ""),
            sklical=v.get("Seja - sklical", ""), vodil=v.get("Seja - vodil", ""),
            zapisnikar=v.get("Seja - zapisnikar", ""), overovatelj1=v.get("Seja - overovatelj 1", ""),
            overovatelj2=v.get("Seja - overovatelj 2", ""), navzocih=v.get("Seja - navzočih", ""),
            vabljenih=v.get("Seja - vabljenih", ""), ura_konca=v.get("Seja - ura konca", ""),
            clen_sklepcnosti=v.get("Statut - člen sklepčnosti", ""),
            clen_dolocitve=v.get("Statut - člen določitve", ""),
            znak=v.get("Znak stranke v imenu liste", ""),
        ),
    )


def _zapisnik(s):
    """Zapisnik o delu organa stranke se pripravi za stranke in skupne liste, ne za skupine volivcev."""
    return P.cell(s.get("tip_predlagatelja")).lower() != "skupina volivcev"


def _kandidat(r):
    g = lambda k: P.cell(r.get(K[k]))  # noqa: E731
    deli = dict(ulica=g("ulica"), hs=g("hs"), naselje=g("naselje"))
    if not any(deli.values()):  # starejša priprava brez teh stolpcev
        deli = _deli_naslova(g("naslov1"), g("naslov2"))
    return dict(ime=g("ime"), priimek=g("priimek"), emso=g("emso"), datum_rojstva=g("datum"),
                spol=g("spol"), naslov1=g("naslov1"), naslov2=g("naslov2"), raven=g("raven"),
                strokovni_naslov=g("naziv"), delo=g("delo"), obcina_naslov=g("obcina_naslov"), **deli)


def _oseba(r):
    return _ime_datoteke(f"{P.cell(r.get(K['priimek']))} {P.cell(r.get(K['ime']))}")


def generiraj(priprava, izhod, prepisi=False):
    izhod = Path(izhod)
    if izhod.exists() and any(izhod.iterdir()):
        if not prepisi:
            raise SystemExit(f"Mapa '{izhod}' ni prazna. Ročni popravki bi se izgubili - uporabi drugo "
                             f"mapo (-o) ali dodaj --prepisi.")
        shutil.rmtree(izhod)
    xl = pd.read_excel(priprava, sheet_name=None, dtype=str, keep_default_na=False)
    nast = {P.cell(r["Nastavitev"]): P.cell(r["Vrednost"]) for r in xl["Nastavitve"].to_dict("records")}
    obcine = {P.cell(r["Občina"]): r for r in xl["Občine"].to_dict("records")}
    rows = [r for r in xl["Kandidati"].to_dict("records") if P.da(r.get(K["vkljuci"]))]

    ustvarjeno, opozorila = [], []
    for kdo, e, napake in napake_emso(xl):
        opozorila += [f"{kdo}: EMŠO {e}: {n}" for n in napake]

    def shrani(doc, obcina, ime, podmapa=None):
        mapa = izhod / _ime_datoteke(obcina or "Brez občine")
        if podmapa:
            mapa = mapa / _ime_datoteke(podmapa)
        mapa.mkdir(parents=True, exist_ok=True)
        pot = mapa / (_ime_datoteke(ime) + ".docx")
        doc.save(pot)
        ustvarjeno.append(pot)

    # --- občinski svet ---
    svet = defaultdict(list)
    for r in rows:
        if P.da(r.get(K["svet"])):
            if not P.cell(r.get(K["obcina"])):
                opozorila.append(f"{_oseba(r)}: kandidira v občinski svet, a občina ni določena - preskočeno")
                continue
            svet[P.cell(r[K["obcina"]])].append(r)
    for obcina, kandidati in sorted(svet.items()):
        orow = obcine.get(obcina, {})
        s = _nastavitve(nast, orow)
        sistem = P.cell(orow.get("Volilni sistem")).lower()
        if sistem not in P.SISTEMI:
            opozorila.append(f"{obcina}: volilni sistem ni izbran (list 'Občine') - lista/kandidatura "
                             f"za občinski svet ni ustvarjena, soglasja so")
        vec_ve = vec_volilnih_enot(orow)
        po_ve = defaultdict(list)
        for r in kandidati:
            po_ve[P.cell(r.get(K["ve"])) if vec_ve else ""].append(r)
        for ve, kk in sorted(po_ve.items()):
            kk.sort(key=lambda r: _red(r.get(K["red"])))
            opozorila.extend(_preveri_red(kk, K["red"], f"{obcina}{' ' + ve if ve else ''}"))
            if vec_ve and not ve:
                opozorila.append(f"{obcina}: kandidati brez volilne enote: " + ", ".join(map(_oseba, kk)))
            pripona = f" - {ve}" if ve else ""
            c = [_kandidat(r) for r in kk]
            if sistem == "proporcionalni":
                shrani(obrazci.lista_proporcionalni(c, s, volilna_enota=ve), obcina,
                       f"Lista kandidatov - občinski svet{pripona}")
                if _zapisnik(s):
                    shrani(zapisniki.zapisnik_lista(c, s, obcina, _ve_st(ve, vec_ve)), obcina,
                           f"Zapisnik - lista kandidatov - občinski svet{pripona}")
            elif sistem == "večinski":
                shrani(obrazci.kandidatura_vecinski(c, s, volilna_enota=ve), obcina,
                       f"Kandidatura - občinski svet{pripona}")
                if _zapisnik(s):
                    shrani(zapisniki.zapisnik_vecinski(c, s, obcina, _ve_st(ve, vec_ve)), obcina,
                           f"Zapisnik - kandidatura - občinski svet{pripona}")
        lista_naziv = (s["ime_liste"] or s["predlagatelj"]) if sistem == "proporcionalni" else s["predlagatelj"]
        for r in kandidati:
            shrani(obrazci.soglasje_svet(_kandidat(r), obcina, s, lista_naziv), obcina,
                   f"Soglasje - občinski svet - {_oseba(r)}")

    # --- župan ---
    for r in rows:
        if not P.da(r.get(K["zupan"])):
            continue
        obcina = P.cell(r.get(K["obcina"]))
        if not obcina:
            opozorila.append(f"{_oseba(r)}: kandidira za župana, a občina ni določena - preskočeno")
            continue
        s = _nastavitve(nast, obcine.get(obcina))
        shrani(obrazci.kandidatura_zupan(_kandidat(r), s), obcina, f"Kandidatura za župana - {_oseba(r)}")
        shrani(obrazci.soglasje_zupan(_kandidat(r), obcina, s), obcina, f"Soglasje - župan - {_oseba(r)}")
        if _zapisnik(s):
            shrani(zapisniki.zapisnik_zupan(_kandidat(r), s, obcina), obcina, f"Zapisnik - župan - {_oseba(r)}")

    # --- svet četrtne / krajevne / vaške skupnosti ---
    sk = defaultdict(list)
    for r in rows:
        if not P.da(r.get(K["skupnost"])):
            continue
        obcina = P.cell(r.get(K["obcina"])) or P.cell(r.get(K["obcina_naslov"]))
        tip, naziv = P.cell(r.get(K["tip_sk"])).lower(), P.cell(r.get(K["naziv_sk"]))
        if tip not in P.TIPI_SKUPNOSTI or not naziv:
            opozorila.append(f"{_oseba(r)}: svet ČS/KS - manjka tip ali naziv skupnosti - preskočeno")
            continue
        sk[(obcina, tip, naziv, P.cell(r.get(K["ve_sk"])))].append(r)
    for (obcina, tip, naziv, ve), kk in sorted(sk.items()):
        kk.sort(key=lambda r: _red(r.get(K["red_sk"])))
        opozorila.extend(_preveri_red(kk, K["red_sk"], naziv))
        orow = obcine.get(obcina, {})
        s = _nastavitve(nast, orow)
        c = [_kandidat(r) for r in kk]
        pripona = f" - {ve}" if ve else ""
        ve_st = _ve_st(ve, vec_ve=True)  # VE skupnosti ni v registru - brez vpisa ostane prazno
        if tip == "četrtna" and P.cell(orow.get("Volilni sistem ČS")).lower() == "proporcionalni":
            shrani(obrazci.lista_proporcionalni(c, s, tip=tip, naziv_skupnosti=naziv, volilna_enota=ve),
                   obcina, f"Lista kandidatov - {naziv}{pripona}", podmapa=naziv)
            if _zapisnik(s):
                shrani(zapisniki.zapisnik_lista(c, s, obcina, ve_st, tip=tip, naziv_skupnosti=naziv),
                       obcina, f"Zapisnik - lista kandidatov - {naziv}{pripona}", podmapa=naziv)
        else:
            shrani(obrazci.kandidatura_vecinski(c, s, tip=tip, naziv_skupnosti=naziv, volilna_enota=ve),
                   obcina, f"Kandidatura - {naziv}{pripona}", podmapa=naziv)
            if _zapisnik(s):
                shrani(zapisniki.zapisnik_vecinski(c, s, obcina, ve_st, tip=tip, naziv_skupnosti=naziv),
                       obcina, f"Zapisnik - kandidatura - {naziv}{pripona}", podmapa=naziv)

    # --- povzetek ---
    izhod.mkdir(parents=True, exist_ok=True)
    povzetek = ["USTVARJENI OBRAZCI", ""] + [str(p.relative_to(izhod)) for p in ustvarjeno]
    if opozorila:
        povzetek += ["", "OPOZORILA", ""] + opozorila
    (izhod / "povzetek.txt").write_text("\n".join(povzetek) + "\n", encoding="utf-8")
    print(f"Ustvarjenih {len(ustvarjeno)} obrazcev v '{izhod}'.")
    for o in opozorila:
        print("  ! " + o)
