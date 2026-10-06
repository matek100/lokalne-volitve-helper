"""Preverjanje EMŠO v pripravi (priprava.xlsx) ali v surovem izvozu obrazca."""
from collections import Counter

import pandas as pd

from . import podatki as P
from .priprava import K, VHOD, preberi_tabelo


def napake_emso(xl):
    """xl: dict listov (pd.read_excel(sheet_name=None)) ali en DataFrame izvoza.
    Vrne seznam (kdo, emso, [napake])."""
    out = []
    if isinstance(xl, pd.DataFrame):  # surov izvoz
        df, c_emso, c_datum, c_spol = xl, VHOD["emso"], VHOD["rojstvo"], None
        ime = lambda r: P.cell(r.get(VHOD["ime"]))  # noqa: E731
    else:
        df, c_emso, c_datum, c_spol = xl["Kandidati"], K["emso"], K["datum"], K["spol"]
        df = df[df[K["vkljuci"]].map(P.da)] if K["vkljuci"] in df else df
        ime = lambda r: f"{P.cell(r.get(K['ime']))} {P.cell(r.get(K['priimek']))}"  # noqa: E731
    rows = df.to_dict("records")
    for r in rows:
        n = P.preveri_emso(r.get(c_emso), r.get(c_datum), r.get(c_spol) if c_spol else None)
        if n:
            out.append((ime(r), P.cell(r.get(c_emso)), n))
    # podvojeni EMŠO
    stevec = Counter(P.emso_norm(r.get(c_emso)) for r in rows if P.cell(r.get(c_emso)))
    for e, n in stevec.items():
        if n > 1:
            imena = ", ".join(ime(r) for r in rows if P.emso_norm(r.get(c_emso)) == e)
            out.append((imena, e, [f"isti EMŠO se pojavi {n}-krat"]))
    if isinstance(xl, dict):
        # predstavnik kandidature (privzeti in prepisi po občinah); prazen ni napaka
        nast = {P.cell(r["Nastavitev"]): P.cell(r["Vrednost"]) for r in xl["Nastavitve"].to_dict("records")}
        predstavniki = [("privzeto", nast)] + [(P.cell(r.get("Občina")), r) for r in xl["Občine"].to_dict("records")]
        for kje, v in predstavniki:
            e = P.cell(v.get("Predstavnik EMŠO"))
            if e:
                n = P.preveri_emso(e, v.get("Predstavnik datum rojstva"), dan_volitev=None)
                if n:
                    out.append((f"predstavnik kandidature ({kje})", e, n))
    return out


def preveri_datoteko(pot):
    try:
        xl = pd.read_excel(pot, sheet_name=None, dtype=str, keep_default_na=False)
        if "Kandidati" not in xl:
            xl = preberi_tabelo(pot)
    except ValueError:  # csv
        xl = preberi_tabelo(pot)
    return napake_emso(xl)
