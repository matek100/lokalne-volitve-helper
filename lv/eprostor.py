"""Poizvedbe na javni WFS servis GURS (eProstor): register naslovov in register prostorskih enot.

Rezultati se shranjujejo v lokalni predpomnilnik (.cache/eprostor.json), da ponovni zagon ne
obremenjuje servisa.
"""
import json
import re
from pathlib import Path

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

RN_WFS = "https://ipi.eprostor.gov.si/wfs-si-gurs-rn/wfs"
RPE_WFS = "https://ipi.eprostor.gov.si/wfs-si-gurs-rpe/wfs"

CACHE_PATH = Path(__file__).resolve().parent.parent / ".cache" / "eprostor.json"

NASLOV_POLJA = ",".join([
    "OBCINA_NAZIV", "NASELJE_NAZIV", "ULICA_NAZIV", "POSTNI_OKOLIS_SIFRA", "POSTNI_OKOLIS_NAZIV",
    "HS_STEVILKA", "HS_DODATEK", "EID_OBCINA", "EID_LOKALNA_VOLILNA_ENOTA",
    "EID_CETRTNA_SKUPNOST", "EID_KRAJEVNA_SKUPNOST", "EID_VASKA_SKUPNOST",
])

# sloje skupnosti: tip -> (sloj, polje EID)
SKUPNOSTI = {
    "četrtna": ("SI.GURS.RPE:CETRTNE_SKUPNOSTI", "EID_CETRTNA_SKUPNOST"),
    "krajevna": ("SI.GURS.RPE:KRAJEVNE_SKUPNOSTI", "EID_KRAJEVNA_SKUPNOST"),
    "vaška": ("SI.GURS.RPE:VASKE_SKUPNOSTI", "EID_VASKA_SKUPNOST"),
}


def _cql_str(s):
    return "'" + s.replace("'", "''") + "'"


class EProstor:
    def __init__(self, cache_path=CACHE_PATH):
        self.cache_path = Path(cache_path)
        self.session = requests.Session()
        # servis občasno prekine povezavo - ponovi poskus s postopnim čakanjem
        ponovi = Retry(total=5, connect=5, read=5, backoff_factor=0.5,
                       status_forcelist=(429, 500, 502, 503, 504), allowed_methods=("GET",))
        self.session.mount("https://", HTTPAdapter(max_retries=ponovi))
        try:
            self.cache = json.loads(self.cache_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            self.cache = {}
        self._obcine = None
        self._lve = None

    def save(self):
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        self.cache_path.write_text(json.dumps(self.cache, ensure_ascii=False), encoding="utf-8")

    def _get(self, base, layer, cql=None, props=None, count=None):
        key = f"{layer}|{cql}|{props}|{count}"
        if key in self.cache:
            return self.cache[key]
        params = dict(service="WFS", version="2.0.0", request="GetFeature", typeNames=layer,
                      outputFormat="application/json")
        if cql:
            params["CQL_FILTER"] = cql
        if props:
            params["propertyName"] = props
        if count:
            params["count"] = count
        r = self.session.get(base, params=params, timeout=120)
        r.raise_for_status()
        rows = [f["properties"] for f in r.json().get("features", [])]
        self.cache[key] = rows
        return rows

    # --- šifranti ---------------------------------------------------------------------------
    @property
    def obcine(self):
        """EID_OBCINA -> uradni naziv občine"""
        if self._obcine is None:
            rows = self._get(RPE_WFS, "SI.GURS.RPE:OBCINE", props="EID_OBCINA,SIFRA,NAZIV")
            self._obcine = {r["EID_OBCINA"]: r["NAZIV"] for r in rows}
        return self._obcine

    @property
    def volilne_enote(self):
        """EID_LOKALNA_VOLILNA_ENOTA -> dict(sifra, naziv, eid_obcina)"""
        if self._lve is None:
            rows = self._get(RPE_WFS, "SI.GURS.RPE:LOKALNE_VOLILNE_ENOTE",
                             props="EID_LOKALNA_VOLILNA_ENOTA,SIFRA,NAZIV,EID_OBCINA")
            self._lve = {r["EID_LOKALNA_VOLILNA_ENOTA"]: dict(
                sifra=r["SIFRA"], naziv=re.sub(r"\s+", " ", r["NAZIV"] or "").strip(),
                eid_obcina=r["EID_OBCINA"]) for r in rows}
        return self._lve

    def st_volilnih_enot(self, obcina_naziv):
        eids = [e for e, n in self.obcine.items() if n == obcina_naziv]
        return sum(1 for v in self.volilne_enote.values() if v["eid_obcina"] in eids)

    def skupnost(self, tip, eid):
        if not eid:
            return None
        layer, polje = SKUPNOSTI[tip]
        rows = self._get(RPE_WFS, layer, cql=f"{polje}={_cql_str(eid)}", props="NAZIV,SIFRA")
        return rows[0]["NAZIV"] if rows else None

    # --- naslovi ----------------------------------------------------------------------------
    def _naslovi(self, cql):
        rows = self._get(RN_WFS, "SI.GURS.RN:REGISTER_NASLOVOV", cql=cql, props=NASLOV_POLJA, count=200)
        # register vsebuje vrstico za vsako stanovanje; združimo po hišni številki
        uniq = {}
        for r in rows:
            k = (r["EID_OBCINA"], r["NASELJE_NAZIV"], r["ULICA_NAZIV"], r["HS_STEVILKA"], r["HS_DODATEK"])
            uniq.setdefault(k, r)
        return list(uniq.values())

    def najdi_naslov(self, ulica, hs, dodatek=None, posta=None, kraj=None, obcina=None):
        """Vrne (zapis, opomba). Zapis je slovar iz registra naslovov ali None."""
        hs_f = f"HS_STEVILKA={int(hs)}"
        dod_f = f"HS_DODATEK ILIKE {_cql_str(dodatek)}" if dodatek else "HS_DODATEK IS NULL"
        post_f = f" AND POSTNI_OKOLIS_SIFRA={int(posta)}" if posta else ""
        u = _cql_str(ulica)
        # ohlapen vzorec: okrajšave ("ul.", "c.") in manjkajoče šumnike (c/č, s/š, z/ž)
        tokens = [t.rstrip(".") for t in re.split(r"[\s.]+", ulica) if t.strip(".")]
        loose = "%".join(re.sub(r"[cčćsšzž]", "_", t, flags=re.I) for t in tokens)
        loose = _cql_str(loose + "%")  # 'Vegova' -> 'Vegova ulica'
        poskusi = [
            (f"ULICA_NAZIV ILIKE {u}", None),
            (f"ULICA_NAZIV IS NULL AND NASELJE_NAZIV ILIKE {u}", None),
            (f"ULICA_NAZIV ILIKE {loose}", "približno ujemanje imena ulice"),
            (f"ULICA_NAZIV IS NULL AND NASELJE_NAZIV ILIKE {loose}", "približno ujemanje imena naselja"),
        ]
        # najprej natančno (z dodatkom ali brez njega), nato še ne glede na dodatek k hišni številki
        variante = [([dod_f], None), ([], "hišna številka se razlikuje v dodatku (npr. 12 / 12a)")]
        for dodatni, dod_opomba in variante:
            for pogoj, opomba in poskusi:
                for pf in ([post_f, ""] if post_f else [""]):
                    cql = " AND ".join([pogoj, hs_f] + dodatni) + pf
                    rows = self._naslovi(cql)
                    if not rows:
                        continue
                    op = [o for o in (opomba, dod_opomba) if o]
                    if post_f and not pf:
                        op.append("poštna številka se ne ujema")
                    # več zadetkov: razloči po kraju in po občini, v kateri oseba kandidira
                    k = (kraj or "").lower()
                    for pogoj_r in (
                        lambda r: k and k in ((r["POSTNI_OKOLIS_NAZIV"] or "").lower(), (r["NASELJE_NAZIV"] or "").lower()),
                        lambda r: obcina and r["OBCINA_NAZIV"] == obcina,
                    ):
                        ujem = [r for r in rows if pogoj_r(r)]
                        if ujem and len(rows) > 1:
                            rows = ujem
                    if len(rows) > 1:
                        op.append(f"več možnih naslovov ({len(rows)}), izbran prvi")
                    return rows[0], "; ".join(op)
        return None, "naslov ni najden v registru naslovov"
