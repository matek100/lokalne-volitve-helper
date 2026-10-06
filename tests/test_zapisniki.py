"""Zapisniki o delu organa politične stranke."""
import re

import pytest

from conftest import odstavki, oznaceno
from lv import zapisniki


def vrstice(doc):
    return [t.strip() for t in odstavki(doc)]


def stevilo(doc, vzorec):
    return sum(1 for t in odstavki(doc) if re.search(vzorec, t))


class TestSkupnaGlava:
    @pytest.mark.parametrize("fn", ["lista", "vecinski", "zupan"])
    def test_seja(self, kandidat, nastavitve, fn):
        if fn == "zupan":
            doc = zapisniki.zapisnik_zupan(kandidat, nastavitve, "Ljubljana")
        else:
            doc = getattr(zapisniki, f"zapisnik_{fn}")([kandidat], nastavitve, "Ljubljana", "1")
        t = vrstice(doc)
        assert "Ime politične stranke: Primer stranka" in t
        assert "Naziv organa: Svet stranke" in t
        assert "je bila sklican(a) za dne 1. 10. 2026 ob 18.00 uri v kraju Ljubljana" in t
        assert "Organ je sklical: Ana Predstavnica" in t
        assert "Delo organa je/so vodil/i: Vodja Seje" in t
        assert "Zapisnik je vodil: Zapis Nikar" in t
        assert "Over Ena" in t and "Over Dva" in t
        assert any("navzočih 12 članov od skupnega števila 15" in x and "podlagi 20. člena" in x for x in t)
        assert any(x.endswith("občine Ljubljana") for x in t)
        assert any("končano ob 19.30 uri" in x for x in t)
        assert any("na podlagi" in x and "35. člena" in x for x in t if x.startswith("Predsedujoči je"))

    def test_prazna_seja_pusti_podcrtaje(self, kandidat, nastavitve):
        nastavitve["seja"] = {}
        t = vrstice(zapisniki.zapisnik_lista([kandidat], nastavitve, "Ljubljana", "1"))
        assert any(x.startswith("je bila sklican(a) za dne ____") for x in t)


class TestLista:
    @pytest.mark.parametrize("n", [1, 2, 5, 12])
    def test_stevilo_kandidatov(self, kandidat, nastavitve, n):
        doc = zapisniki.zapisnik_lista([kandidat] * n, nastavitve, "Ljubljana", "1")
        assert stevilo(doc, r"^\s*\d+\. Janez Novak, Slovenska cesta 10") == n  # predlog
        assert stevilo(doc, r"^\s*\d+\. Janez Novak\s+_+$") == n  # glasovi
        assert stevilo(doc, r"Ime in priimek: Janez Novak") == n  # določena lista
        assert stevilo(doc, r"[…]") == 0  # nadomestne vrstice so odstranjene
        assert f"Na listi kandidatov je (število) {n} kandidatov." in vrstice(doc)

    def test_ena_volilna_enota(self, kandidat, nastavitve):
        doc = zapisniki.zapisnik_lista([kandidat], nastavitve, "Ljubljana", "3")
        t = vrstice(doc)
        assert stevilo(doc, r"^\s*V volilni enoti št\.") == 1
        assert stevilo(doc, r"^\s*10\.\d") == 1
        assert stevilo(doc, r"Predstavnik liste kandidatov v volilni enoti") == 1
        assert stevilo(doc, r"a2\)") == 0
        assert "V volilni enoti št. 3" in t
        assert "a) za eno volilno enoto št. 3" in t
        assert "a) za eno volilno enoto" in oznaceno(doc)
        assert "Predstavnik liste kandidatov v volilni enoti št. 3 je:" in t

    def test_kandidat_in_predstavnik(self, kandidat, nastavitve):
        t = vrstice(zapisniki.zapisnik_lista([kandidat], nastavitve, "Ljubljana", "1"))
        assert f"EMŠO: {kandidat['emso']}" in t
        assert "občina: Ljubljana" in t and "naselje/kraj: Ljubljana" in t
        assert "ulica: Slovenska cesta hišna št. 10" in t
        assert "Ime liste kandidatov je Lista Primer" in t
        assert "Ime in priimek Ana Predstavnica roj. 1. 1. 1980" in t
        assert "tel. št. 040 000 000, e-naslov: info@primer.si" in t

    def test_oznake(self, kandidat, nastavitve):
        doc = zapisniki.zapisnik_lista([kandidat], nastavitve, "Ljubljana", "1")
        o = oznaceno(doc)
        assert "Zapisnik za listo kandidatov za OBČINSKI SVET" in o
        assert "JE" in o and "NI" not in o
        assert kandidat["raven"] in o
        nastavitve["seja"]["znak"] = "Ne"
        o = oznaceno(zapisniki.zapisnik_lista([kandidat], nastavitve, "Ljubljana", "1"))
        assert "NI" in o and "JE" not in o

    def test_cetrtna(self, kandidat, nastavitve):
        doc = zapisniki.zapisnik_lista([kandidat], nastavitve, "Ljubljana", "", tip="četrtna",
                                       naziv_skupnosti="ČS Center")
        assert "Zapisnik za listo kandidatov za SVET ČETRTNE SKUPNOSTI" in oznaceno(doc)
        assert any(x.endswith("kandidatura(e) Ljubljana, ČS Center") for x in vrstice(doc))


class TestVecinski:
    @pytest.mark.parametrize("n", [1, 2, 4])
    def test_stevilo_kandidatov(self, kandidat, kandidatka, nastavitve, n):
        kk = [kandidat, kandidatka] * n
        doc = zapisniki.zapisnik_vecinski(kk[:n], nastavitve, "Straža", "7")
        assert stevilo(doc, r"Ime in priimek: ") == n
        st = [int(m[1]) for t in odstavki(doc) if (m := re.match(r"^\s*(\d+)\.\s+Ime in priimek:", t))]
        assert st == list(range(1, n + 1))
        assert stevilo(doc, r"[…]") == 0
        assert stevilo(doc, r"^\s*ker je:") == 1

    def test_ve(self, kandidat, nastavitve):
        doc = zapisniki.zapisnik_vecinski([kandidat], nastavitve, "Straža", "7")
        t = vrstice(doc)
        assert "Za volilno enoto št. 7" in t and "1. za volilno enoto št. 7" in t
        assert "V volilni enoti št. 7" in t
        assert stevilo(doc, r"^\s*Za volilno enoto št\.") == 1

    def test_skupnost(self, kandidat, nastavitve):
        doc = zapisniki.zapisnik_vecinski([kandidat], nastavitve, "Kranjska Gora", "", tip="krajevna",
                                          naziv_skupnosti="Podkoren")
        assert "Zapisnik za kandidaturo za SVET KRAJEVNE/ČETRTNE/VAŠKE SKUPNOSTI" in oznaceno(doc)

    def test_predstavnik(self, kandidat, nastavitve):
        t = vrstice(zapisniki.zapisnik_vecinski([kandidat], nastavitve, "Straža", "7"))
        p = nastavitve["predstavnik"]
        assert f"Ime in priimek Ana Predstavnica,  datum rojstva 1. 1. 1980 EMŠO {p['emso']}" in t


class TestZupan:
    def test_vsebina(self, kandidatka, nastavitve):
        doc = zapisniki.zapisnik_zupan(kandidatka, nastavitve, "Maribor")
        t = vrstice(doc)
        assert "1. Mojca Kovač, Trg svobode 1, 2000 Maribor" in t
        assert stevilo(doc, r"^\s*\d+\. _{5,}$") == 2  # dnevni red 2. in 3. ostaneta za ročni vpis
        assert stevilo(doc, r"^\s*\d+\. Mojca Kovač\s+_+$") == 1
        assert "Ime in priimek: Mojca Kovač" in t
        assert f"Stopnja izobrazbe: {kandidatka['raven']}" in t
        assert "Naziv izobrazb: gimnazijska maturantka" in t
        assert "Delo, ki ga opravlja: študentka" in t
        assert any("Občine Maribor" in x for x in t)
        assert "Ime in priimek Ana Predstavnica roj. 1. 1. 1980" in t
