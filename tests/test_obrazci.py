"""Izpolnjevanje dejanskih predlog iz templates/."""
import re

import pytest

from conftest import najdi, odstavki, oznaceno
from lv import obrazci


def emso_vrstice(doc):
    return [t for t in odstavki(doc) if re.match(r"^\s*\d+\s*\.\s*EMŠO", t)]


class TestListaProporcionalni:
    def test_glava(self, kandidat, kandidatka, nastavitve):
        doc = obrazci.lista_proporcionalni([kandidat, kandidatka], nastavitve, volilna_enota="Volilna enota 2")
        assert najdi(doc, "B. Volilna enota").endswith("Volilna enota 2")
        assert "vpisanih: 2 " in najdi(doc, "C. V listo")
        assert najdi(doc, "E . Ime liste").endswith("Lista Primer")
        assert najdi(doc, "F. Organ").endswith("Svet stranke")
        assert "Volitve v občinski svet" in oznaceno(doc)
        assert "b) Skupna lista" in oznaceno(doc)

    def test_predstavnik(self, kandidat, nastavitve):
        doc = obrazci.lista_proporcionalni([kandidat], nastavitve)
        t = odstavki(doc)
        i = next(i for i, x in enumerate(t) if "Predstavnik kandidature" in x)
        blok = "\n".join(t[i:i + 10])
        p = nastavitve["predstavnik"]
        for v in (p["emso"], "Ime: Ana Priimek: Predstavnica", p["datum_rojstva"], p["naslov1"], p["naslov2"],
                  p["eposta"], p["telefon"]):
            assert v in blok

    def test_kandidati_po_vrsti(self, kandidat, kandidatka, nastavitve):
        doc = obrazci.lista_proporcionalni([kandidatka, kandidat], nastavitve)
        vrstice = emso_vrstice(doc)
        assert len(vrstice) == 2
        assert vrstice[0].startswith("1.") and kandidatka["emso"] in vrstice[0]
        assert vrstice[1].startswith("2.") and kandidat["emso"] in vrstice[1]
        assert "Ž" in oznaceno(doc) and "M" in oznaceno(doc)
        assert kandidat["raven"].strip() in [x.strip() for x in oznaceno(doc)]

    @pytest.mark.parametrize("n", [1, 3, 15, 16, 40])
    def test_stevilo_mest(self, kandidat, nastavitve, n):
        doc = obrazci.lista_proporcionalni([kandidat] * n, nastavitve)
        vrstice = emso_vrstice(doc)
        assert len(vrstice) == n
        assert [int(re.match(r"\s*(\d+)", v)[1]) for v in vrstice] == list(range(1, n + 1))
        assert sum(1 for t in odstavki(doc) if "Delo, ki ga opravlja: inženir" in t) == n
        assert najdi(doc, "I. Listo kandidatov")  # zaključni del ostane

    def test_cetrtna_skupnost(self, kandidat, nastavitve):
        doc = obrazci.lista_proporcionalni([kandidat], nastavitve, tip="četrtna", naziv_skupnosti="ČS Center")
        assert najdi(doc, "Volitve v svet četrtne skupnosti").endswith("ČS Center")
        assert "Volitve v svet četrtne skupnosti" in "".join(oznaceno(doc))
        assert "Volitve v občinski svet" not in oznaceno(doc)


class TestKandidaturaVecinski:
    @pytest.mark.parametrize("tip,besedilo", [("krajevna", "Volitve v svet krajevne skupnosti"),
                                               ("vaška", "Volitve v svet vaške skupnosti")])
    def test_skupnosti(self, kandidat, nastavitve, tip, besedilo):
        doc = obrazci.kandidatura_vecinski([kandidat], nastavitve, tip=tip, naziv_skupnosti="Podkoren")
        assert najdi(doc, besedilo).endswith("Podkoren")
        assert besedilo in "".join(oznaceno(doc))

    def test_stevilo_mest(self, kandidat, nastavitve):
        for n in (1, 7, 9):
            assert len(emso_vrstice(obrazci.kandidatura_vecinski([kandidat] * n, nastavitve))) == n

    def test_prazna_ve_ostane_prazna(self, kandidat, nastavitve):
        doc = obrazci.kandidatura_vecinski([kandidat], nastavitve)
        assert najdi(doc, "B. Volilna enota").rstrip().endswith("_")

    def test_podpis(self, kandidat, nastavitve):
        doc = obrazci.kandidatura_vecinski([kandidat], nastavitve)
        assert najdi(doc, "Kraj:").startswith("Kraj: Ljubljana")
        assert najdi(doc, "Datum:").startswith("Datum: 10. 10. 2026")


class TestZupan:
    def test_kandidatura(self, kandidatka, nastavitve):
        doc = obrazci.kandidatura_zupan(kandidatka, nastavitve)
        t = "\n".join(odstavki(doc))
        d = t.index("D. Kandidat")
        assert nastavitve["predstavnik"]["emso"] in t[:d]
        assert kandidatka["emso"] in t[d:]
        assert "Ime: Mojca Priimek: Kovač" in t[d:]
        assert "Delo, ki ga opravlja: študentka" in t
        assert "Strokovni ali znanstveni naslov: gimnazijska maturantka" in t
        assert "b) Skupna lista" in oznaceno(doc) and "Ž" in oznaceno(doc)

    def test_soglasje(self, kandidatka, nastavitve):
        doc = obrazci.soglasje_zupan(kandidatka, "Maribor", nastavitve)
        t = [x.strip() for x in odstavki(doc)]
        assert "Kandidat Mojca Kovač" in t
        assert f"Datum rojstva  12. 11. 1999, EMŠO {kandidatka['emso']}," in t
        assert t[t.index("da soglašam s kandidaturo za župana občine:") + 1] == "Maribor"
        assert t[t.index("kot kandidat/kandidatka naslednjega predlagatelja:") + 1] == "Primer stranka"
        assert "V Ljubljana, dne 10. 10. 2026" in t


class TestSoglasjeSvet:
    def test_izpolnjeno(self, kandidat, nastavitve):
        doc = obrazci.soglasje_svet(kandidat, "Ljubljana", nastavitve, "Lista Primer")
        t = [x.strip() for x in odstavki(doc)]
        assert "Kandidat Janez Novak" in t
        assert "Naslov stalnega/začasnega prebivališča: Slovenska cesta 10" in t
        assert "1000 Ljubljana" in t
        assert any(x.endswith("občinskega sveta občine: Ljubljana") for x in t)
        assert any(x.endswith("listi kandidatov: Lista Primer") for x in t)

    def test_brez_kraja_ostanejo_podcrtaji(self, kandidat, nastavitve):
        nastavitve.update(kraj="", datum="")
        doc = obrazci.soglasje_svet(kandidat, "Ljubljana", nastavitve)
        assert najdi(doc, "V_").startswith("V___")


def test_predloge_se_ne_spremenijo(kandidat, nastavitve):
    pred = (obrazci.TEMPLATES / "soglasje_svet.docx").read_bytes()
    obrazci.soglasje_svet(kandidat, "Ljubljana", nastavitve)
    assert (obrazci.TEMPLATES / "soglasje_svet.docx").read_bytes() == pred
