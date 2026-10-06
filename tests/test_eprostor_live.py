"""Testi proti pravemu servisu eProstor (GURS). Zaženi z: python -m pytest --network

Preverijo, da se oblika servisa ni spremenila (imena slojev in polj)."""
import pytest

from lv.eprostor import EProstor

pytestmark = pytest.mark.network


@pytest.fixture(scope="module")
def ep(tmp_path_factory):
    return EProstor(cache_path=tmp_path_factory.mktemp("cache") / "eprostor.json")


def test_sifranti(ep):
    assert len(ep.obcine) == 212
    assert "Ljubljana" in ep.obcine.values()
    assert len(ep.volilne_enote) > 212
    assert ep.st_volilnih_enot("Straža") > 1


def test_naslov_z_ulico(ep):
    z, opomba = ep.najdi_naslov("Slovenska cesta", 10, posta="1000", kraj="Ljubljana")
    assert z and z["OBCINA_NAZIV"] == "Ljubljana" and opomba == ""
    assert z["EID_LOKALNA_VOLILNA_ENOTA"] in ep.volilne_enote
    assert ep.skupnost("četrtna", z["EID_CETRTNA_SKUPNOST"]) == "Četrtna skupnost Center"


def test_naslov_brez_ulice(ep):
    z, _ = ep.najdi_naslov("Podkoren", 12)
    assert z and z["OBCINA_NAZIV"] == "Kranjska Gora" and z["ULICA_NAZIV"] is None


def test_priblizno_ujemanje(ep):
    z, opomba = ep.najdi_naslov("Gradiska ul.", 29, posta="8351")
    assert z and z["ULICA_NAZIV"] == "Gradiška ulica"
    assert "približno" in opomba


def test_razlocitev_po_obcini(ep):
    z, _ = ep.najdi_naslov("Rimska cesta", 9, obcina="Ljubljana")
    assert z and z["OBCINA_NAZIV"] == "Ljubljana"


def test_ni_najden(ep):
    z, opomba = ep.najdi_naslov("Neobstoječa ulica Xyz", 9999)
    assert z is None and "ni najden" in opomba
