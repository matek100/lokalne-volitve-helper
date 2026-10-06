import datetime as dt

import pytest

from lv import podatki as P

R = P.RAVNI_IZOBRAZBE
OBCINE = ["Ljubljana", "Maribor", "Koper", "Straža", "Kranjska Gora", "Rače-Fram", "Novo mesto"]


# --- EMŠO --------------------------------------------------------------------------------------
class TestEmso:
    def test_sestavi_je_veljaven(self):
        for d, m, y, s in [(1, 1, 1950, "M"), (31, 12, 2000, "Ž"), (29, 2, 1996, "M"), (5, 5, 2008, "Ž")]:
            e = P.sestavi_emso(d, m, y, s)
            assert len(e) == 13
            assert P.preveri_emso(e, dan_volitev=None) == []
            assert P.emso_datum(e) == dt.date(y, m, d)
            assert P.emso_spol(e) == s

    def test_znan_veljaven(self):
        assert P.emso_veljaven("0304985501237")

    def test_vodilna_nicla(self):
        assert P.emso_norm("304985501237") == "0304985501237"
        assert P.emso_norm(304985501237.0) == "0304985501237"  # Excel ga prebere kot število
        assert P.preveri_emso("304985501237") == []

    def test_presledki_in_locila(self):
        assert P.emso_norm("0304 985 501 237") == "0304985501237"
        assert P.preveri_emso("0304-985-501237") == []

    @pytest.mark.parametrize("vnos,del_napake", [
        ("", "manjka"),
        ("12345", "5 števk"),
        ("03049855012370", "14 števk"),
        ("0304985501238", "kontrolna števka"),
        ("03049855O1237", "nedovoljene znake"),
        ("3102990500001", "neveljaven datum"),
    ])
    def test_napake(self, vnos, del_napake):
        napake = P.preveri_emso(vnos)
        assert any(del_napake in n for n in napake), napake

    def test_ostanek_10(self):
        # poiščemo prvih 12 števk, pri katerih je kontrolni ostanek 10
        for zap in range(1000):
            e12 = f"0101990500{zap:03d}"[:12]
            if P.emso_kontrolna(e12) is None:
                break
        assert any("ne more obstajati" in n for n in P.preveri_emso(e12 + "0"))

    def test_polnoletnost(self):
        dan = P.DAN_VOLITEV
        polnoleten = P.sestavi_emso(dan.day, dan.month, dan.year - 18)
        en_dan_premlad = P.sestavi_emso(dan.day + 1, dan.month, dan.year - 18)
        assert P.preveri_emso(polnoleten) == []
        assert any("polnoleten" in n for n in P.preveri_emso(en_dan_premlad))

    def test_neujemanje_datuma_in_spola(self):
        e = P.sestavi_emso(3, 4, 1985, "M")
        assert P.preveri_emso(e, datum_rojstva="3. 4. 1985", spol="M") == []
        assert any("datum rojstva" in n for n in P.preveri_emso(e, datum_rojstva="4. 3. 1985"))
        assert any("spol" in n for n in P.preveri_emso(e, spol="Ž"))
        assert P.preveri_emso(e, datum_rojstva=dt.date(1985, 4, 3)) == []

    def test_leto_2000_plus(self):
        assert P.emso_datum(P.sestavi_emso(1, 7, 2001)) == dt.date(2001, 7, 1)
        assert P.emso_datum(P.sestavi_emso(1, 7, 1899)) == dt.date(1899, 7, 1)


# --- naslov ------------------------------------------------------------------------------------
@pytest.mark.parametrize("vnos,ulica,hs,dod,posta,kraj", [
    ("Slovenska cesta 10a, 1000 Ljubljana", "Slovenska cesta", 10, "A", "1000", "Ljubljana"),
    ("Ob železnici 14 Ljubljana", "Ob železnici", 14, None, None, "Ljubljana"),
    ("Domžale, Dragomelj 153c", "Dragomelj", 153, "C", None, "Domžale"),
    ("Ravne pri Mlinšah 11, 1411, Izlake", "Ravne pri Mlinšah", 11, None, "1411", "Izlake"),
    ("Ulica Slavka gruma 108 8000 novo mesto", "Ulica Slavka gruma", 108, None, "8000", "novo mesto"),
    ("Cesta 24. junija 5, 1231 Ljubljana - Črnuče", "Cesta 24. junija", 5, None, "1231", "Ljubljana - Črnuče"),
    ("Podkoren 12", "Podkoren", 12, None, None, ""),
    ("Trg 1, Mirna Peč", "Trg", 1, None, None, "Mirna Peč"),
])
def test_razcleni_naslov(vnos, ulica, hs, dod, posta, kraj):
    n = P.razcleni_naslov(vnos)
    assert (n["ulica"], n["hs"], n["dodatek"], n["posta"], n["kraj"]) == (ulica, hs, dod, posta, kraj)


def test_razcleni_naslov_brez_stevilke():
    n = P.razcleni_naslov("Ajdovščina")
    assert n["hs"] is None
    assert n["vrstica1"] == "Ajdovščina"


def test_vrstici_naslova():
    n = P.razcleni_naslov("Slovenska cesta 10, 1000 Ljubljana")
    assert (n["vrstica1"], n["vrstica2"]) == ("Slovenska cesta 10", "1000 Ljubljana")


# --- občina ------------------------------------------------------------------------------------
@pytest.mark.parametrize("vnos,pricakovano", [
    ("Ljubljana", "Ljubljana"), ("LJUBLJANA ", "Ljubljana"), ("MOL", "Ljubljana"),
    ("Mestna občina Ljubljana", "Ljubljana"), ("MO Ljubljana", "Ljubljana"), ("Občina Straža", "Straža"),
    ("Straza", "Straža"), ("kranjska gora", "Kranjska Gora"), ("Race-Fram", "Rače-Fram"),
    ("Ljubjana", "Ljubljana"),  # tipkarska napaka
])
def test_normaliziraj_obcino(vnos, pricakovano):
    assert P.normaliziraj_obcino(vnos, OBCINE)[0] == pricakovano


def test_obcina_kot_naslov():
    obcina, opomba = P.normaliziraj_obcino("Ulica Slavka Gruma 108 8000 Novo mesto", OBCINE)
    assert obcina == "" and "naslov" in opomba


def test_neznana_obcina():
    obcina, opomba = P.normaliziraj_obcino("Atlantida", OBCINE)
    assert obcina == "" and opomba


# --- izobrazba ---------------------------------------------------------------------------------
@pytest.mark.parametrize("vnos,raven", [
    ("7. stopnja, univ. dipl. inž. el.", R[9]),
    ("Univ. dipl. filozof", R[9]),
    ("V. Gimnazijski maturant", R[5]),
    ("Gimnazijski maturant", R[5]),
    ("5. stopnja, Tehnik mehatronike", R[4]),
    ("6. Stopnja, komercialist", R[6]),
    ("VI.STOPNJA - PROMETNI TEHNIK", R[6]),
    ("VI/2 prva bolonjska stopnja, diplomirana filozofinja (UN)", R[8]),
    ("dipl. ekon. (VS)", R[8]),
    ("Magistrica varstva narave (4. letnik doktorskega študija)", R[10]),
    ("mag. znanosti", R[12]),
    ("dr. med.", R[13]),
    ("Nižja poklicna", R[2]),
    ("osnovna šola", R[1]),
    ("Vizualni umetnik", ""),
    ("", ""),
])
def test_oceni_raven(vnos, raven):
    assert P.oceni_raven(vnos) == raven


@pytest.mark.parametrize("vnos,naziv", [
    ("V. Gimnazijski maturant", "Gimnazijski maturant"),
    ("7. stopnja, univ. dipl. inž. el.", "univ. dipl. inž. el."),
    ("VI.STOPNJA - PROMETNI TEHNIK", "PROMETNI TEHNIK"),
    ("Vizualni umetnik", "Vizualni umetnik"),
    ("V.", ""),
])
def test_strokovni_naslov(vnos, naziv):
    assert P.strokovni_naslov(vnos) == naziv


# --- ostalo ------------------------------------------------------------------------------------
def test_razdeli_ime():
    assert P.razdeli_ime("Janez Novak") == ("Janez", "Novak")
    assert P.razdeli_ime("  Janez   Novak Kranjc ") == ("Janez", "Novak Kranjc")
    assert P.razdeli_ime("Janez") == ("Janez", "")


def test_cell_in_da():
    assert P.cell(None) == "" and P.cell(float("nan")) == "" and P.cell(5.0) == "5"
    assert P.cell(dt.date(2026, 1, 2)) == "2. 1. 2026"
    assert P.da("Da") and P.da(" da ") and not P.da("Ne") and not P.da("")


def test_datum_iz_besedila():
    assert P.datum_iz_besedila("Ljubljana, 3. 4. 1985") == dt.date(1985, 4, 3)
    assert P.datum_iz_besedila("3.4.1985") == dt.date(1985, 4, 3)
    assert P.datum_iz_besedila("Ljubljana") is None
    assert P.datum_iz_besedila("31. 2. 1985") is None
