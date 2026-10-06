# Navodila za uporabo – izpolnjevanje obrazcev za lokalne volitve 2026

Ta priročnik opisuje celoten postopek: od izvoza prijav kandidatov do izpolnjenih Wordovih
obrazcev, ki jih stranka vloži na občinsko volilno komisijo (ObVK).

**Vsebina**

1. [Kaj orodje naredi](#1-kaj-orodje-naredi)
2. [Namestitev](#2-namestitev)
3. [Hiter začetek](#3-hiter-začetek)
4. [Vhodni podatki](#4-vhodni-podatki)
5. [Korak 1: priprava](#5-korak-1-priprava-pripravi)
6. [Pregled in dopolnjevanje priprava.xlsx](#6-pregled-in-dopolnjevanje-pripravaxlsx)
7. [Opombe in kaj z njimi narediti](#7-opombe-in-kaj-z-njimi-narediti)
8. [Korak 2: generiranje obrazcev](#8-korak-2-generiranje-obrazcev-generiraj)
9. [Kaj se izpolni v posameznem obrazcu](#9-kaj-se-izpolni-v-posameznem-obrazcu)
10. [Preverjanje EMŠO](#10-preverjanje-emšo)
11. [Nove prijave in ponovni zagon](#11-nove-prijave-in-ponovni-zagon)
12. [Pogosta vprašanja in težave](#12-pogosta-vprašanja-in-težave)
13. [Za razvijalce](#13-za-razvijalce)

---

## 1. Kaj orodje naredi

```
 izvoz prijav (.xlsx/.csv/.ods)
        │
        ▼
 python volitve.py pripravi ──► priprava.xlsx  ◄── ročni pregled in dopolnitve
        │   (register naslovov GURS:                 (volilni sistem, predlagatelj,
        │    volilna enota, ČS/KS, uradni naslov)     predstavnik, podatki o seji ...)
        ▼
 python volitve.py generiraj ──► output/<Občina>/*.docx  ◄── ročni popravki v Wordu, podpisi
```

Orodje izpolni naslednje obrazce DVK (instruktivni obrazci LV 2026):

| Obrazec | Predloga | Kdaj se ustvari |
|---|---|---|
| Lista kandidatov (proporcionalni sistem) | `templates/lista_proporcionalni.docx` | občinski svet s proporcionalnim sistemom – ena na volilno enoto; svet ČS, če je izbran proporcionalni sistem |
| Kandidatura (večinski sistem) | `templates/kandidatura_vecinski.docx` | občinski svet z večinskim sistemom – ena na volilno enoto; sveti KS, VS in ČS (privzeto) |
| Kandidatura za župana | `templates/kandidatura_zupan.docx` | vsak kandidat za župana |
| Soglasje s kandidaturo za člana občinskega sveta | `templates/soglasje_svet.docx` | vsak kandidat za občinski svet |
| Soglasje s kandidaturo za župana | `templates/soglasje_zupan.docx` | vsak kandidat za župana |
| Zapisnik o delu organa stranke – lista (proporcionalni) | `templates/zapisnik_lista.docx` | ob vsaki listi kandidatov |
| Zapisnik o delu organa stranke – večinski sistem | `templates/zapisnik_vecinski.docx` | ob vsaki kandidaturi po večinskem sistemu |
| Zapisnik o delu organa stranke – župan | `templates/zapisnik_zupan.docx` | ob vsaki kandidaturi za župana |

Zapisniki se ne ustvarijo, če je tip predlagatelja *Skupina volivcev* (ta kandidature podpre
s podpisi, ne z odločitvijo organa).

Vsi izhodni dokumenti so `.docx`, da jih lahko pred tiskom še ročno popravite. Podatki, ki jih
orodje ne pozna (npr. število glasov na seji), ostanejo podčrtani za ročni vpis.

## 2. Namestitev

Potrebujete **Python 3.10 ali novejši** ([python.org](https://www.python.org/downloads/); pri
namestitvi na Windows obkljukajte *Add Python to PATH*).

V mapi projekta zaženite:

```bash
pip install -r requirements.txt
```

Preverite namestitev:

```bash
python volitve.py --help
```

Za korak 1 je potreben dostop do interneta (javni servis eProstor). Korak 2 deluje brez interneta.

## 3. Hiter začetek

```bash
python volitve.py pripravi "Evidentacija kandidatov (responses).xlsx"
```

Odprite `priprava.xlsx` in:

1. na listu **Občine** za vsako občino izberite **Volilni sistem** (proporcionalni / večinski),
2. na listu **Nastavitve** vpišite predlagatelja, organ, ime liste, predstavnika in podatke o seji,
3. na listu **Kandidati** preglejte stolpec **Opombe** in oranžne stolpce.

```bash
python volitve.py generiraj
```

Obrazci so v mapi `output/`. Seznam ustvarjenih datotek in opozorila sta v `output/povzetek.txt`.

Celoten primer z izmišljenimi podatki je v mapi [`primer/`](../primer):

```bash
python volitve.py generiraj primer/priprava_primer.xlsx -o output_primer
```

## 4. Vhodni podatki

Vhod je izvoz odgovorov obrazca za evidentiranje kandidatov (Google Forms, Microsoft Forms ...)
v obliki `.xlsx`, `.csv` ali `.ods`. Potrebni so stolpci (ime stolpca mora biti enako, velike
črke in odvečni presledki niso pomembni):

| Stolpec | Primer | Opomba |
|---|---|---|
| Ime in priimek | `Ana Marija Horvat` | prva beseda postane ime, ostalo priimek – preverite dvojna imena |
| Naslov in kraj stalnega bivališča | `Slovenska cesta 10, 1000 Ljubljana` | dopušča tudi `Ob železnici 14 Ljubljana`, `Domžale, Dragomelj 153c`, okrajšave (`ul.`, `c.`) in manjkajoče šumnike |
| Stopnja in naziv izobrazbe | `7. stopnja, univ. dipl. inž. el.` | prosto besedilo; preslika se v raven izobrazbe |
| Kraj in datum rojstva | `Ljubljana, 3. 4. 1985` | datum se prebere iz besedila, sicer iz EMŠO |
| Delo, ki ga opravljate | `inženir` | |
| EMŠO | `0304985500001` | izgubljena vodilna ničla (Excel) se doda samodejno |
| Občina kjer želite kandidirati v občinski svet | `MOL`, `Mestna občina Ljubljana` | prazno = ne kandidira v občinski svet |
| Želite kandidirati v četrtni/krajevni svet? | `Da` / `Ne` | |
| Želite kandidirati za župana? | `Da` / `Ne` | |

Ostali stolpci (časovni žig, kontakt, motivacija ...) se ne uporabijo.

## 5. Korak 1: priprava (`pripravi`)

```bash
python volitve.py pripravi <izvoz> [-o priprava.xlsx] [--brez-eprostor]
```

Za vsakega kandidata orodje:

- **poišče naslov v registru naslovov GURS** (eProstor) in iz njega prevzame:
  - **volilno enoto** za občinski svet (sloj `LOKALNE_VOLILNE_ENOTE`, veljaven za LV 2026),
  - **četrtno, krajevno ali vaško skupnost**,
  - **uradni zapis naslova** (npr. `Gradiska ul. 29, 8351 Straza` → `Gradiška ulica 29`,
    `8351 Straža pri Novem mestu`; popravi tudi napačno poštno številko),
  - **dele naslova** za zapisnike: občina, naselje, ulica, hišna številka;
- **prepozna občino** (`MOL`, `LJUBLJANA `, `Straza` → uradni naziv); če je v polju občine
  vpisan naslov, prevzame občino iz naslova;
- iz **EMŠO** izračuna datum rojstva in spol ter **preveri EMŠO** (glej [10](#10-preverjanje-emšo));
- **oceni raven izobrazbe** iz prostega besedila in loči strokovni naslov;
- določi **vrstni red** na listi (po vrstnem redu prijav znotraj občine in volilne enote);
- označi **podvojene prijave** (isti EMŠO – velja zadnja).

Z `--brez-eprostor` korak deluje brez interneta, a volilne enote in skupnosti ostanejo prazne.

Odgovori servisa se shranjujejo v `.cache/eprostor.json`, zato je ponovni zagon hiter.

## 6. Pregled in dopolnjevanje `priprava.xlsx`

Barve glav stolpcev: **siva** – izvirni podatki iz prijave, **rumena** – izračunano,
**oranžna** – preverite oz. dopolnite ročno. **Rdeča celica EMŠO** pomeni neveljaven EMŠO.

### List *Kandidati*

| Stolpec | Pomen |
|---|---|
| Vključi | `Ne` izključi kandidata iz vseh obrazcev (npr. odstopil) |
| Ime, Priimek | razdeljeno iz »Ime in priimek« – popravite dvojna imena |
| EMŠO (13 mest) | normaliziran EMŠO (vedno besedilo, da Excel ne izgubi ničel) |
| Datum rojstva, Spol | iz prijave / EMŠO |
| Naslov - ulica in hišna št., Naslov - pošta in kraj | dve vrstici naslova na obrazcih |
| Raven izobrazbe | spustni seznam z ravnmi, kot so na obrazcu; ta se na obrazcu obkroži |
| Strokovni ali znanstveni naslov | npr. `univ. dipl. inž. el.` |
| Delo | delo, ki ga opravlja |
| Občinski svet | `Da` = kandidira v občinski svet |
| **Občina** | občina kandidature (za svet in župana) |
| **Volilna enota** | npr. `Volilna enota 2`; pomembna le v občinah z več VE |
| **Vrstni red** | mesto na listi znotraj občine in VE (1 = nosilec liste) |
| Župan | `Da` = kandidira za župana |
| Svet ČS/KS/VS | `Da` = kandidira v svet ožjega dela občine |
| **Tip skupnosti**, **Naziv skupnosti** | `četrtna` / `krajevna` / `vaška` in ime skupnosti |
| **Volilna enota skupnosti**, **Vrstni red v skupnosti** | VE skupnosti ni v registru – vpišite ročno, če jih je več |
| Občina po naslovu | občina stalnega prebivališča (iz registra) |
| Naslov - naselje, Naslov - ulica, Naslov - hišna št. | deli naslova za zapisnike |
| Opombe | kaj je treba preveriti (glej [7](#7-opombe-in-kaj-z-njimi-narediti)) |

Kandidate lahko tudi ročno **dodate** (nova vrstica z izpolnjenimi rumenimi/oranžnimi stolpci).

### List *Občine*

| Stolpec | Pomen |
|---|---|
| Občina | uradni naziv |
| Št. volilnih enot | iz registra; če je 1, se VE na obrazcih ne vpisuje (v zapisnikih se vpiše »1«) |
| **Volilni sistem** | **obvezno**: `proporcionalni` ali `večinski` – brez tega se lista/kandidatura za občinski svet ne ustvari |
| Volilni sistem ČS | `večinski` (privzeto) ali `proporcionalni` za svete četrtnih skupnosti |
| ostali stolpci | enaki kot na listu *Nastavitve*; neprazna vrednost prepiše privzeto samo za to občino (npr. drug predstavnik, drug kraj in datum seje lokalnega odbora) |

Volilni sistem za posamezno občino je določen z zakonom in statutom občine (manjši občinski
sveti se praviloma volijo po večinskem sistemu) – preverite ga v statutu ali pri ObVK.

### List *Nastavitve*

| Nastavitev | Uporaba |
|---|---|
| Predlagatelj | ime stranke; soglasja, zapisniki (»Ime politične stranke«) |
| Tip predlagatelja | Politična stranka / Skupna lista / Skupina volivcev – obkroži se na obrazcih |
| Organ | organ, ki je določil kandidaturo (npr. Svet stranke) |
| Ime liste | proporcionalni sistem; prazno = ime predlagatelja |
| Predstavnik EMŠO, ime, priimek, datum rojstva, naslov, občina, e-pošta, telefon | predstavnik kandidature na vseh obrazcih in zapisnikih |
| Kraj podpisa, Datum podpisa | kraj in datum na obrazcih in soglasjih; prazno = ročni vpis |
| Število podpisov | samo pri kandidaturah, ki jih podpirajo volivci |
| Seja - datum, ura, kraj | zapisnik: kdaj in kje je bila seja organa |
| Seja - sklical, vodil, zapisnikar, overovatelj 1/2 | zapisnik: osebe (ime, priimek, naslov) |
| Seja - navzočih, vabljenih, ura konca | zapisnik: sklepčnost in konec seje |
| Statut - člen sklepčnosti, člen določitve | zapisnik: člena statuta (npr. `20.`) |
| Znak stranke v imenu liste | zapisnik liste: `Da` / `Ne` (obkroži JE / NI) |

## 7. Opombe in kaj z njimi narediti

| Opomba | Kaj narediti |
|---|---|
| napačna kontrolna števka EMŠO / EMŠO ima N števk / neveljaven datum v EMŠO | preverite EMŠO pri kandidatu – ObVK kandidature z napačnim EMŠO ne sprejme |
| na dan volitev kandidat ne bo polnoleten | kandidat ne more kandidirati |
| datum rojstva se ne ujema z EMŠO | eden od podatkov je napačen |
| naslov ni najden v registru naslovov | preverite naslov; vpišite volilno enoto in ČS/KS ročno |
| naslov ni razčlenjen (manjka hišna številka) | v prijavi ni hišne številke – dopolnite naslov in VE |
| približno ujemanje imena ulice / naselja | preverite, da je najden pravi naslov (stolpca Naslov) |
| hišna številka se razlikuje v dodatku | npr. prijava `12`, v registru `12a` – preverite |
| poštna številka se ne ujema | naslov je bil najden z drugo pošto; uradna pošta je že vpisana |
| več možnih naslovov, izbran prvi | ista ulica obstaja v več krajih – preverite občino in VE |
| občina ugotovljena približno / ni bilo mogoče prepoznati / je videti kot naslov | preverite stolpec Občina |
| občina prevzeta iz naslova | v prijavi ni bilo prepoznavne občine; uporabljena je občina prebivališča |
| kandidira v drugi občini kot prebiva | VE ni izpolnjena – vpišite jo ročno (in preverite, ali kandidat sploh lahko kandidira v tej občini) |
| ČS/KS po naslovu ni najdena | občina morda nima ožjih delov; vpišite tip in naziv ročno ali nastavite »Svet ČS/KS/VS« na Ne |
| preveri raven izobrazbe | besedila ni bilo mogoče zanesljivo preslikati – izberite raven iz seznama |
| podvojena prijava (velja zadnja) | kandidat se je prijavil večkrat; ostala je zadnja prijava |

## 8. Korak 2: generiranje obrazcev (`generiraj`)

```bash
python volitve.py generiraj [priprava.xlsx] [-o output] [--prepisi]
```

Struktura izhoda:

```
output/
  povzetek.txt
  Ljubljana/
    Lista kandidatov - občinski svet.docx
    Zapisnik - lista kandidatov - občinski svet.docx
    Soglasje - občinski svet - Horvat Ana Marija.docx
    ...
    Četrtna skupnost Center/
      Kandidatura - Četrtna skupnost Center.docx
      Zapisnik - kandidatura - Četrtna skupnost Center.docx
  Straža/
    Kandidatura - občinski svet - Volilna enota 7.docx
    Zapisnik - kandidatura - občinski svet - Volilna enota 7.docx
    Kandidatura za župana - Golob Marko.docx
    Soglasje - župan - Golob Marko.docx
    Zapisnik - župan - Golob Marko.docx
```

- **Varnost ročnih popravkov:** če mapa `output/` že obstaja in ni prazna, se ukaz ustavi.
  Z `--prepisi` se mapa izbriše in ustvari znova; z `-o druga_mapa` pišete drugam.
- **Opozorila** (izpisana in v `povzetek.txt`): neizbran volilni sistem, kandidati brez VE v
  občini z več VE, podvojen ali neštevilski vrstni red, neveljavni EMŠO (tudi predstavnika),
  manjkajoča občina ali skupnost.

## 9. Kaj se izpolni v posameznem obrazcu

»Obkroži« je na obrazcih izvedeno kot **krepko besedilo v okvirju**.

**Lista kandidatov / Kandidatura (večinski):** tip glasovanja (obkroženo, z nazivom skupnosti),
volilna enota (le pri več VE), število kandidatov, tip predlagatelja (obkroženo), ime liste,
organ, predstavnik kandidature, za vsakega kandidata EMŠO, ime, priimek, datum rojstva, spol
(obkroženo), naslov, raven izobrazbe (obkroženo), strokovni naslov, delo; kraj in datum.
Število mest za kandidate se prilagodi (odvečna mesta se odstranijo, manjkajoča dodajo).
*Ostane prazno:* obče ime, število podpisov (razen če je nastavljeno), podpis.

**Kandidatura za župana:** kot zgoraj za enega kandidata.

**Soglasja:** ime in priimek, datum rojstva, EMŠO, naslov, občina, predlagatelj oz. ime liste,
kraj in datum. *Ostane prazno:* podpis kandidata.

**Zapisniki:** obkrožena vrsta zapisnika, občina, številka VE, ime stranke, organ, datum, ura
in kraj seje, sklicatelj, vodja, zapisnikar, overovatelja, sklepčnost, dnevni red (točka 1),
predlog kandidatov (ime in naslov), seznam za glasovanje (imena), člen statuta, določeni
kandidati z vsemi podatki in obkroženo ravnjo izobrazbe, ime liste in znak stranke, predstavnik,
ura konca. Ohrani se samo blok za eno volilno enoto. *Ostane prazno:* izid glasovanja (število
glasov ZA/PROTI, neveljavne glasovnice), način glasovanja in razlog izbire (»ker je«), dodatne
točke dnevnega reda, podpisi. Priložiti je treba seznam udeležencev seje.

## 10. Preverjanje EMŠO

```bash
python volitve.py preveri-emso 0304985500001
python volitve.py preveri-emso priprava.xlsx
python volitve.py preveri-emso "Evidentacija kandidatov (responses).xlsx"
```

Preveri se: 13 števk (ena izgubljena vodilna ničla se doda), dovoljeni znaki, veljaven datum
rojstva, ki ni v prihodnosti, kontrolna števka (modul 11), polnoletnost na dan volitev
(15. 11. 2026), ujemanje z vpisanim datumom rojstva in spolom ter podvojeni EMŠO.
Pri datoteki `priprava.xlsx` se preverijo kandidati z `Vključi = Da` in vsi predstavniki
kandidature. Izhodna koda je 1, če je kateri EMŠO neveljaven.

Isto preverjanje se izvede samodejno ob `pripravi` in `generiraj`.

## 11. Nove prijave in ponovni zagon

Ko dobite nov izvoz prijav, zaženite `pripravi` z isto izhodno datoteko:

```bash
python volitve.py pripravi "nov izvoz.xlsx"
```

- obstoječi kandidati (prepoznani po EMŠO) ostanejo **točno taki, kot ste jih uredili**,
- novi kandidati se dodajo na konec in dobijo naslednjo prosto številko v vrstnem redu,
- ročno dodani kandidati, ki jih ni v izvozu, ostanejo,
- listi *Občine* in *Nastavitve* se ohranita; nove občine se dodajo.

Če želite kandidata obdelati na novo (npr. po popravku naslova v prijavi), izbrišite njegovo
vrstico v `priprava.xlsx` in ponovno zaženite `pripravi`.

## 12. Pogosta vprašanja in težave

**Excel prikaže EMŠO kot 3,05E+12 ali brez ničle.** V `priprava.xlsx` je EMŠO shranjen kot
besedilo. Če ga prepisujete ročno, vnesite ga z apostrofom (`'0304985500001`) ali celico
oblikujte kot besedilo. Izgubljeno vodilno ničlo orodje doda samo.

**Volilna enota je prazna.** Naslov ni bil najden ali kandidat kandidira v drugi občini – glej
stolpec Opombe in vpišite VE ročno (npr. `Volilna enota 3`; za zapisnik se uporabi številka).

**Lista se ni ustvarila.** Na listu *Občine* ni izbran volilni sistem – glej `povzetek.txt`.

**Napaka »Mapa 'output' ni prazna«.** Namenoma – zaščita ročnih popravkov. Uporabite `--prepisi`
ali `-o nova_mapa`.

**Napaka »V vhodni datoteki manjkajo stolpci«.** Ime stolpca v izvozu se razlikuje od
pričakovanega (glej [4](#4-vhodni-podatki)); preimenujte stolpec v izvozu.

**Napaka eProstor / brez interneta.** Ob prekinjeni povezavi orodje poizvedbo samo ponovi do 5-krat. Kandidati, pri katerih kljub temu ne uspe, dobijo opombo; zaženite `pripravi`
znova, ko je povezava na voljo (izbrišite njihove vrstice), ali uporabite `--brez-eprostor`.

**Datoteke se ne da prepisati (Permission denied).** Datoteka je odprta v Excelu/Wordu – zaprite jo.

**Obkrožena možnost ni prava.** Popravite v Wordu: odstranite okvir (Oblikovanje → Obrobe) ali
popravite podatek v `priprava.xlsx` in generirajte znova.

## 13. Za razvijalce

### Struktura

```
volitve.py            ukazna vrstica (pripravi, generiraj, preveri-emso)
lv/podatki.py         čiščenje podatkov: EMŠO, naslov, občina, izobrazba, ime
lv/eprostor.py        WFS poizvedbe GURS (register naslovov, prostorske enote) + predpomnilnik
lv/priprava.py        korak 1: izvoz -> priprava.xlsx (+ ohranjanje ročnih popravkov)
lv/docxfill.py        nizkonivojsko izpolnjevanje Worda: podčrtaji, okvirji, rune
lv/obrazci.py         kandidature, liste, soglasja
lv/zapisniki.py       zapisniki o delu organa stranke
lv/generiraj.py       korak 2: priprava.xlsx -> output/
lv/preverjanje.py     preverjanje EMŠO v datotekah
templates/            predloge .docx (izvirniki .doc v templates/src/)
primer/               primeri; ustvari_primere.py jih ustvari znova
tests/                testi (pytest)
```

### Testi

```bash
pip install -r requirements-dev.txt
python -m pytest              # brez omrežja (lažni register naslovov)
python -m pytest --network    # tudi testi proti pravemu servisu eProstor
```

Primere ustvarite znova z `python primer/ustvari_primere.py` (potrebuje eProstor).

### Kako deluje izpolnjevanje

Predloge niso spremenjene z oznakami – orodje poišče besedilo oznake (npr. `EMŠO:`,
`Delo, ki ga opravlja:`) in zamenja prvi niz podčrtajev za njo. Besedilo odstavka je v Wordu
razdeljeno na rune z različnim oblikovanjem; `docxfill.replace_span` spremembo preslika nazaj
na rune, tako da se oblikovanje ohrani. »Obkroži« doda runu okvir (`w:bdr`) in krepko pisavo.
Mesta za kandidate so prepoznana po vzorcu (`N. EMŠO`, `Ime in priimek:`) in se klonirajo.

### Posodobitev za naslednje volitve

1. Nove obrazce DVK shranite v `templates/` (`.doc` pretvorite v `.docx`, npr.
   `soffice --headless --convert-to docx obrazec.doc`), z enakimi imeni datotek.
2. Posodobite `DAN_VOLITEV` v `lv/podatki.py` (preverjanje polnoletnosti).
3. Če so se spremenile oznake na obrazcih ali seznam ravni izobrazbe, posodobite
   `RAVNI_IZOBRAZBE` in regularne izraze v `lv/obrazci.py` / `lv/zapisniki.py`.
4. Zaženite `python -m pytest` – testi v `tests/test_obrazci.py` in `tests/test_zapisniki.py`
   preverijo vsako izpolnjeno polje.
