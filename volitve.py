"""Izpolnjevanje obrazcev za lokalne volitve 2026.

  python volitve.py pripravi <izvoz.xlsx|.csv|.ods> [-o priprava.xlsx]
  python volitve.py generiraj [priprava.xlsx] [-o output] [--prepisi]
  python volitve.py preveri-emso <EMŠO ali datoteka> [...]
"""
import argparse
import sys
from pathlib import Path

from lv import podatki
from lv.generiraj import generiraj
from lv.preverjanje import preveri_datoteko
from lv.priprava import pripravi


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="ukaz", required=True)

    a = sub.add_parser("pripravi", help="izvoz obrazca -> urejljiva priprava.xlsx (z volilnimi enotami)")
    a.add_argument("vhod")
    a.add_argument("-o", "--izhod", default="priprava.xlsx")
    a.add_argument("--brez-eprostor", action="store_true", help="ne poizveduj po registru naslovov")

    b = sub.add_parser("generiraj", help="priprava.xlsx -> izpolnjeni obrazci .docx po občinah")
    b.add_argument("priprava", nargs="?", default="priprava.xlsx")
    b.add_argument("-o", "--izhod", default="output")
    b.add_argument("--prepisi", action="store_true", help="izbriši in na novo ustvari izhodno mapo")

    c = sub.add_parser("preveri-emso", help="preveri EMŠO (posamezne številke ali vse v datoteki)")
    c.add_argument("vnosi", nargs="+", help="EMŠO ali pot do priprava.xlsx / izvoza")

    args = ap.parse_args()
    if args.ukaz == "pripravi":
        pripravi(args.vhod, args.izhod, uporabi_eprostor=not args.brez_eprostor)
    elif args.ukaz == "generiraj":
        generiraj(args.priprava, args.izhod, prepisi=args.prepisi)
    else:
        sys.exit(preveri_emso(args.vnosi))


def preveri_emso(vnosi):
    """Izpiše napake; izhodna koda 1, če je kateri EMŠO neveljaven."""
    napacnih = 0
    for v in vnosi:
        if Path(v).is_file():
            napake = preveri_datoteko(v)
            print(f"{v}: {'vsi EMŠO so veljavni' if not napake else f'{len(napake)} težav'}")
            for kdo, e, nn in napake:
                print(f"  {kdo} ({e}):")
                for n in nn:
                    print(f"    - {n}")
            napacnih += len(napake)
        else:
            nn = podatki.preveri_emso(v)
            e = podatki.emso_norm(v)
            if nn:
                print(f"{v}: NEVELJAVEN")
                for n in nn:
                    print(f"    - {n}")
            else:
                d = podatki.emso_datum(e)
                print(f"{v}: veljaven (rojen {podatki.datum_str(d)}, spol {podatki.emso_spol(e)})")
            napacnih += bool(nn)
    return 1 if napacnih else 0


if __name__ == "__main__":
    main()
