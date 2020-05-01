""" Delete graphic components ID rev and Bloc revision from Drafts.
"""

import sys

from api import Api

blocks_to_delete = [
    "ID rev",
    "ID de REV",
    "Bloc revision",
    "Bloc revision 1",
    "Bloc revision - ENGLISH",
]


def remove_symbols():
    try:
        session = Api()
        print("Author: recs@premiertech.com")
        print("Maintainer: Rechdi, Slimane")
        print("Last update: 2020-04-23")
        session.check_valid_version("Solid Edge ST7", "Solid Edge 2019")
        draft = session.active_document()
        print("part: %s\n" % draft.name)
        assert draft.name.endswith(".dft"), (
            "This macro only works on .psm not %s" % draft.name[-4:]
        )

        for ball in draft.ActiveSheet.Balloons:
            if ball.BalloonType == 7:  # type 7 filter the triangle balloons.
                print("[-] %s, \tdeleted" % ball.name)
                ball.delete()

        for symbol in draft.Blocks:
            if symbol.name in blocks_to_delete:
                print("[-] %s, \tdeleted" % symbol.name)
                symbol.delete()
                
    except AssertionError as err:
        print(err.args)
    except Exception as ex:
        print(ex.args)
    finally:
        raw_input("\n(Press any key to exit ;)")
        sys.exit()


def confirmation(func):
    response = raw_input("""Replace background, (Press y/[Y] to proceed.): """)
    if response.lower() not in ["y"]:
        print("Process canceled")
        sys.exit()
    else:
        func()


if __name__ == "__main__":
    confirmation(remove_symbols)
