""" Delete graphic components ID rev and Bloc revision from Drafts.
"""

import sys

from api import Api


def remove_symbols():
    try:
        session = Api()
        print("Author: recs")
        print("Last update: 2020-04-08")
        session.check_valid_version("Solid Edge ST7", "Solid Edge 2019")
        draft = session.active_document()
        print("part: %s\n" % draft.name)
        assert draft.name.endswith(".dft"), (
            "This macro only works on .psm not %s" % draft.name[-4:]
        )
    except AssertionError as err:
        print(err.args)
    except Exception as ex:
        print(ex.args)
    else:
        for symbol in draft.Blocks:
            if symbol.name in [
                "ID rev",
                "Bloc revision",
                "Bloc revision 1",
                "Bloc revision - ENGLISH",
            ]:
                symbol.delete()
                print(" %s deleted" % symbol.name)
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
