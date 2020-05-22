""" Delete graphic components ID rev and Bloc revision from Drafts.
"""
import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System.Runtime.InteropServices")

import System.Runtime.InteropServices as SRI
import SolidEdgeDraft as se_draft

import sys

blocks_to_delete = [
    "ID rev",
    "ID de REV",
    "Bloc revision",
    "Bloc revision 1",
    "Bloc revision - ENGLISH",
]


def remove_symbols():
    try:
        session = SRI.Marshal.GetActiveObject("SolidEdge.Application")
        print("Author: recs@premiertech.com")
        print("Maintainer: Rechdi, Slimane")
        print("Last update: 2020-04-23")
        print("version solidedge: %s" % session.Value)
        assert session.Value in ["Solid Edge ST7", "Solid Edge 2019"], "Unvalid version of solidedge"
        draft = session.ActiveDocument
        print("part: %s\n" % draft.Name)
        assert draft.Name.lower().endswith(".dft"), (
            "This macro only works on .psm not %s" % draft.Name[-4:]
        )

        rev = prompt_revision()
        if rev == "00":
            revision_00(draft)
        elif rev == "01":
            revision_01(draft)
        elif rev == "testing":
            revision_inspect(draft)

    except AssertionError as err:
        print(err.args)
    except Exception as ex:
        print(ex.args)
    finally:
        raw_input("\nPress any key to exit...")
        sys.exit()


def revision_00(draft):
    for symbol in draft.Blocks:
        if symbol.Name in blocks_to_delete:
            print("[-] %s, \tdeleted" % symbol.Name)
            symbol.delete()

    # for ball in draft.ActiveSheet.Balloons:
    if draft.Balloons:
        for ball in draft.Balloons:
            if ball.BalloonType == 7:  # type 7 filter the triangle balloons.
                print("[-] %s, \tdeleted" % ball.Name)
                ball.Delete()
    else:
        pass


def revision_01(draft):
    for i in draft.Blocks:
        print(i.Name)
    print(dir(draft.Blocks))
    block_revision = r"J:\PTCR\_Solidedge\Draft_Symboles\Bloc revision - ENGLISH.dft"
    # draft.Blocks.AddBlockByFile(block_revision)
    draft.Blocks.Add(block_revision)
    # draft.ActiveSheet.BlockOccurrences.Add(block_revision)


def revision_inspect(draft):
    for i in draft.Blocks:
        print(i.Name)
        print(dir(i))
    print(draft.Sheets["Sheet1"].Name)
    print(draft.Sheets["Sheet1"])
    block_revision = r"J:\PTCR\_Solidedge\Draft_Symboles\Bloc revision - ENGLISH.dft"
    se_draft.BlockOccurrences.Add(block_revision, 0.2, 0.2)



def prompt_revision():
    revision = raw_input(
        "select revision:\n\t0) Rev.00\n\t1) Rev.01\n\t2) Rev.02 and above.\n(press any key to cancel):\n>"
    )
    return {"0": "00", "1": "01", "2": "testing"}.get(revision)


def confirmation(func):
    response = raw_input(
        """Delete graphic components ID rev and Bloc revision,\n(Press y/[Y] to proceed.)\n>"""
    )
    if response.lower() not in ["y"]:
        print("Process canceled")
        sys.exit()
    else:
        func()


if __name__ == "__main__":
    confirmation(remove_symbols)
