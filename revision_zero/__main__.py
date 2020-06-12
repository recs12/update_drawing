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


def revision():
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

        revision_doc = get_document_revision(draft)
        print("Document Revision: %s" % revision_doc)

        rev = prompt_revision()
        if rev == "00":
            remove_blocks(draft)
        elif rev == "01":
            insert_blocks(draft)
        elif rev == "testing":
            revision_inspect(draft)

    except AssertionError as err:
        print(err.args)
    except Exception as ex:
        print(ex.args)
    finally:
        raw_input("\nPress any key to exit...")
        sys.exit()


def get_document_revision(draft):
    """Revision of the draft
    """
    return draft.Properties.Item['ProjectInformation']['Revision'].Value


def remove_blocks(draft):
    """Remove the revision blocks and balloones
    """
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


def insert_blocks(draft):
    block_revision = "J:\\PTCR\\_Solidedge\\Draft_Symboles\\Bloc revision - ENGLISH.dft"
    block_triangle = "J:\\PTCR\\_Solidedge\\Draft_Symboles\\ID rev.dft"

    Sheet1 = draft.Sheets[1]
    Sheet1.Activate()
    # insert a revision block
    # insert a triangles
    for _ in draft.Blocks:
        print(_.Name)
    blocks = draft.Blocks
    blocks.AddBlockByFile(block_triangle, 0.25, 0.25)

    for _ in Sheet1.BlockOccurrences:
        print(_)
    # Sheet1.BlockOccurrences.Add(block_revision, 0.00, 0.00)
    draft.Blocks.AddBlockByFile(block_revision)
    # *** code here ***

    # access the revision block and modify the content
    # block = Sheet1.BlockOccurrences.Item[1]
    # for _ in block.BlockLabelOccurrences:
    #     print(_.Name)
    #     _.Value = "0"



def prompt_revision():
    revision = raw_input(
        "\nselect revision:\n\t0) REVISION.00\n\t1) REVISION.01 and above.\n(Press any key to cancel)\n>"
    )
    return {"0": "00", "1": "01"}.get(revision)


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
    confirmation(revision)
