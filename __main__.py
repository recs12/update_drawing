# -*- coding: utf-8 -*-

""" Check the revision of the document draft then either
    - Delete graphic blocks components <ID rev> & <Bloc revision>
    - or add <ID rev> & <Bloc revision> if revision is above 00.
"""

from System import Console
import System.Runtime.InteropServices as SRI
import System
import sys
import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System")
clr.AddReference("System.Runtime.InteropServices")


blocks_to_delete = [
    "ID rev",
    "ID de REV",
    "Bloc revision",
    "Bloc revision 1",
    "Bloc revision - ENGLISH",
]


__project__ = "update_drawing"
__author__ = "recs"
__version__ = "0.0.1"
__update__ = "2020-10-30"


def get_document_revision(draft):
    """Revision of the draft
    """
    rev = draft.Properties.Item["ProjectInformation"]["Revision"].Value
    return int(rev)


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


def insert_blocks(draft, current_revision, user):

    # Material
    block_revision = "J:\PTCR\_Solidedge\Draft_Symboles\Bloc revision - ENGLISH.dft"
    block_triangle = "J:\PTCR\_Solidedge\Draft_Symboles\ID rev.dft"

    Sheet1 = draft.Sheets[1]
    Sheet1.Activate()
    blocks = draft.Blocks

    Y = get_height(current_revision)

    # Triangle
    blocks.AddBlockByFile(block_triangle)
    Sheet1.BlockOccurrences.Add("ID rev", 0.298, Y)
    count = Sheet1.BlockOccurrences.Count
    block = Sheet1.BlockOccurrences.Item(count)
    labels = block.BlockLabelOccurrences
    labels[1].Value = current_revision

    # Revision block
    blocks.AddBlockByFile(block_revision)
    Sheet1.BlockOccurrences.Add("Bloc revision - ENGLISH", 0.309499, Y)
    count = Sheet1.BlockOccurrences.Count
    block = Sheet1.BlockOccurrences.Item(count)
    labels = block.BlockLabelOccurrences

    date_today = System.DateTime.Today.ToString("yyyy-MM-dd")
    comment = "MISE A JOUR."

    # Split comment in two lines
    comment1 = comment[:43]
    comment2 = comment[43:]

    # Add info to block revision
    labels[1].Value = comment1.upper()
    labels[2].Value = comment2.upper()
    labels[3].Value = user.upper()
    labels[4].Value = date_today
    labels[5].Value = current_revision


def confirmation(func):
    response = raw_input(
    """Would you like to delete graphic components (ID rev & Block revision) or add a Revision Block to this draft ?
    (Press y/[Y] to proceed.): 
    """)
    if response.lower() not in ["y"]:
        print("Process canceled")
        sys.exit()
    else:
        func()


def raw_input(message):
    Console.WriteLine(message)
    return Console.ReadLine()


def username():
    return System.Environment.UserName


def get_height(current_revision):
    revision = current_revision  # integer of revision e.g 0,1,2,...
    HEIGHT = 0.0065  # height of revision block
    if revision == 1:
        return 0.0655828
    if revision > 1:
        return 0.0655828 + ((revision - 1) * HEIGHT)
    else:
        raise ValueError


def main():
    try:
        application = SRI.Marshal.GetActiveObject("SolidEdge.Application")
        draft = application.ActiveDocument
        print("Document name: %s\n" % draft.Name)
        assert draft.Type == 2, ("This macro only works on draft")

        # Collect info for blocks
        current_revision = get_document_revision(draft)
        user = username()

        if not current_revision:
            remove_blocks(draft)
        else:
            insert_blocks(draft, current_revision, user)

    except AssertionError as ae:
        print(ae.args)

    except ValueError as ve:
        print(ve.args)

    except Exception as ex:
        print(ex.args)

    finally:
        raw_input("\nPress any key to exit...")
        sys.exit()


if __name__ == "__main__":
    print("%s\n--author: %s --version: %s --last-update : %s \n" %
          (__project__, __author__, __version__, __update__))
    confirmation(main)
