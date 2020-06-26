# -*- coding: utf-8 -*-

""" Check the revision of the document draft then either
    delete graphic blocks components <ID rev> & <Bloc revision>
    or add <ID rev> & <Bloc revision> if revision is above 00.
"""

import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System")
clr.AddReference("System.Runtime.InteropServices")

import sys
import System
import System.Runtime.InteropServices as SRI
from System import Console

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
        assert session.Value in [
            "Solid Edge ST7",
            "Solid Edge 2019",
        ], "Unvalid version of solidedge"
        draft = session.ActiveDocument
        print("part: %s\n" % draft.Name)
        assert draft.Name.lower().endswith(".dft"), (
            "This macro only works on .psm not %s" % draft.Name[-4:]
        )
        # Collect info for blocks
        current_revision = get_document_revision(draft)
        # current_revision = "02"
        print("Document Revision: %s" % current_revision)
        comment = raw_input("Add description:\>")
        user = username()
        print("User: %s" % user)
        date_today = System.DateTime.Today.ToString("yyyy-MM-dd")
        print("Date: %s" % date_today)

        if current_revision == "00":
            remove_blocks(draft)
        else:
            insert_blocks(draft, current_revision, user, date_today, comment)

    except AssertionError as ae:
        print(ae.args)

    except ValueError as ve:
        print(ve.args)

    except Exception as ex:
        print(ex.args)

    finally:
        raw_input("\nPress any key to exit...")
        sys.exit()


def get_document_revision(draft):
    """Revision of the draft
    """
    return draft.Properties.Item["ProjectInformation"]["Revision"].Value


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


def insert_blocks(draft, current_revision, user, date, comment):

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
    block = Sheet1.BlockOccurrences.Item(Sheet1.BlockOccurrences.Count)
    labels = block.BlockLabelOccurrences
    labels[1].Value = current_revision[-1]

    # Revision block
    blocks.AddBlockByFile(block_revision)
    Sheet1.BlockOccurrences.Add("Bloc revision - ENGLISH", 0.309499 , Y)
    block = Sheet1.BlockOccurrences.Item(Sheet1.BlockOccurrences.Count)
    labels = block.BlockLabelOccurrences

    # Split comment in two lines
    comment1 = comment[:43]
    comment2 = comment[43:]

    # Add info to block revision
    labels[1].Value = comment1.upper()
    labels[2].Value = comment2.upper()
    labels[3].Value = user.upper()
    labels[4].Value = date
    labels[5].Value = current_revision[-1]


def confirmation(func):
    response = raw_input(
        """Delete graphic components ID rev and Bloc revision,\n(Press y/[Y] to proceed.)"""
    )
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
    revision = int(current_revision[-1])    # integer of revision e.g 0,1,2,...
    HEIGHT = 0.0065                         # height of revision block
    if revision == 1:
        return 0.0655828
    if revision > 1:
        return 0.0655828 + ((revision-1) * HEIGHT)
    else:
        raise ValueError

if __name__ == "__main__":
    confirmation(revision)


