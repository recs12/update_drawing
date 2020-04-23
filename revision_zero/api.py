
"""
Api solidedge
=======================

"""

import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System.Runtime.InteropServices")

import sys
import SolidEdgeFramework as SEFramework
import SolidEdgePart as SEPart
import SolidEdgeConstants as SEConstants
import System.Runtime.InteropServices as SRI

class Api():

    def __init__(self):
        # Connect to a running instance of Solid Edge
        self.api = SRI.Marshal.GetActiveObject("SolidEdge.Application")

    def check_valid_version(self, *valid_version):
        #validate solidedge version - 'Solid Edge ST7'
        print("version solidedge: %s" %self.api.Value)
        assert self.api.Value in valid_version, "Unvalid version of solidedge"

    def active_document(self):
        return self.api.ActiveDocument

    @property
    def name(cls, part):
        return cls.part.Name

    def delete(cls, part):
        return cls.part.Delete()


