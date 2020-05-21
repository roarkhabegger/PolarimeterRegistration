import comtypes
import comtypes.server.localserver
import threading as th
import subprocess as sp
import serial as s
from serial.tools import list_ports as list
#import serial.tools.list_ports.ListPortInfo as Info
import time
#import PolControlForm as form
import io as io


#Execute following each time .idl file changes
from comtypes.client import GetModule
GetModule("PseudoMaxIm.tlb")
startMarker = '['
endMarker = ']'



from comtypes.gen.PseudoMaxImTypeLib import PseudoMaxIm
class PseudoMaxImImp(PseudoMaxIm):
    # Registry entries
    _reg_threading_ = "Free"
    _reg_progid_ = "PseudoMaxIm"
    _reg_novers_progid_ = "PseudoMaxIm"
    _reg_desc_ = "Simple Com Server"
    _regcls_ = comtypes.server.localserver.REGCLS_MULTIPLEUSE
    _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER
    #_bool_form_ = False
    #_my_directory_ = r"C:\Users\Roark Habegger\OneDrive\Documents\UNC\Polarimeter Research\PolControl"
    #COM Properties
    ReadyForDownload = False

    #COM Methods
    def Expose(self,duration,light,filter):
        return True






if __name__ == "__main__":
    from comtypes.server.register import UseCommandLine
    UseCommandLine(PseudoMaxImImp)
