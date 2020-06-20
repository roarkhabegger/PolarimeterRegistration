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

import win32com.client as com
import time

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

    # MaxIm Object
    maxIm_obj = com.Dispatch("MaxIm.CCDCamera")
    maxIm_obj.linkEnabled = True

    # Polarimeter obj
    pol_obj = com.Dispatch("PolControl")
    pol_obj.Simulation = True
    pol_obj.FindSerialPorts()

    #COM Methods
    def Expose(self, duration, light, filter):
        pol_positions = [0, 1, 2, 3]
        pol_filters = ["I", "V"]
        pol_locations = [0,2]

        # Expose on every position and filter
        for i in range(len(pol_filters):
            pol_filter = pol_filters[i]
            self.pol_obj.SendCommand("L",pol_locations[i])
            for pol_position in pol_positions:
                self.pol_obj.SendCommand(pol_filter, pol_position)
                time.sleep(2) # add artificial delay
                self.maxIm_obj.Expose(duration, light, filter)
                while self.maxIm_obj.ImageReady == False:
                    time.sleep(1)
                self.maxIm_obj.saveImage('C:\\Users\\astro\\Desktop\\Work\\polarimeter\\9June2020\\Images\\image_{}_{}.fits'.format(pol_filter, pol_position))

        # Home the polarimeter after exposing
        pol_obj.home()

        return True

    #Expose(maxIm_obj, pol_obj, 3.0, 0, 1)


if __name__ == "__main__":
    from comtypes.server.register import UseCommandLine
    UseCommandLine(PseudoMaxImImp)
