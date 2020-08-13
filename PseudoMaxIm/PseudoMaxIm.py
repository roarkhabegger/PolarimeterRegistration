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

import FitsCreationModified

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
    def Expose(self,duration, light, filter):
        # MaxIm Object
        maxIm_obj = com.Dispatch("MaxIm.CCDCamera")
        maxIm_obj.linkEnabled = True
        self.ReadyForDownload = False

        # Polarimeter obj
        if filter<=0:
            pol_obj = com.Dispatch("PolControl")
            pol_obj.Simulation = True
            pol_obj.FindSerialPorts()

            pol_positions = [0, 1, 2, 3]
            pol_filters = ["V", "I"]
            pol_locations = [0,2]
            pol_outsLA = ["100","001"]
            # Expose on every position and filter
            for i in range(len(pol_filters)):
                pol_filter = pol_filters[i]
                pol_obj.SendCommand("L", pol_locations[i])
                #Actually wait for polarimeter to move (put time delay in Sim)
                isReady = False
                while not isReady:
                    state = pol_obj.ReadState()
                    #print(state)
                    if state[2:5] == pol_outsLA[i]:
                        isReady = True

                for pol_position in pol_positions:
                    pol_obj.SendCommand(pol_filter, pol_position)
                    time.sleep(2) # add artificial delay
                    maxIm_obj.Expose(duration, light, filter)
                    #Fake exposure for no MaxIm:
                    #print("Exposing!")
                    #time.sleep(duration)
                    #print("Done Exposing")
                    while maxIm_obj.ImageReady == False:
                        time.sleep(1)
                        print('waiting for exposure to be ready..')
                    maxIm_obj.saveImage('C:\\Users\\astro\\Desktop\\Work\\polarimeter\\22June2020\\images\\image_{}_{}.fits'.format(pol_filter, pol_position))

            # Home the polarimeter after exposing
            pol_obj.home()

            # Create a FITS cube and delete original images
            input = 'C:\\Users\\astro\\Desktop\\Work\\polarimeter\\22June2020\\images'
            output = 'C:\\Users\\astro\\Desktop\\Work\\polarimeter\\22June2020\\cubeImages'

            try:
                FitsCreationModified.SaveFitsFiles(input, output)
                FitsCreationModified.DeleteFitsFiles(input)
            except IOError:
                print("Missing files..")

        else:
            maxIm_obj.Expose(duration,light,filter)
            while maxIm_obj.ImageReady == False:
                time.sleep(1)
                print('waiting for exposure to be ready..')

        self.ReadyForDownload = True
        return True


if __name__ == "__main__":
    from comtypes.server.register import UseCommandLine
    UseCommandLine(PseudoMaxImImp)
