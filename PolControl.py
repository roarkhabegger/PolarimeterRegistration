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
GetModule("PolControl.tlb")
startMarker = '['
endMarker = ']'

def sendToArduino(ser,sendStr):
    #sendStr = sendStr + '\n'
    waitForArduino(ser)
    ser.write(sendStr)
    print(waitForArduino(ser))
    return

def recvFromArduino(ser):
    #global startMarker, endMarker
    ck = ""
    x = "1" # any value that is not an end- or startMarker
    byteCount = -1 # to allow for the fact that the last increment will be one too many
    # wait for the start character
    while x != startMarker:
        x = str(ser.read())[2:-1]

        # save data until the end marker is found
    while x != endMarker:
        if x != startMarker:
            ck = ck + x
            byteCount += 1
        x = str(ser.read())[2:-1]

    return startMarker+ck+endMarker

def waitForArduino(ser):
    # wait until the Arduino sends 'Arduino Ready' - allows time for Arduino reset
    # it also ensures that any bytes left over from a previous message are discarded
    #global startMarker, endMarker
    msg = ""
    while msg.find(startMarker) == -1 & msg.find(endMarker)==-1:
        while ser.inWaiting() == 0:
            pass
        msg = recvFromArduino(ser)

    return msg

from comtypes.gen.PolControlTypeLib import PolControl
class PolControlImp(PolControl):
    # Registry entries
    _reg_threading_ = "Free"
    _reg_progid_ = "PolControl"
    _reg_novers_progid_ = "PolControl"
    _reg_desc_ = "Simple Com Server"
    _regcls_ = comtypes.server.localserver.REGCLS_MULTIPLEUSE
    _reg_clsctx_ = comtypes.CLSCTX_LOCAL_SERVER
    #_bool_form_ = False
    #_my_directory_ = r"C:\Users\Roark Habegger\OneDrive\Documents\UNC\Polarimeter Research\PolControl"
    Connection = False
    ControlPort = -1
    StatePort = -1

    #COM Methods
    def FindSerialPorts(self):
        result = 0
        ActiveComPorts = list.comports()
        myPorts = []
        for i in range(len(ActiveComPorts)):
            port = ActiveComPorts[i]
            #print(port)
            #time.sleep(2)
            try:
                arduino = s.Serial(port[0],57600,timeout=0.1)
                boolVal = arduino.is_open
                if boolVal==True:
                    desc = port[1]
                    if desc == 'Arduino Uno '+"("+port[0]+")":
                        myPorts.append(str(port[0]))
                arduino.close()
            except Exception as e:
                print(e)


        #Test COM Ports for being the State or Control Machines
        for i in range(len(myPorts)):
            ard1 = s.Serial(myPorts[i],57600,timeout=0.1)
            str1 = waitForArduino(ard1)
            if str1.find("R")!=-1 or str1.find("L")!=-1:
                #print("Found State")
                self.StatePort = int(myPorts[i][-1])
            if str1.find("+")!=-1:
                #print("Found Control")
                self.ControlPort = int(myPorts[i][-1])
            #str1 = ""
            ard1.close()
        #print(self.StatePort)
        #print(self.ControlPort)
        #time.sleep(2)
        if self.ControlPort!=-1 & self.StatePort!=-1:
            result = 1
        return result

    def SendCommand(self, motor, position):
        #cmdStr = motor+'='
        #cmdStr += position
        result = 0
        if self.TestConnect()==1:
            try:
                ardCont = s.Serial("COM"+str(self.ControlPort),57600)
                sendToArduino(ardCont,b'V=1')
                ardCont.close()
                result=1
            except Exception as e:
                strOut = "Function SendCommand exited with the following error: \n"
                strOut += str(e)
                time.sleep(4)
        return result

    def Home(self):
        cmdStr = b'H'
        result = 0
        if self.TestConnect()==1:
            try:
                ardCont = s.Serial("COM"+str(self.ControlPort),57600)
                sendToArduino(ardCont,cmdStr)
                print('Message Sent')
                ardCont.close()
                result=1
            except Exception as e:
                strOut = "Function SendCommand exited with the following error: \n"
                strOut += str(e)
        return result

    def TestConnect(self):
        result = 0
        if self.ControlPort==-1:
            print("NO CONTROL PORT DATA")
            result = -1
        if self.StatePort==-1:
            print("NO State PORT DATA")
            result = -1
        if result==-1:
            result = 0
        else:
            try:
                ard1 = s.Serial("COM"+str(self.StatePort),57600,timeout=0.1)
                ard2 = s.Serial("COM"+str(self.ControlPort),57600,timeout=0.1)
                if ard1.is_open & ard2.is_open:
                    self.Connection=True
                    result = 1
                ard1.close()
                ard2.close()

                #self.Connection = arduino.is_open
                #print("Connection Status = "+str(self.Connection))
            except Exception as e:
                print("Function TestConnect exited with the following error: \n")
                print(e)

        return result

    def ReadState(self):
        strOut='NULL'
        if self.TestConnect()==1:
            try:
                ardState = s.Serial("COM"+str(self.StatePort),57600,timeout=0.1)
                str1 = waitForArduino(ardState)
                strOut = str1
                ardState.close()
            except Exception as e:
                strOut = "Function ReadState exited with the following error: \n"
                strOut += str(e)
        return strOut

    def ReadControl(self):
        strOut='NULL'
        if self.TestConnect()==1:
            try:
                ardCont = s.Serial("COM"+str(self.ControlPort),57600,timeout=0.1)
                str1 = waitForArduino(ardCont)
                strOut = str1
                ardCont.close()
            except Exception as e:
                strOut = "Function ReadControl exited with the following error: \n"
                strOut += str(e)
        return strOut





if __name__ == "__main__":
    from comtypes.server.register import UseCommandLine
    UseCommandLine(PolControlImp)
