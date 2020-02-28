#script.py
import win32com.client as com
import time
import comtypes
import comtypes.server.localserver
import threading as th
import subprocess as sp
import serial as s
from serial.tools import list_ports as list
#import serial.tools.list_ports.ListPortInfo as Info
import time
#import PolControlForm as pcF
#pcF.MakeForm()
obj = com.Dispatch("PolControl")
print(obj)
time.sleep(2)
#test = 0
#print(obj.FindSerialPorts())
#print(obj.ReadState())
#print(obj.ReadControl())
#print(obj.ReadControl())
#print(obj.SendCommand(b"V",0))
#print(obj.ReadControl())
#test = obj.TestConnect()
#print(test)#
#time.sleep(20)

#FindSerialPorts()
print("script is done")
