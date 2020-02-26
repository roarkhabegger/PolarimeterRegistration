import os
import subprocess as sp

def main():
    pyStr = input("Type your python3 Command: ")
    print("Checking important python libraries")
    instArr = []
    try:
        import serial.tools
    except Exception as e:
        #instArr.append("python -m pip install serial")
        instArr.append(pyStr+" -m pip install pySerial")

    try:
        import comtypes.server
    except Exception as e:
        instArr.append(pyStr+" -m pip install comtypes")

    try:
        import win32com.client as com
    except Exception as e:
        instArr.append(pyStr+" -m pip install pywin32")

    print(instArr)
    if len(instArr)!=0:
        for entry in instArr:
            File0 = open("install.bat","w+")
            File0.truncate()
            File0.writelines("%s \n" % entry)#+" \n",cmd3])
            File0.close()
            inst = sp.call("install.bat")

    print("Starting Registration")
    my_cwd = os.getcwd()
    my_env = os.environ
    filename = r"PolControl.idl"
    vsDir = r"C:\Program Files (x86)\Microsoft Visual Studio"

    ###search through Vis Studio for vcvars32.bat
    tree = os.scandir(vsDir)
    vers = 0
    for entry in tree :
        if entry.is_dir():
            dir = entry.name
            try:
                newvers = int(dir)
                if newvers>vers:
                    vers = newvers
            except:
                continue

    vsDir += "\\"+str(vers)
    tree = os.scandir(vsDir)
    for entry in tree :
        if entry.is_dir():
            nom = entry.name
    vsDir += "\\"+nom+r"\VC\Auxiliary\Build"
    batDir = vsDir
    batFile = r"vcvars32.bat"

    try:
        open(vsDir+'\\'+batFile)
    except Exception as e:
        print(e)
        exit()

    #Now we have the required registration material
    #Script to compile midl file
    envArr = []
    envArr2 = []
    cmdArr = []

    envArr.append('cd "'+batDir+ '"')
    envArr.append("cd")
    envArr.append(batFile+' &&midl /win32 "'+my_cwd+r'"\Polcontrol.idl')
    envArr2.append('cd "'+batDir+ '"')
    envArr2.append("cd")
    envArr2.append(batFile+' && "'+my_cwd+r'"\Register.bat') #batFile#+
    cmdArr.append('cd "'+my_cwd+'"')
    cmdArr.append(pyStr+' PolControl.py -regserver')

    File1 = open("SetEnv1.bat","w+")
    File1.truncate()
    File1.writelines("%s \n" % c for c in envArr)#+" \n",cmd3])
    File1.close()

    File2 = open("SetEnv2.bat","w+")
    File2.truncate()
    File2.writelines("%s \n" % c for c in envArr2)#+" \n",cmd3])
    File2.close()

    File3 = open("Register.bat","w+")
    File3.truncate()
    File3.writelines("%s \n" % c for c in cmdArr)#+" \n",cmd3])
    File3.close()
    inst = sp.call("SetEnv1.bat")
    print("Defining server")
    inst = sp.call("SetEnv2.bat")
    #inst = sp.call("Register.bat")
main()
