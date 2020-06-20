import win32com.client as com
myPS = com.Dispatch("PseudoMaxIm")
myPS.Expose(1.0,0,1)
