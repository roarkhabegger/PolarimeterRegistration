import win32com.client as com
myPS = com.Dispatch("PseudoMaxIm")
print("ReadyForDownload Default Value = ",myPS.ReadyForDownload)
print("Let's do an exposure!")
myVal = myPS.Expose(1.0,0,-1)
print("Done Exposing!")
print("Post-Expose ReadyForDownload Value = ",myPS.ReadyForDownload)
