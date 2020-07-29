import win32com.client as com
myPS = com.Dispatch("PseudoMaxIm")
print("ReadyForDownload Default Value = ",myPS.ReadyForDownload)
print("Let's do an exposure!")
myVal = myPS.Expose(1.0,0,1)
print("Done Exposing!")
print("Pose-Expose ReadyForDownload Value = ",myPS.ReadyForDownload)
