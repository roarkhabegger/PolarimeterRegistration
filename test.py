#script.py
import win32com.client as com
import time

obj = com.Dispatch("PolControl")
print(obj)
time.sleep(2)

print("script is done")
