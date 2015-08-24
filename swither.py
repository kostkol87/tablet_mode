# -*- coding: utf-8 -*-
__author__ = 'Konstantin'

import sys
import win32com.client

try:
    strComputer = sys.argv[1]
except IndexError:
    strComputer = "."

objWMIService = win32com.client.Dispatch( "WbemScripting.SWbemLocator" )
objSWbemServices = objWMIService.ConnectServer( strComputer, "root/CIMV2" )
colItems = objSWbemServices.ExecQuery( "SELECT * FROM Win32_Keyboard" )

if colItems.Count == 1:
    print( "1 instance:" )
else:
    print( str( colItems.Count ) + " instances:" )
print

for objItem in colItems:
    # print(objItem.DeviceID)
    print(objItem.DeviceID)


2 instances:
USB\VID_0B05&PID_17E0&MI_00\6&1D8779C&0&0000
HID\INTCFD9&COL01\3&18C083FF&0&0000
>>> ================================ RESTART ================================
>>> 
1 instance:
HID\INTCFD9&COL01\3&18C083FF&0&0000
>>> 
