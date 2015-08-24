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
colItems = objSWbemServices.ExecQuery( "SELECT * FROM Win32_LogicalDisk" )

if colItems.Count == 1:
    print( "1 instance:" )
else:
    print( str( colItems.Count ) + " instances:" )
print

for objItem in colItems:
    if objItem.Access == None:
        print( "Access                       : " )
    else:
        print( "Access                       : " + str( objItem.Access ) )
    if objItem.Availability == None:
        print( "Availability                 : " )
    else:
        print( "Availability                 : " + str( objItem.Availability ) )
    if objItem.BlockSize == None:
        print( "BlockSize                    : " )
    else:
        print( "BlockSize                    : " + str( objItem.BlockSize ) )
    if objItem.Caption == None:
        print( "Caption                      : " )
    else:
        print( "Caption                      : " + str( objItem.Caption ) )
    if objItem.Compressed == None:
        print( "Compressed                   : " )
    else:
        print( "Compressed                   : " + str( objItem.Compressed ) )
    if objItem.ConfigManagerErrorCode == None:
        print( "ConfigManagerErrorCode       : " )
    else:
        print( "ConfigManagerErrorCode       : " + str( objItem.ConfigManagerErrorCode ) )
    if objItem.ConfigManagerUserConfig == None:
        print( "ConfigManagerUserConfig      : " )
    else:
        print( "ConfigManagerUserConfig      : " + str( objItem.ConfigManagerUserConfig ) )
    if objItem.CreationClassName == None:
        print( "CreationClassName            : " )
    else:
        print( "CreationClassName            : " + str( objItem.CreationClassName ) )
    if objItem.Description == None:
        print( "Description                  : " )
    else:
        print( "Description                  : " + str( objItem.Description ) )
    if objItem.DeviceID == None:
        print( "DeviceID                     : " )
    else:
        print( "DeviceID                     : " + str( objItem.DeviceID ) )
    if objItem.DriveType == None:
        print( "DriveType                    : " )
    else:
        print( "DriveType                    : " + str( objItem.DriveType ) )
    if objItem.ErrorCleared == None:
        print( "ErrorCleared                 : " )
    else:
        print( "ErrorCleared                 : " + str( objItem.ErrorCleared ) )
    if objItem.ErrorDescription == None:
        print( "ErrorDescription             : " )
    else:
        print( "ErrorDescription             : " + str( objItem.ErrorDescription ) )
    if objItem.ErrorMethodology == None:
        print( "ErrorMethodology             : " )
    else:
        print( "ErrorMethodology             : " + str( objItem.ErrorMethodology ) )
    if objItem.FileSystem == None:
        print( "FileSystem                   : " )
    else:
        print( "FileSystem                   : " + str( objItem.FileSystem ) )
    if objItem.FreeSpace == None:
        print( "FreeSpace                    : " )
    else:
        print( "FreeSpace                    : " + str( objItem.FreeSpace ) )
    if objItem.InstallDate == None:
        print( "InstallDate                  : " )
    else:
        print( "InstallDate                  : " + str( objItem.InstallDate ) )
    if objItem.LastErrorCode == None:
        print( "LastErrorCode                : " )
    else:
        print( "LastErrorCode                : " + str( objItem.LastErrorCode ) )
    if objItem.MaximumComponentLength == None:
        print( "MaximumComponentLength       : " )
    else:
        print( "MaximumComponentLength       : " + str( objItem.MaximumComponentLength ) )
    if objItem.MediaType == None:
        print( "MediaType                    : " )
    else:
        print( "MediaType                    : " + str( objItem.MediaType ) )
    if objItem.Name == None:
        print( "Name                         : " )
    else:
        print( "Name                         : " + str( objItem.Name ) )
    if objItem.NumberOfBlocks == None:
        print( "NumberOfBlocks               : " )
    else:
        print( "NumberOfBlocks               : " + str( objItem.NumberOfBlocks ) )
    if objItem.PNPDeviceID == None:
        print( "PNPDeviceID                  : " )
    else:
        print( "PNPDeviceID                  : " + str( objItem.PNPDeviceID ) )
    strList = " "
    try:
        for objElem in objItem.PowerManagementCapabilities :
            strList = strList + str( objElem ) + ","
    except:
        strList = strList + 'null'
    print( "PowerManagementCapabilities  :" + strList )
    if objItem.PowerManagementSupported == None:
        print( "PowerManagementSupported     : " )
    else:
        print( "PowerManagementSupported     : " + str( objItem.PowerManagementSupported ) )
    if objItem.ProviderName == None:
        print( "ProviderName                 : " )
    else:
        print( "ProviderName                 : " + str( objItem.ProviderName ) )
    if objItem.Purpose == None:
        print( "Purpose                      : " )
    else:
        print( "Purpose                      : " + str( objItem.Purpose ) )
    if objItem.QuotasDisabled == None:
        print( "QuotasDisabled               : " )
    else:
        print( "QuotasDisabled               : " + str( objItem.QuotasDisabled ) )
    if objItem.QuotasIncomplete == None:
        print( "QuotasIncomplete             : " )
    else:
        print( "QuotasIncomplete             : " + str( objItem.QuotasIncomplete ) )
    if objItem.QuotasRebuilding == None:
        print( "QuotasRebuilding             : " )
    else:
        print( "QuotasRebuilding             : " + str( objItem.QuotasRebuilding ) )
    if objItem.Size == None:
        print( "Size                         : " )
    else:
        print( "Size                         : " + str( objItem.Size ) )
    if objItem.Status == None:
        print( "Status                       : " )
    else:
        print( "Status                       : " + str( objItem.Status ) )
    if objItem.StatusInfo == None:
        print( "StatusInfo                   : " )
    else:
        print( "StatusInfo                   : " + str( objItem.StatusInfo ) )
    if objItem.SupportsDiskQuotas == None:
        print( "SupportsDiskQuotas           : " )
    else:
        print( "SupportsDiskQuotas           : " + str( objItem.SupportsDiskQuotas ) )
    if objItem.SupportsFileBasedCompression == None:
        print( "SupportsFileBasedCompression : " )
    else:
        print( "SupportsFileBasedCompression : " + str( objItem.SupportsFileBasedCompression ) )
    if objItem.SystemCreationClassName == None:
        print( "SystemCreationClassName      : " )
    else:
        print( "SystemCreationClassName      : " + str( objItem.SystemCreationClassName ) )
    if objItem.SystemName == None:
        print( "SystemName                   : " )
    else:
        print( "SystemName                   : " + str( objItem.SystemName ) )
    if objItem.VolumeDirty == None:
        print( "VolumeDirty                  : " )
    else:
        print( "VolumeDirty                  : " + str( objItem.VolumeDirty ) )
    if objItem.VolumeName == None:
        print( "VolumeName                   : " )
    else:
        print( "VolumeName                   : " + str( objItem.VolumeName ) )
    if objItem.VolumeSerialNumber == None:
        print( "VolumeSerialNumber           : " )
    else:
        print( "VolumeSerialNumber           : " + str( objItem.VolumeSerialNumber ) )
    print