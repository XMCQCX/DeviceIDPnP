/*
Script:    DeviceIDPnP.ahk
Author:    XMCQCX
Date:      2023-03-01
Version:   2.0.0
Github:    https://github.com/XMCQCX/DeviceIDPnP
AHK forum: https://www.autohotkey.com/boards/viewtopic.php?f=83&t=114610
*/

#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent
CoordMode "ToolTip", "Screen"
MyDevices := MyDevicesAdd()

;=============================================================================================

MyDevices.Add({DeviceName:"USB 3.0", DeviceID:"USB\VID_0951&PID_1666\E0D55EA573DCF450E97C104C"})

;=============================================================================================

DevicesActions(thisDeviceStatus) {

    if thisDeviceStatus = "USB 3.0 Connected"
        if !WinExist("ahk_exe Notepad.exe")
            Run "Notepad.exe"

    if thisDeviceStatus = "USB 3.0 Disconnected"
        if WinExist("ahk_exe Notepad.exe")
            Winclose "ahk_exe Notepad.exe"
}

;=============================================================================================

Class MyDevicesAdd {
    
    aMyDevices := []

	Add(oItem)
	{
        aDevIDs := [], devCount := 0

        if InStr(oItem.DeviceID, "|&|") {
            for _, devID in StrSplit(oItem.DeviceID, "|&|") {
                aDevIDs.push(devID := Trim(devID))
                oItem.DeviceCount := ++devCount
            }
            if !oItem.HasOwnProp("DevicesMatchMode")
                oItem.DevicesMatchMode := 1
        }
        else {
            aDevIDs.push(oItem.DeviceID := Trim(oItem.DeviceID))
            oItem.DeviceCount := 1
            oItem.DevicesMatchMode := 1
        }   
        
        if !oItem.HasOwnProp("ActionAtStartup")
            oItem.ActionAtStartup := "true"
        
        if !oItem.HasOwnProp("Tooltip")
            oItem.Tooltip := "true"

        oItem.DeviceID := aDevIDs
        this.aMyDevices.push(oItem)
        
        devExist := DevicesExistCheck(aDevIDs, oItem.DeviceCount, oItem.DevicesMatchMode)

        if devExist
            this.aMyDevices[this.aMyDevices.Length].DeviceStatus := "Connected"
        else
            this.aMyDevices[this.aMyDevices.Length].DeviceStatus := "Disconnected"
	}
}

;=============================================================================================

TooltipDevicesActions(Mydevices.aMyDevices)

TooltipDevicesActions(Array) {

    strTooltip := ""

    for _, item in Array
    {
        if item.Tooltip = "true"
            strTooltip .= item.DeviceName A_Space item.DeviceStatus "`n"
        
        if item.ActionAtStartup = "true"
            DevicesActions(item.DeviceName A_Space item.DeviceStatus)
    }

    If strTooltip {
        strTooltip := RTrim(strTooltip, "`n")
        Tooltip strTooltip, 0, 0
        SetTimer () => ToolTip(), -6000
    }
}

;=============================================================================================

DevicesExistCheck(aDevIDs, DeviceCount, DevicesMatchMode) {

    aDevList :=  [], devExistCount := 0
    
    for dev in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_PnPEntity")
        aDevList.Push({DeviceID: dev.DeviceID, DeviceStatus :dev.Status})

    for _, mydevID in aDevIDs
        for _, dev in aDevList
            if mydevID = dev.DeviceID
                if dev.DeviceStatus = "OK"
                    devExistCount++
    
    if DevicesMatchMode = 1
        if DeviceCount = devExistCount
            Return true
    
    if DevicesMatchMode = 2
        if devExistCount
            Return true
}

;=============================================================================================

OnMessage(0x219, WM_DEVICECHANGE)
WM_DEVICECHANGE(wParam, lParam, msg, hwnd) {
    SetTimer DevicesStatusCheck, -1250
}

DevicesStatusCheck() {

    aNewDevStatus := []
    for _, dev in MyDevices.aMyDevices
    {
        devExist := DevicesExistCheck(dev.DeviceID, dev.DeviceCount, dev.DevicesMatchMode)

        if (devExist && dev.DeviceStatus = "Disconnected") {
            dev.DeviceStatus := "Connected"
            aNewDevStatus.Push({DeviceName:dev.DeviceName, DeviceStatus:dev.DeviceStatus, Tooltip:dev.Tooltip, ActionAtStartup:"true"})
        }

        if (!devExist && dev.DeviceStatus = "Connected") {
            dev.DeviceStatus := "Disconnected"
            aNewDevStatus.Push({DeviceName:dev.DeviceName, DeviceStatus:dev.DeviceStatus, Tooltip:dev.Tooltip, ActionAtStartup:"true"})
        }
    }

    If aNewDevStatus.Length >= 1
        TooltipDevicesActions(aNewDevStatus)
}
