/*
Script:    DeviceIDPnP.ahk
Author:    XMCQCX
Date:      2022-09-28
Version:   1.2.0
Links:

Change log:
2022-09-28 -> 1.2.0 - Fixed an issue when connecting/disconnecting multiple devices at the same time or in quick succession.
2022-09-27 -> 1.1.1 - Fixed an issue with some devices status not properly updating when connecting/disconnecting.
2022-09-24 -> 1.1.0 - Added run or close scripts/programs if the devices are connected/disconnected when the script start.
*/

#NoEnv
#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%
CoordMode, ToolTip

;=============================================================================================

oMyDevices := {}
oMyDevices.Push({"DeviceName":"USB Kingston DataTraveler 3.0", "DeviceID":"USB\VID_0951&PID_1666\E0D55EA573DCF450E97C104C"})
oMyDevices.Push({"DeviceName":"BLUETOOTH PLAYSTATION(R)3 Controller", "DeviceID":"BTHPS3BUS\{53F88889-1AAF-4353-A047-556B69EC6DA6}&DEV&VID_054C&PID_0268&04766E9094F3\9&320AC31D&0&0"})
oMyDevices.Push({"DeviceName":"HDMI Samsung TV", "DeviceID":"SWD\MMDEVAPI\{0.0.0.00000000}.{ED3C7A62-B05B-44C6-ACD8-BCAA1E894265}"})

;=============================================================================================

DevicesActions(ThisDeviceStatusHasChanged) {

    If (ThisDeviceStatusHasChanged = "USB Kingston DataTraveler 3.0 Connected")
        If !WinExist("ahk_exe Notepad.exe")
            Run, Notepad.exe

    If (ThisDeviceStatusHasChanged = "USB Kingston DataTraveler 3.0 Disconnected")
        If WinExist("ahk_exe Notepad.exe")
            Winclose, % "ahk_exe Notepad.exe"

    ;=============================================================================================

    If (ThisDeviceStatusHasChanged = "BLUETOOTH PLAYSTATION(R)3 Controller Connected")
        If !WinExist("ahk_exe wordpad.exe")
            Run, wordpad.exe

    If (ThisDeviceStatusHasChanged = "BLUETOOTH PLAYSTATION(R)3 Controller Disconnected")
        If WinExist("ahk_exe wordpad.exe")
            Winclose, % "ahk_exe wordpad.exe"

    ;=============================================================================================

    If (ThisDeviceStatusHasChanged = "HDMI Samsung TV Connected")
        If !WinExist("ahk_exe mspaint.exe")
            Run, mspaint.exe
    
    If (ThisDeviceStatusHasChanged = "HDMI Samsung TV Disconnected")
        If WinExist("ahk_exe mspaint.exe")
            Winclose, % "ahk_exe mspaint.exe"

    ;=============================================================================================
}

; Check devices connected
oDevicesConnected := {}
For Device in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_PnPEntity")
    oDevicesConnected.Push({"DeviceName":Device.Name, "DeviceID":Device.DeviceID, "DevicePNPClass":Device.PNPClass, "DeviceStatus":Device.Status})

;=============================================================================================

; Establish the status of the devices in oMyDevices
For Index, MyDevice in oMyDevices
{
    DeviceFound := ""
    For Index, DeviceConnected in oDevicesConnected
    {
        If (MyDevice.DeviceID = DeviceConnected.DeviceID)
        {
            If (DeviceConnected.DeviceStatus = "OK"), DeviceFound := "Yes"
                MyDevice.DeviceStatus := "Connected"

            If (DeviceConnected.DeviceStatus = "Unknown"), DeviceFound := "Yes"
                MyDevice.DeviceStatus := "Disconnected"
        }
    }
    If !DeviceFound
        MyDevice.DeviceStatus := "Disconnected"
}

;=============================================================================================

; Run or close scripts/programs if the devices are connected/disconnected when the script start.
Loop % oMyDevices.Count()
{
    DeviceStatustStartup := oMyDevices[A_Index].DeviceName A_Space oMyDevices[A_Index].DeviceStatus
    DevicesActions(DeviceStatustStartup)
    DeviceStatustStartup := StrReplace(DeviceStatustStartup, "Disconnected", "Not connected")
    strTooltip .= DeviceStatustStartup "`n"
}
If strTooltip
    strTooltip := RTrim(strTooltip, "`n")
        Tooltip, % strTooltip, 0, 0
            SetTimer, RemoveToolTipDeviceStatus, -6000

;=============================================================================================

OnMessage(0x219, "WM_DEVICECHANGE") 
WM_DEVICECHANGE(wParam, lParam, msg, hwnd)
{
    SetTimer, CheckDevicesStatus , -1250
}
Return

;=============================================================================================

CheckDevicesStatus:

    ;=============================================================================================
    
    ; Check devices connected
    oDevicesConnected.Delete(1, oDevicesConnected.Length())
    For Device in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_PnPEntity")
        oDevicesConnected.Push({"DeviceName":Device.Name, "DeviceID":Device.DeviceID, "DevicePNPClass":Device.PNPClass, "DeviceStatus":Device.Status})

    ;=============================================================================================

    ; Find which devices status has changed in oMyDevices
    oMyDevicesStatusHasChanged := []
    For Index, MyDevice in oMyDevices
    {
        DeviceFound := ""
        For Index, DeviceConnected in oDevicesConnected
        {
            If (MyDevice.DeviceID = DeviceConnected.DeviceID)
                If (DeviceConnected.DeviceStatus = "OK")
                    If (MyDevice.DeviceStatus = "Disconnected"), MyDevice.DeviceStatus := "Connected", DeviceFound := "Yes"
                            oMyDevicesStatusHasChanged.Push(MyDevice.DeviceName " Connected")
                
                If (DeviceConnected.DeviceStatus = "Unknown")
                    If (MyDevice.DeviceStatus = "Connected"), MyDevice.DeviceStatus := "Disconnected", DeviceFound := "Yes"
                            oMyDevicesStatusHasChanged.Push(MyDevice.DeviceName " Disconnected")
        }
        If !DeviceFound
            If (MyDevice.DeviceStatus = "Connected"), MyDevice.DeviceStatus := "Disconnected"
                    oMyDevicesStatusHasChanged.Push(MyDevice.DeviceName " Disconnected")
    }

    ;=============================================================================================

    ; If devices in oMyDevices status has changed go to DevicesActions()
    If (oMyDevicesStatusHasChanged)
    {
        strTooltip := ""
        Loop % oMyDevicesStatusHasChanged.Count()
        {
            DevicesActions(oMyDevicesStatusHasChanged[1])
            strTooltip .= oMyDevicesStatusHasChanged[1] "`n"
            oMyDevicesStatusHasChanged.RemoveAt(1)
        }
        If strTooltip
            strTooltip := RTrim(strTooltip, "`n")
                Tooltip, % strTooltip, 0, 0
                    SetTimer, RemoveToolTipDeviceStatus, -6000
    }
    
    ;=============================================================================================

return

;=============================================================================================

RemoveToolTipDeviceStatus:
ToolTip
return