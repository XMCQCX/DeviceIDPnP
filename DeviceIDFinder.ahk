/*
Script:    DeviceIDFinder.ahk
Author:    XMCQCX
Date:      2022-09-24
Version:   1.0.0
*/

#NoEnv
#SingleInstance, Force
SendMode Input
SetWorkingDir, %A_ScriptDir%

MsgBox, 64, Find deviceID, Plug your device and press OK

;=============================================================================================

; List all devices connected
For Device in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_PnPEntity")
    ListConnectedDeviceIDs .= device.name ":" A_Tab Device.DeviceID "`n"

;=============================================================================================

; Remove duplicates from ListConnectedDeviceIDs
Loop, Parse, ListConnectedDeviceIDs, "`n"
{
    ListConnectedDeviceIDs := (A_Index=1 ? A_LoopField : ListConnectedDeviceIDs . (InStr("`n" ListConnectedDeviceIDs
    . "`n", "`n" A_LoopField "`n") ? "" : "`n" A_LoopField ) )
}

;=============================================================================================

; Add all devices connected in oConnectedDeviceIDs
oConnectedDeviceIDs := {}
Loop, Parse, ListConnectedDeviceIDs, "`n"
    oConnectedDeviceIDs.Push({"DeviceID":A_Loopfield})

;=============================================================================================

MsgBox, 64, Find deviceID, Unplug your device and press OK

;=============================================================================================

; List all devices connected
For Device in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_PnPEntity")
    strListConnectedDeviceIDs .= device.name ":" A_Tab Device.DeviceID "`n"

;=============================================================================================

; Remove duplicates from strListConnectedDeviceIDs
Loop, Parse, strListConnectedDeviceIDs, "`n"
{
    strListConnectedDeviceIDs := (A_Index=1 ? A_LoopField : strListConnectedDeviceIDs . (InStr("`n" strListConnectedDeviceIDs
    . "`n", "`n" A_LoopField "`n") ? "" : "`n" A_LoopField ) )
}

;=============================================================================================

; Find the device that was connected/disconnected
Loop, Parse, strListConnectedDeviceIDs, "`n"
{
    For Index, Element in oConnectedDeviceIDs  
    {
        If InStr(strListConnectedDeviceIDs, Element.DeviceID)
            oConnectedDeviceIDs.RemoveAt(Index)
    }
}

;=============================================================================================

; List the IDs of the device
For Index, Element in oConnectedDeviceIDs
    DeviceIDFound .= Element.DeviceID "`n"

;=============================================================================================

; Format the IDs of the device
strDeviceIDFound := ""
For each, line in StrSplit(DeviceIDFound, "`n")
{
    RegExMatch(line, "`nm)^(.*?)" A_TAB "(.*)$", OutputVar)
        strDeviceIDFound .= OutputVar1 "`n" OutputVar2 "`n`n"
}
strDeviceIDFound := RTrim(strDeviceIDFound, "`n`n")

If !strDeviceIDFound
    strDeviceIDFound := "No device found !"

;=============================================================================================

Gui, New
Gui, Add, Text,, DeviceID:
Gui, Add, Edit, vTextDeviceID ReadOnly, %strDeviceIDFound%
Gui, Add, Button, w175 vCopyToClipboard gCopyToClipboard, Copy to Clipboard
Gui, Add, Button, x+10 w175 gFindAnotherDeviceID, Find ID of another device
Gui, Add, Button, x+10 w175 gExit, Exit
GuiControl, Focus, CopyToClipboard
Gui, Show
return

;=============================================================================================

CopyToClipboard:
Clipboard := strDeviceIDFound
MsgBox, 64, Success, Device ID copied to Clipboard !
return

FindAnotherDeviceID:
Reload

Exit:
GuiClose:
Exitapp