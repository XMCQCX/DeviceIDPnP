/*
Script:    DeviceIDFinder.ahk
Author:    XMCQCX
Date:      2023-03-01
Version:   2.0.0
Github:    https://github.com/XMCQCX/DeviceIDPnP
AHK forum: https://www.autohotkey.com/boards/viewtopic.php?f=83&t=114610

Credits: jNizM
This is a modified version of his script.
Example 2: Detect / Monitor Plug and Play device connections and removes
https://www.autohotkey.com/boards/viewtopic.php?f=83&t=105171
*/

#Requires AutoHotkey v2.0
#SingleInstance Force

;=============================================================================================

Main := Gui("+Resize +MinSize500x300", "DeviceIDFinder")
Main.SetFont("s10")
Main.Add("Text", "xm-5", "Connect or disconnect your devices to view their IDs. Use the right-click context menu to copy selected items.")
LV := Main.AddListView("w865 h490 r10 +BackgroundDEDEDE Grid", ["Event", "Time", "Device Name", "Device ID"])
for k, v in ["95", "70", "230", "465"]
	LV.ModifyCol(k, v)

LV.OnEvent("ContextMenu", ShowContextMenu)
Main.OnEvent("Size", GuiSize)
Main.OnEvent("Close", GuiClose)
Main.Show()

;=============================================================================================

WMI := ComObjGet("winmgmts:")
ComObjConnect(Sink := ComObject("WbemScripting.SWbemSink"), "SINK_")
command := "WITHIN 1 WHERE TargetInstance ISA 'Win32_PnPEntity'"
WMI.ExecNotificationQueryAsync(Sink, "SELECT * FROM __InstanceCreationEvent " . command)
WMI.ExecNotificationQueryAsync(Sink, "SELECT * FROM __InstanceDeletionEvent " . command)

;=============================================================================================

SINK_OnObjectReady(Obj, *)
{
	TI := Obj.TargetInstance
	Time := FormatTime(A_Now, "HH:mm:ss")
	switch Obj.Path_.Class
	{
		case "__InstanceCreationEvent": EventType := "Connected"
		case "__InstanceDeletionEvent": EventType := "Disconnected"
	}
	LV.Insert(1,, EventType, Time, TI.Name, TI.DeviceID)
}

;=============================================================================================

ShowContextMenu(LV, Item, IsRightClick, X, Y)
{
    MouseGetPos(,,, &mouseOverClassNN)
    if (Item = 0 || InStr(mouseOverClassNN, "SysHeader"))
        Return

    ContextMenu := Menu()
    ContextMenu.Add("Copy Selected Item", CopyToClipboard)
    ContextMenu.Add("Clear Listview", ClearListview)
    ContextMenu.Show(X, Y)

    CopyToClipboard(*)
    {
        rowNumber := 0
        Loop
        {
            rowNumber := LV.GetNext(rowNumber)
            if not rowNumber
                break
            deviceName := LV.GetText(rowNumber, 3)
            deviceID := LV.GetText(rowNumber, 4)
            strDev .= deviceName ": " deviceID "`n"
        }
        strDev := RTrim(strDev, "`n")
        A_Clipboard := ""
        A_Clipboard := strDev
        if !ClipWait(1)
            return
    }

    ClearListview(*) => LV.Delete()
}

;=============================================================================================

GuiSize(thisGui, MinMax, Width, Height)
{
	if (MinMax = -1)
		return
	LV.Move(,, Width - 20, Height  - 40)
	LV.Redraw()
}

;=============================================================================================

GuiClose(*)
{
	ComObjConnect(Sink)
	ExitApp
}
