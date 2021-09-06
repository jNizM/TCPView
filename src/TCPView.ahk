; SCRIPT DIRECTIVES =============================================================================================================

#Requires AutoHotkey v2.0-beta.1
#DllLoad "iphlpapi.dll"
#DllLoad "shell32.dll"
#DllLoad "ws2_32.dll"


; GLOBALS =======================================================================================================================

app := Map("name", "TCPView", "version", "0.3.1", "release", "2021-09-06", "author", "jNizM", "licence", "MIT")

LV_Header  := ["Process Name", "Process ID", "Protocol", "State", "Local Address", "Local Port", "Remote Address", "Remote Port", "Create Time", "Module Name"]
LV_Options := ["150 Text Left", "100 Integer Right", "80 Text Center", "80 Text Left", "150 Integer Left", "90 Integer Right", "150 Integer Left", "90 Integer Right", "140 Text Right", "180 Text Left"]
SB_Info    := [" Endpoints:", "Established:", "Listening:", "Time Wait:", "Close Wait:", "Update:", "States: (All)"]
SortCol    := 0


; GUI ===========================================================================================================================

OnMessage 0x0135, WM_CTLCOLORBTN
hhr1 := DllCall("gdi32\CreateBitmap", "int", 1, "int", 2, "int", 0x1, "int", 32, "int64*", 0x7f5a5a5a7fa5a5a5, "ptr")
hhr2 := DllCall("gdi32\CreateBitmap", "int", 1, "int", 2, "int", 0x1, "int", 32, "int64*", 0x7fcfcfcf7ffcfcfc, "ptr")

Main   := Gui("+Resize", app["name"])
Main.MarginX := 0
Main.MarginY := 0
Main.BackColor := "FFFFFF"
Main.SetFont("s10", "Segoe UI")

PIC1 := Main.AddPicture("xm ym w1250 h1 BackgroundTrans", "HBITMAP:*" hhr1)
CB1  := Main.AddCheckBox("xm+5 y+4 w80 h27 0x1000 Checked", "TCP v4")
CB2  := Main.AddCheckBox("x+4 yp w80 h27 0x1000", "TCP v6")
CB3  := Main.AddCheckBox("x+4 yp w80 h27 0x1000", "UDP v4")
CB4  := Main.AddCheckBox("x+4 yp w80 h27 0x1000", "UDP v6")
CB5  := Main.AddCheckBox("x+4 yp w80 h27 0x1000 Checked", "Pause")
CB5.OnEvent("Click", CB_Click)
DDL1 := Main.AddDropDownList("x+5 yp+1 w100 Choose2", ["2 Seconds", "5 Seconds", "10 Seconds"])
DDL1.OnEvent("Change", DDL1_Change)
Main.SetFont("s9", "Segoe UI")

PIC2 := Main.AddPicture("xm y+4 w1250 h2 BackgroundTrans", "HBITMAP:*" hhr2)
LV := Main.AddListView("w1250 r30 xm-1 y+0 -LV0x10 LV0x10000 -E0x0200", LV_Header)
loop LV_Options.Length
	LV.ModifyCol(A_Index, LV_Options[A_Index])
ImageListID1 := IL_Create(10)
ImageListID2 := IL_Create(10, 10, true)
LV.SetImageList(ImageListID1)
LV.SetImageList(ImageListID2)
LV.OnEvent("ContextMenu", LV_ContextMenu)

SB := Main.AddStatusBar("")
SB.SetParts(120, 120, 120, 120, 120, 120)
loop SB_Info.Length
	SB.SetText(SB_Info[A_Index], A_Index)

Main.OnEvent("Size", Gui_Size)
Main.OnEvent("Close", Gui_Close)
Main.Show()
SetExplorerTheme(LV.Hwnd)
HideFocusBorder(Main.Hwnd)

SetTimer NetStat, -1000


; WINDOW EVENTS =================================================================================================================

Gui_Size(thisGui, MinMax, Width, Height)
{
	if (MinMax = -1)
		return
	PIC1.Move(,, Width)
	PIC2.Move(,, Width)
	LV.Move(,, Width + 3, Height - 61)
}


Gui_Close(thisGui)
{
	global hBrush
	if (hBrush)
		DllCall("gdi32\DeleteObject", "ptr", hBrush)
	ExitApp
}


; CONTROL EVENTS ================================================================================================================

CB_Click(*)
{
	if (CB5.Value)
	{
		CB5.Text := "Pause"
		Interval := (DDL1.Value = 1) ? 2000 : (DDL1.Value = 2) ? 5000 : (DDL1.Value = 3) ? 10000 : 5000
		SB.SetText("Update: " StrLower(SubStr(DDL1.Text, 1, -4)), 6)
		SetTimer NetStat, Interval
	}
	else
	{
		CB5.Text := "Resume"
		SB.SetText("Paused", 6)
		SetTimer NetStat, 0
	}
}


DDL1_Change(*)
{
	Interval := (DDL1.Value = 1) ? 2000 : (DDL1.Value = 2) ? 5000 : (DDL1.Value = 3) ? 10000 : 5000
	SB.SetText("Update: " StrLower(SubStr(DDL1.Text, 1, -4)), 6)
	SetTimer NetStat, Interval
}


LV_ContextMenu(LV, Item, IsRightClick, X, Y)
{
	ContextMenu := Menu()
	ContextMenu.Add("Select All", Select)
	ContextMenu.SetIcon("Select All", "imageres.dll", 249)
	ContextMenu.Add("Export", Export)
	ContextMenu.SetIcon("Export", "shell32.dll", 259)
	ContextMenu.Add("Copy", Copy)
	ContextMenu.SetIcon("Copy", "shell32.dll", 135)
	ContextMenu.Show(X, Y)


	Select(*)
	{
		LV.Modify(0, "Select")
	}


	Export(*)
	{
		ExportFile := A_Desktop "\tcpview.csv"
		if (FileExist(ExportFile))
			FileDelete(ExportFile)

		loop LV.GetCount() + 1
		{
			i := A_Index - 1, Line := ""
			loop LV.GetCount("Column")
			{
				RetrievedText := LV.GetText(i, A_Index)
				Line .= RetrievedText ";"
			}
			FileAppend(SubStr(Line, 1, -1) "`n", ExportFile, "RAW")
		}
		Main.Opt("+OwnDialogs")
		MsgBox("CSV-Export is done (Desktop)", "CSV-Export", "T3")
	}


	Copy(*)
	{
		LV_List := ListViewGetContent("Selected", LV)
		A_Clipboard := LV_List
	}
}


; FUNCTIONS =====================================================================================================================

HideFocusBorder(wParam, lParam := "", Msg := "", hWnd := "")
{
	static Affected         := Map()
	static WM_UPDATEUISTATE := 0x0128
	static UIS_SET          := 1
	static UISF_HIDEFOCUS   := 0x1
	static SET_HIDEFOCUS    := UIS_SET << 16 | UISF_HIDEFOCUS
	static init             := OnMessage(WM_UPDATEUISTATE, HideFocusBorder)

	if (Msg = WM_UPDATEUISTATE) {
		if (wParam = SET_HIDEFOCUS)
			Affected[hWnd] := true
		else if (Affected.Has(hWnd))
			PostMessage WM_UPDATEUISTATE, SET_HIDEFOCUS, 0,, "ahk_id " hWnd
	}
	else if (DllCall("user32\IsWindow", "ptr", wParam, "uint"))
		PostMessage WM_UPDATEUISTATE, SET_HIDEFOCUS, 0,, "ahk_id " wParam
}


WM_CTLCOLORBTN(*)
{
	global hBrush
	return hBrush := DllCall("gdi32\CreateSolidBrush", "uint", 0xFFFFFF, "uptr")
}


SetExplorerTheme(handle)
{
	if (DllCall("kernel32\GetVersion", "uchar") > 5) {
		VarSetStrCapacity(&ClassName, 1024)
		if (DllCall("user32\GetClassName", "ptr", handle, "str", ClassName, "int", 512, "int")) {
			if (ClassName = "SysListView32") || (ClassName = "SysTreeView32")
				return !DllCall("uxtheme\SetWindowTheme", "ptr", handle, "str", "Explorer", "ptr", 0)
		}
	}
	return false
}


Process32()
{
	static PROCESS_QUERY_INFORMATION := 0x00000400
	static TH32CS_SNAPPROCESS        := 0x00000002

	if (hSnapshot := DllCall("kernel32\CreateToolhelp32Snapshot", "uint", TH32CS_SNAPPROCESS, "uint", 0, "ptr"))
	{
		TABLE := Map()
		PROCESSENTRY32W := Buffer(A_PtrSize = 8 ? 568 : 556, 0)
		NumPut("uint", PROCESSENTRY32W.Size, PROCESSENTRY32W, 0)
		if (DllCall("kernel32\Process32FirstW", "ptr", hSnapshot, "ptr", PROCESSENTRY32W))
		{
			while (DllCall("kernel32\Process32NextW", "ptr", hSnapshot, "ptr", PROCESSENTRY32W))
			{
				ROW := Map(), ProcessID := 0, hIcon := 0, IconNumber := 0
				ROW["ProcessID"] := ProcessID := NumGet(PROCESSENTRY32W, 8, "uint")
				ROW["ExeFile"]                := StrGet(PROCESSENTRY32W.Ptr + (A_PtrSize = 8 ? 44 : 36), "utf-16")

				if (hProcess := DllCall("kernel32\OpenProcess", "uint", PROCESS_QUERY_INFORMATION, "int", 0, "uint", ProcessID, "ptr"))
				{
					Size := VarSetStrCapacity(&ImagePath, 520)
					DllCall("kernel32\QueryFullProcessImageNameW", "ptr", hProcess, "uint", 0, "str", ImagePath, "uint*", Size)
					DllCall("kernel32\CloseHandle", "ptr", hProcess)
					if (ImagePath)
					{
						SHFILEINFOW := Buffer(A_PtrSize + 688, 0)
						if (DllCall("shell32\SHGetFileInfoW", "str", ImagePath, "uint", 0, "ptr", SHFILEINFOW, "uint", SHFILEINFOW.Size, "uint", 0x0101, "ptr"))
						{
							hIcon := NumGet(SHFILEINFOW, 0, "ptr")
							IconNumber := DllCall("comctl32\ImageList_ReplaceIcon", "ptr", ImageListID1, "int", -1, "ptr", hIcon) + 1
							DllCall("comctl32\ImageList_ReplaceIcon", "ptr", ImageListID2, "int", -1, "ptr", hIcon)
							DllCall("user32\DestroyIcon", "ptr", hIcon)
						}
					}
				}

				ROW["IconNumber"] := IconNumber ? IconNumber : 9999999
				TABLE[ProcessID]  := ROW
			}
		}
		return TABLE
	}
	return false
}


GetExtendedTcpTable(PROCESS_TABLE)
{
	static AF_INET := 2
	static ERROR_INSUFFICIENT_BUFFER := 122
	static NO_ERROR := 0
	static TCP_TABLE_OWNER_MODULE_ALL := 8
	static TCP_STATE := ["Closed", "Listen", "Syn Sent", "Syn Received", "Established", "Fin Wait 1", "Fin Wait 2", "Close Wait", "Closing", "Ack", "Time Wait", "Delete TCB"]

	TCP := Buffer(4, 0)
	if (DllCall("iphlpapi\GetExtendedTcpTable", "ptr", TCP, "uint*", &Size := 0, "int", 0, "uint", AF_INET, "uint", TCP_TABLE_OWNER_MODULE_ALL, "uint", 0) = ERROR_INSUFFICIENT_BUFFER)
	{
		TCP := Buffer(Size, 0)
		if (DllCall("iphlpapi\GetExtendedTcpTable", "ptr", TCP, "uint*", Size, "int", 0, "uint", AF_INET, "uint", TCP_TABLE_OWNER_MODULE_ALL, "uint", 0) = NO_ERROR)
		{
			TCP_TABLE := Map()
			NumEntries := NumGet(TCP, 0, "uint")
			loop NumEntries
			{
				TCP_ROW := Map(), ModuleName := ""
				Offset := 8 + ((A_Index - 1) * 160)
				TCP_ROW["State"]           := TCP_STATE[NumGet(TCP, Offset, "uint")]
				TCP_ROW["LocalAddr"]       := InetNtopW(AF_INET, TCP.Ptr + Offset + 4)
				TCP_ROW["LocalPort"]       := ntohs(NumGet(TCP, Offset + 8, "uint"))
				TCP_ROW["RemoteAddr"]      := InetNtopW(AF_INET, TCP.Ptr + Offset + 12)
				TCP_ROW["RemotePort"]      := ntohs(NumGet(TCP, Offset + 16, "uint"))
				TCP_ROW["OwningPID"]       := OwningPID := NumGet(TCP, Offset + 20, "uint")
				TCP_ROW["ProcessName"]     := OwningPID ? PROCESS_TABLE[OwningPID]["ExeFile"] : "[Time Wait]"
				TCP_ROW["CreateTimestamp"] := CreateTime(NumGet(TCP, Offset + 28, "uint") << 32 | NumGet(TCP, Offset + 32, "uint"))
				TCP_ROW["ModuleName"]      := GetOwnerModuleFromTcpEntry(TCP.Ptr + Offset)
				TCP_ROW["IconNumber"]      := OwningPID ? PROCESS_TABLE[OwningPID]["IconNumber"] : 9999999
				TCP_ROW["Protocol"]        := "TCP"
				TCP_TABLE[A_Index]         := TCP_ROW
			}
		}
		return TCP_TABLE
	}
	return false
}


GetExtendedTcp6Table(PROCESS_TABLE)
{
	static AF_INET6 := 23
	static ERROR_INSUFFICIENT_BUFFER := 122
	static NO_ERROR := 0
	static TCP_TABLE_OWNER_MODULE_ALL := 8
	static TCP_STATE := ["Closed", "Listen", "Syn Sent", "Syn Received", "Established", "Fin Wait 1", "Fin Wait 2", "Close Wait", "Closing", "Ack", "Time Wait", "Delete TCB"]

	TCP6 := Buffer(4, 0)
	if (DllCall("iphlpapi\GetExtendedTcpTable", "ptr", TCP6, "uint*", &Size := 0, "int", 0, "uint", AF_INET6, "uint", TCP_TABLE_OWNER_MODULE_ALL, "uint", 0) = ERROR_INSUFFICIENT_BUFFER)
	{
		TCP6 := Buffer(Size, 0)
		if (DllCall("iphlpapi\GetExtendedTcpTable", "ptr", TCP6, "uint*", Size, "int", 0, "uint", AF_INET6, "uint", TCP_TABLE_OWNER_MODULE_ALL, "uint", 0) = NO_ERROR)
		{
			TCP6_TABLE := Map()
			NumEntries := NumGet(TCP6, 0, "uint")
			loop NumEntries
			{
				TCP6_ROW := Map(), ModuleName := ""
				Offset := 8 + ((A_Index - 1) * 192)
				TCP6_ROW["LocalAddr"]       := InetNtopW(AF_INET6, TCP6.Ptr + Offset)
				TCP6_ROW["LocalScopeId"]    := ntohl(NumGet(TCP6, Offset + 16, "uint"))
				TCP6_ROW["LocalPort"]       := ntohs(NumGet(TCP6, Offset + 20, "uint"))
				TCP6_ROW["RemoteAddr"]      := InetNtopW(AF_INET6, TCP6.Ptr + Offset + 24)
				TCP6_ROW["RemoteScopeId"]   := ntohl(NumGet(TCP6, Offset + 40, "uint"))
				TCP6_ROW["RemotePort"]      := ntohs(NumGet(TCP6, Offset + 44, "uint"))
				TCP6_ROW["State"]           := TCP_STATE[NumGet(TCP6, Offset + 48, "uint")]
				TCP6_ROW["OwningPID"]       := OwningPID := NumGet(TCP6, Offset + 52, "uint")
				TCP6_ROW["ProcessName"]     := OwningPID ? PROCESS_TABLE[OwningPID]["ExeFile"] : "[Time Wait]"
				TCP6_ROW["CreateTimestamp"] := CreateTime(NumGet(TCP6, Offset + 60, "uint") << 32 | NumGet(TCP6, Offset + 64, "uint"))
				TCP6_ROW["ModuleName"]      := GetOwnerModuleFromTcp6Entry(TCP6.Ptr + Offset)
				TCP6_ROW["IconNumber"]      := OwningPID ? PROCESS_TABLE[OwningPID]["IconNumber"] : 9999999
				TCP6_ROW["Protocol"]        := "TCPv6"
				TCP6_TABLE[A_Index]         := TCP6_ROW
			}
		}
		return TCP6_TABLE
	}
	return false
}


GetExtendedUdpTable(PROCESS_TABLE)
{
	static AF_INET := 2
	static ERROR_INSUFFICIENT_BUFFER := 122
	static NO_ERROR := 0
	static UDP_TABLE_OWNER_MODULE := 2

	UDP := Buffer(4, 0)
	if (DllCall("iphlpapi\GetExtendedUdpTable", "ptr", UDP, "uint*", &Size := 0, "int", 0, "uint", AF_INET, "uint", UDP_TABLE_OWNER_MODULE, "uint", 0) = ERROR_INSUFFICIENT_BUFFER)
	{
		UDP := Buffer(Size, 0)
		if (DllCall("iphlpapi\GetExtendedUdpTable", "ptr", UDP, "uint*", Size, "int", 0, "uint", AF_INET, "uint", UDP_TABLE_OWNER_MODULE, "uint", 0) = NO_ERROR)
		{
			UDP_TABLE := Map()
			NumEntries := NumGet(UDP, 0, "uint")
			loop NumEntries
			{
				UDP_ROW := Map(), ModuleName := ""
				Offset := 8 + ((A_Index - 1) * 160)
				UDP_ROW["LocalAddr"]       := InetNtopW(AF_INET, UDP.Ptr + Offset)
				UDP_ROW["LocalPort"]       := ntohs(NumGet(UDP, Offset + 4, "uint"))
				UDP_ROW["OwningPID"]       := OwningPID := NumGet(UDP, Offset + 8, "uint")
				UDP_ROW["ProcessName"]     := OwningPID ? PROCESS_TABLE[OwningPID]["ExeFile"] : "[Time Wait]"
				UDP_ROW["CreateTimestamp"] := CreateTime(NumGet(UDP, Offset + 20, "uint") << 32 | NumGet(UDP, Offset + 24, "uint"))
				UDP_ROW["ModuleName"]      := GetOwnerModuleFromUdpEntry(UDP.Ptr + Offset)
				UDP_ROW["IconNumber"]      := OwningPID ? PROCESS_TABLE[OwningPID]["IconNumber"] : 9999999
				UDP_ROW["Protocol"]        := "UDP"
				UDP_ROW["State"]           := ""
				UDP_ROW["RemoteAddr"]      := "*"
				UDP_ROW["RemotePort"]      := ""
				UDP_TABLE[A_Index]         := UDP_ROW
			}
		}
		return UDP_TABLE
	}
	return false
}


GetExtendedUdp6Table(PROCESS_TABLE)
{
	static AF_INET6 := 23
	static ERROR_INSUFFICIENT_BUFFER := 122
	static NO_ERROR := 0
	static UDP_TABLE_OWNER_MODULE := 2

	UDP6 := Buffer(4, 0)
	if (DllCall("iphlpapi\GetExtendedUdpTable", "ptr", UDP6, "uint*", &Size := 0, "int", 0, "uint", AF_INET6, "uint", UDP_TABLE_OWNER_MODULE, "uint", 0) = ERROR_INSUFFICIENT_BUFFER)
	{
		UDP6 := Buffer(Size, 0)
		if (DllCall("iphlpapi\GetExtendedUdpTable", "ptr", UDP6, "uint*", Size, "int", 0, "uint", AF_INET6, "uint", UDP_TABLE_OWNER_MODULE, "uint", 0) = NO_ERROR)
		{
			UDP6_TABLE := Map()
			NumEntries := NumGet(UDP6, 0, "uint")
			loop NumEntries
			{
				UDP6_ROW := Map(), ModuleName := ""
				Offset := 8 + ((A_Index - 1) * 176)
				UDP6_ROW["LocalAddr"]       := NumGet(UDP6, Offset, "uchar")
				UDP6_ROW["LocalAddr"]       := InetNtopW(AF_INET6, UDP6.Ptr + Offset)
				UDP6_ROW["LocalScopeId"]    := ntohl(NumGet(UDP6, Offset + 16, "uint"))
				UDP6_ROW["LocalPort"]       := ntohs(NumGet(UDP6, Offset + 20, "uint"))
				UDP6_ROW["OwningPID"]       := OwningPID := NumGet(UDP6, Offset + 24, "uint")
				UDP6_ROW["ProcessName"]     := OwningPID ? PROCESS_TABLE[OwningPID]["ExeFile"] : "[Time Wait]"
				UDP6_ROW["CreateTimestamp"] := CreateTime(NumGet(UDP6, Offset + 36, "uint") << 32 | NumGet(UDP6, Offset + 40, "uint"))
				UDP6_ROW["ModuleName"]      := GetOwnerModuleFromUdp6Entry(UDP6.Ptr + Offset)
				UDP6_ROW["IconNumber"]      := OwningPID ? PROCESS_TABLE[OwningPID]["IconNumber"] : 9999999
				UDP6_ROW["Protocol"]        := "UDPv6"
				UDP6_ROW["State"]           := ""
				UDP6_ROW["RemoteAddr"]      := "*"
				UDP6_ROW["RemotePort"]      := ""
				UDP6_TABLE[A_Index]         := UDP6_ROW
			}
		}
		return UDP6_TABLE
	}
	return false
}


GetOwnerModuleFromTcpEntry(OWNER_MODULE)
{
	static NO_ERROR                  := 0
	static ERROR_INSUFFICIENT_BUFFER := 122
	static OWNER_MODULE_INFO_CLASS   := 0

	OWNER_MODULE_BASIC_INFO := Buffer(4, 0)
	if (DllCall("iphlpapi\GetOwnerModuleFromTcpEntry", "ptr", OWNER_MODULE, "int", OWNER_MODULE_INFO_CLASS, "ptr", OWNER_MODULE_BASIC_INFO, "uint*", &Size := 0) = ERROR_INSUFFICIENT_BUFFER)
	{
		OWNER_MODULE_BASIC_INFO := Buffer(Size, 0)
		if (DllCall("iphlpapi\GetOwnerModuleFromTcpEntry", "ptr", OWNER_MODULE, "int", OWNER_MODULE_INFO_CLASS, "ptr", OWNER_MODULE_BASIC_INFO, "uint*", Size) = NO_ERROR)
			return StrGet(NumGet(OWNER_MODULE_BASIC_INFO, 0, "ptr"))
	}
	return ""
}


GetOwnerModuleFromTcp6Entry(OWNER_MODULE)
{
	static NO_ERROR                  := 0
	static ERROR_INSUFFICIENT_BUFFER := 122
	static OWNER_MODULE_INFO_CLASS   := 0

	OWNER_MODULE_BASIC_INFO := Buffer(4, 0)
	if (DllCall("iphlpapi\GetOwnerModuleFromTcp6Entry", "ptr", OWNER_MODULE, "int", OWNER_MODULE_INFO_CLASS, "ptr", OWNER_MODULE_BASIC_INFO, "uint*", &Size := 0) = ERROR_INSUFFICIENT_BUFFER)
	{
		OWNER_MODULE_BASIC_INFO := Buffer(Size, 0)
		if (DllCall("iphlpapi\GetOwnerModuleFromTcp6Entry", "ptr", OWNER_MODULE, "int", OWNER_MODULE_INFO_CLASS, "ptr", OWNER_MODULE_BASIC_INFO, "uint*", Size) = NO_ERROR)
			return StrGet(NumGet(OWNER_MODULE_BASIC_INFO, 0, "ptr"))
	}
	return ""
}


GetOwnerModuleFromUdpEntry(OWNER_MODULE)
{
	static NO_ERROR                  := 0
	static ERROR_INSUFFICIENT_BUFFER := 122
	static OWNER_MODULE_INFO_CLASS   := 0

	OWNER_MODULE_BASIC_INFO := Buffer(4, 0)
	if (DllCall("iphlpapi\GetOwnerModuleFromUdpEntry", "ptr", OWNER_MODULE, "int", OWNER_MODULE_INFO_CLASS, "ptr", OWNER_MODULE_BASIC_INFO, "uint*", &Size := 0) = ERROR_INSUFFICIENT_BUFFER)
	{
		OWNER_MODULE_BASIC_INFO := Buffer(Size, 0)
		if (DllCall("iphlpapi\GetOwnerModuleFromUdpEntry", "ptr", OWNER_MODULE, "int", OWNER_MODULE_INFO_CLASS, "ptr", OWNER_MODULE_BASIC_INFO, "uint*", Size) = NO_ERROR)
			return StrGet(NumGet(OWNER_MODULE_BASIC_INFO, 0, "ptr"))
	}
	return ""
}


GetOwnerModuleFromUdp6Entry(OWNER_MODULE)
{
	static NO_ERROR                  := 0
	static ERROR_INSUFFICIENT_BUFFER := 122
	static OWNER_MODULE_INFO_CLASS   := 0

	OWNER_MODULE_BASIC_INFO := Buffer(4, 0)
	if (DllCall("iphlpapi\GetOwnerModuleFromUdp6Entry", "ptr", OWNER_MODULE, "int", OWNER_MODULE_INFO_CLASS, "ptr", OWNER_MODULE_BASIC_INFO, "uint*", &Size := 0) = ERROR_INSUFFICIENT_BUFFER)
	{
		OWNER_MODULE_BASIC_INFO := Buffer(Size, 0)
		if (DllCall("iphlpapi\GetOwnerModuleFromUdp6Entry", "ptr", OWNER_MODULE, "int", OWNER_MODULE_INFO_CLASS, "ptr", OWNER_MODULE_BASIC_INFO, "uint*", Size) = NO_ERROR)
			return StrGet(NumGet(OWNER_MODULE_BASIC_INFO, 0, "ptr"))
	}
	return ""
}


InetNtopW(Family, Addr)
{
	VarSetStrCapacity(&AddrString, Size := (Family = 2) ? 32 : 94)
	if (DllCall("ws2_32\InetNtopW", "int", Family, "ptr", Addr, "str", AddrString, "uint", Size))
		return AddrString
	return ""
}


ntohl(netlong)
{
	return DllCall("ws2_32\ntohl", "uint", netlong, "uint")
}


ntohs(netshort)
{
	return DllCall("ws2_32\ntohs", "ushort", netshort, "ushort")
}


CreateTime(FileTime)
{
	if !(FileTime)
		return ""
	SystemTime := Buffer(16, 0)
	if (DllCall("kernel32\FileTimeToSystemTime", "int64*", FileTime, "ptr", SystemTime))
	{
		LocalTime := Buffer(16, 0)
		if (DllCall("kernel32\SystemTimeToTzSpecificLocalTime", "ptr", 0, "ptr", SystemTime, "ptr", LocalTime))
		{
			return Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}"
                         , NumGet(LocalTime,  0, "ushort")
                         , NumGet(LocalTime,  2, "ushort")
                         , NumGet(LocalTime,  6, "ushort")
                         , NumGet(LocalTime,  8, "ushort")
                         , NumGet(LocalTime, 10, "ushort")
                         , NumGet(LocalTime, 12, "ushort"))
		}
		return false
	}
	return false
}


NetStat()
{
	Interval := (DDL1.Value = 1) ? 2000 : (DDL1.Value = 2) ? 5000 : (DDL1.Value = 3) ? 10000 : 5000
	LVCount := 0, TCPCount := 0, TCP6Count := 0, UDPCount := 0, UDP6Count := 0, LV_TABLE := []
	SetTimer NetStat, Interval

	if !(PROCESS_TABLE := Process32())
	{
		Main.Opt("+OwnDialogs")
		MsgBox("Process32 failed", "TCPView Error", "T5 16")
		ExitApp
	}

	if (CB1.Value)
	{
		if !(TCP_TABLE := GetExtendedTcpTable(PROCESS_TABLE))
		{
			Main.Opt("+OwnDialogs")
			MsgBox("GetExtendedTcpTable failed", "TCPView Error", "T5 16")
			ExitApp
		}
		TCPCount := TCP_TABLE.Count
		for i, v in TCP_TABLE
			LV_TABLE.Push(TCP_TABLE[i])
	}

	if (CB2.Value)
	{
		if !(TCP6_TABLE := GetExtendedTcp6Table(PROCESS_TABLE))
		{
			Main.Opt("+OwnDialogs")
			MsgBox("GetExtendedTcp6Table failed", "TCPView Error", "T5 16")
			ExitApp
		}
		TCP6Count := TCP6_TABLE.Count
		for i, v in TCP6_TABLE
			LV_TABLE.Push(TCP6_TABLE[i])
	}

	if (CB3.Value)
	{
		if !(UDP_TABLE := GetExtendedUdpTable(PROCESS_TABLE))
		{
			Main.Opt("+OwnDialogs")
			MsgBox("GetExtendedUdpTable failed", "TCPView Error", "T5 16")
			ExitApp
		}
		UDPCount := UDP_TABLE.Count
		for i, v in UDP_TABLE
			LV_TABLE.Push(UDP_TABLE[i])
	}

	if (CB4.Value)
	{
		if !(UDP6_TABLE := GetExtendedUdp6Table(PROCESS_TABLE))
		{
			Main.Opt("+OwnDialogs")
			MsgBox("GetExtendedUdp6Table failed", "TCPView Error", "T5 16")
			ExitApp
		}
		UDP6Count := UDP6_TABLE.Count
		for i, v in UDP6_TABLE
			LV_TABLE.Push(UDP6_TABLE[i])
	}

	LV.Opt("-Redraw")

	loop LV_TABLE.Length
	{
		v := LV_TABLE[A_Index]
		if (A_Index > LV.GetCount())
			LV.Add("Icon" . v["IconNumber"], v["ProcessName"], v["OwningPID"], v["Protocol"], v["State"], v["LocalAddr"], v["LocalPort"], v["RemoteAddr"], v["RemotePort"], v["CreateTimestamp"], v["ModuleName"])
		else
			LV.Modify(A_Index, "Icon" . v["IconNumber"], v["ProcessName"], v["OwningPID"], v["Protocol"], v["State"], v["LocalAddr"], v["LocalPort"], v["RemoteAddr"], v["RemotePort"], v["CreateTimestamp"], v["ModuleName"])
	}

	if ((Diff := LV.GetCount() - (TCPCount + TCP6Count + UDPCount + UDP6Count)) > 0)
		loop Diff > 0
			LV.Delete(LV.GetCount() - A_Index)

	LV.Opt("+Redraw")

	SB_C2 := 0, SB_C3 := 0, SB_C4 := 0, SB_C5 := 0
	loop SB_C1 := LV.GetCount()
	{
		if (LV.GetText(A_Index, 4) = "Established")
			SB_C2++
		if (LV.GetText(A_Index, 4) = "Listen")
			SB_C3++
		if (LV.GetText(A_Index, 4) = "Time Wait")
			SB_C4++
		if (LV.GetText(A_Index, 4) = "Close Wait")
			SB_C5++
	}
	SB.SetText(" Endpoints: "  SB_C1, 1)
	SB.SetText("Established: " SB_C2, 2)
	SB.SetText("Listening: "   SB_C3, 3)
	SB.SetText("Time Wait: "   SB_C4, 4)
	SB.SetText("Close Wait: "  SB_C5, 5)
	SB.SetText("Update: " StrLower(SubStr(DDL1.Text, 1, -4)), 6)
}

; ===============================================================================================================================