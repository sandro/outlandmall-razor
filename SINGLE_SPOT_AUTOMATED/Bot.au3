#include <WinAPIGdi.au3>
#include <Array.au3>
#include <Date.au3>
#include <File.au3>
#include <WinAPISysWin.au3>
#include <WindowsConstants.au3>
#include <GuiTab.au3>
#RequireAdmin

; #########
; User Variables
Global $JournalPath = "C:\Program Files (x86)\Ultima Online Outlands\ClassicUO\Data\Client\JournalLogs"
Global $APIKey = "ENTER API KEY HERE"
Global $FilesToUpload = 1
Global $PathToCURL = "C:\Windows\System32\curl.exe"
Global $QuitHotkey = "q"
Global $ClearGumpKey = "e"
Global $StartRazorHotkey = "w"


; ###########
; Program Variables
AutoItSetOption("PixelCoordMode", 2)
Global $scale = _WinAPI_EnumDisplaySettings('', $ENUM_CURRENT_SETTINGS)[0] / @DesktopWidth
Global $logoutGump = 4282424686
Global $connectionLostColor = 8684667
ConsoleWrite("Scale = " & _WinAPI_EnumDisplaySettings('', $ENUM_CURRENT_SETTINGS)[0] & " / " & @DesktopWidth & " = " & $scale & @CRLF)
Global $GUI = GUICreate("Outlands Mall", 400, 300)
Global $GUICTRL = GUICtrlCreateEdit("", 0, 0, 400, 300)
GUISetState(@SW_SHOW)

Func Print($str)
	ConsoleWrite($str & @CRLF)
EndFunc

Func GUILog($str)
	$str = _NowTime() & " " & $str & @CRLF
	GUICtrlSetData($GUICTRL, $str, 1)
EndFunc

Func GetLogoutGumpCoords($wh)
	Local $size = WinGetClientSize($wh)
	Local $coords[4]
	$coords[0] = $size[0]/2 - 25 *$scale
	$coords[1] = $size[1]/2 + 15 *$scale
	$coords[2] = $size[0]/2 + 15 *$scale
	$coords[3] = $size[1]/2 + 25 *$scale
	return $coords
EndFunc

Func GetLogoutChecksum($wh)
	Send($QuitHotkey) ; QuitGame macro
	Sleep(2000)
	Local $coords = GetLogoutGumpCoords($wh)
	Local $cksum = GetScaledChecksum($coords[0], $coords[1], $coords[2], $coords[3], $wh)
	Print("Logout Checksum: " & $cksum)
	return $cksum
EndFunc

Func GetScaledChecksum($left, $top, $right, $bottom, $wh)
	Local $cksum = PixelChecksum($left*$scale, $top*$scale, $right*$scale, $bottom*$scale, 1, $wh, 1)
	return $cksum
EndFunc

Func IsConnectionLost($wh)
	Local $size = WinGetClientSize($wh)
	WinActivate($wh)
	WinWaitActive($wh)
	MouseMove(0,0,0)
	Sleep(100)
	Local $color = PixelGetColor(($size[0]/2)*$scale,($size[1]/2)*$scale, $wh)
	Return $color == $connectionLostColor
EndFunc

Func LoginToUO($wh)
	GUILog("Logging back in	" & WinGetTitle($wh))
	While StringLeft(WinGetTitle($wh), 4) <> "UO -"
		WinActivate($wh)
		WinWaitActive($wh)
		Send("{ENTER}")
		Sleep(2000)
		GUILog("Enter and Wait	" & WinGetTitle($wh))
	WEnd
EndFunc

Func GetTZOffset()
	local $info = _Date_Time_GetTimeZoneInformation()
	local $offset = $info[1]
	If $info[0] == 2 Then
		$offset += $info[7]
	EndIf
	Return $offset
EndFunc

Func SyncJournals()
	GUILog("Syncing journals")
	local $journalFiles = _FileListToArray($JournalPath, "*.txt", 1, True)
	If Not IsArray($journalFiles) Then
		Msgbox(0,"","Sorry but " & $JournalPath & " has no files")
		Exit
	EndIf

	local $sortableFiles[$journalFiles[0]][2]
	Print("found " & $journalFiles[0] & " journal files")
	For $i=1 to UBound($journalFiles) - 1
		$sortableFiles[$i-1][0] = $journalFiles[$i]
		$sortableFiles[$i-1][1] = FileGetTime($journalFiles[$i],0,1)
	Next
	_ArraySort($sortableFiles,1,0,0,1)

	For $i=0 To UBound($sortableFiles) - 1
		local $file = $sortableFiles[$i][0]
		If $i >= $FilesToUpload Then
			ExitLoop
		EndIf
		Local $TimeZoneOffset = GetTZOffset()*-60
		Local $command = $PathToCURL & ' --header "Authorization: Bearer ' & $APIKey & '" -X POST https://outlandmalls.com/uploads -F file=@"' & $file & '" -F tz=' & $TimeZoneOffset
		GUILog($command)
		Local $ok = RunWait($command)
		GUILog("Upload ok? " & String($ok == 0) & " file num: " & $i)
	Next
	;_ArrayDisplay($sortableFiles)
EndFunc

Func GatherUOHandles()
	Local $windows = WinList("UO - ")
	Local $UOHandles[$windows[0][0]]
	Local $numWindows = $windows[0][0]
	GUILog("UO Windows found: " & $numWindows)

	If $numWindows < 1 then
		Msgbox(0,"","NO UO Windows Found")
		Exit
	EndIf

	For $x = 1 to $windows[0][0]
		Print($windows[$x][1])
		_ArrayPush($UOHandles, $windows[$x][1])
	Next
	return $UOHandles
EndFunc

Func GatherRazorHandles($UOHandles)
	Local $handles[UBound($UOHandles)]
	Local $matches
	AutoItSetOption("WinTitleMatchMode", 2)
	For $x = 0 to UBound($UOHandles) - 1
		$matches = StringRegExp(WinGetTitle($UOHandles[$x]), 'UO - (.+?) -', $STR_REGEXPARRAYMATCH)
		_ArrayPush($handles, WinGetHandle($matches[0] & " ([None])"))
	Next
	AutoItSetOption("WinTitleMatchMode", 1)
	return $handles
EndFunc

Func SyncAndLogin($wh)
	SyncJournals()
	Sleep(60000*30)
	LoginToUO($wh)
	Sleep(30000)
	GUILog("Waiting 180s before starting script")
	Sleep(60000*3)
	GUILog("Starting razor script " & WinGetTitle($wh))
	WinActivate($wh)
	WinWaitActive($wh)
	Send($StartRazorHotkey) ; Start Script macro
	Sleep(3000)
EndFunc

Func Main()
	Local $UOHandles = GatherUOHandles()
	Local $RazorHandles = GatherRazorHandles($UOHandles)

	For $wh in $UOHandles
		WinActivate($wh)
		WinWaitActive($wh)
		SendKeepActive($wh)
		Sleep(1000)
		GUILog("Starting script in window " & WinGetTitle($wh))
		Send($StartRazorHotkey) ; Start Script macro
		Sleep(3000)
	Next

	While True
		For $i = 0 to UBound($UOHandles) - 1
			Local $wh = $UOHandles[$i]
			Local $rwh = $RazorHandles[$i]
			GUILog("Checking if script is done " & WinGetTitle($rwh))
			Local $tabbar = ControlGetHandle($rwh, "", "[CLASS:WindowsForms10.SysTabControl32.app.0.1_r3_ad1]")
			_GUICtrlTab_ClickTab($tabbar, 9) ; Click on the Scripts tab
			Sleep(1000)
			Local $scriptPlaying = ControlGetText($rwh, "", "[TEXT:Stop]")
			If Not $scriptPlaying Then
				WinActivate($wh)
				WinWaitActive($wh)
				SendKeepActive($wh)
				Send($ClearGumpKey) ; Clear Gump
				GUILog("Sending Quit Hotkey for " & WinGetTitle($wh))
				Send($QuitHotkey) ; QuitGame macro
				Sleep(1000)
				Local $coords = GetLogoutGumpCoords($wh)
				MouseClick("", $coords[2], $coords[3])
				Sleep(2000)
				SyncAndLogin($wh)
			ElseIf IsConnectionLost($wh) Then
				Local $size = WinGetClientSize($wh)
				MouseClick("", $size[0]/2, ($size[1]/2)+90)
				Sleep(60000*15)
				SyncAndLogin($wh)
			Else
				Sleep(7000)
			EndIf
		Next
		SendKeepActive("")
	WEnd

EndFunc

Main()
