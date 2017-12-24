#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Icon.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Description=A simple discrete glucose tooltip for Nightscout under Windows
#AutoIt3Wrapper_Res_Fileversion=0.8.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Mathias Noack
#AutoIt3Wrapper_Res_Language=1031
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <WinHttp.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <TrayConstants.au3>
#include <StringConstants.au3>
#include <InetConstants.au3>
#include <Misc.au3>

; App title
Local $sTitle = "GlucoTT"

; Delete existing 'update files'
Local $iFileExists = FileExists(@ScriptDir & "\Update.*")
If $iFileExists Then
	$CMD = "del Update.*"
	RunWait(@ComSpec & " /c " & $CMD)
EndIf

; Check for a new version
Local $sFileVersion = FileGetVersion(@ScriptDir & "\" & $sTitle & ".exe")
Local $sFileNewVersion = InetRead("https://github.com/Matze1985/GlucoTT/blob/master/GlucoTT.au3")
Local $sFileCompareVersion = StringRegExp($sFileNewVersion, $sFileVersion, $STR_REGEXPMATCH)

; Returns 0 (no match)
If $sFileCompareVersion <> 1 Then
	Switch MsgBox($MB_YESNO, "Update", "New version available!" & @CRLF & @CRLF & "Download now?")
		Case $IDYES
			; Save the downloaded file to folder
			Local $sDownloadExeFilePath = (@ScriptDir & "\Update.exe")
			Local $sDownloadCmdFilePath = (@ScriptDir & "\Update.bat")

			; Download the files 'INET_FORCERELOAD'
			Local $hDownloadExe = InetGet("https://github.com/Matze1985/GlucoTT/blob/master/GlucoTT.exe?raw=true", $sDownloadExeFilePath, $INET_FORCERELOAD)
			Local $hDownloadCmd = InetGet("https://github.com/Matze1985/GlucoTT/blob/master/Update.bat?raw=true", $sDownloadCmdFilePath, $INET_FORCERELOAD)

			; Close the handle returned by InetGet.
			InetClose($hDownloadExe)
			InetClose($hDownloadCmd)

			MsgBox($MB_OK, "Info", "Download completed, the application must be restarted!")

			; Run cmd script
			Run($sDownloadCmdFilePath, "", @SW_SHOWDEFAULT)
			Exit
	EndSwitch
EndIf

; Settings file location
Local $sFilePath = @ScriptDir & "\Settings.txt"

; Open the file for reading and store the handle to a variable.
Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
If $hFileOpen = -1 Then
	MsgBox($MB_ICONERROR, "Error", "An error occurred when reading the file.")
	Exit
EndIf

TrayItemSetText($TRAY_ITEM_PAUSE, "Pause") ; Set the text of the default 'Pause' item.
TrayItemSetText($TRAY_ITEM_EXIT, "Close app") ; Set the text of the default 'Exit' item.

; Check for another instance of this program
If _Singleton($sTitle, 1) = 0 Then
	MsgBox(16, "Error", "Another instance of this program is already running.")
	Exit
EndIf

; Definition: Settings
Local $sDomain = FileReadLine($hFileOpen, 3)
Local $sDesktopW = FileReadLine($hFileOpen, Int(7))
Local $sDesktopH = FileReadLine($hFileOpen, Int(11))
Local $sInterval = FileReadLine($hFileOpen, Int(15))
Local $sReadOption = FileReadLine($hFileOpen, 19)
Local $sAlertLow = FileReadLine($hFileOpen, Int(23))
Local $sAlertHigh = FileReadLine($hFileOpen, Int(27))

; API-Page
Local $sPage = "/api/v1/entries/sgv?count=1"

; ErrorHandling
Local $checkNumbers = StringRegExp($sDesktopW & $sDesktopH & $sInterval, '^[0-9]+$', $STR_REGEXPMATCH)
If Not $checkNumbers Then
	MsgBox($MB_ICONERROR, "Error", "Only numbers allowed in the fields, please check:" & @CRLF & @CRLF & "- Desktop width (minus)" & @CRLF & "- Desktop height (minus)" & @CRLF & "- Interval (ms) for updating glucose" & @CRLF & "- Alert glucose lower then" & @CRLF & "- Alert glucose higher then")
	Exit
EndIf

Local $checkDomain = StringRegExp($sDomain, '^((https?|ftp|smtp):\/\/)?(www.)?[a-z0-9]+(\.[a-z]{2,}){1,3}(#?\/?[a-zA-Z0-9#]+)*\/?(\?[a-zA-Z0-9-_]+=[a-zA-Z0-9-%]+&?)?$', $STR_REGEXPMATCH)
If Not $checkDomain Then
	MsgBox($MB_ICONERROR, "Error", "Wrong URL in file!" & @CRLF & @CRLF & "Example:" & @CRLF & "https://account.azurewebsites.net")
	Exit
EndIf

Local $checkReadOption = StringRegExp($sReadOption, '^(|mmol\/l)$', $STR_REGEXPMATCH)
If Not $checkReadOption Then
	MsgBox($MB_ICONERROR, "Error", "Wrong read option!" & @CRLF & "Use empty for mg/dl or use mmol/l in the field as read option!")
	Exit
EndIf

Local $checkAlert = StringRegExp($sAlertLow & $sAlertHigh, '^[0-9.]+$', $STR_REGEXPMATCH)
If Not $checkAlert Then
	MsgBox($MB_ICONERROR, "Error", "Only point and numbers allowed in the following fields:" & @CRLF & @CRLF & "- Alert glucose lower then" & @CRLF & "- Alert glucose higher then")
	Exit
EndIf

; Initialize and get session handle
Local $hOpen = _WinHttpOpen()

While 1
	; Get connection handle
	Local $hConnect = _WinHttpConnect($hOpen, $sDomain)

	; Make a SimpleSSL request
	Local $hRequestSSL = _WinHttpSimpleSendSSLRequest($hConnect, Default, $sPage)

	; Read RequestSSL
	Local $sReturned = _WinHttpSimpleReadData($hRequestSSL)

	; Match result variables
	Local $sText = StringRegExpReplace($sReturned, "([0-9]{13})|([0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}.[0-9]{3}\+[0-9]{4})|([	])", " ")
	Local $iGlucose = StringRegExpReplace($sText, "[^0-9]+", "")

	; TrendArrows
	If StringInStr($sText, "DoubleUp") Then
		$iTrend = "⇈"
	EndIf
	If StringInStr($sText, "Flat") Then
		$iTrend = "→︎"
	EndIf
	If StringInStr($sText, "SingleUp") Then
		$iTrend = "↑"
	EndIf
	If StringInStr($sText, "FortyFiveUp") Then
		$iTrend = "↗"
	EndIf
	If StringInStr($sText, "FortyFiveDown") Then
		$iTrend = "↘"
	EndIf
	If StringInStr($sText, "SingleDown") Then
		$iTrend = "↓"
	EndIf
	If StringInStr($sText, "DoubleDown") Then
		$iTrend = "⇊"
	EndIf

	; Check mmol/l option
	If $sReadOption = "mmol/l" Then
		; Calculate mmol/l
		$iGlucoseMmol = Int($iGlucose) * 0.0555
		$iGlucoseResult = Round($iGlucoseMmol, 1)
	Else
		$iGlucoseResult = Int($iGlucose)
	EndIf

	If $iGlucoseResult < $sAlertLow Or $iGlucoseResult > $sAlertHigh Then
		ToolTip($iGlucoseResult & " " & $iTrend, @DesktopWidth - $sDesktopW, @DesktopHeight - $sDesktopH, $sTitle, 2, 2)
	Else
		ToolTip($iGlucoseResult & " " & $iTrend, @DesktopWidth - $sDesktopW, @DesktopHeight - $sDesktopH, $sTitle, 1, 2)
	EndIf
	Sleep($sInterval)
WEnd
FileClose($hFileOpen)
