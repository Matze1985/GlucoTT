#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
   #AutoIt3Wrapper_Icon=Icon.ico
   #AutoIt3Wrapper_UseX64=n
   #AutoIt3Wrapper_Res_Description=A simple discrete glucose tooltip for Nightscout under Windows
   #AutoIt3Wrapper_Res_Fileversion=1.4.3.0
   #AutoIt3Wrapper_Res_LegalCopyright=Mathias Noack
   #AutoIt3Wrapper_Res_Language=1031
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <WinHttp.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <TrayConstants.au3>
#include <StringConstants.au3>
#include <CheckUpdate.au3>

; App title
Local $sTitle = "GlucoTT"

; Check for another instance of this program
If _Singleton($sTitle, 1) = 0 Then
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Another instance of this program is already running.")
   Exit
EndIf

; Check internet connection on Start/Loop - Return 1 for ON | Return 0 for OFF
Func _CheckConnection()
   Local $isReturn = DllCall("wininet.dll", "int", "InternetGetConnectedState", "int", 0, "int", 0)
   If (@error) Or ($isReturn[0] = 0) Then Return SetError(1, 0, 0)
   Return 1
EndFunc

; Check for updates
CheckUpdate($sTitle & (@Compiled ? ".exe" : ".au3"), $sVersion, "https://raw.githubusercontent.com/Matze1985/GlucoTT/master/Update/CheckUpdate.txt")
If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " @error " & @error & @CRLF)

; Settings file location
Local $sFilePath = @ScriptDir & "\Settings.txt"

; Open the file for reading and store the handle to a variable.
Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
If $hFileOpen = -1 Then
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "An error occurred when reading the file.")
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

; ErrorHandling: Checks
Local $checkNumbers = StringRegExp($sDesktopW & $sDesktopH & $sInterval, '^[0-9]+$', $STR_REGEXPMATCH)
If Not $checkNumbers Then
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Only numbers allowed in the fields, please check:" & @CRLF & @CRLF & "- Desktop width (minus)" & @CRLF & "- Desktop height (minus)" & @CRLF & "- Interval (ms) for updating glucose" & @CRLF & "- Alert glucose lower then" & @CRLF & "- Alert glucose higher then")
   Exit
EndIf

Local $checkDomain = StringRegExp($sDomain, '^((https):\/\/)?(www.)?[a-z0-9]+(\.[a-z]{2,}){1,3}(#?\/?[a-zA-Z0-9#]+)*\/?(\?[a-zA-Z0-9-_]+=[a-zA-Z0-9-%]+&?)?$', $STR_REGEXPMATCH)
If Not $checkDomain Then
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Wrong URL in file!" & @CRLF & @CRLF & "Example:" & @CRLF & "https://account.azurewebsites.net")
   Exit
EndIf

Local $checkReadOption = StringRegExp($sReadOption, '^(|mmol\/l)$', $STR_REGEXPMATCH)
If Not $checkReadOption Then
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Wrong read option!" & @CRLF & "Use empty for mg/dl or use mmol/l in the field as read option!")
   Exit
EndIf

Local $checkAlert = StringRegExp($sAlertLow & $sAlertHigh, '^[0-9.]+$', $STR_REGEXPMATCH)
If Not $checkAlert Then
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Only point and numbers allowed in the following fields:" & @CRLF & @CRLF & "- Alert glucose lower then" & @CRLF & "- Alert glucose higher then")
   Exit
EndIf

; API-Page
Global $sPage = "/api/v1/entries/sgv?count=2"

; Initialize and get session handle and get connection handle
Local $hOpen = _WinHttpOpen()
Local $hConnect = _WinHttpConnect($hOpen, $sDomain)

Func _Tooltip()

   ; Check connection
   Local $checkInet = _CheckConnection()
   Local $checkInetMsg

   If $checkInet <> 1 Then
      $checkInetMsg = "✕"
   EndIf

   ; Make a SimpleSSL request
   Local $hRequestSSL = _WinHttpSimpleSendSSLRequest($hConnect, Default, $sPage)

   ; Read RequestSSL
   Local $sReturned = _WinHttpSimpleReadData($hRequestSSL)

   ; Match result variables
   Local $sFirstTextLine = StringRegExpReplace($sReturned, "(.*$)", "")
   Local $sSecondTextLine = StringRegExpReplace($sReturned, "(\A.*)", "")
   Local $sCountMatch = "([0-9]{13})|([0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}.[0-9]{3}\+[0-9]{4})|([	]|(openLibreReader-ios-blueReader-[0-9])|(\.[0-9]{1,4}))"
   Local $sText = StringRegExpReplace($sFirstTextLine, $sCountMatch, " ")
   Local $sLastText = StringRegExpReplace($sSecondTextLine, $sCountMatch, " ")
   Local $sGlucose = StringRegExpReplace($sText, "[^0-9]+", "")
   Local $sLastGlucose = StringRegExpReplace($sLastText, "[^0-9]+", "")

   ; TrendArrows
   Local $sTrend
   If StringInStr($sText, "DoubleUp") Then
      $sTrend = "⇈"
   EndIf
   If StringInStr($sText, "Flat") Then
      $sTrend = "→︎"
   EndIf
   If StringInStr($sText, "SingleUp") Then
      $sTrend = "↑"
   EndIf
   If StringInStr($sText, "FortyFiveUp") Then
      $sTrend = "↗"
   EndIf
   If StringInStr($sText, "FortyFiveDown") Then
      $sTrend = "↘"
   EndIf
   If StringInStr($sText, "SingleDown") Then
      $sTrend = "↓"
   EndIf
   If StringInStr($sText, "DoubleDown") Then
      $sTrend = "⇊"
   EndIf

   ; Check mmol/l option
   Local $sCalcGlucose, $sGlucoseMmol, $sLastGlucoseMmol, $sGlucoseResult, $sLastGlucoseResult
   If $sReadOption = "mmol/l" Then
      ; Calculate mmol/l
      $sCalcGlucose = 0.0555
      $sGlucoseMmol = Int($sGlucose) * $sCalcGlucose
      $sLastGlucoseMmol = Int($sLastGlucose) * $sCalcGlucose
      $sGlucoseResult = Round($sGlucoseMmol, 1)
      $sLastGlucoseResult = Round($sLastGlucoseMmol, 1)
   Else
      $sGlucoseResult = Int($sGlucose)
      $sLastGlucoseResult = Int($sLastGlucose)
   EndIf

   ; Calculate delta
   Local $sDelta
   Local $sTmpDelta = Round($sGlucoseResult - $sLastGlucoseResult, 1)
   If $sTmpDelta > 0 Then
      $sDelta = "+" & $sTmpDelta
   Else
	  $sDelta = $sTmpDelta
   EndIf
   If $sTmpDelta = 0 Then
      $sDelta = "±" & $sTmpDelta
   EndIf

   ; Check tooltip alarm
   Local $sAlarm

   If $sGlucoseResult < $sAlertLow Or $sGlucoseResult > $sAlertHigh Then
      $sAlarm = "2"
   Else
      $sAlarm = "1"
   EndIf

   ; Running tooltip
   If $checkInet <> 1 Then
      ToolTip($checkInetMsg, @DesktopWidth - $sDesktopW, @DesktopHeight - $sDesktopH, $checkInetMsg, 3, 2)
   Else
      ToolTip("   " & $sDelta, @DesktopWidth - $sDesktopW, @DesktopHeight - $sDesktopH, "   " & $sGlucoseResult & " " & $sTrend & "  ", $sAlarm, 2)
   EndIf

   Sleep($sInterval)

EndFunc

; TrayMenu
Opt("TrayAutoPause", 0) ; The script will not pause when selecting the tray icon.
Opt("TrayMenuMode", 2)

Local $idNightscout = TrayCreateItem("Nightscout")
TrayItemSetText($TRAY_ITEM_PAUSE, "Pause") ; Set the text of the default 'Pause' item.
TrayItemSetText($TRAY_ITEM_EXIT, "Close app")

TraySetState()
While 1
   Local $msg = TrayGetMsg()
   Select
      Case $msg = 0
         _Tooltip()
         ContinueLoop
      Case $msg = $idNightscout
         ShellExecute($sDomain)
   EndSelect
WEnd

; Close File/WinHttp
_WinHttpCloseHandle($hConnect)
FileClose($hFileOpen)