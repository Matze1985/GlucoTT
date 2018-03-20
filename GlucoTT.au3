#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
   #AutoIt3Wrapper_Icon=Icon.ico
   #AutoIt3Wrapper_UseX64=n
   #AutoIt3Wrapper_Res_Description=A simple discrete glucose tooltip for Nightscout under Windows
   #AutoIt3Wrapper_Res_Fileversion=2.0.0.0
   #AutoIt3Wrapper_Res_LegalCopyright=Mathias Noack
   #AutoIt3Wrapper_Res_Language=1031
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <WinHttp.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <TrayConstants.au3>
#include <StringConstants.au3>
#include <Include\CheckUpdate.au3>
#include <WinAPIFiles.au3>

Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Opt("TrayOnEventMode", 1) ; Enable TrayOnEventMode.
Opt("WinTitleMatchMode", 2) ;1=start, 2=subStr, 3=exact, 4=advanced, -1 to -4=Nocase

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

; Read settings from ini file
Local $sFile = "Settings.ini"
Local $sIniCategory = "Settings"
Local $sIniTitleNightscout = "Nightscout"
Local $sIniDefaultNightscout = "https://<account>.azurewebsites.net"
Local $sIniTitleDesktopWidth = "Desktop width (minus)"
Local $sIniDefaultDesktopWidth = "43"
Local $sIniTitleDesktopHeight = "Desktop height (minus)"
Local $sIniDefaultDesktopHeight = "68"
Local $sInputDomain = IniRead($sFile, $sIniCategory, $sIniTitleNightscout, $sIniDefaultNightscout)
Local $sInputDesktopWidth = IniRead($sFile, $sIniCategory, $sIniTitleDesktopWidth, $sIniDefaultDesktopWidth)
Local $sInputDesktophHeight = IniRead($sFile, $sIniCategory, $sIniTitleDesktopHeight, $sIniDefaultDesktopHeight)

; Check Domain settings
Local $checkDomain = StringRegExp($sInputDomain, '^(?:https?:\/\/)?(?:www\.)?([^\s\:\/\?\[\]\@\!\$\&\"\(\)\*\+\,\;\=\<\>\#\%\''\"{"\}\|\\\^\`]{1,63}\.(?:[a-z]{2,}))(?:\/|:[0-9]{1,7}|\?|\&|\s|$)\/?')
If Not $checkDomain Then
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Wrong URL for Nightscout set!" & @CRLF & @CRLF & "Example:" & @CRLF & "https://account.azurewebsites.net")
   _Settings()
EndIf

; Check desktop width and height settings
Local $checkNumbers = StringRegExp($sInputDesktopWidth & $sInputDesktophHeight, '^[0-9]+$')
If Not $checkNumbers Then
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Only numbers allowed in the fields, please check:" & @CRLF & @CRLF & "- Desktop width (minus)" & @CRLF & "- Desktop height (minus)")
   _Settings()
EndIf

; After closing restarts the app
Func _Restart()
   Run(@ComSpec & " /c " & 'TIMEOUT /T 1 & START "" "' & @ScriptFullPath & '"', "", @SW_HIDE)
   Exit
EndFunc

; API-Pages
Local $sPage = "/api/v1/entries/sgv?count=2"
Local $sJsonState = "/api/v1/status.json"

; Initialize and get session handle and get connection handle
Local $hOpen = _WinHttpOpen()
Local $hConnect = _WinHttpConnect($hOpen, $sInputDomain)

Func _Tooltip()

   ; Check connection
   Local $checkInet = _CheckConnection()
   Local $wrongMsg

   If $checkInet <> 1 Then
      $wrongMsg = "✕"
   EndIf

   ; Make a SimpleSSL request
   Local $hRequestSSL = _WinHttpSimpleSendSSLRequest($hConnect, Default, $sPage)
   Local $hRequestJsonSSL = _WinHttpSimpleSendSSLRequest($hConnect, Default, $sJsonState)

   ; Read RequestSSL
   Local $sReturned = _WinHttpSimpleReadData($hRequestSSL)
   Local $sReturnedJson = _WinHttpSimpleReadData($hRequestJsonSSL)

   ; Match result variables from page
   Local $sFirstTextLine = StringRegExpReplace($sReturned, "(.*$)", "")
   Local $sSecondTextLine = StringRegExpReplace($sReturned, "(\A.*)", "")
   Local $sCountMatch = "([0-9]{13})|([0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}.[0-9]{3}\+[0-9]{4})|([	]|(openLibreReader-ios-blueReader-[0-9])|(\.[0-9]{1,4}))"
   Local $sText = StringRegExpReplace($sFirstTextLine, $sCountMatch, " ")
   Local $sLastText = StringRegExpReplace($sSecondTextLine, $sCountMatch, " ")
   Local $sGlucose = StringRegExpReplace($sText, "[^0-9]+", "")
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Glucose: " & $sGlucose & @CRLF)
   Local $sLastGlucose = StringRegExpReplace($sLastText, "[^0-9]+", "")

   ; Check time readings
   Local $sYear = StringLeft($sFirstTextLine, 4)
   Local $sMonth = StringMid($sFirstTextLine, 6, 2)
   Local $sDay = StringMid($sFirstTextLine, 9, 2)
   Local $sHour = StringMid($sFirstTextLine, 12, 2)
   Local $sMin = StringMid($sFirstTextLine, 15, 2)
   Local $sSec = StringMid($sFirstTextLine, 18, 2)
   Local $iDateCalc = _DateDiff('n', $sYear & "/" & $sMonth & "/" & $sDay & " " & $sHour & ":" & $sMin & ":" & $sSec, _NowCalc())

   ; Settings check for json: urgentRes
   Local $sReadIntervalMin = StringRegExpReplace($sReturnedJson, '.*"urgentRes":([^"]+),".*', '\1')
   Local $checkReadIntervalMin = StringRegExp($sReadIntervalMin, '([A-Za-z,":-{}])')
   If $checkReadIntervalMin = 1 Then
      $sReadIntervalMin = StringRegExpReplace($sReturnedJson, '.*"urgentRes":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Read interval (min): " & $sReadIntervalMin & @CRLF)

   ; Calculate read interval
   Local $sReadDiffMin = Int($sReadIntervalMin) - Int($iDateCalc)

   ; Check negative minutes and replace "-"
   If $sReadDiffMin < 0 Then
      $sReadDiffMin = StringReplace($sReadDiffMin, "-", "")
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Read diff (min): " & $sReadDiffMin & @CRLF)

   Local $sReadIntervalMs = Int($sReadDiffMin) * 60000
   Local $sReadInterval = Int($sReadIntervalMs) - Int($sReadIntervalMs) + 60000

   ; Settings check for json: units
   Local $sReadOption = StringRegExpReplace($sReturnedJson, '.*"units":"([^"]+)",.*', '\1')
   Local $checkReadOption = StringRegExp($sReadOption, '([{"A-Z0-9abcefhijknpqrstuvwxyz}:,])')
   If $checkReadOption = 1 Then
      $sReadOption = StringRegExpReplace($sReturnedJson, '.*"units":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Units: " & $sReadOption & @CRLF)

   ; Check mmol or mg/dl
   Local $checkReadOptionValues = StringRegExp($sReadOption, '(mmol|mg\/dl)')

   If Not $checkReadOptionValues Then
      _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", 'Wrong read option!' & @CRLF & 'Please set:' & @CRLF & 'DISPLAY_UNITS' & @CRLF & 'in the "Nightscout" application settings' & @CRLF & 'with a "mmol" or "mg/dl" value!')
   EndIf

   ; Alarm settings check for json: alarmLow
   Local $sAlertLow = StringRegExpReplace($sReturnedJson, '.*"alarmLow":"([^"]+)",".*', '\1')
   Local $checkAlertLow = StringRegExp($sAlertLow, '([A-Za-z,":-{}])')
   If $checkAlertLow = 1 Then
      $sAlertLow = StringRegExpReplace($sReturnedJson, '.*"alarmLow":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Low: " & $sAlertLow & @CRLF)

   ; Alarm settings check for json: alarmHigh
   Local $sAlertHigh = StringRegExpReplace($sReturnedJson, '.*"alarmHigh":"([^"]+)",".*', '\1')
   Local $checkAlertHigh = StringRegExp($sAlertHigh, '([A-Za-z,":-{}])')
   If $checkAlertHigh = 1 Then
      $sAlertHigh = StringRegExpReplace($sReturnedJson, '.*"alarmHigh":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " High: " & $sAlertHigh & @CRLF)

   ; Alarm settings check for json: alarmUrgentLow
   Local $sAlertLowUrgent = StringRegExpReplace($sReturnedJson, '.*"alarmUrgentLow":"([^"]+)",".*', '\1')
   Local $checkAlertLowUrgent = StringRegExp($sAlertLowUrgent, '([A-Za-z,":-{}])')
   If $checkAlertLowUrgent = 1 Then
      $sAlertLowUrgent = StringRegExpReplace($sReturnedJson, '.*"alarmUrgentLow":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " UrgentLow: " & $sAlertLowUrgent & @CRLF)

   ; Alarm settings check for json: alarmUrgentHigh
   Local $sAlertHighUrgent = StringRegExpReplace($sReturnedJson, '.*"alarmUrgentHigh":"([^"]+)",.*', '\1')
   Local $checkAlertHighUrgent = StringRegExp($sAlertHighUrgent, '([A-Za-z,":-{}])')
   If $checkAlertHighUrgent = 1 Then
      $sAlertHighUrgent = StringRegExpReplace($sReturnedJson, '.*"alarmUrgentHigh":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " UrgentHigh: " & $sAlertHighUrgent & @CRLF)

   ; Check false alarms and mg/dl values
   Local $checkAlertValues = StringRegExp($sAlertLow & $sAlertHigh & $sAlertLowUrgent & $sAlertHighUrgent, '([0-9]{1,3})')
   If Not $checkAlertValues Then
      _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", 'Please set:' & @CRLF & 'ALARM_LOW, ALARM_HIGH' & @CRLF & 'or' & @CRLF & 'ALARM_URGENT_LOW, ALARM_URGENT_HIGH' & @CRLF & 'in the "Nightscout" application settings' & @CRLF & 'with a "mg/dl" value!')
   EndIf

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
   If $sReadOption = "mmol" Then
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
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Glucose result: " & $sGlucoseResult & @CRLF)
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Last glucose result: " & $sLastGlucoseResult & @CRLF)

   ; Check valid glucose with msg box
   Local $sCheckGlucose = StringRegExp($sGlucoseResult, '(^[0-9]{4,})')
   If $sCheckGlucose Then
      _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", 'Nightscout is not working!')
      Exit
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
   If Int($sGlucose) <= Int($sAlertLow) Or Int($sGlucose) >= Int($sAlertHigh) Or Int($sGlucose) <= Int($sAlertLowUrgent) Or Int($sGlucose) >= Int($sAlertHighUrgent) Then
      $sAlarm = "2" ;=Warning icon
   Else
      $sAlarm = "1" ;=Info icon
   EndIf

   ; Running tooltip
   Local $sDesktopWidth = IniRead($sFile, $sIniCategory, $sIniTitleDesktopWidth, $sIniDefaultDesktopWidth)
   Local $sDesktopHeight = IniRead($sFile, $sIniCategory, $sIniTitleDesktopHeight, $sIniDefaultDesktopHeight)

   ; If 0 min then 1 min
   If $iDateCalc = 0 Then
      $iDateCalc = 1
   EndIf

   If $checkInet <> 1 Or $sGlucoseResult = 0 And $sDelta = 0 Then
      ToolTip($wrongMsg, @DesktopWidth - $sDesktopWidth, @DesktopHeight - $sDesktopHeight, $wrongMsg, 3, 2)
   Else
      ToolTip("   " & $sDelta & " " & @CR & "   " & $iDateCalc & " min", @DesktopWidth - $sDesktopWidth, @DesktopHeight - $sDesktopHeight, "   " & $sGlucoseResult & " " & $sTrend & "  ", $sAlarm, 2)
   EndIf

   Sleep($sReadInterval)

EndFunc

; Start TrayMenu with tooltip
_TrayMenu()

Func _TrayMenu()
   TrayCreateItem("Nightscout")
   TrayItemSetOnEvent(-1, "_Nightscout")
   TrayCreateItem("") ; Create a separator line.
   TrayCreateItem("Settings")
   TrayItemSetOnEvent(-1, "_Settings")
   TrayCreateItem("") ; Create a separator line.
   TrayCreateItem("Close app")
   TrayItemSetOnEvent(-1, "_Exit")
   TraySetState($TRAY_ICONSTATE_SHOW) ; Show the tray menu.

   While 1
      _Tooltip() ; An idle loop.
   WEnd
EndFunc

; Set function for buttons in TrayMenu
Func _Nightscout()
   ShellExecute($sInputDomain)
EndFunc

; Set settings in GUI
Func _Settings()
   Global $hSave, $msg
   GUICreate($sIniCategory, 320, 155, @DesktopWidth / 2 - 160, @DesktopHeight / 2 - 45)
   GUICtrlCreateLabel($sIniTitleNightscout, 10, 5, 70)
   Global $hInputDomain = GUICtrlCreateInput($sInputDomain, 10, 20, 300, 20)
   GUICtrlCreateLabel($sIniTitleDesktopWidth, 10, 45, 140)
   Global $hInputDesktopWidth = GUICtrlCreateInput($sInputDesktopWidth, 10, 60, 50, 20, $ES_NUMBER)
   GUICtrlCreateLabel($sIniTitleDesktopHeight, 10, 85, 140)
   Global $hInputDesktopHeight = GUICtrlCreateInput($sInputDesktophHeight, 10, 100, 50, 20, $ES_NUMBER)
   $hSave = GUICtrlCreateButton("Save", 10, 125, 300, 20)
   GUISetState()
   $msg = 0
   While $msg <> $GUI_EVENT_CLOSE
      $msg = GUIGetMsg()
      Select
         Case $msg = $hSave
            IniWrite($sFile, $sIniCategory, $sIniTitleNightscout, GUICtrlRead($hInputDomain))
            IniWrite($sFile, $sIniCategory, $sIniTitleDesktopWidth, GUICtrlRead($hInputDesktopWidth))
            IniWrite($sFile, $sIniCategory, $sIniTitleDesktopHeight, GUICtrlRead($hInputDesktopHeight))
            _Restart()
         Case $msg = $GUI_EVENT_CLOSE
            GUIDelete($sIniCategory)
      EndSelect
   WEnd
EndFunc

Func _Exit()
   Exit
EndFunc

; Close WinHttp
_WinHttpCloseHandle($hConnect)