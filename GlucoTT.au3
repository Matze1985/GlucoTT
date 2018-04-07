#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
   #AutoIt3Wrapper_Icon=Icon.ico
   #AutoIt3Wrapper_UseX64=n
   #AutoIt3Wrapper_Res_Description=A simple discrete glucose tooltip for Nightscout under Windows
   #AutoIt3Wrapper_Res_Fileversion=2.2.0.0
   #AutoIt3Wrapper_Res_LegalCopyright=Mathias Noack
   #AutoIt3Wrapper_Res_Language=1031
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <WinHttp.au3>
#include <TrayConstants.au3>
#include <Include\CheckUpdate.au3>
#include <Include\_Startup.au3>

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

; Read settings from ini file
Local $sFilePath = @ScriptDir & "\"
Local $sFile = "Settings.ini"
Local $sFileFullPath = $sFilePath & $sFile
Local $sIniCategory = "Settings"
Local $sIniTitleNightscout = "Nightscout"
Local $sIniDefaultNightscout = "https://<account>.azurewebsites.net"
Local $sIniTitleDesktopWidth = "Desktop width (minus)"
Local $sIniDefaultDesktopWidth = "43"
Local $sIniTitleDesktopHeight = "Desktop height (minus)"
Local $sIniDefaultDesktopHeight = "68"
Local $sIniTitleOptions = "Options"
Local $sIniTitleAutostart = "Autostart"
Local $sIniDefaultCheckboxAutostart = "1"
Local $sIniTitleUpdate = "Update"
Local $sIniDefaultCheckboxUpdate = "1"
Local $sInputDomain = IniRead($sFileFullPath, $sIniCategory, $sIniTitleNightscout, $sIniDefaultNightscout)
Local $sInputDesktopWidth = IniRead($sFileFullPath, $sIniCategory, $sIniTitleDesktopWidth, $sIniDefaultDesktopWidth)
Local $sInputDesktophHeight = IniRead($sFileFullPath, $sIniCategory, $sIniTitleDesktopHeight, $sIniDefaultDesktopHeight)
Local $sCheckboxAutostart = IniRead($sFileFullPath, $sIniCategory, $sIniTitleAutostart, $sIniDefaultCheckboxAutostart)
Local $sCheckboxUpdate = IniRead($sFileFullPath, $sIniCategory, $sIniTitleUpdate, $sIniDefaultCheckboxUpdate)

; Check for updates
If $sCheckboxUpdate = "1" Then
   CheckUpdate($sTitle & (@Compiled ? ".exe" : ".au3"), $sVersion, "https://raw.githubusercontent.com/Matze1985/GlucoTT/master/Update/CheckUpdate.txt")
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " @error " & @error & @CRLF)
EndIf

; Check Shortcut state from ini file
If Int($sCheckboxAutostart) < 4 Then
   _StartupRegistry_Install() ; Add the running EXE to the Current Users Run registry key.
Else
   _StartupRegistry_Uninstall() ; Remove the running EXE from the Current Users Run registry key.
EndIf

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

; Check checkbox options
If $sCheckboxAutostart = 1 Or $sCheckboxAutostart = 4 And $sCheckboxUpdate = 1 Or $sCheckboxUpdate = 4 Then
Else
   ShellExecute(@ScriptDir & "\" & $sFile)
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Only numbers for options allowed, please check your ini file:" & @CRLF & @CRLF & "- Checkbox number: 1 (on)" & @CRLF & "- Checkbox number: 4 (off)")
   WinWaitClose($sFile)
   _Restart()
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

   ; Set wrong msg in tooltip
   Local $wrongMsg = "✕"

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
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Glucose : " & $sGlucose & @CRLF)
   Local $sLastGlucose = StringRegExpReplace($sLastText, "[^0-9]+", "")

   ; Check time readings
   Local $sYear = StringLeft($sFirstTextLine, 4)
   Local $sMonth = StringMid($sFirstTextLine, 6, 2)
   Local $sDay = StringMid($sFirstTextLine, 9, 2)
   Local $sHour = StringMid($sFirstTextLine, 12, 2)
   Local $sMin = StringMid($sFirstTextLine, 15, 2)
   Local $sSec = StringMid($sFirstTextLine, 18, 2)
   Local $sLastYearMonthDayHourMinSec = $sYear & "/" & $sMonth & "/" & $sDay & " " & $sHour & ":" & $sMin & ":" & $sSec
   Local $sCurrentYearMonthDayHourMinSec = @YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC

   ; Reading last glucose (min)
   Local $sLastReadingGlucoseMin = _DateDiff('n', $sLastYearMonthDayHourMinSec, $sCurrentYearMonthDayHourMinSec)
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Last reading glucose (min) : " & $sLastReadingGlucoseMin & @CRLF)

   ; Interval for loop interval to read glucose
   Local $sReadInterval = Int(60000)
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Read interval (ms) : " & $sReadInterval & @CRLF)

   ; Settings check for json: status
   Local $sStatus = StringRegExpReplace($sReturnedJson, '.*"status":"([^"]+)",.*', '\1')
   Local $checkStatus = StringRegExp($sStatus, '([{"A-Z0-9abcdefghijlmnpqrstuvwxyz}:,-.])')
   If $checkStatus = 1 Then
      $sStatus = StringRegExpReplace($sReturnedJson, '.*"status":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Status : " & $sStatus & @CRLF)

   ; Settings check for json: units
   Local $sReadOption = StringRegExpReplace($sReturnedJson, '.*"units":"([^"]+)",.*', '\1')
   Local $checkReadOption = StringRegExp($sReadOption, '([{"A-Z0-9abcefhijknpqrstuvwxyz}:,-.])')
   If $checkReadOption = 1 Then
      $sReadOption = StringRegExpReplace($sReturnedJson, '.*"units":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Units : " & $sReadOption & @CRLF)

   ; Alarm settings check for json: alarmLow
   Local $sAlertLow = StringRegExpReplace($sReturnedJson, '.*"alarmLow":"([^"]+)",".*', '\1')
   Local $checkAlertLow = StringRegExp($sAlertLow, '([A-Za-z,":-{}])')
   If $checkAlertLow = 1 Then
      $sAlertLow = StringRegExpReplace($sReturnedJson, '.*"alarmLow":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Low : " & $sAlertLow & @CRLF)

   ; Alarm settings check for json: alarmHigh
   Local $sAlertHigh = StringRegExpReplace($sReturnedJson, '.*"alarmHigh":"([^"]+)",".*', '\1')
   Local $checkAlertHigh = StringRegExp($sAlertHigh, '([A-Za-z,":-{}])')
   If $checkAlertHigh = 1 Then
      $sAlertHigh = StringRegExpReplace($sReturnedJson, '.*"alarmHigh":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " High : " & $sAlertHigh & @CRLF)

   ; Alarm settings check for json: alarmUrgentLow
   Local $sAlertLowUrgent = StringRegExpReplace($sReturnedJson, '.*"alarmUrgentLow":"([^"]+)",".*', '\1')
   Local $checkAlertLowUrgent = StringRegExp($sAlertLowUrgent, '([A-Za-z,":-{}])')
   If $checkAlertLowUrgent = 1 Then
      $sAlertLowUrgent = StringRegExpReplace($sReturnedJson, '.*"alarmUrgentLow":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " UrgentLow : " & $sAlertLowUrgent & @CRLF)

   ; Alarm settings check for json: alarmUrgentHigh
   Local $sAlertHighUrgent = StringRegExpReplace($sReturnedJson, '.*"alarmUrgentHigh":"([^"]+)",.*', '\1')
   Local $checkAlertHighUrgent = StringRegExp($sAlertHighUrgent, '([A-Za-z,":-{}])')
   If $checkAlertHighUrgent = 1 Then
      $sAlertHighUrgent = StringRegExpReplace($sReturnedJson, '.*"alarmUrgentHigh":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " UrgentHigh : " & $sAlertHighUrgent & @CRLF)

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
   Local $sCalcGlucose, $sGlucoseMmol, $sLastGlucoseMmol, $sGlucoseResult, $sLastGlucoseResult, $sGlucoseResultTmp, $sLastGlucoseResultTmp
   If $sReadOption = "mmol" Then
      ; Calculate mmol/l
      $sCalcGlucose = Number(18.01559 * 10 / 10, 3) ; 3=the result is double
	  $sGlucoseMmol = Number($sGlucose / $sCalcGlucose, 3) ; 3=the result is double
	  $sGlucoseResultTmp = Round($sGlucoseMmol, 2) ; Round 0.00
      $sGlucoseResult = Round($sGlucoseResultTmp, 1) ; Round 0.0
      ; Calculate last glucose
      $sLastGlucoseMmol = Number($sLastGlucose / $sCalcGlucose, 3) ; 3=the result is double
	  $sLastGlucoseResultTmp = Round($sLastGlucoseMmol, 2) ; Round 0.00
      $sLastGlucoseResult = Round($sLastGlucoseResultTmp, 1) ; Round 0.0
   Else
      $sGlucoseResult = Int($sGlucose)
      $sLastGlucoseResult = Int($sLastGlucose)
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Glucose result : " & $sGlucoseResult & @CRLF)
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Last glucose result : " & $sLastGlucoseResult & @CRLF)

   ; Calculate delta
   Local $sDeltaTmp = Number($sGlucoseResult - $sLastGlucoseResult, 3) ; 3=the result is double
   Local $sDelta = Round($sDeltaTmp, 1) ; Round 0.0
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Delta : " & $sDelta & @CRLF)
   If $sDelta > 0 Then
      $sDelta = "+" & $sDelta
   EndIf
   If $sDelta = 0 Then
      $sDelta = "±" & $sDelta
   EndIf

   ; Check tooltip alarm
   Local $sAlarm
   If Int($sGlucose) <= Int($sAlertLow) Or Int($sGlucose) >= Int($sAlertHigh) Or Int($sGlucose) <= Int($sAlertLowUrgent) Or Int($sGlucose) >= Int($sAlertHighUrgent) Then
      $sAlarm = "2" ;=Warning icon
   Else
      $sAlarm = "1" ;=Info icon
   EndIf
   If Int($sAlertLow) = 0 And Int($sAlertHigh) = 0 And Int($sAlertLowUrgent) = 0 And Int($sAlertHighUrgent) = 0 Then
      $sAlarm = "0" ;=None icon
   EndIf

   ; Check connections to Nightscout
   Local $sCheckGlucose = StringRegExp($sGlucose, '^([0-9]{1,3})')
   ; Check values for mmol or mg/dl
   Local $checkReadOptionValues = StringRegExp($sReadOption, '(mmol|mg\/dl)')
   If $checkReadOptionValues <> 1 Or $sCheckGlucose <> 1 Or $checkInet <> 1 Then
      ToolTip($wrongMsg, @DesktopWidth - $sInputDesktopWidth, @DesktopHeight - $sInputDesktophHeight, $wrongMsg, 3, 2)
   Else
      ToolTip("   " & $sDelta & " " & @CR & "   " & $sLastReadingGlucoseMin & " min", @DesktopWidth - $sInputDesktopWidth, @DesktopHeight - $sInputDesktophHeight, "   " & $sGlucoseResult & " " & $sTrend & "  ", $sAlarm, 2)
   EndIf

   ; Sleep
   If Mod(@SEC, $sSec) = 0 Then Sleep($sReadInterval)

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
   TrayCreateItem("Help")
   TrayItemSetOnEvent(-1, "_Help")
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

Func _Help()
   ShellExecute("https://github.com/Matze1985/GlucoTT/wiki")
EndFunc

; Set settings in GUI
Func _Settings()
   Local $hSave, $msg
   GUICreate($sIniCategory, 320, 155, @DesktopWidth / 2 - 160, @DesktopHeight / 2 - 45)
   GUICtrlCreateLabel($sIniTitleNightscout, 10, 5, 70)
   Local $hInputDomain = GUICtrlCreateInput($sInputDomain, 10, 20, 300, 20)
   GUICtrlCreateLabel($sIniTitleDesktopWidth, 10, 45, 140)
   Local $hInputDesktopWidth = GUICtrlCreateInput($sInputDesktopWidth, 10, 60, 50, 20, $ES_NUMBER)
   GUICtrlCreateLabel($sIniTitleDesktopHeight, 10, 85, 140)
   Local $hInputDesktopHeight = GUICtrlCreateInput($sInputDesktophHeight, 10, 100, 50, 20, $ES_NUMBER)
   GUICtrlCreateLabel($sIniTitleOptions, 200, 45, 140)
   Local $hCheckboxAutostart = GUICtrlCreateCheckbox($sIniTitleAutostart, 200, 60, 150, 20)
   GUICtrlSetState($hCheckboxAutostart, $sCheckboxAutostart)
   Local $hCheckboxUpdate = GUICtrlCreateCheckbox($sIniTitleUpdate, 200, 80, 150, 20)
   GUICtrlSetState($hCheckboxUpdate, $sCheckboxUpdate)
   $hSave = GUICtrlCreateButton("Save", 10, 125, 300, 20)
   GUISetState()
   $msg = 0
   While $msg <> $GUI_EVENT_CLOSE
      $msg = GUIGetMsg()
      Select
         Case $msg = $hSave
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleNightscout, GUICtrlRead($hInputDomain))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleDesktopWidth, GUICtrlRead($hInputDesktopWidth))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleDesktopHeight, GUICtrlRead($hInputDesktopHeight))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleAutostart, GUICtrlRead($hCheckboxAutostart))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleUpdate, GUICtrlRead($hCheckboxUpdate))
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