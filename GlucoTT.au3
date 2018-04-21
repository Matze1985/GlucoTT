#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
   #AutoIt3Wrapper_Icon=Icon.ico
   #AutoIt3Wrapper_UseX64=n
   #AutoIt3Wrapper_Res_Description=A simple discrete glucose tooltip for Nightscout under Windows
   #AutoIt3Wrapper_Res_Fileversion=2.3.0.0
   #AutoIt3Wrapper_Res_LegalCopyright=Mathias Noack
   #AutoIt3Wrapper_Res_Language=1031
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <TrayConstants.au3>
#include <Include\CheckUpdate.au3>
#include <Include\WinHttp.au3>
#include <Include\ExtMsgBox.au3>
#include <Include\_Startup.au3>
#include <Include\TTS UDF.au3>

Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Opt("TrayOnEventMode", 1) ; Enable TrayOnEventMode.
Opt("WinTitleMatchMode", 2) ;1=start, 2=subStr, 3=exact, 4=advanced, -1 to -4=Nocase

; App title
Local $sTitle = StringRegExpReplace(@ScriptName, ".au3|.exe", "")

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
Local $iIniDefaultDesktopWidth = 43
Local $sIniTitleDesktopHeight = "Desktop height (minus)"
Local $iIniDefaultDesktopHeight = 68
Local $sIniTitleOptions = "Options"
Local $sIniTitleAutostart = "Autostart"
Local $iIniDefaultCheckboxAutostart = 1
Local $sIniTitleUpdate = "Update"
Local $iIniDefaultCheckboxUpdate = 1
Local $sIniTitleTextToSpeech = "TTS"
Local $iIniDefaultCheckboxTextToSpeech = 4
Local $sIniTitlePlayAlarm = "Play alarm"
Local $iIniDefaultCheckboxPlayAlarm = 4
Local $sInputDomain = IniRead($sFileFullPath, $sIniCategory, $sIniTitleNightscout, $sIniDefaultNightscout)
Local $sInputDesktopWidth = IniRead($sFileFullPath, $sIniCategory, $sIniTitleDesktopWidth, $iIniDefaultDesktopWidth)
Local $sInputDesktophHeight = IniRead($sFileFullPath, $sIniCategory, $sIniTitleDesktopHeight, $iIniDefaultDesktopHeight)
Local $iCheckboxAutostart = IniRead($sFileFullPath, $sIniCategory, $sIniTitleAutostart, $iIniDefaultCheckboxAutostart)
Local $iCheckboxUpdate = IniRead($sFileFullPath, $sIniCategory, $sIniTitleUpdate, $iIniDefaultCheckboxUpdate)
Local $iCheckboxTextToSpeech = IniRead($sFileFullPath, $sIniCategory, $sIniTitleTextToSpeech, $iIniDefaultCheckboxTextToSpeech)
Local $iCheckboxPlayAlarm = IniRead($sFileFullPath, $sIniCategory, $sIniTitlePlayAlarm, $iIniDefaultCheckboxPlayAlarm)

; Check for updates
If $iCheckboxUpdate = 1 Then
   CheckUpdate($sTitle & (@Compiled ? ".exe" : ".au3"), $sVersion, "https://raw.githubusercontent.com/Matze1985/GlucoTT/master/Update/CheckUpdate.txt")
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " @error " & @error & @CRLF)
EndIf

; Check Shortcut state from ini file
If $iCheckboxAutostart < 4 Then
   _StartupRegistry_Install() ; Add the running EXE to the Current Users Run registry key.
Else
   _StartupRegistry_Uninstall() ; Remove the running EXE from the Current Users Run registry key.
EndIf

; Check Domain settings
Local $iCheckDomain = StringRegExp($sInputDomain, '^(?:https?:\/\/)?(?:www\.)?([^\s\:\/\?\[\]\@\!\$\&\"\(\)\*\+\,\;\=\<\>\#\%\''\"{"\}\|\\\^\`]{1,63}\.(?:[a-z]{2,}))(?:\/|:[0-9]{1,7}|\?|\&|\s|$)\/?')
If $iCheckDomain = 0 Then
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Wrong URL for Nightscout set!" & @CRLF & @CRLF & "Example:" & @CRLF & "https://account.azurewebsites.net")
   _Settings()
EndIf

; Check desktop width and height settings
Local $iCheckNumbers = StringRegExp($sInputDesktopWidth & $sInputDesktophHeight, '^[0-9]+$')
If $iCheckNumbers = 0 Then
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Only numbers allowed in the fields, please check:" & @CRLF & @CRLF & "- Desktop width (minus)" & @CRLF & "- Desktop height (minus)")
   _Settings()
EndIf

; Check checkbox options
If $iCheckboxAutostart = 1 Or $iCheckboxAutostart = 4 And $iCheckboxUpdate = 1 Or $iCheckboxUpdate = 4 And $iCheckboxTextToSpeech = 1 Or $iCheckboxTextToSpeech = 4 And $iCheckboxPlayAlarm = 1 Or $iCheckboxPlayAlarm = 4 Then
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
   Local $iCheckInet = _CheckConnection()

   ; Set wrong msg in tooltip
   Local $sWrongMsg = "✕"

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
   Local $iGlucose = Int(StringRegExpReplace($sText, "[^0-9]+", ""))
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Glucose : " & $iGlucose & @CRLF)
   Local $iLastGlucose = Int(StringRegExpReplace($sLastText, "[^0-9]+", ""))

   ; Check time readings
   Local $iYear = Int(StringLeft($sFirstTextLine, 4))
   Local $iMonth = Int(StringMid($sFirstTextLine, 6, 2))
   Local $iDay = Int(StringMid($sFirstTextLine, 9, 2))
   Local $iHour = Int(StringMid($sFirstTextLine, 12, 2))
   Local $iMin = Int(StringMid($sFirstTextLine, 15, 2))
   Local $iSec = Int(StringMid($sFirstTextLine, 18, 2))
   Local $sLastYearMonthDayHourMinSec = $iYear & "/" & $iMonth & "/" & $iDay & " " & $iHour & ":" & $iMin & ":" & $iSec
   Local $sCurrentYearMonthDayHourMinSec = @YEAR & "/" & @MON & "/" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
   Local $bMod = Mod(@SEC, $iSec) = @SEC

   ; Settings check for json: urgentRes
   Local $iReadIntervalMin = Int(StringRegExpReplace($sReturnedJson, '.*"urgentRes":([^"]+),".*', '\1'))
   Local $iCheckReadIntervalMin = StringRegExp($iReadIntervalMin, '([A-Za-z,":-{}])')
   If $iCheckReadIntervalMin = 1 Then
      $iReadIntervalMin = Int(StringRegExpReplace($sReturnedJson, '.*"urgentRes":([^"]+),.*', '\1'))
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Read interval (min): " & $iReadIntervalMin & @CRLF)

   ; Reading last glucose (min)
   Local $iLastReadingGlucoseMin = _DateDiff('n', $sLastYearMonthDayHourMinSec, $sCurrentYearMonthDayHourMinSec)
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Last reading glucose (min) : " & $iLastReadingGlucoseMin & @CRLF)

   ; Interval for loop interval to read glucose
   Local $iReadInterval = 60000
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Read interval (ms) : " & $iReadInterval & @CRLF)

   ; Settings check for json: status
   Local $sStatus = StringRegExpReplace($sReturnedJson, '.*"status":"([^"]+)",.*', '\1')
   Local $iCheckStatus = StringRegExp($sStatus, '([{"A-Z0-9abcdefghijlmnpqrstuvwxyz}:,-.])')
   If $iCheckStatus = 1 Then
      $sStatus = StringRegExpReplace($sReturnedJson, '.*"status":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Status : " & $sStatus & @CRLF)

   ; Settings check for json: units
   Local $sReadOption = StringRegExpReplace($sReturnedJson, '.*"units":"([^"]+)",.*', '\1')
   Local $iCheckReadOption = StringRegExp($sReadOption, '([{"A-Z0-9abcefhijknpqrstuvwxyz}:,-.])')
   If $iCheckReadOption = 1 Then
      $sReadOption = StringRegExpReplace($sReturnedJson, '.*"units":([^"]+),.*', '\1')
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Units : " & $sReadOption & @CRLF)

   ; Alarm settings check for json: alarmLow
   Local $iAlertLow = Int(StringRegExpReplace($sReturnedJson, '.*"alarmLow":"([^"]+)",".*', '\1'))
   Local $iCheckAlertLow = StringRegExp($iAlertLow, '([A-Za-z,":-{}])')
   If $iCheckAlertLow = 1 Then
      $iAlertLow = Int(StringRegExpReplace($sReturnedJson, '.*"alarmLow":([^"]+),.*', '\1'))
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Low : " & $iAlertLow & @CRLF)

   ; Alarm settings check for json: alarmHigh
   Local $iAlertHigh = Int(StringRegExpReplace($sReturnedJson, '.*"alarmHigh":"([^"]+)",".*', '\1'))
   Local $iCheckAlertHigh = StringRegExp($iAlertHigh, '([A-Za-z,":-{}])')
   If $iCheckAlertHigh = 1 Then
      $iAlertHigh = Int(StringRegExpReplace($sReturnedJson, '.*"alarmHigh":([^"]+),.*', '\1'))
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " High : " & $iAlertHigh & @CRLF)

   ; Alarm settings check for json: alarmUrgentLow
   Local $iAlertLowUrgent = Int(StringRegExpReplace($sReturnedJson, '.*"alarmUrgentLow":"([^"]+)",".*', '\1'))
   Local $iCheckAlertLowUrgent = StringRegExp($iAlertLowUrgent, '([A-Za-z,":-{}])')
   If $iCheckAlertLowUrgent = 1 Then
      $iAlertLowUrgent = Int(StringRegExpReplace($sReturnedJson, '.*"alarmUrgentLow":([^"]+),.*', '\1'))
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " UrgentLow : " & $iAlertLowUrgent & @CRLF)

   ; Alarm settings check for json: alarmUrgentHigh
   Local $iAlertHighUrgent = Int(StringRegExpReplace($sReturnedJson, '.*"alarmUrgentHigh":"([^"]+)",.*', '\1'))
   Local $iCheckAlertHighUrgent = StringRegExp($iAlertHighUrgent, '([A-Za-z,":-{}])')
   If $iCheckAlertHighUrgent = 1 Then
      $iAlertHighUrgent = Int(StringRegExpReplace($sReturnedJson, '.*"alarmUrgentHigh":([^"]+),.*', '\1'))
   EndIf
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " UrgentHigh : " & $iAlertHighUrgent & @CRLF)

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
   Local $fCalcGlucose, $fGlucoseMmol, $lastGlucoseMmol, $i_fGlucoseResult, $i_fLastGlucoseResult, $fGlucoseResultTmp, $fLastGlucoseResultTmp
   If $sReadOption = "mmol" Then
      ; Calculate mmol/l
      $fCalcGlucose = Number(18.01559 * 10 / 10, 3) ; 3=the result is double
      $fGlucoseMmol = Number($iGlucose / $fCalcGlucose, 3) ; 3=the result is double
      $fGlucoseResultTmp = Round($fGlucoseMmol, 2) ; Round 0.00
      $i_fGlucoseResult = Round($fGlucoseResultTmp, 1) ; Round 0.0
      ; Calculate last glucose
      $fLastGlucoseMmol = Number($iLastGlucose / $fCalcGlucose, 3) ; 3=the result is double
      $fLastGlucoseResultTmp = Round($fLastGlucoseMmol, 2) ; Round 0.00
      $i_fLastGlucoseResult = Round($fLastGlucoseResultTmp, 1) ; Round 0.0
   Else
      $i_fGlucoseResult = Int($iGlucose)
      $i_fLastGlucoseResult = Int($iLastGlucose)
   EndIf

   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Glucose result : " & $i_fGlucoseResult & @CRLF)
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Last glucose result : " & $i_fLastGlucoseResult & @CRLF)

   ; Calculate delta
   Local $fDeltaTmp = Number($i_fGlucoseResult - $i_fLastGlucoseResult, 3) ; 3=the result is double
   Local $s_fDelta = Round($fDeltaTmp, 1) ; Round 0.0
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Delta : " & $s_fDelta & @CRLF)
   If $s_fDelta > 0 Then
      $s_fDelta = "+" & $s_fDelta
   EndIf
   If $s_fDelta = 0 Then
      $s_fDelta = "±" & $s_fDelta
   EndIf

   ; Check tooltip alarm
   Local $iAlarm
   If $iGlucose <= $iAlertLow Or $iGlucose >= $iAlertHigh Or $iGlucose <= $iAlertLowUrgent Or $iGlucose >= $iAlertHighUrgent Then
      $iAlarm = 2 ;=Warning icon
      ; Play alarm from windows media folder (tada.wav)
      If $iCheckboxPlayAlarm = 1 Then
         If @MIN = @MIN + $iLastReadingGlucoseMin And $bMod = True Then
            SoundPlay(@WindowsDir & "\media\tada.wav", 0)
         EndIf
      EndIf
   Else
      $iAlarm = 1 ;=Info icon
   EndIf
   If $iAlertLow = 0 And $iAlertHigh = 0 And $iAlertLowUrgent = 0 And $iAlertHighUrgent = 0 Then
      $iAlarm = 0 ;=None icon
   EndIf

   ; Check connections to Nightscout
   Local $iCheckGlucose = StringRegExp($iGlucose, '^([0-9]{1,3})')
   ; Check values for mmol or mg/dl
   Local $iCheckReadOptionValues = StringRegExp($sReadOption, '(mmol|mg\/dl)')
   If $iCheckReadOptionValues <> 1 Or $iCheckGlucose <> 1 Or $iCheckInet <> 1 Then
      ToolTip($sWrongMsg, @DesktopWidth - $sInputDesktopWidth, @DesktopHeight - $sInputDesktophHeight, $sWrongMsg, 3, 2)
   Else
      ToolTip("   " & $s_fDelta & " " & @CR & "   " & $iLastReadingGlucoseMin & " min", @DesktopWidth - $sInputDesktopWidth, @DesktopHeight - $sInputDesktophHeight, "   " & $i_fGlucoseResult & " " & $sTrend & "  ", $iAlarm, 2)
   EndIf

   ; Check TTS option
   Local $oSapi = _SpeechObject_Create()
   Local $sGlucoseTextToSpeech = StringReplace($i_fGlucoseResult, ".", ",")

   If $iCheckboxTextToSpeech = 1 Then
      ; Read interval (@MIN + last reading glucose)
      If @MIN = @MIN + $iLastReadingGlucoseMin Then
         _SpeechObject_Say($oSapi, $sGlucoseTextToSpeech)
      EndIf
   EndIf

   ; Sleep
   If $bMod = True Then Sleep($iReadInterval)

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
   GUICtrlCreateLabel($sIniTitleOptions, 160, 45, 140)
   Local $hCheckboxAutostart = GUICtrlCreateCheckbox($sIniTitleAutostart, 160, 60, 60, 20)
   GUICtrlSetState($hCheckboxAutostart, $iCheckboxAutostart)
   Local $hCheckboxUpdate = GUICtrlCreateCheckbox($sIniTitleUpdate, 160, 80, 60, 20)
   GUICtrlSetState($hCheckboxUpdate, $iCheckboxUpdate)
   Local $hCheckboxTextToSpeech = GUICtrlCreateCheckbox($sIniTitleTextToSpeech, 160, 100, 60, 20)
   GUICtrlSetState($hCheckboxTextToSpeech, $iCheckboxTextToSpeech)
   Local $hCheckboxPlayAlarm = GUICtrlCreateCheckbox($sIniTitlePlayAlarm, 230, 60, 80, 20)
   GUICtrlSetState($hCheckboxPlayAlarm, $iCheckboxPlayAlarm)
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
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleTextToSpeech, GUICtrlRead($hCheckboxTextToSpeech))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitlePlayAlarm, GUICtrlRead($hCheckboxPlayAlarm))
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