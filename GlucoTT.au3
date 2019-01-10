#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
   #AutoIt3Wrapper_Icon=Icon.ico
   #AutoIt3Wrapper_UseX64=n
   #AutoIt3Wrapper_Res_Description=A simple discrete glucose tooltip for Nightscout under Windows
   #AutoIt3Wrapper_Res_Fileversion=3.0.0.5
   #AutoIt3Wrapper_Res_LegalCopyright=Mathias Noack
   #AutoIt3Wrapper_Res_Language=1031
   #AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <TrayConstants.au3>
#include <Array.au3>
#include <Debug.au3>
#include "Include\CheckUpdate.au3"
#include "Include\WinHttp.au3"
#include "Include\ExtMsgBox.au3"
#include "Include\_Startup.au3"
#include "Include\TTS UDF.au3"

Opt("TrayMenuMode", 3) ; The default tray menu items will not be shown and items are not checked when selected. These are options 1 and 2 for TrayMenuMode.
Opt("TrayOnEventMode", 1) ; Enable TrayOnEventMode.
Opt("WinTitleMatchMode", 2) ;1=start, 2=subStr, 3=exact, 4=advanced, -1 to -4=Nocase
Opt("TrayIconHide", 1) ;Hides the AutoIt tray icon.

; Set Hotkeys [CTRL+ALT ...]
HotKeySet("^!n", "_Nightscout")
HotKeySet("^!s", "_Settings")
HotKeySet("^!h", "_Help")
HotKeySet("^!e", "_Exit")

; App title
Local $sTitle = StringRegExpReplace(@ScriptName, ".au3|.exe", "")

; Check for another instance of this program
If _Singleton($sTitle, 1) = 0 Then
   _DebugOut("Another instance is already running")
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Another instance of this program is already running.")
   Exit
EndIf

; Check internet connection on Start/Loop - Return 1 for ON | Return 0 for OFF
Func _CheckConnection()
   Local $iReturn = DllCall("wininet.dll", "int", "InternetGetConnectedState", "int", 0, "int", 0)
   If (@error) Or ($iReturn[0] = 0) Then Return SetError(1, 0, 0)
   Return 1
EndFunc

; Read settings from ini file
Local $sFilePath = @ScriptDir & "\"
Local $sFile = "Settings.ini"
Local $sFileFullPath = $sFilePath & $sFile
Local $sIniCategory = "Settings"
Local $sIniTitleNightscout = "Nightscout"
Local $sIniDefaultNightscout = "https://<account>.herokuapp.com"
Local $sIniTitleCgmUpdate = "GitHub-Nightscout-Update"
Local $sIniTitleGithubAccount = "GitHub-User"
Local $sIniDefaultGithubAccount = "Input a GitHub-User to update cgm-remote-monitor"
Local $iIniDefaultCheckboxCgmUpdate = 4
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
Local $sIniTitleLog = "Log"
Local $iIniDefaultCheckboxTextToSpeech = 4
Local $iIniDefaultCheckboxLog = 4
Local $sIniTitlePlayAlarm = "Play alarm"
Local $iIniDefaultCheckboxPlayAlarm = 4
Local $sInputDomain = IniRead($sFileFullPath, $sIniCategory, $sIniTitleNightscout, $sIniDefaultNightscout)
Local $sInputDesktopWidth = IniRead($sFileFullPath, $sIniCategory, $sIniTitleDesktopWidth, $iIniDefaultDesktopWidth)
Local $sInputDesktophHeight = IniRead($sFileFullPath, $sIniCategory, $sIniTitleDesktopHeight, $iIniDefaultDesktopHeight)
Local $iCheckboxAutostart = IniRead($sFileFullPath, $sIniCategory, $sIniTitleAutostart, $iIniDefaultCheckboxAutostart)
Local $iCheckboxUpdate = IniRead($sFileFullPath, $sIniCategory, $sIniTitleUpdate, $iIniDefaultCheckboxUpdate)
Local $iCheckboxTextToSpeech = IniRead($sFileFullPath, $sIniCategory, $sIniTitleTextToSpeech, $iIniDefaultCheckboxTextToSpeech)
Local $iCheckboxLog = IniRead($sFileFullPath, $sIniCategory, $sIniTitleLog, $iIniDefaultCheckboxLog)
Local $iCheckboxPlayAlarm = IniRead($sFileFullPath, $sIniCategory, $sIniTitlePlayAlarm, $iIniDefaultCheckboxPlayAlarm)
Local $iCheckboxCgmUpdate = IniRead($sFileFullPath, $sIniCategory, $sIniTitleCgmUpdate, $iIniDefaultCheckboxCgmUpdate)
Local $sInputGithubAccount = IniRead($sFileFullPath, $sIniCategory, $sIniTitleGithubAccount, $sIniDefaultGithubAccount)

; Start displaying debug environment
If $iCheckboxLog = 1 Then
   _DebugSetup($sTitle, True, 4, $sTitle & "_" & @YEAR & "-" & @MON & "-" & @MDAY & ".log", True)
EndIf
If Not @Compiled Then
   _DebugSetup($sTitle, True)
EndIf

; Check empty input for Nightscout
If StringRegExp($sInputDomain, '^\s*$') Then
   _DebugOut("Empty domain entered")
   $sInputDomain = $sIniDefaultNightscout
EndIf

; Check empty input for GitHub-User
If StringRegExp($sInputGithubAccount, '^\s*$') Then
   _DebugOut("Empty GitHub user entered")
   $sInputGithubAccount = $sIniDefaultGithubAccount
EndIf

; Check Shortcut Status from ini file
If $iCheckboxAutostart < 4 Then
   _DebugOut("Set autostart in registry")
   _StartupRegistry_Install() ; Add the running EXE to the Current Users Run registry key.
Else
   _DebugOut("Remove autostart from registry")
   _StartupRegistry_Uninstall() ; Remove the running EXE from the Current Users Run registry key.
EndIf

; Check Domain settings
Local $iCheckDomain = StringRegExp($sInputDomain, '^(?:https?:\/\/)?(?:www\.)?([^\s\:\/\?\[\]\@\!\$\&\"\(\)\*\+\,\;\=\<\>\#\%\''\"{"\}\|\\\^\`]{1,63}\.(?:[a-z]{2,}))(?:\/|:[0-9]{1,7}|\?|\&|\s|$)\/?')
If $iCheckDomain = 0 Then
   _DebugOut("Wrong input for domain")
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Wrong URL for Nightscout set!" & @CRLF & @CRLF & "Example:" & @CRLF & "https://account.azurewebsites.net")
   _Settings()
EndIf

; Check upper case in url
If StringRegExp($sInputDomain, '([A-Z])') Then
   _DebugOut("Wrong domain with upper case")
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Upper case in url not allowed!")
   _Settings()
EndIf

; Check desktop width and height settings
Local $iCheckNumbers = StringRegExp($sInputDesktopWidth & $sInputDesktophHeight, '^[0-9]+$')
If $iCheckNumbers = 0 Then
   _DebugOut("Only numbers allowed for Desktop width/height")
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Only numbers allowed in the fields, please check:" & @CRLF & @CRLF & "- Desktop width (minus)" & @CRLF & "- Desktop height (minus)")
   _Settings()
EndIf

; Check checkbox options
If $iCheckboxAutostart = 1 Or $iCheckboxAutostart = 4 And $iCheckboxUpdate = 1 Or $iCheckboxUpdate = 4 And $iCheckboxTextToSpeech = 1 Or $iCheckboxTextToSpeech = 4 And $iCheckboxPlayAlarm = 1 Or $iCheckboxPlayAlarm = 4 And $iCheckboxCgmUpdate = 1 Or $iCheckboxCgmUpdate = 4 Then
Else
   _DebugOut("Only number 1 (on) or 4 (off) for checkboxes allowed")
   ShellExecute(@ScriptDir & "\" & $sFile)
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Only numbers for options allowed, please check your ini file:" & @CRLF & @CRLF & "- Checkbox number: 1 (on)" & @CRLF & "- Checkbox number: 4 (off)")
   WinWaitClose($sFile)
   _Restart()
EndIf

; Check GitHub-Account
Local $iCheckGithub = StringRegExp($sInputGithubAccount, '^[A-Za-z0-9_-]{3,15}$')
If $iCheckGithub = 0 And $iCheckboxCgmUpdate = 1 Then
   _DebugOut("Wrong GitHub Username")
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Wrong GitHub Username!" & @CRLF & @CRLF & "Example:" & @CRLF & "https://github.com/USERNAME" & @CRLF & @CRLF & "For no update check:" & @CRLF & "Disable the option " & $sIniTitleCgmUpdate & "!")
   _Settings()
EndIf

; After closing restarts the app
Func _Restart()
   _DebugOut("Restart app")
   Run(@ComSpec & " /c " & 'TIMEOUT /T 1 & START "" "' & @ScriptFullPath & '"', "", @SW_HIDE)
   Exit
EndFunc

Func _Tooltip()
   ; API-Pages
   Local $sPageJsonEntries = "/api/v1/entries/sgv.json?count=2"
   Local $sPageJsonStatus = "/api/v1/status.json"

   ; Initialize and get session handle and get connection handle
   Global $hOpen = _WinHttpOpen()
   Local $hConnect = _WinHttpConnect($hOpen, $sInputDomain)
   _DebugReportVar("$sInputDomain", $sInputDomain)

   ; Check connection
   Local $iCheckInet = _CheckConnection()
   _DebugReportVar("$iCheckInet", $iCheckInet)

   ; Set wrong msg in tooltip
   Local $sWrongMsg = "✕"

   ; Make a SimpleSSL request
   Local $hRequestPageJsonEntriesSSL = _WinHttpSimpleSendSSLRequest($hConnect, Default, $sPageJsonEntries)
   Local $hRequestPageJsonStatusSSL = _WinHttpSimpleSendSSLRequest($hConnect, Default, $sPageJsonStatus)

   ; Read RequestSSL
   Local $sReturnedPageJsonEntries = _WinHttpSimpleReadData($hRequestPageJsonEntriesSSL)
   Local $sReturnedPageJsonStatus = _WinHttpSimpleReadData($hRequestPageJsonStatusSSL)

   Local $iGlucose = Int(_ArrayToString(StringRegExp($sReturnedPageJsonEntries, '"sgv":([0-9]{1,3}),"', 1))) ; First array result
   _DebugReportVar("$iGlucose", $iGlucose)
   Local $sLastGlucoseValues = _ArrayToString(StringRegExp($sReturnedPageJsonEntries, '"sgv":([0-9]{1,3}),"', 3)) ; Save all array results
   _DebugReportVar("$sLastGlucoseValues", $sLastGlucoseValues)
   Local $iLastGlucose = Int(StringRegExpReplace($sLastGlucoseValues, '.*\|(.*)', '\1')) ; Save 2nd array result
   _DebugReportVar("$iLastGlucose", $iLastGlucose)

   ; Check time readings
   Local $iDate = Int(_ArrayToString(StringRegExp($sReturnedPageJsonEntries, '"date":([0-9]{13}),"', 1)))
   _DebugReportVar("$iDate", $iDate)
   Local $sLastDates = _ArrayToString(StringRegExp($sReturnedPageJsonEntries, '"date":([0-9]{13}),"', 3)) ; Save all array results
   _DebugReportVar("$sLastDates", $sLastDates)
   Local $iLastDate = Int(StringRegExpReplace($sLastDates, '.*\|(.*)', '\1')) ; Save 2nd array result
   _DebugReportVar("$iLastDate", $iLastDate)
   Local $iServerTimeEpoch = Int(_ArrayToString(StringRegExp($sReturnedPageJsonStatus, '"serverTimeEpoch":([0-9]{13}),"', 1)))
   _DebugReportVar("$iServerTimeEpoch", $iServerTimeEpoch)
   Local $fMin = Round(Int($iServerTimeEpoch - $iDate) / 60000, 1)
   _DebugReportVar("$fMin", $fMin)
   Local $iMin = Int($fMin)
   _DebugReportVar("$iMin", $iMin)
   Local $i_fZeroMin = Number(StringMid(Number($fMin - Int($fMin) - 1), 2))
   _DebugReportVar("$i_fZeroMin", $i_fZeroMin)
   Local $iMinInterval = Int(Round(Int($iDate - $iLastDate) / 60000, 0))
   _DebugReportVar("$iMinInterval", $iMinInterval)

   ; Calculate ms for sleep with interval logic
   Local $i_MsWait = Int(Round(60000 * $i_fZeroMin, 0))
   Local $fMinIntervalCheck = Number(0.1) ; Wait 60000 x 0.1 = max. 6000 ms
   _DebugReportVar("$fMinIntervalCheck", $iMinInterval)

   ; Check the first three intervals
   If $iMin == $iMinInterval And $fMin <= $iMinInterval + $fMinIntervalCheck Then
      $i_MsWait = 1000
   EndIf
   If $iMin == 2 * $iMinInterval And $fMin <= 2 * $iMinInterval + $fMinIntervalCheck Then
      $i_MsWait = 1000
   EndIf
   If $iMin == 3 * $iMinInterval And $fMin <= 3 * $iMinInterval + $fMinIntervalCheck Then
      $i_MsWait = 1000
   EndIf
   _DebugReportVar("$i_MsWait", $i_MsWait)

   ; bMod for alarm and speech - modulus operation
   Local $bMod = Mod($iDate, $iServerTimeEpoch)

   ; Check Nightscout version
   Global $sNightscoutVersion = StringRegExpReplace($sReturnedPageJsonStatus, '.*"version":("|)([^"]+)("|),.*', '\2')
   _DebugReportVar("$sNightscoutVersion", $sNightscoutVersion)

   ; Settings check for json: status
   Local $sStatus = StringRegExpReplace($sReturnedPageJsonStatus, '.*"status":("|)([^"]+)("|),.*', '\2')
   _DebugReportVar("$sStatus", $sStatus)

   ; Settings check for json: units
   Local $sReadOption = StringRegExpReplace($sReturnedPageJsonStatus, '.*"units":("|)([^"]+)("|),.*', '\2')
   _DebugReportVar("$sReadOption", $sReadOption)

   ; Alarm settings check for json: alarm bgLow
   Local $iAlertLow = Int(StringRegExpReplace($sReturnedPageJsonStatus, '.*"bgLow":("|)([^"]+)("|),.*', '\2'))
   _DebugReportVar("$iAlertLow", $iAlertLow)

   ; Alarm settings check for json: alarm bgHigh
   Local $iAlertHigh = Int(StringRegExpReplace($sReturnedPageJsonStatus, '.*"bgHigh":("|)([^"]+)("|),.*', '\2'))
   _DebugReportVar("$iAlertHigh", $iAlertHigh)

   ; Alarm settings check for json: alarm UrgentLow for bgTargetBottom
   Local $iAlertLowUrgent = Int(StringRegExpReplace($sReturnedPageJsonStatus, '.*"bgTargetBottom":("|)([^"]+)("|),.*', '\2'))
   _DebugReportVar("$iAlertLowUrgent", $iAlertLowUrgent)

   ; Alarm settings check for json: alarm UrgentHigh for bgTargetTop
   Local $iAlertHighUrgent = Int(StringRegExpReplace($sReturnedPageJsonStatus, '.*"bgTargetTop":("|)([^"]+)("|),.*', '\2'))
   _DebugReportVar("$iAlertHighUrgent", $iAlertHighUrgent)

   ; TrendArrows
   Local $sDirection = _ArrayToString(StringRegExp($sReturnedPageJsonEntries, '"direction":"(.{1,13})","type"', 1))
   _DebugReportVar("$sDirection", $sDirection)

   Local $sTrend
   If StringInStr($sDirection, "DoubleUp") Then
      $sTrend = "⇈"
   EndIf
   If StringInStr($sDirection, "Flat") Then
      $sTrend = "→︎"
   EndIf
   If StringInStr($sDirection, "SingleUp") Then
      $sTrend = "↑"
   EndIf
   If StringInStr($sDirection, "FortyFiveUp") Then
      $sTrend = "↗"
   EndIf
   If StringInStr($sDirection, "FortyFiveDown") Then
      $sTrend = "↘"
   EndIf
   If StringInStr($sDirection, "SingleDown") Then
      $sTrend = "↓"
   EndIf
   If StringInStr($sDirection, "DoubleDown") Then
      $sTrend = "⇊"
   EndIf
   _DebugReportVar("$sTrend", $sTrend)

   ; Check mmol/l option
   Local $fCalcGlucose, $fGlucoseMmol, $fLastGlucoseMmol, $i_fGlucoseResult, $i_fLastGlucoseResult, $fGlucoseResultTmp, $fLastGlucoseResultTmp
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

   _DebugReportVar("$i_fGlucoseResult", $i_fGlucoseResult)
   _DebugReportVar("$i_fLastGlucoseResult", $i_fLastGlucoseResult)

   ; Calculate delta
   Local $fDeltaTmp = Number($i_fGlucoseResult - $i_fLastGlucoseResult, 3) ; 3=the result is double
   Local $s_fDelta = Round($fDeltaTmp, 1) ; Round 0.0
   If $s_fDelta > 0 Then
      $s_fDelta = "+" & $s_fDelta
   EndIf
   If $s_fDelta = 0 Then
      $s_fDelta = "±" & $s_fDelta
   EndIf
   _DebugReportVar("$s_fDelta", $s_fDelta)

   ; Check tooltip alarm
   Local $iAlarm
   If $iGlucose <= $iAlertLow Or $iGlucose >= $iAlertHigh Or $iGlucose <= $iAlertLowUrgent Or $iGlucose >= $iAlertHighUrgent Then
      $iAlarm = 2 ;=Warning icon
      ; Play alarm from Nightscout (alarm.mp3)
      If $iCheckboxPlayAlarm = 1 Then
         If $iMin == 0 And $bMod = True Then
            SoundPlay($sInputDomain & "/audio/alarm.mp3", 0)
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

   ; Tooltip
   If $iCheckReadOptionValues <> 1 Or $iCheckGlucose <> 1 Or $iCheckInet <> 1 Then
      ToolTip($sWrongMsg, @DesktopWidth - $sInputDesktopWidth, @DesktopHeight - $sInputDesktophHeight, $sWrongMsg, 3, 2)
   Else
      ToolTip("   " & $s_fDelta & " " & @CR & "   " & $iMin & " min", @DesktopWidth - $sInputDesktopWidth, @DesktopHeight - $sInputDesktophHeight, "   " & $i_fGlucoseResult & " " & $sTrend & "  ", $iAlarm, 2)
   EndIf

   ; Check TTS option and locals, globals
   Local $oSapi = _SpeechObject_Create()
   Local $sGlucoseTextToSpeech = StringReplace($i_fGlucoseResult, ".", ",")
   Global $sUpdateWindowTitle = "Nightscout-Update"

   ; Sleep
   If $bMod = True Then
      ; Check upload to Nightscout
      If $iDate == $iLastDate And StringRegExp($iDate & $iLastDate, '([0-9]{13})') Then
         _DebugOut("More than one upload method selected")
         _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Please use one upload method to Nightscout!" & @CRLF & @CRLF & "Check your application settings, which transmits the values!" & @CRLF & @CRLF & "Otherwise, " & $sTitle & " does not work properly!")
      EndIf
      If $iCheckboxTextToSpeech = 1 Then
         ; Read every zero minutes
         If $iMin == 0 And $bMod = True Then
            _SpeechObject_Say($oSapi, $sGlucoseTextToSpeech)
         EndIf
      EndIf
      Sleep($i_MsWait)
      ; Check for a cgm-remote-monitor update when the update window not exists (Important: API update 60 requests per hour!)
      If $i_MsWait = 60000 Then
         ; Check for app updates
         If $iCheckboxUpdate = 1 Then
            CheckUpdate($sTitle & (@Compiled ? ".exe" : ".au3"), $sVersion, "https://raw.githubusercontent.com/Matze1985/GlucoTT/master/Update/CheckUpdate.txt")
         EndIf
         If $iCheckboxCgmUpdate = 1 Then
            If Not WinExists($sUpdateWindowTitle) Then
               _CgmUpdateCheck()
            EndIf
         EndIf
      EndIf
   EndIf

   ; Close WinHttp
   _WinHttpCloseHandle($hConnect)

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
   _DebugOut("Open Nightscout")
   ShellExecute($sInputDomain)
EndFunc

Func _Help()
   _DebugOut("Open Help")
   ShellExecute("https://github.com/Matze1985/GlucoTT/wiki")
EndFunc

; Set settings in GUI
Func _Settings()
   _DebugOut("Open Settings")
   Local $hDLL = DllOpen("user32.dll")
   Local $hSave, $hDonate, $msg
   GUICreate($sIniCategory, 320, 190, @DesktopWidth / 2 - 160, @DesktopHeight / 2 - 45)
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
   Local $hCheckboxLog = GUICtrlCreateCheckbox($sIniTitleLog, 160, 120, 60, 20)
   GUICtrlSetState($hCheckboxLog, $iCheckboxLog)
   Local $hCheckboxPlayAlarm = GUICtrlCreateCheckbox($sIniTitlePlayAlarm, 230, 60, 80, 20)
   GUICtrlSetState($hCheckboxPlayAlarm, $iCheckboxPlayAlarm)
   $hDonate = GUICtrlCreateButton("Donate", 230, 80, 80, 40)
   Local $hCheckboxCgmUpdate = GUICtrlCreateCheckbox($sIniTitleCgmUpdate, 10, 120, 145, 20)
   GUICtrlSetState($hCheckboxCgmUpdate, $iCheckboxCgmUpdate)
   Local $hInputGithubAccount = GUICtrlCreateInput($sInputGithubAccount, 10, 140, 300, 20)
   $hSave = GUICtrlCreateButton("Save", 10, 165, 300, 20)
   GUISetState()
   $msg = 0
   While $msg <> $GUI_EVENT_CLOSE
      $msg = GUIGetMsg()
      Select
         Case $msg = $hSave Or _IsPressed("0D", $hDLL)
            _DebugOut("Save settings")
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleNightscout, GUICtrlRead($hInputDomain))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleDesktopWidth, GUICtrlRead($hInputDesktopWidth))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleDesktopHeight, GUICtrlRead($hInputDesktopHeight))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleAutostart, GUICtrlRead($hCheckboxAutostart))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleUpdate, GUICtrlRead($hCheckboxUpdate))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleTextToSpeech, GUICtrlRead($hCheckboxTextToSpeech))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleLog, GUICtrlRead($hCheckboxLog))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitlePlayAlarm, GUICtrlRead($hCheckboxPlayAlarm))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleCgmUpdate, GUICtrlRead($hCheckboxCgmUpdate))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleGithubAccount, GUICtrlRead($hInputGithubAccount))
            _Restart()
         Case $msg = $hDonate
            _DebugOut("Click on Donate")
            ShellExecute("https://www.paypal.me/MathiasN")
         Case $msg = $GUI_EVENT_CLOSE Or _IsPressed("1B", $hDLL)
            _DebugOut("Close settings")
            GUIDelete($sIniCategory)
      EndSelect
   WEnd
EndFunc

Func _CgmUpdateCheck()
   ; Check GitHub update for cgm-remote-monitor
   Local $iRetCheckUpdateValue, $hRequestCgmUpdateCompare, $sOldClip, $sBranch, $sDeployOn
   Local $sCheckGithubStatus = '("status":( |)"(diverged|ahead)")' ; ahead, diverged (update available), behind (after update)
   Local $sCheckUrlAzure = '([Aa][Zz][Uu][Rr][Ee])'
   Local $sCheckUrlHeroku = '([Hh][Ee][Rr][Oo][Kk][Uu])'
   Local $sSubdomainName = StringRegExpReplace($sInputDomain, "(?:http[s]*\:\/\/)|[0-9a-z-]+\.(?:\.)?(com|net)|[.]|[\/]", "")
   _DebugReportVar("$sSubdomainName", $sSubdomainName)
   Local $sUpdateWindowButtons = "Update | Close"
   Local $sUpdateWindowMsg = "Update on GitHub available!"
   Local $sWndGithubTitle = "cgm-remote-monitor"
   Local $sUpdateWindowMsgHelp = "Help for update on GitHub!" & @CRLF & @CRLF & "1. Login " & @CRLF & "2. Create pull request" & @CRLF & "3. Merge and confirm pull request" & @CRLF & "4. Deploy branch on Heruko or Azure" & @CRLF & "5. Close this window, when finished!"
   Local $iCheckVersionDev = StringRegExp($sNightscoutVersion, "(dev)")
   _DebugReportVar("$iCheckVersionDev", $iCheckVersionDev)
   Local $iCheckVersionRelease = StringRegExp($sNightscoutVersion, "(release)")
   _DebugReportVar("$iCheckVersionRelease", $iCheckVersionRelease)
   Local $iCheckGithubAccount = StringRegExp($sInputGithubAccount, "(" & $sIniDefaultGithubAccount & ")")
   _DebugReportVar("$iCheckGithubAccount", $iCheckGithubAccount)

   Local $hConnectCgmUpdateCompare = _WinHttpConnect($hOpen, "https://api.github.com")

   If $iCheckVersionDev = 1 Then
      $hRequestCgmUpdateCompare = _WinHttpSimpleSendSSLRequest($hConnectCgmUpdateCompare, Default, "/repos/" & $sInputGithubAccount & "/cgm-remote-monitor/compare/dev...nightscout:dev")
   EndIf
   If $iCheckVersionRelease = 1 Then
      $hRequestCgmUpdateCompare = _WinHttpSimpleSendSSLRequest($hConnectCgmUpdateCompare, Default, "/repos/" & $sInputGithubAccount & "/cgm-remote-monitor/compare/master...nightscout:master")
   EndIf

   Local $sReturnedCgmUpdateCompare = _WinHttpSimpleReadData($hRequestCgmUpdateCompare)

   ; Check valid GitHub-User with repository
   If StringInStr($sReturnedCgmUpdateCompare, '"message":"Not Found"') Then
      _DebugOut('GitHub-User has no repository "cgm-remote-monitor"')
      _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Input a valid GitHub-User!" & @CRLF & "The cgm remote monitor repository must exist!")
      _Settings()
   EndIf

   ; Write in variable $sDeployOn
   If StringRegExp($sInputDomain, "(azure)") Then
      $sDeployOn = "Azure"
   EndIf
   If StringRegExp($sInputDomain, "(heroku)") Then
      $sDeployOn = "Heroku"
   Else
      $sDeployOn = "Page for Nightscout"
   EndIf

   Local $iCheckMerge = StringRegExp($sReturnedCgmUpdateCompare, $sCheckGithubStatus)
   If $iCheckMerge = 1 Then
      $iRetCheckUpdateValue = _ExtMsgBox($EMB_ICONINFO, $sUpdateWindowButtons, $sIniTitleCgmUpdate, $sUpdateWindowMsg)
      Switch $iRetCheckUpdateValue
         Case 1
            If $iCheckVersionDev = 1 Then
               ShellExecute("https://github.com/" & $sInputGithubAccount & "/cgm-remote-monitor/compare/dev...nightscout:dev")
               $sBranch = "dev"
            EndIf
            If $iCheckVersionRelease = 1 Then
               ShellExecute("https://github.com/" & $sInputGithubAccount & "/cgm-remote-monitor/compare/master...nightscout:master")
               $sBranch = "master"
            EndIf
            _ExtMsgBox($EMB_ICONINFO, "Close", $sIniTitleCgmUpdate, $sUpdateWindowMsgHelp)
            If StringRegExp($sInputDomain, $sCheckUrlAzure) Then
               If WinExists($sWndGithubTitle) Then
                  WinActivate($sWndGithubTitle)
                  Send("^t")
                  Sleep(500)
                  $sOldClip = ClipGet()
                  ClipPut("https://portal.azure.com")
                  Send("^v{ENTER}")
                  ClipPut($sOldClip)
                  Sleep(500)
               EndIf
            EndIf
            If StringRegExp($sInputDomain, $sCheckUrlHeroku) Then
               If WinExists($sWndGithubTitle) Then
                  WinActivate($sWndGithubTitle)
                  Send("^t")
                  Sleep(500)
                  $sOldClip = ClipGet()
                  ClipPut("https://dashboard.heroku.com/apps/" & $sSubdomainName & "/deploy/github")
                  Send("^v{ENTER}")
                  ClipPut($sOldClip)
                  Sleep(500)
               EndIf
            EndIf
            _ExtMsgBox($EMB_ICONINFO, "Close", "Deploy-Nightscout-Update", "Finally, please open " & $sDeployOn & " and login and choose the " & $sBranch & " branch for deploy!")
         Case 2
            ;Exit
      EndSwitch
   EndIf
   _DebugReportVar("$sReturnedCgmUpdateCompare", $sReturnedCgmUpdateCompare)
   _DebugReportVar("$iCheckMerge", $iCheckMerge)
   _WinHttpCloseHandle($hConnectCgmUpdateCompare)
EndFunc

Func _Exit()
   _DebugOut("Close app")
   Exit
EndFunc
