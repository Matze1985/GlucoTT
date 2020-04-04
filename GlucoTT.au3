#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
   #AutoIt3Wrapper_Icon=Icon.ico
   #AutoIt3Wrapper_UseX64=n
   #AutoIt3Wrapper_Res_Description=GlucoTT
   #AutoIt3Wrapper_Res_Fileversion=3.5.0.5
   #AutoIt3Wrapper_Res_LegalCopyright=Mathias Noack
   #AutoIt3Wrapper_Res_Language=1031
   #AutoIt3Wrapper_Run_Tidy=y
;~    #AutoIt3Wrapper_Run_Au3Stripper=y
   #Au3Stripper_Parameters=/mo /SCI=1
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <TrayConstants.au3>
#include <Array.au3>
#include <Debug.au3>
#include "Include\CheckUpdate.au3"
#include "Include\WinHttp.au3"
#include "Include\ExtMsgBox.au3"
#include "Include\_Startup.au3"
#include "Include\TTS UDF.au3"
#include "Include\_AudioEndpointVolume.au3"

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

; Check internet connection on Start/Loop - Return 1 for ON | Return 0 for OFF
Func _CheckConnection()
   Local $iReturn = DllCall("wininet.dll", "int", "InternetGetConnectedState", "int", 0, "int", 0)
   If (@error) Or ($iReturn[0] = 0) Then Return SetError(1, 0, 0)
   Return 1
EndFunc

; Read settings from ini file
Local $sFilePath = @ScriptDir & "\"
Local $sFileSettings = "Settings.ini"
Local $sFileFullPath = $sFilePath & $sFileSettings
Local $sIniCategory = "Settings"
Local $sIniTitleNightscout = "Nightscout"
Local $sIniDefaultNightscout = "<account>.herokuapp.com"
Local $sIniTitleCgmUpdate = "GitHub-Nightscout-Update"
Local $sIniTitleGithubAccount = "GitHub-User"
Local $sIniDefaultGithubAccount = "Input a GitHub-User to update cgm-remote-monitor"
Local $sIniTitleCheckboxToken = "Nightscout-Token"
Local $sIniTitleInputToken = "Token of Nightscout"
Local $iIniDefaultCheckboxToken = 4
Local $sIniDefaultInputToken = "Token for Nightscout"
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
Local $sIniTitleTrayTip = "TrayTip"
Local $iIniDefaultCheckboxTrayTip = 4
Local $sInputDomain = IniRead($sFileFullPath, $sIniCategory, $sIniTitleNightscout, $sIniDefaultNightscout)
Local $sInputDesktopWidth = IniRead($sFileFullPath, $sIniCategory, $sIniTitleDesktopWidth, $iIniDefaultDesktopWidth)
Local $sInputDesktophHeight = IniRead($sFileFullPath, $sIniCategory, $sIniTitleDesktopHeight, $iIniDefaultDesktopHeight)
Local $sInputGithubAccount = IniRead($sFileFullPath, $sIniCategory, $sIniTitleGithubAccount, $sIniDefaultGithubAccount)
Local $sInputToken = IniRead($sFileFullPath, $sIniCategory, $sIniTitleInputToken, $sIniDefaultInputToken)
Local $iCheckboxAutostart = IniRead($sFileFullPath, $sIniCategory, $sIniTitleAutostart, $iIniDefaultCheckboxAutostart)
Local $iCheckboxUpdate = IniRead($sFileFullPath, $sIniCategory, $sIniTitleUpdate, $iIniDefaultCheckboxUpdate)
Local $iCheckboxTextToSpeech = IniRead($sFileFullPath, $sIniCategory, $sIniTitleTextToSpeech, $iIniDefaultCheckboxTextToSpeech)
Local $iCheckboxLog = IniRead($sFileFullPath, $sIniCategory, $sIniTitleLog, $iIniDefaultCheckboxLog)
Local $iCheckboxPlayAlarm = IniRead($sFileFullPath, $sIniCategory, $sIniTitlePlayAlarm, $iIniDefaultCheckboxPlayAlarm)
Local $iCheckboxTrayTip = IniRead($sFileFullPath, $sIniCategory, $sIniTitleTrayTip, $iIniDefaultCheckboxTrayTip)
Local $iCheckboxCgmUpdate = IniRead($sFileFullPath, $sIniCategory, $sIniTitleCgmUpdate, $iIniDefaultCheckboxCgmUpdate)
Local $iCheckboxToken = IniRead($sFileFullPath, $sIniCategory, $sIniTitleCheckboxToken, $iIniDefaultCheckboxToken)

; Start displaying debug environment
Local $sDebugAction = "@@ Debug : Action -> "
Local $sDebugInfo = "@@ Debug : Info -> "
Local $sFileLog = $sTitle & "@Debug.log"
Local $sFileLogFullPath = @DesktopDir & "\" & $sFileLog

; Check existing log
If $iCheckboxLog == 1 Then
   _DebugSetup($sTitle, True, 4, $sFileLogFullPath, True)
   _DebugOut($sDebugInfo & "Writing in log -> " & $sFileLogFullPath)
EndIf
If Not @Compiled Then
   _DebugSetup($sTitle, True)
EndIf

; Check for another instance of this program
If _Singleton($sTitle, 1) = 0 Then
   _DebugOut($sDebugInfo & "Another instance is already running")
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Another instance of this program is already running.")
   Exit
EndIf

; Check empty input for Nightscout
If StringRegExp($sInputDomain, '^\s*$') Then
   _DebugOut($sDebugInfo & "Empty domain entered")
   $sInputDomain = $sIniDefaultNightscout
EndIf

; Check empty input for GitHub-User
If StringRegExp($sInputGithubAccount, '^\s*$') Then
   _DebugOut($sDebugInfo & "Empty GitHub user entered")
   $sInputGithubAccount = $sIniDefaultGithubAccount
EndIf

; Check Shortcut Status from ini file
If $iCheckboxAutostart < 4 Then
   _DebugOut($sDebugInfo & "Set autostart in registry")
   _StartupRegistry_Install() ; Add the running EXE to the Current Users Run registry key.
Else
   _DebugOut($sDebugInfo & "Remove autostart from registry")
   _StartupRegistry_Uninstall() ; Remove the running EXE from the Current Users Run registry key.
EndIf

; Check Domain settings
Local $iCheckDomain = StringRegExp($sInputDomain, '^(?:https?:\/\/)?(?:www\.)?([^\s\:\/\?\[\]\@\!\$\&\"\(\)\*\+\,\;\=\<\>\#\%\''\"{"\}\|\\\^\`]{1,63}\.(?:[a-z]{2,}))(?:\/|:[0-9]{1,7}|\?|\&|\s|$)\/?')
If $iCheckDomain == 0 Then
   _DebugOut($sDebugInfo & "Wrong input for domain")
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Wrong URL for Nightscout set!" & @CRLF & @CRLF & "Example:" & @CRLF & "https://account.azurewebsites.net")
   _Settings()
EndIf

; Check upper case in url
If StringRegExp($sInputDomain, '([A-Z])') Then
   _DebugOut($sDebugInfo & "Wrong domain with upper case")
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Upper case in url not allowed!")
   _Settings()
EndIf

; Check desktop width and height settings
Local $iCheckNumbers = StringRegExp($sInputDesktopWidth & $sInputDesktophHeight, '^[0-9]+$')
If $iCheckNumbers == 0 Then
   _DebugOut($sDebugInfo & "Only numbers allowed for Desktop width/height")
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Only numbers allowed in the fields, please check:" & @CRLF & @CRLF & "- Desktop width (minus)" & @CRLF & "- Desktop height (minus)")
   _Settings()
EndIf

; Check checkbox options
If StringRegExp($iCheckboxAutostart & $iCheckboxUpdate & $iCheckboxTextToSpeech & $iCheckboxPlayAlarm & $iCheckboxTrayTip & $iCheckboxCgmUpdate & $iCheckboxToken, '1|4') Then
   _DebugOut($sDebugInfo & "Checkbox options ok")
Else
   _DebugOut($sDebugInfo & "Only number 1 (on) or 4 (off) for checkboxes allowed")
   ShellExecute(@ScriptDir & "\" & $sFileSettings)
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Only numbers for options allowed, please check your ini file:" & @CRLF & @CRLF & "- Checkbox number: 1 (on)" & @CRLF & "- Checkbox number: 4 (off)")
   WinWaitClose($sFileSettings)
   _Restart()
EndIf

; Check GitHub-Account
Local $iCheckGithub = StringRegExp($sInputGithubAccount, '^[A-Za-z0-9_-]{3,15}$')
If $iCheckGithub == 0 And $iCheckboxCgmUpdate == 1 Then
   _DebugOut($sDebugAction & "Wrong GitHub Username")
   _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Wrong GitHub Username!" & @CRLF & @CRLF & "Example:" & @CRLF & "https://github.com/USERNAME" & @CRLF & @CRLF & "For no update check:" & @CRLF & "Disable the option " & $sIniTitleCgmUpdate & "!")
   _Settings()
EndIf

; After closing restarts the app
Func _Restart()
   _DebugOut($sDebugAction & "Restart app")
   Run(@ComSpec & " /c " & 'TIMEOUT /T 1 & START "" "' & @ScriptFullPath & '"', "", @SW_HIDE)
   Exit
EndFunc

Func _Tooltip()

   ; API-Pages
   Local $sApiSecret
   Local $sPageJsonEntries = "api/v1/entries/sgv.json?count=2"
   Local $sPageJsonStatus = "api/v1/status.json"

   ; Check checkbox of Token
   If $iCheckboxToken <> 1 Then
      $sApiSecret = ""
   Else
      $sApiSecret = "API-SECRET: " & $sInputToken
   EndIf

   ; Initialize and get session handle and get connection handle
   Global $hOpen = _WinHttpOpen()
   Global $hConnect = _WinHttpConnect($hOpen, StringRegExpReplace($sInputDomain, "http[s]:\/\/", "")) ; Remove http[s]

   Local $hRequestJsonStatus = _WinHttpOpenRequest($hConnect, Default, $sPageJsonStatus)
   _WinHttpAddRequestHeaders($hRequestJsonStatus, $sApiSecret)
   Local $hRequestJsonEntries = _WinHttpOpenRequest($hConnect, Default, $sPageJsonEntries)
   _WinHttpAddRequestHeaders($hRequestJsonEntries, $sApiSecret)

   ; Send request
   _WinHttpSendRequest($hRequestJsonStatus)
   _WinHttpSendRequest($hRequestJsonEntries)

   ; Wait for the response
   _WinHttpReceiveResponse($hRequestJsonStatus)
   _WinHttpReceiveResponse($hRequestJsonEntries)

   ; Read RequestSSL
   Local $sReturnedPageJsonEntries = _WinHttpSimpleReadData($hRequestJsonEntries)
   Local $sReturnedPageJsonStatus = _WinHttpSimpleReadData($hRequestJsonStatus)

   _DebugReportVar("$sReturnedPageJsonEntries", $sReturnedPageJsonEntries)
   _DebugReportVar("$sReturnedPageJsonStatus", $sReturnedPageJsonStatus)

   If StringInStr($sReturnedPageJsonEntries & $sReturnedPageJsonStatus, '"status":401,"message":"Unauthorized"') Then
      _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Please enable, insert or check your API-SECRET-Token!")
      _Settings()
   EndIf

   _DebugReportVar("$sInputDomain", $sInputDomain)

   ; Check connection
   Local $iCheckInet = _CheckConnection()
   _DebugReportVar("$iCheckInet", $iCheckInet)

   ; Set wrong msg in tooltip
   Local $sWrongMsg = "✕"

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
   Local $i_MsWaitDiff = Int(Round(60000 * $i_fZeroMin, 0))
   Local $i_MsWait = $i_MsWaitDiff
   Local $fMinIntervalCheck = Number(0.1) ; Wait 60000 x 0.1 = max. 6000 ms
   _DebugReportVar("$fMinIntervalCheck", $fMinIntervalCheck)

   ; Calculate all intervals
   Local $iIntervalRound
   Local $fIntervalRound = Number($iMin / $iMinInterval)
   _DebugReportVar("$fIntervalRound", $fIntervalRound)
   If StringRegExp($fIntervalRound, '(.)') Then
      $iIntervalRound = Int($fIntervalRound + 1)
   Else
      $iIntervalRound = $iMinInterval
   EndIf
   If StringRegExp($fIntervalRound, '(#INF)') Then
      $iIntervalRound = Int(2)
   EndIf
   _DebugReportVar("$iIntervalRound", $iIntervalRound)

   ; Check new glucose updates
   Local $bUpdateWait = False
   If $iMin == $iIntervalRound * $iMinInterval - $iMinInterval And $fMin <= $iIntervalRound * $iMinInterval - $iMinInterval + $fMinIntervalCheck Then
      $i_MsWait = 1000
      $bUpdateWait = True
   Else
      $bUpdateWait = False
   EndIf

   ; After 0 min read normal wait
   If $iMin == 0 Then
      $i_MsWait = $i_MsWaitDiff
   EndIf
   _DebugReportVar("$i_MsWait", $i_MsWait)

   ; Check Nightscout version
   Global $sNightscoutVersion = StringRegExpReplace($sReturnedPageJsonStatus, '.*"version":("|)([^"]+)("|),.*', '\2')
   _DebugReportVar("$sNightscoutVersion", $sNightscoutVersion)

   ; Settings check for json: status
   Local $sStatus = StringRegExpReplace($sReturnedPageJsonStatus, '.*"status":("|)([^"]+)("|),.*', '\2')
   _DebugReportVar("$sStatus", $sStatus)

   ; Name for nightscout
   Local $sCustomTitle = StringRegExpReplace($sReturnedPageJsonStatus, '.*"customTitle":("|)([^"]+)("|),.*', '\2')
   _DebugReportVar("$sCustomTitle", $sCustomTitle)

   ; Settings check for json: units
   Local $sUnit = StringRegExpReplace($sReturnedPageJsonStatus, '.*"units":("|)([^"]+)("|),.*', '\2')
   _DebugReportVar("$sUnit", $sUnit)

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
   If $sUnit == "mmol" Then
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
   If $i_fGlucoseResult < $i_fLastGlucoseResult Then
      If StringInStr($sDirection, "FortyFiveUp") Or StringInStr($sDirection, "SingleUp") Or StringInStr($sDirection, "DoubleUp") Then
         $fDeltaTmp = Number(StringReplace($fDeltaTmp, "-", ""), 3) ; 3=the result is double
      EndIf
   EndIf

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
      If $iCheckboxPlayAlarm == 1 Then
         If $iMin == 0 Then
            SoundPlay($sInputDomain & "/audio/alarm.mp3", 0)
         EndIf
      EndIf
   Else
      $iAlarm = 1 ;=Info icon
   EndIf
   If $iAlertLow == 0 And $iAlertHigh == 0 And $iAlertLowUrgent == 0 And $iAlertHighUrgent == 0 Then
      $iAlarm = 0 ;=None icon
   EndIf

   ; Check connections to Nightscout
   Local $iCheckGlucose = StringRegExp($iGlucose, '^([0-9]{1,3})')
   ; Check values for mmol or mg/dl
   Local $iCheckReadOptionValues = StringRegExp($sUnit, '(mmol|mg\/dl)')

   ; TrayTip variables
   Local $sTrayTipMsg = $i_fGlucoseResult & " " & $sUnit & "  •  " & $sTrend & "  •  " & $s_fDelta & "  •  " & $iMin & " min"

   If $iCheckboxTrayTip == 1 Then
      If $iMin == 0 Then
         _DebugOut($sDebugInfo & "Display -> TrayTip")
         If $iCheckReadOptionValues <> 1 Or $iCheckGlucose <> 1 Or $iCheckInet <> 1 Then
            TrayTip($sCustomTitle, $sWrongMsg, 10, 3 + $TIP_NOSOUND)
         Else
            TrayTip($sCustomTitle, $sTrayTipMsg, 10, $iAlarm + $TIP_NOSOUND)
         EndIf
      EndIf
      If $iMin == $iMinInterval * $iIntervalRound - $iMinInterval And $iMin <> 0 And $bUpdateWait == False Then
         TrayTip($sCustomTitle, $sTrayTipMsg, 10, 2 + $TIP_NOSOUND)
      EndIf
   Else
      _DebugOut($sDebugInfo & "Display -> ToolTip")
      If $iCheckReadOptionValues <> 1 Or $iCheckGlucose <> 1 Or $iCheckInet <> 1 Then
         ToolTip($sWrongMsg, @DesktopWidth - $sInputDesktopWidth, @DesktopHeight - $sInputDesktophHeight, $sWrongMsg, 3, 2)
      Else
         ToolTip("   " & $s_fDelta & " " & @CR & "   " & $iMin & " min", @DesktopWidth - $sInputDesktopWidth, @DesktopHeight - $sInputDesktophHeight, "   " & $i_fGlucoseResult & " " & $sTrend & "  ", $iAlarm, 2)
      EndIf
   EndIf

   ; Check TTS option and locals, globals
   Local $oSapi = _SpeechObject_Create()
   Local $sGlucoseTextToSpeech = StringReplace($i_fGlucoseResult, ".", ",")
   Global $sUpdateWindowTitle = "Nightscout-Update"

   ; Check upload to Nightscout
   If $iDate == $iLastDate And StringRegExp($iDate & $iLastDate, '([0-9]{13})') Then
      _DebugOut($sDebugInfo & "More than one upload method selected")
      _ExtMsgBox($MB_ICONERROR, $MB_OK, "Error", "Please use one upload method to Nightscout!" & @CRLF & @CRLF & "Check your application settings, which transmits the values!" & @CRLF & @CRLF & "Otherwise, " & $sTitle & " does not work properly!")
   EndIf
   _DebugReportVar("_GetMute()", _GetMute())
   If $iCheckboxTextToSpeech == 1 And _GetMute() == 0 Then
      ; Read every zero minutes
      If $iMin == 0 Then
         If StringInStr($sGlucoseTextToSpeech, "-") Then
            ; Speak no negative values
         Else
            _SpeechObject_Say($oSapi, $sGlucoseTextToSpeech)
         EndIf
      EndIf
   EndIf

   ; Sleep
   Sleep($i_MsWait)

;~       ; Remove temporary files after 60 seconds
;~       Local $sFileProc = "proc.cmd"
;~       If FileExists($sFileProc) Then
;~          _DebugOut($sDebugInfo & 'Remove temp. file "' & $sFileProc & '"')
;~          FileDelete($sFileProc)
;~       EndIf
   ; Check for a cgm-remote-monitor update when the update window not exists (Important: API update 60 requests per hour!)
   If $i_MsWait == 60000 Then
      ; Check for app updates
      If $iCheckboxUpdate == 1 Then
         CheckUpdate($sTitle & (@Compiled ? ".exe" : ".au3"), $sVersion, "https://raw.githubusercontent.com/Matze1985/GlucoTT/master/Update/CheckUpdate.txt")
      EndIf
      If $iCheckboxCgmUpdate == 1 Then
         If Not WinExists($sUpdateWindowTitle) Then
            _CgmUpdateCheck()
         EndIf
      EndIf
   EndIf

   ; Close handles
   _WinHttpCloseHandle($hRequestJsonEntries)
   _WinHttpCloseHandle($hRequestJsonStatus)
   _WinHttpCloseHandle($hConnect)
   _WinHttpCloseHandle($hOpen)

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
   _DebugOut($sDebugAction & "TrayMenu -> Nightscout")
   ShellExecute($sInputDomain)
EndFunc

Func _Help()
   _DebugOut($sDebugAction & "TrayMenu -> Help")
   ShellExecute("https://github.com/Matze1985/GlucoTT/wiki")
EndFunc

Func _Exit()
   _DebugOut($sDebugAction & "TrayMenu -> Close app")
   Exit
EndFunc

; Set settings in GUI
Func _Settings()
   _DebugOut($sDebugAction & "Settings -> Open")
   Local $hDLL = DllOpen("user32.dll")
   Local $hSave, $hDonate, $msg
   GUICreate($sIniCategory, 320, 230, @DesktopWidth / 2 - 160, @DesktopHeight / 2 - 45)
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
   Local $hCheckboxTrayTip = GUICtrlCreateCheckbox($sIniTitleTrayTip, 230, 80, 80, 20)
   GUICtrlSetState($hCheckboxTrayTip, $iCheckboxTrayTip)
   $hDonate = GUICtrlCreateButton("Donate", 230, 100, 80, 35)
   Local $hCheckboxCgmUpdate = GUICtrlCreateCheckbox($sIniTitleCgmUpdate, 10, 120, 145, 20)
   GUICtrlSetState($hCheckboxCgmUpdate, $iCheckboxCgmUpdate)
   Local $hInputGithubAccount = GUICtrlCreateInput($sInputGithubAccount, 10, 140, 300, 20)
   Local $hCheckboxToken = GUICtrlCreateCheckbox($sIniTitleCheckboxToken, 10, 160, 145, 20)
   GUICtrlSetState($hCheckboxToken, $iCheckboxToken)
   Local $hInputToken = GUICtrlCreateInput($sInputToken, 10, 180, 300, 20)
   $hSave = GUICtrlCreateButton("Save", 10, 205, 300, 20)
   GUISetState()
   $msg = 0
   While $msg <> $GUI_EVENT_CLOSE
      $msg = GUIGetMsg()
      Select
         Case $msg = $hSave Or _IsPressed("0D", $hDLL)
            _DebugOut($sDebugAction & "Settings -> Save")
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleNightscout, GUICtrlRead($hInputDomain))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleDesktopWidth, GUICtrlRead($hInputDesktopWidth))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleDesktopHeight, GUICtrlRead($hInputDesktopHeight))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleAutostart, GUICtrlRead($hCheckboxAutostart))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleUpdate, GUICtrlRead($hCheckboxUpdate))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleTextToSpeech, GUICtrlRead($hCheckboxTextToSpeech))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleLog, GUICtrlRead($hCheckboxLog))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitlePlayAlarm, GUICtrlRead($hCheckboxPlayAlarm))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleTrayTip, GUICtrlRead($hCheckboxTrayTip))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleCgmUpdate, GUICtrlRead($hCheckboxCgmUpdate))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleGithubAccount, GUICtrlRead($hInputGithubAccount))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleCheckboxToken, GUICtrlRead($hCheckboxToken))
            IniWrite($sFileFullPath, $sIniCategory, $sIniTitleInputToken, GUICtrlRead($hInputToken))
            _Restart()
         Case $msg = $hDonate
            _DebugOut($sDebugAction & "Settings -> Donate")
            ShellExecute("https://www.paypal.me/MathiasN")
         Case $msg = $GUI_EVENT_CLOSE Or _IsPressed("1B", $hDLL)
            _DebugOut($sDebugAction & "Settings -> Close")
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
   Local $sGithubApiUrl = "https://api.github.com"
   Local $iCheckGithubAccount = StringRegExp($sInputGithubAccount, "(" & $sIniDefaultGithubAccount & ")")
   _DebugReportVar("$iCheckGithubAccount", $iCheckGithubAccount)

   ; Repo informations --> https://api.github.com/repos/<User>/cgm-remote-monitor
   Local $hConnectCgmToRepo = _WinHttpConnect($hOpen, $sGithubApiUrl)
   $hRequestCgmToRepo = _WinHttpSimpleSendSSLRequest($hConnectCgmToRepo, Default, "/repos/" & $sInputGithubAccount & "/cgm-remote-monitor")
   Local $sReturnedCgmToRepo = _WinHttpSimpleReadData($hRequestCgmToRepo)
   _DebugReportVar("$sReturnedCgmToRepo", $sReturnedCgmToRepo)
   Local $sGithubDefaultBranch = StringRegExpReplace(_ArrayToString(StringRegExp($sReturnedCgmToRepo, '"default_branch":(.*?),', 1)), ' |"', "") ; default branch
   _DebugReportVar("$sGithubDefaultBranch", $sGithubDefaultBranch)
   _WinHttpCloseHandle($hConnectCgmToRepo)

   ; Check update for branch
   ; Example: https://github.com/<User>/cgm-remote-monitor/compare/dev...nightscout:dev
   ; Example: https://api.github.com/repos/<User>/cgm-remote-monitor/compare/dev...nightscout:dev
   Local $hConnectCgmUpdateCompare = _WinHttpConnect($hOpen, $sGithubApiUrl)
   Local $hRequestCgmUpdateCompare = _WinHttpSimpleSendSSLRequest($hConnectCgmUpdateCompare, Default, "/repos/" & $sInputGithubAccount & "/cgm-remote-monitor/compare/" & $sGithubDefaultBranch & "...nightscout:" & $sGithubDefaultBranch)
   Local $sReturnedCgmUpdateCompare = _WinHttpSimpleReadData($hRequestCgmUpdateCompare)
   _DebugReportVar("$sReturnedCgmUpdateCompare", $sReturnedCgmUpdateCompare)
   _WinHttpCloseHandle($hConnectCgmUpdateCompare)

   ; Check valid GitHub-User with repository
   If StringInStr($sReturnedCgmUpdateCompare, '"message":"Not Found"') Then
      _DebugOut($sDebugInfo & 'GitHub-User has no repository "cgm-remote-monitor"')
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
   If $iCheckMerge == 1 Then
      $iRetCheckUpdateValue = _ExtMsgBox($EMB_ICONINFO, $sUpdateWindowButtons, $sIniTitleCgmUpdate, $sUpdateWindowMsg)
      Switch $iRetCheckUpdateValue
         Case 1
            ShellExecute("https://github.com/" & $sInputGithubAccount & "/cgm-remote-monitor/compare/" & $sGithubDefaultBranch & "...nightscout:" & $sGithubDefaultBranch)
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
            _ExtMsgBox($EMB_ICONINFO, "Close", "Deploy-Nightscout-Update", "Finally, please open " & $sDeployOn & " and login and choose the " & $sGithubDefaultBranch & " branch for deploy!")
         Case 2
            ;Exit
      EndSwitch
   EndIf
   _DebugReportVar("$sReturnedCgmUpdateCompare", $sReturnedCgmUpdateCompare)
   _DebugReportVar("$iCheckMerge", $iCheckMerge)
EndFunc
