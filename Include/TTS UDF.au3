#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.10.2
 Author:         Kanashius
 Rights:		 Open Source UDF
 Script Function:
				UDF zum wiedergeben von Text mittels Text-To-Speech
 URL:			http://www.autoitscript.com/forum/index.php?showtopic=173934
#ce ----------------------------------------------------------------------------
#include <WinAPI.au3>

; #CURRENT# =====================================================================================================================
; _SpeechObject_Create
; _SpeechObject_Say
; _SpeechObject_setRate
; _SpeechObject_SetVolume
; _SpeechObject_SetVoice
; _SpeechObject_SetOutput
; _SpeechObject_Pause
; _SpeechObject_Resume
; _SpeechObject_Stop
; _SpeechObject_isReady
; _SpeechObject_getOutputsName
; _SpeechObject_getVoicesName
; ===============================================================================================================================

;===============================================================================
;
; Function Name:    _SpeechObject_Create()
; Description:      Create TTS-Object
; Parameter(s):     none.
; Requirement(s):   none.
; Return Value(s):  Returns an Object
; Author(s):        Kanashius
;
;===============================================================================
Func _SpeechObject_Create()
	$oSpeech=ObjCreate('SAPI.SpVoice')
	if @error then
		return -1
	endif
	$oSpeech.Rate = 1
    $oSpeech.Volume = 100
	return $oSpeech
EndFunc

;===============================================================================
;
; Function Name:    _SpeechObject_Say()
; Description:      Read a text.
; Parameter(s):     $oSpeech - SpeechObejct as returned by _SpeechObject_Create()
;					$sText - String to read
; Requirement(s):   none.
; Return Value(s):  none.
; Author(s):        Kanashius
;
;===============================================================================
Func _SpeechObject_Say($oSpeech,$sText)
	_SpeechObject_Stop($oSpeech)
	$oSpeech.Speak($sText,1)
EndFunc

;===============================================================================
;
; Function Name:    _SpeechObject_setRate()
; Description:      Set Rate of an Speech-Object (reading-speed)
; Parameter(s):     $oSpeech - SpeechObejct as returned by _SpeechObject_Create()
;					$iRate - Int Between -10 and 10
; Requirement(s):   none.
; Return Value(s):  none.
; Author(s):        Kanashius
;
;===============================================================================
Func _SpeechObject_SetRate($oSpeech,$iRate)
	$oSpeech.Rate=$iRate
EndFunc

;===============================================================================
;
; Function Name:    _SpeechObject_SetVolume()
; Description:      Set the Volume of the Speech-Object
; Parameter(s):     $oSpeech - SpeechObejct as returned by _SpeechObject_Create()
;					$iVolume - int between 0 and 100
; Requirement(s):   none.
; Return Value(s):  none.
; Author(s):        Kanashius
;
;===============================================================================
Func _SpeechObject_SetVolume($oSpeech,$iVolume)
	$oSpeech.Volume=$iVolume
EndFunc

;===============================================================================
;
; Function Name:    _SpeechObject_SetVoice()
; Description:      Set the voice of the Speech-Object.
; Parameter(s):     $oSpeech - SpeechObejct as returned by _SpeechObject_Create()
;					$sName - Name (String) of an Voice as returned by _SpeechObject_getVoicesName()
; Requirement(s):   none.
; Return Value(s):  none.
; Author(s):        Kanashius
;
;===============================================================================
Func _SpeechObject_SetVoice($oSpeech,$sName)
	Dim $SOTokens = $oSpeech.GetVoices('', '')
	For $Token In $SOTokens
        if $Token.GetDescription=$sName then
			$oSpeech.Voice=$Token
		endif
	Next
EndFunc

;===============================================================================
;
; Function Name:    _SpeechObject_SetOutput()
; Description:      Set the Output of the Speech-Object.
; Parameter(s):     $oSpeech - SpeechObejct as returned by _SpeechObject_Create()
;					$sName - Name (String) of an Outputdevice as returned by _SpeechObject_getOutputsName()
; Requirement(s):   none.
; Return Value(s):  none.
; Author(s):        Kanashius
;
;===============================================================================
Func _SpeechObject_SetOutput($oSpeech,$sName)
	Dim $SOTokens = $oSpeech.GetAudioOutputs('','')
	For $Token In $SOTokens
        if $Token.GetDescription=$sName then
			$oSpeech.AudioOutput=$Token
		endif
	Next
EndFunc

;===============================================================================
;
; Function Name:    _SpeechObject_Pause()
; Description:      Pauses the Speech-Object while reading.
; Parameter(s):     $oSpeech - SpeechObejct as returned by _SpeechObject_Create()
; Requirement(s):   none.
; Return Value(s):  none.
; Author(s):        Kanashius
;
;===============================================================================
Func _SpeechObject_Pause($oSpeech)
	$oSpeech.Pause()
EndFunc

;===============================================================================
;
; Function Name:    _SpeechObject_Resume()
; Description:      Resumes the Speech-Object when it is paused.
; Parameter(s):     $oSpeech - SpeechObejct as returned by _SpeechObject_Create()
; Requirement(s):   none.
; Return Value(s):  none.
; Author(s):        Kanashius
;
;===============================================================================
Func _SpeechObject_Resume($oSpeech)
	$oSpeech.Resume()
EndFunc

;===============================================================================
;
; Function Name:    _SpeechObject_Stop()
; Description:      Stops an Speech-Obekt while reading
; Parameter(s):     $oSpeech - SpeechObejct as returned by _SpeechObject_Create()
; Requirement(s):   none.
; Return Value(s):  none.
; Author(s):        Kanashius
;
;===============================================================================
Func _SpeechObject_Stop($oSpeech)
    Local $Output = $oSpeech.AudioOutput
    Local $Voice = $oSpeech.Voice
    Local $Rate = $oSpeech.Rate
    Local $Volume = $oSpeech.Volume
    $oSpeech = ObjCreate("SAPI.SpVoice")
    $oSpeech.AudioOutput = $Output
    $oSpeech.Voice = $Voice
    $oSpeech.Rate = $Rate
    $oSpeech.Volume = $Volume
EndFunc

;===============================================================================
;
; Function Name:    _SpeechObject_isReady()
; Description:      Check if the Speech-Object is ready.
; Parameter(s):     $oSpeech - SpeechObejct as returned by _SpeechObject_Create()
; Requirement(s):   none.
; Return Value(s):  true - if Speech-Object is ready
;					false - if Speech-Object is reading
; Author(s):        Kanashius
;
;===============================================================================
Func _SpeechObject_isReady($oSpeech)
	if _WinAPI_WaitForSingleObject($oSpeech.SpeakCompleteEvent,0)<>258 then
		return true
	endif
EndFunc

;===============================================================================
;
; Function Name:    _SpeechObject_getOutputsName()
; Description:      Return the Names of all avaible AudioOutput-Devices
; Parameter(s):     $oSpeech - SpeechObejct as returned by _SpeechObject_Create()
;					$bAsArray - true if names have to be returned as an Array
; Requirement(s):   none.
; Return Value(s):  String where the Devicenames are seperated with a "|"
;					or Array with Names.
; Author(s):        Kanashius
;
;===============================================================================
Func _SpeechObject_getOutputsName($oSpeech,$bAsArray=false)
	Dim $SOTokens = $oSpeech.GetAudioOutputs('','')
	$sString=""
	For $Token In $SOTokens
        $sString&="|"&$Token.GetDescription
	Next
	$sString=StringTrimLeft($sString,1)
	if $bAsArray then
		return StringSplit($sString,"|",2)
	else
		return $sString
	endif
EndFunc

;===============================================================================
;
; Function Name:    _SpeechObject_getVoicesName()
; Description:      Return the Names of all avaible Voices
; Parameter(s):     $oSpeech - SpeechObejct as returned by _SpeechObject_Create()
;					$bAsArray - true if names have to be returned as an Array
; Requirement(s):   none.
; Return Value(s):  String where the Voicenames are seperated with a "|"
;					or Array with Names.
; Author(s):        Kanashius
;
;===============================================================================
Func _SpeechObject_getVoicesName($oSpeech,$bAsArray=false)
	Dim $SOTokens = $oSpeech.GetVoices('', '')
	$sString=""
	For $Token In $SOTokens
        $sString&="|"&$Token.GetDescription
	Next
	$sString=StringTrimLeft($sString,1)
	if $bAsArray then
		return StringSplit($sString,"|",2)
	else
		return $sString
	endif
EndFunc