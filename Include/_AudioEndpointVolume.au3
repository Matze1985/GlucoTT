#include-once

; ===============================================================================================================================
; Title:	_AudioEndpointVolume
; Author:	Erik Pilsits
; Version:	1.0.0.0
; ===============================================================================================================================

Global Const $CLSID_MMDeviceEnumerator = "{BCDE0395-E52F-467C-8E3D-C4579291692E}"
Global Const $IID_IMMDeviceEnumerator = "{A95664D2-9614-4F35-A746-DE8DB63617E6}"
Global Const $tagIMMDeviceEnumerator = _
	"EnumAudioEndpoints hresult(int;dword;ptr*);" & _
	"GetDefaultAudioEndpoint hresult(int;int;ptr*);" & _
	"GetDevice hresult(wstr;ptr*);" & _
	"RegisterEndpointNotificationCallback hresult(ptr);" & _
	"UnregisterEndpointNotificationCallback hresult(ptr)"

Global Const $IID_IMMDevice = "{D666063F-1587-4E43-81F1-B948E807363F}"
Global Const $tagIMMDevice = _
	"Activate hresult(struct*;dword;ptr;ptr*);" & _
	"OpenPropertyStore hresult(dword;ptr*);" & _
	"GetId hresult(wstr*);" & _
	"GetState hresult(dword*)"
Global Const $eRender = 0, $eConsole = 0

Global Const $IID_IAudioEndpointVolume = "{5CDF2C82-841E-4546-9722-0CF74078229A}"
Global Const $tagIAudioEndpointVolume = _
	"RegisterControlChangeNotify hresult(ptr);" & _
	"UnregisterControlChangeNotify hresult(ptr);" & _
	"GetChannelCount hresult(uint*);" & _
	"SetMasterVolumeLevel hresult(float;ptr);" & _
	"SetMasterVolumeLevelScalar hresult(float;ptr);" & _
	"GetMasterVolumeLevel hresult(float*);" & _
	"GetMasterVolumeLevelScalar hresult(float*);" & _
	"SetChannelVolumeLevel hresult(uint;float;ptr);" & _
	"SetChannelVolumeLevelScalar hresult(uint;float;ptr);" & _
	"GetChannelVolumeLevel hresult(uint;float*);" & _
	"GetChannelVolumeLevelScalar hresult(uint;float*);" & _
	"SetMute hresult(int;ptr);" & _
	"GetMute hresult(int*);" & _
	"GetVolumeStepInfo hresult(uint*;uint*);" & _
	"VolumeStepUp hresult(ptr);" & _
	"VolumeStepDown hresult(ptr);" & _
	"QueryHardwareSupport hresult(dword*);" & _
	"GetVolumeRange hresult(float*;float*;float*)"
Global Const $CLSCTX_INPROC_SERVER = 1

Global $__g_oEndpointVolume = 0

OnAutoItExitRegister("_EndpointVolume_Close")

; #FUNCTION# ====================================================================================================================
; Name ..........: _EndpointVolume_Init
; Description ...: Create global endpoint volume object used by other functions
; Syntax ........: _EndpointVolume_Init()
; Parameters ....:
; Return values .: Failure - Sets @error
;                          | 1 - Failed to create IMMDeviceEnumerator interface
;                          | 2 - Failed to get default audio endpoint interface pointer
;                          | 3 - Failed to create IMMDevice interface
;                          | 4 - Failed to get endpoint volume interface pointer
;                          | 5 - Failed to create IAudioEndpointVolume interface
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......: Calling a function before calling Init will automatically call Init. In this case if an error occurs,
;                  @error is set to -1. Explicitly call Init to get the actual @error code.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _EndpointVolume_Init()
	If IsObj($__g_oEndpointVolume) Then Return
	$__g_oEndpointVolume = __GetDefaultEndpointVolume()
	If @error Then Return SetError(@error)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _EndpointVolume_Close
; Description ...: Releases global endpoint volume object
; Syntax ........: _EndpointVolume_Close()
; Parameters ....:
; Return values .: None
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _EndpointVolume_Close()
	If IsObj($__g_oEndpointVolume) Then $__g_oEndpointVolume = 0
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetChannelCount
; Description ...: Get the number of channels in the default audio endpoint
; Syntax ........: _GetChannelCount()
; Parameters ....:
; Return values .: Success - Number of channels
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GetChannelCount()
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	Local $nCount
	If __SUCCEEDED($__g_oEndpointVolume.GetChannelCount($nCount)) Then
		Return $nCount
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _SetMasterVolumeLevel
; Description ...: Set master volume level in decibels.
; Syntax ........: _SetMasterVolumeLevel($idB)
; Parameters ....: $idB                 - Volume level in decibels (float)
; Return values .: Success - 1
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......: Call _GetVolumeRange for possible values. See _GetVolumeRange remarks for more information.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _SetMasterVolumeLevel($idB)
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	If __SUCCEEDED($__g_oEndpointVolume.SetMasterVolumeLevel($idB, 0)) Then
		Return 1
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _SetMasterVolumeLevelScalar
; Description ...: Set master volume level in scalar units, 0.0 to 100.0.
; Syntax ........: _SetMasterVolumeLevelScalar($iVol)
; Parameters ....: $iVol                - Volume level in scalar units, 0.0 to 100.0 (float)
; Return values .: Success - 1
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _SetMasterVolumeLevelScalar($iVol)
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	If __SUCCEEDED($__g_oEndpointVolume.SetMasterVolumeLevelScalar($iVol / 100, 0)) Then
		Return 1
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetMasterVolumeLevel
; Description ...: Get master volume level in decibels
; Syntax ........: _GetMasterVolumeLevel()
; Parameters ....:
; Return values .: Success - Master volume level in decibels (float)
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GetMasterVolumeLevel()
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	Local $idB
	If __SUCCEEDED($__g_oEndpointVolume.GetMasterVolumeLevel($idB)) Then
		Return $idB
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetMasterVolumeLevelScalar
; Description ...: Get master volume level in scalar units, 0.0 to 100.0.
; Syntax ........: _GetMasterVolumeLevelScalar()
; Parameters ....:
; Return values .: Success - Master volume level in scalar units, 0.0 to 100.0 (float)
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GetMasterVolumeLevelScalar()
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	Local $iVol
	If __SUCCEEDED($__g_oEndpointVolume.GetMasterVolumeLevelScalar($iVol)) Then
		Return ($iVol * 100)
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _SetChannelVolumeLevel
; Description ...: Set volume level for a specific channel in decibels.
; Syntax ........: _SetChannelVolumeLevel($iChannel, $idB)
; Parameters ....: $iChannel            - Channel number to modify, 0 based index
;                  $idB                 - Volume level in decibels (float)
; Return values .: Success - 1
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......: Call _GetChannelCount for number of channels, _GetVolumeRange for possible values.
;                  See _GetVolumeRange remarks for more information.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _SetChannelVolumeLevel($iChannel, $idB)
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	If __SUCCEEDED($__g_oEndpointVolume.SetChannelVolumeLevel($iChannel, $idB, 0)) Then
		Return 1
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _SetChannelVolumeLevelScalar
; Description ...: Set volume level for a specific channel in scalar units, 0.0 to 100.0.
; Syntax ........: _SetChannelVolumeLevelScalar($iChannel, $iVol)
; Parameters ....: $iChannel            - Channel number to modify, 0 based index
;                  $iVol                - Volume level in scalar units, 0.0 to 100.0 (float)
; Return values .: Success - 1
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......: Call _GetChannelCount for number of channels.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _SetChannelVolumeLevelScalar($iChannel, $iVol)
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	If __SUCCEEDED($__g_oEndpointVolume.SetChannelVolumeLevelScalar($iChannel, $iVol / 100, 0)) Then
		Return 1
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetChannelVolumeLevel
; Description ...: Get volume level for a specific channel in decibels.
; Syntax ........: _GetChannelVolumeLevel($iChannel)
; Parameters ....: $iChannel            - Channel number to query, 0 based index
; Return values .: Success - Channel volume level in decibels (float)
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......: Call _GetChannelCount for number of channels.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GetChannelVolumeLevel($iChannel)
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	Local $idB
	If __SUCCEEDED($__g_oEndpointVolume.GetChannelVolumeLevel($iChannel, $idB)) Then
		Return $idB
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetChannelVolumeLevelScalar
; Description ...: Get volume of a specific channel in scalar units, 0.0 to 100.0.
; Syntax ........: _GetChannelVolumeLevelScalar($iChannel)
; Parameters ....: $iChannel            - Channel number to query, 0 based index
; Return values .: Success - Channel volume in scalar units, 0.0 to 100.0 (float)
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......: Call _GetChannelCount for number of channels.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GetChannelVolumeLevelScalar($iChannel)
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	Local $iVol
	If __SUCCEEDED($__g_oEndpointVolume.GetChannelVolumeLevelScalar($iChannel, $iVol)) Then
		Return ($iVol * 100)
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _SetMute
; Description ...: Set mute state
; Syntax ........: _SetMute($bMute)
; Parameters ....: $bMute               - Desired mute state: 1 to mute, 0 to unmute
; Return values .: Success - 1
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _SetMute($bMute)
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	If __SUCCEEDED($__g_oEndpointVolume.SetMute($bMute, 0)) Then
		Return 1
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetMute
; Description ...: Get current mute state
; Syntax ........: _GetMute()
; Parameters ....:
; Return values .: Success - Current mute state
;                          | 0 - not muted
;                          | 1 - muted
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GetMute()
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	Local $bMute
	If __SUCCEEDED($__g_oEndpointVolume.GetMute($bMute)) Then
		Return $bMute
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetVolumeStepInfo
; Description ...: Get current volume step and range.
; Syntax ........: _GetVolumeStepInfo()
; Parameters ....:
; Return values .: Success - Two element array containing step info
;                          [0] - Current volume step
;                          [1] - Volume step range
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......: Volume step values range from 0 to (range-1).
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GetVolumeStepInfo()
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	Local $iCurrentStep, $iStepRange
	If __SUCCEEDED($__g_oEndpointVolume.GetVolumeStepInfo($iCurrentStep, $iStepRange)) Then
		Local $aRet[2] = [$iCurrentStep, $iStepRange]
		Return $aRet
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _VolumeStepUp
; Description ...: Increase volume by one step.
; Syntax ........: _VolumeStepUp()
; Parameters ....:
; Return values .: Success - 1
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......: Call _GetVolumeStepInfo for current step and range.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _VolumeStepUp()
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	If __SUCCEEDED($__g_oEndpointVolume.VolumeStepUp(0)) Then
		Return 1
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _VolumeStepDown
; Description ...: Decrease volume by one step.
; Syntax ........: _VolumeStepDown()
; Parameters ....:
; Return values .: Success - 1
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......: Call _GetVolumeStepInfo for current step and range.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _VolumeStepDown()
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	If __SUCCEEDED($__g_oEndpointVolume.VolumeStepDown(0)) Then
		Return 1
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _QueryHardwareSupport
; Description ...: Query audio endpoint device for hardware supported functions.
; Syntax ........: _QueryHardwareSupport()
; Parameters ....:
; Return values .: Success - Dword mask of bitwise OR'd hardware function constants.
;                  Failure - 0 and sets @error
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _QueryHardwareSupport()
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	Local $iMask
	If __SUCCEEDED($__g_oEndpointVolume.QueryHardwareSupport($iMask)) Then
		Return $iMask
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetVolumeRange
; Description ...: Get volume range information.
; Syntax ........: _GetVolumeRange()
; Parameters ....:
; Return values .: Success - Three element array containing volume range information
;                          [0] - Minimum volume in decibels (float)
;                          [1] - Maximum volume in decibels (float)
;                          [2] - Volume increment value in decibels (float)
; Author ........: Erik Pilsits
; Modified ......:
; Remarks .......: The volume range is divided into equal steps such that nSteps = (max - min) / increment.
;                  A call to SetMasterVolumeLevel or SetChannelVolumeLevel that falls between steps will set the volume
;                  to the nearest step.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _GetVolumeRange()
	_EndpointVolume_Init()
	If @error Then Return SetError(-1, 0, 0)
	;
	Local $iMindB, $iMaxdB, $iIncrementdB
	If __SUCCEEDED($__g_oEndpointVolume.GetVolumeRange($iMindB, $iMaxdB, $iIncrementdB)) Then
		Local $aRet[3] = [$iMindB, $iMaxdB, $iIncrementdB]
		Return $aRet
	Else
		Return SetError(1, 0, 0)
	EndIf
EndFunc

#region INTERNAL FUNCTIONS
Func __GetDefaultEndpointVolume()
	Local $oIAudioEndpointVolume = 0, $err = 1
	; create device enumerator interface
	Local $oDevEnum = ObjCreateInterface($CLSID_MMDeviceEnumerator, $IID_IMMDeviceEnumerator, $tagIMMDeviceEnumerator)
	If IsObj($oDevEnum) Then
		$err = 2
		; get default audio endpoint interface pointer
		Local $pDefaultDevice = 0
		If __SUCCEEDED($oDevEnum.GetDefaultAudioEndpoint($eRender, $eConsole, $pDefaultDevice)) Then
			$err = 3
			; create default audio endpoint interface
			Local $oIMMDevice = ObjCreateInterface($pDefaultDevice, $IID_IMMDevice, $tagIMMDevice)
			If IsObj($oIMMDevice) Then
				$err = 4
				; get endpoint volume interface pointer
				Local $pEndpointVolume = 0
				If __SUCCEEDED($oIMMDevice.Activate(__uuidof($IID_IAudioEndpointVolume), $CLSCTX_INPROC_SERVER, 0, $pEndpointVolume)) Then
					$err = 5
					; create endpoint volume interface
					$oIAudioEndpointVolume = ObjCreateInterface($pEndpointVolume, $IID_IAudioEndpointVolume, $tagIAudioEndpointVolume)
				EndIf
				$oIMMDevice = 0
			EndIf
		EndIf
		$oDevEnum = 0
	EndIf
	;
	If IsObj($oIAudioEndpointVolume) Then
		Return $oIAudioEndpointVolume
	Else
		Return SetError($err, 0, 0)
	EndIf
EndFunc

Func __SUCCEEDED($hr)
	Return ($hr >= 0)
EndFunc

Func __uuidof($sGUID)
	Local $tGUID = DllStructCreate("ulong Data1;ushort Data2;ushort Data3;byte Data4[8]")
	DllCall("ole32.dll", "long", "CLSIDFromString", "wstr", $sGUID, "struct*", $tGUID)
	If @error Then Return SetError(@error, @extended, 0)
	Return $tGUID
EndFunc
#endregion INTERNAL FUNCTIONS
