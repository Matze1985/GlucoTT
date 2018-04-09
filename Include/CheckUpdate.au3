#include <Misc.au3>
#include <Date.au3>
#include <StringConstants.au3>
#include <File.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <ButtonConstants.au3>
#include <ExtMsgBox.au3>

; Update a script (compiled or not) by downloading the updated version from the Internet
;  script version must have format x.x.x.x !
;	as per format above (#AutoIt3Wrapper_Res_Fileversion=1.0.1.0)
;   If you use a different version number, you will have to change the logic in function CheckUpdate
;	You have to compile the script using wrapper functions, because the version number is fetched with FileGetVersion()
;----------------------------------------------------------------------------------------------------------------------------------
;	Author ........: GreenCan
;   Modified.......: 14.01.2018 - Matze1985 - Change normal MsgBox to Extended Message Box
; 	UDF-URL........: https://www.autoitscript.com/forum/topic/162107-checkupdate-autoupdate-a-running-script-or-exe-over-the-web/
;----------------------------------------------------------------------------------------------------------------------------------

; Globals
Global $_CRC32_CodeBuffer, $_CRC32_CodeBufferMemory ; CRC checksum

#Region version
   Global $sVersion = FileGetVersion(@ScriptName)
   Local $sHeaderLine
   ; get version for non-compiled script (get the script version (#AutoIt3Wrapper_Res_Fileversion=1.0.1.0)
   If Not @Compiled Then
      For $i = 2 To 10
         $sHeaderLine = FileReadLine(@ScriptDir & "\" & @ScriptName, $i)
         If StringLeft($sHeaderLine, 10) = "#EndRegion" Then ExitLoop ; not found
         If StringInStr($sHeaderLine, "Res_Fileversion") > 0 Then
            $sVersion = StringTrimLeft($sHeaderLine, 32)
            ExitLoop
         EndIf
      Next
   EndIf
#EndRegion version

#Region update
   ; #FUNCTION# ================================================================================
   ; Name...........: CheckUpdate
   ; Description ...: Check version of the script or executable on Internet download site and Updates it if newer, if user agrees
   ; Syntax.........: CheckUpdate($sFileToUpdate, $sCurrentVersion, $sUpdateINI[, $InetForceReload = 1])
   ; Parameters ....: $sFileToUpdate - Script of Compiled script to Update
   ;				   $sCurrentVersion - Current version of the script ( as set in #AutoIt3Wrapper_Res_Fileversion consisiting of 4 version levels x.x.x.x)
   ;				   $sUpdateINI - url to INI file containing the update information
   ;				   $InetForceReload - force (uncached) download of the INI file
   ; Return values .: Success - exit the current script and restart new version
   ;                  Failure - 1, sets @error
   ;                  |1 - No newer version available
   ;                  |2 - User refused Update
   ;                  |3 - Update file not available
   ;                  |4 - Version info not available in Update file
   ;                  |5 - Download failure (and error message displayed)
   ;                  |6 - File size error, update aborted (and error message displayed)
   ;                  |7 - CRC Checksum error, update aborted (and error message displayed)
   ;
   ; 					Note: In case of error, just continue with current script...
   ;
   ; Author ........: GreenCan
   ; Modified.......:
   ; Remarks .......: 		The script version has format x.x.x.x !
   ;						In case the script (or exe) has been renamed (and does not correspond to the original file name),
   ;						The orginal file will not be renamed
   ;						If the file has it original file name, then the original file will be renamed to the filename_version.extension
   ;						example:
   ;							1. UpdateTest.au3 version 1.0.1.0 will be renamed to UpdateTest_1.0.1.0.au3
   ;								and UpdateTest.au3 version 1.0.1.2 will take its place
   ;							2. WhathEverNewName.au3 (UpdateTest.au3 version 1.0.1.0) will not be renamed and remane WhathEverNewName.au3
   ;								and UpdateTest.au3 version 1.0.1.2 will be started
   ;
   ;						ini file for this example
   ;						[UpdateTest.au3]
   ;						version=1.0.1.1
   ;						date=2014/06/17 10:00
   ;						Filesize=16720
   ;						CRC=FA7B28EB
   ;						download=http://users.telenet.be/GreenCan/AutoIt/Updates/UpdateTest_1.0.1.1.au3
   ;
   ;						[UpdateTest.exe]
   ;						version=1.0.1.1
   ;						date=2014/06/17 10:00
   ;						Filesize=473088
   ;						CRC=BC2B80BE
   ;						download=http://users.telenet.be/GreenCan/AutoIt/Updates/UpdateTest_1.0.1.1.exe
   ;
   ; Related .......:
   ; Link ..........:
   ; Example .......:
   ; ===========================================================================================
   Func CheckUpdate($sFileToUpdate, $sCurrentVersion, $sUpdateINI, $InetForceReload = 1)
      Local $sINI_Data, $sNewVersion, $sDate, $sURL, $sChangesURL, $sDestinationFile, $szDrive, $szDir, $szFName, $szExt, $UpdateScript, $sTempFile, $Return
      Local $procwatchPID, $aVersion, $iNewVersion, $iCurrentVersion, $iBufferSize = 0x80000, $iCRC, $iCRC32 = 0, $sData, $FileSize, $i_FileSize, $hFile
      $sINI_Data = InetRead($sUpdateINI, $InetForceReload) ; get the ini file
      If Not @error Then
         If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Update file found" & @CRLF)
         $sINI_Data = BinaryToString($sINI_Data)
         ; read the ini file in memory
         $sNewVersion = IniMemoryRead($sINI_Data, $sFileToUpdate, "version", "")
         If $sNewVersion <> "" Then
            $sDate = IniMemoryRead($sINI_Data, $sFileToUpdate, "date", "")
            $sDate = _DateTimeFormat($sDate, 2) & " - " & _DateTimeFormat($sDate, 4) ; set current time in Locale format
            $sURL = IniMemoryRead($sINI_Data, $sFileToUpdate, "download", "")
            $iFilesize = IniMemoryRead($sINI_Data, $sFileToUpdate, "Filesize", 0)
            $iCRC = IniMemoryRead($sINI_Data, $sFileToUpdate, "CRC", 0)
            $sChangesURL = IniMemoryRead($sINI_Data, $sFileToUpdate, "changes", 0)

            If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Application version: Current: " & $sCurrentVersion & " - New: " & $sNewVersion & "<" & @CRLF)

            ; convert version x.x.x.x to a number where each x can go up to 999, so the max number can be 999 999 999 999
            ; so version 1.0.1.29 will be converted to 1000001029  (001.000.001.029)
            $aVersion = StringSplit($sNewVersion, ".")
            $iNewVersion = $aVersion[4] + ($aVersion[3] * 1000) + ($aVersion[2] * 1000000) + ($aVersion[1] * 1000000000)

            ; do the same for current version
            $aVersion = StringSplit($sCurrentVersion, ".")
            $iCurrentVersion = $aVersion[4] + ($aVersion[3] * 1000) + ($aVersion[2] * 1000000) + ($aVersion[1] * 1000000000)
            ; Update only to a newer version
            If $sNewVersion > $sCurrentVersion Then
;~ 				If MsgBox(36, "Update " & $sFileToUpdate & " ver " & $sCurrentVersion, "A new version of " & $sFileToUpdate & " is available since " & $sDate & "." & @CRLF & "Download version " & $sNewVersion & " now? ") = 6 Then ;  No is Default
               While "loop view History"
				  _ExtMsgBoxSet(-1, -1, -1, -1, -1, -1, 700)
                  $Return = _ExtMsgBox(32, " ChangeLog | Yes | No", "Update " & $sFileToUpdate & " ver " & $sCurrentVersion, "A new version of " & $sFileToUpdate & " is available since " & $sDate & "." & @CRLF & "Download version " & $sNewVersion & " now? ")
                  Switch $Return
                     Case 1
                        $Return = 6
                     Case 2
                        $Return = 7
                     Case 3
                        $Return = 2
                  EndSwitch

                  ; Escape = 2, No = 2, View Changes = 6, Download = 7 (No is Default)
                  If $Return <> 6 Then ExitLoop
                  If $sChangesURL = "" Then
                     _ExtMsgBox(0, $MB_OK, "History of changes " & $sFileToUpdate, "Sorry but release information is not available at the moment.")
                  Else
                     $sContent = BinaryToString(InetRead($sChangesURL, $InetForceReload))
                     If Not @error Then
                        If StringInStr($sContent, @CRLF) = 0 Then $sContent = StringReplace($sContent, @LF, @CRLF) ;
                        Text_Viewer("History of changes " & $sFileToUpdate, $sContent, 10)
                     Else
                        _ExtMsgBox(0, $MB_OK, "History of changes " & $sFileToUpdate, "Sorry but release information is not available at the moment.")
                     EndIf
                  EndIf
               WEnd

               If $Return = 7 Then ;  download
                  _PathSplit($sFileToUpdate, $szDrive, $szDir, $szFName, $szExt)
                  If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") : Downloading " & @ScriptDir & "\" & $szFName & "_" & $sNewVersion & $szExt & @CRLF)
                  InetgetProgress($sURL, @ScriptDir & "\" & $szFName & "_" & $sNewVersion & $szExt)
                  If @error Then
                     _ExtMsgBox(16, $MB_OK, "Error", "Download failure " & @CR & _
                           $sURL & @CR & "Please retry later.")
                     If FileExists(@ScriptDir & "\" & $szFName & "_" & $sNewVersion & $szExt) Then _
                           FileDelete(@ScriptDir & "\" & $szFName & "_" & $sNewVersion & $szExt) ; delete the file because it may be corrupt
                     Return SetError(5, 0, 0)
                  Else
                     ; checksum verification
                     $i_FileSize = FileGetSize(@ScriptDir & "\" & $szFName & "_" & $sNewVersion & $szExt)
                     If $i_FileSize <> $iFilesize Then
                        _ExtMsgBox(16, $MB_OK, "Error", "Download failure, File size error!" & @CR & _
                              $sURL & @CR & "Please retry later.")
                        If FileExists(@ScriptDir & "\" & $szFName & "_" & $sNewVersion & $szExt) Then _
                              FileDelete(@ScriptDir & "\" & $szFName & "_" & $sNewVersion & $szExt) ; delete the file because it may be corrupt
                        Return SetError(6, 0, 0)
                     Else
                        $hFile = FileOpen(@ScriptDir & "\" & $szFName & "_" & $sNewVersion & $szExt, 16)

                        For $i = 1 To Ceiling($i_FileSize / $iBufferSize)
                           $sData = FileRead($hFile, $iBufferSize)
                           $iCRC32 = _CRC32($sData, BitNOT($iCRC32))
                        Next
                        FileClose($hFile)
                        If Hex($iCRC32, 8) <> $iCRC Then
                           _ExtMsgBox(16, $MB_OK, "Error", "Download failure, CRC Checksum error!" & @CR & _
                                 $sURL & @CR & "Please retry later.")
                           If FileExists(@ScriptDir & "\" & $szFName & "_" & $sNewVersion & $szExt) Then _
                                 FileDelete(@ScriptDir & "\" & $szFName & "_" & $sNewVersion & $szExt) ; delete the file because it may be corrupt
                           Return SetError(7, 0, 0)
                        Else
                           If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") : CRC Checksum successful " & Hex($iCRC32, 8) & @CRLF)
                           ; download successful, exit this program just after starting the temporary batch file that does following:
                           ; 1. wait for a few seconds to enable to current script to Exit (using ping trick)
                           ; 2. delete old file
                           ; 3. then rename the new downloaded version to the initially used script name (scriptname.au3 or scriptname.exe)
                           ; start the new script
                           ; finally auto-delete the temporary script
                           ; quit

                           $UpdateScript = '@ECHO ON' & _
                                 @CRLF & _
                                 'ping 127.0.0.1 -n 5 -w 5000' & _
                                 @CRLF & _
                                 'DEL /F "' & @ScriptDir & '\' & $szFName & '.exe' & _
                                 @CRLF & _
                                 'rename "' & @ScriptDir & "\" & $szFName & "_" & $sNewVersion & $szExt & '" ' & $szFName & $szExt & _
                                 @CRLF & _
                                 'start ' & $szFName & $szExt & _
                                 @CRLF & _
                                 'DEL /F "' & @ScriptDir & '\proc.cmd"' & _
                                 @CRLF

                           If Not @Compiled Then ConsoleWrite($UpdateScript & @CR)

                           $sTempFile = FileOpen(@ScriptDir & "\proc.cmd", 2)
                           FileWrite($sTempFile, $UpdateScript)
                           FileClose($sTempFile)
                           _ExtMsgBox(0, $MB_OK, "Restart", "Wait a moment. Restarting the program", 2)
                           Run('"' & @ScriptDir & '\proc.cmd"', "", @SW_HIDE)
                           ; done, now exit
                           Exit
                        EndIf
                     EndIf
                  EndIf
               Else
                  If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " User refused Update" & @CRLF)
                  Return SetError(2, 0, 0)
               EndIf
            Else
               If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " No newer version available" & @CRLF)
               Return SetError(1, 0, 0)
            EndIf
         Else
            If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Version info not available in Update file" & @CRLF)
            Return SetError(4, 0, 0)
         EndIf
      Else
         If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") :" & " Update file not available" & @CRLF)
         Return SetError(3, 0, 0)
      EndIf

   EndFunc

   ; #FUNCTION# ================================================================================
   ; Name...........: Text_Viewer
   ; Description ...: Simple Text Viewer GUI
   ; Syntax.........: Text_Viewer($s_Title, $s_Text[, $i_FontSize = 8.5])
   ; Parameters ....: $s_Title - Window title
   ;				   $s_Text - Text to display
   ;				   $i_FontSize - Font size (default = 8.5)
   ; Return values .: none
   ;
   ; Author ........: GreenCan
   ; Modified.......:
   ; Remarks .......:
   ; Related .......:
   ; Link ..........:
   ; Example .......:
   ; ===========================================================================================
   Func Text_Viewer($s_Title, $s_Text, $i_FontSize = 8.5)
      Local $iMsg
      Local $iWindow_width = 700
      Local $iWindow_heigth = 500
      Local $hGUI_Viewer = GUICreate($s_Title, $iWindow_width, $iWindow_heigth, -1, -1, $WS_CAPTION, $WS_EX_TOPMOST)
      GUICtrlCreateEdit($s_Text, 5, 5, $iWindow_width - 10, $iWindow_heigth - 35, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_READONLY, $ES_WANTRETURN, $WS_HSCROLL, $WS_VSCROLL))
      GUICtrlSetFont(-1, $i_FontSize)
      GUICtrlSetResizing(-1, $GUI_DOCKTOP + $GUI_DOCKBOTTOM)

      Local $hButtonOK = GUICtrlCreateButton("OK", 10, $iWindow_heigth - 25, 80, 20, $BS_DEFPUSHBUTTON)

      GUISetState()
      Do
         $iMsg = GUIGetMsg()
         Select
            Case $iMsg = $hButtonOK Or $iMsg = -3
               GUIDelete($hGUI_Viewer)
               Return

         EndSelect
      Until $iMsg = $GUI_EVENT_CLOSE
   EndFunc

   ; #FUNCTION# ================================================================================
   ; Name...........: IniMemoryRead
   ; Description ...: IniRead from an ini file loaded into memory
   ; Syntax.........: IniMemoryRead($sIniContent, $sSection, $sKey, $sDefault)
   ; Parameters ....: $sIniContent - The ini file loaded into memory
   ;				   $sSection    - The section name in the .ini file.
   ;				   $sKey 	    - The key name in the .ini file.
   ;				   $sDefault    - The default value to return if the requested key is not found.

   ; Return values .: Success - The requested key value as a string.
   ;                  Failure - The default string if requested key not found.
   ; Author ........: GreenCan
   ; Modified.......:
   ; Remarks .......:
   ; Related .......:
   ; Link ..........:
   ; Example .......:
   ; ===========================================================================================
   Func IniMemoryRead($sIniContent, $sSection, $sKey, $sDefault)
      Local $aIniContent
      $aIniContent = StringSplit($sIniContent, @CRLF)
      ; find section (case unsensitive)
      For $i = 1 To $aIniContent[0]
         $aIniContent[$i] = StringLower(StringStripWS($aIniContent[$i], $STR_STRIPLEADING + $STR_STRIPTRAILING)) ; remove blanks
         If $aIniContent[$i] == "[" & StringLower($sSection) & "]" Then ExitLoop
      Next
      If $i > $aIniContent[0] Then Return $sDefault

      ; find key (case unsensitive)
      For $i = $i + 1 To $aIniContent[0]
         If $aIniContent[$i] = "" Then ContinueLoop ; skip empty line
         $aIniContent[$i] = StringStripWS($aIniContent[$i], $STR_STRIPLEADING + $STR_STRIPTRAILING) ; remove blanks
         If StringLeft($aIniContent[$i], 1) = ";" Then ContinueLoop ; skip remarks
         If StringLeft($aIniContent[$i], 1) = "[" And StringRight($aIniContent[$i], 1) = "]" Then Return $sDefault ; OK passed the complete section so return the Default result
         If StringLower(StringLeft($aIniContent[$i], StringLen($sKey))) = StringLower($sKey) Then Return StringStripWS(StringTrimLeft($aIniContent[$i], StringLen($sKey) + 1), $STR_STRIPLEADING + $STR_STRIPTRAILING) ; return Key
      Next
      Return $sDefault
   EndFunc

   ; -----------------------------------------------------------------------------
   ; CRC Checksum Machine Code UDF
   ; Purpose: Provide The Machine Code Version of CRC16/CRC32 Algorithm In AutoIt
   ; Author: Ward
   ; http://www.autoitscript.com/forum/topic/121985-autoit-machine-code-algorithm-collection/
   ; -----------------------------------------------------------------------------
   Func _CRC32_Exit()
      $_CRC32_CodeBuffer = 0
      _MemVirtualFree($_CRC32_CodeBufferMemory, 0, $MEM_RELEASE)
   EndFunc

   Func _CRC32($Data, $Initial = -1, $Polynomial = 0xEDB88320)
      If Not IsDllStruct($_CRC32_CodeBuffer) Then
         If @AutoItX64 Then
            Local $Opcode = '0xC80004004989CA680001000059678D41FF516A0859D1E873034431C8E2F75989848DFCFBFFFFE2E589D14489C04D85D2741B67E318418A1230C2480FB6D2C1E80833849500FCFFFF49FFC2E2E8F7D0C9C3'
         Else
            Local $Opcode = '0xC8000400538B5514B9000100008D41FF516A0859D1E8730231D0E2F85989848DFCFBFFFFE2E78B5D088B4D0C8B451085DB7416E3148A1330C20FB6D2C1E80833849500FCFFFF43E2ECF7D05BC9C21000'
         EndIf
         $Opcode = Binary($Opcode)

         $_CRC32_CodeBufferMemory = _MemVirtualAlloc(0, BinaryLen($Opcode), $MEM_COMMIT, $PAGE_EXECUTE_READWRITE)
         $_CRC32_CodeBuffer = DllStructCreate("byte[" & BinaryLen($Opcode) & "]", $_CRC32_CodeBufferMemory)
         DllStructSetData($_CRC32_CodeBuffer, 1, $Opcode)
         OnAutoItExitRegister("_CRC32_Exit")
      EndIf

      $Data = Binary($Data)
      Local $InputLen = BinaryLen($Data)
      Local $Input = DllStructCreate("byte[" & $InputLen & "]")
      DllStructSetData($Input, 1, $Data)

      Local $Ret = DllCall("user32.dll", "uint", "CallWindowProc", "ptr", DllStructGetPtr($_CRC32_CodeBuffer), _
            "ptr", DllStructGetPtr($Input), _
            "uint", $InputLen, _
            "uint", $Initial, _
            "uint", $Polynomial)

      Return $Ret[0]
   EndFunc
#EndRegion update

; #FUNCTION# ================================================================================
; Name...........: InetgetProgress
; Description ...: Download a file with progress bar
; Syntax.........: InetgetProgress($sURL, $sFilename)
; Parameters ....: $sURL - URL of the file to be downloaded
;				   $sFilename - Local name of the destination file to be downloaded
; Return values .: Success - 0
;                  Failure - 1, sets @error
;                  |1 - Download failed
; Author ........: GreenCan
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===========================================================================================
Func InetgetProgress($sURL, $sFilename)
   Local $iSize, $iTotalSize, $hDownload, $iSec, $iCurrentBytes, $iReadBytes, $szDrive, $szDir, $szFName, $szExt
   _PathSplit($sURL, $szDrive, $szDir, $szFName, $szExt)
   $iSize = InetGetSize($sURL)
   $iTotalSize = Round($iSize / 1024)
   If Not @Compiled Then ConsoleWrite("@@ Debug(" & @ScriptLineNumber & ") : $szFName & $szExt: " & $szFName & $szExt & " $sURL: " & $sURL & @CRLF)
   $hDownload = InetGet($sURL, $sFilename, 16, 1) ; from InetConstants.au3:  $INET_FORCEBYPASS (16) = By-pass forcing the connection online, $INET_DOWNLOADBACKGROUND (1) = Background download
   ProgressOn("Download " & $szFName & $szExt, "Download progress")
   Do
      $iSec = @SEC
      $iCurrentBytes = Round(InetGetInfo($hDownload, 0))
      While @SEC = $iSec
         Sleep(1000)
      WEnd
      $iReadBytes = Round(InetGetInfo($hDownload, 0))
      $iTotalSize = $iTotalSize - (($iReadBytes - $iCurrentBytes) / 1024)
      ProgressSet(100 - Round($iTotalSize / $iSize * 100000), 100 - Round($iTotalSize / $iSize * 100000) & "%")
   Until InetGetInfo($hDownload, 2)
   ProgressOff()
   If Not InetGetInfo($hDownload, 3) Then Return SetError(1, 0, 0)
EndFunc