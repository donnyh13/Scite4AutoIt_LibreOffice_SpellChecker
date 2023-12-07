#Region ; *** Dynamically added Include files ***
#include <StringConstants.au3>                               ; added:12/07/23 15:41:13
#EndRegion ; *** Dynamically added Include files ***
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

Global $bReturn
Global $iError

If ($CmdLine[0] > 0) Then
	__Print_To_Error()
	If ($CmdLine[0] = 4) Then ; Single Word Check.
		$bReturn = _SpellCheck($CmdLine[1], $CmdLine[2], $CmdLine[3], $CmdLine[4])
		$iError = @error

	Else ; Entire Script Word Check.
		$bReturn = _SpellCheck($CmdLine[1], $CmdLine[2], $CmdLine[3])
		$iError = @error
	EndIf

	If ($iError > 0) Then ; Something went wrong.
		Exit 2

	ElseIf ($bReturn = False) Then ; Word(s) are incorrectly spelled.
		__Print_To_Error(Null) ; Delete the Error file.
		Exit 1

	Else ; Word(s) are correctly spelled.
		__Print_To_Error(Null) ; Delete the Error file.
		Exit 0

	EndIf

Else ; If not params passed, exit.
	__Print_To_Error("No Parameters passed")
	Exit 2

EndIf

Func _SpellCheck($sWordToCheck, $sLanguage = "en", $sCountry = "US", $bReturnWords = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __SpellCheck_ComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__S4A_LO_SC_INPUT_ERROR = 1, $__S4A_LO_SC_INIT_ERROR = 2, $__S4A_LO_SC_PROCESS_ERROR = 3
	Local Const $FO_READ = 0, $FO_OVERWRITE = 2, $FO_CREATEPATH = 8
	Local $sSpCheckFile = @ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini"
	Local $oServiceManager, $oSpellChecker
	Local $aEmptyArgs[1]
	Local $tLocale
	Local $vReturn
	Local $hFile

	If Not IsString($sWordToCheck) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 1, __Print_To_Error("Word called to check is not a String."))
	If Not IsString($sLanguage) Or ((StringLen($sLanguage) <> 2) And (StringLen($sLanguage) <> 3)) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 2, __Print_To_Error("Language code called is not a string, or is not 2 or 3 characters long."))
	If Not IsString($sCountry) Or (StringLen($sCountry) <> 2) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 3, __Print_To_Error("Country code called is not a string, or is not 2 characters long." & $sCountry))
	If IsString($bReturnWords) And ($bReturnWords = "true") Then $bReturnWords = True
	If IsString($bReturnWords) And ($bReturnWords = "false") Then $bReturnWords = False
	If Not IsBool($bReturnWords) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 4, __Print_To_Error("Return Words called is is not a Boolean."))

	If Not FileExists($sSpCheckFile) Then
		$hFile = FileOpen($sSpCheckFile, $FO_CREATEPATH) ; Make sure file exists.
		If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __Print_To_Error("Failed to open Scite4AutoIt_LO_SpellChecker.ini. #1"))
		FileClose($hFile)
	EndIf

	$sLanguage = StringLower($sLanguage)
	$sCountry = StringUpper($sCountry)

	$tLocale = __SpellCheck_CreateStruct("com.sun.star.lang.Locale")
	If @error Then Return SetError($__S4A_LO_SC_INIT_ERROR, 2, __Print_To_Error("Failed to create com.sun.star.lang.Locale Structure."))
	$tLocale.Country = $sCountry
	$tLocale.Language = $sLanguage

	$aEmptyArgs[0] = __LOWriter_SetPropertyValue("", "")
	If @error Then Return SetError($__S4A_LO_SC_INIT_ERROR, 3, __Print_To_Error("Failed to create Property Structure."))

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If @error Then Return SetError($__S4A_LO_SC_INIT_ERROR, 4, __Print_To_Error("Failed to create com.sun.star.ServiceManager Object."))

	$oSpellChecker = $oServiceManager.createInstance("com.sun.star.linguistic2.SpellChecker")
	If Not IsObj($oSpellChecker) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 5, __Print_To_Error("Failed to create com.sun.star.linguistic2.SpellChecker Object."))

	If Not $oSpellChecker.hasLocale($tLocale) Then Return SetError($__S4A_LO_SC_PROCESS_ERROR, 1, __Print_To_Error("Language and Country Combination is not valid."))

	If ($bReturnWords = True) Then ; If Return Words = True, then I am checking a single word.
		$hFile = FileOpen($sSpCheckFile, $FO_OVERWRITE)
		If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __Print_To_Error("Failed to open Scite4AutoIt_LO_SpellChecker.ini. #2"))
		FileFlush($hFile)

		$vReturn = __SingleWordCheck($sWordToCheck, $hFile, $oSpellChecker, $tLocale, $aEmptyArgs)
		FileClose($hFile)
		Return SetError(@error, @extended, $vReturn)

	Else ; Entire Script check.
		$hFile = FileOpen($sSpCheckFile, $FO_READ)
		If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __Print_To_Error("Failed to open Scite4AutoIt_LO_SpellChecker.ini. #3"))
		$vReturn = __ScriptWordCheck($hFile, $oSpellChecker, $tLocale, $aEmptyArgs)
		FileClose($hFile)
		Return SetError(@error, @extended, $vReturn)

	EndIf

EndFunc   ;==>_SpellCheck


Func __SingleWordCheck($sWordToCheck, ByRef $hFile, ByRef $oSpellChecker, ByRef $tLocale, ByRef $aEmptyArgs)
	Local Const $__S4A_LO_SC_SUCCESS = 0, $__S4A_LO_SC_INIT_ERROR = 2
	Local $oSpell
	Local $asArray

	If $oSpellChecker.isValid($sWordToCheck, $tLocale, $aEmptyArgs) Then Return SetError($__S4A_LO_SC_SUCCESS, 0, True)

	$oSpell = $oSpellChecker.Spell($sWordToCheck, $tLocale, $aEmptyArgs)
	If Not IsObj($oSpell) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __Print_To_Error("Failed to retrieve Spelling Object."))

	$asArray = $oSpell.getAlternatives()
	If Not IsArray($asArray) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 2, __Print_To_Error("Failed to retrieve array of Alternative words."))

	If (UBound($asArray) > 0) Then
		For $i = 0 To UBound($asArray) - 1
			FileWrite($hFile, $asArray[$i] & @CRLF)
		Next

		FileFlush($hFile)

	EndIf

	Return SetError($__S4A_LO_SC_SUCCESS, 0, False)
EndFunc   ;==>__SingleWordCheck

Func __ScriptWordCheck(ByRef $hFile, ByRef $oSpellChecker, ByRef $tLocale, ByRef $aEmptyArgs)
	Local Const $__S4A_LO_SC_SUCCESS = 0, $__S4A_LO_SC_INIT_ERROR = 2, $__S4A_LO_SC_PROCESS_ERROR = 3
	Local Const $FO_OVERWRITE = 2
	Local $asWords[0]
	Local $iLines = 0, $iCount = 0
	Local $sWordToCheck = "", $sSpCheckFile = @ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini"

	$asWords = FileReadToArray($hFile)
	If @error Then Return SetError($__S4A_LO_SC_PROCESS_ERROR, 1, __Print_To_Error("Failed to Read File to Array."))
	$iLines = @extended

	FileClose($hFile)

	$hFile = FileOpen($sSpCheckFile, $FO_OVERWRITE)
	If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __Print_To_Error("Failed to open Scite4AutoIt_LO_SpellChecker.ini. #4"))
	FileFlush($hFile)

	For $i = 0 To $iLines - 1
		$sWordToCheck = StringLeft($asWords[$i], StringInStr($asWords[$i], @TAB) - 1)
		If $oSpellChecker.isValid($sWordToCheck, $tLocale, $aEmptyArgs) Then $asWords[$i] = 0 ; Overwrite the word with 0 to indicate it is spell correctly.

	Next

	For $i = 0 To $iLines - 1

		If IsString($asWords[$i]) Then
			FileWrite($hFile, $asWords[$i] & @CRLF)
			$iCount += 1
		EndIf

	Next

	FileFlush($hFile)

	Return ($iCount > 0) ? SetError($__S4A_LO_SC_SUCCESS, 0, False) : SetError($__S4A_LO_SC_SUCCESS, 0, True) ; Return false if words were processed.
EndFunc   ;==>__ScriptWordCheck

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __SpellCheck_CreateStruct
; Description ...: Retrieves a Struct.
; Syntax ........: __SpellCheck_CreateStruct($sStructName)
; Parameters ....: $sStructName	- a string value. Name of structure to create.
; Return values .:Success: Structure.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $sStructName Value not a string
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.ServiceManager" Object
;				   @Error 2 @Extended 2 Return 0 = Error retrieving requested Structure.
;				   --Success--
;				   @Error 0 @Extended 0 Return Structure = Success. Property Structure Returned
; Author ........: mLipok;
; Modified ......: donnyh13 - Added error checking.
; Remarks .......: From WriterDemo.au3 as modified by mLipok from WriterDemo.vbs found in the LibreOffice SDK examples.
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/204665-libreopenoffice-writer/?do=findComment&comment=1471711
; Example .......: No
; ===============================================================================================================================
Func __SpellCheck_CreateStruct($sStructName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __SpellCheck_ComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__S4A_LO_SC_SUCCESS = 0, $__S4A_LO_SC_INPUT_ERROR = 1, $__S4A_LO_SC_INIT_ERROR = 2
	Local $oServiceManager, $tStruct

	If Not IsString($sStructName) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 1, __Print_To_Error("Structure Name called is not a String."))
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __Print_To_Error("Failed to create com.sun.star.ServiceManager."))
	$tStruct = $oServiceManager.Bridge_GetStruct($sStructName)
	If Not IsObj($tStruct) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 2, __Print_To_Error("Failed to create requested Structure."))

	Return SetError($__S4A_LO_SC_SUCCESS, 0, $tStruct)
EndFunc   ;==>__SpellCheck_CreateStruct

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __LOWriter_SetPropertyValue
; Description ...: Creates a property value struct object.
; Syntax ........: __LOWriter_SetPropertyValue($sName, $vValue)
; Parameters ....: $sName               - a string value. Property name.
;                  $vValue              - a variant value. Property value.
; Return values .:Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = Property $sName Value was not a string
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Properties Object failed to be created
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Property Object Returned
; Author ........: Leagnus, GMK
; Modified ......: donnyh13 - added CreateStruct function. Modified variable names.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __LOWriter_SetPropertyValue($sName, $vValue)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __SpellCheck_ComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__S4A_LO_SC_SUCCESS = 0, $__S4A_LO_SC_INPUT_ERROR = 1, $__S4A_LO_SC_INIT_ERROR = 2
	Local $tProperties

	If Not IsString($sName) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 1, __Print_To_Error("Property Name called is not a String."))
	$tProperties = __SpellCheck_CreateStruct("com.sun.star.beans.PropertyValue")
	If @error Or Not IsObj($tProperties) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __Print_To_Error("Failed to create a Property Structure."))
	$tProperties.Name = $sName
	$tProperties.Value = $vValue

	Return SetError($__S4A_LO_SC_SUCCESS, 0, $tProperties)
EndFunc   ;==>__LOWriter_SetPropertyValue

Func __Print_To_Error($sError = "")
	Local Const $__sError_File = @ScriptDir & "\Scite4AutoIt_LO_SpellChecker_ERROR.ini"
	Local Const $__S4A_LO_SC_SUCCESS = 0, $__S4A_LO_SC_INIT_ERROR = 2
	Local $hFile
	Local Const $FO_READ = 0, $FO_APPEND = 1, $FO_OVERWRITE = 2

	$hFile = FileOpen($__sError_File, $FO_READ)

	If ($hFile = -1) Then ; Create File
		$hFile = FileOpen($__sError_File, $FO_OVERWRITE)
		If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, 0)
	EndIf

	FileClose($hFile)

	If ($sError = "") Then ; Clear the error file.
		$hFile = FileOpen($__sError_File, $FO_OVERWRITE)
		If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, 0)

	ElseIf ($sError = Null) Then
		FileDelete($__sError_File)

	Else ; Write the error.
		$hFile = FileOpen($__sError_File, $FO_APPEND)
		If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, 0)
	EndIf

	FileWrite($hFile, $sError & @CRLF)

	FileFlush($hFile)

	FileClose($hFile)

	Return SetError($__S4A_LO_SC_SUCCESS, 0, 0)
EndFunc   ;==>__Print_To_Error

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __SpellCheck_ComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __SpellCheck_ComErrorHandler(Byref $oComError)
; Parameters ....: $oComError           - [in/out] an object. The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13 - Added ConsoleWrite option.
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __SpellCheck_ComErrorHandler(ByRef $oComError)
	#forceref $oComError
	__Print_To_Error("A COM Error was thrown.")
	; This function does nothing.
EndFunc   ;==>__SpellCheck_ComErrorHandler
