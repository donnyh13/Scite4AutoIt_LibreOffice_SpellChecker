#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

Global $bReturn

If ($CmdLine[0] > 0) Then
	__Print_To_Error()

	$bReturn = _SpellCheck($CmdLine[1], $CmdLine[2], $CmdLine[3], $CmdLine[4], $CmdLine[5])

	If @error Then ; Something went wrong.
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

Func _SpellCheck($sWordToCheck, $sLanguage, $sCountry, $bReturnWords, $iMaxSuggestions)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __SpellCheck_ComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__S4A_LO_SC_SUCCESS = 0, $__S4A_LO_SC_INPUT_ERROR = 1, $__S4A_LO_SC_INIT_ERROR = 2, $__S4A_LO_SC_PROCESS_ERROR = 3
	Local Const $FO_READ = 0, $FO_OVERWRITE = 2, $FO_CREATEPATH = 8
	Local $sSpCheckFile = @ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini"
	Local $oServiceManager, $oSpellChecker
	Local $aEmptyArgs[1]
	Local $asLang[2], $asCountry[2]
	Local $atLocale[0]
	Local $tLocale
	Local $vReturn
	Local $hFile

	If Not IsString($sWordToCheck) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 1, __Print_To_Error("Word called to check is not a String."))
	If Not IsString($sLanguage) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 2, __Print_To_Error("Language code called is not a string."))
	If Not IsString($sCountry) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 3, __Print_To_Error("Country code called is not a string."))
	If IsString($bReturnWords) And ($bReturnWords = "true") Then
		$bReturnWords = True

	ElseIf IsString($bReturnWords) And ($bReturnWords = "false") Then
		$bReturnWords = False

	EndIf

	If Not IsBool($bReturnWords) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 4, __Print_To_Error("ReturnWords parameter called is is not a Boolean."))
	If StringRegExp($iMaxSuggestions, "[^0-9]") Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 5, __Print_To_Error("Maximum Suggestion per Language parameter called is is not a number."))

	$sCountry = StringUpper($sCountry)
	$sLanguage = StringLower($sLanguage)
	$iMaxSuggestions = Int($iMaxSuggestions)

	If ($iMaxSuggestions < 1) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 6, __Print_To_Error("Maximum Suggestion per Language parameter called is less than 1."))

	If StringInStr($sLanguage, ";") Then
		$asLang = StringSplit($sLanguage, ";")
		If @error Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 7, __Print_To_Error("Failed to split Language codes."))

		For $i = 1 To $asLang[0]
			If ((StringLen($asLang[$i]) <> 2) And (StringLen($asLang[$i]) <> 3)) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 8, __Print_To_Error("Language code " & $asLang[$i] & " called is not 2 or 3 characters long."))
		Next

	Else
		If ((StringLen($sLanguage) <> 2) And (StringLen($sLanguage) <> 3)) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 9, __Print_To_Error("Language code called is not 2 or 3 characters long."))
		$asLang[0] = 1
		$asLang[1] = $sLanguage

	EndIf

	If StringInStr($sCountry, ";") Then
		$asCountry = StringSplit($sCountry, ";")
		If @error Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 10, __Print_To_Error("Failed to split Country codes."))

		For $i = 1 To $asCountry[0]
			If (StringLen($asCountry[$i]) <> 2) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 11, __Print_To_Error("Country code " & $asCountry[$i] & " called is not 2 characters long."))
		Next

	Else
		If (StringLen($sCountry) <> 2) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 12, __Print_To_Error("Country code called is not 2 characters long."))
		$asCountry[0] = 1
		$asCountry[1] = $sCountry

	EndIf

	If ($asLang[0] <> $asCountry[0]) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 13, __Print_To_Error("Country Codes and Language codes contain unequal amount of values."))

	If Not FileExists($sSpCheckFile) Then
		$hFile = FileOpen($sSpCheckFile, $FO_CREATEPATH) ; Make sure file exists.
		If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __Print_To_Error("Failed to open Scite4AutoIt_LO_SpellChecker.ini. #1"))
		FileClose($hFile)
	EndIf

	ReDim $atLocale[$asLang[0]]

	For $i = 1 To $asLang[0]
		$tLocale = __SpellCheck_CreateStruct("com.sun.star.lang.Locale")
		If @error Then Return SetError($__S4A_LO_SC_INIT_ERROR, 2, __Print_To_Error("Failed to create com.sun.star.lang.Locale Structure."))

		$tLocale.Language = $asLang[$i]
		$tLocale.Country = $asCountry[$i]
		$atLocale[$i - 1] = $tLocale

	Next

	$aEmptyArgs[0] = __LOWriter_SetPropertyValue("", "")
	If @error Then Return SetError($__S4A_LO_SC_INIT_ERROR, 3, __Print_To_Error("Failed to create Property Structure."))

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If @error Then Return SetError($__S4A_LO_SC_INIT_ERROR, 4, __Print_To_Error("Failed to create com.sun.star.ServiceManager Object."))

	$oSpellChecker = $oServiceManager.createInstance("com.sun.star.linguistic2.SpellChecker")
	If Not IsObj($oSpellChecker) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 5, __Print_To_Error("Failed to create com.sun.star.linguistic2.SpellChecker Object."))

	For $i = 0 To UBound($atLocale) - 1
		If Not $oSpellChecker.hasLocale($atLocale[$i]) Then Return SetError($__S4A_LO_SC_PROCESS_ERROR, 1, __Print_To_Error("Language (" & $atLocale[$i].Language() & ") and Country (" & $atLocale[$i].Country() & ") combination is not valid."))
	Next

	If ($bReturnWords = True) Then ; If Return Words = True, then I am checking a single word.
		$hFile = FileOpen($sSpCheckFile, $FO_OVERWRITE)
		If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __Print_To_Error("Failed to open Scite4AutoIt_LO_SpellChecker.ini. #2"))
		FileFlush($hFile)

		$vReturn = __SingleWordCheck($sWordToCheck, $hFile, $oSpellChecker, $atLocale, $aEmptyArgs, $iMaxSuggestions)
		Return SetError(@error, FileClose($hFile), $vReturn)

	Else ; Entire Script check.
		$hFile = FileOpen($sSpCheckFile, $FO_READ)
		If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __Print_To_Error("Failed to open Scite4AutoIt_LO_SpellChecker.ini. #3"))

		$vReturn = __ScriptWordCheck($hFile, $oSpellChecker, $atLocale, $aEmptyArgs)
		Return SetError(@error, FileClose($hFile), $vReturn)

	EndIf

EndFunc   ;==>_SpellCheck


Func __SingleWordCheck($sWordToCheck, ByRef $hFile, ByRef $oSpellChecker, ByRef $atLocale, ByRef $aEmptyArgs, $iMaxSuggestions)
	Local Const $__S4A_LO_SC_SUCCESS = 0, $__S4A_LO_SC_INIT_ERROR = 2
	Local $oSpell
	Local $iCount = 0
	Local $asArray[0]
	Local $aasWords[UBound($atLocale)]

	For $i = 0 To UBound($atLocale) - 1

		If $oSpellChecker.isValid($sWordToCheck, $atLocale[$i], $aEmptyArgs) Then Return SetError($__S4A_LO_SC_SUCCESS, 0, True)

		$oSpell = $oSpellChecker.Spell($sWordToCheck, $atLocale[$i], $aEmptyArgs)
		If Not IsObj($oSpell) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __Print_To_Error("Failed to retrieve Spelling Object."))

		If ($oSpell.getAlternativesCount() > 0) Then
			$asArray = $oSpell.getAlternatives()
			If Not IsArray($asArray) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 2, __Print_To_Error("Failed to retrieve array of Alternative words."))

			$aasWords[$i] = $asArray

		EndIf
	Next

	__DuplicateWordScan($aasWords)

	For $i = 0 To UBound($aasWords) - 1
		If IsArray($aasWords[$i]) Then

			For $j = 0 To UBound($aasWords[$i]) - 1
				If IsString(($aasWords[$i])[$j]) Then
					FileWrite($hFile, ($aasWords[$i])[$j] & @CRLF)
					$iCount += 1
					If ($iCount >= $iMaxSuggestions) Then ExitLoop
				EndIf
			Next

		EndIf
	Next

	FileFlush($hFile)

	Return SetError($__S4A_LO_SC_SUCCESS, 0, False)
EndFunc   ;==>__SingleWordCheck

Func __ScriptWordCheck(ByRef $hFile, ByRef $oSpellChecker, ByRef $atLocale, ByRef $aEmptyArgs)
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

	For $j = 0 To UBound($atLocale) - 1
		For $i = 0 To $iLines - 1
			If IsString($asWords[$i]) Then
				$sWordToCheck = StringLeft($asWords[$i], StringInStr($asWords[$i], @TAB) - 1)
				If $oSpellChecker.isValid($sWordToCheck, $atLocale[$j], $aEmptyArgs) Then $asWords[$i] = 0 ; Overwrite the word with 0 to indicate it is spelled correctly.
			EndIf
		Next
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

Func __DuplicateWordScan(ByRef $aasWords)
	Local Const $__S4A_LO_SC_SUCCESS = 0
	Local $asWords[0]

	If (UBound($aasWords) > 1) Then
		For $i = 0 To UBound($aasWords) - 1

			If IsArray($aasWords[$i]) Then

				For $j = 0 To UBound($aasWords[$i]) - 1

					If IsString(($aasWords[$i])[$j]) Then

						For $k = $i + 1 To UBound($aasWords) - 1
							For $m = 0 To UBound($aasWords[$k]) - 1
								If IsString(($aasWords[$k])[$m]) And (($aasWords[$i])[$j] == ($aasWords[$k])[$m]) Then
									$asWords = $aasWords[$k]
									$asWords[$m] = 0
									$aasWords[$k] = $asWords
									ExitLoop
								EndIf
							Next
						Next
					EndIf

				Next

			EndIf

		Next
	EndIf

	Return SetError($__S4A_LO_SC_SUCCESS, 0, 1)
EndFunc   ;==>__DuplicateWordScan

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
