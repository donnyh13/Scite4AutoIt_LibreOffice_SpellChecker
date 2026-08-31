#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

; Upon the user invoking either a single word spellcheck or a selection/full script spellcheck, the lua script will collect the word(s) and their start/stop positions.
; + If a multi-word spellcheck is performed, the lua script writes them and their start/stop positions to Scite4AutoIt_LO_SpellChecker.ini, then calls this AutoIt script with the parameters described below.
; + If a single word spellcheck is triggered, this AutoIt script is called with the word to check, plus the other parameters described below.
; Once this script is called, the parameters are processed. And the LibreOffice Service manager is initialized.
; + If a multi-word spellcheck, all the words are read from Scite4AutoIt_LO_SpellChecker.ini and spellchecked, any misspelled words are left in the ini file.
; + If a single word spellcheck, the word is processed, and a list of possible correct words are retrieved for each language option, then they are written to Scite4AutoIt_LO_SpellChecker.ini for the lua script.
; - If the au3 script encounters an error, the exit code is 2 and the error is written to Scite4AutoIt_LO_SpellChecker_ERROR.ini.
; - If the word or any of multiple words being checked are misspelled, the exit code is 1.
; - Otherwise exit code 0 means all words are correct.
; Once the au3 script is done, the lua script interprets the exit code, and proceeds as necessary, either reading the error file, or marking misspelled word(s), and offering spelling suggestions for a single word check.

; This script expects to be called with 5 parameters.
; - Param 1: The word to spellcheck, if performing a single word check. Otherwise ## is passed, and the words to check are found in the Scite4AutoIt_LO_SpellChecker.ini file.
; - Param 2: The 2 to 3 character ISO 639 Language Code.
; - Param 3: The 2 character long ISO 3166 Country Code. (Each Language and country code pair must be in the same order to pair with each other.)
; - Param 4: A string (either "single" or "multi"), whether the current check is a single or multi-word check.
; - Param 5: An Integer of the maximum number of suggested words to return per-language.

Global $bReturn

If ($CmdLine[0] = 5) Then
	; Execute the Spell Checker with the called parameters.
	$bReturn = _S4A_SpChk_SpellCheck($CmdLine[1], $CmdLine[2], $CmdLine[3], $CmdLine[4], $CmdLine[5])

	If @error Then     ; Something went wrong.
		Exit 2

	ElseIf $bReturn Then     ; Word(s) are incorrectly spelled.
		Exit 1

	Else     ; Word(s) are correctly spelled.
		Exit 0

	EndIf

Else ; If no parameters passed, exit and write an error.
	__S4A_SpChk_Print_To_Error("Wrong number of Parameters passed: " & $CmdLine[0])
	Exit 2

EndIf

; #FUNCTION# ====================================================================================================================
; Name ..........: _S4A_SpChk_SpellCheck
; Description ...: The Main Spell Checking Function.
; Syntax ........: _S4A_SpChk_SpellCheck($sWordToCheck, $sLanguage, $sCountry, $sCheckMode, $iMaxSuggestions)
; Parameters ....: $sWordToCheck        - The Word to check if I am checking a single word.
;                  $sLanguage           - The Language(s) to use to check the word(s).
;                  $sCountry            - The Country code(s) to use to check the word(s).
;                  $sCheckMode          - Either "single" or "multi", whether I'm checking a single word or multiple.
;                  $iMaxSuggestions     - The Max number of suggestions per language to return.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   @error 1 = Input error.
;				   @error 2 = Initialization error.
;				   @error 3 = Processing error.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _S4A_SpChk_SpellCheck($sWordToCheck, $sLanguage, $sCountry, $sCheckMode, $iMaxSuggestions)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __S4A_SpChk_ComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__S4A_LO_SC_INPUT_ERROR = 1, $__S4A_LO_SC_INIT_ERROR = 2, $__S4A_LO_SC_PROCESS_ERROR = 3
	Local Const $__FO_READ = 0, $__FO_OVERWRITE = 2, $FO_CREATEPATH = 8
	Local $sSpCheckFile = @ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini"
	Local $oServiceManager, $oSpellChecker
	Local $asLang[2], $asCountry[2]
	Local $atLocale[0]
	Local $tLocale
	Local $vReturn
	Local $hFile

	If Not IsString($sWordToCheck) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 1, __S4A_SpChk_Print_To_Error("Word called to check is not a String."))
	If Not IsString($sLanguage) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 2, __S4A_SpChk_Print_To_Error("Language code called is not a string."))
	If Not IsString($sCountry) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 3, __S4A_SpChk_Print_To_Error("Country code called is not a string."))
	If Not IsString($sCheckMode) Or (($sCheckMode <> "single") And ($sCheckMode <> "multi")) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 4, __S4A_SpChk_Print_To_Error("Check mode called is not a string, or not equal to 'single'/'multi'."))
	If StringRegExp($iMaxSuggestions, "[^\d]") Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 5, __S4A_SpChk_Print_To_Error("Maximum Suggestion per Language parameter called is not a number."))

	$sCountry = StringUpper($sCountry) ; Country codes are always uppercase, force them to be uppercase to prevent any issues.
	$sLanguage = StringLower($sLanguage) ; Language codes are always lowercase, force them to be lowercase to prevent any issues.
	$iMaxSuggestions = Int($iMaxSuggestions)

	If ($iMaxSuggestions < 1) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 6, __S4A_SpChk_Print_To_Error("Maximum Suggestion per Language parameter called with value less than 1."))

	; Parse and check the language and country codes and store them in separate arrays, they are to be separated by ";".
	$asLang = StringSplit($sLanguage, ";")
	If Not IsArray($asLang) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 7, __S4A_SpChk_Print_To_Error("Failed to split Language codes."))

	For $i = 1 To $asLang[0]
		If ((StringLen($asLang[$i]) <> 2) And (StringLen($asLang[$i]) <> 3)) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 8, __S4A_SpChk_Print_To_Error("Language code " & $asLang[$i] & " called is not 2 or 3 characters long."))
	Next

	$asCountry = StringSplit($sCountry, ";")
	If Not IsArray($asCountry) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 10, __S4A_SpChk_Print_To_Error("Failed to split Country codes."))

	For $i = 1 To $asCountry[0]
		If (StringLen($asCountry[$i]) <> 2) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 11, __S4A_SpChk_Print_To_Error("Country code " & $asCountry[$i] & " called is not 2 characters long."))
	Next

	If ($asLang[0] <> $asCountry[0]) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 13, __S4A_SpChk_Print_To_Error("Country Codes and Language codes contain unequal amount of values."))

	; Make sure the word list file exists.
	If Not FileExists($sSpCheckFile) Then
		$hFile = FileOpen($sSpCheckFile, $FO_CREATEPATH)
		If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __S4A_SpChk_Print_To_Error("Failed to open Scite4AutoIt_LO_SpellChecker.ini. #1"))
		FileClose($hFile)
	EndIf

	ReDim $atLocale[$asLang[0]]

	; Initialize the LibreOffice Service Manager, this is needed to create the Spell checking engine.
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If @error Then Return SetError($__S4A_LO_SC_INIT_ERROR, 3, __S4A_SpChk_Print_To_Error("Failed to create com.sun.star.ServiceManager Object."))

	; Create an instance of the Spell Checker Engine.
	$oSpellChecker = $oServiceManager.createInstance("com.sun.star.linguistic2.SpellChecker")
	If Not IsObj($oSpellChecker) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 4, __S4A_SpChk_Print_To_Error("Failed to create com.sun.star.linguistic2.SpellChecker Object."))

	; Create the LibreOffice Locale Struct for each Language/Country code pair and store it in an array for use in the spellchecker.
	For $i = 1 To $asLang[0]
		$tLocale = __S4A_SpChk_CreateStruct("com.sun.star.lang.Locale")
		If @error Then Return SetError($__S4A_LO_SC_INIT_ERROR, 2, __S4A_SpChk_Print_To_Error("Failed to create com.sun.star.lang.Locale Structure."))

		$tLocale.Language = $asLang[$i]
		$tLocale.Country = $asCountry[$i]

		; Make sure the Language/Country pairing is valid before adding it to the array.
		If Not $oSpellChecker.hasLocale($tLocale) Then Return SetError($__S4A_LO_SC_PROCESS_ERROR, 1, __S4A_SpChk_Print_To_Error("Language (" & $tLocale.Language() & ") and Country (" & $tLocale.Country() & ") combination is not valid."))

		$atLocale[$i - 1] = $tLocale
	Next

	If ($sCheckMode = "single") Then ; Single word check.
		$hFile = FileOpen($sSpCheckFile, $__FO_OVERWRITE) ; Clear the Scite4AutoIt_LO_SpellChecker.ini file.
		If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __S4A_SpChk_Print_To_Error("Failed to open Scite4AutoIt_LO_SpellChecker.ini. #2"))
		FileFlush($hFile)

		$vReturn = __S4A_SpChk_SingleWordCheck($sWordToCheck, $hFile, $oSpellChecker, $atLocale, $iMaxSuggestions)
		Return SetError(@error, FileClose($hFile), $vReturn)

	Else ; Multi-word check.
		$hFile = FileOpen($sSpCheckFile, $__FO_READ)
		If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __S4A_SpChk_Print_To_Error("Failed to open Scite4AutoIt_LO_SpellChecker.ini. #3"))

		$vReturn = __S4A_SpChk_ScriptWordCheck($hFile, $oSpellChecker, $atLocale)
		Return SetError(@error, FileClose($hFile), $vReturn)

	EndIf

EndFunc   ;==>_S4A_SpChk_SpellCheck

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __S4A_SpChk_SingleWordCheck
; Description ...: Spell Check a single word.
; Syntax ........: __S4A_SpChk_SingleWordCheck($sWordToCheck, ByRef $hFile, ByRef $oSpellChecker, ByRef $atLocale, $iMaxSuggestions)
; Parameters ....: $sWordToCheck        - The word to check.
;                  $hFile               - The file to write suggested words to.
;                  $oSpellChecker       - The Spell Checker Engine object.
;                  $atLocale            - Array of Locale Structures.
;                  $iMaxSuggestions     - The Maximum suggestions to return per language.
; Return values .: Success: Boolean. True if the word is misspelled.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   @error 2 = Initialization error.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __S4A_SpChk_SingleWordCheck($sWordToCheck, ByRef $hFile, ByRef $oSpellChecker, ByRef $atLocale, $iMaxSuggestions)
	Local Const $__S4A_LO_SC_SUCCESS = 0, $__S4A_LO_SC_INIT_ERROR = 2
	Local $oSpell
	Local $iCount = 0
	Local $asArray[0]
	Local $aEmptyArgs[0]
	Local $aasWords[UBound($atLocale)]

	; Cycle through each Locale to check the word.
	For $i = 0 To UBound($atLocale) - 1

		If $oSpellChecker.isValid($sWordToCheck, $atLocale[$i], $aEmptyArgs) Then Return SetError($__S4A_LO_SC_SUCCESS, 0, False)

		; If word is invalid, initiate the Spell engine with the locale so I can get spelling suggestions.
		$oSpell = $oSpellChecker.Spell($sWordToCheck, $atLocale[$i], $aEmptyArgs)
		If Not IsObj($oSpell) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __S4A_SpChk_Print_To_Error("Failed to retrieve Spelling Object."))

		; If there are spelling suggestions, retrieve them and store them in an array in an array to be written to the Scite4AutoIt_LO_SpellChecker.ini file.
		If ($oSpell.getAlternativesCount() > 0) Then
			$asArray = $oSpell.getAlternatives()
			If Not IsArray($asArray) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 2, __S4A_SpChk_Print_To_Error("Failed to retrieve array of Alternative words."))

			$aasWords[$i] = $asArray

		EndIf
	Next

	; Multiple dictionaries could return the same word suggestion.
	; Scan the suggested words Array for duplicated suggested words and remove them. Array is modified directly by ByRef.
	__S4A_SpChk_DuplicateWordScan($aasWords)

	; Write the suggested words to file.
	For $i = 0 To UBound($aasWords) - 1
		If IsArray($aasWords[$i]) Then

			For $j = 0 To UBound($aasWords[$i]) - 1
				If IsString(($aasWords[$i])[$j]) Then
					FileWrite($hFile, ($aasWords[$i])[$j] & @CRLF)
					$iCount += 1
					If ($iCount >= $iMaxSuggestions) Then ExitLoop ; If more suggestions than max suggestions desired, exit loop.
				EndIf
			Next

		EndIf
	Next

	FileFlush($hFile)

	Return SetError($__S4A_LO_SC_SUCCESS, 0, True)
EndFunc   ;==>__S4A_SpChk_SingleWordCheck

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __S4A_SpChk_ScriptWordCheck
; Description ...: Spell Check an entire Script.
; Syntax ........: __S4A_SpChk_ScriptWordCheck(ByRef $hFile, ByRef $oSpellChecker, ByRef $atLocale)
; Parameters ....: $hFile               - The File to read the words to spellcheck from.
;                  $oSpellChecker       - The Spell Checker Engine object.
;                  $atLocale            - Array of Locale Structures.
; Return values .: Success: Boolean. True if some or all words are misspelled.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   @error 2 = Initialization error.
;				   @error 3 = Processing error.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __S4A_SpChk_ScriptWordCheck(ByRef $hFile, ByRef $oSpellChecker, ByRef $atLocale)
	Local Const $__S4A_LO_SC_SUCCESS = 0, $__S4A_LO_SC_INIT_ERROR = 2, $__S4A_LO_SC_PROCESS_ERROR = 3
	Local Const $__FO_OVERWRITE = 2
	Local $asWords[0]
	Local $aEmptyArgs[0]
	Local $iLines = 0, $iMisspellCount = 0
	Local $sWordToCheck = "", $sSpCheckFile = @ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini"

	; Read the words to check to an array. Each line will contain the word to check, plus its start/stop position in the script file.
	; The word is first, then a @TAB is inserted, then the start position, another @TAB and the end position.
	$asWords = FileReadToArray($hFile)
	If @error Then Return SetError($__S4A_LO_SC_PROCESS_ERROR, 1, __S4A_SpChk_Print_To_Error("Failed to Read File to Array."))
	$iLines = @extended

	FileClose($hFile)

	; Open and clear the file.
	$hFile = FileOpen($sSpCheckFile, $__FO_OVERWRITE)
	If ($hFile = -1) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __S4A_SpChk_Print_To_Error("Failed to open Scite4AutoIt_LO_SpellChecker.ini. #4"))
	FileFlush($hFile)

	; Check all the words for each language.
	For $j = 0 To UBound($atLocale) - 1
		For $i = 0 To $iLines - 1
			If IsString($asWords[$i]) Then
				; The word is written to file as such: word @TAB Start-position @TAB End-position.
				$sWordToCheck = StringLeft($asWords[$i], StringInStr($asWords[$i], @TAB) - 1)
				If $oSpellChecker.isValid($sWordToCheck, $atLocale[$j], $aEmptyArgs) Then $asWords[$i] = 0 ; Overwrite the word entry with 0 to indicate it is spelled correctly.
			EndIf
		Next
	Next

	; Write the words and positions that are misspelled back to the file so the lua script can color them.
	For $i = 0 To $iLines - 1
		If IsString($asWords[$i]) Then
			FileWriteLine($hFile, $asWords[$i])
			$iMisspellCount += 1
		EndIf
	Next

	FileFlush($hFile)

	Return ($iMisspellCount = 0) ? SetError($__S4A_LO_SC_SUCCESS, 0, False) : SetError($__S4A_LO_SC_SUCCESS, 0, True) ; If no misspellings were found, return False.
EndFunc   ;==>__S4A_SpChk_ScriptWordCheck

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __S4A_SpChk_DuplicateWordScan
; Description ...: Check if the suggested word array contains duplicates, and remove them.
; Syntax ........: __S4A_SpChk_DuplicateWordScan(ByRef $aasWords)
; Parameters ....: $aasWords            - [in/out] an array of arrays of strings. The Array of Arrays containing strings to look for duplicates in.
; Return values .: 1
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __S4A_SpChk_DuplicateWordScan(ByRef $aasWords)
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
EndFunc   ;==>__S4A_SpChk_DuplicateWordScan

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __S4A_SpChk_CreateStruct
; Description ...: Retrieves a Struct.
; Syntax ........: __S4A_SpChk_CreateStruct($sStructName)
; Parameters ....: $sStructName	- The name of the LibreOffice structure to create.
; Return values .: Success: Structure.
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
Func __S4A_SpChk_CreateStruct($sStructName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __S4A_SpChk_ComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__S4A_LO_SC_SUCCESS = 0, $__S4A_LO_SC_INPUT_ERROR = 1, $__S4A_LO_SC_INIT_ERROR = 2
	Local $oServiceManager, $tStruct

	If Not IsString($sStructName) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 1, __S4A_SpChk_Print_To_Error("Structure Name called is not a String."))

	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __S4A_SpChk_Print_To_Error("Failed to create com.sun.star.ServiceManager."))

	$tStruct = $oServiceManager.Bridge_GetStruct($sStructName)
	If Not IsObj($tStruct) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 2, __S4A_SpChk_Print_To_Error("Failed to create requested Structure."))

	Return SetError($__S4A_LO_SC_SUCCESS, 0, $tStruct)
EndFunc   ;==>__S4A_SpChk_CreateStruct

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __S4A_SpChk_SetPropertyValue
; Description ...: Creates a property value struct object.
; Syntax ........: __S4A_SpChk_SetPropertyValue($sName, $vValue)
; Parameters ....: $sName               - a string value. Property name.
;                  $vValue              - a variant value. Property value.
; Return values .: Success: Object
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
Func __S4A_SpChk_SetPropertyValue($sName, $vValue)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __S4A_SpChk_ComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local Const $__S4A_LO_SC_SUCCESS = 0, $__S4A_LO_SC_INPUT_ERROR = 1, $__S4A_LO_SC_INIT_ERROR = 2
	Local $tProperties

	If Not IsString($sName) Then Return SetError($__S4A_LO_SC_INPUT_ERROR, 1, __S4A_SpChk_Print_To_Error("Property Name called is not a String."))

	$tProperties = __S4A_SpChk_CreateStruct("com.sun.star.beans.PropertyValue")
	If @error Or Not IsObj($tProperties) Then Return SetError($__S4A_LO_SC_INIT_ERROR, 1, __S4A_SpChk_Print_To_Error("Failed to create a Property Structure."))

	$tProperties.Name = $sName
	$tProperties.Value = $vValue

	Return SetError($__S4A_LO_SC_SUCCESS, 0, $tProperties)
EndFunc   ;==>__S4A_SpChk_SetPropertyValue

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __S4A_SpChk_Print_To_Error
; Description ...: Prints an error to an Error file.
; Syntax ........: __S4A_SpChk_Print_To_Error($sError)
; Parameters ....: $sError              - Default is "". The Error message to print to the error script.
; Return values .: None
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This functions writes an error message to a ini file so the lua script can output it to the console.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __S4A_SpChk_Print_To_Error($sError)
	Local Const $__sError_File = @ScriptDir & "\Scite4AutoIt_LO_SpellChecker_ERROR.ini"
	Local Const $__S4A_LO_SC_SUCCESS = 0

	FileWriteLine($__sError_File, $sError)

	Return SetError($__S4A_LO_SC_SUCCESS, 0, 0)
EndFunc   ;==>__S4A_SpChk_Print_To_Error

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __S4A_SpChk_ComErrorHandler
; Description ...: ComError Handler
; Syntax ........: __S4A_SpChk_ComErrorHandler(ByRef $oComError)
; Parameters ....: $oComError           - The Com Error Object passed by Autoit.Error.
; Return values .: None
; Author ........: mLipok
; Modified ......: donnyh13
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __S4A_SpChk_ComErrorHandler(ByRef $oComError)
	; Output the COM error to the error file, so that the lua script can output it to the console.
	__S4A_SpChk_Print_To_Error("A COM Error was thrown." & @CRLF & _
			"!--COM Error-Begin--" & @CRLF & _
			"Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
			"WinDescription: " & $oComError.windescription & @CRLF & _
			"Source: " & $oComError.source & @CRLF & _
			"Error Description: " & $oComError.description & @CRLF & _
			"HelpFile: " & $oComError.helpfile & @CRLF & _
			"HelpContext: " & $oComError.helpcontext & @CRLF & _
			"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
			"At line: " & $oComError.scriptline & @CRLF & _
			"!--COM-Error-End--" & @CRLF)
EndFunc   ;==>__S4A_SpChk_ComErrorHandler
