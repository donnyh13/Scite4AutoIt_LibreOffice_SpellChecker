#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#Region ; *** Dynamically added Include files ***
#include <FileConstants.au3>                                 ; added:11/09/23 13:35:59
#include <Date.au3>                                          ; added:11/17/23 08:52:28
#include <File.au3>                                          ; added:12/06/23 17:23:15
#include <AutoItConstants.au3>                               ; added:12/07/23 06:18:57
#EndRegion ; *** Dynamically added Include files ***

#include <Array.au3>

Func __RetrieveProps($oObj)
Local Const $iURLFrameCreate = 8 ;frame will be created if not found
Local $iError = 0
	Local $oDoc, $oServiceManager, $oDesktop, $oScript
	Local $atProperties[3], $avParamArray[1], $aDummyArray[0], $asReturn[0]
	Local $vProperty
	Local $sFileURL = "private:factory/swriter"

ObjEvent("AutoIt.Error", __SpellCheck_ComErrorHandler)

If Not IsObj($oObj) Then Return SetError(1,1,0)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError(2, 3, 0)
	$oDesktop = $oServiceManager.createInstance("com.sun.star.frame.Desktop")
	If Not IsObj($oDesktop) Then Return SetError(2, 4, 0)

$vProperty = __LOWriter_SetPropertyValue("Hidden", True)
If @error Then 	$iError = BitOR($iError,1)
If Not BitAND($iError,1) Then $atProperties[0] = $vProperty

$vProperty = __LOWriter_SetPropertyValue("MacroExecutionMode", 4)
If @error Then 	$iError = BitOR($iError,2)
If Not BitAND($iError,2) Then $atProperties[1] = $vProperty

$vProperty = __LOWriter_SetPropertyValue("ReadOnly", True)
If @error Then 	$iError = BitOR($iError,4)
If Not BitAND($iError,4) Then $atProperties[2] = $vProperty

$oDoc = $oDesktop.loadComponentFromURL($sFileURL, "_default", $iURLFrameCreate, $atProperties)
If Not IsObj($oDoc) Then Return SetError(2,5,0)

$oScript = $oDoc.getScriptProvider().getScript("vnd.sun.star.script:Debug.GetDebug.DisplayDbgInfo?language=Basic&location=application")
If Not IsObj($oScript) Then Return SetError(2,2,0);WorkAround method.

$avParamArray[0] = $oObj

$asReturn = $oScript.Invoke($avParamArray,$aDummyArray, $aDummyArray)

$oDoc.Close(True)

For $i = 0 To UBound($asReturn) - 1
MsgBox(0,"Properties",$asReturn[$i])
Next

Return ($iError > 0) ? SetError(3,$iError,$oDoc) : SetError(0,1,$oDoc)

EndFunc


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

	Local $tProperties

	If Not IsString($sName) Then Return SetError(1, 1, 0)
	$tProperties = __SpellCheck_CreateStruct("com.sun.star.beans.PropertyValue")
	If @error Or Not IsObj($tProperties) Then Return SetError(2, 1, 0)
	$tProperties.Name = $sName
	$tProperties.Value = $vValue

	Return SetError(0, 0, $tProperties)
EndFunc   ;==>__LOWriter_SetPropertyValue

;~ MsgBox(0,"",$CmdLine[0])
Global $bReturn
Global $iError, $iExtended
;~ ConsoleWrite("Yes" & @CRLF)
;~ MsgBox(0,"",$CmdLine[0])
If ($CmdLine[0] > 0) Then
	If ($CmdLine[0] = 4) Then
	$bReturn = _SpellCheck($CmdLine[1], $CmdLine[2], $CmdLine[3], $CmdLine[4])
	$iError = @error
	$iExtended = @extended

	Else
	$bReturn = _SpellCheck($CmdLine[1], $CmdLine[2], $CmdLine[3])
	$iError = @error
	$iExtended = @extended
	EndIf

		If ($iError > 0) Then
		Exit 2
	ElseIf ($bReturn = False) Then
		Exit 1

		Else
Exit 0
EndIf
;~ MsgBox(0,"",$CmdLine[1])

EndIf




Exit 0 ;0 = True, anything else = Null in script

;~ _SpellCheck("Hoshel")
;~ ConsoleWrite(@error & @CRLF & @extended & @CRLF)

;~ ConsoleWrite(_OOo_SpellChecker("fim","en","US",True) & @CRLF)
;~ Local $aArray = _OOo_SpellChecker("Hoshel","en","US",True)
;~ _ArrayDisplay($aArray)

Func _SpellCheck($sWordToCheck, $sLanguage = "en", $sCountry = "US", $bReturnWords = False)
Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __SpellCheck_ComErrorHandler)
	#forceref $oCOM_ErrorHandler

Local $oServiceManager, $oSpellChecker
Local $aEmptyArgs[1]
Local $tLocale
Local $vReturn
Local $hFile

If Not IsString($sWordToCheck) Then Return SetError(1,1,0)
If Not IsString($sLanguage) Or (StringLen($sLanguage) <> 2) Then Return SetError(1,2,0)
If Not IsString($sCountry) Or (StringLen($sCountry) <> 2) Then Return SetError(1,3,0)
If Not IsBool($bReturnWords) Then Return SetError(1,4,0)

If Not FileExists(@ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini") Then
$hFile = FileOpen(@ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini",$FO_CREATEPATH); Make sure file exists.
If ($hFile = -1) Then Return SetError(2, 1, 0)
FileClose($hFile)
EndIf

;~ If ($sWordToCheck = "##") Then Exit MsgBox(0,"","")

$sLanguage = StringLower($sLanguage)
$sCountry = StringUpper($sCountry)

$tLocale = __SpellCheck_CreateStruct("com.sun.star.lang.Locale")
If @error Then Return SetError(2,2,0)
$tLocale.Country = $sCountry
$tLocale.Language = $sLanguage

$aEmptyArgs[0] = __LOWriter_SetPropertyValue("","")
If @error Then Return SetError(2,3,0)

$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
If @error Then Return SetError(2,4,0)

$oSpellChecker = $oServiceManager.createInstance("com.sun.star.linguistic2.SpellChecker")
	If Not IsObj($oSpellChecker) Then Return SetError(2,5,0)

 If Not $oSpellChecker.hasLocale($tLocale) Then Return SetError(3, 1, 0)

If ($bReturnWords = True) Then; If Return Words = True, then I am checking a single word.
$hFile = FileOpen(@ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini",$FO_OVERWRITE)
If ($hFile = -1) Then Return SetError(2, 1, 0)
FileFlush($hFile)
;~ FileClose($hFile)
$vReturn = __SingleWordCheck($sWordToCheck, $hFile, $oSpellChecker, $tLocale, $aEmptyArgs)
FileClose($hFile)
Return SetError(@error, @extended, $vReturn)

Else; Entire Script check.
	$hFile = FileOpen(@ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini",$FO_READ)
	If ($hFile = -1) Then Return SetError(2, 1, 0)
$vReturn = __ScriptWordCheck($hFile, $oSpellChecker, $tLocale, $aEmptyArgs)
FileClose($hFile)
Return SetError(@error, @extended, $vReturn)

EndIf





;~ If $oSpellChecker.isValid($sWordToCheck,$tLocale,$aEmptyArgs) Then Return True

;~ If ($bReturnWords = False) Then Return False
;~ ConsoleWrite($oSpellChecker.isValid($sWordToCheck,$tLocale,$aEmptyArgs) & @CRLF)

;~ $vReturn = $oSpellChecker.Spell($sWordToCheck, $tLocale, $aEmptyArgs)

;~ ConsoleWrite($vReturn.getAlternativesCount() & @CRLF)
;~ Local $anArray = $vReturn.getAlternatives()
;~ _ArrayDisplay($anArray)


;~ If (UBound($anArray) > 0) Then
;~ 	Local $hFile = FileOpen(@ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini",$FO_OVERWRITE)



;~ 	For $i = 0 To UBound($anArray) -1
;~ FileWrite($hFile,$anArray[$i] & @CRLF)
;~ Next

;~ FileFlush($hFile)

;~ FileClose($hFile)


;~ 	EndIf


;~ If ($vReturn = Null) Then Return SetError(0,0,True)

;~ If IsArray($vReturn) Then _ArrayDisplay($vReturn)
;~ ConsoleWrite($vReturn.getAlternativesCount() & @CRLF)




EndFunc


Func __SingleWordCheck($sWordToCheck, ByRef $hFile, ByRef $oSpellChecker, ByRef $tLocale, ByRef $aEmptyArgs)
Local $vReturn
Local $anArray

If $oSpellChecker.isValid($sWordToCheck,$tLocale,$aEmptyArgs) Then Return True

;~ If ($bReturnWords = False) Then Return False
;~ ConsoleWrite($oSpellChecker.isValid($sWordToCheck,$tLocale,$aEmptyArgs) & @CRLF)

$vReturn = $oSpellChecker.Spell($sWordToCheck, $tLocale, $aEmptyArgs)

;~ ConsoleWrite($vReturn.getAlternativesCount() & @CRLF)
$anArray = $vReturn.getAlternatives()
;~ _ArrayDisplay($anArray)


If (UBound($anArray) > 0) Then
;~ 	Local $hFile = FileOpen(@ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini",$FO_OVERWRITE)

	For $i = 0 To UBound($anArray) -1
FileWrite($hFile,$anArray[$i] & @CRLF)
Next

FileFlush($hFile)

	EndIf


;~ If ($vReturn = Null) Then Return SetError(0,0,True)

;~ If IsArray($vReturn) Then _ArrayDisplay($vReturn)
;~ ConsoleWrite($vReturn.getAlternativesCount() & @CRLF)

Return False
EndFunc

Func __ScriptWordCheck(ByRef $hFile, ByRef $oSpellChecker, ByRef $tLocale, ByRef $aEmptyArgs)
;~ Local $vReturn
Local $asWords[0]
Local $iLines = 0, $iCount = 0
Local $sWordToCheck = ""

$asWords = FileReadToArray($hFile)
If @error Then Return SetError(3, 2, 0)
$iLines = @extended

FileClose($hFile)

$hFile = FileOpen(@ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini",$FO_OVERWRITE)
If ($hFile = -1) Then Return SetError(2, 1, 0)
FileFlush($hFile)

For $i = 0 To $iLines - 1
$sWordToCheck = StringLeft($asWords[$i],StringInStr($asWords[$i],@TAB) - 1)

If $oSpellChecker.isValid($sWordToCheck,$tLocale,$aEmptyArgs) Then $asWords[$i] = 0

Next

;~ If ($bReturnWords = False) Then Return False
;~ ConsoleWrite($oSpellChecker.isValid($sWordToCheck,$tLocale,$aEmptyArgs) & @CRLF)

;~ $vReturn = $oSpellChecker.Spell($sWordToCheck, $tLocale, $aEmptyArgs)

;~ ConsoleWrite($vReturn.getAlternativesCount() & @CRLF)
;~ $anArray = $vReturn.getAlternatives()
;~ _ArrayDisplay($anArray)


For $i = 0 To $iLines - 1
;~ 	Local $hFile = FileOpen(@ScriptDir & "\Scite4AutoIt_LO_SpellChecker.ini",$FO_OVERWRITE)

;~ 	For $i = 0 To UBound($anArray) -1
If IsString($asWords[$i]) Then
	FileWrite($hFile,$asWords[$i] & @CRLF)
	$iCount += 1
	EndIf
;~ Next
Next

FileFlush($hFile)


;~ 	EndIf

;~ If ($vReturn = Null) Then Return SetError(0,0,True)

;~ If IsArray($vReturn) Then _ArrayDisplay($vReturn)
;~ ConsoleWrite($vReturn.getAlternativesCount() & @CRLF)

Return ($iCount > 0) ? False : True
EndFunc



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

	Local $oServiceManager, $tStruct

	If Not IsString($sStructName) Then Return SetError(1, 1, 0)
	$oServiceManager = ObjCreate("com.sun.star.ServiceManager")
	If Not IsObj($oServiceManager) Then Return SetError(2, 1, 0)
	$tStruct = $oServiceManager.Bridge_GetStruct($sStructName)
	If Not IsObj($tStruct) Then Return SetError(2, 2, 0)

	Return SetError(0, 0, $tStruct)
EndFunc

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
	; This function does nothing.

;~ #cs
				ConsoleWrite("!--COM Error-Begin--" & @CRLF & _
						"Number: 0x" & Hex($oComError.number, 8) & @CRLF & _
						"WinDescription: " & $oComError.windescription & @CRLF & _
						"Source: " & $oComError.source & @CRLF & _
						"Error Description: " & $oComError.description & @CRLF & _
						"HelpFile: " & $oComError.helpfile & @CRLF & _
						"HelpContext: " & $oComError.helpcontext & @CRLF & _
						"LastDLLError: " & $oComError.lastdllerror & @CRLF & _
						"At line: " & $oComError.scriptline & @CRLF & _
						"!--COM-Error-End--" & @CRLF)
;~ #ce

EndFunc




;Add auto Local?

;Can get list of Locals?

#cs mLipok Example

#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <WinAPILocale.au3>

Global $oMyError = ObjEvent('AutoIt.Error', '_ErrFunc')

_Example()
Func _Example()
    Local $sLocale = 'pl'
    Local $sCountry = 'PL'
    Local $sWord = InputBox("Spell Checker", "Type single word to check (empty for end) ?")
    While $sWord <> ""
        Global $vSpell = _OOo_SpellChecker($sWord, $sLocale, $sCountry)
        If @error Then
            MsgBox($MB_ICONERROR, '_OOo_SpellChecker', '@error = ' & @error & @CRLF & '@extended = ' & @extended)
        ElseIf $vSpell Then
            MsgBox($MB_OK, "Spell Checker", $sWord & " is valid")
        Else
            If MsgBox($MB_YESNO, "Spell Checker", $sWord & " is NOT valid. Would you like to see alternatives?") = $IDYES Then
                Global $aAlternatives = _OOo_SpellChecker($sWord, $sLocale, $sCountry, True)
                _ArrayDisplay($aAlternatives, "Spell Checker")
            EndIf
        EndIf
        $sWord = InputBox("Spell Checker", "File to check (empty for end)?")
    WEnd
EndFunc

#ce




#cs Pitonyak macros

Listing 5.57: Spell, hyphenate, and use a thesaurus.
Sub SpellCheckExample
  Dim s() As Variant
  Dim vReturn As Variant, i As Integer
  Dim emptyArgs(0) As New com.sun.star.beans.PropertyValue
  Dim aLocale As New com.sun.star.lang.Locale


  aLocale.Language = "en"
  aLocale.Country = "US"

  s = Array("hello", "anesthesiologist", _
         "PNEUMONOULTRAMICROSCOPICSILICOVOLCANOCONIOSIS", _
         "Pitonyak", "misspell")

  '*********Spell Check Example!
  Dim vSpeller As Variant
  vSpeller = createUnoService("com.sun.star.linguistic2.SpellChecker")
  'Use vReturn = vSpeller.spell(s, aLocale, emptyArgs()) if you want options!
  For i = LBound(s()) To UBound(s())
    vReturn = vSpeller.isValid(s(i), aLocale, emptyArgs())
    MsgBox "Spell check on " & s(i) & " returns " & vReturn
  Next

  '******Hyphenation Example!
  Dim vHyphen As Variant
  vHyphen = createUnoService("com.sun.star.linguistic2.Hyphenator")
  For i = LBound(s()) To UBound(s())
    'vReturn = vHyphen.hyphenate(s(i), aLocale, 0, emptyArgs())
    vReturn = vHyphen.createPossibleHyphens(s(i), aLocale, emptyArgs())
    If IsNull(vReturn) Then
      'hyphenation is probablly off in the configuration
      MsgBox "Hyphenating " & s(i) & " returns null"
    Else
      MsgBox "Hyphenating " & s(i) & _
                    " returns " & vReturn.getPossibleHyphens()
    End If
  Next

  '******Thesaurus Example!
  Dim vThesaurus As Variant, j As Integer, k As Integer
  vThesaurus = createUnoService("com.sun.star.linguistic2.Thesaurus")
    s = Array("hello", "stamp", "cool")
  For i = LBound(s()) To UBound(s())
    vReturn = vThesaurus.queryMeanings(s(i), aLocale, emptyArgs())
    If UBound(vReturn) < 0 Then
      Print "Thesaurus found nothing for " & s(i)
    Else
      Dim sTemp As String
      sTemp = "Hyphenated " & s(i)
      For j = LBound(vReturn) To UBound(vReturn)
        sTemp = sTemp & Chr(13) & "Meaning = " & _
                        vReturn(j).getMeaning() & Chr(13)
        Dim vSyns As Variant
        vSyns = vReturn(j).querySynonyms()

     For k = LBound(vSyns) To UBound(vSyns)
          sTemp = sTemp & vSyns(k) & " "
        Next
        sTemp = sTemp & Chr(13)
      Next
      MsgBox sTemp
    End If
  Next
End Sub
#ce

#cs mLipok Mod of GMK
Func _OOo_SpellChecker($sWord, $sLocale = 'en', $sCountry = 'US', $bShowAlternatives = False)
    Local Enum _
            $eErr_Success, _
            $eErr_GeneralError, _
            $eErr_InvalidType, _
            $eErr_InvalidValue, _
            $eErr_NoMatch

    If Not IsString($sWord) Then Return SetError($eErr_InvalidType, 1, 0)
    If StringRegExp($sWord, '\d|\s') Or $sWord = '' Then Return SetError($eErr_InvalidValue, 1, 0)
    If Not IsString($sLocale) Then Return SetError($eErr_InvalidType, 2, 0)
    If StringLen($sLocale) <> 2 Then Return SetError($eErr_InvalidValue, 2, 0)
    If Not IsString($sCountry) Then Return SetError($eErr_InvalidType, 3, 0)
    If StringLen($sCountry) <> 2 Then Return SetError($eErr_InvalidValue, 3, 0)

    Local $oSM = ObjCreate('com.sun.star.ServiceManager')
    If Not IsObj($oSM) Then Return SetError($eErr_GeneralError, 10, 0)

    Local $oLocale = $oSM.Bridge_GetStruct('com.sun.star.lang.Locale')
    If Not IsObj($oLocale) Then Return SetError($eErr_GeneralError, 11, 0)

    $oLocale.Language = $sLocale
    $oLocale.Country = $sCountry
    Local $oLinguService = $oSM.createInstance('com.sun.star.linguistic2.LinguServiceManager')
    If Not IsObj($oLinguService) Then Return SetError($eErr_GeneralError, 12, 0)

    Local $oSpellChecker = $oLinguService.getSpellChecker
    If Not IsObj($oSpellChecker) Then Return SetError($eErr_GeneralError, 13, 0)
    If Not $oSpellChecker.hasLocale($oLocale) Then Return SetError($eErr_NoMatch, 2, 0)

    Local $oPropertyValue = $oSM.Bridge_GetStruct('com.sun.star.beans.PropertyValue')
    If Not IsObj($oPropertyValue) Then Return SetError($eErr_GeneralError, 14, 0)

    Local $aPropertyValue[1] = [$oPropertyValue]
    Local $nRandom = Random(0, 0.5)
    Local $bReturn = $oSpellChecker.isValid($sWord, $nRandom, $aPropertyValue)
    If @error Then Return SetError($eErr_GeneralError, 15, 0)

    If Not $bReturn And $bShowAlternatives Then
        Local $oSpell = $oSpellChecker.spell($sWord, $nRandom, $aPropertyValue)
        If Not IsObj($oSpell) Then Return SetError($eErr_GeneralError, 16, 0)

        Local $aReturn = $oSpell.getAlternatives()
        Local $iAlternatives = UBound($aReturn)
        ReDim $aReturn[$iAlternatives + 1]
        $iAlternatives += 1
        For $i = $iAlternatives - 1 To 1 Step -1
            $aReturn[$i] = $aReturn[$i - 1]
        Next
        $aReturn[0] = $oSpell.getAlternativesCount()
        If @error Then Return SetError($eErr_GeneralError, 17, 0)
        Return SetError($eErr_Success, 0, $aReturn)
    EndIf

    Return SetError($eErr_Success, 0, $bReturn)
EndFunc   ;==>_OOo_SpellChecker
#ce


#cs Pitonyak OOME **


The spell-checker, hyphenation, and thesaurus all require a locale to function. They will not function,
however, if they aren’t properly configured. Use Tools | Options | Language Settings | Writing Aids to
configure these in OOo. The macro in Listing 313 obtains a SpellChecker, Hyphenator, and Thesaurus, all of
which require a Locale object.
Listing 313. Spell Check, Hyphenate, and Thesaurus.
Sub SpellCheckExample
  Dim s()          'Contains the words to check
  Dim vReturn      'Value returned from SpellChecker, Hyphenator, and Thesaurus
  Dim i As Integer 'Utility index variable


 Dim msg$         'Message string

  REM Although I create an empty argument array, I could also
  REM use Array() to return an empty array.
  Dim emptyArgs() as new com.sun.star.beans.PropertyValue

  Dim aLocale As New com.sun.star.lang.Locale
  aLocale.Language = "en"  'Use the English language
  aLocale.Country = "US"   'Use the United States as the country

  REM Words to check for spelling, hyphenation, and thesaurus
  s = Array("hello", "anesthesiologist",_
      "PNEUMONOULTRAMICROSCOPICSILICOVOLCANOCONIOSIS",_
      "Pitonyak", "misspell")

  REM *********Spell Check Example!
  Dim vSpeller As Variant
  vSpeller = createUnoService("com.sun.star.linguistic2.SpellChecker")
  'Use vReturn = vSpeller.spell(s, aLocale, emptyArgs()) if you want options!
  For i = LBound(s()) To UBound(s())
    vReturn = vSpeller.isValid(s(i), aLocale, emptyArgs())
    msg = msg & vReturn & " for " & s(i) & CHR$(10)
  Next
  MsgBox msg, 0, "Spell Check Words"
  msg = ""

  '******Hyphenation Example!
  Dim vHyphen As Variant
  vHyphen = createUnoService("com.sun.star.linguistic2.Hyphenator")
  For i = LBound(s()) To UBound(s())
    'vReturn = vHyphen.hyphenate(s(i), aLocale, 0, emptyArgs())
    vReturn = vHyphen.createPossibleHyphens(s(i), aLocale, emptyArgs())
    If IsNull(vReturn) Then
      'hyphenation is probablly off in the configuration
      msg = msg & " null for " & s(i) & CHR$(10)
    Else
      msg = msg & vReturn.getPossibleHyphens() & " for " & s(i) & CHR$(10)
    End If
  Next
  MsgBox msg, 0, "Hyphenate Words"
  msg = ""


 '******Thesaurus Example!
  Dim vThesaurus As Variant
  Dim j As Integer, k As Integer
  vThesaurus = createUnoService("com.sun.star.linguistic2.Thesaurus")
  s = Array("hello", "stamp", "cool")
  For i = LBound(s()) To UBound(s())
    vReturn = vThesaurus.queryMeanings(s(i), aLocale, emptyArgs())
    If UBound(vReturn) < 0 Then
      Print "Thesaurus found nothing for " & s(i)
    Else
      msg = "Word " & s(i) & " has the following meanings:" & CHR$(10)
      For j = LBound(vReturn) To UBound(vReturn)
        msg=msg & CHR$(10) & "Meaning = " & vReturn(j).getMeaning() & CHR$(10)
        msg = msg & Join(vReturn(j).querySynonyms(), " ") & CHR$(10)
      Next
      MsgBox msg, 0, "Althernate Meanings"
    End If
  Next
End Sub
It is possible to obtain the default locale that is configured for OOo. Laurent Godard, an active
OpenOffice.org volunteer, wrote the macro in Listing 314, which obtains OOo’s currently configured locale.
Listing 314. Currently configured language (Locale).
Sub OOoLang()
  'Retreives the running OOO version
  'Author : Laurent Godard
  'e-mail : listes.godard@laposte.net
  '
  Dim aSettings, aConfigProvider
  Dim aParams2(0) As new com.sun.star.beans.PropertyValue
  Dim sProvider$, sAccess$
  sProvider = "com.sun.star.configuration.ConfigurationProvider"
  sAccess   = "com.sun.star.configuration.ConfigurationAccess"
  aConfigProvider = createUnoService(sProvider)
  aParams2(0).Name = "nodepath"
  aParams2(0).Value = "/org.openoffice.Setup/L10N"
  aSettings = aConfigProvider.createInstanceWithArguments(sAccess, aParams2())

  Dim OOLangue as string
  OOLangue= aSettings.getbyname("ooLocale")    'en-US
  MsgBox "OOo is configured with Locale " & OOLangue, 0, "OOo Locale"
End Sub
The above macro is reputed to fail on AOO, but, the following should work correctly. The macro does
function correctly with LO. The macro shown below should work for both AOO and LO.
Listing 315. You can use the tools library to get the current locale as a string.
GlobalScope.BasicLibraries.loadLibrary( "Tools" )
Print GetRegistryKeyContent("org.openoffice.Setup/L10N",FALSE).getByName("ooLocale")


#ce

