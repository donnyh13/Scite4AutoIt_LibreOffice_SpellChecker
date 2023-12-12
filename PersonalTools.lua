-- ==>LUA example file for personal LUA scripts.
-- ==>Copy this file to your %USERPROFILE% directory
-------------------------------------------------------------------------------
-- required line ... do not remove
PersonalTools = EventClass:new(Common)
-------------------------------------------------------------------------------

-- Single Word Spell-Check
-- to be able to use this script you add the following to your SciTEUSer.properties (without the leading "--"):
--#~ Single Word Spell-Check lua Script
--command.name.45.$(au3)=Sp-Check Current Word
--command.mode.45.$(au3)=subsystem:lua,savebefore:no
--command.shortcut.45.$(au3)=Ctrl+Shift+F
--command.45.$(au3)=InvokeTool PersonalTools.SingleWordCheck
--
--------------------------------------------------------------------------------
function PersonalTools:SingleWordCheck()
    local sSciteUserHome = props["SciteUserHome"]
    local sSpChkScript = sSciteUserHome .. "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_SpellChecker.au3"
    local sSpChkWordList = sSciteUserHome .. "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_SpellChecker.ini"
    local sSpChkErrorFile = sSciteUserHome ..  "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_SpellChecker_ERROR.ini"
    local sCurWord, sText, sLang, sCountry = "", "", "en", "US"
    local bReturn
    local iExit, iSignal, iCount, iSpChkIndicator, iOldIndic, iCurWordStart, iCurWordEnd = nil, nil, 0, 10, nil, nil, nil
    local aWords = {}
    local hFile
    local old_seperator

    if output.LinesOnScreen == 0 then scite.MenuCommand(IDM_TOGGLEOUTPUT) end
        output:AppendText("- Scite4AutoIt_LibreOffice_SpellChecker -\n")

        --Check if all necessary files exist.
        hFile = io.open(sSpChkScript, "r")
        if (hFile == nil) then
           output:AppendText("! Scite4AutoIt_LO_SpellChecker.au3 file missing. -- Aborting\n")
        return
    end
    io.close(hFile)

    hFile = io.open(sSpChkWordList, "r")
    if (hFile == nil) then
        output:AppendText("! Scite4AutoIt_LO_SpellChecker.ini file missing -- Creating file.\n")
        hFile = io.open(sSpChkWordList, "w")
    end
    io.close(hFile)

    if (props["personal.tools.Language"] == "") then sLang = "en" else sLang = props["personal.tools.Language"] end
    if (props["personal.tools.Country"] == "") then sCountry = "US" else sCountry = props["personal.tools.Country"] end

    iCurWordStart = editor:WordStartPosition(editor.CurrentPos, true)
    iCurWordEnd = editor:WordEndPosition(editor.CurrentPos, true)

    editor:SetTargetRange(iCurWordStart, iCurWordEnd)
    if (string.find(editor.TargetText,"[; %(%)%[\\//," .. '"' .. "]") == 1) and
        (string.find(editor.TargetText,"[%a]") ~= nil) then -- If there is a semicolon, space or beginning round bracket at the beginning of the word, adjust the selection to skip it.
        iCurWordStart = iCurWordStart + 1
        editor:SetTargetRange(iCurWordStart, iCurWordEnd)
    end

    if (string.find(editor.TargetText,"[ %.\\//]",string.len(editor.TargetText) - 1) ~= nil) and
        (string.find(editor.TargetText,"[%a]") ~= nil) then -- If there is a space or period at the end of the word, adjust the selection to skip it.
        iCurWordEnd = iCurWordEnd - 1
        editor:SetTargetRange(iCurWordStart, iCurWordEnd)
    end

    sText = editor.TargetText

    editor:SetTargetRange(iCurWordEnd, iCurWordEnd + 1)

    if (string.find(editor.TargetText, "'")) and editor:IsRangeWord(iCurWordStart, iCurWordEnd + 2) then
        iCurWordEnd = iCurWordEnd + 2
    end

    editor:SetTargetRange(iCurWordStart, iCurWordEnd)

    sCurWord = editor.TargetText

    if (sCurWord == '') then return output:AppendText("! Cursor is NOT located in a word.\n") end

    -- Clear any SpellChecking markings for the current word.
    iOldIndic = editor.IndicatorCurrent
    editor.IndicatorCurrent = iSpChkIndicator
    editor:IndicatorClearRange(iCurWordStart, iCurWordEnd - iCurWordStart)
    editor.IndicatorCurrent = iOldIndic

    bReturn, iExit, iSignal = os.execute('"' .. sSpChkScript .. '"' .. " " .. sCurWord .. " " .. sLang .. " " .. sCountry .. " " .. "true")

     if (iSignal == 1) then

        hFile = io.open(sSpChkWordList, "r")

        repeat
            sText = hFile:read("l")

            if (sText ~= nil) and (sText ~= "") then
                table.insert(aWords,sText)
                iCount = iCount + 1
            end

        until (sText == nil)

        io.close(hFile)

        if (iCount > 0) then
            editor:GotoPos(iCurWordEnd)

            old_seperator = editor.AutoCSeparator
            editor.AutoCSeparator = string.byte(';')
            editor:UserListShow(11, table.concat(aWords, ';'))
            editor.AutoCSeparator = old_seperator
            output:AppendText("> The word " .. '"' .. sCurWord .. '"' .. " is misspelled, found " .. iCount .. " suggestion(s)." .. "\n")

        else
            output:AppendText("! The word " .. '"' .. sCurWord .. '"' .. " is misspelled, but found no spelling suggestions." .. "\n")

        end

    elseif (iSignal == 0) then
        output:AppendText("+ The word " .. '"' .. sCurWord .. '"' .. " is spelled correctly." .. "\n")

    else -- Error of some form.
        output:AppendText("! Scite4AutoIt_LibreOffice_SpellChecker encountered an Error." .. "\n")
        hFile = io.open(sSpChkErrorFile, "r")
        if (hFile == nil) then
            output:AppendText("! Failed to open Error File." .. "\n")
            return
        end
        repeat
            sText = hFile:read("l")

            if (sText ~= nil) and (sText ~= "") then
                output:AppendText("! " .. sText .. "\n")
            end

        until (sText == nil)

        io.close(hFile)
        os.remove(sSpChkErrorFile)
    end
    output:AppendText("++ Scite4AutoIt_LibreOffice_SpellChecker completed. ++\n")
end

-- Entire Script Spell-Check
-- to be able to use this script you add the following to your SciTEUSer.properties (without the leading "--"):
--#~ Entire Script Spell-Check lua Script
--command.name.46.$(au3)=Sp-Check Entire Script
--command.mode.46.$(au3)=subsystem:lua,savebefore:no
--command.shortcut.46.$(au3)=Ctrl+Shift+G
--command.46.$(au3)=InvokeTool PersonalTools.CheckScript
--
--------------------------------------------------------------------------------
function PersonalTools:CheckScript()
    local sSciteUserHome = props["SciteUserHome"]
    local sSpChkScript = sSciteUserHome .. "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_SpellChecker.au3"
    local sSpChkWordList = sSciteUserHome .. "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_SpellChecker.ini"
    local sSpChkIgnoredWords = sSciteUserHome .. "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_IgnoredWords.ini"
    local sSpChkErrorFile = sSciteUserHome ..  "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_SpellChecker_ERROR.ini"
    local iLineStart, iLineEnd, iPos, iWordStart, iWordEnd, iCStyle, iOldIndic, iExit, iSignal, iFirstTab, iSpChkIndicator, iCount
    local sText, sLang, sCountry = "", "en", "US"
    local bReturn
    local asIgnoredWords = {}
    local asWordsToCheck = {}
    local hFile

    iSpChkIndicator = 10

    if output.LinesOnScreen == 0 then scite.MenuCommand(IDM_TOGGLEOUTPUT) end
    output:AppendText("- Scite4AutoIt_LibreOffice_SpellChecker -\n")

    --Check if all necessary files exist.
    hFile = io.open(sSpChkScript, "r")
    if (hFile == nil) then
        output:AppendText("! Scite4AutoIt_LO_SpellChecker.au3 file missing. -- Aborting\n")
        return
    end
    io.close(hFile)

    hFile = io.open(sSpChkWordList, "r")
    if (hFile == nil) then
        output:AppendText("! Scite4AutoIt_LO_SpellChecker.ini file missing -- Creating file.\n")
        hFile = io.open(sSpChkWordList, "w")
    end
    io.close(hFile)

    hFile = io.open(sSpChkIgnoredWords, "r")
    if (hFile == nil) then
        output:AppendText("! Scite4AutoIt_LO_IgnoredWords.ini file missing -- Creating file.\n")
        hFile = io.open(sSpChkWordList, "w")
    end
    io.close(hFile)

    if (props["personal.tools.Language"] ~= "") then sLang = props["personal.tools.Language"] end
    if (props["personal.tools.Country"] ~= "") then sCountry = props["personal.tools.Country"] end

    PersonalTools:ClearSpChk(true) -- Clear any previous spell checking marks.

    hFile = io.open(sSpChkIgnoredWords, "r")

    repeat
        sText = hFile:read("l")

        if (sText ~= nil) and (sText ~= "") then
            sText = string.gsub(sText," ", "")
            table.insert(asIgnoredWords, sText)
        end

    until (sText == nil)

    io.close(hFile)

    -- ## WordIsIgnored Func ##
    -- Function for checking if a word is on the ignored list.
    function WordIsIgnored(sWord)
    for k, v in pairs(asIgnoredWords) do
        if (v == sWord) then return true end
    end

    return false
    end
    -- ## WordIsIgnored Func ##

    editor:Colourise(0, -1)

    for iLine = 0, editor.LineCount do

        iLineStart = editor:PositionFromLine(iLine)
        iLineEnd = editor.LineEndPosition[iLine]

        iPos = iLineStart

        repeat
            iWordStart = editor:WordStartPosition(iPos)
            iWordEnd = editor:WordEndPosition(iPos)

            if (iWordStart ~= iWordEnd) and (editor:LineFromPosition(iWordEnd) == iLine) then

                iCStyle = editor.StyleAt[iWordStart + 1]
                if (iCStyle == SCE_AU3_COMMENT) or (iCStyle == SCE_AU3_COMMENTBLOCK) or (iCStyle == SCE_AU3_STRING) then

                    editor:SetTargetRange(iWordStart, iWordEnd)
                    if (string.find(editor.TargetText,"[; %(%)%[\\//," .. '"' .. "]") == 1) and
                        (string.find(editor.TargetText,"[%a]") ~= nil) then -- If there is a semicolon, space or beginning round bracket at the beginning of the word, adjust the selection to skip it.
                        iWordStart = iWordStart + 1
                        editor:SetTargetRange(iWordStart, iWordEnd)
                    end

                    if (string.find(editor.TargetText,"[ %.\\//]",string.len(editor.TargetText) - 1) ~= nil) and
                        (string.find(editor.TargetText,"[%a]") ~= nil) then -- If there is a space or period at the end of the word, adjust the selection to skip it.
                        iWordEnd = iWordEnd - 1
                        editor:SetTargetRange(iWordStart, iWordEnd)
                    end

                sText = editor.TargetText

                editor:SetTargetRange(iWordEnd, iWordEnd + 1)
                if (string.find(editor.TargetText, "'")) and editor:IsRangeWord(iWordStart, iWordEnd + 2) then
                    iWordEnd = iWordEnd + 2
                    editor:SetTargetRange(iWordStart, iWordEnd)
                    sText = editor.TargetText
                end

                sText = string.gsub(sText," ", "") -- Remove any spaces in the string.
                if string.find(sText, "[%a]") and (string.find(sText,"[%.$#@_://\\]") == nil) and 
                    (WordIsIgnored(sText) == false) then
                    table.insert(asWordsToCheck,sText .. "\t" .. iWordStart .. "\t" .. iWordEnd)

                end

            end

            iPos = iWordEnd
            end


        iPos = iPos + 1
        until (iPos >= iLineEnd)

    end

    hFile = io.open(sSpChkWordList, "w")

    for k, v in pairs(asWordsToCheck) do
        hFile:write(v .. "\n")

    end
    hFile:flush()
    hFile:close()

    bReturn, iExit, iSignal = os.execute('"' .. sSpChkScript .. '"' .. ' ' .. "##" .. ' ' .. sLang .. ' ' .. sCountry)

    if (iSignal == 1) then

        iOldIndic = editor.IndicatorCurrent
        editor.IndicatorCurrent = iSpChkIndicator
        editor.IndicStyle[iSpChkIndicator] = INDIC_STRAIGHTBOX
        editor.IndicFore[iSpChkIndicator] = 0xFF00FF
        editor.IndicAlpha[iSpChkIndicator] = 100
        editor.IndicOutlineAlpha[iSpChkIndicator] = 100
        hFile = io.open(sSpChkWordList, "r")
        iCount = 0
        editor:BeginUndoAction()
        repeat
            sText = hFile:read("l")

            if (sText ~= nil) and (sText ~= "") then
                iWordStart = string.gsub(string.match(sText, "\t[%d]+\t"), "\t","")
                iFirstTab = string.find(sText, "\t")
                iWordEnd = string.gsub(string.match(sText, "\t[%d]+", iFirstTab + 1), "\t","")
                editor:IndicatorFillRange(iWordStart, iWordEnd - iWordStart)
                iCount = iCount + 1
            end

        until (sText == nil)
        editor:EndUndoAction()
        hFile:close()
        output:AppendText("> Found " .. iCount .. " misspelled word(s).\n")
        editor.IndicatorCurrent = iOldIndic

    elseif (iSignal == 0) then
        output:AppendText("+ No Spelling mistakes found.\n")

    else -- Error of some form.
        output:AppendText("! Scite4AutoIt_LibreOffice_SpellChecker encountered an Error." .. "\n")
        hFile = io.open(sSpChkErrorFile, "r")
        if (hFile == nil) then
            output:AppendText("! Failed to open Error File.\n")
            return
        end
        repeat
            sText = hFile:read("l")

            if (sText ~= nil) and (sText ~= "") then
                output:AppendText("! " .. sText .. "\n")
            end

        until (sText == nil)

        io.close(hFile)
        os.remove(sSpChkErrorFile)

    end
    output:AppendText("++ Scite4AutoIt_LibreOffice_SpellChecker completed. ++\n")
end

--------------------------------------------------------------------------------
-- Clear Spell Checking Marks
-- to be able to use this script you add the following to your SciTEUSer.properties (without the leading "--"):
--#x Clear Spell Checking Marking lua Script
--command.name.47.$(au3)=Sp-Check Clear
--command.mode.47.$(au3)=subsystem:lua,savebefore:no
--command.shortcut.47.$(au3)=Ctrl+Shift+H
--command.47.$(au3)=InvokeTool PersonalTools.ClearSpChk
--
--------------------------------------------------------------------------------
function PersonalTools:ClearSpChk(bInternalCall)
    local  iOldIndic, iSpChkIndicator = nil, 10
    if (bInternalCall ~= true) then output:AppendText("- Scite4AutoIt_LibreOffice_SpellChecker -\n") end
    iOldIndic = editor.IndicatorCurrent
    editor.IndicatorCurrent = iSpChkIndicator
    editor:IndicatorClearRange(0, editor.LineEndPosition[editor.LineCount]) -- Clear any previous spell checking marks.
    output:AppendText("> Spell Check markings successfully cleared.\n")
    editor.IndicatorCurrent = iOldIndic
    if (bInternalCall ~= true) then output:AppendText("++ Scite4AutoIt_LibreOffice_SpellChecker completed. ++\n") end

end