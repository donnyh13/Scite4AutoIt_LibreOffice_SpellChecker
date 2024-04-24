-------------------------------------------------------------------------------
-- required line ... do not remove
PersonalTools = EventClass:new(Common)
-------------------------------------------------------------------------------

--------------------------------------------------------------------------------
-- SingleWordCheck
--
-- Spell-Checks the current word under the caret, activating a Auto-Complete list with any suggested words.
--
-- Single Word Spell-Check
-- to be able to use this script, add the following to your SciTEUser.properties (without the leading "--"):
--#~ Single Word Spell-Check lua Script
--command.name.45.$(au3)=Sp-Check Current Word
--command.mode.45.$(au3)=subsystem:lua,savebefore:no
--command.shortcut.45.$(au3)=Ctrl+Shift+F
--command.45.$(au3)=InvokeTool PersonalTools.SingleWordCheck
--
--------------------------------------------------------------------------------
function PersonalTools.SingleWordCheck()
    -- Settings that can be modified for compatibility with other LUA Scripts.
    local iSpChkIndicator = 10
    local iListType = 18
    local iMarker = 18
    --##########################
    local sSciteUserHome = props["SciteUserHome"]
    local sSpChkScript = sSciteUserHome .. "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_SpellChecker.au3"
    local sSpChkWordList = sSciteUserHome .. "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_SpellChecker.ini"
    local sSpChkErrorFile = sSciteUserHome ..  "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_SpellChecker_ERROR.ini"
    local sCurWord, sLine
    local iSignal,  iOldIndic, iLine, iLineStart, iLineEnd, iPos, iStart, iEnd
    local iCount, iTimer, iMaxSuggestions = 0, os.clock(), 10
    local sLang, sCountry = "en", "US"
    local aWords = {}
    local hFile
    local old_separator
    local sSearchPattern = "[%s%p]?([$@]-[%w_]+['%.,-]?[%w_]+)[%s%p]?"
    local iMarkerMask = (1 << iMarker)

    -- Read the properties.
    if (props["S4A.SpellCheck.Language"] ~= "") then sLang = props["S4A.SpellCheck.Language"] end
    if (props["S4A.SpellCheck.Country"] ~= "") then sCountry = props["S4A.SpellCheck.Country"] end
    if (props["S4A.SpellCheck.MaxSuggestions"] ~= "") then iMaxSuggestions = props["S4A.SpellCheck.MaxSuggestions"] end

    -- Expand the output screen if it is collapsed.
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

    iPos = editor.CurrentPos
    iLine = editor:LineFromPosition(iPos)
    -- Retrieve the Line start and end positions.
    iLineStart = editor:PositionFromLine(iLine)
    iLineEnd = editor.LineEndPosition[iLine]

    if (iLineStart == iLineEnd) then return output:AppendText("! Current line is empty. " .. string.sub(string.format(os.clock() - iTimer), 1, 4) .. " Seconds.\n") end

    if (editor.SelectionStart == editor.SelectionEnd) then

        sLine = editor:GetLine(iLine)
        iStart, iEnd, sCurWord = string.find(sLine, sSearchPattern)

        while (iStart ~= nil) do
            if ((iLineStart + iStart) <= iPos) and ((iLineStart + iEnd) >= iPos) then break end
            iStart, iEnd, sCurWord = string.find(sLine, sSearchPattern, iEnd)
        end

        if (iStart == nil) then return output:AppendText("! Failed to identify valid word at cursor. " .. string.sub(string.format(os.clock() - iTimer), 1, 4) .. " Seconds.\n") end

    else -- Selected word check.
        iStart = editor.SelectionStart
        iEnd = editor.SelectionEnd
        editor:SetTargetRange(iStart, iEnd)

        iStart = iStart - iLineStart
        iEnd = iEnd - iLineStart
        -- Retrieve the selected word.
        sCurWord = string.gsub(editor.TargetText, " ", "")
    end

    -- Clear any SpellChecking markings for the current word.
    iOldIndic = editor.IndicatorCurrent
    editor.IndicatorCurrent = iSpChkIndicator
    editor:IndicatorClearRange((iLineStart + iStart), iEnd - iStart)
    editor.IndicatorCurrent = iOldIndic

    -- Delete the marker if there are no spelling errors remaining on the line.
    if (editor:MarkerGet(iLine) & iMarkerMask ~= iMarkerMask) and
        (editor:LineFromPosition(editor:IndicatorEnd(iSpChkIndicator, iLineStart)) ~= iLine) then editor:MarkerDelete(iLine, iMarker) end
    -- if (editor:LineFromPosition(editor:IndicatorEnd(iSpChkIndicator, iLineStart)) ~= iLine) then editor:MarkerDelete(iLine, iMarker) end

    -- Execute the Spell Checking Script.
    _, _, iSignal = os.execute('"' .. sSpChkScript .. '"' .. " " .. sCurWord .. " " .. sLang .. " " .. sCountry .. " " .. "true" .. " " .. iMaxSuggestions)

    -- iSignal will be either, 0 = Word is spelled correctly, 1 = word it misspelled, or 2 = An error occurred executing Spell Check Script.
    if (iSignal == 0) then
       output:AppendText("+ The word " .. '"' .. sCurWord .. '"' .. " is spelled correctly." .. "\n")

    elseif (iSignal == 1) then

        -- Open and read the Spell Check Word List that will contain a list of suggested words.
        hFile = io.open(sSpChkWordList, "r")

        repeat
            sLine = hFile:read("l")

            if (sLine ~= nil) and (sLine ~= "") then
                table.insert(aWords,sLine)
                iCount = iCount + 1
            end

        until (sLine == nil)

        io.close(hFile)

        if (iCount > 0) then
            old_separator = editor.AutoCSeparator
            editor.AutoCSeparator = string.byte(';')
            editor:UserListShow(iListType, table.concat(aWords, ';'))
            editor.AutoCSeparator = old_separator
            output:AppendText("> The word " .. '"' .. sCurWord .. '"' .. " is misspelled, found " .. iCount .. " suggestion(s)." .. "\n")

        else
            output:AppendText("! The word " .. '"' .. sCurWord .. '"' .. " is misspelled, but found no spelling suggestions." .. "\n")

       end

    else -- Error of some form.
        output:AppendText("! Scite4AutoIt_LibreOffice_SpellChecker encountered an Error." .. "\n")

        -- Open the Error file and read the errors the Au3 Script encountered, and output them.
        hFile = io.open(sSpChkErrorFile, "r")
        if (hFile == nil) then
            output:AppendText("! Failed to open Error File." .. "\n")
            return
        end
        repeat
            sLine = hFile:read("l")

            if (sLine ~= nil) and (sLine ~= "") then
                output:AppendText("! " .. sLine .. "\n")
            end

        until (sLine == nil)

        io.close(hFile)

        -- Delete the error log file.
        os.remove(sSpChkErrorFile)
    end

    output:AppendText("++ Scite4AutoIt_LibreOffice_SpellChecker completed. " .. string.sub(string.format(os.clock() - iTimer), 1, 4) .. " Seconds. ++\n")
end

--------------------------------------------------------------------------------
-- CheckScript
--
-- Spell-Checks the entire Script, marking any misspelled words.
--
-- to be able to use this script you add the following to your SciTEUser.properties (without the leading "--"):
--#~ Entire Script Spell-Check lua Script
--command.name.46.$(au3)=Sp-Check Entire Script
--command.mode.46.$(au3)=subsystem:lua,savebefore:no
--command.shortcut.46.$(au3)=Ctrl+Shift+G
--command.46.$(au3)=InvokeTool PersonalTools.CheckScript
--
--------------------------------------------------------------------------------
function PersonalTools:CheckScript()
    -- Settings that can be modified for compatibility with other LUA Scripts.
    local iSpChkIndicator = 10
    local iMarker = 18
    --##########################
    local sSciteUserHome = props["SciteUserHome"]
    local sSpChkScript = sSciteUserHome .. "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_SpellChecker.au3"
    local sSpChkWordList = sSciteUserHome .. "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_SpellChecker.ini"
    local sSpChkIgnoredWords = sSciteUserHome .. "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_IgnoredWords.ini"
    local sSpChkErrorFile = sSciteUserHome ..  "\\Scite4AutoIt_LO_SpellChecker\\Scite4AutoIt_LO_SpellChecker_ERROR.ini"
    local iLineStart, iLineEnd, iWordStart, iWordEnd, iCStyle, iOldIndic, iSignal, iStart, iEnd, iFirstLine, iLastLine
    local iCount, iMaxSuggestions, iTimer  = 0, 10, os.clock()
    local sLang, sCountry, sHighlight = "en", "US", "0xFF00FF"
    local sLine, sText
    local asIgnoredWords = {}
    local hFile
    local sSearchPattern = "[%s%p]?([$@]-[%w_]+['%.,-]?[%w_]+)[%s%p]?"
    local iMarkerMask = (1 << iMarker)
    local bUseMarkers = false

    -- Expand the output screen if it is collapsed.
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

    -- Read the properties.
    if (props["S4A.SpellCheck.Language"] ~= "") then sLang = string.gsub(props["S4A.SpellCheck.Language"], " ", "") end -- Strip spaces
    if (props["S4A.SpellCheck.Country"] ~= "") then sCountry = string.gsub(props["S4A.SpellCheck.Country"], " ", "") end -- Strip spaces
    if (props["S4A.SpellCheck.MaxSuggestions"] ~= "") then iMaxSuggestions = props["S4A.SpellCheck.MaxSuggestions"] end
    if (props["S4A.SpellCheck.UseMarkers"] ~= "n") then bUseMarkers = true end
    if (props["S4A.SpellCheck.Highlight"] ~= "") then
        sHighlight = string.gsub(props["S4A.SpellCheck.Highlight"]," ", "")
        if (string.len(sHighlight) ~= 8) or (string.find(sHighlight,"[^a-fxA-F0-9]") ~= nil) or (string.find(sHighlight,"0x") ~= 1) then
            if (string.len(sHighlight) ~= 8) then output:AppendText("! User designated highlight color is the wrong length, (" .. string.len(sHighlight) .. " characters instead of 8 characters), reverting to original color.\n") end
            if (string.find(sHighlight,"[^a-fxA-F0-9]") ~= nil) then output:AppendText("! User designated highlight color contains invalid character(s)-->" .. string.gsub(sHighlight, "[a-fA-Fx0-9]", "") .. "<--reverting to original color.\n") end
            if (string.find(sHighlight,"0x") ~= 1) then output:AppendText("! User designated highlight color does not begin with 0x, reverting to original color.\n") end

            sHighlight = "0xFF00FF"
        end
    end

    -- Clear any previous spell checking marks.
    self:ClearSpChk(true)

    -- Open the Ignored Words file.
    hFile = io.open(sSpChkIgnoredWords, "r")

    repeat
        sLine = hFile:read("l")

        if (sLine ~= nil) and (sLine ~= "") then
            sLine = string.gsub(sLine," ", "")
            table.insert(asIgnoredWords, sLine)
        end

    until (sLine == nil)

    io.close(hFile)

    -- ## WordIsIgnored Func ##
    -- Function for checking if a word is on the ignored list.
    function WordIsIgnored(sWord)
        for k = 1, #asIgnoredWords do
            if (asIgnoredWords[k] == sWord) then return true end
        end
        return false
        end
    -- ## WordIsIgnored Func ##

    -- Lexer the entire script so all is styled correctly so this script can determine where comments, strings etc are.
    editor:Colourise(0, -1)

    if (editor.SelectionStart == editor.SelectionEnd) then
        iFirstLine = 0
        iLastLine = editor.LineCount

    else
        iFirstLine = editor:LineFromPosition(editor.SelectionStart)
        iLastLine = editor:LineFromPosition(editor.SelectionEnd)
    end

    -- Open and write the words to check to my Check Words file.
    hFile = io.open(sSpChkWordList, "w")

    -- Cycle through the Script lines beginning at the beginning.
    for iLine = iFirstLine, iLastLine do

        -- Retrieve the Line start and end positions.
        iLineStart = editor:PositionFromLine(iLine)
        iLineEnd = editor.LineEndPosition[iLine]

        if (iLineStart ~= iLineEnd) then

            sLine = editor:GetLine(iLine)
            iStart, iEnd, sText = string.find(sLine, sSearchPattern)

            while (iStart ~= nil) do
                iCStyle = editor.StyleAt[iLineStart + iStart + 1]
                if (iCStyle == SCE_AU3_COMMENT) or (iCStyle == SCE_AU3_COMMENTBLOCK) or (iCStyle == SCE_AU3_STRING) then
                    -- If the String contains letters, but not a period, $, #, @, _, :, / or \ then check the word if it is spelled correctly.
                    if string.find(sText, "[%a]") and (string.find(sText,"[%.$#@_://\\]") == nil) and
                            (WordIsIgnored(sText) == false) then-- check if word is ignored.
                        -- insert the word into my file with its start and stop positions, separated by tabs.
                        hFile:write(sText .. "\t" .. (iLineStart + iStart) .. "\t" .. (iLineStart + iEnd - 1) .. "\n")
                    end
                end
                iStart, iEnd, sText = string.find(sLine, sSearchPattern, iEnd)
            end
        end
    end

    hFile:flush()
    hFile:close()

    -- Run my Spell Check script.
    _, _, iSignal = os.execute('"' .. sSpChkScript .. '" ' .. "## " .. sLang .. " " .. sCountry .. " " .. "false " .. iMaxSuggestions)

    -- iSignal will be either, 0 = Words are spelled correctly, 1 = words are misspelled, or 2 = An error occurred executing Spell Check Script.
    if (iSignal == 0) then
        output:AppendText("+ No Spelling mistakes found.\n")

    elseif (iSignal == 1) then

        -- Set up the Spell Checking indicator.
        iOldIndic = editor.IndicatorCurrent
        editor.IndicatorCurrent = iSpChkIndicator
        editor.IndicStyle[iSpChkIndicator] = INDIC_STRAIGHTBOX
        editor.IndicFore[iSpChkIndicator] = sHighlight
        editor.IndicAlpha[iSpChkIndicator] = 100
        editor.IndicOutlineAlpha[iSpChkIndicator] = 100

        -- Open and read the Word list of misspelled words.
        hFile = io.open(sSpChkWordList, "r")

        if bUseMarkers then 
            editor:MarkerDefine(iMarker, 32) -- 32 = downward flag.
            editor.MarkerFore[iMarker] = "0x000000"
            editor.MarkerBack[iMarker] = "0x0000FF"
        end

        editor:BeginUndoAction()
        repeat
            sLine = hFile:read("l")

            if (sLine ~= nil) and (sLine ~= "") then
                -- Identify the misspelled words start/end Position. They will be located after the first and second tab.
                _, _, iWordStart, iWordEnd = string.find(sLine, "\t([%d]+)\t([%d]+)")
                -- Mark the word.
                editor:IndicatorFillRange(iWordStart, iWordEnd - iWordStart)
                if bUseMarkers and ((editor:MarkerGet(editor:LineFromPosition(iWordStart)) & iMarkerMask) ~= iMarkerMask) then -- If Marker not already present, add one.
                    editor:MarkerAdd(editor:LineFromPosition(iWordStart), iMarker)
                end
                -- Unfold any folds the word is in.
                editor:EnsureVisible(editor:LineFromPosition(iWordStart))
                iCount = iCount + 1

            end
        until (sLine == nil)
        editor:EndUndoAction()
        hFile:close()
        output:AppendText("> Found " .. iCount .. " misspelled word(s).\n")
        editor.IndicatorCurrent = iOldIndic

    else -- Error of some form.
        output:AppendText("! Scite4AutoIt_LibreOffice_SpellChecker encountered an Error." .. "\n")

        -- Open the Error file and read the errors the Au3 Script encountered, and output them.
        hFile = io.open(sSpChkErrorFile, "r")
        if (hFile == nil) then
            output:AppendText("! Failed to open Error File.\n")
            return
        end

        repeat
            sLine = hFile:read("l")

            if (sLine ~= nil) and (sLine ~= "") then
                output:AppendText("! " .. sLine .. "\n")
            end

        until (sLine == nil)

        io.close(hFile)

        -- Delete the error log.
        os.remove(sSpChkErrorFile)

    end

    output:AppendText("++ Scite4AutoIt_LibreOffice_SpellChecker completed. " .. string.sub(string.format(os.clock() - iTimer), 1, 4) .. " Seconds. ++\n")
end

--------------------------------------------------------------------------------
-- ClearSpChk
--
-- Clears all Spell Checking Marks in the current script.
--
-- Parameters:
--	bInternalCall - If True, the call is from another function, suppress any introductory messages.
--
-- to be able to use this script you add the following to your SciTEUser.properties (without the leading "--"):
--#x Clear Spell Checking Marking lua Script
--command.name.47.$(au3)=Sp-Check Clear
--command.mode.47.$(au3)=subsystem:lua,savebefore:no
--command.shortcut.47.$(au3)=Ctrl+Shift+H
--command.47.$(au3)=InvokeTool PersonalTools.ClearSpChk
--
--------------------------------------------------------------------------------
function PersonalTools.ClearSpChk(bInternalCall)
    -- Settings that can be modified for compatibility with other LUA Scripts.
    local iSpChkIndicator = 10
    local iMarker = 18
    --##########################
    local iOldIndic, iStartSel, iEndSel, iFirstLine, iLastLine, iFoundLine, iLineStart
    local iTimer = os.clock()
    local iMarkerMask = (1 << iMarker)
    local sMsg = " for range"
    local bClearAllMarkers = false

    -- Expand the output screen if it is collapsed.
    if output.LinesOnScreen == 0 then scite.MenuCommand(IDM_TOGGLEOUTPUT) end
    if (bInternalCall ~= true) then output:AppendText("- Scite4AutoIt_LibreOffice_SpellChecker -\n") end

    iStartSel = editor.SelectionStart
    iEndSel = editor.SelectionEnd

    if (iStartSel == iEndSel) then
        iStartSel = 0
        iEndSel = editor.TextLength
        sMsg = ""
        bClearAllMarkers = true
    end

    iOldIndic = editor.IndicatorCurrent
    editor.IndicatorCurrent = iSpChkIndicator
    editor:IndicatorClearRange(iStartSel, iEndSel - iStartSel) -- Clear any previous spell checking marks for the desired range.

    if bClearAllMarkers then
        editor:MarkerDeleteAll(iMarker)

    else -- Remove markers for a range.
        iFirstLine = editor:LineFromPosition(iStartSel)
        iLastLine = editor:LineFromPosition(iEndSel)

        iFoundLine = editor:MarkerNext(iFirstLine, iMarkerMask)

        while (iFoundLine ~= -1) do
            if ((iFoundLine < iFirstLine) or (iFoundLine > iLastLine)) then break end

            if (iFoundLine == iFirstLine) or (iFoundLine == iLastLine) then

                iLineStart = editor:PositionFromLine(iFoundLine)
                if (editor:LineFromPosition(editor:IndicatorEnd(iSpChkIndicator, iLineStart)) ~= iFoundLine) then -- Delete the marker if there are no spelling errors remaining on the line.
                    editor:MarkerDelete(iFoundLine, iMarker)
                else
                    iFoundLine = iFoundLine + 1

                end

            else
                editor:MarkerDelete(iFoundLine, iMarker)

            end

            iFoundLine = editor:MarkerNext(iFoundLine, iMarkerMask)
        end

    end

    output:AppendText("> Spell Check markings successfully cleared" .. sMsg .. ".\n")
    editor.IndicatorCurrent = iOldIndic

    if (bInternalCall ~= true) then output:AppendText("++ Scite4AutoIt_LibreOffice_SpellChecker completed. " .. string.sub(string.format(os.clock() - iTimer), 1, 4) .. " Seconds. ++\n") end
end

-------------------------------------------------------------------------------
-- OnUserListSelection
--
-- Replaces the current word with the Selection made in a User List.
--
-- Parameters:
--	iListType - The User List Type.
--	sSel - The word selected by the user from the User List.
--------------------------------------------------------------------------------
function PersonalTools.OnUserListSelection(iListType, sSel)
    -- Settings that can be modified for compatibility with other LUA Scripts.
    local iMyListType = 18
    --##########################
    -- The List Style I use for Spelling suggestions is 18, if the list is mine, perform the word replacement.
    if iListType == iMyListType then
        local iLine, iStart, iEnd, iLineStart
        local sLine
        local sSearchPattern = "[%s%p]?([$@]-[%w_]+['%.,-]?[%w_]+)[%s%p]?"
        local iPos = editor.CurrentPos

        if (editor.SelectionStart == editor.SelectionEnd) then -- No Selection.
            iLine = editor:LineFromPosition(iPos)
            -- Retrieve the Line start and end positions.
            iLineStart = editor:PositionFromLine(iLine)

            sLine = editor:GetLine(iLine)
            iStart, iEnd = string.find(sLine, sSearchPattern)

            if (iStart == nil) then return output:AppendText("! Failed to identify valid word to replace.\n") end

            while (iStart ~= nil) do
                if ((iLineStart + iStart) <= iPos) and ((iLineStart + iEnd) >= iPos) then break end

                iStart, iEnd = string.find(sLine, sSearchPattern, iEnd)
            end

            editor:SetTargetRange((iLineStart + iStart), (iLineStart + iEnd - 1))

        else -- Selected word replace.
            editor:SetTargetRange(editor.SelectionStart, editor.SelectionEnd)

        end

        editor:BeginUndoAction()

        -- Replace the word with the new chosen word.
        editor:ReplaceTarget(sSel)
        -- Go to the end of the new word.
        editor:GotoPos(editor.CurrentPos + string.len(sSel))
        editor:EndUndoAction()

        -- Lexer the new word.
        editor:Colourise(editor.CurrentPos - string.len(sSel), editor.CurrentPos)
    end
end
