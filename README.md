# Scite4AutoIt_LibreOffice_SpellChecker

## Description
 Integrates LibreOffice's Spell Checker ability into Scite4AutoIt. Allowing for single word Spell Checking, or whole script misspell marking.

## Requirements
This SpellChecker requires the following to function:
- Full installation of Scite4AutoIt
- LibreOffice (**Installed** version only, portable **WILL NOT** work.)
- Desired Spell Checking language installed in Libre Office.

## Installation

To "Install" this AddOn for Scite4AutoIt, do the following.
- Navigate to %LocalAppData%\Local\AutoIt v3\SciTE
  (You can copy and paste the above into Start search box and push enter)
- Copy the "Scite4AutoIt_LO_SpellChecker" folder into this directory.
- Do one of the following, depending on if you already have custom lua scripts present in PersonalTools.lua or not.
  - If you do not have custom scripts present in PersonalTools.lua, 
    - find in this directory the file PersonalTools.lua, and rename it to PersonalTools.lua.old 
    - Copy the PersonalTools.lua file from the Add-On folder to the %LocalAppData%\Local\AutoIt v3\SciTE directory.
  - If you do have custom scripts present in PersonalTools.lua,
    - Open the PersonalTools.lua file in the Add-On folder and copy everything in the file below the following entry,
      > -- required line ... do not remove  
      > PersonalTools = EventClass:new(Common)
    - Paste the contents into the end of the PersonalTools.lua file found in the %LocalAppData%\Local\AutoIt v3\SciTE directory.
- Open Scite editor, Select "Options" in the toolbar, Select "Open User Options File".
- Copy and paste the following into the "User Options" File. (If you have previous custom Lua Scripts, you may need to modify the command numbering.)
#~ Single Word Spell-Check lua Script
command.name.45.$(au3)=Sp-Check Current Word
command.mode.45.$(au3)=subsystem:lua,savebefore:no
command.shortcut.45.$(au3)=Ctrl+Shift+F
command.45.$(au3)=InvokeTool PersonalTools.SingleWordCheck
#~

#~ Entire Script Spell-Check lua Script
command.name.46.$(au3)=Sp-Check Entire Script
command.mode.46.$(au3)=subsystem:lua,savebefore:no
command.shortcut.46.$(au3)=Ctrl+Shift+G
command.46.$(au3)=InvokeTool PersonalTools.CheckScript
#~

#~ Clear Spell Checking Marking lua Script
command.name.47.$(au3)=Sp-Check Clear
command.mode.47.$(au3)=subsystem:lua,savebefore:no
command.shortcut.47.$(au3)=Ctrl+Shift+H
command.47.$(au3)=InvokeTool PersonalTools.ClearSpChk

  *You may modify any Shortcut key combinations as necessary.*
- Restart Scite

- You may set the following settings by pasting the below properties in your "User Options" File, and modifying the value after "="

personal.tools.Language=en
personal.tools.Country=US

  - Language must be a LOWER CASE two or three character ISO 639 Language Code. You can find a list of these codes online, such as https://iso639-3.sil.org/code_tables/639/data. 
  - Country must be a UPPER CASE two character ISO 3166 Country Code. You can find a list of these codes online, such as https://en.wikipedia.org/wiki/ISO_3166-1_alpha-2
  - **YOU MUST HAVE THE desired Language installed in Libre Office already for this to work**.

- You can add words to ignore in the file "Scite4AutoIt_LO_IgnoredWords.ini" that you placed in "%LocalAppData%\Local\AutoIt v3\SciTE\Scite4AutoIt_LO_SpellChecker".
  - ONLY one word per line.
  - No puctuation characters please.
  
## Release

## Changes

Please see the [Changelog](CHANGELOG.md)

## License

Distributed under the MIT License. See the [LICENSE](LICENSE) for more information.

## Acknowledgements

- Opportunity by [GitHub](https://github.com)
- Scripting ability by [AutoIt](https://www.autoitscript.com/site/autoit/)
- Thanks to the authors of the Third-Party UDFs:
  - *OpenOffice/LibreOffice Spell Checker* by @GMK, @mLipok. [OpenOffice/LibreOffice Spell Checker](https://www.autoitscript.com/forum/topic/185932-openofficelibreoffice-spell-checker/)

## Links 

[License](LICENSE) <br>
[AutoIt](https://www.autoitscript.com/site/autoit/) <br>
