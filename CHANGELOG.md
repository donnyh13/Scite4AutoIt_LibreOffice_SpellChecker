# Scite4AutoIt_LibreOffice_SpellChecker Changelog

All notable changes to "Scite4AutoIt_LO_SpellChecker" will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
This project also adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Go to [legend](#legend---types-of-changes) for further information about the types of changes.

## Releases

|    Version           |    Changes                              |    Download                    |     Released   |    Compare on GitHub       |
|:---------------------|:---------------------------------------:|:------------------------------:|:--------------:|:---------------------------|
|    **v0.1.0**        | [Change Log](#010---2024-04-)            | [v0.1.0][v0.1.0]               | 2024-04-??     | [Compare][v0.1.0-Compare]  |
|    **v0.0.1-alpha**  | [Change Log](#001-alpha---2023-12-20)   | [v0.0.1-alpha][v0.0.1-alpha]   | 2023-12-20     |                            |

## [0.1.0] - 2024-04-??

### Added

- Spell Checker now can mark lines with misspelled words using a Margin marker also.
- Single Word check now will check the currently selected word.
- Script check now will be restrained to a selected portion if a selection exists.
- ClearSpChk function will now clear only a selection if one exists.
- ClearSpChk function now also clears the new Margin Markers also.
- Selected words that are single word spell checked can also be replaced using the drop down.

## Fixed

- Some apostrophied words not being correctly selected, such as "I'll".
- Some markings incorrectly included as part of the word at the beginning or end, such as dashes, quotations etc.
- Two if blocks in the program were incorrectly testing for nill, instead of nil.

## Changed

- Rewrote the word recognition method to work better and quicker.
- Seconds read out at end of Script is now maximum 4 characters long.
- Words to spell check are now directly written to the file, instead of to a table and then to a file, saving several milliseconds.

## [0.0.1-alpha] - 2023-12-20

### Fixed

- Spell Checking sometimes not scanning whole document.

### Changed

- Single Word Check to eliminate periods at the end of the word if present.
- Various punctuation characters are skipped at the beginning of words.
- Includes apostrophied words.
- Changed "personal.tools.Language" to "S4A.SpellCheck.Language"
- Changed "personal.tools.Country" to  "S4A.SpellCheck.Country"
- Scite4AutoIt_LO_SpellChecker.au3 to use all functions without using global variables.

### Added

- Initial Release.
- S4A.SpellCheck.Highlight property.
- S4A.SpellCheck.MaxSuggestions property.
- Multi-Language spell checking support.
- Details to ReadMe.
- Comments in Script.
- Misspelled words found in folds are now unfolded.

---

### Legend - Types of changes

- `Added` for new features.
- `Changed` for changes in existing functionality.
- `Deprecated` for soon-to-be removed features.
- `Fixed` for any bug fixes.
- `Removed` for now removed features.
- `Security` in case of vulnerabilities.
- `Project` for documentation or overall project improvements.

##

[To the top](#scite4autoitlibreofficespellchecker-changelog)

---

[v0.1.0-Compare]:	https://github.com/donnyh13/Au3LibreOffice/compare/v0.0.1-alpha...v0.1.0

[v0.1.0]:	https://github.com/donnyh13/Au3LibreOffice/releases/tag/v0.1.0
[v0.0.1-alpha]:	https://github.com/donnyh13/Scite4AutoIt_LibreOffice_SpellChecker/releases/tag/0.0.1-alpha
