@echo off

REM - Set Console Title
title Grep To Table

REM - Set Console Color
REM (background)(foreground)
REM | 0 = Black  | 1 = Blue  | 2 = Green
REM | 3 = Aqua   | 4 = Red   | 5 = Purple
REM | 6 = Yellow | 7 = White | 8 = Gray
Color 02

REM --- RMDIR ---
REM Removes (deletes) a directory.
REM RMDIR [/S] [/Q] [drive:]path
REM RD [/S] [/Q] [drive:]path
REM     /S      Removes all directories and files in the specified directory
REM             in addition to the directory itself.  Used to remove a directory
REM             tree.
REM     /Q      Quiet mode, do not ask if ok to remove a directory tree with /S
REM
REM --- MKDIR ---
REM Creates a directory.
REM MKDIR [drive:]path
REM MD [drive:]path
REM If Command Extensions are enabled MKDIR changes as follows:
REM MKDIR creates any intermediate directories in the path, if needed.
REM For example, assume \a does not exist then:
REM     mkdir \a\b\c\d
REM is the same as:
REM     mkdir \a
REM     chdir \a
REM     mkdir b
REM     chdir b
REM     mkdir c
REM     chdir c
REM     mkdir d
REM which is what you would have to type if extensions were disabled.

SETLOCAL
set "TempFolder=.\temp"

CALL :CleanTempFolder %TempFolder%

echo ---Get Settings
set /p "FpSetting=File path of the setting: "

echo ---Print Options
cscript //nologo .\src\PrintOptions.vbs %FpSetting%

echo ---Pattern Index Selection
set /p "PatternIndex=Pattern Index: "
echo.

echo ---Get Input Folder
set /p "DirIn=Folder with data files: "
echo DirOut=%DirIn%
echo.

echo ---Get output Folder
set /p "DirOut=Folder where output file is stored: "
echo DirOut=%DirOut%
echo.

echo ---Run GrepToTable.vbs
cscript //nologo .\src\GrepToTable.vbs %DirIn% %DirOut% %PatternIndex% .\src\Settings.xml .\temp

CALL :CleanTempFolder  %TempFolder%

EXIT /B %ERRORLEVEL%

REM Delete the temp folder if it exists
REM Create a new folder.
:CleanTempFolder
echo.
echo ---Press Enter to clean up temp folder.
pause
if exist %~1 (
RMDIR /S /Q %~1
)
MKDIR %~1
EXIT /B 0
