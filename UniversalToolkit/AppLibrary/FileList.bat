@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: ============================================================
::  FILE / FOLDER NAME LISTER
::  - Browse dialogs for source and output
::  - OPTIONAL recursive subfolder search
::  - Files grouped by EXTENSION (.EXE, .TXT, .PDF, etc.)
::  - DEFAULT save location = DESKTOP
:: ============================================================

title Directory File & Folder Name Lister

cls
echo ============================================================
echo      DIRECTORY FILE ^& FOLDER NAME LISTER
echo ============================================================
echo.

:: ============================================================
::  Get Desktop path reliably
:: ============================================================
set "DESKTOP_DIR="
for /f "delims=" %%P in ('
    powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"
') do set "DESKTOP_DIR=%%P"

if "!DESKTOP_DIR!"=="" set "DESKTOP_DIR=%USERPROFILE%\Desktop"

:: ============================================================
::  STEP 1: Browse for SOURCE directory
:: ============================================================
echo Step 1 of 4: Select the SOURCE folder to scan...
echo.

call :BROWSE_FOLDER "Select the folder to scan for files and folders" "SOURCE_DIR"

if "!SOURCE_DIR!"=="" (
    echo ERROR: No folder was selected.
    pause
    exit /b 1
)

echo Source folder: !SOURCE_DIR!
echo.

:: ============================================================
::  STEP 2: Subfolder search option
:: ============================================================
echo Step 2 of 4: Include subfolders?
echo   [1] Top-level only ^(default^) - scan just this folder
echo   [2] Recursive - search ALL subfolders too
echo.

:SUBFOLDER_CHOICE
set "SUB_CHOICE="
set /p "SUB_CHOICE=Enter your choice [1/2] (press Enter = top-level only): "

if "%SUB_CHOICE%"=="" (
    set "RECURSIVE=0"
    goto SUB_SELECTED
)
if "%SUB_CHOICE%"=="1" (
    set "RECURSIVE=0"
    goto SUB_SELECTED
)
if "%SUB_CHOICE%"=="2" (
    set "RECURSIVE=1"
    goto SUB_SELECTED
)

echo Invalid choice. Please enter 1 or 2.
echo.
goto SUBFOLDER_CHOICE

:SUB_SELECTED
if %RECURSIVE%==1 (
    echo Mode: Recursive ^(includes all subfolders^)
) else (
    echo Mode: Top-level only
)
echo.

:: ============================================================
::  STEP 3: Choose OUTPUT directory (DEFAULT = DESKTOP)
:: ============================================================
echo Step 3 of 4: Choose where to save the output file:
echo   [1] Save to DESKTOP ^(default^)  -  !DESKTOP_DIR!
echo   [2] Save in the SOURCE folder   -  !SOURCE_DIR!
echo   [3] Browse for a different folder
echo.

:OUTPUT_CHOICE
set "OUT_CHOICE="
set /p "OUT_CHOICE=Enter your choice [1/2/3] (press Enter = Desktop): "

if "%OUT_CHOICE%"=="" (
    set "OUTPUT_DIR=!DESKTOP_DIR!"
    goto OUTPUT_SELECTED
)
if "%OUT_CHOICE%"=="1" (
    set "OUTPUT_DIR=!DESKTOP_DIR!"
    goto OUTPUT_SELECTED
)
if "%OUT_CHOICE%"=="2" (
    set "OUTPUT_DIR=!SOURCE_DIR!"
    goto OUTPUT_SELECTED
)
if "%OUT_CHOICE%"=="3" (
    echo.
    call :BROWSE_FOLDER "Select where to save the output file" "OUTPUT_DIR"
    if "!OUTPUT_DIR!"=="" (
        echo ERROR: No output folder was selected.
        pause
        exit /b 1
    )
    goto OUTPUT_SELECTED
)

echo Invalid choice. Please enter 1, 2, or 3.
echo.
goto OUTPUT_CHOICE

:OUTPUT_SELECTED
echo Output folder: !OUTPUT_DIR!
echo.

:: ============================================================
::  STEP 4: Choose output format
:: ============================================================
echo Step 4 of 4: Choose output format:
echo   [1] Plain Text (.TXT)
echo   [2] Rich Text Format (.RTF)
echo.

:CHOOSE_FORMAT
set "FORMAT_CHOICE="
set /p "FORMAT_CHOICE=Enter your choice (1 or 2): "

if "%FORMAT_CHOICE%"=="1" (
    set "OUTPUT_EXT=txt"
    goto FORMAT_SELECTED
)
if "%FORMAT_CHOICE%"=="2" (
    set "OUTPUT_EXT=rtf"
    goto FORMAT_SELECTED
)

echo Invalid choice. Please enter 1 or 2.
echo.
goto CHOOSE_FORMAT

:FORMAT_SELECTED

:: ============================================================
::  Set output filename with timestamp
:: ============================================================
set "TIMESTAMP="
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set "TIMESTAMP=%%I"
set "TIMESTAMP=%TIMESTAMP:~0,8%_%TIMESTAMP:~8,6%"

set "OUTPUT_FILE=!OUTPUT_DIR!\FileList_!TIMESTAMP!.!OUTPUT_EXT!"

:: ============================================================
::  COLLECT DATA
:: ============================================================
set "FOLDER_LIST=%temp%\dir_folders_temp.txt"
set "EXT_LIST=%temp%\dir_exts_temp.txt"
set "EXT_SORTED=%temp%\dir_exts_sorted.txt"

:: Clean any previous temp files
if exist "!FOLDER_LIST!" del "!FOLDER_LIST!"
if exist "!EXT_LIST!" del "!EXT_LIST!"
if exist "!EXT_SORTED!" del "!EXT_SORTED!"
del "%temp%\dir_ext_*.tmp" 2>nul

set /a FOLDER_COUNT=0
set /a FILE_COUNT=0

if %RECURSIVE%==1 (
    call :COLLECT_RECURSIVE
) else (
    call :COLLECT_TOPLEVEL
)

:: Sort extensions alphabetically
if exist "!EXT_LIST!" (
    sort "!EXT_LIST!" /o "!EXT_SORTED!"
) else (
    type nul > "!EXT_SORTED!"
)

:: Count files per extension
for /f "usebackq delims=" %%E in ("!EXT_SORTED!") do (
    set /a EXT_COUNT_%%E=0
    if exist "%temp%\dir_ext_%%E.tmp" (
        for /f %%C in ('find /c /v "" ^< "%temp%\dir_ext_%%E.tmp"') do set "EXT_COUNT_%%E=%%C"
    )
)

:: ============================================================
::  Write output based on format
:: ============================================================
if "!OUTPUT_EXT!"=="txt" (
    call :WRITE_TXT
) else (
    call :WRITE_RTF
)

:: ============================================================
::  Cleanup all temp files
:: ============================================================
del "!FOLDER_LIST!" 2>nul
del "!EXT_LIST!" 2>nul
del "!EXT_SORTED!" 2>nul
del "%temp%\dir_ext_*.tmp" 2>nul

:: ============================================================
::  Done
:: ============================================================
cls
echo.
echo ============================================================
echo  SUCCESS!
echo ============================================================
echo  Scanned:     !SOURCE_DIR!
if %RECURSIVE%==1 (echo  Mode:        Recursive ^(all subfolders^)
) else (echo  Mode:        Top-level only)
echo  Folders:     !FOLDER_COUNT!
echo  Files:       !FILE_COUNT!
echo  Saved to:    !OUTPUT_FILE!
echo ============================================================
echo.
pause

:: Open the file automatically
start "" "!OUTPUT_FILE!"

endlocal
exit /b 0

:: ============================================================
::  SUBROUTINE: COLLECT_TOPLEVEL
:: ============================================================
:COLLECT_TOPLEVEL
:: Collect folders
for /d %%D in ("!SOURCE_DIR!\*") do (
    echo %%~nxD>> "!FOLDER_LIST!"
    set /a FOLDER_COUNT+=1
)

:: Collect files and group by extension
for %%F in ("!SOURCE_DIR!\*") do (
    set /a FILE_COUNT+=1
    set "FNAME=%%~nxF"
    set "FEXT=%%~xF"
    if "!FEXT!"=="" set "FEXT=(NONE)"
    echo !FNAME!>> "%temp%\dir_ext_!FEXT!.tmp"
    findstr /x /c:"!FEXT!" "!EXT_LIST!" 2>nul || echo !FEXT!>> "!EXT_LIST!"
)
goto :eof

:: ============================================================
::  SUBROUTINE: COLLECT_RECURSIVE
:: ============================================================
:COLLECT_RECURSIVE
:: Get source dir length for relative path calculation
set "SRC_LEN=0"
call :STRLEN "!SOURCE_DIR!" SRC_LEN
set /a SRC_LEN+=1  :: +1 for the trailing backslash

:: Collect all folders recursively (relative paths)
for /f "delims=" %%D in ('dir "!SOURCE_DIR!" /s /b /ad 2^>nul') do (
    set "FULLPATH=%%D"
    call :GET_RELPATH
    echo !RELPATH!>> "!FOLDER_LIST!"
    set /a FOLDER_COUNT+=1
)

:: Collect all files recursively (relative paths), group by extension
for /f "delims=" %%F in ('dir "!SOURCE_DIR!" /s /b /a-d 2^>nul') do (
    set /a FILE_COUNT+=1
    set "FULLPATH=%%F"
    set "FNAME=%%~nxF"
    set "FEXT=%%~xF"
    if "!FEXT!"=="" set "FEXT=(NONE)"
    call :GET_RELPATH
    echo !RELPATH!>> "%temp%\dir_ext_!FEXT!.tmp"
    findstr /x /c:"!FEXT!" "!EXT_LIST!" 2>nul || echo !FEXT!>> "!EXT_LIST!"
)
goto :eof

:: ============================================================
::  SUBROUTINE: GET_RELPATH - sets RELPATH from FULLPATH
:: ============================================================
:GET_RELPATH
set "RELPATH=!FULLPATH:~%SRC_LEN%!"
goto :eof

:: ============================================================
::  SUBROUTINE: STRLEN - %1=string, %2=var to store length
:: ============================================================
:STRLEN
setlocal enabledelayedexpansion
set "s=%~1"
set "len=0"
:STRLEN_LOOP
if defined s (
    set "s=!s:~1!"
    set /a len+=1
    goto STRLEN_LOOP
)
endlocal & set "%~2=%len%"
goto :eof

:: ============================================================
::  SUBROUTINE: BROWSE_FOLDER
:: ============================================================
:BROWSE_FOLDER
set "BROWSE_TITLE=%~1"
set "BROWSE_VAR=%~2"

for /f "delims=" %%P in ('
    powershell -NoProfile -Command "$null = [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms'); $dlg = New-Object System.Windows.Forms.FolderBrowserDialog; $dlg.Description = '%BROWSE_TITLE:'=''%'; $dlg.ShowNewFolderButton = $true; if ($dlg.ShowDialog() -eq 'OK') { Write-Output $dlg.SelectedPath }"
') do set "%BROWSE_VAR%=%%P"

goto :eof

:: ============================================================
::  SUBROUTINE: WRITE TXT
:: ============================================================
:WRITE_TXT
(
    echo ============================================================
    echo  DIRECTORY LISTING
    echo ============================================================
    echo  Source:    !SOURCE_DIR!
    if %RECURSIVE%==1 (echo  Mode:      Recursive ^(includes subfolders^)
    ) else (echo  Mode:      Top-level only)
    echo  Generated: %date% %time%
    echo  Folders:   !FOLDER_COUNT!
    echo  Files:     !FILE_COUNT!
    echo ============================================================
    echo.
    echo ------------------------------------------------------------
    echo  FOLDERS ^(!FOLDER_COUNT!^)
    echo ------------------------------------------------------------
    if !FOLDER_COUNT! gtr 0 (
        type "!FOLDER_LIST!"
    ) else (
        echo ^(none^)
    )
    echo.
    echo ------------------------------------------------------------
    echo  FILES BY EXTENSION ^(!FILE_COUNT! total^)
    echo ------------------------------------------------------------
) > "!OUTPUT_FILE!"

if !FILE_COUNT! gtr 0 (
    for /f "usebackq delims=" %%E in ("!EXT_SORTED!") do (
        set "CEXT=%%E"
        set "CNT=!EXT_COUNT_%%E!"
        echo.>> "!OUTPUT_FILE!"
        echo [.!CEXT! FILES]  ^(!CNT!^)>> "!OUTPUT_FILE!"
        echo ---------------------------------------->> "!OUTPUT_FILE!"
        type "%temp%\dir_ext_%%E.tmp">> "!OUTPUT_FILE!"
    )
) else (
    echo.>> "!OUTPUT_FILE!"
    echo ^(no files found^)>> "!OUTPUT_FILE!"
)

(
    echo.
    echo ============================================================
    echo  END OF LIST
    echo ============================================================
) >> "!OUTPUT_FILE!"

goto :eof

:: ============================================================
::  SUBROUTINE: WRITE RTF
:: ============================================================
:WRITE_RTF
setlocal enabledelayedexpansion

:: Escape for RTF
set "RTF_SOURCE=!SOURCE_DIR:\=\\!"
set "RTF_SOURCE=!RTF_SOURCE:{=\{}!"
set "RTF_SOURCE=!RTF_SOURCE:}=\}}!"

if %RECURSIVE%==1 (
    set "RTF_MODE=Recursive \\i (includes subfolders)\\i0"
) else (
    set "RTF_MODE=Top-level only"
)

:: RTF Header
(
    echo {\rtf1\ansi\deff0
    echo {\fonttbl{\f0 Arial;}{\f1 Courier New;}}
    echo {\colortbl ;\red0\green0\blue128;\red128\green0\blue0;\red0\green100\blue0;\red128\green0\blue128;}
    echo \f0\fs24
    echo \b\fs32 DIRECTORY LISTING\b0\fs24\par
    echo \par
    echo \b Source:\b0  !RTF_SOURCE!\par
    echo \b Mode:\b0  !RTF_MODE!\par
    echo \b Generated:\b0  %date% %time%\par
    echo \b Folders:\b0  !FOLDER_COUNT!\par
    echo \b Files:\b0  !FILE_COUNT!\par
    echo \par
    echo \b\fs28\cf1 FOLDERS \b0\i ^(!FOLDER_COUNT!^)\i0\fs24\par
    echo \f1
) > "!OUTPUT_FILE!"

:: Folders list
if !FOLDER_COUNT! gtr 0 (
    for /f "usebackq delims=" %%L in ("!FOLDER_LIST!") do (
        set "LINE=%%L"
        set "LINE=!LINE:\=\\!"
        set "LINE=!LINE:{=\{}!"
        set "LINE=!LINE:}=\}}!"
        echo !LINE!\par >> "!OUTPUT_FILE!"
    )
) else (
    echo \i ^(none^)\i0\par >> "!OUTPUT_FILE!"
)

:: Files by extension header
(
    echo \f0\par
    echo \b\fs28\cf2 FILES BY EXTENSION \b0\i ^(!FILE_COUNT! total^)\i0\fs24\par
) >> "!OUTPUT_FILE!"

if !FILE_COUNT! gtr 0 (
    for /f "usebackq delims=" %%E in ("!EXT_SORTED!") do (
        set "CEXT=%%E"
        set "CNT=!EXT_COUNT_%%E!"
        (
            echo \f0\par
            echo \b\fs26\cf3 .!CEXT! FILES \b0\i ^(!CNT!^)\i0\fs24\par
            echo \f1
        ) >> "!OUTPUT_FILE!"
        for /f "usebackq delims=" %%L in ("%temp%\dir_ext_%%E.tmp") do (
            set "LINE=%%L"
            set "LINE=!LINE:\=\\!"
            set "LINE=!LINE:{=\{}!"
            set "LINE=!LINE:}=\}}!"
            echo !LINE!\par >> "!OUTPUT_FILE!"
        )
    )
) else (
    echo \f1\par\i ^(no files found^)\i0\par >> "!OUTPUT_FILE!"
)

:: RTF Footer
(
    echo \f0\par
    echo \cf4\i --- End of list ---\i0\cf0\par
    echo }
) >> "!OUTPUT_FILE!"

endlocal
goto :eof
