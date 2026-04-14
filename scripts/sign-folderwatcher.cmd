@echo off
REM Sign all FolderWatcher executables with EV code signing certificate
REM Requires: USB token plugged in, will prompt for PIN

SET SIGNTOOL="D:\Repos\GB\InnoSetup\includes\signtool.exe"
SET CERT_HASH=AB7147D41404DF8459937ACF61216B5117A32CDD
SET BUILD=D:\Repos\Git\tb-folder-watcher\Build

echo Signing FolderWatcher executables...
echo.

FOR %%F IN ("%BUILD%\FolderWatcher_win32.exe" "%BUILD%\FolderWatcher_win64.exe") DO (
    IF EXIST %%F (
        echo Signing %%~nxF...
        %SIGNTOOL% sign /sha1 %CERT_HASH% /tr http://timestamp.comodoca.com /td sha256 /fd sha256 /du https://grandjean.net /d "FolderWatcher" %%F
        IF ERRORLEVEL 1 (
            echo FAILED: %%~nxF
            pause
            exit /b 1
        )
        echo.
    ) ELSE (
        echo SKIPPED: %%~nxF not found
        echo.
    )
)

echo.
echo Verifying signatures...
echo.

FOR %%F IN ("%BUILD%\FolderWatcher_win32.exe" "%BUILD%\FolderWatcher_win64.exe") DO (
    IF EXIST %%F (
        echo Verifying %%~nxF...
        %SIGNTOOL% verify /pa %%F
        IF ERRORLEVEL 1 (
            echo FAILED verification: %%~nxF
            pause
            exit /b 1
        )
        echo.
    )
)

echo All FolderWatcher executables signed and verified.
pause
