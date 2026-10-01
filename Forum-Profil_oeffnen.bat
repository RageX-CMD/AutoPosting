@echo off
setlocal
set "PROFILE=%LOCALAPPDATA%\EventPostGenerator\EdgeProfile"
set "EDGE_X86=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
set "EDGE_X64=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"

if exist "%EDGE_X86%" (
    start "" "%EDGE_X86%" --user-data-dir="%PROFILE%" --profile-directory="Default"
    goto :done
)

if exist "%EDGE_X64%" (
    start "" "%EDGE_X64%" --user-data-dir="%PROFILE%" --profile-directory="Default"
    goto :done
)

where msedge.exe >nul 2>nul
if %errorlevel% equ 0 (
    start "" msedge.exe --user-data-dir="%PROFILE%" --profile-directory="Default"
    goto :done
)

echo Microsoft Edge wurde nicht gefunden.
pause
exit /b 1

:done
echo Das dauerhafte Forum-Profil wurde in Edge geoeffnet.
echo Bitte dieses Edge-Fenster schliessen, bevor die Posting-App gestartet wird.
endlocal
