@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --clean EventPostGenerator.spec
if exist "dist\EventPostGenerator.exe" copy /Y "dist\EventPostGenerator.exe" "EventPostGenerator.exe"
echo.
echo Fertig: EventPostGenerator.exe
pause
