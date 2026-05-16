@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  py -m venv .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements-windows.txt

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

set ICON_ARG=
if exist "assets\japanmoto.ico" set ICON_ARG=--icon "assets\japanmoto.ico"
if exist "japanmoto.ico" set ICON_ARG=--icon "japanmoto.ico"

pyinstaller ^
  --noconfirm ^
  --clean ^
  --noupx ^
  --onedir ^
  --windowed ^
  --name JapanMoto ^
  --collect-submodules win32com ^
  --collect-data win32com ^
  --add-data "6055.xls;." ^
  --add-data "6055_MOTO_template.xls;." ^
  --add-data "DOGOVIR_6055_template.doc;." ^
  --add-data "vidatkova.xls;." ^
  %ICON_ARG% ^
  app_6055.py

echo.
echo Build complete.  Folder: dist\JapanMoto\
echo Executable:      dist\JapanMoto\JapanMoto.exe

:: Create desktop shortcut (current user desktop, fallback to public desktop)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop'; ^
   $exe = (Resolve-Path '%~dp0dist\JapanMoto\JapanMoto.exe').Path; ^
   $work = (Resolve-Path '%~dp0dist\JapanMoto').Path; ^
   $desktop = [Environment]::GetFolderPath('Desktop'); ^
   if (-not (Test-Path $desktop)) { $desktop = [Environment]::GetFolderPath('CommonDesktopDirectory') }; ^
   $lnk = Join-Path $desktop 'Japan moto.lnk'; ^
   $ws = New-Object -ComObject WScript.Shell; ^
   $s = $ws.CreateShortcut($lnk); ^
   $s.TargetPath = $exe; ^
   $s.WorkingDirectory = $work; ^
   $s.Description = 'Japan moto - акт, договір, видаткова'; ^
   $s.IconLocation = $exe + ',0'; ^
   $s.Save(); ^
   Write-Host ('Desktop shortcut created: ' + $lnk)" || echo WARNING: shortcut creation failed
