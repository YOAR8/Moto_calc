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
if exist "icon\iconwn.ico" set ICON_ARG=--icon "icon\iconwn.ico"

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
  --add-data "DOGOVIR_6055_template.doc;." ^
  --add-data "vidatkova.xls;." ^
  --add-data "icon\iconwn.ico;icon" ^
  --add-data "icon\iconwn 48.ico;icon" ^
  --add-data "icon\iconwn 64.ico;icon" ^
  --add-data "\320\206\320\235\320\241\320\242\320\240\320\243\320\232\320\246\320\206\320\257.md;." ^
  %ICON_ARG% ^
  app_6055.py

:: Copy shortcut helper and readme to dist folder
if exist "dist\JapanMoto\" (
  copy /Y "\320\206\320\235\320\241\320\242\320\240\320\243\320\232\320\246\320\206\320\257.md" "dist\JapanMoto\" >nul 2>&1
  copy /Y "\320\257\321\200\320\273\320\270\320\272 \320\275\320\260 \321\200\320\276\320\261\320\276\321\207\320\270\320\271 \321\201\321\202\321\226\320\273.bat" "dist\JapanMoto\" >nul 2>&1
)

echo.
echo Build complete.  Folder: dist\JapanMoto\
echo Executable:      dist\JapanMoto\JapanMoto.exe

:: Create desktop shortcut (current user desktop, fallback to public desktop)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop'; ^
   $exe = (Resolve-Path '%~dp0dist\JapanMoto\JapanMoto.exe').Path; ^
   $work = (Resolve-Path '%~dp0dist\JapanMoto').Path; ^
   $ico = Join-Path $work 'icon\iconwn.ico'; ^
   $desktop = [Environment]::GetFolderPath('Desktop'); ^
   if (-not (Test-Path $desktop)) { $desktop = [Environment]::GetFolderPath('CommonDesktopDirectory') }; ^
   $lnk = Join-Path $desktop 'Japan moto.lnk'; ^
   $ws = New-Object -ComObject WScript.Shell; ^
   $s = $ws.CreateShortcut($lnk); ^
   $s.TargetPath = $exe; ^
   $s.WorkingDirectory = $work; ^
   $s.Description = 'Japan moto - акт, договір, видаткова'; ^
   if (Test-Path $ico) { $s.IconLocation = "$ico,0" } else { $s.IconLocation = "$exe,0" }; ^
   $s.Save(); ^
   Write-Host ('Desktop shortcut created: ' + $lnk)" || echo WARNING: shortcut creation failed
