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
  app_6055.py

echo.
echo Build complete.  Folder: dist\JapanMoto\
echo Executable:      dist\JapanMoto\JapanMoto.exe

:: Create desktop shortcut
powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; ^
   $s  = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Japan moto.lnk'); ^
   $s.TargetPath   = '%~dp0dist\JapanMoto\JapanMoto.exe'; ^
   $s.WorkingDirectory = '%~dp0dist\JapanMoto'; ^
   $s.Description  = 'Japan moto — акт, договір, видаткова'; ^
   $s.Save()"

echo Desktop shortcut created: Japan moto.lnk
