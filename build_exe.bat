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
  --onefile ^
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
echo Build complete. Run: dist\JapanMoto.exe
echo Output files will be created in %%USERPROFILE%%\Documents\JapanMoto\out
echo Full generation creates a separate buyer folder using the short FIO.
