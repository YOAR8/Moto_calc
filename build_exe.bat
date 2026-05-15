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
  --windowed ^
  --name MotoCalc ^
  --collect-submodules win32com ^
  --collect-data win32com ^
  app_6055.py

if not exist dist\MotoCalc mkdir dist\MotoCalc
copy /Y 6055.xls dist\MotoCalc\6055.xls >nul
copy /Y 6055_MOTO_template.xls dist\MotoCalc\6055_MOTO_template.xls >nul
copy /Y DOGOVIR_6055_template.doc dist\MotoCalc\DOGOVIR_6055_template.doc >nul
copy /Y vidatkova.xls dist\MotoCalc\vidatkova.xls >nul

echo.
echo Build complete. Run: dist\MotoCalc\MotoCalc.exe
echo Output files will be created in %%USERPROFILE%%\Documents\MotoCalc\out
