@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  py -m venv .venv
)

call .venv\Scripts\activate
python -m pip install -r requirements-windows.txt
python app_6055.py
