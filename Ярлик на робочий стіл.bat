@echo off
:: Japan moto — створити ярлик на робочому столі
setlocal

set EXE=%~dp0JapanMoto.exe
set ICO=%~dp0icon\iconwn.ico

if not exist "%EXE%" (
  echo Файл JapanMoto.exe не знайдено поруч із цим скриптом.
  echo Переконайтесь, що ви запускаєте цей файл з папки, де знаходиться JapanMoto.exe
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$exe = '%EXE:\=\\%'; ^
   $ico = '%ICO:\=\\%'; ^
   $work = Split-Path $exe; ^
   $desktop = [Environment]::GetFolderPath('Desktop'); ^
   $lnk = Join-Path $desktop 'Japan moto.lnk'; ^
   $ws = New-Object -ComObject WScript.Shell; ^
   $s = $ws.CreateShortcut($lnk); ^
   $s.TargetPath = $exe; ^
   $s.WorkingDirectory = $work; ^
   $s.Description = 'Japan moto - акт, договір, видаткова'; ^
   if (Test-Path $ico) { $s.IconLocation = $ico } else { $s.IconLocation = $exe + ',0' }; ^
   $s.Save(); ^
   Write-Host 'Ярлик створено на робочому столі: ' + $lnk"

if %errorlevel% == 0 (
  echo.
  echo Готово! Ярлик "Japan moto" з'явився на робочому столі.
) else (
  echo.
  echo Не вдалося створити ярлик. Спробуйте запустити від імені адміністратора.
)
pause
