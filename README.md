# MotoCalc (Windows)

Повноцінний додаток з одним вікном для роботи з файлами:
- 6055.xls (джерело)
- 6055_MOTO_template.xls (акт)
- DOGOVIR_6055_template.doc (договір)

## Що робить програма

- Редагує ключові поля у 6055.xls:
  - A3 (номер + дата)
  - C15 (ПІБ)
  - C43 (номер рами) і синхронізує у C39/C42/C43
- Кнопкою Зробити все переносить дані:
  - в акт (Excel)
  - у договір (Word bookmarks)
- Результати пише в: %USERPROFILE%\Documents\MotoCalc\out

## Запуск на Python (fallback)

```powershell
run_app.bat
```

Це запасний режим, якщо антивірус блокує exe.

## Збірка EXE

```powershell
build_exe.bat
```

Після збірки запускати:

```powershell
dist\MotoCalc\MotoCalc.exe
```

## Порада щодо антивірусу

- Менше false positive зазвичай дає режим onedir (вже використано), а не onefile.
- Не використовувати упаковку/обфускацію.
- Збирати на чистій Windows 10/11 VM.
- Якщо є корпоративний AV на Windows Server 2019, додати папку релізу у allowlist.
- Для production бажано підписувати exe code-sign сертифікатом.

## GitHub репозиторій

Цільовий репозиторій:
https://github.com/YOAR8/Moto_calc
