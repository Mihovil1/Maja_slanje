@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PY=
where py >nul 2>nul && set PY=py
if not defined PY (where python >nul 2>nul && set PY=python)
if not defined PY (
  echo.
  echo  Python nije instaliran.
  echo  Instalirajte ga s https://www.python.org/downloads/
  echo  i kod instalacije OBAVEZNO ukljucite "Add python.exe to PATH".
  echo.
  pause
  exit /b
)
if not exist ".instalirano" (
  echo Prvo pokretanje - instaliram potrebne pakete, pricekajte...
  %PY% -m pip install --user --quiet openpyxl pywin32
  if errorlevel 1 (
    echo Instalacija paketa nije uspjela.
    pause
    exit /b
  )
  echo ok> .instalirano
)
start "" %PY%w slanje_obavijesti.py 2>nul || start "" %PY% slanje_obavijesti.py
