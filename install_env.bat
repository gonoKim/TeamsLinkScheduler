@echo off
setlocal
if not exist .venv (
  py -3 -m venv .venv
)
call .venv\Scripts\python -m pip install --upgrade pip
call .venv\Scripts\pip install -r requirements.txt
echo.
echo [OK] venv �\�z�����Brun_gui.bat �����s���Ă��������B
pause
