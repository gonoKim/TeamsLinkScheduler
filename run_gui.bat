@echo off
setlocal
if not exist .venv (
  echo [!] venv ������܂���Binstall_env.bat ���Ɏ��s���Ă��������B
  pause
  exit /b 1
)
call .venv\Scripts\activate
python main.py
