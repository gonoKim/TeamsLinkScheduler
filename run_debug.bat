@echo off
setlocal
if not exist .venv (
  echo [!] venv ������܂���Binstall_env.bat ���Ɏ��s���Ă��������B
  pause
  exit /b 1
)
call .venv\Scripts\activate
set DEBUG=1
python main.py 1>nul 2>debug_stderr.log
echo.
echo [OK] DEBUG=1 �ŋN�����܂����B�G���[�� debug_stderr.log ���m�F���Ă��������B
pause
