@echo off
setlocal
if not exist .venv (
  echo [!] venv がありません。install_env.bat を先に実行してください。
  pause
  exit /b 1
)
call .venv\Scripts\activate
set DEBUG=1
python main.py 1>nul 2>debug_stderr.log
echo.
echo [OK] DEBUG=1 で起動しました。エラーは debug_stderr.log を確認してください。
pause
