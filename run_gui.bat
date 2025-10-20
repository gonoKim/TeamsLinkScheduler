@echo off
setlocal
if not exist .venv (
  echo [!] venv ‚ª‚ ‚è‚Ü‚¹‚ñBinstall_env.bat ‚ğæ‚ÉÀs‚µ‚Ä‚­‚¾‚³‚¢B
  pause
  exit /b 1
)
call .venv\Scripts\activate
python main.py
