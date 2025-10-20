# Teams Link Scheduler (COM, v1.1)

## 実行
1. install_env.bat
2. run_gui.bat

## デバッグ
- run_debug.bat で DEBUG=1 にして起動（stderrは debug_stderr.log）。
- PowerShell:
  py -3 -m venv .venv
  .\.venv\Scripts\Activate.ps1
  pip install -r requirements.txt
  $env:DEBUG="1"
  python main.py

## 備考
- 管理者タスクは管理者で起動したシェルから実行。
- エラー 0x8007007B はパス不正。v1.1 は絶対パスの cmd.exe を使用。
