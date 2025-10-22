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

## exeファイル
Remove-Item -Recurse -Force .\dist, .\build -ErrorAction SilentlyContinue
pyinstaller `
  --noconfirm --clean --onefile --windowed `
  --name "TeamsLinkScheduler" `
  --collect-submodules win32com `
  --hidden-import pythoncom `
  --hidden-import pywintypes `
  --add-data "logo.png;." `
  --add-data "logo_spi.png;." `
  .\main.py

## 実行ファイル（.exe）のアイコンを直接埋め込む

1. 「Resource Hacker」 という無料ツールをインストール。
    https://www.angusj.com/resourcehacker/

2. Resource Hacker を開いて
   → TeamsLinkScheduler.exe を開く。

3. 左のツリーで Icon Group を選択。
   → 右クリックして「Replace Resource...」を選択。

4.  新しい .ico ファイルを指定して「Replace」。

5. 「File」→「Save As」で保存（元のファイルを上書きしないように注意）。

## 備考
- 管理者タスクは管理者で起動したシェルから実行。
- エラー 0x8007007B はパス不正。v1.1 は絶対パスの cmd.exe を使用。
