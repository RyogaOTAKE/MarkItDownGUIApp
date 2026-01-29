@echo off
setlocal

REM 仮想環境を作成します
if not exist .venv (
  py -3 -m venv .venv
)

REM 仮想環境を有効化します
call .venv\Scripts\activate.bat

REM 依存を入れます
python -m pip install --upgrade pip
pip install -r requirements.txt

REM ビルドします
REM --collect-all markitdown はサブモジュールやデータを取りこぼしにくくします
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name MarkItDownGUI ^
  --collect-all markitdown ^
  --copy-metadata markitdown ^
  app.py

echo.
echo Build done. dist\MarkItDownGUI.exe
endlocal
