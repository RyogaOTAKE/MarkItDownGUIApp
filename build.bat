@echo off
setlocal

REM Create the virtual environment.
if not exist .venv (
  uv venv .venv
)

REM Activate the virtual environment.
call .venv\Scripts\activate.bat

REM Install dependencies.
uv pip install -r requirements.txt

REM Build with the spec file.
REM Include magika model files and tkinterdnd2 data files.
pyinstaller --noconfirm --clean MarkItDownGUI.spec

echo.
echo Build done. dist\MarkItDownGUI.exe
endlocal
