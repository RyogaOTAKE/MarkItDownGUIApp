# MarkItDown GUI

## 対象
- .docx .pdf .pptx .xlsx を Markdown(.md) に変換します
- 単一ファイル / 複数ファイル変換と、フォルダ一括変換に対応します
- ウィンドウへのドラッグ & ドロップで、ファイルまたはフォルダを選択できます
- 出力は入力ファイルと同じ場所に同名 .md を作ります

## 実行(開発)
1. `uv venv .venv`
2. .venv\Scripts\activate
3. `uv pip install -r requirements.txt`
4. python app.py

## 使い方
- `選択` ボタンから、単一ファイルまたは複数ファイルを選択できます
- ファイルをウィンドウへドラッグ & ドロップすると、ファイル選択欄へ反映されます
- フォルダをウィンドウへドラッグ & ドロップすると、フォルダ一括変換欄へ反映されます
- ファイルとフォルダの同時ドロップには対応していません

## exe作成
- build.bat を実行します
- dist\MarkItDownGUI.exe ができます
