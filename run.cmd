@echo off
REM LibreOfficeをUNOリスニングモードでバックグラウンド起動します。
REM startコマンドは待機しません。
REM LibreOfficeの起動コマンドは、ユーザーが確認した正しい形式を使用します。
start "" "C:\Program Files\LibreOffice\program\soffice.exe" --accept=socket,host=localhost,port=2002;urp; --headless --nologo --nofirststartwizard

echo ===== LibreOffice startup command initiated. Attempting to run conversion script with retry logic. =====

REM 変換スクリプトを実行します。
REM convert.py内のリトライロジックがLibreOfficeへの接続を待ちます。
REM ここで入力ファイルと出力ファイルの絶対パスを渡します。
"C:\Program Files\LibreOffice\program\python.exe" convert.py "%cd%\%1" "%cd%\%2" %3

echo ===== Conversion script finished =====

REM LibreOfficeプロセスを終了します。
REM 特定のコマンドライン引数を持つsoffice.exeプロセスを検索し、強制終了させます。
echo Attempting to terminate LibreOffice process using taskkill...
REM taskkill /IM <イメージ名> /FI "COMMAND LINE like '*<コマンドライン引数>*'" /F /T
REM 検索フィルターの構文を調整
taskkill /IM soffice.exe /F /T

echo LibreOffice process termination command executed.

pause
