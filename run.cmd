@echo off
REM ==========================================================================
REM LibreOffice UNOリスニングモードでバックグラウンド起動し、変換スクリプトを実行
REM ==========================================================================

REM Green Red とリセット(NONE) のエスケープシーケンスを定義
set GBack=[42m
set G=[32m
set R=[31m
set N=[0m

REM LibreOfficeをUNOリスニングモードでバックグラウンド起動します。
REM startコマンドは待機しません。
REM LibreOfficeの起動コマンドは、ユーザーが確認した正しい形式を使用します。
SET /P X="- Starting LibreOffice in UNO listening mode..."<NUL
start "" "C:\Program Files\LibreOffice\program\soffice.exe" --accept=socket,host=localhost,port=2002;urp; --headless --nologo --nofirststartwizard
echo %G%[EXECUTED]%N%

REM 変換スクリプトを実行します。
REM convert.py内のリトライロジックがLibreOfficeへの接続を待ちます。
REM ここで入力ファイルと出力ファイルの絶対パスを渡します。
SET /P X="- Starting Conversion..."<NUL
echo %G%[EXECUTING...]%N%
"C:\Program Files\LibreOffice\program\python.exe" convert.py "%cd%\%1" "%cd%\%2" %3


REM LibreOfficeプロセスを終了します。
REM taskkillを使ってsoffice.exeプロセスを強制終了させます。
SET /P X="- Attempting to terminate LibreOffice process..."<NUL
taskkill /IM soffice.exe /F /T > NUL
IF %ERRORLEVEL% EQU 0 (
    echo %G%[OK]%N%
) ELSE (
    REM プロセスが見つからなかった場合や、終了に失敗した場合
    echo %R%[ERROR]%N% Failed to terminate LibreOffice process. Error Code: %ERRORLEVEL%
    IF %ERRORLEVEL% EQU 128 (
        echo %R%[ERROR]%N% LibreOffice process was not found
    ) ELSE (
        echo %R%[ERROR]%N% Termination failed with unexpected error code: %ERRORLEVEL%
    )
)

SET /P X="%GBack%[DONE] Conversion Finished! Press Any Key to exit...%N%"<NUL
pause>NUL