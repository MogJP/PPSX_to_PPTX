@echo off
REM ==========================================================================
REM LibreOffice UNOƒŠƒXƒjƒ“ƒOƒ‚[ƒh‚ÅƒoƒbƒNƒOƒ‰ƒEƒ“ƒh‹N“®‚µA•ÏŠ·ƒXƒNƒŠƒvƒg‚ğÀs
REM ==========================================================================

REM Green Red ‚ÆƒŠƒZƒbƒg(NONE) ‚ÌƒGƒXƒP[ƒvƒV[ƒPƒ“ƒX‚ğ’è‹`
set GBack=[42m
set G=[32m
set R=[31m
set N=[0m

REM LibreOffice‚ğUNOƒŠƒXƒjƒ“ƒOƒ‚[ƒh‚ÅƒoƒbƒNƒOƒ‰ƒEƒ“ƒh‹N“®‚µ‚Ü‚·B
REM startƒRƒ}ƒ“ƒh‚Í‘Ò‹@‚µ‚Ü‚¹‚ñB
REM LibreOffice‚Ì‹N“®ƒRƒ}ƒ“ƒh‚ÍAƒ†[ƒU[‚ªŠm”F‚µ‚½³‚µ‚¢Œ`®‚ğg—p‚µ‚Ü‚·B
SET /P X="- Starting LibreOffice in UNO listening mode..."<NUL
start "" "C:\Program Files\LibreOffice\program\soffice.exe" --accept=socket,host=localhost,port=2002;urp; --headless --nologo --nofirststartwizard
echo %G%[EXECUTED]%N%

REM •ÏŠ·ƒXƒNƒŠƒvƒg‚ğÀs‚µ‚Ü‚·B
REM convert.py“à‚ÌƒŠƒgƒ‰ƒCƒƒWƒbƒN‚ªLibreOffice‚Ö‚ÌÚ‘±‚ğ‘Ò‚¿‚Ü‚·B
REM ‚±‚±‚Å“ü—Íƒtƒ@ƒCƒ‹‚Æo—Íƒtƒ@ƒCƒ‹‚Ìâ‘ÎƒpƒX‚ğ“n‚µ‚Ü‚·B
SET /P X="- Starting Conversion..."<NUL
echo %G%[EXECUTING...]%N%
"C:\Program Files\LibreOffice\program\python.exe" convert.py "%cd%\%1" "%cd%\%2" %3


REM LibreOfficeƒvƒƒZƒX‚ğI—¹‚µ‚Ü‚·B
REM taskkill‚ğg‚Á‚Äsoffice.exeƒvƒƒZƒX‚ğ‹­§I—¹‚³‚¹‚Ü‚·B
SET /P X="- Attempting to terminate LibreOffice process..."<NUL
taskkill /IM soffice.exe /F /T > NUL
IF %ERRORLEVEL% EQU 0 (
    echo %G%[OK]%N%
) ELSE (
    REM ƒvƒƒZƒX‚ªŒ©‚Â‚©‚ç‚È‚©‚Á‚½ê‡‚âAI—¹‚É¸”s‚µ‚½ê‡
    echo %R%[ERROR]%N% Failed to terminate LibreOffice process. Error Code: %ERRORLEVEL%
    IF %ERRORLEVEL% EQU 128 (
        echo %R%[ERROR]%N% LibreOffice process was not found
    ) ELSE (
        echo %R%[ERROR]%N% Termination failed with unexpected error code: %ERRORLEVEL%
    )
)

SET /P X="%GBack%[DONE] Conversion Finished! Press Any Key to exit...%N%"<NUL
pause>NUL