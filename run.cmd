@echo off
REM ==========================================================================
REM LibreOffice UNO���X�j���O���[�h�Ńo�b�N�O���E���h�N�����A�ϊ��X�N���v�g�����s
REM ==========================================================================

REM Green Red �ƃ��Z�b�g(NONE) �̃G�X�P�[�v�V�[�P���X���`
set GBack=[42m
set G=[32m
set R=[31m
set N=[0m

REM LibreOffice��UNO���X�j���O���[�h�Ńo�b�N�O���E���h�N�����܂��B
REM start�R�}���h�͑ҋ@���܂���B
REM LibreOffice�̋N���R�}���h�́A���[�U�[���m�F�����������`�����g�p���܂��B
SET /P X="- Starting LibreOffice in UNO listening mode..."<NUL
start "" "C:\Program Files\LibreOffice\program\soffice.exe" --accept=socket,host=localhost,port=2002;urp; --headless --nologo --nofirststartwizard
echo %G%[EXECUTED]%N%

REM �ϊ��X�N���v�g�����s���܂��B
REM convert.py���̃��g���C���W�b�N��LibreOffice�ւ̐ڑ���҂��܂��B
REM �����œ��̓t�@�C���Əo�̓t�@�C���̐�΃p�X��n���܂��B
SET /P X="- Starting Conversion..."<NUL
echo %G%[EXECUTING...]%N%
"C:\Program Files\LibreOffice\program\python.exe" convert.py "%cd%\%1" "%cd%\%2" %3


REM LibreOffice�v���Z�X���I�����܂��B
REM taskkill���g����soffice.exe�v���Z�X�������I�������܂��B
SET /P X="- Attempting to terminate LibreOffice process..."<NUL
taskkill /IM soffice.exe /F /T > NUL
IF %ERRORLEVEL% EQU 0 (
    echo %G%[OK]%N%
) ELSE (
    REM �v���Z�X��������Ȃ������ꍇ��A�I���Ɏ��s�����ꍇ
    echo %R%[ERROR]%N% Failed to terminate LibreOffice process. Error Code: %ERRORLEVEL%
    IF %ERRORLEVEL% EQU 128 (
        echo %R%[ERROR]%N% LibreOffice process was not found
    ) ELSE (
        echo %R%[ERROR]%N% Termination failed with unexpected error code: %ERRORLEVEL%
    )
)

SET /P X="%GBack%[DONE] Conversion Finished! Press Any Key to exit...%N%"<NUL
pause>NUL