@echo off
REM LibreOffice��UNO���X�j���O���[�h�Ńo�b�N�O���E���h�N�����܂��B
REM start�R�}���h�͑ҋ@���܂���B
REM LibreOffice�̋N���R�}���h�́A���[�U�[���m�F�����������`�����g�p���܂��B
start "" "C:\Program Files\LibreOffice\program\soffice.exe" --accept=socket,host=localhost,port=2002;urp; --headless --nologo --nofirststartwizard

echo ===== LibreOffice startup command initiated. Attempting to run conversion script with retry logic. =====

REM �ϊ��X�N���v�g�����s���܂��B
REM convert.py���̃��g���C���W�b�N��LibreOffice�ւ̐ڑ���҂��܂��B
REM �����œ��̓t�@�C���Əo�̓t�@�C���̐�΃p�X��n���܂��B
"C:\Program Files\LibreOffice\program\python.exe" convert.py "%cd%\%1" "%cd%\%2" %3

echo ===== Conversion script finished =====

REM LibreOffice�v���Z�X���I�����܂��B
REM ����̃R�}���h���C������������soffice.exe�v���Z�X���������A�����I�������܂��B
echo Attempting to terminate LibreOffice process using taskkill...
REM taskkill /IM <�C���[�W��> /FI "COMMAND LINE like '*<�R�}���h���C������>*'" /F /T
REM �����t�B���^�[�̍\���𒲐�
taskkill /IM soffice.exe /F /T

echo LibreOffice process termination command executed.

pause
