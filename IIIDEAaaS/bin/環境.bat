@ECHO OFF

SETLOCAL
SET FILENAME=%ProgramData%\Anaconda3\Scripts\activate.bat

IF NOT EXIST %FILENAME% (
	SET FILENAME=%USERPROFILE%\Anaconda3\Scripts\activate.bat
	ECHO %FILENAME%
) else (
	ECHO %FILENAME%
)
CALL %FILENAME%

pip install janome
pip install sumy
pip install tinysegmenter
pip install pysummarization
pip install MeCab

exit /b 0

