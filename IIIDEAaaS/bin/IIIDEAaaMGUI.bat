SETLOCAL
SET FILENAME=%ProgramData%\Anaconda3\Scripts\activate.bat

IF NOT EXIST %FILENAME% (
	SET FILENAME=%USERPROFILE%\Anaconda3\Scripts\activate.bat
	ECHO %FILENAME%
) else (
	ECHO %FILENAME%
)
CALL %FILENAME%

cd /d %~dp0
powershell start-process IIIDEAaaMGUI.exe
exit /b
